using System;
using System.Collections.Generic;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// Pure logic behind fresh-mode merging (v3.MD D19): term matching for gap-swept
    /// items (the sweep has no CONTAINS engine, so the query terms are re-applied
    /// client-side over subject/body/sender) and boundary de-duplication between index
    /// hits and swept items. Unit-tested in T1; no COM, no I/O.
    /// </summary>
    public static class FreshMerge
    {
        /// <summary>Sweep refusal: recipient ('to') filters have no sweep-side equivalent.</summary>
        public const string RecipientFilterNotSweepable = "RecipientFilterNotSweepable";

        /// <summary>Sweep refusal: attachment CONTENT is matched by the index tier alone (D47).</summary>
        public const string AttachmentContentNotSweepable = "AttachmentContentNotSweepable";

        /// <summary>Freshness verdict: the sweep ran and covered everything it was asked to.</summary>
        public const string FreshnessLive = "live";

        /// <summary>Freshness verdict: the sweep ran but left at least one coverage hole.</summary>
        public const string FreshnessPartial = "partial";

        /// <summary>Freshness verdict: the sweep did not run at all, so nothing was checked live.</summary>
        public const string FreshnessIndexOnly = "index-only";

        /// <summary>Coverage hole: the sweep ran but swept no folder whatsoever.</summary>
        public const string GapNothingSwept = "nothing_swept";

        /// <summary>Coverage hole: Outlook would not enumerate one or more folders.</summary>
        public const string GapFoldersFailed = "folders_failed";

        /// <summary>Coverage hole: the per-folder item cap truncated a folder's window.</summary>
        public const string GapItemCap = "item_cap";

        /// <summary>Coverage hole: the subtree walk stopped at its folder cap.</summary>
        public const string GapFolderCap = "folder_cap";

        /// <summary>Coverage hole: the subtree walk stopped at its time budget.</summary>
        public const string GapTimeBudget = "time_budget";

        /// <summary>Coverage hole: the subtree walk refused folders past its depth guard.</summary>
        public const string GapDepthLimit = "depth_limit";

        /// <summary>
        /// Coverage hole: folders were skipped because they could not be resolved or
        /// enumerated. A default folder the store does not HAVE is not one of these - it is
        /// counted in <see cref="SweepInfo.FoldersAbsent"/> and raises nothing, because a
        /// folder that does not exist cannot be hiding mail.
        /// </summary>
        public const string GapFoldersSkipped = "folders_skipped";

        /// <summary>
        /// Every way a sweep that RAN can still have covered less than its scope, decided
        /// in one pure place and pinned in T1 - the same shape as
        /// <see cref="SweepRefusalReason"/>.
        /// <para>
        /// Returns null for a sweep that covered everything, and for one that never ran:
        /// "did not run" is <see cref="FreshnessIndexOnly"/>, a different state with a
        /// different remedy, and reporting it as a partial sweep would blur the two.
        /// </para>
        /// <para>
        /// Order is severity-first (no coverage at all, then lost folders, then truncation)
        /// because it is also the order the matching advice sentences are emitted in.
        /// <see cref="GapFoldersSkipped"/> is deliberately suppressed once a BOUND stopped
        /// the walk: the folders past a cap, a budget or the depth guard are counted as
        /// skipped too, and reporting them a second time as unreadable would misattribute
        /// them - the bound's own code already says what happened.
        /// </para>
        /// <para>
        /// <see cref="SweepInfo.FoldersAbsent"/> is read by NO branch here, on purpose. A
        /// default folder the store does not have is not a hole in the coverage: there is
        /// nothing behind it to cover. Counting absence as a skip made every non-folder-
        /// scoped search on a profile with such a store report <c>degraded: true</c> and
        /// <c>freshness: "partial"</c> - a flag that cries wolf is worse than no flag.
        /// </para>
        /// </summary>
        public static IReadOnlyList<string>? DescribeCoverageGaps(SweepInfo sweep)
        {
            if (sweep == null)
            {
                throw new ArgumentNullException(nameof(sweep));
            }

            if (!sweep.Performed)
            {
                return null;
            }

            List<string> gaps = new List<string>();
            if (sweep.FoldersSwept == 0)
            {
                gaps.Add(GapNothingSwept);
            }

            if (sweep.FoldersFailed > 0)
            {
                gaps.Add(GapFoldersFailed);
            }

            if (sweep.FolderCapReached == true)
            {
                gaps.Add(GapFolderCap);
            }

            if (sweep.TimeBudgetExceeded == true)
            {
                gaps.Add(GapTimeBudget);
            }

            if (sweep.DepthLimitReached == true)
            {
                gaps.Add(GapDepthLimit);
            }

            if (sweep.FolderCapReached != true
                && sweep.TimeBudgetExceeded != true
                && sweep.DepthLimitReached != true
                && sweep.FoldersSkipped > 0)
            {
                gaps.Add(GapFoldersSkipped);
            }

            if (sweep.ItemCappedFolders != null && sweep.ItemCappedFolders.Count > 0)
            {
                gaps.Add(GapItemCap);
            }

            return gaps.Count == 0 ? null : gaps;
        }

        /// <summary>
        /// The freshness verdict for one search, from its sweep block alone: three states,
        /// not two (<see cref="FreshnessLive"/> / <see cref="FreshnessPartial"/> /
        /// <see cref="FreshnessIndexOnly"/>). <c>degraded</c> is derived from it - anything
        /// that is not <see cref="FreshnessLive"/> is degraded.
        /// <para>
        /// A null sweep means no sweep block exists at all, which only the internal
        /// index-only escape hatch produces (<c>SearchRequest.IndexOnly</c>, not exposed on
        /// the MCP tool). That caller asked for index rows and got exactly them, so nothing
        /// was withheld and the verdict stays <see cref="FreshnessLive"/> - unchanged from
        /// before this classification existed.
        /// </para>
        /// </summary>
        public static string ClassifyFreshness(SweepInfo? sweep)
        {
            if (sweep == null)
            {
                return FreshnessLive;
            }

            if (!sweep.Performed)
            {
                return FreshnessIndexOnly;
            }

            IReadOnlyList<string>? gaps = DescribeCoverageGaps(sweep);
            return gaps == null || gaps.Count == 0 ? FreshnessLive : FreshnessPartial;
        }

        /// <summary>
        /// Why the freshness sweep cannot answer this request at all, or null when it can
        /// (D47 - the rule stated in ONE place, and pinned in T1).
        /// <para>
        /// The sweep reads <c>MailItem.Subject</c> and <c>MailItem.Body</c> through COM. It
        /// never opens an attachment, so every row it can produce is a MESSAGE row. That
        /// makes the two attachment flags asymmetric on purpose, and the asymmetry is the
        /// whole rule:
        /// </para>
        /// <list type="bullet">
        /// <item><description><b>attachment hits ONLY</b> - the caller asked for rows the
        /// sweep can never produce, so merging its message rows would return exactly what
        /// the filter excludes. The sweep is refused (and reported), matching the
        /// exhaustive tier, which refuses such a search outright for the same reason.</description></item>
        /// <item><description><b>attachment hits EXCLUDED</b> - the caller asked for
        /// message rows, which is all the sweep produces. It runs normally; nothing to
        /// filter.</description></item>
        /// <item><description><b>attachment hits INCLUDED</b> (the default) - the sweep
        /// contributes its message rows and simply adds no attachment ones.</description></item>
        /// </list>
        /// </summary>
        public static string? SweepRefusalReason(bool hasRecipientFilter, bool attachmentHitsOnly)
        {
            if (hasRecipientFilter)
            {
                return RecipientFilterNotSweepable;
            }

            if (attachmentHitsOnly)
            {
                return AttachmentContentNotSweepable;
            }

            return null;
        }

        /// <summary>
        /// Applies the search terms (ANDed; a trailing '*' marks a prefix stem and is
        /// matched as a case-insensitive substring, slightly over-matching the index's
        /// word-prefix semantics - the acceptable direction for a freshness sweep) within
        /// <paramref name="searchIn"/>.
        /// <para>
        /// Tier alignment (D40/SF-6): the scopes mean the same thing here as in the index
        /// tier - subject, body, or either - so a hit does not appear and then vanish once
        /// the index frontier passes the item. Sender name/address used to be matched here
        /// as well ("approximating the index's all-properties CONTAINS" - a belief SF-6
        /// disproved); it no longer is, because the index tier never matched senders by
        /// term. Sender matching is the <c>from</c> filter's job in every tier.
        /// </para>
        /// </summary>
        public static bool MatchesTerms(ComMailBrief item, IReadOnlyList<string>? terms, SearchIn searchIn)
        {
            if (item == null)
            {
                throw new ArgumentNullException(nameof(item));
            }

            if (terms == null || terms.Count == 0)
            {
                return true;
            }

            bool matchSubject = searchIn != SearchIn.BodyOnly;
            bool matchBody = searchIn != SearchIn.SubjectOnly;

            foreach (string term in terms)
            {
                string stem = term.EndsWith("*", StringComparison.Ordinal)
                    ? term.Substring(0, term.Length - 1)
                    : term;
                if (stem.Length == 0)
                {
                    continue;
                }

                bool found = (matchSubject && ContainsIgnoreCase(item.Subject, stem))
                    || (matchBody && ContainsIgnoreCase(item.Body, stem));
                if (!found)
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// True when a swept item is the same message as an index hit: same store, same
        /// leaf folder (Sent vs Inbox copies of a self-send must stay distinct), same
        /// subject, received within <paramref name="toleranceSeconds"/>.
        /// </summary>
        public static bool IsDuplicate(ComMailBrief item, IndexHit hit, int toleranceSeconds)
        {
            if (item == null)
            {
                throw new ArgumentNullException(nameof(item));
            }

            if (hit == null)
            {
                throw new ArgumentNullException(nameof(hit));
            }

            if (!string.Equals(item.StoreDisplayName, ResolveHitStore(hit), StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string? hitFolderLeaf = hit.FolderSegments.Count > 0 ? hit.FolderSegments[hit.FolderSegments.Count - 1] : null;
            if (item.FolderName != null && hitFolderLeaf != null
                && !string.Equals(item.FolderName, hitFolderLeaf, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (!string.Equals(item.Subject ?? string.Empty, hit.Subject ?? string.Empty, StringComparison.Ordinal))
            {
                return false;
            }

            if (!hit.DateReceivedUtc.HasValue)
            {
                return item.ReceivedTime == null;
            }

            return OutlookComSession.ReceivedTimeMatches(item.ReceivedTime, hit.DateReceivedUtc.Value, toleranceSeconds);
        }

        /// <summary>
        /// Filters swept items down to the ones NOT already present among the index
        /// hits. <paramref name="duplicateCount"/> reports how many were dropped.
        /// </summary>
        public static IReadOnlyList<ComMailBrief> SelectFreshOnly(
            IReadOnlyList<ComMailBrief> sweptItems,
            IReadOnlyList<IndexHit> indexHits,
            int toleranceSeconds,
            out int duplicateCount)
        {
            if (sweptItems == null)
            {
                throw new ArgumentNullException(nameof(sweptItems));
            }

            if (indexHits == null)
            {
                throw new ArgumentNullException(nameof(indexHits));
            }

            List<ComMailBrief> fresh = new List<ComMailBrief>();
            int duplicates = 0;
            HashSet<string> seenEntryIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ComMailBrief item in sweptItems)
            {
                if (!seenEntryIds.Add(item.EntryId))
                {
                    duplicates++;
                    continue;
                }

                bool isDuplicate = false;
                foreach (IndexHit hit in indexHits)
                {
                    if (IsDuplicate(item, hit, toleranceSeconds))
                    {
                        isDuplicate = true;
                        break;
                    }
                }

                if (isDuplicate)
                {
                    duplicates++;
                }
                else
                {
                    fresh.Add(item);
                }
            }

            duplicateCount = duplicates;
            return fresh;
        }

        /// <summary>
        /// Store display name a hit belongs to under the delegate rule (v3.MD section
        /// 0.8 fact 3: /1/ subtree items live in the store named by the first folder
        /// segment).
        /// </summary>
        public static string? ResolveHitStore(IndexHit hit)
        {
            if (hit == null)
            {
                throw new ArgumentNullException(nameof(hit));
            }

            if (hit.StoreType == 1 && hit.FolderSegments.Count > 0)
            {
                return hit.FolderSegments[0];
            }

            return hit.StoreDisplayName;
        }

        private static bool ContainsIgnoreCase(string? haystack, string needle)
        {
            return haystack != null && haystack.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
