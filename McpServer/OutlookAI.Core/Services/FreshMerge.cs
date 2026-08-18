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

        /// <summary>
        /// Coverage hole: the sweep ran but swept no folder whatsoever, and at least one
        /// folder in its scope existed and went uncovered. A scope whose folders are all
        /// ABSENT sweeps nothing too and raises this not at all - see
        /// <see cref="DescribeCoverageGaps"/>.
        /// </summary>
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
        /// Coverage hole: the index holds NO mail at all for part of this search's scope, so
        /// there was no frontier to open the sweep window from and it fell back to a fixed
        /// span (<c>MailService.EmptyIndexSweepWindow</c>). Everything older than that span,
        /// in the store(s) named by <see cref="SweepInfo.StoresWithoutIndex"/>, is in NEITHER
        /// tier: not in the index, which has no rows for it, and not in the sweep, which
        /// starts after it.
        /// <para>
        /// A CODE rather than a fourth <see cref="FreshnessLive"/>/<see cref="FreshnessPartial"/>/
        /// <see cref="FreshnessIndexOnly"/> value, and the reasoning is the one the three-value
        /// contract was designed around. Those three answer "did the LIVE check run, and did it
        /// cover its scope", <c>degraded</c> is derived from them and is the single thing agents
        /// are taught to branch on, and a fourth value would be a token no caller has been
        /// taught - it would read as "unknown", i.e. as a reason to doubt an answer without
        /// saying what to do about it, and every existing `freshness == "live"` check would
        /// silently stop matching a state that is genuinely not live. The codes are the axis
        /// this contract already extends along: each one names ONE hole, each earns its own
        /// advice sentence, and each makes the search <see cref="FreshnessPartial"/> and
        /// <c>degraded: true</c> through machinery that already exists. This hole is simply the
        /// widest of them, so it sorts first.
        /// </para>
        /// <para>
        /// It is also the only code raised by something the sweep did not do: it describes the
        /// INDEX tier, and it therefore survives a sweep that never ran
        /// (<see cref="DescribeCoverageGaps"/>). "The sweep could not run" and "the index has
        /// nothing for this store" are independent facts and an answer missing both tiers must
        /// say both.
        /// </para>
        /// </summary>
        public const string GapNoIndexFrontier = "no_index_frontier";

        /// <summary>Whether a search's window leaves the freshness sweep anything to find.</summary>
        public enum SweepWindowVerdict
        {
            /// <summary>Part of the requested window lies past the index frontier - sweep it.</summary>
            Needed = 0,

            /// <summary>
            /// The requested window ends at or before the sweep would start, so the index
            /// already covers all of it and a sweep could not add a single item.
            /// </summary>
            NotNeeded = 1,
        }

        /// <summary>
        /// Whether the freshness sweep has anything to do for this request - the same
        /// distinction <c>ClassifyDefaultFolder</c> makes between a folder that is ABSENT
        /// and one that is UNREADABLE, one level up.
        /// <para>
        /// A search bounded by <c>before</c> to mail older than the index frontier cannot
        /// be missing recent mail: there is no recent mail inside its window. Until this
        /// existed, that case set <c>performed = false</c> - the value that means "the
        /// sweep could not run" - so a search deliberately aimed at old mail was told it
        /// was <c>degraded</c> and <c>index-only</c>, i.e. that it might be missing exactly
        /// the mail its own bounds exclude.
        /// </para>
        /// <para>
        /// The boundary is inclusive on purpose: at <c>before == gapStart</c> the window is
        /// empty (<c>before</c> is exclusive), so there is still nothing to sweep.
        /// </para>
        /// </summary>
        public static SweepWindowVerdict DecideSweepWindow(DateTime gapStartUtc, DateTime? beforeUtc)
        {
            return beforeUtc.HasValue && beforeUtc.Value <= gapStartUtc
                ? SweepWindowVerdict.NotNeeded
                : SweepWindowVerdict.Needed;
        }

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
        /// <see cref="SweepInfo.FoldersAbsent"/> RAISES no branch here and SUPPRESSES exactly
        /// one, and the difference between those two sentences is a fix and a regression.
        /// A default folder the store does not have is not a hole in the coverage: there is
        /// nothing behind it to cover. Counting absence as a skip made every non-folder-
        /// scoped search on a profile with such a store report <c>degraded: true</c> and
        /// <c>freshness: "partial"</c> - a flag that cries wolf is worse than no flag.
        /// </para>
        /// <para>
        /// The comment that stood here said absence was read by no branch at all, and while
        /// the counters were whole-sweep totals that was very nearly true: it took a profile
        /// with no arrival-path folder ANYWHERE for the absent count to be the whole story.
        /// Once the counters became per store (2026-08-18) it stopped being true for a far more
        /// ordinary shape - a PST, an archive-only store, a shared mailbox mounted without
        /// the four defaults - where <c>foldersSwept: 0, foldersAbsent: 4</c> is simply what
        /// a complete sweep of that store looks like. Reading only the zero re-created the
        /// very alarm the absent counter was introduced to remove, one scope down.
        /// </para>
        /// <para>
        /// So <see cref="GapNothingSwept"/> now asks WHY nothing was swept. Nothing to sweep
        /// is not a gap; nothing swept because folders failed or were skipped still is, and
        /// so is a sweep that reached a store it was asked about and covered none of it (all
        /// four counters zero - the store the sweep never got to). Absence only ever
        /// suppresses on its own: one absent folder next to one unreadable folder is still a
        /// hole, because that other folder exists and holds mail nobody read.
        /// </para>
        /// </summary>
        public static IReadOnlyList<string>? DescribeCoverageGaps(SweepInfo sweep)
        {
            if (sweep == null)
            {
                throw new ArgumentNullException(nameof(sweep));
            }

            // FIRST, and before the "did it run" gate below: this one is a fact about the
            // INDEX tier, not about the sweep, so a sweep that could not run or had nothing
            // to do does not make it go away. It is also the widest hole here - a whole tier
            // contributed nothing to part of the scope - and the gap order is the order the
            // advice sentences are emitted in.
            List<string> gaps = new List<string>();
            if (sweep.IndexFrontierMissing == true)
            {
                gaps.Add(GapNoIndexFrontier);
            }

            if (!sweep.Performed)
            {
                return gaps.Count == 0 ? null : gaps;
            }

            // Everything the sweep meant to walk here turned out not to exist. That is a
            // complete answer about this scope, not an empty one.
            bool nothingExistedToSweep = sweep.FoldersAbsent > 0
                && sweep.FoldersSkipped == 0
                && sweep.FoldersFailed == 0;

            if (sweep.FoldersSwept == 0 && !nothingExistedToSweep)
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
        /// <para>
        /// A sweep that was NOT NEEDED (<see cref="SweepInfo.NotNeeded"/>) is
        /// <see cref="FreshnessLive"/> for the same reason: the answer is not missing
        /// anything a sweep could have found, which is precisely what these three values
        /// exist to say. It is deliberately not a fourth value - <c>degraded</c> is what
        /// the tool description tells an agent to relay and it must be absent here, while
        /// a value no agent has been taught would read as a reason to doubt the answer.
        /// The sweep block still says what happened: <c>performed: false</c> with
        /// <c>notNeeded: true</c>.
        /// </para>
        /// <para>
        /// ONE exception to that, and it is the whole of <see cref="GapNoIndexFrontier"/>:
        /// "not needed" is a claim that the INDEX already covers the requested window, and
        /// over a scope with no index rows that claim is simply false. A search bounded by
        /// <c>before</c> to mail older than the fallback window would otherwise return an
        /// empty list, out of a store the index has never seen, and call itself <c>live</c>.
        /// So a not-needed sweep is live only while the frontier it was measured against
        /// exists.
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
                if (sweep.NotNeeded == true)
                {
                    return sweep.IndexFrontierMissing == true ? FreshnessPartial : FreshnessLive;
                }

                // Could not run. Still index-only - that value means exactly "the sweep
                // never ran" and callers pin it - and a missing frontier is reported
                // alongside it as a coverage code, not by renaming this state.
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
