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
