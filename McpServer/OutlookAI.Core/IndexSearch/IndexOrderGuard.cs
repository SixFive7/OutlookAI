using System;
using System.Collections.Generic;

namespace OutlookAI.Core.IndexSearch
{
    /// <summary>
    /// Keeps rows the ORDER BY column cannot rank from taking result slots away from rows
    /// it can. The guarantee it exists to hold is one sentence: <b>a row the index cannot
    /// date can never reduce the number of dated rows a search returns.</b>
    /// <para>
    /// WHY THIS EXISTS. Since gap B3 a message-level row is admitted whatever its item
    /// class, so a store-scoped statement no longer carries a Kind predicate at all and the
    /// candidate set now includes the appointments and contacts of folders the COM tiers
    /// never open. Those rows carry no <c>System.Message.DateReceived</c>. The statement
    /// asks for <c>SELECT TOP n ... ORDER BY System.Message.DateReceived DESC</c>, and the
    /// provider applies that ranking BEFORE the cut, so where a NULL sorts under DESC is
    /// what decides whether those rows fill the n. This project has never measured the
    /// Windows Search provider's NULL ordering. If NULLs sort last the rows are harmless;
    /// if they sort first they fill the cut and the real mail behind them never leaves the
    /// provider - a search returning a full page of appointments and no mail at all, with
    /// nothing short, nothing dropped and nothing to notice. The widening would then have
    /// cost mail, which is the one direction the maintainer's rule forbids.
    /// </para>
    /// <para>
    /// WHY NOT JUST ASK FOR THE ORDERING TO PUT THEM LAST. WS-SQL has no NULLS LAST and no
    /// COALESCE in ORDER BY, so the ordering cannot be made explicit. Over-fetching more
    /// candidates does not help either: a Calendar folder can hold more undated rows than
    /// any bounded over-fetch, so a factor large enough to be safe does not exist. And
    /// excluding undated rows outright would work, but it would re-narrow the tier the B3
    /// decision widened, which is not this guard's call to make.
    /// </para>
    /// <para>
    /// SO THE GUARD IS TWO PARTS, both pure and both pinned in T1.
    /// <list type="number">
    /// <item><see cref="RankableFirst"/> decides the order the SERVICE trims in. The trim to
    /// <c>Top</c> used to take the provider's first n admitted rows, so under a NULLs-first
    /// provider it kept the undated ones even when the dated rows were already in hand.
    /// Ranked rows now come first, ordered by their key, and unranked rows follow in
    /// provider order - the same "undated last" convention MailService already applies when
    /// it merges sweep hits into the same list.</item>
    /// <item><see cref="NeedsOrderKeyRefetch"/> decides when the PROVIDER may have hidden
    /// rows the trim can no longer recover, and the service then re-runs the statement with
    /// <see cref="WsSqlBuilder.BuildOrderKeyPresence"/> and merges. The two results are
    /// unioned, so the refetch can only ever add.</item>
    /// </list>
    /// </para>
    /// <para>
    /// WHAT THIS DELIBERATELY DOES NOT PROTECT AGAINST. A meeting request, bounce report or
    /// read receipt carries a received date, so it competes with mail on the same axis and a
    /// newer one can push an older mail off the end of a Top-n list. That is the B3 decision
    /// working, not displacement: those classes ARE mail under the admission rule
    /// (<see cref="OutlookAI.Core.Mapi.MailItemAdmission"/>). What is ruled out here is a row
    /// with no position in the ordering at all taking a slot by accident of NULL collation.
    /// </para>
    /// </summary>
    public static class IndexOrderGuard
    {
        /// <summary>
        /// True when <paramref name="hit"/> carries a value in the column
        /// <paramref name="order"/> ranks by, so the provider can place it.
        /// </summary>
        public static bool HasOrderKey(IndexHit hit, IndexOrder order)
        {
            if (hit == null)
            {
                throw new ArgumentNullException(nameof(hit));
            }

            switch (order)
            {
                case IndexOrder.SizeDescending:
                    return hit.SizeBytes.HasValue;
                case IndexOrder.DateReceivedDescending:
                    return hit.DateReceivedUtc.HasValue;
                default:
                    throw new ArgumentException("Unknown IndexOrder value.", nameof(order));
            }
        }

        /// <summary>
        /// True when any row in <paramref name="rows"/> lacks the ordering key. Answered over
        /// EVERY row the statement returned rather than over the admitted ones, because a row
        /// this tier later refuses still consumed a slot at the provider.
        /// </summary>
        public static bool AnyRowMissingOrderKey(IReadOnlyList<IndexHit> rows, IndexOrder order)
        {
            if (rows == null)
            {
                throw new ArgumentNullException(nameof(rows));
            }

            for (int i = 0; i < rows.Count; i++)
            {
                if (!HasOrderKey(rows[i], order))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Whether the statement must be re-run for the rows the ordering may have hidden.
        /// Both conditions are necessary and together they are sufficient under ANY NULL
        /// collation the provider might use - which is the point, since the collation is the
        /// thing nobody here has measured.
        /// <list type="bullet">
        /// <item><paramref name="rowsReturned"/> below <paramref name="sqlTop"/> means the
        /// statement was not cut off, so every matching row is already in hand and nothing
        /// can have been displaced, whatever order they arrived in.</item>
        /// <item>No unranked row in the returned block means no unranked row sorted above the
        /// cut: had one done so it would be in the block. So none of them took a slot.</item>
        /// </list>
        /// A refetch when neither holds would be wasted work; skipping one when both hold is
        /// the silent loss this guard exists to prevent, so the test is deliberately generous
        /// in the safe direction. A pool that exactly fills TOP without truncating triggers a
        /// refetch that simply returns a subset of what is already known.
        /// </summary>
        public static bool NeedsOrderKeyRefetch(int rowsReturned, int sqlTop, bool anyRowMissingOrderKey)
        {
            if (sqlTop < 1)
            {
                throw new ArgumentException("sqlTop must be positive.", nameof(sqlTop));
            }

            if (rowsReturned < 0)
            {
                throw new ArgumentException("rowsReturned must not be negative.", nameof(rowsReturned));
            }

            return anyRowMissingOrderKey && rowsReturned >= sqlTop;
        }

        /// <summary>
        /// Unions two admitted-row lists by item URL, keeping the first occurrence. Used to
        /// fold a refetch into the statement's own rows: the result is a superset of both, so
        /// the refetch cannot subtract even if the provider answers it inconsistently.
        /// </summary>
        public static IReadOnlyList<IndexHit> Merge(IReadOnlyList<IndexHit> first, IReadOnlyList<IndexHit> second)
        {
            if (first == null)
            {
                throw new ArgumentNullException(nameof(first));
            }

            if (second == null)
            {
                throw new ArgumentNullException(nameof(second));
            }

            List<IndexHit> merged = new List<IndexHit>(first.Count + second.Count);
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            AppendUnseen(merged, seen, first);
            AppendUnseen(merged, seen, second);
            return merged;
        }

        /// <summary>
        /// Orders <paramref name="hits"/> so that a row the ordering cannot rank can never
        /// hold a place ahead of one it can: ranked rows first, by their key descending, then
        /// the unranked ones in the order they arrived.
        /// <para>
        /// The sort is STABLE in both halves - ties fall back to arrival position - so where
        /// the provider already ordered correctly this reproduces its answer exactly rather
        /// than inventing a second opinion about equal keys.
        /// </para>
        /// </summary>
        public static IReadOnlyList<IndexHit> RankableFirst(IReadOnlyList<IndexHit> hits, IndexOrder order)
        {
            if (hits == null)
            {
                throw new ArgumentNullException(nameof(hits));
            }

            List<int> ranked = new List<int>(hits.Count);
            List<IndexHit> unranked = new List<IndexHit>();
            for (int i = 0; i < hits.Count; i++)
            {
                if (HasOrderKey(hits[i], order))
                {
                    ranked.Add(i);
                }
                else
                {
                    unranked.Add(hits[i]);
                }
            }

            ranked.Sort((left, right) =>
            {
                int byKey = CompareKeyDescending(hits[left], hits[right], order);
                return byKey != 0 ? byKey : left.CompareTo(right);
            });

            List<IndexHit> ordered = new List<IndexHit>(hits.Count);
            for (int i = 0; i < ranked.Count; i++)
            {
                ordered.Add(hits[ranked[i]]);
            }

            ordered.AddRange(unranked);
            return ordered;
        }

        private static void AppendUnseen(List<IndexHit> target, HashSet<string> seen, IReadOnlyList<IndexHit> source)
        {
            for (int i = 0; i < source.Count; i++)
            {
                IndexHit hit = source[i];
                if (hit == null)
                {
                    continue;
                }

                // A row with no URL cannot be de-duplicated, and cannot be admitted either
                // (IndexRowFilter refuses it), so keeping it costs nothing while dropping it
                // here would be a second admission rule in a place nobody would look for one.
                if (string.IsNullOrEmpty(hit.ItemUrl) || seen.Add(hit.ItemUrl))
                {
                    target.Add(hit);
                }
            }
        }

        private static int CompareKeyDescending(IndexHit left, IndexHit right, IndexOrder order)
        {
            switch (order)
            {
                case IndexOrder.SizeDescending:
                    return right.SizeBytes!.Value.CompareTo(left.SizeBytes!.Value);
                case IndexOrder.DateReceivedDescending:
                    return DateTime.Compare(right.DateReceivedUtc!.Value, left.DateReceivedUtc!.Value);
                default:
                    throw new ArgumentException("Unknown IndexOrder value.", nameof(order));
            }
        }
    }
}
