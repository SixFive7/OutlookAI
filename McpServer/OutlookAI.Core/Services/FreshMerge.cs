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

        /// <summary>
        /// Coverage hole: the per-folder item cap truncated a folder's window, newest-first,
        /// so what is missing is the OLDEST part of that folder's freshness window.
        /// <para>
        /// Raised only for folders whose table Outlook actually sorted. The rest raise
        /// <see cref="GapItemCapUnsorted"/>, and the split is what keeps this code's advice
        /// sentence true of every folder it names (see
        /// <see cref="SortedItemCappedFolders"/>).
        /// </para>
        /// </summary>
        public const string GapItemCap = "item_cap";

        /// <summary>
        /// Coverage hole: the per-folder item cap truncated a folder's window whose table
        /// Outlook would not sort, so the cap kept an ARBITRARY slice of the window and what
        /// is missing is unknown (gap H2).
        /// <para>
        /// Its own code rather than a qualifier on <see cref="GapItemCap"/>, because the two
        /// lead somewhere different. <c>item_cap</c> says the oldest end of the window is
        /// missing: an agent can narrow with <c>after</c>, page the bound, or tell the user
        /// which mail to expect to be absent. This one says none of that is available - the
        /// hole is a random 200-item slice - so the only remedies are to make the window
        /// small enough to fit under the cap or to read the folder with
        /// <c>exhaustive: true</c>.
        /// </para>
        /// <para>
        /// The failure it reports was previously swallowed with the note that an unsorted
        /// sweep still works. The SWEEP does; the sentence attached to the cap does not, and
        /// a sentence an agent relays to a human as fact is the thing this whole contract
        /// exists to keep honest.
        /// </para>
        /// </summary>
        public const string GapItemCapUnsorted = "item_cap_unsorted";

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

        /// <summary>
        /// Coverage hole: a tier reached rows it could not turn into items - the row carried
        /// no usable EntryID, Outlook refused to open it, or its item class could not be
        /// read. Each one is a mail that matched the tier's own filter and is missing from
        /// the answer anyway.
        /// <para>
        /// The specific defect (gap H1): a swept row that failed to open was skipped AND did
        /// not count toward the per-folder cap, so a folder where EVERY row failed came back
        /// as <c>SweepOutcome.Complete</c> with zero items and was counted in
        /// <c>foldersSwept</c> - a folder reporting full coverage having produced nothing.
        /// The folder-level counters cannot express this, because the folder really was
        /// enumerated; only a row-level counter can.
        /// </para>
        /// <para>
        /// Shared with the exhaustive scan (<see cref="ScanGapRowsUnreadable"/>, gap F5),
        /// which loses rows the same three ways. One token, because it is one fact about one
        /// kind of loss and an agent should not have to learn a second name for it per tier.
        /// </para>
        /// </summary>
        public const string GapRowsUnreadable = "rows_unreadable";

        /// <summary>
        /// Coverage hole: a swept item was dropped because a property one of the caller's
        /// own filters needs could not be read from it - <c>unread_only</c> needs
        /// <c>IsRead</c>, <c>has_attachments</c> needs <c>HasAttachments</c>,
        /// <c>before</c>/<c>after</c> need <c>ReceivedTime</c> (gap I1).
        /// <para>
        /// DROPPING IS THE DELIBERATE CHOICE, and this code is the other half of it. A
        /// filter the caller asked for has to be honoured: admitting an item that cannot be
        /// shown to match would quietly corrupt the answer in the opposite direction, and an
        /// agent that asked for unread mail would relay read mail as unread. So the item
        /// stays out - and the count of what was dropped, plus which filters could not be
        /// evaluated, goes in the payload, because "we dropped some of your results" is
        /// exactly the class of fact this whole contract exists to stop being silent.
        /// </para>
        /// <para>
        /// The remedy is the caller's, which is why the filter NAMES travel with the code
        /// (<see cref="SweepInfo.FiltersUnevaluated"/>): re-running without the filter in
        /// question returns those items, and the caller can judge them.
        /// </para>
        /// <para>
        /// One honest caveat, carried here rather than settled here: a null
        /// <c>ReceivedTime</c> is ambiguous. Outlook reports 4501-01-01 for "no value" and
        /// the COM snapshot maps that to null exactly as it maps a failed read
        /// (<c>OutlookComSession.TryGetDateTime</c>), so a date filter dropping such an item
        /// cannot say which of the two happened. The advice sentence therefore says "no
        /// usable value", never "the read failed". Known to occur for real on drafts (the
        /// Phase-1 completeness oracle had to exclude them on both sides for that reason);
        /// whether Sent Items carry it too is the open question this deliberately leaves
        /// open - and reporting the drop is what will eventually answer it, since the code
        /// names the filter and counts the items each time it happens.
        /// </para>
        /// </summary>
        public const string GapFilterUnreadable = "filter_unreadable";

        /// <summary>
        /// Coverage hole: the live conversation walk stopped at the requested member cap,
        /// so it did not see the whole conversation. Unlike a search's <c>top</c>, which
        /// caps a date-SORTED match set, the walk reads the conversation table in Outlook's
        /// own order and stops - so the member it did not reach may be the newest one, which
        /// is exactly the member the live tier exists to find. Remedy: raise <c>top</c>.
        /// </summary>
        public const string ThreadGapMemberCap = "member_cap";

        /// <summary>
        /// Coverage hole: the conversation has members in a store the live walk did not
        /// cover. Outlook's Conversation object walks the store of the item it was opened
        /// from, so a conversation spanning two accounts gets a live check of one of them
        /// and index coverage of the rest - fine for anything already indexed, blind to a
        /// reply that arrived in the OTHER store moments ago. Remedy: pass an <c>id</c> from
        /// the store in question.
        /// </summary>
        public const string ThreadGapUnwalkedStore = "unwalked_store";

        /// <summary>
        /// Coverage hole: the profile holds a store the local index has no mail for, and the
        /// live walk did not cover it either - so whether this conversation reaches into it
        /// cannot be established by any tier this lookup has.
        /// <para>
        /// THE SILENT HALF OF GAP C4, and it is silent for a reason worth stating.
        /// <see cref="ThreadGapUnwalkedStore"/> is raised from the stores the INDEX ROWS for
        /// this conversation name, so it can only ever fire for a store the index knows.
        /// Point it at the unindexed-PST profile the A-group work is about and it says
        /// nothing at all: the index has no rows from that store to name, the walk covered
        /// one other store, and half a conversation can be missing with the payload reading
        /// <c>freshness: "live"</c>.
        /// </para>
        /// <para>
        /// It is deliberately a WEAKER claim than its sibling. <c>unwalked_store</c> says the
        /// conversation demonstrably has members elsewhere; this one says the question could
        /// not be asked, which is a different thing to relay and the reason it is a second
        /// code rather than a widening of the first. Both have the same remedy - call
        /// <c>thread</c> again with an <c>id</c> from the store in question - because both
        /// are answered by walking that store's own conversation graph.
        /// </para>
        /// <para>
        /// The verdict per store is the same pure <c>MailService.StoresMissingFromIndex</c>
        /// that closed A3 and A1's residue: only a PROBE's "no" counts, absence from the
        /// discovery catalog is never evidence, and a store that could not be settled is
        /// reported neither way. So this code means "there is a store here neither tier can
        /// see", never "a store we could not check".
        /// </para>
        /// </summary>
        public const string ThreadGapUnindexedStore = "unindexed_store";

        /// <summary>
        /// Exhaustive-scan coverage hole: the scan's wall-clock budget stopped it, so the
        /// folders it had not reached yet were never opened. Deliberately the SAME token the
        /// freshness sweep uses (<see cref="GapTimeBudget"/>) - the hole is the same hole,
        /// and a caller should not have to learn one vocabulary per tier.
        /// </summary>
        public const string ScanGapTimeBudget = GapTimeBudget;

        /// <summary>
        /// Exhaustive-scan coverage hole: folders whose table neither filter engine would
        /// open. Same token as the sweep's (<see cref="GapFoldersSkipped"/>), same reason.
        /// </summary>
        public const string ScanGapFoldersSkipped = GapFoldersSkipped;

        /// <summary>
        /// Exhaustive-scan coverage hole: a subtree was refused for sitting deeper than the
        /// folder-walk depth guard, so its folders were never opened (gap F4). The SAME token
        /// the sweep's walk uses (<see cref="GapDepthLimit"/>) - one bound, one name, one
        /// remedy - because it is literally the same guard at the same number.
        /// <para>
        /// New in the sense that this walk had no bound to report: it recursed without one,
        /// so a cyclic folder graph ended the COM host process rather than truncating an
        /// answer. Adding the guard without this code would have replaced a crash with a
        /// silent hole, which in the mode chosen FOR completeness is the worse of the two.
        /// </para>
        /// </summary>
        public const string ScanGapDepthLimit = GapDepthLimit;

        /// <summary>
        /// Exhaustive-scan coverage hole: the result cap stopped the walk partway through
        /// the folder tree.
        /// <para>
        /// Its own token rather than the sweep's <see cref="GapItemCap"/>, because the two
        /// caps lose different things. The sweep's per-folder cap truncates ONE folder's
        /// newest-first window, and the folders after it are still swept; the exhaustive cap
        /// stops the whole walk in the order the tree happened to come (gap F2), so what is
        /// missing is whole folders that were never opened, and raising <c>top</c> is capped
        /// at 100 rather than pageable.
        /// </para>
        /// </summary>
        public const string ScanGapResultCap = "result_cap";

        /// <summary>
        /// Exhaustive-scan coverage hole: rows the scan could not turn into items (gap F5).
        /// The same token as the sweep's <see cref="GapRowsUnreadable"/> - one fact, one
        /// name.
        /// <para>
        /// Counts only the FAILURES: a row with no usable EntryID, one Outlook would not
        /// open, one whose <c>Class</c> would not read. A row dropped because its class is
        /// not <c>IPM.Note</c> is a deliberate filter, not a failure, and is counted apart
        /// in <see cref="ExhaustiveInfo.RowsDropped"/> without raising anything - the same
        /// distinction <see cref="SweepInfo.FoldersAbsent"/> draws one level up, and for the
        /// same reason: a flag that cries wolf is worse than no flag.
        /// </para>
        /// </summary>
        public const string ScanGapRowsUnreadable = GapRowsUnreadable;

        /// <summary>
        /// Exhaustive-scan coverage hole: a scanned item was dropped because a property one
        /// of the caller's own filters needs could not be read - the same hole gap I1 names
        /// in the sweep, in the tier that post-filters the same snapshots, so the same token.
        /// <para>
        /// Only <c>unread_only</c> and <c>has_attachments</c> can raise it here: this mode's
        /// date bounds go into the DASL filter rather than a post-filter, so there is no
        /// <c>before</c>/<c>after</c> read to fail.
        /// </para>
        /// </summary>
        public const string ScanGapFilterUnreadable = GapFilterUnreadable;

        /// <summary>
        /// Exhaustive-scan coverage hole: the scan's result cap stopped the walk, and the
        /// request's <c>from</c> / <c>unread_only</c> / <c>has_attachments</c> filters were
        /// then applied to what the cap had already kept (gap F3).
        /// <para>
        /// The cap counts items the DASL filter matched, before any of those three run, so
        /// the returned list can be a handful of rows while the store holds thousands of
        /// matches - and <c>truncated: true</c> beside two results reads as "a couple more
        /// exist" rather than "the scan stopped after <c>top</c> candidates and most of them
        /// were then discarded". Neither number on its own says which happened; only the
        /// pairing does.
        /// </para>
        /// <para>
        /// Raised ONLY together with <see cref="ScanGapResultCap"/>, which is what makes it
        /// free of false alarms: with no cap the post-filter saw the whole matched set and
        /// removed exactly what the caller asked to remove, and with no such filter the cap
        /// truncated a list nothing later thinned. So this code adds no degradation of its
        /// own - the scan is already <c>partial</c> when it fires - it adds the explanation
        /// the result cap's own sentence could not give.
        /// </para>
        /// <para>
        /// Distinct from <see cref="ScanGapFilterUnreadable"/>, which is about items those
        /// same filters could not be EVALUATED on. Here they were evaluated and worked; the
        /// hole is that they ran over a truncated input.
        /// </para>
        /// </summary>
        public const string ScanGapPostCapFilter = "post_cap_filter";

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
        /// The item-capped folders whose window really was cut newest-first: everything in
        /// <see cref="SweepInfo.ItemCappedFolders"/> that is not in
        /// <see cref="SweepInfo.ItemCappedFoldersUnsorted"/> (gap H2).
        /// <para>
        /// One split read by both renderings - the gap codes and the advice sentences - so
        /// neither can name a folder the other does not. Written as a difference rather than
        /// as a second list from the COM layer because the sweep result already carries the
        /// whole capped set and its exception; deriving the complement here keeps the two
        /// lists from drifting into disagreement about which folders were capped at all.
        /// </para>
        /// <para>
        /// Matched ordinally: both lists are built from the same
        /// <c>store/folder</c> labels in the same walk, so a case-insensitive compare would
        /// buy nothing and could collapse two folders a mailbox genuinely distinguishes.
        /// </para>
        /// </summary>
        public static IReadOnlyList<string> SortedItemCappedFolders(SweepInfo sweep)
        {
            if (sweep == null)
            {
                throw new ArgumentNullException(nameof(sweep));
            }

            IReadOnlyList<string>? capped = sweep.ItemCappedFolders;
            if (capped == null || capped.Count == 0)
            {
                return Array.Empty<string>();
            }

            IReadOnlyList<string>? unsorted = sweep.ItemCappedFoldersUnsorted;
            if (unsorted == null || unsorted.Count == 0)
            {
                return capped;
            }

            HashSet<string> arbitrary = new HashSet<string>(unsorted, StringComparer.Ordinal);
            List<string> sorted = new List<string>(capped.Count);
            foreach (string folder in capped)
            {
                if (!arbitrary.Contains(folder))
                {
                    sorted.Add(folder);
                }
            }

            return sorted;
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
            // The one place the capped set is split, so the code list and the advice
            // sentences cannot end up describing different folders (gap H2).

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

            // Two codes over one cap, because "newest-first" is true of one set of folders
            // and false of the other, and a single code would have to carry both (gap H2).
            if (SortedItemCappedFolders(sweep).Count > 0)
            {
                gaps.Add(GapItemCap);
            }

            if (sweep.ItemCappedFoldersUnsorted != null && sweep.ItemCappedFoldersUnsorted.Count > 0)
            {
                gaps.Add(GapItemCapUnsorted);
            }

            // Row- and item-level holes last: they are the narrowest, and they sit INSIDE
            // folders the counters above already report as swept. That is exactly why they
            // need their own codes - no folder counter can express "this folder was read
            // and some of its mail was still lost".
            if (sweep.RowsUnreadable > 0)
            {
                gaps.Add(GapRowsUnreadable);
            }

            if (sweep.ItemsFilterUnreadable > 0)
            {
                gaps.Add(GapFilterUnreadable);
            }

            return gaps.Count == 0 ? null : gaps;
        }

        /// <summary>
        /// Every way an EXHAUSTIVE scan that ran can still have covered less than its scope -
        /// the third tier's <see cref="DescribeCoverageGaps"/>, pure for the same reason and
        /// pinned in T1 the same way.
        /// <para>
        /// The scan always runs (it IS the live check), so unlike the sweep and the thread
        /// walk there is no "did it run" gate here and null means one thing only: it covered
        /// what it was asked.
        /// </para>
        /// <para>
        /// Order is severity-first, and it is also the order the advice sentences come out
        /// in: the two bounds that stopped the WALK (so whole folders were never opened),
        /// then the folders it reached and could not filter, then the rows it opened and
        /// lost. <see cref="ExhaustiveInfo.RowsDropped"/> raises nothing on its own - a row
        /// dropped for its item class is the mode working as designed, not a hole.
        /// </para>
        /// </summary>
        public static IReadOnlyList<string>? DescribeExhaustiveCoverageGaps(ExhaustiveInfo exhaustive)
        {
            if (exhaustive == null)
            {
                throw new ArgumentNullException(nameof(exhaustive));
            }

            List<string> gaps = new List<string>();
            if (exhaustive.TimedOut)
            {
                gaps.Add(ScanGapTimeBudget);
            }

            if (exhaustive.Truncated)
            {
                gaps.Add(ScanGapResultCap);
            }

            if (exhaustive.FoldersSkipped > 0)
            {
                gaps.Add(ScanGapFoldersSkipped);
            }

            if (exhaustive.DepthLimitReached)
            {
                gaps.Add(ScanGapDepthLimit);
            }

            if (exhaustive.RowsUnreadable > 0)
            {
                gaps.Add(ScanGapRowsUnreadable);
            }

            if (exhaustive.ItemsFilterUnreadable > 0)
            {
                gaps.Add(ScanGapFilterUnreadable);
            }

            // Last, and conditional on the cap having actually fired: it explains a hole the
            // codes above already declared rather than declaring one of its own (gap F3).
            if (exhaustive.Truncated
                && exhaustive.PostCapFilters != null
                && exhaustive.PostCapFilters.Count > 0)
            {
                gaps.Add(ScanGapPostCapFilter);
            }

            return gaps.Count == 0 ? null : gaps;
        }

        /// <summary>
        /// The request filters an exhaustive scan can only apply AFTER its own result cap,
        /// in the order they appear on the request, or null when the request passed none
        /// (gap F3).
        /// <para>
        /// These three are read off the item snapshots the scan returns, so the scan cannot
        /// push them into its DASL filter and cannot count toward its cap only the items
        /// that pass them. Contrast <c>before</c>/<c>after</c>, which DO go into the filter
        /// and therefore bound the scan itself; they are absent here for that reason and not
        /// by omission.
        /// </para>
        /// <para>
        /// Pure and T1-pinned, so the names in the payload and the names in the advice
        /// sentence are one list. They are the caller's own parameter names, which is what
        /// makes the sentence actionable: the remedy is to drop one of them and let the cap
        /// count the matches instead.
        /// </para>
        /// </summary>
        public static IReadOnlyList<string>? PostCapFilters(
            bool hasFrom, bool hasUnreadOnly, bool hasAttachmentsFilter)
        {
            List<string> names = new List<string>(3);
            if (hasFrom)
            {
                names.Add("from");
            }

            if (hasUnreadOnly)
            {
                names.Add("unread_only");
            }

            if (hasAttachmentsFilter)
            {
                names.Add("has_attachments");
            }

            return names.Count == 0 ? null : names;
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
        /// Every way a conversation walk that RAN can still have covered less than the
        /// conversation - the <c>thread</c> analogue of <see cref="DescribeCoverageGaps"/>,
        /// and pure for the same reason.
        /// <para>
        /// Returns null for a walk that covered everything and for one that never ran: "did
        /// not run" is <see cref="FreshnessIndexOnly"/>, a different state with a different
        /// remedy (pass <c>id</c>), and folding it in here would blur the two exactly as it
        /// would on a sweep.
        /// </para>
        /// <para>
        /// <paramref name="indexHitStores"/> is the set of stores the INDEX rows for this
        /// conversation came from, which is the only evidence available that the
        /// conversation reaches past the store Outlook walked. A store the walk covered
        /// raises nothing; a store neither tier could name (null or blank) raises nothing
        /// either, because "unknown" is not "uncovered" and a gap code that fires on missing
        /// metadata is a false alarm on every profile that has any.
        /// </para>
        /// </summary>
        public static IReadOnlyList<string>? DescribeThreadCoverageGaps(
            ThreadLiveInfo live,
            IReadOnlyCollection<string?>? indexHitStores)
        {
            if (live == null)
            {
                throw new ArgumentNullException(nameof(live));
            }

            if (!live.Performed)
            {
                return null;
            }

            List<string> gaps = new List<string>();
            if (live.MembersWalked > 0 && !string.IsNullOrEmpty(live.AnchorStore) && indexHitStores != null)
            {
                foreach (string? store in indexHitStores)
                {
                    if (!string.IsNullOrEmpty(store)
                        && !string.Equals(store, live.AnchorStore, StringComparison.OrdinalIgnoreCase))
                    {
                        gaps.Add(ThreadGapUnwalkedStore);
                        break;
                    }
                }
            }

            // Read off the block rather than recomputed here, exactly as the sweep's
            // no_index_frontier is read off SweepInfo.IndexFrontierMissing: establishing it
            // needs Outlook's store list and an index probe per store, neither of which
            // belongs in a pure classifier. The RULE that turns those into this list is pure
            // and lives in UnwalkedUnindexedStores.
            if (live.StoresWithoutIndex != null && live.StoresWithoutIndex.Count > 0)
            {
                gaps.Add(ThreadGapUnindexedStore);
            }

            if (live.MemberCapReached)
            {
                gaps.Add(ThreadGapMemberCap);
            }

            return gaps.Count == 0 ? null : gaps;
        }

        /// <summary>
        /// Which of the stores the index holds no mail for were left OUTSIDE this
        /// conversation walk - the rule behind <see cref="ThreadGapUnindexedStore"/>, pure so
        /// T1 pins it without a profile, an index or a mailbox.
        /// <para>
        /// The anchor's own store is excluded, and that exclusion is the whole substance of
        /// the rule: Outlook enumerated the conversation THERE, member by member, so its
        /// coverage of that store is complete whatever the index does or does not hold. An
        /// unindexed store that happens to be the one walked is therefore not a hole, and
        /// reporting it would fire this code on every single-store PST profile - the
        /// cry-wolf failure that makes a completeness flag worthless.
        /// </para>
        /// <para>
        /// Nothing is claimed when the walk did not run, returned no members, or could not
        /// name the store it covered. All three leave <c>AnchorStore</c> unusable, so the
        /// exclusion above cannot be applied - and a list built without it would name the
        /// very store that WAS walked, which is worse than saying nothing. "Did not run" has
        /// its own state (<see cref="FreshnessIndexOnly"/>) with its own remedy, and a
        /// zero-member walk is judged the same way <see cref="DescribeThreadCoverageGaps"/>
        /// already judges it for <see cref="ThreadGapUnwalkedStore"/>.
        /// </para>
        /// </summary>
        public static IReadOnlyList<string> UnwalkedUnindexedStores(
            ThreadLiveInfo live,
            IReadOnlyList<string>? storesWithoutIndex)
        {
            if (live == null)
            {
                throw new ArgumentNullException(nameof(live));
            }

            if (!live.Performed
                || live.MembersWalked <= 0
                || string.IsNullOrEmpty(live.AnchorStore)
                || storesWithoutIndex == null
                || storesWithoutIndex.Count == 0)
            {
                return Array.Empty<string>();
            }

            List<string> unwalked = new List<string>(storesWithoutIndex.Count);
            foreach (string store in storesWithoutIndex)
            {
                if (string.IsNullOrEmpty(store)
                    || string.Equals(store, live.AnchorStore, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                unwalked.Add(store);
            }

            return unwalked;
        }

        /// <summary>
        /// The freshness verdict for one <c>thread</c> lookup, from its live block alone -
        /// the same three values, with the same meanings, that a search carries.
        /// <para>
        /// <c>thread</c> had none of this: no <c>degraded</c>, no <c>freshness</c>, no
        /// staleness, and with a single index row for the conversation the COM walk never
        /// ran at all. So the tool that promises the FULL conversation was the only one with
        /// no way to say it had returned part of one, and a reply newer than the index
        /// frontier was absent from a payload that read as complete.
        /// </para>
        /// <para>
        /// <see cref="FreshnessIndexOnly"/> is reachable here and means what it always
        /// means - the live check did not run - but it has a remedy the sweep's version does
        /// not: the walk needs a concrete item to start from, so a caller who passed only
        /// <c>conversation_id</c> can fix it by passing <c>id</c> as well.
        /// </para>
        /// </summary>
        public static string ClassifyThreadFreshness(
            ThreadLiveInfo live,
            IReadOnlyCollection<string?>? indexHitStores)
        {
            if (live == null)
            {
                throw new ArgumentNullException(nameof(live));
            }

            if (!live.Performed)
            {
                return FreshnessIndexOnly;
            }

            // Recomputed from the same inputs rather than read off the block, so the verdict
            // and the codes cannot disagree about the same lookup - the shape ClassifyFreshness
            // already uses for the sweep.
            IReadOnlyList<string>? gaps = DescribeThreadCoverageGaps(live, indexHitStores);
            return gaps == null || gaps.Count == 0 ? FreshnessLive : FreshnessPartial;
        }

        /// <summary>
        /// The freshness verdict for an EXHAUSTIVE search - the same three-value contract,
        /// over the one tier that mode has.
        /// <para>
        /// It set neither <c>freshness</c> nor <c>degraded</c> at all (gap F1), so a scan
        /// that timed out after four folders and skipped nine more came back looking
        /// identical, on the exact two fields the search description teaches agents to read
        /// and relay, to one that covered everything. The facts were already in
        /// <c>exhaustive.*</c>; the two flags an agent branches on were the ones missing.
        /// </para>
        /// <para>
        /// WHAT THE THREE VALUES MEAN HERE, and why no fourth is added. They answer "did the
        /// LIVE check run, and did it cover its scope". An exhaustive scan IS the live check
        /// - it reads Outlook's folders directly and consults no index - so it has run by
        /// the time there is a result to classify, and <see cref="FreshnessIndexOnly"/> is
        /// therefore unreachable rather than merely unused: that value means "nothing was
        /// checked live", which is the precise inverse of this mode. What remains is whether
        /// the scan covered what it was asked, which is <see cref="FreshnessLive"/> against
        /// <see cref="FreshnessPartial"/>, decided by the counters the scan already reports.
        /// A fourth value such as "com-only" would say something true about the METHOD and
        /// nothing about the COVERAGE, and every caller already taught to treat
        /// <c>live</c> as "complete" would stop matching a complete answer.
        /// </para>
        /// <para>
        /// TRUNCATION COUNTS HERE AND DOES NOT IN THE MERGED PATH, and that asymmetry is the
        /// point rather than an oversight. An indexed search's <c>truncated</c> caps a
        /// SORTED complete match set - the provider ordered every match by date and the
        /// caller got the newest <c>top</c> of them, so raising <c>top</c> or moving the
        /// <c>before</c> bound pages through the rest. The exhaustive scan stops walking the
        /// moment it has <c>top</c> items, in whatever order the folder tree happened to
        /// come (gap F2), so its truncation means folders were never opened at all. That is
        /// a coverage hole in the plainest sense, and it is not pageable.
        /// </para>
        /// </summary>
        public static string ClassifyExhaustiveFreshness(ExhaustiveInfo? exhaustive)
        {
            if (exhaustive == null)
            {
                throw new ArgumentNullException(nameof(exhaustive));
            }

            // Recomputed from the gap codes rather than restating their conditions, so the
            // verdict and the codes cannot disagree about the same scan - the shape
            // ClassifyFreshness and ClassifyThreadFreshness already use for their tiers.
            // Before this, adding a hole meant remembering to widen a boolean expression
            // here as well, which is precisely how a code ships next to freshness "live".
            IReadOnlyList<string>? gaps = DescribeExhaustiveCoverageGaps(exhaustive);
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
        /// Whether this query's terms could match INSIDE an attachment, which is a thing
        /// only the index tier can do (gap B2).
        /// <para>
        /// The index tier's body scope is <c>System.Search.Contents</c>, which is body plus
        /// attachment content. The freshness sweep reads <c>MailItem.Subject</c> and
        /// <c>MailItem.Body</c> through COM and never opens an attachment, and the
        /// exhaustive scan reads <c>urn:schemas:httpmail:textdescription</c>. So for any
        /// query whose terms are matched against the body scope, attachment text is covered
        /// by exactly one of the three tiers - and the one that does not cover it is the one
        /// responsible for everything newer than the index frontier.
        /// </para>
        /// <para>
        /// The attachment-ONLY case was already refused and reported
        /// (<see cref="AttachmentContentNotSweepable"/>). This is the DEFAULT case, which
        /// said nothing at all: the sweep contributes its message rows, the merge succeeds,
        /// and the answer reports <c>freshness: "live"</c> while a term sitting inside a
        /// PDF that arrived four minutes ago is in neither tier. <c>live</c> is the reading
        /// an agent trusts most, so the shortfall belongs in the payload.
        /// </para>
        /// <para>
        /// It depends on <paramref name="searchIn"/> and on there being terms at all, and
        /// deliberately not on the attachment-hit flags: with attachment rows INCLUDED the
        /// index can return the attachment itself, with them EXCLUDED it can still match a
        /// message row on its attachment's text, and with attachment rows ONLY the sweep
        /// does not run. All three leave the sweep unable to see attachment text; only the
        /// third already says so.
        /// </para>
        /// </summary>
        public static bool AttachmentTextMatchable(SearchIn searchIn, bool hasTerms)
        {
            return hasTerms && searchIn != SearchIn.SubjectOnly;
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
