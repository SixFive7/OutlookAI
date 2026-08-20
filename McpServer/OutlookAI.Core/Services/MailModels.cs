using System;
using System.Collections.Generic;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;

namespace OutlookAI.Core.Services
{
    /// <summary>Parameters for one search call (mirrors the MCP tool arguments).</summary>
    public sealed class SearchRequest
    {
        /// <summary>Free-text query; whitespace-separated terms are ANDed. Optional.</summary>
        public string? Query { get; set; }

        /// <summary>
        /// Which properties <see cref="Query"/> terms must appear in: subject OR
        /// body/attachment content (default), subject only, or body only. Honored
        /// identically by all three tiers - index, freshness sweep and exhaustive scan
        /// (D40/SF-6). Sender matching is never a search_in scope; that is <see cref="From"/>.
        /// </summary>
        public SearchIn SearchIn { get; set; } = SearchInValues.Default;

        /// <summary>
        /// True = folder/date-bounded COM scan that BYPASSES the index (correctness
        /// beats speed; also works when the SystemIndex is broken). Requires a store
        /// plus a bound (folder or after date) to avoid multi-minute scans. False
        /// (default) = index search + freshness gap-sweep, merged and deduped (D19/D34).
        /// </summary>
        public bool Exhaustive { get; set; }

        /// <summary>
        /// Test/diagnostic escape hatch: skip the freshness gap-sweep and return index
        /// results only. NOT exposed on the MCP tool - since D34 the sweep is always on
        /// for agents (with graceful degradation when it cannot run).
        /// </summary>
        public bool IndexOnly { get; set; }

        /// <summary>Store display name to scope to (as returned by list_accounts).</summary>
        public string? Store { get; set; }

        /// <summary>Store-relative folder path ('/'-separated) to scope to; requires <see cref="Store"/>.</summary>
        public string? Folder { get; set; }

        /// <summary>
        /// Whether <see cref="Folder"/> covers its SUBFOLDERS too. Default true (user
        /// decision, soak fix 15) - it matches what the index tier always did and removes
        /// the old asymmetry where an exhaustive folder scan silently covered less ground
        /// than the same folder search.
        /// <para>
        /// Honored by all three tiers: index (recursive SCOPE vs SCOPE + folder-path
        /// equality), freshness sweep (subtree walk vs single folder - and the flag is part
        /// of the sweep cache key) and exhaustive scan (the ScanFolderTree recurse flag).
        /// Ignored without a <see cref="Folder"/>: a whole store is recursive either way.
        /// </para>
        /// <para>
        /// ⚠ Delegate mailboxes are indexed FLAT (no folder nesting), so the index tier
        /// covers a delegate subtree by matching each contained folder NAME. When that set
        /// cannot be built or is too large the query widens to the whole delegate mailbox
        /// and says so in advice - it is never narrowed silently.
        /// </para>
        /// </summary>
        public bool IncludeSubfolders { get; set; } = true;

        /// <summary>Sender filter (index-backed per-column CONTAINS).</summary>
        public string? From { get; set; }

        /// <summary>Recipient filter (To or Cc, index-backed per-column CONTAINS).</summary>
        public string? To { get; set; }

        /// <summary>Only items received at or after this UTC instant.</summary>
        public DateTime? AfterUtc { get; set; }

        /// <summary>Only items received before this UTC instant.</summary>
        public DateTime? BeforeUtc { get; set; }

        /// <summary>True = only unread mail.</summary>
        public bool? UnreadOnly { get; set; }

        /// <summary>Filter on attachment presence.</summary>
        public bool? HasAttachments { get; set; }

        /// <summary>
        /// Include indexed attachment-content entries of ANY attachment type - documents,
        /// images, embedded messages, invites, media (soak fix 16). Default true.
        /// </summary>
        public bool IncludeAttachmentHits { get; set; } = true;

        /// <summary>ONLY attachment-content entries, any type. Overrides <see cref="IncludeAttachmentHits"/>.</summary>
        public bool AttachmentHitsOnly { get; set; }

        /// <summary>Order results by size instead of date (big-mail discovery; index path only).</summary>
        public bool OrderBySizeDescending { get; set; }

        /// <summary>Maximum hits returned (compact payloads - T1-pinned caps in <see cref="MailService"/>).</summary>
        public int Top { get; set; } = MailService.SearchTopDefault;

        /// <summary>Snippet length per hit (0 disables snippets).</summary>
        public int SnippetChars { get; set; } = MailService.SnippetCharsDefault;

        /// <summary>
        /// Continuation handle from a previous exhaustive scan's
        /// <c>exhaustive.nextToken</c>: continue that walk instead of starting a new one
        /// (F2). Opaque, and only meaningful with <see cref="Exhaustive"/>.
        /// <para>
        /// Every other argument must arrive UNCHANGED; a resume whose question differs is
        /// refused rather than silently honoured or silently ignored, because both of those
        /// answer a different question under a claim of continuity.
        /// <see cref="Top"/> and <see cref="SnippetChars"/> are the exceptions and may
        /// differ per page: they shape the presentation of one page, not the question.
        /// </para>
        /// </summary>
        public string? ResumeToken { get; set; }
    }

    /// <summary>
    /// How far a paged exhaustive scan has got, in fields a caller can act on WITHOUT the
    /// token (F2). Present exactly when the scan stopped early.
    /// <para>
    /// It is what makes the continuation survivable: the token lives in one server process,
    /// so a server restart loses it - and a caller holding this block continues by hand with
    /// <c>folder</c> and <c>before</c>, which are parameters <c>search</c> already has. That
    /// is why the token itself can afford to be ten characters of context instead of a
    /// kilobyte of self-describing state.
    /// </para>
    /// </summary>
    public sealed class ScanPositionInfo
    {
        /// <summary>Mail folders the chain has finished, across every page so far.</summary>
        public int FoldersDone { get; set; }

        /// <summary>Mail folders in scope, counted by this page's own ordered enumeration.</summary>
        public int FoldersTotal { get; set; }

        /// <summary>Store-relative path of the folder the next page starts in.</summary>
        public string? ResumeFolder { get; set; }

        /// <summary>True when the next page resumes PART WAY THROUGH that folder rather than at its top.</summary>
        public bool ResumeWithinFolder { get; set; }

        /// <summary>
        /// The inclusive received-date bound the next page restricts on, when the folder's
        /// table sorted. Absent on the other two rungs, which have no date to resume from.
        /// </summary>
        public DateTime? ResumeCursorUtc { get; set; }

        /// <summary>
        /// Which rung of the resumption ladder the next page will use: <c>date</c> (the
        /// folder sorted, so resumption is a narrower query), <c>ordinal</c> (the sort was
        /// refused, so it is a verified row skip) or <c>restart</c> (the folder is re-read
        /// from the top with duplicate suppression).
        /// <para>
        /// It is a cost signal and it is also evidence: <c>date</c> on a folder means
        /// <c>Table.Sort</c> works there, which is the open question behind the freshness
        /// sweep's own item cap.
        /// </para>
        /// </summary>
        public string? ResumeTier { get; set; }

        /// <summary>Pages of this chain served so far, this one included.</summary>
        public int Page { get; set; }
    }

    /// <summary>One agent-facing hit: compact triage payload (v3.MD sections 8/12).</summary>
    public sealed class HitSummary
    {
        /// <summary>Opaque hit id for read/save_attachment/thread ("h1", "h2", ...). Cached per server process.</summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>
        /// "index" (SystemIndex row), "live" (freshness gap-sweep result, D19), "exhaustive"
        /// (COM scan) or "com" (thread's live conversation walk). Everything but "index"
        /// came from Outlook directly, so it carries no index-only field - notably
        /// <see cref="ConversationId"/>, which the COM snapshot does not read.
        /// </summary>
        public string Source { get; set; } = "index";

        /// <summary>Subject.</summary>
        public string? Subject { get; set; }

        /// <summary>Sender display name.</summary>
        public string? FromName { get; set; }

        /// <summary>Sender address.</summary>
        public string? FromAddress { get; set; }

        /// <summary>Received timestamp, UTC.</summary>
        public DateTime? ReceivedUtc { get; set; }

        /// <summary>Store display name (delegate hits already routed to the delegate store).</summary>
        public string? Store { get; set; }

        /// <summary>Folder path within the store ('/'-separated; leaf name for live hits).</summary>
        public string? Folder { get; set; }

        /// <summary>"inbox"/"sent" for live (gap-sweep) hits; absent for index hits.</summary>
        public string? FolderKind { get; set; }

        /// <summary>Index snippet (System.Search.AutoSummary), truncated for triage.</summary>
        public string? Snippet { get; set; }

        /// <summary>Item size in bytes.</summary>
        public long? SizeBytes { get; set; }

        /// <summary>Read state.</summary>
        public bool? IsRead { get; set; }

        /// <summary>Whether the item has attachments.</summary>
        public bool? HasAttachments { get; set; }

        /// <summary>True when this hit is an attachment-CONTENT match; read resolves the parent mail.</summary>
        public bool IsAttachmentHit { get; set; }

        /// <summary>Matched attachment file name, for attachment hits.</summary>
        public string? AttachmentFileName { get; set; }

        /// <summary>Conversation id for the thread tool.</summary>
        public string? ConversationId { get; set; }

        /// <summary>
        /// What this hit IS, when it is not ordinary mail - present only then, so it costs
        /// nothing on the usual result and its presence is itself the signal.
        /// <para>
        /// It exists because the three search tiers stopped filtering by item class (gap
        /// B3): bounce reports, read receipts, meeting requests, responses, posts and
        /// sharing invitations now come back beside mail in every tier, and a widened result
        /// set a caller cannot tell apart is worse than the narrow one it replaced.
        /// </para>
        /// <para>
        /// TWO VOCABULARIES, ON PURPOSE, because the two kinds of tier know different
        /// things. A hit from Outlook (<c>source</c> <c>live</c>/<c>exhaustive</c>/<c>com</c>)
        /// carries the real MAPI message class - <c>REPORT.IPM.Note.NDR</c>,
        /// <c>IPM.Schedule.Meeting.Request</c>. A hit from the index carries
        /// <c>kind:&lt;System.Kind&gt;</c> instead (<c>kind:calendar</c>,
        /// <c>kind:unknown</c>), because that tier never opens the item and reporting a bare
        /// class name would claim an authority it does not have. The prefix is what keeps
        /// the two apart at a glance;
        /// <see cref="OutlookAI.Core.Mapi.MailItemAdmission"/> owns both renderings.
        /// </para>
        /// </summary>
        public string? ItemClass { get; set; }
    }

    /// <summary>Freshness gap-sweep diagnostics attached to (non-exhaustive) search results.</summary>
    public sealed class SweepInfo
    {
        /// <summary>
        /// Whether the sweep ran. False has two meanings, told apart by
        /// <see cref="NotNeeded"/>: it could not run (see <see cref="Error"/>), or it had
        /// nothing to do.
        /// </summary>
        public bool Performed { get; set; }

        /// <summary>
        /// True when the sweep did not run because it could not have found anything: the
        /// search's <c>before</c> bound ends at or before <see cref="GapStartUtc"/>, so the
        /// index already covers the whole requested window
        /// (<see cref="FreshMerge.DecideSweepWindow"/>). Null otherwise.
        /// <para>
        /// The distinction is the point: "did not need to run" is a COMPLETE answer, while
        /// "could not run" is a degraded one. Both used to be <c>performed: false</c> with
        /// no error, so a search deliberately bounded to older mail was reported as
        /// <c>degraded</c> and <c>freshness: "index-only"</c> - told it might be missing
        /// recent mail that its own bounds exclude.
        /// </para>
        /// </summary>
        public bool? NotNeeded { get; set; }

        /// <summary>
        /// True when this result was served from the short-lived sweep cache (D34) - no
        /// COM call was made; the swept data is at most <see cref="SweepCache.DefaultTimeToLive"/>
        /// old. Omitted (null) when the sweep ran live.
        /// </summary>
        public bool? Cached { get; set; }

        /// <summary>Age of the cached sweep data in seconds (present only when Cached=true).</summary>
        public double? CacheAgeSeconds { get; set; }

        /// <summary>
        /// Sweep window start (UTC): the WIDEST window this sweep opened.
        /// <para>
        /// One number over what is now a per-store decision. An unscoped sweep opens one
        /// window per store, each from that store's own index frontier, so a store whose
        /// index lags by hours is swept back hours while a current one is swept back
        /// minutes; this reports the earliest of them, i.e. how far back the sweep looked in
        /// the store that needed it most. A store-scoped sweep opens exactly one window and
        /// this is it.
        /// </para>
        /// <para>
        /// The earliest is the honest single number even though it is not the window every
        /// store got, because the claim it supports - "the merged answer covers everything
        /// from here to now" - is true of every store in scope: a store swept from a LATER
        /// start was swept from its own frontier, and its index covers the span in front of
        /// that. Reporting the latest instead would understate coverage that was actually
        /// delivered.
        /// </para>
        /// </summary>
        public DateTime? GapStartUtc { get; set; }

        /// <summary>
        /// True when part of this search's scope had NO index frontier to open a window
        /// from - the index holds no mail whatsoever for it - so the window fell back to a
        /// fixed span and everything older than that span is in neither tier. Null
        /// otherwise; see <see cref="FreshMerge.GapNoIndexFrontier"/>, which this raises,
        /// and <see cref="StoresWithoutIndex"/>, which names the stores where known.
        /// <para>
        /// Set from the frontier probe, so it is a fact about the INDEX and survives a sweep
        /// that could not run or did not need to. It is the input the pure classifier reads;
        /// the code and the advice sentence are its two renderings.
        /// </para>
        /// </summary>
        public bool? IndexFrontierMissing { get; set; }

        /// <summary>
        /// The store(s) in scope the index holds no mail for, when they could be named.
        /// Null when none were found, and also null in the one case where the fact is known
        /// but the names are not: an unscoped search whose PROFILE-wide probe found no mail
        /// anywhere (an unindexed profile - the flag is still set).
        /// <para>
        /// The same fact <c>list_accounts</c> reports as <c>inLocalIndex: false</c>, said
        /// here because this is where it changes an answer: a store the index does not know
        /// is searchable only as far back as the fallback window reaches.
        /// </para>
        /// <para>
        /// CAPPED at <see cref="MailService.UnindexedStoreListCap"/> since 2026-08-18 (Q7b).
        /// It was the one list in this server with no bound, in the payload and in its
        /// advice sentence alike, which on a profile of many unindexed PSTs would have put
        /// the whole store list in both. See <see cref="StoresWithoutIndexTruncated"/> and
        /// <see cref="StoresWithoutIndexTotal"/> - a cap this server does not report is the
        /// defect, not the cap.
        /// </para>
        /// </summary>
        public IReadOnlyList<string>? StoresWithoutIndex { get; set; }

        /// <summary>
        /// True when <see cref="StoresWithoutIndex"/> lists fewer stores than were found;
        /// null when it is complete. The has-more half of the cap, so a short list is never
        /// mistaken for the whole set.
        /// </summary>
        public bool? StoresWithoutIndexTruncated { get; set; }

        /// <summary>
        /// How many stores were actually found without index rows, when that is more than
        /// the list shows. Null otherwise - the list is then its own total.
        /// </summary>
        public int? StoresWithoutIndexTotal { get; set; }

        /// <summary>
        /// What the sweep covered, following the search scope (soak fix 13):
        /// <c>"folder"</c> = the searched folder (plus its subfolders when
        /// include_subfolders is on); <c>"default folders (Inbox, Sent Items, Deleted
        /// Items, Junk Email)"</c> = those folders in the searched store, or in every
        /// store when the search is not store-scoped. Lets an agent see the freshness
        /// coverage of its query.
        /// <para>
        /// A SENTENCE, so branch on <see cref="ScopeShape"/> instead - this one is rendered
        /// from that token (<see cref="MailService.DescribeSweepScope"/>) and the two cannot
        /// describe different breadths.
        /// </para>
        /// </summary>
        public string? Scope { get; set; }

        /// <summary>
        /// The same breadth as <see cref="Scope"/>, as a token software can act on:
        /// <c>default_folders</c>, <c>folder</c> or <c>folder_only</c>
        /// (<see cref="MailService.ClassifySweepScope"/>, gap E2). Null exactly when
        /// <see cref="Scope"/> is - a sweep that was refused or could not run planned no
        /// coverage to describe.
        /// <para>
        /// <c>default_folders</c> is the one that carries a warning: it means Inbox, Sent
        /// Items, Deleted Items and Junk Email of the store(s) in scope and NOT their
        /// subfolders, so mail a server-side rule filed into a subfolder before the indexer
        /// reached it is covered by neither tier. The remedy is a folder scope, which sweeps
        /// that folder's whole subtree.
        /// </para>
        /// <para>
        /// Deliberately NOT a coverage code and never <c>degraded</c>: this shape holds for
        /// nearly every search, and a flag that fires always devalues the ones that fire
        /// rarely (the reasoning <see cref="AttachmentTextCovered"/> is built on).
        /// </para>
        /// </summary>
        public string? ScopeShape { get; set; }

        /// <summary>Folders swept.</summary>
        public int FoldersSwept { get; set; }

        /// <summary>
        /// The swept folders as <c>store/folder path</c>, listed while the set is small
        /// enough to be useful (omitted for a wide all-stores sweep - the count and
        /// <see cref="Scope"/> describe those).
        /// </summary>
        public IReadOnlyList<string>? Folders { get; set; }

        /// <summary>
        /// True when <see cref="Folders"/> was dropped because the sweep covered more
        /// than <see cref="MailService.SweptFolderListCap"/> folders. Without this the
        /// omission is indistinguishable from "no folders to report" - a cap must never
        /// be invisible (section-12 discipline). Null when the list is present.
        /// </summary>
        public bool? FolderListOmitted { get; set; }

        /// <summary>
        /// Folders skipped (unresolvable, unenumerable, or past the folder cap of a scoped
        /// sweep) - each one a hole where fresh mail may be hiding. A default folder the
        /// store does not HAVE is not counted here (see <see cref="FoldersAbsent"/>):
        /// counting it here made every search on a profile with such a store report itself
        /// degraded over a folder that cannot hold anything.
        /// </summary>
        public int FoldersSkipped { get; set; }

        /// <summary>
        /// Default folders the store(s) in scope do not have (a data file with no Junk
        /// Email, say). Present so the coverage arithmetic is checkable against the four
        /// folders <see cref="Scope"/> names - swept plus skipped plus absent - and
        /// deliberately NOT a coverage gap: it raises no gap code, adds no advice and never
        /// degrades the search, because a folder that does not exist cannot be holding
        /// mail. Null when every folder in scope exists, which is the usual case.
        /// <para>
        /// It does one more job, and only <see cref="FreshMerge.DescribeCoverageGaps"/>
        /// does it: when the scope's folders are ALL absent, <see cref="FoldersSwept"/> is
        /// 0 and this counter is the only thing that says the sweep was complete rather
        /// than empty-handed. Without it a store with no arrival-path folders reported
        /// itself degraded on every search that named it.
        /// </para>
        /// </summary>
        public int? FoldersAbsent { get; set; }

        /// <summary>
        /// Folders whose item enumeration FAILED, so they have no freshness coverage at
        /// all. Until soak fix 15 these were counted as successfully swept.
        /// </summary>
        public int FoldersFailed { get; set; }

        /// <summary>
        /// Folders where the per-folder item cap (<c>SweepPerFolderCap</c>) truncated the
        /// window. The sweep reads newest-first, so the OLDEST fresh mail in those folders
        /// was dropped. Null when nothing was truncated.
        /// <para>
        /// That newest-first claim holds for the folders in this list MINUS
        /// <see cref="ItemCappedFoldersUnsorted"/>; the two lists together are always the
        /// whole capped set.
        /// </para>
        /// </summary>
        public IReadOnlyList<string>? ItemCappedFolders { get; set; }

        /// <summary>
        /// The subset of <see cref="ItemCappedFolders"/> whose table Outlook would not sort
        /// by received time, so the cap kept an ARBITRARY slice of the freshness window
        /// rather than its newest end (gap H2). Null when every capped folder sorted.
        /// <para>
        /// It exists because a coverage code alone cannot carry a claim that is sometimes
        /// false. <c>item_cap</c> means "a folder's window was truncated" and its advice
        /// sentence says which part is missing - the oldest - which is what makes narrowing
        /// with <c>after</c> a remedy and what an agent relays to the user. Outlook's
        /// <c>Table.Sort</c> can fail (the property must exist as a column, and late-bound
        /// COM turns E_INVALIDARG into an ArgumentException), and that failure was swallowed
        /// on the grounds that an unsorted sweep still works. It does; the sentence does
        /// not. An arbitrary cut leaves an arbitrary hole, so the caller is not merely told
        /// less than the truth, it is told the wrong thing about which mail to go looking
        /// for.
        /// </para>
        /// <para>
        /// So the fact rides in the payload and both renderings are computed from it: the
        /// folders named here raise <see cref="FreshMerge.GapItemCapUnsorted"/> and its own
        /// sentence, the rest raise <see cref="FreshMerge.GapItemCap"/> and keep theirs, and
        /// neither sentence can name a folder it is not true of
        /// (<see cref="FreshMerge.SortedItemCappedFolders"/> is the one split both read).
        /// </para>
        /// </summary>
        public IReadOnlyList<string>? ItemCappedFoldersUnsorted { get; set; }

        /// <summary>
        /// Folders where the sweep added the received-date COLUMN successfully and
        /// <c>Table.Sort</c> then threw anyway. Null when none did.
        /// <para>
        /// A pure diagnostic: it raises no coverage code, changes no advice and never
        /// degrades an answer. It is here to settle from real sweeps, rather than from a
        /// probe, WHY the sort has been observed not to apply - Microsoft documents that a
        /// sort property may be referenced "by their explicit string names only; cannot
        /// reference properties by their namespaces", and the shipped call passes a
        /// namespace. If that is the cause, this equals the folders swept on every store and
        /// every profile; if a provider is at fault it varies by store. Until 2026-08-19 one
        /// <c>catch</c> covered the column add and the sort together, so
        /// <see cref="ItemCappedFoldersUnsorted"/> could not say which had failed.
        /// </para>
        /// </summary>
        public int? SortRefusedFolders { get; set; }

        /// <summary>
        /// True when the scoped sweep hit <c>MaxScopedSweepFolders</c> and stopped walking
        /// the subtree, so folders past the cut-off were never visited.
        /// </summary>
        public bool? FolderCapReached { get; set; }

        /// <summary>
        /// True when the scoped sweep's subtree walk refused a folder deeper than
        /// <c>OutlookComSession.FolderWalkDepthGuard</c>. Null when the walk stayed
        /// inside the guard (which every real folder tree does).
        /// </summary>
        public bool? DepthLimitReached { get; set; }

        /// <summary>
        /// True when the scoped sweep's subtree walk stopped on
        /// <c>OutlookComSession.ScopedSweepTimeBudgetMs</c>, so the folders it had not
        /// reached yet have no freshness coverage. Null when the walk finished inside
        /// its budget.
        /// </summary>
        public bool? TimeBudgetExceeded { get; set; }

        /// <summary>
        /// True when the WHOLE sweep ran out of <c>MailService.SweepWorkBudgetMs</c> and
        /// stopped at a store or folder boundary, so folders it never reached have no
        /// freshness coverage. Null when it finished inside its budget.
        /// <para>
        /// Separate from <see cref="TimeBudgetExceeded"/> because the remedies point in
        /// different directions: that one says a SUBTREE is wide (scope narrower, or drop
        /// <c>include_subfolders</c>), this one says the PROFILE is big (name a store or a
        /// folder). Before the budget existed, this state was not reported at all - the
        /// outer gateway deadline expired instead, the COM host was killed as unresponsive,
        /// and the folders already swept were discarded.
        /// </para>
        /// </summary>
        public bool? SweepBudgetExpired { get; set; }

        /// <summary>
        /// Every coverage hole this sweep left, as machine-readable codes
        /// (<c>FreshMerge.Gap*</c>). Null when the sweep covered its whole scope, or when
        /// it never ran (that is <c>freshness: "index-only"</c>, not a partial sweep).
        /// <para>
        /// The counters above each state ONE fact about the walk; this states the
        /// CONCLUSION drawn from all of them, which is the thing a caller has to act on.
        /// Until this existed, a sweep that hit a cap, lost a folder or ran out of time
        /// reported the shortfall only as integers here and as prose in <c>advice</c>,
        /// while the two top-level markers still said <c>degraded</c> absent and
        /// <c>freshness: "live"</c> - so an agent reading fields rather than prose, which
        /// is the sensible way to read a payload, was told a partial answer was complete.
        /// Each code has exactly one advice sentence, emitted from this same list
        /// (<see cref="MailService.DescribeSweepCoverage"/>), so the codes and the prose
        /// cannot drift apart.
        /// </para>
        /// </summary>
        public IReadOnlyList<string>? CoverageGaps { get; set; }

        /// <summary>
        /// Table rows the sweep could not turn into items: no usable EntryID column, or
        /// Outlook refused to open the row's item (<see cref="FreshMerge.GapRowsUnreadable"/>,
        /// gap H1). Each one is mail inside the freshness window that the sweep saw and did
        /// not deliver.
        /// <para>
        /// Attributed per store like the folder counters beside it, so a store-scoped search
        /// reads its own losses rather than another account's.
        /// </para>
        /// <para>
        /// Nothing counted these, and the folder counters could not: such a row was skipped
        /// AND did not count toward the per-folder cap, so a folder where every row failed
        /// returned "complete" with zero items and was counted in
        /// <see cref="FoldersSwept"/> as fully covered.
        /// </para>
        /// </summary>
        public int RowsUnreadable { get; set; }

        /// <summary>
        /// Swept items dropped because a property one of the request's own filters needs
        /// could not be read (<see cref="FreshMerge.GapFilterUnreadable"/>, gap I1).
        /// <see cref="FiltersUnevaluated"/> names which filters.
        /// </summary>
        public int ItemsFilterUnreadable { get; set; }

        /// <summary>
        /// The request filters that could not be evaluated on at least one swept item -
        /// <c>"unread_only"</c>, <c>"has_attachments"</c>, <c>"before"</c>, <c>"after"</c> -
        /// in that order, each named once. Null when every filter could be evaluated on
        /// every item, which is the usual case.
        /// <para>
        /// The names are the remedy: they are the request parameters the caller passed, so
        /// re-running without the one named returns the dropped items.
        /// </para>
        /// </summary>
        public IReadOnlyList<string>? FiltersUnevaluated { get; set; }

        /// <summary>Items in the window before term filtering.</summary>
        public int ItemsSeen { get; set; }

        /// <summary>
        /// Swept items in scope whose BODY was cut before it crossed the COM-host pipe
        /// (<c>OutlookComSession.SweepBodyCharsCap</c> per item,
        /// <c>OutlookComSession.SweepBodyBytesBudget</c> across the whole sweep). Null when
        /// nothing was cut, which is every ordinary search.
        /// <para>
        /// A FACT, not a hole, and it raises no code on its own: these bodies are matched
        /// against and never shown, so an item that was cut and matched anyway lost nothing.
        /// <see cref="ItemsBodyCappedUnmatched"/> is the subset that could have cost a
        /// result, and that is what <c>FreshMerge.GapBodyCap</c> reads. Reported beside it so
        /// the ratio is legible: 1 of 1 and 1 of 400 mean very different things about how
        /// close this profile runs to the bound.
        /// </para>
        /// <para>
        /// Counted from the per-item <c>ComMailBrief.BodyTruncated</c> flags in the same loop
        /// that applies the store scope, NOT from the sweep's whole-sweep total, so a cached
        /// all-stores sweep serving a store-scoped search reports this store's cuts and not
        /// another account's.
        /// </para>
        /// </summary>
        public int? ItemsBodyCapped { get; set; }

        /// <summary>
        /// The subset of <see cref="ItemsBodyCapped"/> that then failed to match the query
        /// terms - the only items where cutting the body can have cost a hit. Raises
        /// <c>FreshMerge.GapBodyCap</c>. Null when there are none.
        /// <para>
        /// The two facts it is built from ARE separable and are separated: the cut is measured
        /// in the COM layer, per item, and the match is decided here, for the same item. What
        /// no measurement can settle is whether the term really sat past the cut, since that
        /// needs the part of the body the bound refused to carry - so the field counts
        /// candidates and the advice says "may be missing".
        /// </para>
        /// <para>
        /// Structurally zero for a subject-only search (the body is never consulted) and for
        /// a term-less one (everything matches), which is why the code cannot cry wolf on the
        /// searches where a cut body is harmless.
        /// </para>
        /// </summary>
        public int? ItemsBodyCappedUnmatched { get; set; }

        /// <summary>
        /// True when the whole-sweep body budget ran out, so items swept after that point
        /// carried little or none of their body; null when only the per-item ceiling cut, or
        /// when nothing was cut at all.
        /// <para>
        /// It changes the remedy, which is the only reason it is in the payload: a per-item
        /// cut points at ONE enormous mail, which <c>read</c> can page in full, while an
        /// exhausted budget points at the sweep's own breadth and is answered with a narrower
        /// store, folder or window.
        /// </para>
        /// <para>
        /// Carried only when this scope actually suffered a cut. The budget belongs to the
        /// FRAME, which spans every store the sweep visited, so reporting it in a store-scoped
        /// answer that lost nothing would import another account's condition - the
        /// cross-store leak the per-store counters exist to prevent.
        /// </para>
        /// </summary>
        public bool? BodyBudgetExhausted { get; set; }

        /// <summary>
        /// <c>false</c> when this query's terms could have matched inside an attachment and
        /// the sweep cannot look there (gap B2); null when the question does not arise - a
        /// subject-only search, or one with no terms at all.
        /// <para>
        /// Never true, and that is deliberate rather than an oversight. The value is a
        /// statement about the TIER, not about this particular sweep: the sweep reads
        /// <c>MailItem.Subject</c> and <c>MailItem.Body</c> through COM and never opens an
        /// attachment, so there is no state of the world in which it covers attachment text.
        /// A field that is either <c>false</c> or absent says exactly that, and an agent
        /// branching on <c>=== false</c> gets the right answer without having to know it.
        /// </para>
        /// <para>
        /// WHY IT IS HERE. The attachment-ONLY search is refused outright and reported
        /// (<c>sweep.error</c>, <c>freshness: "index-only"</c>, <c>degraded</c>). The
        /// DEFAULT search said nothing: the index tier matches
        /// <c>System.Search.Contents</c>, which is body PLUS attachment content, the sweep
        /// re-matches subject and body only, and the merged answer then reports
        /// <c>freshness: "live"</c> - so a term inside an attachment of mail that arrived
        /// after the index frontier is invisible under the one word an agent reads as
        /// "nothing is missing".
        /// </para>
        /// <para>
        /// It raises no coverage code and does not degrade the search, and the reason is
        /// arithmetic rather than taste: every gap code makes a search <c>partial</c>
        /// (<c>FreshMerge.ClassifyFreshness</c> derives the verdict from the code list), and
        /// this condition holds for nearly every search anyone runs. A flag that fires
        /// always tells a caller nothing and devalues the flag that fires rarely. The
        /// advice sentence is narrower still and fires only where the gap could actually be
        /// hiding something - see <c>MailService.DescribeAttachmentTextGap</c>.
        /// </para>
        /// </summary>
        public bool? AttachmentTextCovered { get; set; }

        /// <summary>Swept items dropped as already present in the index results.</summary>
        public int Duplicates { get; set; }

        /// <summary>Sweep wall-clock cost.</summary>
        public long ElapsedMs { get; set; }

        /// <summary>
        /// Stores this sweep covered under a <c>Com.StoreNaming</c> label because Outlook
        /// would not report their display name (gap G2). Null when every store named itself.
        /// <para>
        /// A REPORT rather than a coverage gap, and the distinction is the point. Such a
        /// store used to be abandoned at the failed name read - four folders added to
        /// <see cref="FoldersSkipped"/>, no per-store bucket, its fresh mail simply absent -
        /// and it is swept now, so there is no hole to raise. What is left is a NAMING hole
        /// with a consequence the caller has to know: mail from this store is in the answer,
        /// and the store cannot be used as a <c>store</c> scope for a follow-up.
        /// </para>
        /// </summary>
        public int? StoresUnnamed { get; set; }

        /// <summary>Content-free error when the sweep could not run.</summary>
        public string? Error { get; set; }
    }

    /// <summary>Exhaustive-scan diagnostics attached to exhaustive:true results.</summary>
    public sealed class ExhaustiveInfo
    {
        /// <summary>Term engine used: "ci_phrasematch" (index-backed DASL), "like" (substring scan), or "ci_phrasematch+like".</summary>
        public string Engine { get; set; } = string.Empty;

        /// <summary>Store.IsInstantSearchEnabled as reported by Outlook (the ci_* gate).</summary>
        public bool InstantSearchEnabled { get; set; }

        /// <summary>Mail folders the scan filtered.</summary>
        public int FoldersScanned { get; set; }

        /// <summary>Folders where the filter failed under both engines.</summary>
        public int FoldersSkipped { get; set; }

        /// <summary>True when the result cap stopped the scan (results may be incomplete).</summary>
        public bool Truncated { get; set; }

        /// <summary>True when the time budget stopped the scan (results may be incomplete).</summary>
        public bool TimedOut { get; set; }

        /// <summary>
        /// True when the scan refused one or more subtrees for sitting deeper than
        /// <c>OutlookComSession.FolderWalkDepthGuard</c>, so those folders were never opened
        /// (<see cref="FreshMerge.ScanGapDepthLimit"/>, gap F4).
        /// <para>
        /// The walk had NO depth bound at all until 2026-08-18, which made a cyclic or
        /// pathological folder graph an uncatchable <c>StackOverflowException</c> that ended
        /// the COM host - reported to the caller as Outlook having disappeared. The bound is
        /// the one the sweep walk and the folder-listing walk already use; this flag is what
        /// keeps hitting it from being the silent truncation those two were fixed for
        /// (gaps G3/G4).
        /// </para>
        /// </summary>
        public bool DepthLimitReached { get; set; }

        /// <summary>
        /// Rows the scan examined and did not admit, for ANY reason - the same counter, with
        /// the same meaning, that the index tier keeps as
        /// <c>IndexSearch.IndexSearchResult.RowsDropped</c>. One shape across tiers rather
        /// than a second vocabulary for the third one.
        /// <para>
        /// It is a DIAGNOSTIC, not a coverage hole, and raises nothing on its own. Most of
        /// it is the scan's deliberate item-class filter: only <c>IPM.Note</c> mail is
        /// admitted, so meeting requests and responses, NDRs and read receipts, posts and
        /// sharing invitations are counted here and dropped. Subtract
        /// <see cref="RowsUnreadable"/> to get exactly that number - which is the measurement
        /// the tier-asymmetry question (gap B3) needs and nothing has ever reported.
        /// </para>
        /// </summary>
        public int RowsDropped { get; set; }

        /// <summary>
        /// The subset of <see cref="RowsDropped"/> that was a FAILURE rather than a filter:
        /// a row with no usable EntryID, one Outlook would not open, one whose item class
        /// could not be read (<see cref="FreshMerge.ScanGapRowsUnreadable"/>, gap F5). Any
        /// of them may have been a match.
        /// </summary>
        public int RowsUnreadable { get; set; }

        /// <summary>
        /// Scanned items dropped because a property one of the request's own filters needs
        /// could not be read (<see cref="FreshMerge.ScanGapFilterUnreadable"/>) - the sweep's
        /// gap I1 in this tier. <see cref="FiltersUnevaluated"/> names which filters.
        /// </summary>
        public int ItemsFilterUnreadable { get; set; }

        /// <summary>
        /// The request filters that could not be evaluated on at least one scanned item.
        /// Only <c>unread_only</c> and <c>has_attachments</c> are reachable here - this
        /// mode's date bounds are applied by the DASL filter, not read back off the item.
        /// Null when every filter could be evaluated on every item.
        /// </summary>
        public IReadOnlyList<string>? FiltersUnevaluated { get; set; }

        /// <summary>
        /// The request filters this mode could only apply AFTER its own result cap -
        /// <c>from</c>, <c>unread_only</c>, <c>has_attachments</c>, in request order
        /// (<see cref="FreshMerge.PostCapFilters"/>, gap F3). Null when the request passed
        /// none.
        /// <para>
        /// The scan's <c>maxItems</c> is <c>top</c> and it counts items the DASL filter
        /// matched. These three are read off the returned snapshots, so they thin the list
        /// the cap already closed - which is how an exhaustive search returns 2 rows with
        /// <see cref="Truncated"/> set while thousands of items match. Present whether or
        /// not the cap fired, because it says how the answer was BUILT; it is the pairing
        /// with <see cref="Truncated"/> that raises
        /// <see cref="FreshMerge.ScanGapPostCapFilter"/>.
        /// </para>
        /// </summary>
        public IReadOnlyList<string>? PostCapFilters { get; set; }

        /// <summary>
        /// Scanned items one of <see cref="PostCapFilters"/> evaluated and EXCLUDED - the
        /// caller asked for them to go, so this is not a coverage hole and raises nothing on
        /// its own (gap F3).
        /// <para>
        /// It is the size of the thinning, and it is what turns
        /// <see cref="FreshMerge.ScanGapPostCapFilter"/> from an abstract warning into a
        /// measurement: a scan that stopped at <c>top</c> = 25 and returned 2 rows spent 23
        /// of its 25 on items the filter then dropped, so the cap was reached on candidates
        /// rather than on results. Distinct from <see cref="ItemsFilterUnreadable"/>, which
        /// counts items those filters could not be evaluated on at all.
        /// </para>
        /// </summary>
        public int ItemsFilteredOut { get; set; }

        /// <summary>
        /// Every coverage hole this scan left, as machine-readable codes
        /// (<c>FreshMerge.ScanGap*</c>), on the same contract
        /// <see cref="SweepInfo.CoverageGaps"/> carries for the sweep and
        /// <c>ThreadLiveInfo</c> for the conversation walk. Null when the scan covered its
        /// whole scope.
        /// <para>
        /// The counters above each state ONE fact; this states the conclusion drawn from all
        /// of them, and <see cref="SearchOutcome.Freshness"/> is recomputed from it
        /// (<see cref="FreshMerge.ClassifyExhaustiveFreshness"/>) so a code can never ship
        /// beside <c>freshness: "live"</c>. Each code has exactly one advice sentence,
        /// emitted from this same list by <c>MailService.DescribeExhaustiveCoverage</c>.
        /// </para>
        /// </summary>
        public IReadOnlyList<string>? CoverageGaps { get; set; }

        /// <summary>Scan wall-clock cost.</summary>
        public long ElapsedMs { get; set; }

        /// <summary>
        /// Which bound ENDED the walk: <c>complete</c>, <c>time_budget</c> or
        /// <c>result_cap</c> (<c>ComScanStopReasons</c>). RECORDED by the walk at the moment
        /// it stopped, never derived here from <see cref="Truncated"/> and
        /// <see cref="TimedOut"/>.
        /// <para>
        /// Those two are independent booleans and both can be true, while their remedies
        /// point in opposite directions - measured: on Exchange the token exists because of
        /// the time budget (a 108 144-item folder at roughly 12 items/s), on a local PST
        /// because of the result cap (roughly 1 200 items/s, so the clock never fires). A
        /// budget stop means "keep resuming, there is no cheaper route"; a cap stop means
        /// "keep resuming, or narrow with folder/after, and narrowing is cheaper". An agent
        /// that cannot tell them apart gives the wrong advice about half the time.
        /// </para>
        /// <para>
        /// <see cref="DepthLimitReached"/> is deliberately NOT a value here: the depth guard
        /// bounds one subtree and every sibling branch is still walked, so it never ends a
        /// walk however true it is.
        /// </para>
        /// </summary>
        public string? StopReason { get; set; }

        /// <summary>
        /// The handle that continues this scan, or null when there is nothing left to do
        /// (F2). Pass it back as <c>resume_token</c> with every other argument unchanged.
        /// <para>
        /// Absent EXACTLY when the walk covered its scope, so a paging caller terminates on
        /// this field being missing rather than on a count - a count cannot distinguish "the
        /// last page happened to be short" from "there is no more".
        /// </para>
        /// <para>
        /// It can also be absent on a scan that DID stop, when no resumable position could be
        /// formed; the advice says so in that case rather than leaving the caller to infer
        /// completeness from a missing field.
        /// </para>
        /// </summary>
        public string? NextToken { get; set; }

        /// <summary>
        /// Hits returned across every page of this token chain, this page included.
        /// <para>
        /// <c>top</c> counts PER PAGE, which is how every other paging surface here behaves,
        /// and the accumulation across pages is exactly the context cost the <c>top</c> cap of
        /// 100 exists to bound. This number is what makes that cost visible so an agent can
        /// stop deliberately instead of discovering it afterwards.
        /// </para>
        /// </summary>
        public int? ItemsReturnedTotal { get; set; }

        /// <summary>Where the next page carries on. Present exactly when the walk stopped early.</summary>
        public ScanPositionInfo? Position { get; set; }

        /// <summary>
        /// Folders that appeared in scope since the token was issued (added, moved, or
        /// renamed into an earlier position) plus folders the chain had finished that are no
        /// longer there. The appeared ones are SCANNED, never skipped; the departed ones were
        /// already covered. Zero on a scan that was not resumed.
        /// </summary>
        public int? TreeChangedFolders { get; set; }

        /// <summary>True when the folder the previous page stopped inside was gone on resume.</summary>
        public bool? CursorFolderMissing { get; set; }

        /// <summary>True when a resumed folder had to be re-read from its beginning.</summary>
        public bool? ResumedUnsorted { get; set; }

        /// <summary>True when a resumed folder's recorded row position no longer identified the same row.</summary>
        public bool? ResumePositionLost { get; set; }

        /// <summary>True when the per-folder duplicate-suppression set filled, so duplicates may now appear.</summary>
        public bool? DedupCapacityReached { get; set; }

        /// <summary>True when this page continues an earlier scan rather than starting one.</summary>
        public bool? Resumed { get; set; }
    }

    /// <summary>
    /// How a folder-scoped search was actually resolved (present only when the request
    /// carried a folder). Exists so a caller can SEE the answer's real breadth instead of
    /// assuming it: a delegate mailbox can only be covered by folder NAME, and a request
    /// that could not be narrowed is widened, never silently trimmed.
    /// </summary>
    public sealed class SearchScopeInfo
    {
        /// <summary>The folder path as requested.</summary>
        public string? Folder { get; set; }

        /// <summary>Whether subfolders were requested.</summary>
        public bool IncludeSubfolders { get; set; }

        /// <summary>
        /// The resolution shape: <c>folder</c> (recursive folder scope),
        /// <c>folder_only</c> (folder without subfolders), <c>delegate_folders</c> (flat
        /// delegate namespace, matched by folder name) or <c>delegate_store_widened</c>.
        /// </summary>
        public string Shape { get; set; } = string.Empty;

        /// <summary>True when the answer covers MORE than the requested folder subtree.</summary>
        public bool? Widened { get; set; }

        /// <summary>How many flat folder names the delegate query matched (delegate scopes only).</summary>
        public int? FolderNamesMatched { get; set; }

        /// <summary>
        /// True when the COM folder walk that produced those names was itself cut short by
        /// the walk cap or the depth guard, so the name set is SHORT and the delegate scope
        /// under-returns (gap G4). Null when the walk covered the mailbox's tree.
        /// <para>
        /// It matters more here than anywhere else the same walk is used, which is why it is
        /// reported rather than left to the folder listing. The delegate index namespace is
        /// FLAT: a delegate folder scope cannot be a subtree predicate, it can only be an OR
        /// of folder NAMES read out of the COM tree. A missing name is therefore not a
        /// missing row in a listing, it is a folder whose mail no tier looks in - and
        /// <see cref="FolderNamesMatched"/>, being a count, reads exactly the same whether
        /// the walk saw the whole mailbox or stopped halfway.
        /// </para>
        /// </summary>
        public bool? FolderNamesTruncated { get; set; }
    }

    /// <summary>
    /// What the INDEX TIER did on this search: present on every non-exhaustive search,
    /// absent on an exhaustive one, which bypasses the index by design.
    /// <para>
    /// It exists because these three numbers were computed on every search and reached
    /// nobody. <c>IndexSearchResult.RowsDropped</c> was the count of what the tier refused,
    /// and gap B3's whole complaint was that it never surfaced;
    /// <see cref="CandidatesExhausted"/> was the one way the post-filter could hide matches
    /// and said so only in an advice sentence (gap G6), which an agent reading fields rather
    /// than prose - the sensible way to read a payload - never saw.
    /// </para>
    /// <para>
    /// Adding it was previously declined on the grounds that the <c>search</c> tool
    /// description could not afford to document it. Measurement settled that: the client cap
    /// is per STRING, the description sits at 1791 of 2048 units, and a payload block needs
    /// no description text at all - the block is self-describing and the README carries the
    /// reference.
    /// </para>
    /// </summary>
    public sealed class IndexTierInfo
    {
        /// <summary>
        /// True when the index tier did not run AT ALL because the requested <c>store</c>
        /// exists in the Outlook profile but the local index has no scope that addresses it
        /// (a PST, an archive-only data file, a fresh install, indexing off, excluded or
        /// still building). Null on every other search.
        /// <para>
        /// It exists because <see cref="RowsScanned"/> 0 cannot say this. "The statement ran
        /// and matched nothing" and "no statement ran, and none could have" lead to different
        /// next moves: the first means there is no such mail, the second means this answer
        /// rests on the freshness sweep alone and <c>exhaustive:true</c> is the way to read
        /// the store in full. <c>sweep.storesWithoutIndex</c> names the store, and this says
        /// the tier was skipped rather than merely empty.
        /// </para>
        /// <para>
        /// The alternative to skipping was to run the query with NO scope, which is what
        /// <c>thread</c> does with an unresolvable store. Rejected here: <c>search</c>'s
        /// <c>store</c> is a filter on the result set, not a lookup hint, so widening it
        /// would return another account's mail under a scope the caller chose.
        /// </para>
        /// </summary>
        public bool? StoreNotIndexed { get; set; }

        /// <summary>
        /// True when the index tier RAN, returned nothing, and its own probes then found no
        /// indexed row at all under the requested <c>folder</c> bound while the store around
        /// it does hold rows (gap G5). Null on every other search, including one where the
        /// probes could not run.
        /// <para>
        /// It is the folder-level sibling of <see cref="StoreNotIndexed"/> and answers the
        /// same question one scope down: can the index address what the caller named. The
        /// two are not interchangeable - <c>storeNotIndexed</c> means no statement ran, this
        /// means the statement ran and its folder predicate matched nothing - and either way
        /// the consequence is the one worth branching on: for that scope this answer rests
        /// on the freshness sweep alone, so nothing older than the sweep window is covered
        /// and <c>exhaustive: true</c> is the way to read the folder in full.
        /// </para>
        /// <para>
        /// It exists because the fact was PROSE-only and, worse, conditional on the merged
        /// answer being completely empty - so one item the freshness sweep returned removed
        /// the only trace of it. A swept item proves the folder resolves in COM, which is a
        /// different tier and a different question.
        /// </para>
        /// <para>
        /// It does not set <c>degraded</c>: a folder created minutes ago is genuinely absent
        /// from the index and is not a defect, and the probes cannot tell that apart from a
        /// path that does not resolve. Reporting the fact is the honest half; asserting a
        /// hole would be a false alarm on every new folder.
        /// </para>
        /// </summary>
        public bool? FolderNotIndexed { get; set; }

        /// <summary>Rows the SQL statement returned, before admission (the denominator of <see cref="RowsDropped"/>).</summary>
        public int RowsScanned { get; set; }

        /// <summary>
        /// Rows the index tier examined and did not admit - deliberately the same name and
        /// meaning as <see cref="ExhaustiveInfo.RowsDropped"/>, so the two tiers report one
        /// counter shape rather than two vocabularies.
        /// <para>
        /// A DIAGNOSTIC, not a coverage hole: it raises nothing and never degrades a search.
        /// What lands here is rows outside the mapi namespace (only reachable when the
        /// statement has no SCOPE) and rows of the wrong shape for what was asked - an
        /// attachment row under <c>include_attachment_hits: false</c>, a message row under
        /// <c>attachment_hits_only</c>. Since gap B3 no message row is dropped for its item
        /// class, in this tier or any other.
        /// </para>
        /// </summary>
        public int RowsDropped { get; set; }

        /// <summary>
        /// True when this tier could not establish that the list holds every match: the
        /// over-fetched candidate list ran out before enough rows were admitted, or the
        /// follow-up query that recovers rows the result ordering may have hidden failed
        /// (<c>IndexSearch.IndexOrderGuard</c>). Those are the two ways this tier CAN hide
        /// matches, so unlike <see cref="RowsDropped"/> it is worth acting on. Null otherwise
        /// (gap G6: it used to be advice-only).
        /// </summary>
        public bool? CandidatesExhausted { get; set; }
    }

    /// <summary>Index staleness snapshot attached to search results.</summary>
    public sealed class StalenessInfo
    {
        /// <summary>
        /// Newest indexed DateReceived (UTC) in the STORE this search was scoped to, or
        /// across the whole profile when it named no store
        /// (<see cref="MailService.StalenessScopeFor"/>).
        /// <para>
        /// It said "across the searched scope" while the probe ran unscoped for every
        /// search: measured 2026-08-18, five store-scoped searches reported one profile-wide
        /// frontier while the per-store probes spanned 45.4 hours, which pinned a quiet
        /// store's sweep window to a busy store's clock.
        /// </para>
        /// <para>
        /// IT IS NO LONGER THE SWEEP'S WINDOW BASE ON AN UNSCOPED SEARCH, and that is the
        /// point rather than a discrepancy. A store-scoped search still has exactly one
        /// frontier and this is it. An unscoped one opens a window per store, each from that
        /// store's own frontier, because this figure is a MAXIMUM across stores and a maximum
        /// cannot bound anyone else's lag - <see cref="SweepInfo.GapStartUtc"/> is what the
        /// sweep actually looked back to. This stays the profile-wide value because that is
        /// what it has always meant, and because narrowing it to the worst store would make
        /// search and outlook_health report different numbers for the same profile.
        /// </para>
        /// <para>
        /// One exception, stated rather than papered over: an <c>exhaustive</c> search
        /// reports the PROFILE-wide value. That path resolves no index scope by design - it
        /// answers from COM alone - and this block is context there, not the basis of the
        /// answer.
        /// </para>
        /// </summary>
        public DateTime? NewestIndexedUtc { get; set; }

        /// <summary>
        /// The OLDEST index frontier among the stores this search covered - "how far behind
        /// is the worst store in scope", which is the question
        /// <see cref="NewestIndexedUtc"/> structurally cannot answer.
        /// <para>
        /// The two exist side by side rather than one replacing the other, and that is the
        /// decision (Q7a, 2026-08-18). <see cref="NewestIndexedUtc"/> is a MAXIMUM and stays
        /// one, because narrowing it to the worst store would make <c>search</c> and
        /// <c>outlook_health</c> report different numbers for the same profile. A maximum
        /// cannot bound anyone else's lag, though, so an unscoped search reporting it alone
        /// says nothing about the account that is actually behind - which is exactly the
        /// store whose recent mail is at risk.
        /// </para>
        /// <para>
        /// Store-scoped search: the same value as <see cref="NewestIndexedUtc"/>, because
        /// one store is in scope and its frontier is both the newest and the oldest.
        /// Unscoped search: the earliest of the per-store frontiers the sweep planner
        /// measured, i.e. the figure the freshness advice already quotes. Absent when no
        /// per-store frontier was measured at all - an exhaustive search (no index scope by
        /// design) or an unscoped search whose store catalog could not be read - and absence
        /// means "not measured", never "no lag". A store the index holds nothing for has no
        /// frontier to be oldest and is named in <see cref="SweepInfo.StoresWithoutIndex"/>
        /// instead.
        /// </para>
        /// </summary>
        public DateTime? OldestStoreFrontierUtc { get; set; }

        /// <summary>Age of the newest indexed mail in minutes.</summary>
        public double? AgeMinutes { get; set; }

        /// <summary>
        /// Whether OUTLOOK.EXE is running (the index only advances while it runs).
        /// Snapshot taken AFTER the freshness sweep/scan, so an Outlook the sweep just
        /// autostarted reports true (D34 self-consistency fix).
        /// </summary>
        public bool OutlookRunning { get; set; }
    }

    /// <summary>Search outcome (search tool payload).</summary>
    public sealed class SearchOutcome
    {
        /// <summary>
        /// True when this result is NOT fully fresh, so mail that arrived since the last
        /// index update may be missing. Two ways to earn it: the live Outlook check could
        /// not run at all (<see cref="Freshness"/> <c>"index-only"</c>), or it ran and
        /// covered only part of what it was asked to (<c>"partial"</c> - see
        /// <see cref="SweepInfo.CoverageGaps"/> for which holes).
        /// <para>
        /// Deliberately a blunt top-level boolean as well as prose in <c>advice</c>. The
        /// same fact used to live only in an advice sentence and in sweep.performed, which
        /// is easy to skim past; a result that LOOKS complete but is not is the one failure
        /// mode here that can mislead a reader rather than merely inconvenience them.
        /// </para>
        /// <para>
        /// The partial case was added later and WIDENED this flag rather than adding a
        /// second one, because the sentence above was always its documented meaning - a
        /// sweep that hit a cap or lost a folder is not fully fresh either. The widening
        /// only ever turns false into true, so a caller that already keys on
        /// <c>degraded == true</c> keeps working and simply stops being lied to. What such
        /// a caller must NOT do is infer the reason from <c>freshness == "index-only"</c>:
        /// that value still means, exactly as before, that the sweep never ran.
        /// </para>
        /// <para>
        /// On an EXHAUSTIVE search there is no index to lag behind, and the flag keeps the
        /// meaning that survives the change of tier: this answer covers less than it was
        /// asked to. It is set when the scan timed out, skipped folders it could not filter,
        /// or stopped at the result cap partway through the folder tree - facts
        /// <c>exhaustive.*</c> already carried while these two flags stayed absent, which
        /// left the mode chosen FOR completeness the only one that could not say it had
        /// fallen short.
        /// </para>
        /// </summary>
        public bool? Degraded { get; set; }

        /// <summary>
        /// <c>"live"</c> when the freshness sweep ran and covered its whole scope,
        /// <c>"partial"</c> when it ran but left coverage holes (see
        /// <see cref="SweepInfo.CoverageGaps"/>), <c>"index-only"</c> when it could not run.
        /// <para>
        /// An EXHAUSTIVE search carries the same three-value contract read over its own
        /// single tier (<see cref="FreshMerge.ClassifyExhaustiveFreshness"/>): the scan IS
        /// the live check, so it is <c>"live"</c> when it covered its scope and
        /// <c>"partial"</c> when its own counters say it did not - timed out, skipped
        /// folders, or stopped at the result cap mid-walk. <c>"index-only"</c> cannot occur
        /// there; it names the state this mode exists to avoid.
        /// </para>
        /// </summary>
        public string? Freshness { get; set; }

        /// <summary>Merged hits, newest first.</summary>
        public IReadOnlyList<HitSummary> Hits { get; set; } = Array.Empty<HitSummary>();

        /// <summary>
        /// True when the 'top' cap cut the result list - more matches EXIST (section 12
        /// has-more discipline: raise top or narrow the query). Determined by
        /// over-fetching one row past the cap, so true is definite, not a guess.
        /// </summary>
        public bool Truncated { get; set; }

        /// <summary>Index query wall-clock cost (0 for exhaustive searches - the index is bypassed).</summary>
        public long IndexElapsedMs { get; set; }

        /// <summary>Freshness-sweep diagnostics (absent on exhaustive searches).</summary>
        public SweepInfo? Sweep { get; set; }

        /// <summary>Index-tier diagnostics (absent on exhaustive searches, which bypass the index).</summary>
        public IndexTierInfo? Index { get; set; }

        /// <summary>Exhaustive-scan diagnostics (exhaustive searches only).</summary>
        public ExhaustiveInfo? Exhaustive { get; set; }

        /// <summary>How the folder scope resolved (folder-scoped searches only).</summary>
        public SearchScopeInfo? Scope { get; set; }

        /// <summary>Staleness self-report (R7/D19). Best-effort on exhaustive searches.</summary>
        public StalenessInfo Staleness { get; set; } = new StalenessInfo();

        /// <summary>Agent-facing advice when results may be incomplete.</summary>
        public IReadOnlyList<string>? Advice { get; set; }
    }

    /// <summary>Recipient view for read results.</summary>
    public sealed class RecipientView
    {
        /// <summary>"to", "cc" or "bcc".</summary>
        public string Kind { get; set; } = "to";

        /// <summary>Display name.</summary>
        public string? Name { get; set; }

        /// <summary>SMTP address when resolvable.</summary>
        public string? Address { get; set; }
    }

    /// <summary>Attachment view for read results.</summary>
    public sealed class AttachmentView
    {
        /// <summary>1-based index for save_attachment.</summary>
        public int Index { get; set; }

        /// <summary>File name.</summary>
        public string? FileName { get; set; }

        /// <summary>Size in bytes.</summary>
        public long? SizeBytes { get; set; }
    }

    /// <summary>Read outcome (read tool payload).</summary>
    public sealed class ReadOutcome
    {
        /// <summary>The hit id this read resolved (echoed back when one was used).</summary>
        public string? Id { get; set; }

        /// <summary>REAL Outlook EntryID (usable directly in future read/save_attachment calls).</summary>
        public string EntryId { get; set; } = string.Empty;

        /// <summary>Store display name.</summary>
        public string? Store { get; set; }

        /// <summary>Folder path as Outlook reports it.</summary>
        public string? Folder { get; set; }

        /// <summary>Subject.</summary>
        public string? Subject { get; set; }

        /// <summary>Sender display name.</summary>
        public string? FromName { get; set; }

        /// <summary>Sender SMTP address.</summary>
        public string? FromAddress { get; set; }

        /// <summary>Received timestamp, UTC.</summary>
        public DateTime? ReceivedUtc { get; set; }

        /// <summary>Sent timestamp, UTC.</summary>
        public DateTime? SentUtc { get; set; }

        /// <summary>To/Cc/Bcc recipients (capped - check RecipientsTruncated).</summary>
        public IReadOnlyList<RecipientView> Recipients { get; set; } = Array.Empty<RecipientView>();

        /// <summary>True when Recipients was capped at the payload limit (present only then).</summary>
        public bool? RecipientsTruncated { get; set; }

        /// <summary>Real recipient count before capping (present only when capped).</summary>
        public int? RecipientsTotal { get; set; }

        /// <summary>Plain-text body window [bodyOffset, bodyOffset + max_body_chars) of the full body.</summary>
        public string Body { get; set; } = string.Empty;

        /// <summary>
        /// Effective start of the returned body window within the full body (omitted
        /// when 0, i.e. the body starts at its beginning). Continue reading with
        /// body_offset = bodyOffset + body.length while bodyTruncated.
        /// </summary>
        public int? BodyOffset { get; set; }

        /// <summary>Full body length in characters (all windows).</summary>
        public long BodyTotalChars { get; set; }

        /// <summary>True when more body exists BEYOND the returned window.</summary>
        public bool BodyTruncated { get; set; }

        /// <summary>"text" (Outlook's own plain-text rendering), "html-converted", or "none".</summary>
        public string BodyOrigin { get; set; } = "text";

        /// <summary>
        /// Stored HTML body (Outlook's own HTMLBody), only when include_html=true - the
        /// only way to see formatting, signature placement and quote boundaries, which the
        /// plain-text body collapses. Windowed from its start with the same max_body_chars
        /// budget as the text body (batch B, B2).
        /// </summary>
        public string? BodyHtml { get; set; }

        /// <summary>Full HTML body length in characters (only when include_html=true).</summary>
        public long? BodyHtmlTotalChars { get; set; }

        /// <summary>True when BodyHtml was cut at the character budget (only when include_html=true).</summary>
        public bool? BodyHtmlTruncated { get; set; }

        /// <summary>Total item size in bytes.</summary>
        public long? SizeBytes { get; set; }

        /// <summary>Read state.</summary>
        public bool? IsRead { get; set; }

        /// <summary>Conversation id for the thread tool.</summary>
        public string? ConversationId { get; set; }

        /// <summary>Internet Message-ID (durable across moves - use for dedupe, not EntryID).</summary>
        public string? InternetMessageId { get; set; }

        /// <summary>Transport headers (only when include_headers=true; may be truncated).</summary>
        public string? Headers { get; set; }

        /// <summary>True when Headers was cut at the cap.</summary>
        public bool? HeadersTruncated { get; set; }

        /// <summary>
        /// Attachments (save via save_attachment with the same id + index). Capped -
        /// check AttachmentsTruncated; indexes beyond the cap remain saveable, they are
        /// just not listed.
        /// </summary>
        public IReadOnlyList<AttachmentView> Attachments { get; set; } = Array.Empty<AttachmentView>();

        /// <summary>True when Attachments was capped at the payload limit (present only then).</summary>
        public bool? AttachmentsTruncated { get; set; }

        /// <summary>Real attachment count before capping (present only when capped).</summary>
        public int? AttachmentsTotal { get; set; }

        /// <summary>How the hit was located ("cached", "urlSegments", "itemPathDisplay", "directEntryId").</summary>
        public string? LocatedVia { get; set; }

        /// <summary>Locate cost for this call (0 when served from cache).</summary>
        public long? LocateMs { get; set; }
    }

    /// <summary>save_attachment outcome.</summary>
    public sealed class SaveAttachmentOutcome
    {
        /// <summary>The hit id used.</summary>
        public string? Id { get; set; }

        /// <summary>Parent item's real EntryID.</summary>
        public string EntryId { get; set; } = string.Empty;

        /// <summary>1-based attachment index saved.</summary>
        public int AttachmentIndex { get; set; }

        /// <summary>Saved file name (sanitized, uniquified).</summary>
        public string FileName { get; set; } = string.Empty;

        /// <summary>Absolute path of the saved file - read it from here.</summary>
        public string SavedPath { get; set; } = string.Empty;

        /// <summary>Saved file size in bytes.</summary>
        public long SizeBytes { get; set; }
    }

    /// <summary>
    /// What the LIVE tier of one <c>thread</c> lookup did - the conversation walk, which is
    /// to <c>thread</c> what the freshness sweep is to <c>search</c>.
    /// <para>
    /// The walk asks Outlook for the conversation itself rather than sweeping folders, and
    /// that is why <c>thread</c> can afford to be live on every call: the scope IS one
    /// conversation, so there is no window to open, no folder set to choose and no per-folder
    /// cap to hit. It needs an anchor item though - COM cannot look up a conversation by id
    /// string - so it is the one live tier here that can be unavailable for a reason the
    /// caller can fix, by passing <c>id</c>.
    /// </para>
    /// </summary>
    public sealed class ThreadLiveInfo
    {
        /// <summary>Whether the live conversation walk ran. False means <see cref="Error"/> says why.</summary>
        public bool Performed { get; set; }

        /// <summary>
        /// Content-free reason the walk did not run: <c>NoAnchorItem</c> (only
        /// <c>conversation_id</c> was passed, and COM needs a concrete item to walk from),
        /// <c>AnchorNotLocatable</c> (the anchor could not be opened in Outlook), or a COM
        /// failure token. Null when it ran.
        /// </summary>
        public string? Error { get; set; }

        /// <summary>Store display name the walk covered (Outlook walks the anchor's store only).</summary>
        public string? AnchorStore { get; set; }

        /// <summary>Conversation members the walk returned.</summary>
        public int MembersWalked { get; set; }

        /// <summary>
        /// True when the walk stopped at the requested member cap, so it did not see the
        /// whole conversation (<see cref="FreshMerge.ThreadGapMemberCap"/>). Determined by
        /// over-fetching one member past <c>top</c>, so true is definite.
        /// </summary>
        public bool MemberCapReached { get; set; }

        /// <summary>Members the walk contributed that the index did not already hold.</summary>
        public int MembersAdded { get; set; }

        /// <summary>Walk wall-clock cost, including locating the anchor.</summary>
        public long ElapsedMs { get; set; }

        /// <summary>
        /// Stores the profile HAS, the local index holds no mail for, and this walk did not
        /// cover - so a member of this conversation sitting in one of them is in NEITHER
        /// tier (<see cref="FreshMerge.ThreadGapUnindexedStore"/>, gap C4's silent half).
        /// Null when there are none.
        /// <para>
        /// The same fact <c>search</c> reports as <c>sweep.storesWithoutIndex</c>, said here
        /// because it changes the answer here too: Outlook walks a conversation inside ONE
        /// store, and every other store is index-only for this lookup.
        /// </para>
        /// <para>
        /// CAPPED at <see cref="MailService.UnindexedStoreListCap"/> - the same cap, from the
        /// same constant, the sweep's list carries. See
        /// <see cref="StoresWithoutIndexTruncated"/> and <see cref="StoresWithoutIndexTotal"/>.
        /// </para>
        /// </summary>
        public IReadOnlyList<string>? StoresWithoutIndex { get; set; }

        /// <summary>
        /// True when <see cref="StoresWithoutIndex"/> lists fewer stores than were found;
        /// null when it is complete.
        /// </summary>
        public bool? StoresWithoutIndexTruncated { get; set; }

        /// <summary>
        /// How many such stores were found, when that is more than the list shows. Null
        /// otherwise - the list is then its own total.
        /// </summary>
        public int? StoresWithoutIndexTotal { get; set; }

        /// <summary>
        /// Coverage holes of this lookup, in the same vocabulary as
        /// <see cref="SweepInfo.CoverageGaps"/> (<see cref="FreshMerge.DescribeThreadCoverageGaps"/>).
        /// Null when the answer is whole.
        /// </summary>
        public IReadOnlyList<string>? CoverageGaps { get; set; }
    }

    /// <summary>thread outcome.</summary>
    public sealed class ThreadOutcome
    {
        /// <summary>Conversation id the thread was resolved for.</summary>
        public string? ConversationId { get; set; }

        /// <summary>
        /// True when this conversation is NOT fully fresh, so replies newer than the index
        /// may be missing - the same flag, with the same meaning and the same obligation to
        /// relay it, that <see cref="SearchOutcome.Degraded"/> carries.
        /// <para>
        /// <c>thread</c> had no freshness fields at all, which made it the one tool that
        /// could not express a partial answer while being the one whose description promises
        /// "the FULL conversation". With a single index row for the conversation the COM walk
        /// never ran, so a reply that arrived after the index frontier was simply absent and
        /// nothing in the payload said so.
        /// </para>
        /// </summary>
        public bool? Degraded { get; set; }

        /// <summary>
        /// <c>"live"</c> when the conversation walk ran and covered the conversation,
        /// <c>"partial"</c> when it ran but left a hole (see
        /// <see cref="ThreadLiveInfo.CoverageGaps"/>), <c>"index-only"</c> when it could not
        /// run at all. Same three values, same meanings, as on a search.
        /// </summary>
        public string? Freshness { get; set; }

        /// <summary>
        /// Which tier the member list RESTS on: "index" (the ConversationID query returned
        /// rows) or "com" (it did not, so the Outlook Conversation walk is the whole answer).
        /// <para>
        /// It is no longer a statement that only one tier ran - since C1 both do, and their
        /// members are merged. Read <see cref="Live"/> for what the live tier did and each
        /// hit's own <c>source</c> for where that member came from.
        /// </para>
        /// </summary>
        public string Source { get; set; } = "index";

        /// <summary>Thread members, oldest first - index rows and live walk members, deduped.</summary>
        public IReadOnlyList<HitSummary> Hits { get; set; } = Array.Empty<HitSummary>();

        /// <summary>
        /// True when the 'top' cap cut the member list - the conversation HAS more
        /// members (over-fetch-by-one, same contract as search.truncated).
        /// </summary>
        public bool Truncated { get; set; }

        /// <summary>
        /// True when the requested <c>store</c> did not resolve to an index scope and the
        /// lookup ran over the WHOLE profile instead (gap C3). Over-returning is the safe
        /// direction, so the widening stays - but silently widening a scope the caller chose
        /// is not, so it is reported. Null when the scope was honoured or none was asked for.
        /// </summary>
        public bool? ScopeWidened { get; set; }

        /// <summary>The live conversation walk's own report. Always present.</summary>
        public ThreadLiveInfo? Live { get; set; }

        /// <summary>Index staleness snapshot, same block and same meaning as on a search.</summary>
        public StalenessInfo? Staleness { get; set; }

        /// <summary>Agent-facing advice when the conversation may be incomplete.</summary>
        public IReadOnlyList<string>? Advice { get; set; }

        /// <summary>Wall-clock cost of the thread lookup.</summary>
        public long ElapsedMs { get; set; }
    }

    /// <summary>
    /// Per-store index-freshness row of the outlook_health report.
    /// <para>
    /// The rows are the UNION of what the index knows and what Outlook reports, not the
    /// index's list alone. They were the index's alone, so a store Outlook has mounted and
    /// the index has never seen simply did not appear - and that is the exact condition
    /// that makes searches of it fall back to a fixed window, i.e. the one thing this block
    /// most needs to say. Two lists that disagree, with nothing naming the disagreement,
    /// was the defect.
    /// </para>
    /// </summary>
    public sealed class StoreStaleness
    {
        /// <summary>Store display name.</summary>
        public string Store { get; set; } = string.Empty;

        /// <summary>
        /// Newest indexed DateReceived under that store's scope (UTC). Absent when the
        /// index holds no mail for the store, which <see cref="InLocalIndex"/> tells apart
        /// from a probe that could not be run.
        /// </summary>
        public DateTime? NewestIndexedUtc { get; set; }

        /// <summary>
        /// Whether the local index holds anything for this store. False means searches of
        /// it are served by the freshness sweep alone, over a fixed fallback window: mail
        /// older than that is not findable through this server at all until the store is
        /// indexed (or an <c>exhaustive:true</c> search names it). Null when the store was
        /// not probed - never guessed.
        /// </summary>
        public bool? InLocalIndex { get; set; }
    }

    /// <summary>Account view for list_accounts.</summary>
    public sealed class AccountView
    {
        /// <summary>Account SMTP address.</summary>
        public string? SmtpAddress { get; set; }

        /// <summary>Account display name.</summary>
        public string? DisplayName { get; set; }

        /// <summary>Store new mail lands in.</summary>
        public string? DeliveryStore { get; set; }
    }

    /// <summary>Store view for list_accounts (D22/D25 searchability flags).</summary>
    public sealed class StoreView
    {
        /// <summary>Store display name (use as the search tool's store argument).</summary>
        public string DisplayName { get; set; } = string.Empty;

        /// <summary>
        /// True when <see cref="DisplayName"/> is a <c>Com.StoreNaming</c> LABEL because
        /// Outlook would not report this store's real name (gap G2) - so it is the one entry
        /// in this list that CANNOT be used as the <c>store</c> argument, since a store scope
        /// is matched against the display name that could not be read. Null otherwise.
        /// <para>
        /// Such a store was absent from this list entirely until 2026-08-18, which made
        /// <c>list_accounts</c> - the tool whose whole job is to say what stores exist -
        /// quietly incomplete.
        /// </para>
        /// </summary>
        public bool? NameUnreadable { get; set; }

        /// <summary>True for delegate/shared mailbox caches (distinct from the 3 accounts).</summary>
        public bool IsDelegate { get; set; }

        /// <summary>Cached Exchange Mode state; false = server-only.</summary>
        public bool? IsCachedExchange { get; set; }

        /// <summary>Raw OlExchangeStoreType (0 primary, 1 additional/delegate, 2 public folders, 3 not Exchange).</summary>
        public int? ExchangeStoreType { get; set; }

        /// <summary>True for server-only stores (e.g. Online Archives) - invisible to local search (D22/D25).</summary>
        public bool OnlineOnly { get; set; }

        /// <summary>False when the local index cannot see this store; search cannot cover it.</summary>
        public bool LocallySearchable { get; set; }

        /// <summary>Whether any indexed item was found for this store (null = not probed).</summary>
        public bool? InLocalIndex { get; set; }
    }

    /// <summary>list_accounts outcome.</summary>
    public sealed class AccountsOutcome
    {
        /// <summary>Profile mail accounts.</summary>
        public IReadOnlyList<AccountView> Accounts { get; set; } = Array.Empty<AccountView>();

        /// <summary>All stores (accounts + delegates + anything else mounted).</summary>
        public IReadOnlyList<StoreView> Stores { get; set; } = Array.Empty<StoreView>();
    }

    /// <summary>One installed signature for list_signatures.</summary>
    public sealed class SignatureView
    {
        /// <summary>Signature name - the draft tools' 'signature' argument.</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>Short plain-text excerpt (first lines) - use it to detect each signature's language/purpose.</summary>
        public string? Excerpt { get; set; }
    }

    /// <summary>Registry-determined default-signature assignment of one account (list_signatures).</summary>
    public sealed class SignatureAccountView
    {
        /// <summary>Account SMTP address.</summary>
        public string Account { get; set; } = string.Empty;

        /// <summary>Default signature for new messages (absent = unknown, never guessed).</summary>
        public string? NewMessage { get; set; }

        /// <summary>Default signature for replies/forwards (absent = unknown).</summary>
        public string? ReplyForward { get; set; }
    }

    /// <summary>list_signatures outcome.</summary>
    public sealed class SignaturesOutcome
    {
        /// <summary>Installed signatures (name + excerpt), name-sorted.</summary>
        public IReadOnlyList<SignatureView> Signatures { get; set; } = Array.Empty<SignatureView>();

        /// <summary>Per-account default assignments as far as the registry records them (absent when unreadable).</summary>
        public IReadOnlyList<SignatureAccountView>? Accounts { get; set; }

        /// <summary>Explains unknown defaults when assignments are missing (degrade-gracefully contract).</summary>
        public string? Note { get; set; }
    }

    /// <summary>manage_signature outcome (soak fix D38).</summary>
    public sealed class ManageSignatureOutcome
    {
        /// <summary>Executed action: create | update | delete.</summary>
        public string Action { get; set; } = string.Empty;

        /// <summary>Signature name operated on.</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>Absolute paths of the rendition files written (create/update: .htm + .txt + .rtf).</summary>
        public IReadOnlyList<string>? FilesWritten { get; set; }

        /// <summary>Absolute paths removed by delete (files and the _files resource directory).</summary>
        public IReadOnlyList<string>? FilesDeleted { get; set; }

        /// <summary>Backup directory holding the previous file set (always present for update/delete).</summary>
        public string? BackupPath { get; set; }

        /// <summary>Account whose default assignment was recorded (set_default_for).</summary>
        public string? DefaultSetForAccount { get; set; }

        /// <summary>Scope recorded for that account: new | reply | both.</summary>
        public string? DefaultSetScope { get; set; }

        /// <summary>Accounts whose dangling default assignments were cleared by a delete.</summary>
        public IReadOnlyList<string>? DefaultsClearedForAccounts { get; set; }

        /// <summary>Operational guidance (e.g. Outlook restart pickup of default changes).</summary>
        public string? Advice { get; set; }
    }

    /// <summary>Folder view for list_folders.</summary>
    public sealed class FolderView
    {
        /// <summary>Store-relative path ('/'-separated) - use as the search tool's folder argument.</summary>
        public string Path { get; set; } = string.Empty;

        /// <summary>Total items.</summary>
        public long? Items { get; set; }

        /// <summary>Unread items.</summary>
        public long? Unread { get; set; }
    }

    /// <summary>Folders of one store for list_folders.</summary>
    public sealed class StoreFoldersView
    {
        /// <summary>Store display name.</summary>
        public string Store { get; set; } = string.Empty;

        /// <summary>
        /// True when <see cref="Store"/> is a LABEL rather than a name, because Outlook
        /// would not report this store's display name (gap G2, <c>Com.StoreNaming</c>).
        /// <para>
        /// The store used to be dropped from this listing entirely, so its whole folder tree
        /// was missing with nothing saying so. It is listed now, and this flag is the half
        /// that keeps the fix honest: the label CANNOT be passed back as the <c>store</c>
        /// argument of <c>search</c> or <c>list_folders</c>, because a store scope is
        /// resolved by comparing against the very display name that could not be read.
        /// </para>
        /// </summary>
        public bool? NameUnreadable { get; set; }

        /// <summary>This page's folders of the store (full tree, stable traversal order).</summary>
        public IReadOnlyList<FolderView> Folders { get; set; } = Array.Empty<FolderView>();
    }

    /// <summary>list_folders outcome (full tree, offset-paged in stable traversal order).</summary>
    public sealed class FoldersOutcome
    {
        /// <summary>Folder trees per store (this page).</summary>
        public IReadOnlyList<StoreFoldersView> Stores { get; set; } = Array.Empty<StoreFoldersView>();

        /// <summary>
        /// Total folders in the full traversal (all pages) - a LOWER BOUND when
        /// <see cref="WalkCapReached"/> or <see cref="DepthLimitReached"/> is set, since the
        /// walk that produced it stopped before the tree did.
        /// </summary>
        public int FolderTotal { get; set; }

        /// <summary>Echo of a non-zero requested offset (omitted for the first page).</summary>
        public int? Offset { get; set; }

        /// <summary>
        /// True when this answer is not the whole tree - either more folders exist beyond
        /// this page (continue with <see cref="NextOffset"/>) or the walk itself was cut
        /// short (<see cref="WalkCapReached"/> / <see cref="DepthLimitReached"/>, gap G3).
        /// <para>
        /// It used to mean only the first of those, and it was computed against the list the
        /// walk had ALREADY truncated - so the one case it could not see was the one that
        /// mattered: a tree cut off at the walk cap paged out as <c>truncated: false</c>,
        /// i.e. as a complete answer. The two causes are told apart by the flags beside it,
        /// and only the pageable one carries a <see cref="NextOffset"/>: paging cannot get
        /// past a walk cap, because the next call re-walks and stops at the same place.
        /// </para>
        /// </summary>
        public bool Truncated { get; set; }

        /// <summary>The offset that continues the listing (present only when more PAGES exist).</summary>
        public int? NextOffset { get; set; }

        /// <summary>
        /// True when the folder walk stopped at <see cref="MailService.FolderWalkAbsoluteCap"/>,
        /// so folders - and, once the cap falls mid-profile, whole stores - were never
        /// visited. Null when the walk covered the tree. Not pageable: narrow with
        /// <c>store</c> instead.
        /// </summary>
        public bool? WalkCapReached { get; set; }

        /// <summary>
        /// True when the walk refused a folder deeper than
        /// <c>OutlookComSession.FolderWalkDepthGuard</c> (64), so that subtree is missing.
        /// Null when the walk stayed inside the guard, which every real tree does - the
        /// guard is what stops a cyclic tree taking the process down.
        /// </summary>
        public bool? DepthLimitReached { get; set; }

        /// <summary>
        /// Stores listed under a <c>Com.StoreNaming</c> label because Outlook would not
        /// report their display name (gap G2). Null when every store named itself.
        /// </summary>
        public int? StoresUnnamed { get; set; }

        /// <summary>
        /// Stores a store-SCOPED listing had to leave out because they would not report a
        /// name, so they could be neither matched against the requested store nor ruled out
        /// of it. Null on an unscoped listing, where nothing has to be decided and such a
        /// store is listed under its label instead.
        /// </summary>
        public int? StoresUnnamedExcluded { get; set; }

        /// <summary>
        /// What a caller has to know about this listing that the counters alone do not say -
        /// the same role <c>advice</c> plays on a search. Null when the listing is complete
        /// and every store named itself, which is the usual case.
        /// </summary>
        public IReadOnlyList<string>? Advice { get; set; }
    }

    /// <summary>open_in_outlook outcome (v3.MD L3).</summary>
    public sealed class OpenInOutlookOutcome
    {
        /// <summary>The hit id used (when one was).</summary>
        public string? Id { get; set; }

        /// <summary>REAL EntryID of the item now shown in an Outlook Inspector window.</summary>
        public string EntryId { get; set; } = string.Empty;

        /// <summary>Subject of the displayed item.</summary>
        public string? Subject { get; set; }

        /// <summary>Always true on success - the item is on screen.</summary>
        public bool Displayed { get; set; }
    }

    /// <summary>goto_folder outcome (v3.MD L3).</summary>
    public sealed class GotoFolderOutcome
    {
        /// <summary>Store display name navigated to.</summary>
        public string Store { get; set; } = string.Empty;

        /// <summary>Store-relative folder path requested (null = the store's Inbox/root).</summary>
        public string? Folder { get; set; }

        /// <summary>ActiveExplorer().CurrentFolder.FolderPath after navigation (\\Store\Folder\...).</summary>
        public string? ExplorerFolderPath { get; set; }

        /// <summary>Explorer window caption after navigation.</summary>
        public string? ExplorerCaption { get; set; }

        /// <summary>Always true on success - the folder is on screen.</summary>
        public bool Displayed { get; set; }
    }

    /// <summary>Draft-tool outcome (v3.MD L4/D4): the draft is SAVED, never sent.</summary>
    public sealed class DraftOutcome
    {
        /// <summary>"new", "reply", "replyall" or "forward".</summary>
        public string Kind { get; set; } = "new";

        /// <summary>The hit id the source mail was referenced by (derived drafts, when one was used).</summary>
        public string? Id { get; set; }

        /// <summary>EntryID of the source mail (derived drafts).</summary>
        public string? SourceEntryId { get; set; }

        /// <summary>REAL EntryID of the saved draft (usable with read/open_in_outlook).</summary>
        public string EntryId { get; set; } = string.Empty;

        /// <summary>Store the draft was saved in.</summary>
        public string? Store { get; set; }

        /// <summary>Drafts folder name (localized).</summary>
        public string? Folder { get; set; }

        /// <summary>SmtpAddress the draft will send as (SendUsingAccount).</summary>
        public string? Account { get; set; }

        /// <summary>True when SendUsingAccount was pinned from a matched Account object.</summary>
        public bool AccountResolved { get; set; }

        /// <summary>Draft subject (RE:/FW: prefixes included for derived drafts).</summary>
        public string? Subject { get; set; }

        /// <summary>True when Outlook injected the account's DEFAULT signature into the body.</summary>
        public bool SignatureInjected { get; set; }

        /// <summary>The signature name requested via the 'signature' parameter (absent when the account default was used).</summary>
        public string? Signature { get; set; }

        /// <summary>Whether the requested signature override was applied (absent when none was requested; false = the default-signature draft stands).</summary>
        public bool? SignatureApplied { get; set; }

        /// <summary>Content-free reason when a requested override failed (the draft is still valid, with the default signature).</summary>
        public string? SignatureError { get; set; }

        /// <summary>True when the draft was opened in an Outlook window for the user (D4 default).</summary>
        public bool Displayed { get; set; }

        /// <summary>Conversation id (derived drafts thread with their source).</summary>
        public string? ConversationId { get; set; }

        /// <summary>Recipients currently on the draft (capped - check RecipientsTruncated).</summary>
        public IReadOnlyList<RecipientView>? Recipients { get; set; }

        /// <summary>True when Recipients was capped at the payload limit (present only then).</summary>
        public bool? RecipientsTruncated { get; set; }

        /// <summary>Real recipient count before capping (present only when capped).</summary>
        public int? RecipientsTotal { get; set; }

        /// <summary>
        /// Addresses Outlook could NOT resolve against the address book (present only
        /// when there are any). They stay on the draft - the user can fix them in the
        /// compose window - but they are never dropped silently (batch A, A2).
        /// <para>
        /// CAPPED at <see cref="MailService.UnresolvedRecipientsCap"/>. Check
        /// <see cref="UnresolvedRecipientsTruncated"/> before reporting this list as the
        /// complete set of failures.
        /// </para>
        /// </summary>
        public IReadOnlyList<string>? UnresolvedRecipients { get; set; }

        /// <summary>
        /// True when <see cref="UnresolvedRecipients"/> lists fewer addresses than failed to
        /// resolve; null when it is complete. The has-more half of the cap, and the half that
        /// was missing: the list was cut at
        /// <see cref="MailService.UnresolvedRecipientsCap"/> with nothing saying so, on the
        /// one surface where a short list is acted on rather than merely read - the remedy
        /// for an unresolved address is to ask the user about it, so an address past the cap
        /// was never mentioned to anybody.
        /// </summary>
        public bool? UnresolvedRecipientsTruncated { get; set; }

        /// <summary>
        /// How many addresses failed to resolve in total, when that is more than the list
        /// shows. Null otherwise - the list is then its own total, exactly as
        /// <see cref="RecipientsTotal"/> works for the resolved list.
        /// </summary>
        public int? UnresolvedRecipientsTotal { get; set; }

        /// <summary>
        /// What to do about a CUT unresolved list, present only with
        /// <see cref="UnresolvedRecipientsTruncated"/>
        /// (<see cref="MailService.DescribeUnresolvedRecipientCap"/>). Rendered from the same
        /// two values as the fields above, so the sentence and the fields cannot disagree.
        /// </summary>
        public string? UnresolvedRecipientsAdvice { get; set; }

        /// <summary>
        /// Only present when a derived draft's subject was overridden: whether the
        /// original conversation topic could be carried over so the draft still groups
        /// with the source thread (batch A, A3).
        /// </summary>
        public bool? ConversationTopicPreserved { get; set; }

        /// <summary>Draft importance when it is not the default ("low" or "high").</summary>
        public string? Importance { get; set; }

        /// <summary>Present (true) only when a read receipt was requested.</summary>
        public bool? ReadReceiptRequested { get; set; }

        /// <summary>Which body form was used: "text" (from body) or "html" (from body_html).</summary>
        public string? BodyFormat { get; set; }

        /// <summary>
        /// How the body reached the item: "wordEditor" (composed inside the held Inspector
        /// like Outlook's own compose window - the normal path) or "html" (the fallback
        /// wholesale HTMLBody assignment). On "html" the rendering is less faithful, so
        /// check the draft with read include_html=true.
        /// </summary>
        public string? BodyPlacement { get; set; }

        /// <summary>
        /// D49: present (true) only when Outlook was window-less and the compose surface had
        /// to be promoted invisibly to obtain the Word editor. Purely informational - the
        /// draft is composed to the full standard either way; it exists so headless
        /// composition is observable instead of assumed.
        /// </summary>
        public bool? ComposeSurfacePromoted { get; set; }

        /// <summary>
        /// D49: present ONLY when the composition fell back to the HTMLBody splice, naming
        /// the content-free reason the Word surface was unavailable. Never gated on a
        /// signature override having been requested: a fallback draft is a lesser draft
        /// whatever the caller asked for, and D48 shipped precisely because that was silent.
        /// </summary>
        public string? ComposeSurfaceError { get; set; }

        /// <summary>D49: what the caller can do about a degraded composition. Present with <see cref="ComposeSurfaceError"/>.</summary>
        public string? ComposeSurfaceAdvice { get; set; }

        /// <summary>
        /// Only present when body_html was supplied AND normalization changed something:
        /// exactly what was dropped, unwrapped, escaped or repaired, so the agent can see
        /// how its markup was altered instead of guessing (batch B, B1).
        /// </summary>
        public IReadOnlyList<string>? HtmlAdjustments { get; set; }

        /// <summary>
        /// Files actually on the SAVED draft (name + size read back from Outlook, not
        /// echoed from the request - D46/C3). Absent when the draft carries none.
        /// </summary>
        public IReadOnlyList<AttachmentView>? Attachments { get; set; }

        /// <summary>Total size of the attachment set in bytes (present only with attachments).</summary>
        public long? AttachmentsTotalBytes { get; set; }

        /// <summary>
        /// How many files the call asked to attach - present only when attachments were
        /// requested, so a mismatch with the Attachments list is visible rather than silent.
        /// </summary>
        public int? AttachmentsRequested { get; set; }
    }

    /// <summary>
    /// update_draft outcome (v3.MD D46/C1). Everything is read back from the SAVED draft;
    /// <see cref="Changed"/> names exactly which parts this call rewrote, so an agent can
    /// see that an omitted field really was left alone.
    /// </summary>
    public sealed class UpdateDraftOutcome
    {
        /// <summary>Always "updated" on success (a refusal comes back as an error object).</summary>
        public string Status { get; set; } = "updated";

        /// <summary>
        /// Present and true only when this call FINISHED an earlier update whose outcome was
        /// unknown, instead of performing a fresh revision. Absent on an ordinary update, so
        /// an agent never has to read it to know nothing unusual happened.
        /// </summary>
        public bool? Resumed { get; set; }

        /// <summary>What <see cref="Resumed"/> means for the fields beside it; present only with it.</summary>
        public string? ResumedAdvice { get; set; }

        /// <summary>The hit id the draft was referenced by, when one was used.</summary>
        public string? Id { get; set; }

        /// <summary>EntryID of the updated draft.</summary>
        public string EntryId { get; set; } = string.Empty;

        /// <summary>Store holding the draft.</summary>
        public string? Store { get; set; }

        /// <summary>Drafts folder name (localized).</summary>
        public string? Folder { get; set; }

        /// <summary>SmtpAddress the draft will send as.</summary>
        public string? Account { get; set; }

        /// <summary>Subject after the update.</summary>
        public string? Subject { get; set; }

        /// <summary>Which parts this call actually changed (body, subject, to, cc, bcc, ...).</summary>
        public IReadOnlyList<string>? Changed { get; set; }

        /// <summary>True when the revised draft was (re)opened in an Outlook window.</summary>
        public bool Displayed { get; set; }

        /// <summary>Conversation id after the update (unchanged by a subject rewrite - A3).</summary>
        public string? ConversationId { get; set; }

        /// <summary>Recipients on the draft AFTER the update (capped).</summary>
        public IReadOnlyList<RecipientView>? Recipients { get; set; }

        /// <summary>True when Recipients was capped at the payload limit.</summary>
        public bool? RecipientsTruncated { get; set; }

        /// <summary>Real recipient count before capping (present only when capped).</summary>
        public int? RecipientsTotal { get; set; }

        /// <summary>
        /// Addresses Outlook could not resolve; they stay on the draft. Capped at
        /// <see cref="MailService.UnresolvedRecipientsCap"/> - check
        /// <see cref="UnresolvedRecipientsTruncated"/> before reporting it as complete.
        /// </summary>
        public IReadOnlyList<string>? UnresolvedRecipients { get; set; }

        /// <summary>True when the unresolved list was cut by the cap; null when it is whole.</summary>
        public bool? UnresolvedRecipientsTruncated { get; set; }

        /// <summary>How many addresses failed to resolve, when that is more than the list shows.</summary>
        public int? UnresolvedRecipientsTotal { get; set; }

        /// <summary>
        /// What to do about a cut unresolved list, present only with
        /// <see cref="UnresolvedRecipientsTruncated"/>. Same sentence, same source values, as
        /// the draft creators emit - one cap, one wording, whichever tool hit it.
        /// </summary>
        public string? UnresolvedRecipientsAdvice { get; set; }

        /// <summary>Only when the subject was replaced: whether threading survived it (A3).</summary>
        public bool? ConversationTopicPreserved { get; set; }

        /// <summary>Importance when it is not the default.</summary>
        public string? Importance { get; set; }

        /// <summary>Present (true) only when a read receipt is requested.</summary>
        public bool? ReadReceiptRequested { get; set; }

        /// <summary>Signature name requested for this update, when one was.</summary>
        public string? Signature { get; set; }

        /// <summary>Whether the requested signature swap was applied.</summary>
        public bool? SignatureApplied { get; set; }

        /// <summary>"text" or "html" - present only when a body was supplied.</summary>
        public string? BodyFormat { get; set; }

        /// <summary>"wordEditor" - present only when a body was supplied.</summary>
        public string? BodyPlacement { get; set; }

        /// <summary>What the HTML normalizer changed, when body_html was supplied.</summary>
        public IReadOnlyList<string>? HtmlAdjustments { get; set; }

        /// <summary>Files on the draft after the update (read back from the saved item).</summary>
        public IReadOnlyList<AttachmentView>? Attachments { get; set; }

        /// <summary>Total size of the attachment set in bytes.</summary>
        public long? AttachmentsTotalBytes { get; set; }

        /// <summary>File names this call added.</summary>
        public IReadOnlyList<string>? AttachmentsAdded { get; set; }

        /// <summary>File names this call removed.</summary>
        public IReadOnlyList<string>? AttachmentsRemoved { get; set; }

        /// <summary>
        /// Requested removals that matched nothing on the draft - reported rather than
        /// silently ignored, so a misspelled file name is visible.
        /// </summary>
        public IReadOnlyList<string>? AttachmentsNotFound { get; set; }

        /// <summary>
        /// Files Outlook refused at attach time despite passing the pre-COM checks. Never
        /// silent: the rest of the update stands, and these names say exactly what did not
        /// make it onto the draft.
        /// </summary>
        public IReadOnlyList<string>? AttachmentsFailed { get; set; }

        /// <summary>
        /// Present (non-zero) only when the revision LOST inline images the draft already
        /// carried (D47) - a draft whose signature image was still a <c>file:///</c> link
        /// instead of an embedded resource, which Word cannot re-serialize. Paired with
        /// <see cref="InlineImagesAdvice"/>, which names the fix.
        /// </summary>
        public int? InlineImagesDropped { get; set; }

        /// <summary>What to do about <see cref="InlineImagesDropped"/>; null when none were lost.</summary>
        public string? InlineImagesAdvice { get; set; }
    }

    /// <summary>
    /// discard_draft outcome (v3.MD D46/C2, S1 v3). The delete is SOFT: the draft is in
    /// Deleted Items, and <see cref="NewEntryId"/> + <see cref="FromFolder"/> make the
    /// operation reversible exactly like a move (D39).
    /// </summary>
    public sealed class DiscardDraftOutcome
    {
        /// <summary>Always "discarded" on success.</summary>
        public string Status { get; set; } = "discarded";

        /// <summary>True - the draft was soft-deleted.</summary>
        public bool Discarded { get; set; }

        /// <summary>The hit id the draft was referenced by, when one was used.</summary>
        public string? Id { get; set; }

        /// <summary>EntryID the draft had in Drafts (now stale).</summary>
        public string EntryId { get; set; } = string.Empty;

        /// <summary>EntryID in Deleted Items when it could be re-located (EntryIDs change on any move).</summary>
        public string? NewEntryId { get; set; }

        /// <summary>Store the draft lived in.</summary>
        public string? Store { get; set; }

        /// <summary>Folder it was discarded from - the undo address.</summary>
        public string? FromFolder { get; set; }

        /// <summary>Deleted Items folder name it went to.</summary>
        public string? ToFolder { get; set; }

        /// <summary>Subject of the discarded draft.</summary>
        public string? Subject { get; set; }

        /// <summary>How to undo it.</summary>
        public string? Advice { get; set; }
    }

    /// <summary>
    /// send-tool outcome (v3.MD L5/D4 - high-friction two-step flow). Status
    /// "confirmation_required": NOTHING was sent; a one-time confirm token was issued.
    /// Status "sent": the mail went out with the verified identity reported here.
    /// </summary>
    public sealed class SendOutcome
    {
        /// <summary>"confirmation_required" (step 1 - nothing sent) or "sent" (step 2 - transport accepted).</summary>
        public string Status { get; set; } = "confirmation_required";

        /// <summary>True only when the mail was actually handed to the transport.</summary>
        public bool Sent { get; set; }

        /// <summary>Strong policy warning (step 1): confirm with the user before using the token.</summary>
        public string? Warning { get; set; }

        /// <summary>One-time token for the confirming send call (step 1 only).</summary>
        public string? ConfirmToken { get; set; }

        /// <summary>Seconds until the token expires (step 1 only).</summary>
        public double? TokenExpiresInSeconds { get; set; }

        /// <summary>The hit id the draft was referenced by (when one was used).</summary>
        public string? Id { get; set; }

        /// <summary>Draft EntryID this flow operated on (invalid after a successful send - sent items get a new EntryID).</summary>
        public string EntryId { get; set; } = string.Empty;

        /// <summary>Store the draft lives/lived in.</summary>
        public string? Store { get; set; }

        /// <summary>Folder the draft was in when the token was issued (step 1 only).</summary>
        public string? Folder { get; set; }

        /// <summary>SmtpAddress of the sending identity - always the account owning the draft's store.</summary>
        public string? Account { get; set; }

        /// <summary>True when the SendUsingAccount getter readback matched right before Send() (step 2).</summary>
        public bool? AccountVerified { get; set; }

        /// <summary>SentOnBehalfOfName applied to the outgoing mail, when requested.</summary>
        public string? SentOnBehalfOf { get; set; }

        /// <summary>Draft subject (so the model can restate to the user WHAT would be / was sent).</summary>
        public string? Subject { get; set; }

        /// <summary>Recipients the mail would go / went to (confirm these with the user in step 1; capped - check RecipientsTruncated).</summary>
        public IReadOnlyList<RecipientView>? Recipients { get; set; }

        /// <summary>True when Recipients was capped at the payload limit (present only then; the mail still goes to ALL recipients).</summary>
        public bool? RecipientsTruncated { get; set; }

        /// <summary>Real recipient count before capping (present only when capped).</summary>
        public int? RecipientsTotal { get; set; }
    }

    /// <summary>Outlook block of the health report (Phase 7).</summary>
    public sealed class OutlookHealthView
    {
        /// <summary>Whether OUTLOOK.EXE is running for this user.</summary>
        public bool Running { get; set; }

        /// <summary>
        /// True when the running Outlook is headless (no window, tray icon only - the
        /// D17 autostart state; launch Outlook normally to promote it to a windowed
        /// session). False when a window exists; null when Outlook is not running.
        /// </summary>
        public bool? Headless { get; set; }

        /// <summary>Installed classic-Outlook build (OUTLOOK.EXE file version; null when not found).</summary>
        public string? Version { get; set; }

        /// <summary>
        /// Office MAJOR whose registry hive this server reads: "16.0" (Outlook 2016 through
        /// Microsoft 365), "15.0" (2013) or "17.0" (the next one). A different fact from
        /// <see cref="Version"/>, which is the installed build.
        /// <para>
        /// NULL when none of the supported majors is registered - and that null is the whole
        /// point of the field: registry-backed answers (accounts, signature defaults, the
        /// Outlook search settings) then read a fallback hive and can come back empty on a
        /// perfectly healthy Outlook. problems says so in words when this is absent.
        /// </para>
        /// </summary>
        public string? OfficeVersion { get; set; }

        /// <summary>True while the add-in installer holds the OutlookAISetup mutex (D17: COM tools retry later).</summary>
        public bool InstallerMutexHeld { get; set; }

        /// <summary>True when this server holds a COM session that Outlook ANSWERED just now (probed liveness, SF-1).</summary>
        public bool ComConnected { get; set; }

        /// <summary>
        /// Whether Outlook is servicing its message queue, according to Windows itself -
        /// not to us. False means COM calls into it would not return.
        /// <para>
        /// This is judged with a Win32 check that costs microseconds and cannot block, so
        /// it is trustworthy even when everything COM-shaped is stuck. Null when Outlook is
        /// not running.
        /// </para>
        /// </summary>
        public bool? Responding { get; set; }

        /// <summary>"not running", "starting", "responsive" or "not responding".</summary>
        public string? State { get; set; }

        /// <summary>Store count reachable over COM (null when Outlook is not running - health never starts it).</summary>
        public int? StoresReachable { get; set; }

        /// <summary>Reachable store display names.</summary>
        public IReadOnlyList<string>? Stores { get; set; }

        /// <summary>
        /// How Outlook is being reached and whether that path is healthy. With the COM
        /// host this reports the child's state, its PID, and how many times it has been
        /// restarted - a climbing restart count is the visible trace of Outlook wedging
        /// and being recovered from, which was previously indistinguishable from silence.
        /// </summary>
        public ComHostDiagnostics? ComHost { get; set; }
    }

    /// <summary>Index block of the outlook_health report (Phase 7; per-store rows merged from index_status in D37).</summary>
    public sealed class IndexHealthView
    {
        /// <summary>Active index provider ("OleDb"/"AdodbCom") or "unavailable: ..." when unreachable.</summary>
        public string Provider { get; set; } = string.Empty;

        /// <summary>Newest indexed mail timestamp across all stores (UTC).</summary>
        public DateTime? NewestIndexedUtc { get; set; }

        /// <summary>Age of the newest indexed mail in minutes.</summary>
        public double? AgeMinutes { get; set; }

        /// <summary>Per-store newest-indexed timestamps (absent when the index is unreachable).</summary>
        public IReadOnlyList<StoreStaleness>? PerStore { get; set; }

        /// <summary>WSearch service start mode from the registry: automatic|manual|disabled|unknown.</summary>
        public string WSearchStartMode { get; set; } = "unknown";

        /// <summary>Whether SearchIndexer.exe is running (null when the probe failed).</summary>
        public bool? IndexerProcessRunning { get; set; }
    }

    /// <summary>Audit-log block of the health report (Phase 7).</summary>
    public sealed class AuditHealthView
    {
        /// <summary>Audit log file path.</summary>
        public string Path { get; set; } = string.Empty;

        /// <summary>Whether an append handle could be opened (write ops fail-closed without it).</summary>
        public bool Writable { get; set; }

        /// <summary>Content-free failure reason when not writable.</summary>
        public string? Error { get; set; }
    }

    /// <summary>
    /// Tuning block of the health report (Phase 7): read straight from the
    /// HKCU\Software\OutlookAI\Tuning registry state the add-in maintains - the server
    /// stays decoupled from add-in code (R12/section 0.5.3).
    /// </summary>
    public sealed class TuningHealthView
    {
        /// <summary>True when the add-in's tuning state exists in the registry (add-in installed and initialized).</summary>
        public bool Managed { get; set; }

        /// <summary>Master toggle.</summary>
        public bool? Enabled { get; set; }

        /// <summary>Search-registry group toggle (D22 keys).</summary>
        public bool? SearchEnabled { get; set; }

        /// <summary>Full-caching group toggle (D25 keys).</summary>
        public bool? CachingEnabled { get; set; }

        /// <summary>OST-headroom group toggle (D25 caps).</summary>
        public bool? OstEnabled { get; set; }

        /// <summary>True when a tuning change still needs an Outlook restart to take effect.</summary>
        public bool? RestartNeeded { get; set; }

        /// <summary>Group-policy conflicts the reconciler backed off from (';'-joined; null when none).</summary>
        public string? PolicyConflicts { get; set; }

        /// <summary>Last reconcile timestamp (ISO 8601) as recorded by the add-in.</summary>
        public string? LastReconcileUtc { get; set; }

        /// <summary>
        /// EFFECTIVE Outlook UI search backend, read from the live registry - NOT from
        /// desired state (D22/D35): "local" (DisableServerAssistedSearch in force; the
        /// Outlook search box queries the same SystemIndex corpus agent search uses) or
        /// "server-assisted" (value absent/0; UI results are server-capped and
        /// differently ranked, so they can diverge from agent search). The policy hive
        /// wins over the user hive when both carry the value.
        /// </summary>
        public string? UiSearchBackend { get; set; }
    }

    /// <summary>
    /// How this server is registered with Claude Code (Phase 8 / D6 v2 / R11). Two
    /// independent sources, deliberately: the OBSERVED state, read here and now from
    /// ~/.claude.json and compared against the executable this very process is running
    /// from, plus what the add-in's reconcile last recorded in HKCU. They can disagree -
    /// that disagreement is exactly what makes drift visible.
    /// </summary>
    public sealed class McpRegistrationHealthView
    {
        /// <summary>
        /// "ok" (the registered command IS this executable), "drifted" (registered, but
        /// pointing somewhere else - typically a stale developer build path), "absent"
        /// (no outlookai entry), "unreadable" (config present but not parseable, so it is
        /// never rewritten) or "unknown" (config could not be examined).
        /// </summary>
        public string Status { get; set; } = "unknown";

        /// <summary>Executable this process is actually running from.</summary>
        public string? RunningFrom { get; set; }

        /// <summary>Command recorded under mcpServers.outlookai in ~/.claude.json (null when absent).</summary>
        public string? RegisteredCommand { get; set; }

        /// <summary>Status code from the add-in's last registration reconcile (null when the add-in never ran one).</summary>
        public string? AddInStatus { get; set; }

        /// <summary>True when the add-in's last reconcile had to repair the registration.</summary>
        public bool? AddInHealed { get; set; }

        /// <summary>Timestamp (ISO 8601) of the add-in's last registration reconcile.</summary>
        public string? AddInLastReconcileUtc { get; set; }

        /// <summary>Installed server path the add-in resolved, when it recorded one.</summary>
        public string? AddInResolvedServerPath { get; set; }
    }

    /// <summary>
    /// outlook_health tool outcome (Phase 7; index_status merged in D37): compact
    /// self-check of everything the server depends on plus the index freshness report.
    /// </summary>
    public sealed class HealthOutcome
    {
        /// <summary>"ok" when everything the server needs is available, else "degraded" (see Problems).</summary>
        public string Status { get; set; } = "ok";

        /// <summary>What is degraded, one compact line each (present only when Status != ok).</summary>
        public IReadOnlyList<string>? Problems { get; set; }

        /// <summary>Outlook process/COM state.</summary>
        public OutlookHealthView Outlook { get; set; } = new OutlookHealthView();

        /// <summary>SystemIndex + WSearch state (incl. per-store freshness).</summary>
        public IndexHealthView Index { get; set; } = new IndexHealthView();

        /// <summary>Audit log writability (write tools fail-closed without it).</summary>
        public AuditHealthView Audit { get; set; } = new AuditHealthView();

        /// <summary>Outlook tuning state summary (registry read).</summary>
        public TuningHealthView Tuning { get; set; } = new TuningHealthView();

        /// <summary>Whether Claude Code's MCP registration points at this executable (Phase 8).</summary>
        public McpRegistrationHealthView Registration { get; set; } = new McpRegistrationHealthView();

        /// <summary>Actionable freshness advice (distinct from Problems: guidance, not degradation).</summary>
        public IReadOnlyList<string>? Advice { get; set; }
    }

    /// <summary>show_search_results outcome (v3.MD L3).</summary>
    public sealed class ShowSearchResultsOutcome
    {
        /// <summary>The query now in Outlook's search box.</summary>
        public string Query { get; set; } = string.Empty;

        /// <summary>Scope the search ran with ("current_folder"/"subfolders"/"all_folders"/"all_outlook").</summary>
        public string Scope { get; set; } = string.Empty;

        /// <summary>ActiveExplorer().CurrentFolder.FolderPath right after Search (the UI may swap to a results view asynchronously).</summary>
        public string? ExplorerFolderPath { get; set; }

        /// <summary>Explorer window caption right after Search.</summary>
        public string? ExplorerCaption { get; set; }

        /// <summary>Always true on success - the search UI is on screen and populating.</summary>
        public bool Displayed { get; set; }

        /// <summary>
        /// Agent-facing advice (present when the displayed results may not match agent
        /// search - e.g. the Outlook UI search backend is server-assisted, D22/D35).
        /// </summary>
        public IReadOnlyList<string>? Advice { get; set; }
    }

    /// <summary>
    /// Per-item result of move_mail/archive_mail (D39). A move CHANGES the item's
    /// EntryID: <c>oldEntryId</c> is stale, <c>newEntryId</c> is the current identity,
    /// and <c>fromFolder</c> is the undo address (move back = move_mail with
    /// newEntryId and folder=fromFolder).
    /// </summary>
    public sealed class MoveItemView
    {
        /// <summary>Echo of the input id (hit id or EntryID) this result belongs to.</summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>True when the item was moved AND its audit line was written.</summary>
        /// <remarks>
        /// Both halves matter, and that is exactly why <see cref="Ok"/> is not enough on its
        /// own: a move that succeeded and could not be audited reports <c>Ok=false</c> over
        /// an item that really did move. Read <see cref="Outcome"/> before concluding
        /// anything from a false.
        /// </remarks>
        public bool Ok { get; set; }

        /// <summary>Failure reason, present only when not ok.</summary>
        /// <remarks>
        /// This used to be documented as "nothing was moved for this item", which the
        /// product contradicted in its own code twice over: the audit-failure message says
        /// "The item WAS moved...", and the per-item timeout message says the outcome is
        /// UNKNOWN. Both travelled through this field. The parenthesis is gone and
        /// <see cref="Outcome"/> carries the answer.
        /// </remarks>
        public string? Error { get; set; }

        /// <summary>
        /// What happened to this item, when it did not simply move: <c>unchanged</c> (it is
        /// where it was), <c>applied</c> (it MOVED - the failure is downstream of the move,
        /// and newEntryId/fromFolder are absent because the reporting step is what failed),
        /// or <c>unknown</c> (the move may or may not have taken effect - find the item
        /// before acting). Omitted when <see cref="Ok"/> is true, where it would only repeat
        /// the boolean.
        /// </summary>
        public string? Outcome { get; set; }

        /// <summary>Store the item lives in (moves are same-store).</summary>
        public string? Store { get; set; }

        /// <summary>Store-relative path the item came FROM - the undo address.</summary>
        public string? FromFolder { get; set; }

        /// <summary>Store-relative path the item is in now.</summary>
        public string? ToFolder { get; set; }

        /// <summary>EntryID before the move (stale after a successful move).</summary>
        public string? OldEntryId { get; set; }

        /// <summary>EntryID after the move - use this for follow-up read/open_in_outlook/undo.</summary>
        public string? NewEntryId { get; set; }
    }

    /// <summary>move_mail outcome (D39): same-store, audited, reversible moves.</summary>
    public sealed class MoveMailOutcome
    {
        /// <summary>Number of ids requested.</summary>
        public int Requested { get; set; }

        /// <summary>Number of items actually moved (audited).</summary>
        public int Moved { get; set; }

        /// <summary>Number of items that failed (see each item's error).</summary>
        public int Failed { get; set; }

        /// <summary>Echo of the store-relative target folder path.</summary>
        public string TargetFolder { get; set; } = string.Empty;

        /// <summary>
        /// Store-relative paths of folders created for this call (create_folder=true), when
        /// any - INCLUDING folders created by an item whose move then failed or was refused.
        /// Folders are made before the target guard runs, so a refused move can leave one
        /// behind; this used to be filled from successful moves alone and said nothing.
        /// This server cannot delete folders, so an unwanted one has to be removed in Outlook.
        /// </summary>
        public IReadOnlyList<string>? CreatedFolders { get; set; }

        /// <summary>Per-item results, input order.</summary>
        public IReadOnlyList<MoveItemView> Items { get; set; } = Array.Empty<MoveItemView>();

        /// <summary>Standing guidance (EntryID change/undo semantics), present when anything moved.</summary>
        public IReadOnlyList<string>? Advice { get; set; }
    }

    /// <summary>One store's designated Archive folder as archive_mail resolved it (D39).</summary>
    public sealed class ArchiveFolderView
    {
        /// <summary>Store display name.</summary>
        public string Store { get; set; } = string.Empty;

        /// <summary>Store-relative path of the designated Archive folder (localized name - e.g. Archive/Archiveren).</summary>
        public string Folder { get; set; } = string.Empty;

        /// <summary>Resolution mechanism ("outlookDefaultFolder" or "storeArchiveProperty").</summary>
        public string Via { get; set; } = string.Empty;
    }

    /// <summary>archive_mail outcome (D39): one-click-archive semantics, audited, reversible.</summary>
    public sealed class ArchiveMailOutcome
    {
        /// <summary>Number of ids requested.</summary>
        public int Requested { get; set; }

        /// <summary>Number of items archived (audited).</summary>
        public int Archived { get; set; }

        /// <summary>Number of items that failed (see each item's error).</summary>
        public int Failed { get; set; }

        /// <summary>The designated Archive folder per store involved (resolved, never guessed by name).</summary>
        public IReadOnlyList<ArchiveFolderView>? ArchiveFolders { get; set; }

        /// <summary>Per-item results, input order (toFolder = the store's Archive folder).</summary>
        public IReadOnlyList<MoveItemView> Items { get; set; } = Array.Empty<MoveItemView>();

        /// <summary>Standing guidance (EntryID change/undo semantics), present when anything moved.</summary>
        public IReadOnlyList<string>? Advice { get; set; }
    }
}
