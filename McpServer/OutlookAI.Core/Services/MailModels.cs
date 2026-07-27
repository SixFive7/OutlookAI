using System;
using System.Collections.Generic;

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

        /// <summary>Include indexed attachment-content entries (kind=document). Default true.</summary>
        public bool IncludeAttachmentHits { get; set; } = true;

        /// <summary>ONLY attachment-content entries (kind=document). Overrides <see cref="IncludeAttachmentHits"/>.</summary>
        public bool AttachmentHitsOnly { get; set; }

        /// <summary>Order results by size instead of date (big-mail discovery; index path only).</summary>
        public bool OrderBySizeDescending { get; set; }

        /// <summary>Maximum hits returned (compact payloads - T1-pinned caps in <see cref="MailService"/>).</summary>
        public int Top { get; set; } = MailService.SearchTopDefault;

        /// <summary>Snippet length per hit (0 disables snippets).</summary>
        public int SnippetChars { get; set; } = MailService.SnippetCharsDefault;
    }

    /// <summary>One agent-facing hit: compact triage payload (v3.MD sections 8/12).</summary>
    public sealed class HitSummary
    {
        /// <summary>Opaque hit id for read/save_attachment/thread ("h1", "h2", ...). Cached per server process.</summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>"index" (SystemIndex row), "live" (freshness gap-sweep result, D19) or "exhaustive" (COM scan).</summary>
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
    }

    /// <summary>Freshness gap-sweep diagnostics attached to (non-exhaustive) search results.</summary>
    public sealed class SweepInfo
    {
        /// <summary>Whether the sweep ran (false: COM unavailable - see Error).</summary>
        public bool Performed { get; set; }

        /// <summary>
        /// True when this result was served from the short-lived sweep cache (D34) - no
        /// COM call was made; the swept data is at most <see cref="SweepCache.DefaultTimeToLive"/>
        /// old. Omitted (null) when the sweep ran live.
        /// </summary>
        public bool? Cached { get; set; }

        /// <summary>Age of the cached sweep data in seconds (present only when Cached=true).</summary>
        public double? CacheAgeSeconds { get; set; }

        /// <summary>Sweep window start (UTC).</summary>
        public DateTime? GapStartUtc { get; set; }

        /// <summary>
        /// What the sweep covered, following the search scope (soak fix 13):
        /// <c>"folder"</c> = the searched folder (plus its subfolders when
        /// include_subfolders is on); <c>"default folders (Inbox, Sent Items, Deleted
        /// Items, Junk Email)"</c> = those folders in the searched store, or in every
        /// store when the search is not store-scoped. Lets an agent see the freshness
        /// coverage of its query.
        /// </summary>
        public string? Scope { get; set; }

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

        /// <summary>Folders skipped (unresolvable, unenumerable, or past the folder cap of a scoped sweep).</summary>
        public int FoldersSkipped { get; set; }

        /// <summary>
        /// Folders whose item enumeration FAILED, so they have no freshness coverage at
        /// all. Until soak fix 15 these were counted as successfully swept.
        /// </summary>
        public int FoldersFailed { get; set; }

        /// <summary>
        /// Folders where the per-folder item cap (<c>SweepPerFolderCap</c>) truncated the
        /// window. The sweep reads newest-first, so the OLDEST fresh mail in those folders
        /// was dropped. Null when nothing was truncated.
        /// </summary>
        public IReadOnlyList<string>? ItemCappedFolders { get; set; }

        /// <summary>
        /// True when the scoped sweep hit <c>MaxScopedSweepFolders</c> and stopped walking
        /// the subtree, so folders past the cut-off were never visited.
        /// </summary>
        public bool? FolderCapReached { get; set; }

        /// <summary>Items in the window before term filtering.</summary>
        public int ItemsSeen { get; set; }

        /// <summary>Swept items dropped as already present in the index results.</summary>
        public int Duplicates { get; set; }

        /// <summary>Sweep wall-clock cost.</summary>
        public long ElapsedMs { get; set; }

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

        /// <summary>Scan wall-clock cost.</summary>
        public long ElapsedMs { get; set; }
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
    }

    /// <summary>Index staleness snapshot attached to search results.</summary>
    public sealed class StalenessInfo
    {
        /// <summary>Newest indexed DateReceived (UTC) across the searched scope.</summary>
        public DateTime? NewestIndexedUtc { get; set; }

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

    /// <summary>thread outcome.</summary>
    public sealed class ThreadOutcome
    {
        /// <summary>Conversation id the thread was resolved for.</summary>
        public string? ConversationId { get; set; }

        /// <summary>"index" (ConversationID query) or "com" (Outlook Conversation walk).</summary>
        public string Source { get; set; } = "index";

        /// <summary>Thread members, oldest first.</summary>
        public IReadOnlyList<HitSummary> Hits { get; set; } = Array.Empty<HitSummary>();

        /// <summary>
        /// True when the 'top' cap cut the member list - the conversation HAS more
        /// members (over-fetch-by-one, same contract as search.truncated).
        /// </summary>
        public bool Truncated { get; set; }

        /// <summary>Wall-clock cost of the thread lookup.</summary>
        public long ElapsedMs { get; set; }
    }

    /// <summary>Per-store index-freshness row of the outlook_health report.</summary>
    public sealed class StoreStaleness
    {
        /// <summary>Store display name.</summary>
        public string Store { get; set; } = string.Empty;

        /// <summary>Newest indexed DateReceived under that store's scope (UTC).</summary>
        public DateTime? NewestIndexedUtc { get; set; }
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

        /// <summary>This page's folders of the store (full tree, stable traversal order).</summary>
        public IReadOnlyList<FolderView> Folders { get; set; } = Array.Empty<FolderView>();
    }

    /// <summary>list_folders outcome (full tree, offset-paged in stable traversal order).</summary>
    public sealed class FoldersOutcome
    {
        /// <summary>Folder trees per store (this page).</summary>
        public IReadOnlyList<StoreFoldersView> Stores { get; set; } = Array.Empty<StoreFoldersView>();

        /// <summary>Total folders in the full traversal (all pages).</summary>
        public int FolderTotal { get; set; }

        /// <summary>Echo of a non-zero requested offset (omitted for the first page).</summary>
        public int? Offset { get; set; }

        /// <summary>True when more folders exist beyond this page - continue with offset=nextOffset.</summary>
        public bool Truncated { get; set; }

        /// <summary>The offset that continues the listing (present only when truncated).</summary>
        public int? NextOffset { get; set; }
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

        /// <summary>True while the add-in installer holds the OutlookAISetup mutex (D17: COM tools retry later).</summary>
        public bool InstallerMutexHeld { get; set; }

        /// <summary>True when this server holds a COM session that Outlook ANSWERED just now (probed liveness, SF-1).</summary>
        public bool ComConnected { get; set; }

        /// <summary>Store count reachable over COM (null when Outlook is not running - health never starts it).</summary>
        public int? StoresReachable { get; set; }

        /// <summary>Reachable store display names.</summary>
        public IReadOnlyList<string>? Stores { get; set; }
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

        /// <summary>True when the item was moved (and its audit line written).</summary>
        public bool Ok { get; set; }

        /// <summary>Failure reason (present only when not ok; nothing was moved for this item).</summary>
        public string? Error { get; set; }

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

        /// <summary>Store-relative paths of folders created for this call (create_folder=true), when any.</summary>
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
