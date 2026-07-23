using System;
using System.Collections.Generic;

namespace OutlookAI.Core.Services
{
    /// <summary>Search freshness mode (v3.MD D19). Exhaustive arrives in Phase 3.</summary>
    public enum SearchMode
    {
        /// <summary>Index only (8-60 ms; works with Outlook closed, results may be stale).</summary>
        Fast = 0,

        /// <summary>Index + COM gap-sweep of items newer than the newest-indexed timestamp. The default.</summary>
        Fresh = 1,
    }

    /// <summary>Parameters for one search call (mirrors the MCP tool arguments).</summary>
    public sealed class SearchRequest
    {
        /// <summary>Free-text query; whitespace-separated terms are ANDed. Optional.</summary>
        public string? Query { get; set; }

        /// <summary>Freshness mode. Default fresh (D19).</summary>
        public SearchMode Mode { get; set; } = SearchMode.Fresh;

        /// <summary>Store display name to scope to (as returned by list_accounts).</summary>
        public string? Store { get; set; }

        /// <summary>Store-relative folder path ('/'-separated) to scope to; requires <see cref="Store"/>.</summary>
        public string? Folder { get; set; }

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

        /// <summary>Maximum hits returned (compact payloads: default 25, cap 100).</summary>
        public int Top { get; set; } = 25;

        /// <summary>Snippet length per hit (0 disables snippets).</summary>
        public int SnippetChars { get; set; } = 200;
    }

    /// <summary>One agent-facing hit: compact triage payload (v3.MD sections 8/12).</summary>
    public sealed class HitSummary
    {
        /// <summary>Opaque hit id for read/save_attachment/thread ("h1", "h2", ...). Cached per server process.</summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>"index" (SystemIndex row) or "live" (COM gap-sweep result, D19 fresh mode).</summary>
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

    /// <summary>Gap-sweep diagnostics attached to fresh-mode results.</summary>
    public sealed class SweepInfo
    {
        /// <summary>Whether the sweep ran (false: COM unavailable - see Error).</summary>
        public bool Performed { get; set; }

        /// <summary>Sweep window start (UTC).</summary>
        public DateTime? GapStartUtc { get; set; }

        /// <summary>Default folders swept.</summary>
        public int FoldersSwept { get; set; }

        /// <summary>Default folders skipped (unresolvable).</summary>
        public int FoldersSkipped { get; set; }

        /// <summary>Items in the window before term filtering.</summary>
        public int ItemsSeen { get; set; }

        /// <summary>Swept items dropped as already present in the index results.</summary>
        public int Duplicates { get; set; }

        /// <summary>Sweep wall-clock cost.</summary>
        public long ElapsedMs { get; set; }

        /// <summary>Content-free error when the sweep could not run.</summary>
        public string? Error { get; set; }
    }

    /// <summary>Index staleness snapshot attached to search results.</summary>
    public sealed class StalenessInfo
    {
        /// <summary>Newest indexed DateReceived (UTC) across the searched scope.</summary>
        public DateTime? NewestIndexedUtc { get; set; }

        /// <summary>Age of the newest indexed mail in minutes.</summary>
        public double? AgeMinutes { get; set; }

        /// <summary>Whether OUTLOOK.EXE is running (the index only advances while it runs).</summary>
        public bool OutlookRunning { get; set; }
    }

    /// <summary>Search outcome (search tool payload).</summary>
    public sealed class SearchOutcome
    {
        /// <summary>Mode that ran ("fast"/"fresh").</summary>
        public string Mode { get; set; } = "fresh";

        /// <summary>Merged hits, newest first.</summary>
        public IReadOnlyList<HitSummary> Hits { get; set; } = Array.Empty<HitSummary>();

        /// <summary>Index query wall-clock cost.</summary>
        public long IndexElapsedMs { get; set; }

        /// <summary>Sweep diagnostics (fresh mode only).</summary>
        public SweepInfo? Sweep { get; set; }

        /// <summary>Staleness self-report (R7/D19).</summary>
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

        /// <summary>To/Cc/Bcc recipients.</summary>
        public IReadOnlyList<RecipientView> Recipients { get; set; } = Array.Empty<RecipientView>();

        /// <summary>Plain-text body (possibly truncated - check BodyTruncated).</summary>
        public string Body { get; set; } = string.Empty;

        /// <summary>Full body length in characters before truncation.</summary>
        public long BodyTotalChars { get; set; }

        /// <summary>True when Body was cut at max_body_chars.</summary>
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

        /// <summary>Attachments (save via save_attachment with the same id + index).</summary>
        public IReadOnlyList<AttachmentView> Attachments { get; set; } = Array.Empty<AttachmentView>();

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

        /// <summary>Wall-clock cost of the thread lookup.</summary>
        public long ElapsedMs { get; set; }
    }

    /// <summary>Per-store staleness row for index_status.</summary>
    public sealed class StoreStaleness
    {
        /// <summary>Store display name.</summary>
        public string Store { get; set; } = string.Empty;

        /// <summary>Newest indexed DateReceived under that store's scope (UTC).</summary>
        public DateTime? NewestIndexedUtc { get; set; }
    }

    /// <summary>index_status outcome.</summary>
    public sealed class IndexStatusOutcome
    {
        /// <summary>Active index query provider ("OleDb"/"AdodbCom") or "unavailable".</summary>
        public string Provider { get; set; } = string.Empty;

        /// <summary>Whether OUTLOOK.EXE is running.</summary>
        public bool OutlookRunning { get; set; }

        /// <summary>True while the add-in installer holds the OutlookAISetup mutex (D17: retry later).</summary>
        public bool InstallerMutexHeld { get; set; }

        /// <summary>Newest indexed mail timestamp across all stores (UTC).</summary>
        public DateTime? NewestIndexedUtc { get; set; }

        /// <summary>Age of the newest indexed mail in minutes.</summary>
        public double? IndexAgeMinutes { get; set; }

        /// <summary>Per-store newest-indexed timestamps.</summary>
        public IReadOnlyList<StoreStaleness>? PerStore { get; set; }

        /// <summary>Actionable freshness advice.</summary>
        public IReadOnlyList<string> Advice { get; set; } = Array.Empty<string>();
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

        /// <summary>Folders up to the requested depth.</summary>
        public IReadOnlyList<FolderView> Folders { get; set; } = Array.Empty<FolderView>();
    }

    /// <summary>list_folders outcome.</summary>
    public sealed class FoldersOutcome
    {
        /// <summary>Folder trees per store.</summary>
        public IReadOnlyList<StoreFoldersView> Stores { get; set; } = Array.Empty<StoreFoldersView>();

        /// <summary>True when the folder cap cut the listing.</summary>
        public bool Truncated { get; set; }
    }
}
