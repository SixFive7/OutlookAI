using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using OutlookAI.ComHost.Client;
using OutlookAI.ComHost.Supervision;
using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;

namespace OutlookAI.McpServer.Tools;

/// <summary>
/// Process-wide service holder: ONE MailService and ONE COM-host supervisor per server
/// process. Created lazily so starting the server never touches the index or Outlook by
/// itself.
/// <para>
/// Note what this process does NOT hold any more: a COM session, a pumped STA thread, or
/// any Outlook reference. Those live in the OutlookAI.ComHost child, which exists so a
/// wedged Outlook call can be reclaimed by killing it (Docs/com-host.md). Everything here
/// is either pure computation or a bounded round trip.
/// </para>
/// </summary>
internal static class ServerRuntime
{
    private static readonly Lazy<RemoteComGateway> LazyGateway = new(
        () => new RemoteComGateway(allowStartingOutlook: true), LazyThreadSafetyMode.ExecutionAndPublication);

    private static readonly Lazy<MailService> LazyService = new(
        () => new MailService(LazyGateway.Value), LazyThreadSafetyMode.ExecutionAndPublication);

    internal static MailService Service => LazyService.Value;

    /// <summary>The COM-host gateway, for health reporting that must not go through the service layer.</summary>
    internal static RemoteComGateway Gateway => LazyGateway.Value;
}

/// <summary>
/// The MCP tool surface (v3.MD section 0.5): search/thread/read/save_attachment,
/// move_mail/archive_mail (D39), list_accounts/list_folders/list_signatures,
/// manage_signature, the show-me tools, the draft tools, send, and outlook_health.
/// Payloads are compact JSON (camelCase, nulls omitted - section 12 discipline);
/// domain failures come back as an {"error": ...} object instead of protocol faults
/// so agents can react.
/// </summary>
[McpServerToolType]
public static class OutlookTools
{
    private static readonly JsonSerializerOptions Json = new(JsonSerializerDefaults.Web)
    {
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        WriteIndented = false,
    };

    // The tool description is capped by the CLIENT, not by the protocol: Claude Code cuts
    // a tool description at 2048 UTF-16 code units, positionally and silently, so a
    // description that outgrows the cap loses its tail without any diagnostic. That is
    // measured, not documented-and-hoped: 2026-08-18, client 2.1.234, read off the client's
    // own outbound request (DescriptionBudgetCiTests.ClientTruncationBudget carries the
    // numbers and the caveats). This description used to run to 3912 characters and the cut
    // landed mid-word inside DEGRADED RESULTS, taking the rest of that instruction plus the
    // folder-scope, results and exhaustive sections with it. What stays here is what a
    // caller needs BEFORE the call (purpose, matching semantics) plus the one thing that is
    // a shipped behaviour the agent must act on (degraded=true is a SUCCESS the user has to
    // be told about). Everything that describes the ANSWER moved to where the answer already
    // reports it - the advice/scope/sweep blocks - and everything that is per-argument
    // how-to-call detail moved onto the arguments. The same measurement showed the cut is
    // PER STRING with no per-tool bucket, so that move really did buy room rather than
    // shuffling text between two halves of one budget; the arguments carry their own budget,
    // which is a house limit rather than a client one. DescriptionBudgetCiTests measures all
    // of it from the wire.
    //
    // DEGRADED RESULTS was rewritten when freshness gained its third value: it used to
    // state degraded=true and freshness="index-only" as one fact, which stopped being true
    // the moment a sweep that RAN but covered part of its scope also earned degraded=true.
    // A sentence that pairs the flag with one specific value teaches the agent to read the
    // value instead of the flag, so it now leads with what the flag MEANS and lists the two
    // values under it. Rewritten inside the existing budget, not on top of it: the paragraph
    // paid for the new state by dropping the parenthetical listing Outlook's states, which
    // advice already spells out at runtime, and two words elsewhere. Measured from the wire
    // at 1798 before and 1791 after.
    [McpServerTool(Name = "search")]
    [Description("Search locally indexed Outlook mail across all accounts, folders and delegate mailboxes. "
        + "Sub-second and cheap: iterate with refined terms instead of pulling large result sets.\n\n"
        + "MATCHING: query terms are whitespace-separated and ANDed; each term matches whole words, and the terms "
        + "may land in different parts of the mail - one in the subject, another in the body (search_in narrows "
        + "this to one of them). Body matching also covers attachment text - see include_attachment_hits. Append "
        + "* for prefix match (haproxy*). Allowed characters: letters, digits and @.-_'+ ; omit query to "
        + "filter only by from/to/date/flags. Sender and recipients are NOT matched by query terms - use "
        + "from / to.\n\n"
        + "FRESHNESS: results always include mail that arrived after the last index update: the server sweeps it "
        + "live through Outlook (started headless when needed) and merges it in; the sweep block reports what "
        + "that covered.\n\n"
        + "DEGRADED RESULTS: degraded=true means NOT fully fresh - the live check could not run "
        + "(freshness=\"index-only\") or it covered only part of its scope (freshness=\"partial\"; "
        + "sweep.coverageGaps says which). Everything already "
        + "indexed is present and correct; recent mail from the uncovered part may be missing. This is a "
        + "SUCCESSFUL result, not an error - but SAY SO TO THE USER when degraded is true, because an answer that "
        + "looks complete and quietly is not is worse than a visible failure. advice spells out the reason and the "
        + "remedy; outlook_health gives the full picture.\n\n"
        + "RESULTS: each hit's id feeds read, thread, save_attachment, open_in_outlook, move_mail and "
        + "archive_mail (ids last this session). Read advice whenever present and relay what "
        + "concerns the user: every partial "
        + "result, cap, skipped folder, widened scope and freshness gap is reported there, alongside the scope "
        + "and sweep blocks.")]
    public static async Task<CallToolResult> Search(
        [Description("Free-text terms, whitespace-separated, ANDed. Each term may match in the subject or the body (see search_in). Letters/digits plus @.-_'+ only; trailing * for prefix. Omit to filter by sender/date only.")]
        string? query = null,
        [Description("Which part of the mail the query terms must match: subject_and_body (default), subject, or "
            + "body. Narrow it when a term is noisy in one of them.")]
        string? search_in = null,
        [Description("exhaustive=true bypasses the index and scans folders through Outlook instead - requires "
            + "store plus folder and/or after, is far slower, and matches whole words in subject and body only "
            + "(no attachment text). Default false: index + freshness sweep. It follows include_subfolders like "
            + "the other modes, so a folder scope walks the subtree - which on a big subtree can hit the 105 s "
            + "budget; pass include_subfolders=false to scan just the named folder, and check "
            + "foldersScanned/foldersSkipped plus advice for partial coverage. Use it when the index looks stale "
            + "or wrong, or when completeness matters more than speed.")]
        bool exhaustive = false,
        [Description("Store display name to search in (see list_accounts). Omit for all stores (required when "
            + "exhaustive=true). The scope is honoured exactly and never silently widened. A store the profile has "
            + "but Windows Search has not indexed (a PST, a fresh install, indexing off or still building) still "
            + "searches: the index tier is skipped, the freshness sweep is the only tier, and the answer says so "
            + "with index.storeNotIndexed=true, sweep.storesWithoutIndex naming it, coverageGaps no_index_frontier "
            + "and degraded=true - only the recent window is covered, so use exhaustive=true with folder and/or "
            + "after to read that store in full. A name the profile does not have is refused, and the error lists "
            + "the stores that do exist.")] string? store = null,
        [Description("Store-relative folder path (from list_folders), e.g. 'Inbox' or 'Projects/2026'. Requires "
            + "store. Includes its subfolders unless include_subfolders=false. Delegate/shared mailboxes are "
            + "indexed WITHOUT their folder nesting, so a folder scope there matches by folder NAME: the search "
            + "widens to the whole delegate mailbox when the subfolder set cannot be built, and can return mail "
            + "from a second folder with the same name elsewhere in that mailbox - advice says so whenever it "
            + "happens, and the scope block reports what was actually covered. Setting folder also aims the "
            + "freshness sweep at that folder and its subfolders; without it the sweep covers Inbox, Sent Items, "
            + "Deleted Items and Junk Email of the store(s) in scope - those four folders only, not their "
            + "subfolders - so for brand-new mail filed anywhere else, pass store + folder.")] string? folder = null,
        [Description("Whether folder covers its subfolders. Default true, in every mode. Set false to search that "
            + "one folder only - also the cheap way to keep an exhaustive scan bounded.")]
        bool include_subfolders = true,
        [Description("Sender filter: address or name fragment (index-backed).")] string? from = null,
        [Description("Recipient (To/Cc) filter: address fragment (index-backed).")] string? to = null,
        [Description("Only mail received at/after this instant (ISO 8601, e.g. 2026-07-01 or 2026-07-01T08:00:00Z).")] string? after = null,
        [Description("Only mail received before this instant (ISO 8601).")] string? before = null,
        [Description("true = unread mail only.")] bool? unread_only = null,
        [Description("Filter on attachment presence.")] bool? has_attachments = null,
        [Description("Include attachment-CONTENT matches. Body matching covers EVERY attachment type - documents, "
            + "images, embedded messages, calendar invites, media - and those matches come back as separate hits "
            + "with isAttachmentHit=true (read on such a hit opens the parent mail). Default true. Setting it "
            + "false drops only those hits; ordinary subject/body matches, including the freshness sweep's, are "
            + "unaffected. The sweep matches subject and body only - attachment text is matched by the index "
            + "alone, so a term living only inside an attachment is findable only once that mail is indexed.")]
        bool include_attachment_hits = true,
        [Description("Max hits (1-100, default 25). Keep small - iterate instead. truncated=true in the response "
            + "means more matches exist beyond top: narrow with store/folder/from/after rather than raising it.")]
        int top = 25,
        [Description("Snippet length per hit (0-1000, default 200; 0 = no snippets).")] int snippet_chars = 200,
        CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () =>
        {
            SearchRequest request = new()
            {
                Query = query,
                SearchIn = SearchInValues.Parse(search_in),
                Exhaustive = exhaustive,
                Store = store,
                Folder = folder,
                IncludeSubfolders = include_subfolders,
                From = from,
                To = to,
                AfterUtc = ParseUtc(after, "after"),
                BeforeUtc = ParseUtc(before, "before"),
                UnreadOnly = unread_only,
                HasAttachments = has_attachments,
                IncludeAttachmentHits = include_attachment_hits,
                Top = top,
                SnippetChars = snippet_chars,
            };
            return ServerRuntime.Service.Search(request);
        });
    }

    // FRESHNESS was added when thread gained the contract search already had (gap C1): it
    // used to walk Outlook only when the index held nothing, so one indexed row hid every
    // reply newer than the index frontier, and the payload had no degraded/freshness field
    // to say so. The paragraph pays for itself out of the old text - the sentence explaining
    // that the COM walk is a FALLBACK is gone, because it no longer is one - and the whole
    // description still sits far inside the 2048-code-unit client truncation budget that
    // DescriptionBudgetCiTests measures from the wire (measured client behaviour, 2.1.234,
    // 2026-08-18 - see DescriptionBudgetCiTests.ClientTruncationBudget).
    [McpServerTool(Name = "thread")]
    [Description("Fetch the full conversation of a mail. Two complementary lookup keys - pass BOTH when available: "
        + "conversation_id is the fast index path (every search hit already carries it, no locate cost), and id anchors "
        + "the LIVE check - COM cannot look up a conversation by id string, it needs a concrete mail item to walk "
        + "Outlook's conversation graph from.\n\n"
        + "FRESHNESS: given id, the indexed members and a live walk of Outlook's own conversation are merged, so replies "
        + "that arrived after the last index update are included. Given conversation_id alone there is no live check: "
        + "degraded=true with freshness=\"index-only\", and the newest replies may be missing - SAY SO TO THE USER, or "
        + "call again with id. degraded is also true when the live walk covered only part of the conversation "
        + "(freshness=\"partial\"; live.coverageGaps says which, advice says what to do).\n\n"
        + "Members are oldest-first; truncated=true means the conversation has more members than 'top'.")]
    public static async Task<CallToolResult> Thread(
        [Description("ConversationId from a search hit or read result - the fast index path. Pass when you have it.")] string? conversation_id = null,
        [Description("Hit id (e.g. h12) or EntryID of any mail in the conversation - anchors the live conversation-graph walk. Pass it whenever you have one: without it the answer is index-only.")] string? id = null,
        [Description("Store display name to scope the index lookup (faster). A name that does not resolve widens the lookup to the whole profile and reports scopeWidened:true.")] string? store = null,
        [Description("Max thread members (default 50, max 200). It caps the live walk too, so raise it when live.coverageGaps reports member_cap.")] int top = 50,
        CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.Thread(conversation_id, id, store, top));
    }

    [McpServerTool(Name = "read")]
    [Description("Read one mail in full by hit id (from search/thread) or EntryID: plain-text body with truncation flags and true total size, "
        + "sender/recipients with SMTP addresses, attachment list, conversation id. For an attachment-content hit this opens the PARENT mail. "
        + "Long bodies page cheaply: when bodyTruncated=true, call again with body_offset = bodyOffset + body.length to CONTINUE reading - "
        + "the next window is served from the already-extracted body, not re-read from the start. "
        + "The plain text HIDES layout: to check how a mail actually renders - or to verify a draft you just created with body_html - "
        + "pass include_html=true and inspect the stored HTML. "
        + "Needs Outlook (starts it if allowed). First read of an index hit locates the item (up to a few seconds); repeats are cached. "
        + "Works on any EntryID, including a draft you just created (drafts are not in the search index; pass the draft's entryId directly).")]
    public static async Task<CallToolResult> Read(
        [Description("Hit id (e.g. h12) or full EntryID hex.")] string id,
        [Description("Body window size in characters (default 20000; 0 = metadata only). bodyTruncated=true means more body exists beyond the window; bodyTotalChars is the full size.")] int max_body_chars = MailService.BodyCharsDefault,
        [Description("Include raw transport headers (capped at 8 KB). Default false.")] bool include_headers = false,
        [Description("Start of the body window in characters (default 0). Use the previous read's bodyOffset + body.length to continue a long body.")] int body_offset = 0,
        [Description("Also return the stored HTML body (Outlook's own HTMLBody) as bodyHtml - the ONLY way to verify formatting, "
            + "signature placement and quoted-thread boundaries, all of which the plain-text body collapses. Use it to check a draft "
            + "you created with body_html. The HTML has its own budget (max_html_chars) and is returned from its start; "
            + "bodyHtmlTotalChars gives the true size and bodyHtmlTruncated says whether it was cut. Default false (HTML is bulky).")]
        bool include_html = false,
        [Description("Budget for bodyHtml in characters (default 100000, max 500000; 0 = omit it). Only used with include_html. "
            + "It is deliberately larger than max_body_chars: Outlook puts ~40000 characters of stylesheet BEFORE the message "
            + "content, so a small window would show CSS instead of your text.")]
        int max_html_chars = MailService.HtmlCharsDefault,
        CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.Read(
            id, max_body_chars, include_headers, MailService.HeaderCharsDefault, body_offset, include_html, max_html_chars));
    }

    [McpServerTool(Name = "save_attachment")]
    [Description("Save one attachment of a mail to disk so you can open/read the file yourself. "
        + "Use the attachment 'index' from a read result. Never overwrites - existing names get a numeric suffix. Returns the absolute path.")]
    public static async Task<CallToolResult> SaveAttachment(
        [Description("Hit id or EntryID of the mail (for attachment-content hits: the hit itself).")] string id,
        [Description("1-based attachment index from read's attachments list.")] int attachment_index,
        [Description("Absolute target directory. Default: %LOCALAPPDATA%\\OutlookAI\\scratch\\attachments.")] string? target_dir = null,
        CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.SaveAttachment(id, attachment_index, target_dir));
    }

    [McpServerTool(Name = "move_mail")]
    [Description("MOVE mail to another folder WITHIN ITS OWN account/store - refile into a project folder, or restore items "
        + "from Deleted Items. Content-preserving, fully audited and REVERSIBLE: each result carries fromFolder plus "
        + "oldEntryId/newEntryId, so any move can be undone by calling move_mail again with newEntryId and folder=fromFolder. "
        + "Takes 1-50 ids (hit ids from search/thread, or EntryIDs); the target folder is a store-relative path (see "
        + "list_folders). SAME-STORE ONLY: each item moves within the store it lives in - cross-store moves are rejected "
        + "with a clear per-item error (run one call per store). Moving CHANGES an item's EntryID: use newEntryId afterwards; "
        + "old hit ids/index rows go stale briefly (re-run search). Moving to Deleted Items or the Outbox is refused - this "
        + "server cannot delete mail. Needs Outlook (starts it headless if needed); never opens windows.")]
    public static async Task<CallToolResult> MoveMail(
        [Description("1-50 hit ids (e.g. h12) or full EntryID hex strings. Each item is moved within its own store.")]
        string[] ids,
        [Description("Store-relative target folder path (from list_folders), e.g. 'Archive/2026' or 'Projects/Acme'.")]
        string folder,
        [Description("Create the target folder (including missing parents) when it does not exist. Default false.")]
        bool create_folder = false,
        [Description("Optional store display name (see list_accounts): when given, items living in a DIFFERENT store fail "
            + "with a cross-store error instead of moving. Omit to move each item within its own store.")]
        string? store = null,
        CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.MoveMail(ids, folder, create_folder, store));
    }

    [McpServerTool(Name = "archive_mail")]
    [Description("ARCHIVE mail with Outlook's own one-click-archive semantics: each item is moved to ITS OWN account's "
        + "DESIGNATED Archive folder - exactly the folder the Archive button (Backspace), mobile swipe-archive and Outlook "
        + "on the web use. The folder is resolved per store from the mailbox designation (localization-proof - e.g. a Dutch "
        + "mailbox's 'Archiveren' - never guessed by name); if a store has no designated Archive folder the item fails with "
        + "a clear error and NOTHING is created. Takes 1-50 ids (hit ids or EntryIDs), which may span accounts - each goes "
        + "to its own account's Archive. Content-preserving, fully audited and REVERSIBLE like move_mail: results carry "
        + "fromFolder + oldEntryId/newEntryId (undo = move_mail with newEntryId and folder=fromFolder). Archiving changes "
        + "EntryIDs; re-run search for fresh ids. Needs Outlook (starts it headless if needed); never opens windows.")]
    public static async Task<CallToolResult> ArchiveMail(
        [Description("1-50 hit ids (e.g. h12) or full EntryID hex strings; may span accounts.")]
        string[] ids,
        CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.ArchiveMail(ids));
    }

    // This tool's whole job is to RETURN a rich payload, so the description only has to get
    // an agent to call it and to read the answer. Everything that explained what a value
    // MEANS moved out to where the payload already states it: outlook.state carries the
    // liveness wording itself, problems spells out a hung Outlook, a COM host restart and
    // its lastFailure, a stopped WSearch and an unwritable audit log, and advice carries the
    // index-freshness guidance including that search covers a gap with its own sweep. Kept
    // deliberately small because the client truncates descriptions at 2048 UTF-16 code units
    // silently and mid-sentence (measured 2026-08-18 against client 2.1.234); this one had
    // reached 80% of that, and a later insertion had already landed mid-sentence inside the
    // comHost paragraph, splitting it in two. DescriptionBudgetCiTests measures it from the
    // wire.
    //
    // officeVersion follows that same rule and is why it is only NAMED here: the field costs
    // eight words in the coverage list plus four in the degradation list, and the sentence
    // explaining what an absent one implies lives in HealthReporting.NoOfficeVersionProblem,
    // which the payload already carries whenever it applies. Measured after that addition:
    // 1445 units of the 2048 client cap, 91 below the 1536 warn line.
    [McpServerTool(Name = "outlook_health")]
    [Description("One-call, read-only health and freshness report covering everything this server depends on: Outlook "
        + "itself (running, installed version, its Office registry major as officeVersion, probed COM session liveness, "
        + "reachable stores, whether it runs headless "
        + "- tray icon only, no window, which a normal Outlook launch promotes - and the state/responding pair Windows "
        + "itself answers for rather than a call attempt), the helper process this server drives Outlook through and "
        + "can restart (comHost: state, pid, restartCount, lastFailure), index freshness globally and per store, the "
        + "Windows Search (WSearch) service, audit-log writability, the OutlookAI tuning state incl. the effective UI "
        + "search backend, and the add-in installer mutex.\n\n"
        + "Attaches only to an Outlook that is ALREADY running and NEVER starts one, and it is bounded (the COM probe "
        + "gives up after 5 s) - so it answers even while Outlook is wedged, which is exactly when it is worth "
        + "calling.\n\n"
        + "READING IT: status=ok means every dependency is available; otherwise status=degraded and problems names each "
        + "degradation in words - Outlook not responding (anything needing it is then refused immediately rather than "
        + "left waiting, and restarting Outlook clears that), a COM host restart and what wedged it, a stopped WSearch, "
        + "an unwritable audit log, no supported Office version (registry answers then read EMPTY). advice carries the "
        + "freshness guidance. Read both and relay whatever concerns the "
        + "user.")]
    public static async Task<CallToolResult> OutlookHealth(CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.Health());
    }

    [McpServerTool(Name = "list_accounts")]
    [Description("List the profile's mail accounts and ALL stores (accounts, delegate/shared caches, archives) with flags: "
        + "isDelegate, onlineOnly/locallySearchable (server-only stores like Online Archives are invisible to local search), "
        + "and inLocalIndex. Store display names are the 'store' argument for search/list_folders.")]
    public static async Task<CallToolResult> ListAccounts(CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.ListAccounts());
    }

    [McpServerTool(Name = "list_folders")]
    [Description("List the FULL folder tree(s) with item/unread counts. Folder paths feed the search tool's 'folder' argument. "
        + "Traversal order is stable: stores sorted by display name, then depth-first with sibling folders sorted by name. "
        + "One call returns up to 1000 folders (virtually always the whole tree); truncated=true means more exist - "
        + "continue with offset=nextOffset to page the remainder in the same stable order.")]
    public static async Task<CallToolResult> ListFolders(
        [Description("Store display name (see list_accounts). Omit for all stores.")] string? store = null,
        [Description("Folders to skip in the stable traversal (default 0). Use the previous result's nextOffset to continue a truncated listing.")] int offset = 0,
        CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.ListFolders(store, offset));
    }

    [McpServerTool(Name = "list_signatures")]
    [Description("List the installed Outlook email signatures: each name plus a short plain-text excerpt (use the excerpt to "
        + "detect a signature's language and purpose), and per-account default assignments where the profile registry records "
        + "them (missing = unknown, e.g. no signature configured or roaming-managed - never guessed). Use this to pick the "
        + "BEST signature for a draft via the draft tools' 'signature' parameter - e.g. match the language of the "
        + "recipient/thread. Read-only, never starts Outlook.")]
    public static async Task<CallToolResult> ListSignatures(CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.ListSignatures());
    }

    [McpServerTool(Name = "manage_signature")]
    [Description("Create, update or DELETE an Outlook email signature by writing its file set (.htm + .txt + .rtf) in the "
        + "Signatures folder. DESTRUCTIVE: 'delete' permanently removes the signature and 'update' replaces its content - "
        + "before either, the previous files are automatically backed up under %LOCALAPPDATA%\\OutlookAI\\signature-backups "
        + "and the backup path is returned as backupPath; double-check the name (see list_signatures) before deleting. "
        + "For create/update supply body_text and/or body_html - the missing renditions are derived. Deleting a signature "
        + "also clears per-account default assignments that referenced it. Optional set_default_for records the signature "
        + "as an account's default (new mail, replies, or both) in the Outlook profile; Outlook picks that up at its next "
        + "start. Never starts or touches Outlook itself; every operation is audit-logged.")]
    public static async Task<CallToolResult> ManageSignature(
        [Description("'create' | 'update' | 'delete'.")] string action,
        [Description("Signature name (as shown by list_signatures and Outlook's signature pickers).")] string name,
        [Description("Plain-text signature body (create/update). When omitted it is derived from body_html.")]
        string? body_text = null,
        [Description("HTML signature body (create/update), fragment or full document. When omitted it is derived from body_text.")]
        string? body_html = null,
        [Description("Optionally record the signature as an account default: {\"account\": SMTP address from list_accounts, "
            + "\"scope\": \"new\"|\"reply\"|\"both\"}. Not allowed with delete.")]
        SetDefaultForArg? set_default_for = null,
        CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.ManageSignature(new ManageSignatureRequest
        {
            Action = action,
            Name = name,
            BodyText = body_text,
            BodyHtml = body_html,
            DefaultForAccount = set_default_for?.Account,
            DefaultForScope = set_default_for?.Scope,
        }));
    }

    /// <summary>manage_signature's set_default_for argument object.</summary>
    public sealed class SetDefaultForArg
    {
        /// <summary>Account SMTP address (see list_accounts).</summary>
        public string? Account { get; set; }

        /// <summary>Which default(s) to record: "new" | "reply" | "both".</summary>
        public string? Scope { get; set; }
    }

    [McpServerTool(Name = "open_in_outlook")]
    [Description("Show the user a mail: opens it in a visible Outlook message window (starts Outlook if needed). "
        + "Use when the user asks to see a mail, or to hand a found mail over for human reading/action. "
        + "Pass a hit id from search/thread or a full EntryID. The window stays open for the user - do not try to close it.")]
    public static async Task<CallToolResult> OpenInOutlook(
        [Description("Hit id (e.g. h12) or full EntryID hex of the mail to display.")] string id,
        CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.OpenInOutlook(id));
    }

    [McpServerTool(Name = "goto_folder")]
    [Description("Navigate the user's Outlook window to a folder (like clicking it in the folder pane). "
        + "Starts Outlook and opens a window if none is visible. Omit 'folder' for the store's Inbox.")]
    public static async Task<CallToolResult> GotoFolder(
        [Description("Store display name (see list_accounts).")] string store,
        [Description("Store-relative folder path (from list_folders), e.g. 'Inbox' or 'Projects/2026'. Omit for Inbox.")] string? folder = null,
        CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.GotoFolder(store, folder));
    }

    [McpServerTool(Name = "show_search_results")]
    [Description("Show the user a result list by driving Outlook's real search UI (the search box fills in and results appear on screen). "
        + "Use this to SHOW findings - for your own searching use the search tool instead. "
        + "query supports Outlook search syntax (free text plus e.g. from:name, hasattachments:yes). "
        + "Optional store/folder navigate the window there first; scope controls the search breadth from that folder. "
        + "When Outlook's UI search runs server-assisted (local search tuning off), the result carries an 'advice' note that "
        + "the displayed list may diverge from agent search results - relay it to the user.")]
    public static async Task<CallToolResult> ShowSearchResults(
        [Description("Search text for the Outlook search box (free text and Outlook query syntax).")] string query,
        [Description("current_folder (default - that folder only, no subfolders; the search tool's folder scope DOES "
            + "include them, so pass subfolders to show the same breadth) | subfolders (that folder and its "
            + "subfolders) | all_folders (current store) | all_outlook (all stores).")] string scope = "current_folder",
        [Description("Store display name to navigate to first (see list_accounts).")] string? store = null,
        [Description("Store-relative folder path to navigate to first. Requires store.")] string? folder = null,
        CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.ShowSearchResults(query, scope, store, folder));
    }

    private const string CcHint = "Cc recipient address(es), separated by ';' or ','. ADDED to the recipients Outlook already "
        + "put on the draft - existing recipients are never replaced. Addresses that do not resolve are reported back in "
        + "unresolvedRecipients (they stay on the draft for the user to fix), never dropped silently.";

    private const string BccHint = "Bcc recipient address(es), separated by ';' or ','. ADDED to the draft like cc; other "
        + "recipients never see them. Unresolvable addresses come back in unresolvedRecipients.";

    private const string ImportanceHint = "Message importance: 'low', 'normal' or 'high'. Omit for Outlook's default (normal).";

    private const string ReadReceiptHint = "Ask for a read receipt (default: the account's own setting). Use only when the "
        + "user asked for one - recipients see the request.";

    private const string BodyHtmlHint = "Formatted HTML body, used INSTEAD of body - supply exactly one of body or body_html. "
        + "The HTML is inserted as REAL HTML into the draft region only: above the signature and above any quoted original, "
        + "both of which stay untouched. Send a FRAGMENT (no <html>/<head>/<body> wrapper needed - they are stripped if present) "
        + "and do NOT escape it or wrap it in <pre>. Supported: h1-h6, p, br, hr, strong/b, em/i, u, s, sub/sup, ol/ul/li, dl/dt/dd, "
        + "blockquote, pre/code, tables (table/caption/thead/tbody/tfoot/tr/th/td), a[href] (http/https/mailto/tel only), and "
        + "span/div - all with inline style attributes, which ARE kept because formatting is the point. Dropped WITH their content: "
        + "script, style blocks, iframe, object/embed, link/meta, audio/video. Dropped but their text is kept: img and any other "
        + "unsupported tag. Also removed: event handlers (on*), id/name/class attributes, and CSS that loads or executes anything. "
        + "Malformed markup is REPAIRED, not rejected - unclosed and mis-nested tags are closed, stray '<' is escaped, a stray "
        + "<li>/<tr>/<td> gets the list or table it needs. Everything that was changed comes back in htmlAdjustments, so read that "
        + "field. VERIFY THE RESULT with read include_html=true - read's plain text hides layout problems.";

    private const string AttachmentsHint = "Files to attach, as ABSOLUTE paths on this machine (e.g. "
        + "\"C:\\\\Users\\\\me\\\\Documents\\\\offer.pdf\"). Any readable file is allowed - there is no folder restriction. "
        + "Every path is checked BEFORE anything is written: if even one is missing, unreadable, empty or a directory, "
        + "NOTHING is attached and no draft is created/changed, and the error names each bad path with its own reason, so "
        + "fix them and call again. Max 20 files and 150 MB per call. The result reports the files that ended up on the "
        + "saved draft with their names and sizes. Attaching a file to a draft that already has a pending send "
        + "confirm_token invalidates that token.";

    private const string DerivedSubjectHint = "Replacement subject line. Omit to keep Outlook's own RE:/FW: subject, which is "
        + "the safe default. The draft keeps threading either way (its ConversationIndex still extends the original and the "
        + "original conversation topic is carried over, reported as conversationTopicPreserved), but a changed subject is what "
        + "the recipient sees, so only override when the user asked to rename the thread. Max 255 characters.";

    [McpServerTool(Name = "new_draft")]
    [Description("Create a NEW email draft for the user - saved into the chosen account's Drafts folder with that account's "
        + "identity and signature, and opened on screen (default) so the user can review, edit and send it themselves. "
        + "NOTHING IS SENT by this tool. Supply EITHER body (plain text, line breaks preserved) OR body_html (real HTML, for a "
        + "formatted letter); either way the text is placed above the signature.")]
    public static async Task<CallToolResult> NewDraft(
        [Description("Sending account SMTP address (see list_accounts) - determines the From identity, the Drafts folder and the signature.")]
        string account,
        [Description("To recipient address(es), separated by ';' or ','.")] string to,
        [Description("Subject line.")] string subject,
        [Description("Plain-text body. Placed ABOVE the account's signature. Use body_html instead when the message needs formatting "
            + "(headings, bold, lists, tables); exactly one of body or body_html is required.")]
        string? body = null,
        [Description(BodyHtmlHint)] string? body_html = null,
        [Description(CcHint)] string? cc = null,
        [Description(BccHint)] string? bcc = null,
        [Description("Open the draft in an Outlook window for the user (default true). Pass false only when the user asked not to see it.")]
        bool display = true,
        [Description("Signature name to apply instead of the account default (see list_signatures). Pick the BEST one for this "
            + "message - e.g. match the recipient's/thread's language using the excerpts; with a single signature the choice is "
            + "trivial. Omit for the account's default signature.")]
        string? signature = null,
        [Description(ImportanceHint)] string? importance = null,
        [Description(ReadReceiptHint)] bool? request_read_receipt = null,
        [Description(AttachmentsHint)] string[]? attachments = null,
        CancellationToken cancellationToken = default)
    {
        CallToolResult? writingRules = WritingRulesRejection(body, body_html);
        if (writingRules != null)
        {
            return writingRules;
        }

        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.NewDraft(
            account, to, cc, subject, body, display, signature, bcc, importance, request_read_receipt, body_html, attachments));
    }

    [McpServerTool(Name = "reply_draft")]
    [Description("Create a REPLY draft to a mail (hit id from search/thread, or EntryID) via Outlook's own Reply - "
        + "threading and the quoted original are preserved and the right account's signature is applied; your text goes above the quote. "
        + "The draft is saved to Drafts and opened on screen (default) for the user to review, edit and send. NOTHING IS SENT.")]
    public static async Task<CallToolResult> ReplyDraft(
        [Description("Hit id (e.g. h12) or full EntryID hex of the mail to reply to.")] string id,
        [Description("Plain-text reply body. Placed ABOVE the quoted original. Use body_html instead when the reply needs formatting; "
            + "exactly one of body or body_html is required.")]
        string? body = null,
        [Description(BodyHtmlHint)] string? body_html = null,
        [Description(CcHint)] string? cc = null,
        [Description(BccHint)] string? bcc = null,
        [Description(DerivedSubjectHint)] string? subject = null,
        [Description("Open the draft in an Outlook window for the user (default true).")] bool display = true,
        [Description("Signature name to apply instead of the account default (see list_signatures). Pick the BEST one for this "
            + "reply - e.g. match the thread's language; with a single signature the choice is trivial. Omit for the default.")]
        string? signature = null,
        [Description(ImportanceHint)] string? importance = null,
        [Description(ReadReceiptHint)] bool? request_read_receipt = null,
        [Description(AttachmentsHint)] string[]? attachments = null,
        CancellationToken cancellationToken = default)
    {
        CallToolResult? writingRules = WritingRulesRejection(body, body_html);
        if (writingRules != null)
        {
            return writingRules;
        }

        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.ReplyDraft(
            id, body, replyAll: false, display, signature, cc, bcc, subject, importance, request_read_receipt, body_html, attachments));
    }

    [McpServerTool(Name = "replyall_draft")]
    [Description("Create a REPLY-ALL draft to a mail (hit id or EntryID) via Outlook's own ReplyAll - all original recipients kept, "
        + "threading and quoted history preserved, correct signature applied, your text above the quote. "
        + "Saved to Drafts and opened on screen (default) for the user to review, edit and send. NOTHING IS SENT.")]
    public static async Task<CallToolResult> ReplyAllDraft(
        [Description("Hit id (e.g. h12) or full EntryID hex of the mail to reply to.")] string id,
        [Description("Plain-text reply body. Placed ABOVE the quoted original. Use body_html instead when the reply needs formatting; "
            + "exactly one of body or body_html is required.")]
        string? body = null,
        [Description(BodyHtmlHint)] string? body_html = null,
        [Description(CcHint)] string? cc = null,
        [Description(BccHint)] string? bcc = null,
        [Description(DerivedSubjectHint)] string? subject = null,
        [Description("Open the draft in an Outlook window for the user (default true).")] bool display = true,
        [Description("Signature name to apply instead of the account default (see list_signatures). Pick the BEST one for this "
            + "reply - e.g. match the thread's language; with a single signature the choice is trivial. Omit for the default.")]
        string? signature = null,
        [Description(ImportanceHint)] string? importance = null,
        [Description(ReadReceiptHint)] bool? request_read_receipt = null,
        [Description(AttachmentsHint)] string[]? attachments = null,
        CancellationToken cancellationToken = default)
    {
        CallToolResult? writingRules = WritingRulesRejection(body, body_html);
        if (writingRules != null)
        {
            return writingRules;
        }

        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.ReplyDraft(
            id, body, replyAll: true, display, signature, cc, bcc, subject, importance, request_read_receipt, body_html, attachments));
    }

    [McpServerTool(Name = "forward_draft")]
    [Description("Create a FORWARD draft of a mail (hit id or EntryID) via Outlook's own Forward - quoted content and attachments "
        + "carried over, correct signature applied, your text above the quote. "
        + "Saved to Drafts and opened on screen (default) for the user to review, edit and send. NOTHING IS SENT.")]
    public static async Task<CallToolResult> ForwardDraft(
        [Description("Hit id (e.g. h12) or full EntryID hex of the mail to forward.")] string id,
        [Description("To recipient address(es), separated by ';' or ','.")] string to,
        [Description("Plain-text body. Placed ABOVE the forwarded mail. Use body_html instead when the message needs formatting; "
            + "exactly one of body or body_html is required.")]
        string? body = null,
        [Description(BodyHtmlHint)] string? body_html = null,
        [Description(CcHint)] string? cc = null,
        [Description(BccHint)] string? bcc = null,
        [Description(DerivedSubjectHint)] string? subject = null,
        [Description("Open the draft in an Outlook window for the user (default true).")] bool display = true,
        [Description("Signature name to apply instead of the account default (see list_signatures). Pick the BEST one for this "
            + "message - e.g. match the new recipient's language; with a single signature the choice is trivial. Omit for the default.")]
        string? signature = null,
        [Description(ImportanceHint)] string? importance = null,
        [Description(ReadReceiptHint)] bool? request_read_receipt = null,
        [Description(AttachmentsHint)] string[]? attachments = null,
        CancellationToken cancellationToken = default)
    {
        CallToolResult? writingRules = WritingRulesRejection(body, body_html);
        if (writingRules != null)
        {
            return writingRules;
        }

        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.ForwardDraft(
            id, body, to, display, signature, cc, bcc, subject, importance, request_read_receipt, body_html, attachments));
    }

    [McpServerTool(Name = "update_draft")]
    [Description("REVISE an existing draft in place - the draft keeps its id, its place in Drafts and its identity, and "
        + "only the parts you pass are touched. Use it to iterate on a draft you already created (fix a sentence, add a "
        + "recipient, attach the file you forgot) instead of creating a second draft. NOTHING IS SENT.\n\n"
        + "THE BODY IS REPLACED, NOT APPENDED: body/body_html rewrite YOUR text only - the account's signature and, on a "
        + "reply or forward, the quoted original are left exactly as they are, below it. Omit both to keep the current text.\n\n"
        + "RECIPIENTS ARE REPLACED (the opposite of the draft tools' cc/bcc, which append): to/cc/bcc each REPLACE that "
        + "whole list, which is the only way to REMOVE someone. To add a recipient, pass the full new list. A list you do "
        + "not pass is left untouched.\n\n"
        + "ATTACHMENTS ARE ADDED: attachments adds files; remove_attachments removes them by file name. Both in one call = "
        + "replace (removals run first).\n\n"
        + "SIGNATURE IMAGES survive a revision: they are stored embedded in the draft, so re-rendering keeps them. The one "
        + "exception is a draft composed by an older version of this server, whose signature image is still LINKED to a "
        + "file on disk - such a link cannot survive the re-render. That is never silent: the result reports "
        + "inlineImagesDropped with advice, and passing signature restores the signature and its images in embedded form.\n\n"
        + "Only saved, UNSENT drafts in a Drafts folder can be updated - a sent mail, a received mail or an item elsewhere "
        + "is refused with a clear reason and nothing is changed. Any pending send confirm_token for the draft is "
        + "invalidated by the update.")]
    public static async Task<CallToolResult> UpdateDraft(
        [Description("The draft to revise: the entryId a draft tool returned (preferred), or a hit id of a saved, UNSENT draft.")]
        string id,
        [Description("New plain-text body. REPLACES your text in the draft region only - the signature and any quoted "
            + "original survive. Omit to leave the body alone; exactly one of body or body_html may be supplied.")]
        string? body = null,
        [Description(BodyHtmlHint)] string? body_html = null,
        [Description("New subject line. Omit to keep the current one. On a reply/forward draft the threading is preserved "
            + "(ConversationIndex and the original topic are restored after the rename, reported as conversationTopicPreserved).")]
        string? subject = null,
        [Description("REPLACEMENT To list, separated by ';' or ','. Replaces every current To recipient - pass the full "
            + "list. Omit to keep the current To recipients.")]
        string? to = null,
        [Description("REPLACEMENT Cc list, separated by ';' or ','. Replaces every current Cc recipient; pass an empty "
            + "string to clear Cc. Omit to keep the current ones. NOTE: this is REPLACE, unlike the draft tools' append.")]
        string? cc = null,
        [Description("REPLACEMENT Bcc list, separated by ';' or ','. Replaces every current Bcc recipient; pass an empty "
            + "string to clear Bcc. Omit to keep the current ones.")]
        string? bcc = null,
        [Description(ImportanceHint)] string? importance = null,
        [Description(ReadReceiptHint)] bool? request_read_receipt = null,
        [Description("Signature name to swap in (see list_signatures). Replaces the draft's signature region only, "
            + "leaving your text and any quoted original untouched. Omit to keep the current signature.")]
        string? signature = null,
        [Description(AttachmentsHint)] string[]? attachments = null,
        [Description("File names to REMOVE from the draft, exactly as reported in the attachments list (e.g. "
            + "\"offer.pdf\"). Removals happen before additions, so listing a name here and attaching a new file with the "
            + "same name replaces it. Names that match nothing come back in attachmentsNotFound instead of failing.")]
        string[]? remove_attachments = null,
        [Description("Re-open the revised draft in an Outlook window for the user (default true, like the draft tools). "
            + "Pass false when the user is not watching or the draft is already open and you do not want it refocused.")]
        bool display = true,
        CancellationToken cancellationToken = default)
    {
        // Gated only when this revision actually rewrites the body. Changing recipients, a
        // subject, a signature or attachments is not writing, and refusing it to hand over
        // rules about wording would be a rejection with nothing to act on.
        CallToolResult? writingRules = WritingRulesRejection(body, body_html);
        if (writingRules != null)
        {
            return writingRules;
        }

        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.UpdateDraft(
            id, body, body_html, subject, to, cc, bcc, importance, request_read_receipt, signature,
            attachments, remove_attachments, display));
    }

    [McpServerTool(Name = "discard_draft")]
    [Description("Throw away a draft YOU just created in this session - the cleanup counterpart of the draft tools, for "
        + "when a draft turned out wrong or is no longer wanted. DESTRUCTIVE but deliberately tiny in reach.\n\n"
        + "WHAT IT CAN TOUCH: only a draft returned by new_draft / reply_draft / replyall_draft / forward_draft / "
        + "update_draft in THIS server session, that is still UNSENT and still in a Drafts folder. All three conditions "
        + "must hold.\n\n"
        + "WHAT IT CAN NEVER TOUCH: mail the user received or wrote themselves, anything already sent, anything outside "
        + "Drafts, a draft from an earlier session (restarting the server clears the list), and the contents of Deleted "
        + "Items. It cannot empty anything and it cannot delete permanently.\n\n"
        + "It is a SOFT delete - exactly like pressing Delete in Outlook: the draft moves to Deleted Items and the result "
        + "carries newEntryId plus fromFolder, so it can be put back with move_mail. Anything it refuses comes back as a "
        + "clear error saying why - it never silently does nothing. Every discard is audit-logged.")]
    public static async Task<CallToolResult> DiscardDraft(
        [Description("The draft to discard: the entryId a draft tool returned for a draft created in THIS session.")]
        string id,
        CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.DiscardDraft(id));
    }

    [McpServerTool(Name = "send")]
    [Description("ACTUALLY SENDS an email - the ONLY tool that sends anything. DO NOT USE THIS BY DEFAULT: the standard OutlookAI "
        + "workflow is new_draft/reply_draft/replyall_draft/forward_draft, which save a draft and open it for the USER to review "
        + "and press Send themselves. Use send ONLY when the user EXPLICITLY asked for automatic sending, or you are certain "
        + "beyond doubt that is what they want; when in doubt, create a draft instead. Deliberately high-friction, two-step: "
        + "a call WITHOUT confirm_token NEVER sends - it returns a warning plus a single-use confirm_token bound to that exact "
        + "draft and its current content. Re-confirm with the user, then call send again WITH the token. Tokens expire after "
        + "~2 minutes, work exactly once, and are invalidated by any change to the draft. The From identity is always the "
        + "account owning the draft's store, hard-verified immediately before transport (mismatch aborts - this tool can never "
        + "send from a different account). Every step is audit-logged.")]
    public static async Task<CallToolResult> Send(
        [Description("The draft to send: the entryId returned by a draft tool (preferred) or a hit id of a saved, UNSENT draft.")]
        string id,
        [Description("One-time confirmation token from the previous send call for this draft. OMIT on the first call.")]
        string? confirm_token = null,
        [Description("Optional Exchange send-on-behalf-of SMTP address (requires server-side permission). Must be identical in both calls.")]
        string? sent_on_behalf_of = null,
        CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.Send(id, confirm_token, sent_on_behalf_of));
    }

    /// <summary>
    /// The writing-rules gate in front of the five drafting tools: null to proceed, or the
    /// rejection that hands the agent the user's own writing rules and asks for one retry.
    /// See <see cref="WritingRulesGate"/> for why the rules travel in an error at all.
    /// <para>
    /// Called BEFORE <see cref="GuardAsync"/>, deliberately, and not from the service layer.
    /// The decision is a registry read and a hash, so the rejection is instant, it is
    /// identical whether Outlook is running, closed or wedged, and no COM host is started to
    /// produce it - which is also what lets the CI tier exercise it on a runner with no
    /// Outlook at all.
    /// </para>
    /// <para>
    /// "Supplies a body" is any non-null body/body_html, blank included: a call that names the
    /// parameter is a call that is writing, and whether the value is usable is the service
    /// layer's judgement to make on the retry.
    /// </para>
    /// </summary>
    private static CallToolResult? WritingRulesRejection(string? body, string? bodyHtml)
    {
        if (!WritingRulesGate.Shared.TryClaimDelivery(body != null || bodyHtml != null, out string rules))
        {
            return null;
        }

        return Error(
            WritingRulesGate.ErrorType,
            WritingRulesGate.RetryMessage,
            WritingRulesGate.Clarification,
            writingRules: rules);
    }

    /// <summary>
    /// Runs a tool operation with a bounded COM budget and turns every failure into a
    /// structured result rather than a hang or an opaque protocol fault.
    /// <para>
    /// The service layer is synchronous, so the work runs on a pool thread while this
    /// method stays async. Blocking that thread is safe here in a way it was not before:
    /// everything it can block on now goes through the COM host, which the supervisor
    /// will kill when the deadline expires.
    /// </para>
    /// </summary>
    private static async Task<CallToolResult> GuardAsync<T>(CancellationToken cancellationToken, Func<T> operation)
    {
        try
        {
            // The ambient context is established here, on the caller's execution context,
            // so it flows into the pool thread with ExecutionContext capture.
            using (ComHostRequestContext.Enter(cancellationToken))
            {
                T value = await Task.Run(operation, CancellationToken.None).ConfigureAwait(false);
                return Success(JsonSerializer.Serialize(value, Json));
            }
        }
        catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
        {
            // The CLIENT cancelled. Rethrowing is what the SDK expects - it suppresses the
            // response entirely, per spec. Anything else would answer a request that is no
            // longer there.
            throw;
        }
        catch (OperationCanceledException ex)
        {
            // A cancellation that is NOT the client's. Left to propagate it would reach the
            // SDK's generic handler and reach the agent as the message-redacted
            // "An error occurred invoking '<tool>'." - silence-adjacent, and exactly the
            // symptom this whole design exists to remove. So it is answered here instead.
            return Error("Cancelled", ex.Message,
                "The operation was cancelled before it completed. Retry; if it repeats, check outlook_health.");
        }
        catch (ComHostStartingException ex)
        {
            return Error("OutlookStarting", ex.Message,
                "Nothing is wrong. Outlook was not running (or is still coming up) and is being started in the "
                + "background; this returned at once instead of making you wait. Retry after the stated delay. search "
                + "works meanwhile and returns indexed mail.",
                retryAfterSeconds: ex.RetryAfterSeconds);
        }
        catch (ComHostUnresponsiveException ex)
        {
            return Error("OutlookUnresponsive", ex.Message,
                "This answered immediately instead of waiting on a call that would not return. TELL THE USER Outlook is "
                + "not responding and that restarting Outlook fixes it now. search still returns indexed mail meanwhile. "
                + "Outlook is re-checked automatically, so this also clears by itself. outlook_health shows the state.",
                retryAfterSeconds: ex.RetryAfterSeconds);
        }
        catch (ComHostTimeoutException ex)
        {
            return Error("Timeout", ex.Message,
                "Outlook did not answer within the time budget and the COM host was restarted, so the next call starts "
                + "clean. Outlook itself may be busy, showing a dialog, or not responding - check outlook_health. "
                + "search still returns indexed results meanwhile.");
        }
        catch (ComHostUnavailableException ex)
        {
            return Error("ComHostUnavailable", ex.Message,
                "The Outlook COM host could not be started or stopped unexpectedly. outlook_health reports its state; "
                + "search still returns indexed results without it.");
        }
        catch (SendRefusedException ex)
        {
            return Error("SendRefused", ex.Message,
                "Nothing was sent. If automatic sending is still explicitly wanted and the draft is unchanged, request a fresh "
                + "token by calling send without confirm_token and re-confirm with the user.",
                ex.Reason);
        }
        catch (DraftRefusedException ex)
        {
            return Error("DraftRefused", ex.Message,
                "Nothing was changed or deleted. Check the draft with read, or create a new draft instead.",
                ex.Reason);
        }
        catch (OutlookUnavailableException ex)
        {
            return Error("OutlookUnavailable", ex.Message,
                "Retry after the add-in update finishes. search still works meanwhile (index results with a freshness "
                + "warning), as does outlook_health.");
        }
        catch (ArgumentException ex)
        {
            return Error("InvalidArgument", ex.Message, null);
        }
        catch (COMException ex)
        {
            return Error(
                "ComFailure",
                string.Format(CultureInfo.InvariantCulture, "{0} 0x{1:X8}", ex.GetType().Name, ex.HResult),
                "Outlook rejected the operation; check outlook_health and retry.");
        }
        catch (ComHostRemoteException ex)
        {
            // A failure raised inside the COM host whose type the parent does not model.
            // The type reported is the CHILD's, not this wrapper's: "ComHostRemoteException"
            // names the pipe the failure crossed and tells an agent nothing it can act on,
            // while the child-side name is the same word it would have seen had the work
            // run in this process. The message is already the child's own.
            return Error(ex.RemoteType, ex.Message, null);
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            return Error(ex.GetType().Name, ex.Message, null);
        }
    }

    /// <summary>A successful tool result. The text payload is byte-identical to what this server has always returned.</summary>
    private static CallToolResult Success(string json)
    {
        return new CallToolResult
        {
            Content = [new TextContentBlock { Text = json }],
        };
    }

    /// <summary>
    /// A failed tool result.
    /// <para>
    /// Now carries <c>isError: true</c>. Previously every domain failure was transported
    /// as a protocol-level SUCCESS whose text happened to contain an error object - a
    /// private convention that the tool descriptions taught the model but that was
    /// invisible to any generic MCP client. The text keeps its exact former shape, so
    /// nothing that already read it breaks; the flag and the structured copy are additive.
    /// </para>
    /// </summary>
    private static CallToolResult Error(
        string type,
        string message,
        string? advice,
        string? reason = null,
        int? retryAfterSeconds = null,
        string? writingRules = null)
    {
        // 'reason' is the machine-readable refusal code (send/draft refusals), and
        // 'retryAfterSeconds' is machine-readable retry guidance for the transient states
        // (Outlook starting, Outlook unresponsive). 'writingRules' carries the user's own
        // writing prompt on the one error that exists to deliver it, and is last because it
        // is the only field that runs to pages. All three are omitted for everything else by
        // the null-ignoring serializer, so the existing error shape is unchanged.
        var payload = new { error = new { type, reason, message, advice, retryAfterSeconds, writingRules } };
        return new CallToolResult
        {
            IsError = true,
            Content = [new TextContentBlock { Text = JsonSerializer.Serialize(payload, Json) }],
            StructuredContent = JsonSerializer.SerializeToElement(payload, Json),
        };
    }

    private static DateTime? ParseUtc(string? value, string name)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        if (DateTimeOffset.TryParse(
                value,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal,
                out DateTimeOffset parsed))
        {
            return parsed.UtcDateTime;
        }

        throw new ArgumentException("'" + name + "' is not a parsable ISO 8601 timestamp: " + value);
    }
}
