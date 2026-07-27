using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;
using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;

namespace OutlookAI.McpServer.Tools;

/// <summary>
/// Process-wide service holder: ONE MailService (and with it one ComGateway / pumped
/// STA thread and one held-open Outlook session) per server process. Created lazily so
/// starting the server never touches the index or Outlook by itself.
/// </summary>
internal static class ServerRuntime
{
    private static readonly Lazy<MailService> LazyService = new(
        MailService.CreateDefault, LazyThreadSafetyMode.ExecutionAndPublication);

    internal static MailService Service => LazyService.Value;
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

    [McpServerTool(Name = "search")]
    [Description("Search locally indexed Outlook mail across all accounts, folders, and delegate mailboxes. "
        + "Sub-second and cheap: iterate with refined terms instead of pulling large result sets.\n\n"
        + "MATCHING: query terms are whitespace-separated and ANDed; each term matches whole words, and the terms "
        + "may land in different parts of the mail - one in the subject, another in the body (search_in narrows "
        + "matching to just the subject or just the body). Body includes "
        + "attachment text of EVERY attachment type - documents, images, embedded messages, calendar invites, "
        + "media - and those matches return as separate hits with isAttachmentHit=true "
        + "(include_attachment_hits=false drops them; read on such a hit opens the parent mail). Sender and "
        + "recipients are NOT matched by query terms - use from / to. Append * for prefix match (haproxy*). "
        + "Allowed characters: letters, digits and @.-_'+ ; omit query entirely to filter only by "
        + "from/to/date/flags.\n\n"
        + "FRESHNESS: results always include mail that arrived after the last index update - the server sweeps it "
        + "live through Outlook (autostarting it headless when needed) and merges it in. The sweep follows your "
        + "scope: with folder set it covers that folder and its subfolders, otherwise Inbox, Sent Items, Deleted "
        + "Items and Junk Email of the store(s) in scope (those four folders only, not their subfolders) - so for "
        + "brand-new mail filed into any other folder, pass store + folder. "
        + "The response's sweep block reports what was covered. The sweep is cached ~10 s, "
        + "so rapid follow-up searches run at index speed. If it cannot run, "
        + "index results are returned with a warning in advice - a search never fails for that reason.\n\n"
        + "FOLDER SCOPE: folder covers that folder AND its subfolders by default in every mode; pass "
        + "include_subfolders=false for that one folder alone. Delegate/shared mailboxes are indexed WITHOUT "
        + "their folder nesting, so a folder scope there matches by folder NAME: if the subfolder set cannot be "
        + "built the search widens to the whole delegate mailbox and says so in advice, and if that mailbox has "
        + "two folders with the same name the results can include both - advice says that too. The scope block "
        + "in the response reports what was actually covered.\n\n"
        + "RESULTS: each hit carries an id for read, thread, save_attachment, open_in_outlook, move_mail and "
        + "archive_mail (ids are valid for this session). truncated=true means more matches exist beyond top "
        + "(max 100) - narrow with store/folder/from/after rather than raising it. Read advice when present and "
        + "relay it to the user when it concerns them: it is where every partial result, cap, skipped folder and "
        + "widened scope is reported.\n\n"
        + "EXHAUSTIVE: exhaustive=true bypasses the index and scans folders through Outlook instead - requires "
        + "store plus folder and/or after, is far slower, and matches whole words in subject and body only (no "
        + "attachment text). It follows include_subfolders like the other modes, so a folder scope walks the "
        + "subtree - which on a big subtree can hit the 120 s budget; pass include_subfolders=false to scan just "
        + "the named folder, and check foldersScanned/foldersSkipped plus advice for partial coverage. Use it "
        + "when the index looks stale or wrong, or when completeness matters more than speed.")]
    public static string Search(
        [Description("Free-text terms, whitespace-separated, ANDed. Each term may match in the subject or the body (see search_in). Letters/digits plus @.-_'+ only; trailing * for prefix. Omit to filter by sender/date only.")]
        string? query = null,
        [Description("Which part of the mail the query terms must match: subject_and_body (default), subject, or "
            + "body. Narrow it when a term is noisy in one of them.")]
        string? search_in = null,
        [Description("true = bounded index-bypassing COM scan (requires store + folder and/or after). Default false: index + freshness sweep.")]
        bool exhaustive = false,
        [Description("Store display name to search in (see list_accounts). Omit for all stores (required when exhaustive=true).")] string? store = null,
        [Description("Store-relative folder path (from list_folders), e.g. 'Inbox' or 'Projects/2026'. Requires store. "
            + "Includes its subfolders unless include_subfolders=false.")] string? folder = null,
        [Description("Whether folder covers its subfolders. Default true, in every mode. Set false to search that "
            + "one folder only - also the cheap way to keep an exhaustive scan bounded.")]
        bool include_subfolders = true,
        [Description("Sender filter: address or name fragment (index-backed).")] string? from = null,
        [Description("Recipient (To/Cc) filter: address fragment (index-backed).")] string? to = null,
        [Description("Only mail received at/after this instant (ISO 8601, e.g. 2026-07-01 or 2026-07-01T08:00:00Z).")] string? after = null,
        [Description("Only mail received before this instant (ISO 8601).")] string? before = null,
        [Description("true = unread mail only.")] bool? unread_only = null,
        [Description("Filter on attachment presence.")] bool? has_attachments = null,
        [Description("Include attachment-CONTENT matches (any attachment type: documents, images, embedded "
            + "messages, invites, media). Default true.")]
        bool include_attachment_hits = true,
        [Description("Max hits (1-100, default 25). Keep small - iterate instead.")] int top = 25,
        [Description("Snippet length per hit (0-1000, default 200; 0 = no snippets).")] int snippet_chars = 200)
    {
        return Guard(() =>
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

    [McpServerTool(Name = "thread")]
    [Description("Fetch the full conversation of a mail. Two complementary lookup keys - pass BOTH when available: "
        + "conversation_id is the fast path (a free index lookup; every search hit already carries it, no locate cost), "
        + "and id anchors the COM fallback that walks Outlook's conversation graph when the index has no rows for the "
        + "conversation - COM cannot look up a conversation by id string, it needs a concrete mail item to start from. "
        + "Members are oldest-first; truncated=true means the conversation has more members than 'top'.")]
    public static string Thread(
        [Description("ConversationId from a search hit or read result - the fast index path. Pass when you have it.")] string? conversation_id = null,
        [Description("Hit id (e.g. h12) or EntryID of any mail in the conversation - anchors the COM conversation-graph fallback (used when the index has no rows).")] string? id = null,
        [Description("Store display name to scope the index lookup (faster).")] string? store = null,
        [Description("Max thread members (default 50).")] int top = 50)
    {
        return Guard(() => ServerRuntime.Service.Thread(conversation_id, id, store, top));
    }

    [McpServerTool(Name = "read")]
    [Description("Read one mail in full by hit id (from search/thread) or EntryID: plain-text body with truncation flags and true total size, "
        + "sender/recipients with SMTP addresses, attachment list, conversation id. For an attachment-content hit this opens the PARENT mail. "
        + "Long bodies page cheaply: when bodyTruncated=true, call again with body_offset = bodyOffset + body.length to CONTINUE reading - "
        + "the next window is served from the already-extracted body, not re-read from the start. "
        + "Needs Outlook (starts it if allowed). First read of an index hit locates the item (up to a few seconds); repeats are cached.")]
    public static string Read(
        [Description("Hit id (e.g. h12) or full EntryID hex.")] string id,
        [Description("Body window size in characters (default 20000; 0 = metadata only). bodyTruncated=true means more body exists beyond the window; bodyTotalChars is the full size.")] int max_body_chars = MailService.BodyCharsDefault,
        [Description("Include raw transport headers (capped at 8 KB). Default false.")] bool include_headers = false,
        [Description("Start of the body window in characters (default 0). Use the previous read's bodyOffset + body.length to continue a long body.")] int body_offset = 0)
    {
        return Guard(() => ServerRuntime.Service.Read(id, max_body_chars, include_headers, MailService.HeaderCharsDefault, body_offset));
    }

    [McpServerTool(Name = "save_attachment")]
    [Description("Save one attachment of a mail to disk so you can open/read the file yourself. "
        + "Use the attachment 'index' from a read result. Never overwrites - existing names get a numeric suffix. Returns the absolute path.")]
    public static string SaveAttachment(
        [Description("Hit id or EntryID of the mail (for attachment-content hits: the hit itself).")] string id,
        [Description("1-based attachment index from read's attachments list.")] int attachment_index,
        [Description("Absolute target directory. Default: %LOCALAPPDATA%\\OutlookAI\\scratch\\attachments.")] string? target_dir = null)
    {
        return Guard(() => ServerRuntime.Service.SaveAttachment(id, attachment_index, target_dir));
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
    public static string MoveMail(
        [Description("1-50 hit ids (e.g. h12) or full EntryID hex strings. Each item is moved within its own store.")]
        string[] ids,
        [Description("Store-relative target folder path (from list_folders), e.g. 'Archive/2026' or 'Projects/Acme'.")]
        string folder,
        [Description("Create the target folder (including missing parents) when it does not exist. Default false.")]
        bool create_folder = false,
        [Description("Optional store display name (see list_accounts): when given, items living in a DIFFERENT store fail "
            + "with a cross-store error instead of moving. Omit to move each item within its own store.")]
        string? store = null)
    {
        return Guard(() => ServerRuntime.Service.MoveMail(ids, folder, create_folder, store));
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
    public static string ArchiveMail(
        [Description("1-50 hit ids (e.g. h12) or full EntryID hex strings; may span accounts.")]
        string[] ids)
    {
        return Guard(() => ServerRuntime.Service.ArchiveMail(ids));
    }

    [McpServerTool(Name = "outlook_health")]
    [Description("One-call Outlook + mail-server health and freshness report: Outlook running + installed version, whether it "
        + "runs headless (no window, tray icon only - the autostart state; a normal Outlook launch promotes it), probed COM "
        + "session liveness (comConnected), store reachability, index freshness (newest indexed mail vs clock, globally AND "
        + "per store - the index only advances while Outlook runs), Windows Search (WSearch) service state, audit-log "
        + "writability, OutlookAI tuning state (incl. the effective UI search backend), and the add-in installer mutex. "
        + "Read-only - attaches to Outlook only when it is already running, NEVER starts it. status=ok means all "
        + "dependencies are available; problems lists each degradation; advice carries freshness guidance (search covers "
        + "any index gap automatically with its COM sweep - this tool only reports).")]
    public static string OutlookHealth()
    {
        return Guard(() => ServerRuntime.Service.Health());
    }

    [McpServerTool(Name = "list_accounts")]
    [Description("List the profile's mail accounts and ALL stores (accounts, delegate/shared caches, archives) with flags: "
        + "isDelegate, onlineOnly/locallySearchable (server-only stores like Online Archives are invisible to local search), "
        + "and inLocalIndex. Store display names are the 'store' argument for search/list_folders.")]
    public static string ListAccounts()
    {
        return Guard(() => ServerRuntime.Service.ListAccounts());
    }

    [McpServerTool(Name = "list_folders")]
    [Description("List the FULL folder tree(s) with item/unread counts. Folder paths feed the search tool's 'folder' argument. "
        + "Traversal order is stable: stores sorted by display name, then depth-first with sibling folders sorted by name. "
        + "One call returns up to 1000 folders (virtually always the whole tree); truncated=true means more exist - "
        + "continue with offset=nextOffset to page the remainder in the same stable order.")]
    public static string ListFolders(
        [Description("Store display name (see list_accounts). Omit for all stores.")] string? store = null,
        [Description("Folders to skip in the stable traversal (default 0). Use the previous result's nextOffset to continue a truncated listing.")] int offset = 0)
    {
        return Guard(() => ServerRuntime.Service.ListFolders(store, offset));
    }

    [McpServerTool(Name = "list_signatures")]
    [Description("List the installed Outlook email signatures: each name plus a short plain-text excerpt (use the excerpt to "
        + "detect a signature's language and purpose), and per-account default assignments where the profile registry records "
        + "them (missing = unknown, e.g. no signature configured or roaming-managed - never guessed). Use this to pick the "
        + "BEST signature for a draft via the draft tools' 'signature' parameter - e.g. match the language of the "
        + "recipient/thread. Read-only, never starts Outlook.")]
    public static string ListSignatures()
    {
        return Guard(() => ServerRuntime.Service.ListSignatures());
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
    public static string ManageSignature(
        [Description("'create' | 'update' | 'delete'.")] string action,
        [Description("Signature name (as shown by list_signatures and Outlook's signature pickers).")] string name,
        [Description("Plain-text signature body (create/update). When omitted it is derived from body_html.")]
        string? body_text = null,
        [Description("HTML signature body (create/update), fragment or full document. When omitted it is derived from body_text.")]
        string? body_html = null,
        [Description("Optionally record the signature as an account default: {\"account\": SMTP address from list_accounts, "
            + "\"scope\": \"new\"|\"reply\"|\"both\"}. Not allowed with delete.")]
        SetDefaultForArg? set_default_for = null)
    {
        return Guard(() => ServerRuntime.Service.ManageSignature(new ManageSignatureRequest
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
    public static string OpenInOutlook(
        [Description("Hit id (e.g. h12) or full EntryID hex of the mail to display.")] string id)
    {
        return Guard(() => ServerRuntime.Service.OpenInOutlook(id));
    }

    [McpServerTool(Name = "goto_folder")]
    [Description("Navigate the user's Outlook window to a folder (like clicking it in the folder pane). "
        + "Starts Outlook and opens a window if none is visible. Omit 'folder' for the store's Inbox.")]
    public static string GotoFolder(
        [Description("Store display name (see list_accounts).")] string store,
        [Description("Store-relative folder path (from list_folders), e.g. 'Inbox' or 'Projects/2026'. Omit for Inbox.")] string? folder = null)
    {
        return Guard(() => ServerRuntime.Service.GotoFolder(store, folder));
    }

    [McpServerTool(Name = "show_search_results")]
    [Description("Show the user a result list by driving Outlook's real search UI (the search box fills in and results appear on screen). "
        + "Use this to SHOW findings - for your own searching use the search tool instead. "
        + "query supports Outlook search syntax (free text plus e.g. from:name, hasattachments:yes). "
        + "Optional store/folder navigate the window there first; scope controls the search breadth from that folder. "
        + "When Outlook's UI search runs server-assisted (local search tuning off), the result carries an 'advice' note that "
        + "the displayed list may diverge from agent search results - relay it to the user.")]
    public static string ShowSearchResults(
        [Description("Search text for the Outlook search box (free text and Outlook query syntax).")] string query,
        [Description("current_folder (default - that folder only, no subfolders; the search tool's folder scope DOES "
            + "include them, so pass subfolders to show the same breadth) | subfolders (that folder and its "
            + "subfolders) | all_folders (current store) | all_outlook (all stores).")] string scope = "current_folder",
        [Description("Store display name to navigate to first (see list_accounts).")] string? store = null,
        [Description("Store-relative folder path to navigate to first. Requires store.")] string? folder = null)
    {
        return Guard(() => ServerRuntime.Service.ShowSearchResults(query, scope, store, folder));
    }

    [McpServerTool(Name = "new_draft")]
    [Description("Create a NEW email draft for the user - saved into the chosen account's Drafts folder with that account's "
        + "identity and signature, and opened on screen (default) so the user can review, edit and send it themselves. "
        + "NOTHING IS SENT by this tool. Body is plain text (line breaks preserved) and is placed above the signature.")]
    public static string NewDraft(
        [Description("Sending account SMTP address (see list_accounts) - determines the From identity, the Drafts folder and the signature.")]
        string account,
        [Description("To recipient address(es), separated by ';' or ','.")] string to,
        [Description("Subject line.")] string subject,
        [Description("Plain-text body. Placed ABOVE the account's signature.")] string body,
        [Description("Cc recipient address(es), separated by ';' or ','. Optional.")] string? cc = null,
        [Description("Open the draft in an Outlook window for the user (default true). Pass false only when the user asked not to see it.")]
        bool display = true,
        [Description("Signature name to apply instead of the account default (see list_signatures). Pick the BEST one for this "
            + "message - e.g. match the recipient's/thread's language using the excerpts; with a single signature the choice is "
            + "trivial. Omit for the account's default signature.")]
        string? signature = null)
    {
        return Guard(() => ServerRuntime.Service.NewDraft(account, to, cc, subject, body, display, signature));
    }

    [McpServerTool(Name = "reply_draft")]
    [Description("Create a REPLY draft to a mail (hit id from search/thread, or EntryID) via Outlook's own Reply - "
        + "threading and the quoted original are preserved and the right account's signature is applied; your text goes above the quote. "
        + "The draft is saved to Drafts and opened on screen (default) for the user to review, edit and send. NOTHING IS SENT.")]
    public static string ReplyDraft(
        [Description("Hit id (e.g. h12) or full EntryID hex of the mail to reply to.")] string id,
        [Description("Plain-text reply body. Placed ABOVE the quoted original.")] string body,
        [Description("Open the draft in an Outlook window for the user (default true).")] bool display = true,
        [Description("Signature name to apply instead of the account default (see list_signatures). Pick the BEST one for this "
            + "reply - e.g. match the thread's language; with a single signature the choice is trivial. Omit for the default.")]
        string? signature = null)
    {
        return Guard(() => ServerRuntime.Service.ReplyDraft(id, body, replyAll: false, display, signature));
    }

    [McpServerTool(Name = "replyall_draft")]
    [Description("Create a REPLY-ALL draft to a mail (hit id or EntryID) via Outlook's own ReplyAll - all original recipients kept, "
        + "threading and quoted history preserved, correct signature applied, your text above the quote. "
        + "Saved to Drafts and opened on screen (default) for the user to review, edit and send. NOTHING IS SENT.")]
    public static string ReplyAllDraft(
        [Description("Hit id (e.g. h12) or full EntryID hex of the mail to reply to.")] string id,
        [Description("Plain-text reply body. Placed ABOVE the quoted original.")] string body,
        [Description("Open the draft in an Outlook window for the user (default true).")] bool display = true,
        [Description("Signature name to apply instead of the account default (see list_signatures). Pick the BEST one for this "
            + "reply - e.g. match the thread's language; with a single signature the choice is trivial. Omit for the default.")]
        string? signature = null)
    {
        return Guard(() => ServerRuntime.Service.ReplyDraft(id, body, replyAll: true, display, signature));
    }

    [McpServerTool(Name = "forward_draft")]
    [Description("Create a FORWARD draft of a mail (hit id or EntryID) via Outlook's own Forward - quoted content and attachments "
        + "carried over, correct signature applied, your text above the quote. "
        + "Saved to Drafts and opened on screen (default) for the user to review, edit and send. NOTHING IS SENT.")]
    public static string ForwardDraft(
        [Description("Hit id (e.g. h12) or full EntryID hex of the mail to forward.")] string id,
        [Description("Plain-text body. Placed ABOVE the forwarded mail.")] string body,
        [Description("To recipient address(es), separated by ';' or ','.")] string to,
        [Description("Open the draft in an Outlook window for the user (default true).")] bool display = true,
        [Description("Signature name to apply instead of the account default (see list_signatures). Pick the BEST one for this "
            + "message - e.g. match the new recipient's language; with a single signature the choice is trivial. Omit for the default.")]
        string? signature = null)
    {
        return Guard(() => ServerRuntime.Service.ForwardDraft(id, body, to, display, signature));
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
    public static string Send(
        [Description("The draft to send: the entryId returned by a draft tool (preferred) or a hit id of a saved, UNSENT draft.")]
        string id,
        [Description("One-time confirmation token from the previous send call for this draft. OMIT on the first call.")]
        string? confirm_token = null,
        [Description("Optional Exchange send-on-behalf-of SMTP address (requires server-side permission). Must be identical in both calls.")]
        string? sent_on_behalf_of = null)
    {
        return Guard(() => ServerRuntime.Service.Send(id, confirm_token, sent_on_behalf_of));
    }

    private static string Guard<T>(Func<T> operation)
    {
        try
        {
            return JsonSerializer.Serialize(operation(), Json);
        }
        catch (SendRefusedException ex)
        {
            return Error("SendRefused", ex.Message,
                "Nothing was sent. If automatic sending is still explicitly wanted and the draft is unchanged, request a fresh "
                + "token by calling send without confirm_token and re-confirm with the user.");
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
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            return Error(ex.GetType().Name, ex.Message, null);
        }
    }

    private static string Error(string type, string message, string? advice)
    {
        var payload = new { error = new { type, message, advice } };
        return JsonSerializer.Serialize(payload, Json);
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
