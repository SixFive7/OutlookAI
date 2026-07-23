using System.ComponentModel;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.Json.Serialization;
using ModelContextProtocol.Server;
using OutlookAI.Core.Com;
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
/// MCP tool surface v1 (v3.MD section 0.5 L1+L2): search, thread, read,
/// save_attachment, index_status, list_accounts, list_folders. Payloads are compact
/// JSON (camelCase, nulls omitted - section 12 discipline); domain failures come back
/// as an {"error": ...} object instead of protocol faults so agents can react.
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
    [Description("Search all locally indexed Outlook mail (3 accounts + delegate stores), including attachment contents. "
        + "Sub-second and cheap: iterate freely with refined terms. Terms in 'query' are ANDed; append * for prefix match. "
        + "mode=fresh (default) also COM-sweeps mail newer than the index frontier so just-arrived mail is found (may start Outlook); "
        + "mode=fast is index-only and works with Outlook closed (results can lag - check staleness); "
        + "mode=exhaustive bypasses the index with a bounded COM folder scan (requires store + folder and/or after; slower - "
        + "use when the index is stale/broken or correctness beats speed; whole-word term matching on subject+body). "
        + "Returns compact hits with an 'id' for read/save_attachment/thread/open_in_outlook; truncated=true means more "
        + "matches exist beyond 'top'.")]
    public static string Search(
        [Description("Free-text terms, whitespace-separated, ANDed. Letters/digits plus @.-_'+ only; trailing * for prefix. Omit to filter by sender/date only.")]
        string? query = null,
        [Description("fast | fresh | exhaustive (default fresh).")] string mode = "fresh",
        [Description("Store display name to search in (see list_accounts). Omit for all stores (required for mode=exhaustive).")] string? store = null,
        [Description("Store-relative folder path (from list_folders), e.g. 'Inbox' or 'Projects/2026'. Requires store.")] string? folder = null,
        [Description("Sender filter: address or name fragment (index-backed).")] string? from = null,
        [Description("Recipient (To/Cc) filter: address fragment (index-backed).")] string? to = null,
        [Description("Only mail received at/after this instant (ISO 8601, e.g. 2026-07-01 or 2026-07-01T08:00:00Z).")] string? after = null,
        [Description("Only mail received before this instant (ISO 8601).")] string? before = null,
        [Description("true = unread mail only.")] bool? unread_only = null,
        [Description("Filter on attachment presence.")] bool? has_attachments = null,
        [Description("Include attachment-CONTENT matches (kind=document rows). Default true.")] bool include_attachment_hits = true,
        [Description("Max hits (1-100, default 25). Keep small - iterate instead.")] int top = 25,
        [Description("Snippet length per hit (0-1000, default 200; 0 = no snippets).")] int snippet_chars = 200)
    {
        return Guard(() =>
        {
            SearchMode parsedMode = mode.ToLowerInvariant() switch
            {
                "fast" => SearchMode.Fast,
                "fresh" => SearchMode.Fresh,
                "exhaustive" => SearchMode.Exhaustive,
                _ => throw new ArgumentException("mode must be 'fast', 'fresh' or 'exhaustive'."),
            };

            SearchRequest request = new()
            {
                Query = query,
                Mode = parsedMode,
                Store = store,
                Folder = folder,
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
    [Description("Fetch the full conversation for a hit: pass conversation_id (from a search hit) and/or id (hit id or EntryID). "
        + "Uses the index first; falls back to Outlook's conversation graph via COM when the index has no rows. Members are oldest-first; "
        + "truncated=true means the conversation has more members than 'top'.")]
    public static string Thread(
        [Description("ConversationId from a search hit or read result.")] string? conversation_id = null,
        [Description("Hit id (e.g. h12) or EntryID whose conversation to fetch (enables the COM fallback).")] string? id = null,
        [Description("Store display name to scope the index lookup (faster).")] string? store = null,
        [Description("Max thread members (default 50).")] int top = 50)
    {
        return Guard(() => ServerRuntime.Service.Thread(conversation_id, id, store, top));
    }

    [McpServerTool(Name = "read")]
    [Description("Read one mail in full by hit id (from search/thread) or EntryID: plain-text body with truncation flags and true total size, "
        + "sender/recipients with SMTP addresses, attachment list, conversation id. For an attachment-content hit this opens the PARENT mail. "
        + "Needs Outlook (starts it if allowed). First read of an index hit locates the item (up to a few seconds); repeats are cached.")]
    public static string Read(
        [Description("Hit id (e.g. h12) or full EntryID hex.")] string id,
        [Description("Body cap in characters (default 20000; 0 = metadata only). bodyTruncated+bodyTotalChars flag cuts.")] int max_body_chars = MailService.BodyCharsDefault,
        [Description("Include raw transport headers (capped at 8 KB). Default false.")] bool include_headers = false)
    {
        return Guard(() => ServerRuntime.Service.Read(id, max_body_chars, include_headers));
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

    [McpServerTool(Name = "index_status")]
    [Description("Freshness self-report: newest indexed mail vs clock (global and per store), whether Outlook is running "
        + "(the index only advances while it runs), and advice on when to use search mode=fresh. Never starts Outlook.")]
    public static string IndexStatus()
    {
        return Guard(() => ServerRuntime.Service.IndexStatus());
    }

    [McpServerTool(Name = "health")]
    [Description("Compact health check of everything this server depends on: Outlook running + installed version, store "
        + "reachability, index freshness, Windows Search (WSearch) service state, audit-log writability, OutlookAI tuning "
        + "state, and the add-in installer mutex. Read-only - attaches to Outlook only when it is already running, never "
        + "starts it. status=ok means all dependencies are available; otherwise problems lists each degradation.")]
    public static string Health()
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
    [Description("List folder trees with item/unread counts. Folder paths feed the search tool's 'folder' argument. "
        + "Depth-capped for compact output; raise depth for deeper trees.")]
    public static string ListFolders(
        [Description("Store display name (see list_accounts). Omit for all stores.")] string? store = null,
        [Description("Tree depth 1-6 (default 2).")] int depth = 2)
    {
        return Guard(() => ServerRuntime.Service.ListFolders(store, depth));
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
        + "Optional store/folder navigate the window there first; scope controls the search breadth from that folder.")]
    public static string ShowSearchResults(
        [Description("Search text for the Outlook search box (free text and Outlook query syntax).")] string query,
        [Description("current_folder (default) | subfolders | all_folders (current store) | all_outlook (all stores).")] string scope = "current_folder",
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
        bool display = true)
    {
        return Guard(() => ServerRuntime.Service.NewDraft(account, to, cc, subject, body, display));
    }

    [McpServerTool(Name = "reply_draft")]
    [Description("Create a REPLY draft to a mail (hit id from search/thread, or EntryID) via Outlook's own Reply - "
        + "threading and the quoted original are preserved and the right account's signature is applied; your text goes above the quote. "
        + "The draft is saved to Drafts and opened on screen (default) for the user to review, edit and send. NOTHING IS SENT.")]
    public static string ReplyDraft(
        [Description("Hit id (e.g. h12) or full EntryID hex of the mail to reply to.")] string id,
        [Description("Plain-text reply body. Placed ABOVE the quoted original.")] string body,
        [Description("Open the draft in an Outlook window for the user (default true).")] bool display = true)
    {
        return Guard(() => ServerRuntime.Service.ReplyDraft(id, body, replyAll: false, display));
    }

    [McpServerTool(Name = "replyall_draft")]
    [Description("Create a REPLY-ALL draft to a mail (hit id or EntryID) via Outlook's own ReplyAll - all original recipients kept, "
        + "threading and quoted history preserved, correct signature applied, your text above the quote. "
        + "Saved to Drafts and opened on screen (default) for the user to review, edit and send. NOTHING IS SENT.")]
    public static string ReplyAllDraft(
        [Description("Hit id (e.g. h12) or full EntryID hex of the mail to reply to.")] string id,
        [Description("Plain-text reply body. Placed ABOVE the quoted original.")] string body,
        [Description("Open the draft in an Outlook window for the user (default true).")] bool display = true)
    {
        return Guard(() => ServerRuntime.Service.ReplyDraft(id, body, replyAll: true, display));
    }

    [McpServerTool(Name = "forward_draft")]
    [Description("Create a FORWARD draft of a mail (hit id or EntryID) via Outlook's own Forward - quoted content and attachments "
        + "carried over, correct signature applied, your text above the quote. "
        + "Saved to Drafts and opened on screen (default) for the user to review, edit and send. NOTHING IS SENT.")]
    public static string ForwardDraft(
        [Description("Hit id (e.g. h12) or full EntryID hex of the mail to forward.")] string id,
        [Description("Plain-text body. Placed ABOVE the forwarded mail.")] string body,
        [Description("To recipient address(es), separated by ';' or ','.")] string to,
        [Description("Open the draft in an Outlook window for the user (default true).")] bool display = true)
    {
        return Guard(() => ServerRuntime.Service.ForwardDraft(id, body, to, display));
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
                "Retry after the add-in update finishes, or use search mode=fast / index_status meanwhile.");
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
                "Outlook rejected the operation; check index_status and retry.");
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
