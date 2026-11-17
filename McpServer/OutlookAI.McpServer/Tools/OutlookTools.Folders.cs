using System.ComponentModel;
using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using OutlookAI.Core.Com;
using OutlookAI.Core.Mapi;

namespace OutlookAI.McpServer.Tools;

/// <summary>
/// Folders and stores addressed by ID (Q131-Q136, decided by the maintainer 2026-10-11): the two
/// tools that exist because of it - find_folder (Q136 (ii)) and create_folder (Q134 (i)) - the one
/// sentence every folder_id and store_id argument carries, and the refusal of the arguments this
/// release retired (Q135 (i), <see cref="RetiredArguments"/>).
/// </summary>
public static partial class OutlookTools
{
    /// <summary>
    /// What every folder_id (and parent_id) argument says about IDs - written once, so the seven
    /// arguments cannot drift apart. An agent reads it without any doc: what an ID is, where it comes
    /// from, and that a name or a path is refused.
    /// </summary>
    internal const string FolderIdRule = "Folders are addressed ONLY by ID: get one from find_folder (by the folder's "
        + "name) or list_folders, or the folderId a read, a move or a hit carries. A name or path such as 'Inbox' or "
        + "'Projects/2026' is refused - look the folder up first. An ID stays valid when the folder is renamed or moved; "
        + "copy it exactly (its last character is a check that catches a typo).";

    /// <summary>What every store_id argument says about IDs.</summary>
    internal const string StoreIdRule = "Stores are addressed ONLY by ID - a display name or an address is refused, "
        + "because two stores may share one.";

    [McpServerTool(Name = "find_folder", ReadOnly = true, Destructive = false, Idempotent = true, OpenWorld = false)]
    [Description("Find folders by NAME and get their IDs. Every tool addresses a folder by its folderId, never by a "
        + "name or path: call this whenever the user names a folder ('my Projects folder', 'Archive/2026') and pass "
        + "the chosen folder's folderId as folder_id to search, move_mail, goto_folder or show_search_results, or as "
        + "parent_id to create_folder.\n\n"
        + "Read-only and it never acts: it returns CANDIDATES - folderId, name, store and storeId, and the folder's "
        + "path (for reading) with its item counts - exact name matches first, then folders whose path ends with the "
        + "text (so 'Projects/2026' finds 2026 inside Projects), then names containing it. Names match "
        + "case-insensitively. When several folders match, pick by store and path; if that does not settle it, ASK "
        + "THE USER - never guess between folders that share a name. No match means no such folder: list_folders "
        + "shows the whole tree, and create_folder makes a new one.\n\n"
        + "Needs Outlook (starts it headless if needed); walks the same tree list_folders does.")]
    public static async Task<CallToolResult> FindFolder(
        [Description("The folder's name, or part of it, or the end of its path ('2026', 'Projects/2026'). Matched "
            + "case-insensitively; nothing is created or changed.")]
        string name,
        [Description("storeId (from list_accounts) to look in one store only. Omit to look in every store. " + StoreIdRule)]
        string? store_id = null,
        CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.FindFolder(name, store_id));
    }

    [McpServerTool(Name = "create_folder")]
    [Description("Create ONE new mail folder named 'name' directly under the folder 'parent_id', and get its folderId - "
        + "ready for move_mail. For a top-level folder, parent_id is the store's rootFolderId (from list_accounts or "
        + "list_folders).\n\n"
        + "The name is used EXACTLY as given: a '/' is an ordinary character in it, NOT a path separator - "
        + "create_folder(parent_id, 'Q3/Q4') makes one folder called 'Q3/Q4'. To make nested folders, create the "
        + "parent first and pass its folderId as parent_id for the next.\n\n"
        + "Refused, with nothing created: under Deleted Items or anything in it, under the Outbox, under a folder that "
        + "is not a mail folder, and when the parent already has a folder of that name (Outlook compares names "
        + "case-insensitively) - that refusal gives the existing folder's folderId: use it. Check with find_folder "
        + "first when the folder may already exist. This server cannot rename or delete folders. Audit-logged. Needs "
        + "Outlook (starts it headless if needed)."
        + OutcomeHint)]
    public static async Task<CallToolResult> CreateFolder(
        [Description("The folderId of the folder to create the new folder in - from find_folder or list_folders, or a "
            + "store's rootFolderId for a top-level folder. " + FolderIdRule)]
        string parent_id,
        [Description("The new folder's name, exactly as it should read - '/' included, which is part of the name, never "
            + "a separator. At most 255 characters, no line breaks or tabs.")]
        string name,
        CancellationToken cancellationToken = default)
    {
        return await GuardAsync(cancellationToken, () => ServerRuntime.Service.CreateFolder(parent_id, name));
    }

    /// <summary>
    /// The refusal of a call that passed a retired argument (<see cref="RetiredArguments"/>): an ordinary
    /// InvalidArgument error, provably unchanged - it is answered before the tool runs.
    /// </summary>
    internal static CallToolResult RetiredArgumentError(string message)
    {
        return Error("InvalidArgument", message, null, outcome: MutationOutcome.Unchanged);
    }
}

/// <summary>
/// The arguments this release retired (Q135 (i): a clean break in one release, path input refused
/// with a pointer to find_folder). An MCP server binds arguments by name and ignores names it does
/// not know - so without this, a call written for the old surface would not fail: search with
/// <c>store</c> and <c>folder</c> would run UNSCOPED and answer from every mailbox, and move_mail with
/// <c>folder</c> would fail on the missing folder_id with no word about why. Every call passes through
/// <see cref="Filter"/> first; one carrying a retired name is refused, naming what replaced it - and,
/// for a folder, the find_folder call that gets its ID.
/// </summary>
internal static class RetiredArguments
{
    /// <summary>Tool name, then the retired argument names it refuses.</summary>
    internal static readonly IReadOnlyDictionary<string, IReadOnlyList<string>> ByTool =
        new Dictionary<string, IReadOnlyList<string>>(StringComparer.Ordinal)
        {
            ["search"] = new[] { "store", "folder" },
            ["thread"] = new[] { "store" },
            ["move_mail"] = new[] { "folder", "create_folder", "store" },
            ["list_folders"] = new[] { "store" },
            ["goto_folder"] = new[] { "store", "folder" },
            ["show_search_results"] = new[] { "store", "folder" },
        };

    /// <summary>The call-tool filter (registered in <see cref="ServerHost"/>).</summary>
    internal static McpRequestHandler<CallToolRequestParams, CallToolResult> Filter(
        McpRequestHandler<CallToolRequestParams, CallToolResult> next)
    {
        return async (context, cancellationToken) =>
        {
            string? refusal = Refusal(context.Params?.Name, context.Params?.Arguments);
            return refusal != null
                ? OutlookTools.RetiredArgumentError(refusal)
                : await next(context, cancellationToken).ConfigureAwait(false);
        };
    }

    /// <summary>
    /// The refusal for a call to <paramref name="tool"/> with <paramref name="arguments"/>, or null when it
    /// passes no retired argument. Pure, so T1 pins every message.
    /// </summary>
    internal static string? Refusal(string? tool, IDictionary<string, JsonElement>? arguments)
    {
        if (tool == null || arguments == null || !ByTool.TryGetValue(tool, out IReadOnlyList<string>? retired))
        {
            return null;
        }

        foreach (string name in retired)
        {
            if (arguments.TryGetValue(name, out JsonElement value) && value.ValueKind != JsonValueKind.Null)
            {
                return Describe(tool, name, value.ValueKind == JsonValueKind.String ? value.GetString() : value.ToString());
            }
        }

        return null;
    }

    private static string Describe(string tool, string argument, string? value)
    {
        string opening = tool + " no longer takes '" + argument + "' - it was removed when folders and stores became "
            + "addressed by ID; names and paths are for display only. ";
        switch (argument)
        {
            case "folder":
                return opening + OutlookIds.FindFolderPointer(value) + " Then pass it as folder_id.";
            case "store":
                return opening + "List the stores with list_accounts and pass the store's storeId as store_id"
                    + (tool == "move_mail"
                        ? " - or nothing at all: move_mail's folder_id names its store."
                        : ".");
            case "create_folder":
                return opening + "move_mail never creates folders now: make the folder with create_folder(parent_id, "
                    + "name) first, then pass its folderId to move_mail as folder_id.";
            default:
                return opening;
        }
    }
}
