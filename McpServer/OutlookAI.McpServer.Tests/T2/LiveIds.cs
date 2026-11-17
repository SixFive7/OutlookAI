using OutlookAI.Core.Services;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The IDs a live test addresses stores and folders by (Q131-Q136), looked up the way an agent looks
/// them up - through list_accounts, list_folders and create_folder - never built from names. Every
/// lookup is by the product's own read-only tools; the one write, <see cref="EnsureTestFolder"/>,
/// creates only folders carrying <see cref="LiveOutlookTestMailer.TestFolderNamePrefix"/>, through the
/// product's own create_folder, and refuses any other name.
/// </summary>
internal static class LiveIds
{
    /// <summary>The storeId of the ONE store list_accounts shows under <paramref name="storeDisplayName"/>; fails on none or several.</summary>
    internal static string StoreId(MailService service, string storeDisplayName)
    {
        StoreView store = Store(service, storeDisplayName);
        return store.StoreId ?? throw new InvalidOperationException("list_accounts shows no storeId for '" + storeDisplayName + "'.");
    }

    /// <summary>The rootFolderId of that store - parent_id for a top-level create_folder.</summary>
    internal static string RootFolderId(MailService service, string storeDisplayName)
    {
        StoreView store = Store(service, storeDisplayName);
        return store.RootFolderId ?? throw new InvalidOperationException("list_accounts shows no rootFolderId for '" + storeDisplayName + "'.");
    }

    /// <summary>The folderId of the folder at <paramref name="path"/> (names joined by '/') in the store; fails when there is none.</summary>
    internal static string FolderId(MailService service, string storeDisplayName, string path)
    {
        return TryFolderId(service, storeDisplayName, path)
            ?? throw new InvalidOperationException("list_folders shows no folder '" + path + "' in '" + storeDisplayName + "'.");
    }

    /// <summary>The folderId of the folder at <paramref name="path"/>, or null when list_folders shows none.</summary>
    internal static string? TryFolderId(MailService service, string storeDisplayName, string path)
    {
        FolderView? folder = Folders(service, storeDisplayName)
            .FirstOrDefault(f => string.Equals(f.Path, path, StringComparison.OrdinalIgnoreCase));
        return folder?.FolderId;
    }

    /// <summary>
    /// The folderId of the ONE folder of the store whose own name is <paramref name="name"/>, as find_folder
    /// offers it (an exact match); fails on none or several - a test names a folder it knows to be unique.
    /// </summary>
    internal static string FolderIdByName(MailService service, string storeDisplayName, string name)
    {
        List<FolderCandidateView> exact = service.FindFolder(name, StoreId(service, storeDisplayName)).Candidates
            .Where(c => c.Match == "exact")
            .ToList();
        if (exact.Count != 1)
        {
            throw new InvalidOperationException(
                "find_folder offers " + exact.Count + " folders named '" + name + "' in '" + storeDisplayName + "' - a test names exactly one.");
        }

        return exact[0].FolderId;
    }

    /// <summary>Every folder list_folders shows for the store, all pages (<see cref="LiveFolderProbe.FolderTree"/>).</summary>
    internal static IReadOnlyList<FolderView> Folders(MailService service, string storeDisplayName)
    {
        return LiveFolderProbe.FolderTree(service, storeDisplayName);
    }

    /// <summary>
    /// The folderId of the test folder at <paramref name="path"/> in the store, each missing part made
    /// with create_folder - the product's own folder-creating write since Q134 (i). Every part it MAKES
    /// must carry <see cref="LiveOutlookTestMailer.TestFolderNamePrefix"/> (parts that exist already, such as
    /// Inbox, are only passed through); a name without it is refused before anything is created. Each
    /// part is ONE name: a '/' inside a test folder's name is not expressible here, and
    /// <see cref="CreateTestFolder"/> takes such a name directly.
    /// </summary>
    internal static string EnsureTestFolder(MailService service, string storeDisplayName, string path)
    {
        string[] parts = path.Split('/');
        string parentId = RootFolderId(service, storeDisplayName);
        string walked = string.Empty;
        foreach (string part in parts)
        {
            walked = walked.Length == 0 ? part : walked + "/" + part;
            string? existing = TryFolderId(service, storeDisplayName, walked);
            parentId = existing ?? CreateTestFolder(service, storeDisplayName, parentId, part).FolderId;
        }

        return parentId;
    }

    /// <summary>
    /// create_folder for a TEST folder: the name must carry <see cref="LiveOutlookTestMailer.TestFolderNamePrefix"/>,
    /// or nothing is asked of the product at all.
    /// </summary>
    internal static CreateFolderOutcome CreateTestFolder(MailService service, string storeDisplayName, string parentId, string name)
    {
        if (!name.StartsWith(LiveOutlookTestMailer.TestFolderNamePrefix, StringComparison.Ordinal))
        {
            throw new InvalidOperationException(
                "Refused: a live test creates only folders named with '" + LiveOutlookTestMailer.TestFolderNamePrefix
                + "', and '" + name + "' is not one.");
        }

        LiveStoreWriteGuard.Writable(storeDisplayName, StoreWriteKind.Folder, "create_folder");
        return service.CreateFolder(parentId, name);
    }

    /// <summary>
    /// A folder ID that no folder has: the store's root folder ID with its folder part replaced by one
    /// no folder is given (a PST node id far past any made, an Exchange counter of all ones).
    /// </summary>
    internal static string UnknownFolderIdIn(MailService service, string storeDisplayName)
    {
        string root = RootFolderId(service, storeDisplayName);
        if (!OutlookAI.Core.Mapi.OutlookIds.TryDecodeFolderId(root, "root", out string hex, out string? problem))
        {
            throw new InvalidOperationException(problem);
        }

        byte[] bytes = Convert.FromHexString(hex);
        if (bytes.Length == 24)
        {
            BitConverter.GetBytes(0x00FFFFE2u).CopyTo(bytes, 20);
        }
        else if (bytes.Length == 46)
        {
            for (int i = 38; i < 44; i++)
            {
                bytes[i] = 0xFF;
            }
        }
        else
        {
            bytes[^1] ^= 0x5A;
        }

        return OutlookAI.Core.Mapi.OutlookIds.FolderIdFromEntryId(bytes);
    }

    /// <summary>
    /// The stores list_accounts showed, by display name, for the rest of the run: a store's ID does not change
    /// while it stays attached, and list_accounts probes the index once per store, which a test asking for IDs
    /// inside a loop would otherwise pay on every call.
    /// </summary>
    private static readonly System.Collections.Concurrent.ConcurrentDictionary<string, StoreView> Stores =
        new(StringComparer.OrdinalIgnoreCase);

    private static StoreView Store(MailService service, string storeDisplayName)
    {
        return Stores.GetOrAdd(storeDisplayName, name => LookUpStore(service, name));
    }

    private static StoreView LookUpStore(MailService service, string storeDisplayName)
    {
        List<StoreView> named = service.ListAccounts().Stores
            .Where(s => string.Equals(s.DisplayName, storeDisplayName, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (named.Count != 1)
        {
            throw new InvalidOperationException(
                named.Count + " stores of this profile are named '" + storeDisplayName + "' - a live test addresses exactly one.");
        }

        return named[0];
    }
}
