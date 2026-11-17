using System.Collections.Concurrent;
using System.Security.Cryptography;
using System.Text;
using OutlookAI.Core.Com;
using OutlookAI.Core.Mapi;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Synthetic folder and store IDs for T1 stand-in sessions (Q131-Q136): every store name gets a
/// deterministic PST-format root folder entry id - four zero flag bytes, a 16-byte provider UID drawn
/// from the name, node id 0x8022 - and every folder path in it a deterministic folder entry id with
/// its own node id. A stand-in answers <c>TryDescribeFolder</c> for any ID made here
/// (<see cref="Describe"/>), so a test addresses a folder exactly as an agent does: by its ID. Every
/// id is synthetic - live ids never enter the repository.
/// </summary>
internal static class TestFolders
{
    private static readonly ConcurrentDictionary<string, (string Store, string Path)> Known =
        new(StringComparer.OrdinalIgnoreCase);

    /// <summary>The 16-byte provider UID of a store name.</summary>
    internal static byte[] Uid(string store)
    {
        byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes("store:" + store));
        return hash.Take(16).ToArray();
    }

    /// <summary>A store's root folder entry id (hex).</summary>
    internal static string RootEntryId(string store) => EntryId(store, 0x8022);

    /// <summary>A store's ID as the tools show it.</summary>
    internal static string StoreId(string store) => OutlookIds.StoreIdOfEntryId(RootEntryId(store))!;

    /// <summary>A folder's entry id (hex): its node id drawn from the path, folder type 0x02.</summary>
    internal static string FolderEntryId(string store, string path)
    {
        byte[] hash = SHA256.HashData(Encoding.UTF8.GetBytes("folder:" + store + "|" + path));
        uint nid = ((BitConverter.ToUInt32(hash, 0) & 0x00FFFFE0u) | 0x02u) | 0x00010000u;
        string entryId = EntryId(store, nid);
        Known[entryId] = (store, path);
        return entryId;
    }

    /// <summary>A folder's ID as the tools show it.</summary>
    internal static string FolderId(string store, string path) => OutlookIds.FolderIdFromEntryId(FolderEntryId(store, path))!;

    /// <summary>The detail, given the root folder entry id of its store.</summary>
    internal static ComStoreDetail WithRoot(ComStoreDetail d)
    {
        return new ComStoreDetail(
            d.DisplayName, d.StoreId, d.ExchangeStoreType, d.IsCachedExchange, d.NameUnreadable, d.FilePath,
            d.MappingSignatureHex, d.ExchangeProfileSectionHex, d.OwnName, RootEntryId(d.DisplayName));
    }

    /// <summary>Every detail given its root folder entry id.</summary>
    internal static IReadOnlyList<ComStoreDetail> WithRoots(IEnumerable<ComStoreDetail> details) => details.Select(WithRoot).ToList();

    /// <summary>
    /// A stand-in's answer to <c>TryDescribeFolder(folderEntryId, out error)</c> for a folder made by
    /// <see cref="FolderEntryId"/>: its names split on '/', a mail folder nothing refuses. Sets the out
    /// slot; <c>FolderNotFound</c> for an id it never made.
    /// </summary>
    internal static ComFolderDetail? Describe(object?[]? args, string? storeEntryId = null)
    {
        string entryId = (args?[0] as string) ?? string.Empty;
        if (!Known.TryGetValue(entryId, out (string Store, string Path) folder))
        {
            if (args != null && args.Length > 1)
            {
                args[1] = "FolderNotFound";
            }

            return null;
        }

        if (args != null && args.Length > 1)
        {
            args[1] = null;
        }

        string[] segments = folder.Path.Split('/');
        return new ComFolderDetail(
            entryId, segments[^1], segments, folder.Store, storeEntryId, RootEntryId(folder.Store),
            defaultItemType: 0, itemCount: 1, unreadCount: 0, isStoreRoot: false, refusedAsTarget: null);
    }

    private static string EntryId(string store, uint nid)
    {
        byte[] id = new byte[24];
        Array.Copy(Uid(store), 0, id, 4, 16);
        BitConverter.GetBytes(nid).CopyTo(id, 20);
        return Convert.ToHexString(id);
    }
}
