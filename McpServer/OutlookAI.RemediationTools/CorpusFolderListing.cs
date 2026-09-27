using System.Globalization;
using OutlookAI.Core.Com;

namespace OutlookAI.RemediationTools;

/// <summary>What the non-creating resolver says about ONE default folder of a store.</summary>
/// <param name="FolderId">The <c>OlDefaultFolders</c> value asked about.</param>
/// <param name="Resolution">Resolved, Absent or Unreadable - never by the lookup that creates one.</param>
/// <param name="Invisible">
/// Resolved to a folder that is NOT in the visible tree - the store's hidden root, as
/// <c>GetDefaultFolder(olFolderInbox)</c> answered for a PST attached with <c>AddStoreEx</c>.
/// </param>
/// <param name="Name">The folder's display name when it resolved to a visible folder.</param>
public sealed record ListedDefaultFolder(
    int FolderId,
    OutlookComSession.DefaultFolderResolution Resolution,
    bool Invisible,
    string? Name);

/// <summary>One folder of a store's visible tree.</summary>
/// <param name="Path">Its path under the store's root folder, '/'-separated.</param>
/// <param name="Depth">0 for a direct child of the root folder.</param>
/// <param name="Items">Its item count, or null when that would not read.</param>
/// <param name="DefaultItemType"><c>Folder.DefaultItemType</c> (0 mail, 1 appointment, 2 contact, 3 task ...), or null.</param>
public sealed record ListedFolder(string Path, int Depth, int? Items, int? DefaultItemType);

/// <summary>A read-only picture of a store: its default folders and its visible folder tree.</summary>
/// <param name="DefaultFolders">One entry per <see cref="CorpusFolderListing.DefaultFolderIds"/>, in that order.</param>
/// <param name="Tree">Every folder under the root folder, depth-first, parents before children.</param>
/// <param name="Truncated">True when the walk stopped at its folder or depth cap, so the tree is incomplete.</param>
public sealed record FolderListing(
    IReadOnlyList<ListedDefaultFolder> DefaultFolders,
    IReadOnlyList<ListedFolder> Tree,
    bool Truncated);

/// <summary>
/// <c>corpus-folders</c>: what a store's folders ARE, read without changing them - added 2026-09-27
/// for the first guest run of population v2, so the questions only a folder list answers (does a PST
/// attached with <c>AddStoreEx</c> still lack its default folders after a build; does a teardown leave
/// emptied folders in Deleted Items) are answered by the project's own tested tool rather than by an
/// improvised COM script. Pure: the COM half collects a <see cref="FolderListing"/>, and this renders it.
/// </summary>
public static class CorpusFolderListing
{
    /// <summary>The default folders listed, in the order printed: the arrival and mail folders, then the undated kinds'.</summary>
    public static IReadOnlyList<int> DefaultFolderIds { get; } = new[] { 6, 5, 3, 4, 16, 23, 9, 10, 13 };

    /// <summary>How far the tree walk goes before it stops and says so.</summary>
    public const int MaxDepth = 8;

    /// <summary>How many folders the tree walk lists before it stops and says so.</summary>
    public const int MaxFolders = 500;

    /// <summary>The English label of a default folder id, for the listing.</summary>
    public static string Label(int folderId) => folderId switch
    {
        3 => "Deleted Items",
        4 => "Outbox",
        5 => "Sent Items",
        6 => "Inbox",
        9 => "Calendar",
        10 => "Contacts",
        13 => "Tasks",
        16 => "Drafts",
        23 => "Junk Email",
        _ => folderId.ToString(CultureInfo.InvariantCulture),
    };

    /// <summary>The kind a <c>Folder.DefaultItemType</c> value names.</summary>
    public static string KindOf(int? defaultItemType) => defaultItemType switch
    {
        null => "kind?",
        0 => "mail",
        1 => "calendar",
        2 => "contacts",
        3 => "tasks",
        4 => "journal",
        5 => "notes",
        _ => "type " + defaultItemType.Value.ToString(CultureInfo.InvariantCulture),
    };

    /// <summary>The listing as text, one line per default folder and per folder of the tree.</summary>
    public static IReadOnlyList<string> Render(string store, FolderListing listing)
    {
        ArgumentNullException.ThrowIfNull(store);
        ArgumentNullException.ThrowIfNull(listing);
        var lines = new List<string>
        {
            $"== folders of '{store}' - read-only; no folder was asked for by the lookup that creates one ==",
            "default folders:",
        };

        foreach (ListedDefaultFolder folder in listing.DefaultFolders)
        {
            string state = folder.Resolution switch
            {
                OutlookComSession.DefaultFolderResolution.Absent => "ABSENT",
                OutlookComSession.DefaultFolderResolution.Unreadable => "UNREADABLE",
                _ when folder.Invisible => "resolves to a folder NOT in the visible tree (the store's hidden root)",
                _ => "'" + (folder.Name ?? "(nameless)") + "'",
            };
            lines.Add($"  {Label(folder.FolderId),-14} {state}");
        }

        lines.Add("tree (under the root folder):");
        if (listing.Tree.Count == 0)
        {
            lines.Add("  (no folder)");
        }

        foreach (ListedFolder folder in listing.Tree)
        {
            string name = folder.Path.Contains('/', StringComparison.Ordinal)
                ? folder.Path.Substring(folder.Path.LastIndexOf('/') + 1)
                : folder.Path;
            string items = folder.Items == null ? "?" : folder.Items.Value.ToString(CultureInfo.InvariantCulture);
            string created = name.Contains(CorpusManifest.CreatedFolderPrefix, StringComparison.Ordinal)
                ? "  [made by the corpus tool]"
                : string.Empty;
            lines.Add($"  {new string(' ', 2 * folder.Depth)}{name}  items={items}  {KindOf(folder.DefaultItemType)}{created}");
        }

        if (listing.Truncated)
        {
            lines.Add($"  ... TRUNCATED: the walk stops at {MaxFolders.ToString(CultureInfo.InvariantCulture)} folders or depth "
                + $"{MaxDepth.ToString(CultureInfo.InvariantCulture)}, so this tree is not the whole store.");
        }

        return lines;
    }
}
