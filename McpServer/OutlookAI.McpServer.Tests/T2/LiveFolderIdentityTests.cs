using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// T2 live measurement and permanent pin of what a FOLDER's id is and what it survives (Q114 and
/// Q115, 2026-10-03). The maintainer decided that tools will address folders only by a unique id,
/// with names kept for display; these tests pin the facts that decision rests on, on whatever
/// store the run's hub is - a PST on the test guests, an Exchange mailbox on the Exchange guest.
/// <list type="bullet">
/// <item><see cref="FolderId_SurvivesRenameAndMoveWithinItsStore_AndARecreatedFolderGetsANewOne"/>:
/// the id's layout (<see cref="FolderEntryIdLayout"/>), that <c>GetFolderFromID</c> opens it with
/// and without the store id and from either case of hex, and what a rename, a move within the
/// store, a soft delete and a delete-and-recreate under the same name each do to it.</item>
/// <item><see cref="IndexRows_CarryTheirOwnNodeId_SoAPstFolderIdFindsItsScopeWithoutOutlook"/>: what the
/// search index keeps that could turn a folder id into an index scope without Outlook (Q115) - every
/// property the property system can name, read off a folder's own row and an item row. Measured: no
/// row holds a folder's whole id, but every row holds its OWN node id in <c>System.ProviderItemID</c>
/// (<c>N</c> + ten decimal digits), so a PST folder id's scope is found from the index alone - the root
/// by the store UID in its items' URLs, the folder's row by its node id - which this test does end to
/// end and compares with the scope built from the folder's names.</item>
/// <item><see cref="IndexRefilesAFolderRenameAndMove_WithinTheCeiling"/>: how long the index takes
/// to file a folder's items, and the folder's own row, under its new path after a rename and after a
/// move, and whether the old and new paths overlap or leave a gap - what bounds how stale an id-to-scope
/// lookup can be, and whether a folder's row keeps its node id through both.</item>
/// </list>
/// <para>
/// SAFETY: every write targets the hub (S2), every item carries the tag and this run's marker (S3),
/// every folder carries <see cref="LiveOutlookTestMailer.TestFolderNamePrefix"/>, folders are created
/// only through the product's own <c>move_mail</c> and renamed, moved and deleted only through the
/// tested helpers, which refuse anything that is not a test folder; all of it is removed through the
/// tested sweeps. The index and every other store are only read.
/// </para>
/// </summary>
[Collection(LiveCollections.MoveArchive)]
[Trait("Category", "Live")]
public sealed class LiveFolderIdentityTests
{
    /// <summary>Ceiling for the index to take a seeded item, and to re-file it after a rename or a move.</summary>
    private const int IndexWaitSeconds = 300;

    /// <summary>Gap between polls of those waits.</summary>
    private const int IndexPollSeconds = 2;

    /// <summary>Per-statement bound for the polls and the property reads.</summary>
    private const int StatementTimeoutSeconds = 15;

    /// <summary>Columns per property-read statement; a statement the index refuses is split until each bad column stands alone.</summary>
    private const int PropertyBatch = 48;

    /// <summary>Largest folder whose rows one DIRECTORY statement is trusted to list in full.</summary>
    private const int MaxItemsPerFolder = 400;

    private readonly LiveMoveArchiveFixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveFolderIdentityTests(LiveMoveArchiveFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    private MailService Service => _fixture.Service;

    private static string Prefix => LiveOutlookTestMailer.TestFolderNamePrefix;

    // ------------------------------------------------------------------ 1. the id and what it survives

    [Fact]
    [Trait("Requires", "MailAccount")]
    public void FolderId_SurvivesRenameAndMoveWithinItsStore_AndARecreatedFolderGetsANewOne()
    {
        LiveOutlookTestMailer.DeleteTestFolders(Hub);
        string parent = Prefix + "-Id";
        string nameA = Prefix + "-IdA";
        string nameB = Prefix + "-IdB";
        string renamedA = Prefix + "-IdA-renamed";
        List<string> entryIds = new();
        try
        {
            StoreIdentity store = LiveOutlookTestMailer.ReadStoreIdentity(Hub);
            FolderIdentity root = LiveOutlookTestMailer.ReadFolderIdentity(Hub, store.RootFolderEntryId);
            DescribeStore(store, root);

            string itemA = FileNewTaggedItem(parent + "/" + nameA, "A", entryIds);
            FileNewTaggedItem(parent + "/" + nameB, "B", entryIds);
            string parentId = Child(store.RootFolderEntryId, parent);
            string idA = Child(parentId, nameA);
            string idB = Child(parentId, nameB);

            // The id itself.
            FolderIdentity a0 = LiveOutlookTestMailer.ReadFolderIdentity(Hub, idA);
            FolderEntryIdLayout layout = Assert.IsType<FolderEntryIdLayout>(FolderEntryIdLayout.Parse(a0.EntryId));
            FolderEntryIdLayout rootLayout = Assert.IsType<FolderEntryIdLayout>(FolderEntryIdLayout.Parse(root.EntryId));
            Describe("created", a0);
            Assert.NotEqual(FolderEntryIdKind.Unknown, layout.Kind);
            Assert.Equal(rootLayout.Kind, layout.Kind);
            Assert.Equal(rootLayout.ProviderUidHex, layout.ProviderUidHex);
            Assert.True(FolderEntryIdLayout.SameBytes(a0.Properties["PR_ENTRYID"], a0.EntryId), "Folder.EntryID is not the folder's PR_ENTRYID");
            Assert.True(FolderEntryIdLayout.SameBytes(a0.StoreId, store.StoreId), "Folder.StoreID is not its store's StoreID");
            if (layout.Kind == FolderEntryIdKind.Pst)
            {
                Assert.Equal(FolderEntryIdLayout.NidTypeNormalFolder, layout.NidType);
                Assert.True(
                    FolderEntryIdLayout.SameBytes(layout.ProviderUidHex, store.Properties["PR_RECORD_KEY"]),
                    "a PST folder id's provider UID is not its store's PR_RECORD_KEY");

                // A PST folder's own record key is its node id alone - store-scoped, so every PST's Inbox
                // shares it (MAPI: "the scope of a record key for folders and messages is the message store").
                Assert.True(
                    FolderEntryIdLayout.SameBytes(layout.FolderPartHex, a0.Properties["PR_RECORD_KEY"]),
                    "a PST folder's PR_RECORD_KEY is not its node id");
            }

            Assert.True(a0.OpensWithoutStoreId == true, "GetFolderFromID without the store id did not open the same folder: " + a0.WithoutStoreIdError);
            Assert.True(a0.OpensFromLowercaseHex == true, "GetFolderFromID did not open the folder from lower-case hex");

            // A rename.
            TestFolderChange rename = LiveOutlookTestMailer.RenameTestFolder(Hub, idA, renamedA);
            FolderIdentity a1 = LiveOutlookTestMailer.ReadFolderIdentity(Hub, idA);
            Describe("renamed", a1);
            _output.WriteLine("rename: id on the same object " + Same(rename.EntryIdOnSameObject, idA) + ", re-opened " + Same(rename.EntryIdAfter, idA));
            Assert.Equal(renamedA, a1.Name);
            Assert.True(FolderEntryIdLayout.SameBytes(a1.EntryId, idA), "a rename changed the folder's id");
            AssertSameCandidates(a0, a1, "a rename");
            (bool itemOpens1, string? itemParent1, string? itemError1) = LiveOutlookTestMailer.ReadItemParent(Hub, itemA);
            _output.WriteLine("item after the rename: opens " + itemOpens1 + ", in the folder " + Same(itemParent1, idA) + " " + itemError1);
            Assert.True(itemOpens1 && FolderEntryIdLayout.SameBytes(itemParent1, idA), "the item's own id did not survive its folder's rename");

            // A move within the store, under the sibling.
            TestFolderChange move = LiveOutlookTestMailer.MoveTestFolder(Hub, idA, idB);
            _output.WriteLine("move: id on the same object " + Same(move.EntryIdOnSameObject, idA) + ", found under the new parent " + Same(move.EntryIdAfter, idA));
            Assert.False(string.IsNullOrEmpty(move.EntryIdAfter), "the moved folder was not found under its new parent");
            FolderIdentity a2 = LiveOutlookTestMailer.ReadFolderIdentity(Hub, move.EntryIdAfter!);
            Describe("moved", a2);
            Assert.True(FolderEntryIdLayout.SameBytes(a2.ParentEntryId, idB), "the moved folder's parent is not the folder it was moved into");
            Assert.True(FolderEntryIdLayout.SameBytes(move.EntryIdAfter, idA), "a move within the store changed the folder's id");
            AssertSameCandidates(a0, a2, "a move within the store");
            (bool itemOpens2, string? itemParent2, string? itemError2) = LiveOutlookTestMailer.ReadItemParent(Hub, itemA);
            _output.WriteLine("item after the move: opens " + itemOpens2 + ", in the folder " + Same(itemParent2, move.EntryIdAfter) + " " + itemError2);
            Assert.True(itemOpens2 && FolderEntryIdLayout.SameBytes(itemParent2, move.EntryIdAfter), "the item's own id did not survive its folder's move");

            // A soft delete: Folder.Delete, which is what Delete on a folder does in Outlook.
            string deletedItems = LiveOutlookTestMailer.ReadDeletedItemsEntryId(Hub);
            TestFolderChange delete = LiveOutlookTestMailer.SoftDeleteTestFolder(Hub, move.EntryIdAfter!);
            (bool oldOpens, string? oldName, string? oldPath, string? oldError) = LiveOutlookTestMailer.TryOpenFolder(Hub, idA);
            _output.WriteLine(
                "soft delete: found under Deleted Items " + (delete.EntryIdAfter != null) + ", with the id it had "
                + Same(delete.EntryIdAfter, idA) + "; the old id opens " + oldOpens
                + (oldOpens ? " as '" + oldName + "' at '" + oldPath + "'" : " - " + oldError));
            Assert.False(string.IsNullOrEmpty(delete.EntryIdAfter), "the deleted folder was not found under Deleted Items");
            FolderIdentity a3 = LiveOutlookTestMailer.ReadFolderIdentity(Hub, delete.EntryIdAfter!);
            Describe("soft-deleted", a3);
            Assert.True(FolderEntryIdLayout.SameBytes(a3.ParentEntryId, deletedItems), "the deleted folder's parent is not Deleted Items");

            // Recreated under the original name and parent: a new folder, so a new id.
            FileNewTaggedItem(parent + "/" + nameA, "A again", entryIds);
            string idAgain = Child(parentId, nameA);
            Describe("recreated", LiveOutlookTestMailer.ReadFolderIdentity(Hub, idAgain));
            Assert.False(FolderEntryIdLayout.SameBytes(idAgain, idA), "a new folder under a deleted folder's name and parent got the deleted folder's id");
            Assert.False(FolderEntryIdLayout.SameBytes(idAgain, delete.EntryIdAfter), "a new folder got the id of the folder in Deleted Items");

            _output.WriteLine(
                "MEASURED: rename keeps the id; a move within the store keeps it " + FolderEntryIdLayout.SameBytes(move.EntryIdAfter, idA)
                + "; a soft delete keeps it " + FolderEntryIdLayout.SameBytes(delete.EntryIdAfter, idA)
                + " and the old id then opens " + (oldOpens ? "the folder in Deleted Items" : "nothing")
                + "; a folder recreated under the same name gets a new id.");
        }
        finally
        {
            CleanUp(entryIds);
        }

        AssertHubClean();
    }

    // ------------------------------------------------------------------ 2. what the index keeps about a folder

    [Fact]
    [Trait("Requires", "SearchIndex")]
    public void IndexRows_CarryTheirOwnNodeId_SoAPstFolderIdFindsItsScopeWithoutOutlook()
    {
        IIndexClient client = IndexClientFactory.CreateAuto(out string providerReport);
        _output.WriteLine("index client: " + providerReport);
        IReadOnlyList<string> properties = PropertySystemCatalog.CanonicalNames();
        _output.WriteLine("property system: " + properties.Count + " property descriptions");
        Assert.True(properties.Count > 100, "the property system named almost nothing - this test would search nothing");

        List<IndexRoot> roots = ListIndexRoots(client);
        _output.WriteLine("index roots: " + roots.Count);
        int foldersCompared = 0;
        int itemRowsCompared = 0;
        int propertiesRead = 0;
        int mappedEndToEnd = 0;
        foreach (string storeName in _fixture.Settings.IndexedStores)
        {
            StoreIdentity store = LiveOutlookTestMailer.ReadStoreIdentity(storeName);
            IndexRoot? root = MatchRoot(roots, store);
            if (root == null)
            {
                _output.WriteLine("[" + storeName + "] no index root carries this store's hash - not compared");
                continue;
            }

            string rootUrl = root.Url.TrimEnd('/');
            FolderIdentity rootFolder = LiveOutlookTestMailer.ReadFolderIdentity(storeName, store.RootFolderEntryId);
            _output.WriteLine(
                "[" + storeName + "] index root " + rootUrl + "; root folder id " + FolderEntryIdLayout.Parse(rootFolder.EntryId)?.Describe()
                + ", opens without the store id " + rootFolder.OpensWithoutStoreId + ", from lower-case hex " + rootFolder.OpensFromLowercaseHex);
            Assert.True(rootFolder.OpensWithoutStoreId == true, "[" + storeName + "] GetFolderFromID without the store id did not open its root: " + rootFolder.WithoutStoreIdError);
            IReadOnlyList<FolderListing> listing = LiveOutlookTestMailer.ListFolderIdentities(storeName, maxFolders: 60, maxDepth: 1);
            string? rootFolderPath = rootFolder.FolderPath;

            // Top-level folders small enough that one DIRECTORY statement lists every row under them.
            foreach (FolderListing folder in listing.Where(f => f.Depth == 1 && f.ItemCount > 0 && f.ItemCount <= MaxItemsPerFolder && Plain(f.Name)).Take(3))
            {
                IReadOnlyList<string>? segments = RelativeSegments(rootFolderPath, folder.FolderPath);
                if (segments == null || segments.Any(s => !Plain(s)))
                {
                    continue;
                }

                string folderUrl = rootUrl + "/0/" + string.Join("/", segments.Select(MapiUrlSegment.Encode));
                string parentUrl = segments.Count == 1
                    ? rootUrl + "/0"
                    : rootUrl + "/0/" + string.Join("/", segments.Take(segments.Count - 1).Select(MapiUrlSegment.Encode));
                byte[] folderId = Convert.FromHexString(folder.EntryId);

                Dictionary<string, object?>? folderRow = ReadRow(client, parentUrl, folderUrl, properties, out int readFolder)
                    ?? ReadRow(client, parentUrl + "/", folderUrl, properties, out readFolder);
                propertiesRead += readFolder;
                List<string> itemUrls = DirectoryUrls(client, folderUrl, MaxItemsPerFolder);
                string directoryForItems = folderUrl;
                if (itemUrls.Count == 0)
                {
                    itemUrls = DirectoryUrls(client, folderUrl + "/", MaxItemsPerFolder);
                    directoryForItems = folderUrl + "/";
                }
                _output.WriteLine(
                    "  folder '" + string.Join("/", segments) + "' (" + folder.ItemCount + " items): own row "
                    + (folderRow != null ? "FOUND" : "absent") + " at " + folderUrl + "; " + itemUrls.Count + " row(s) under it");
                if (folderRow == null || itemUrls.Count == 0)
                {
                    continue;
                }

                ReportRow("folder row", folderRow);
                FolderEntryIdLayout layout = Assert.IsType<FolderEntryIdLayout>(FolderEntryIdLayout.Parse(folder.EntryId));
                string? folderProviderItemId = Text(folderRow, "System.ProviderItemID");
                _output.WriteLine(
                    "  folder row: carries the whole folder id in " + Carriers(folderRow, folderId) + "; System.ProviderItemID "
                    + (folderProviderItemId ?? "(none)") + "; the id's node part " + (layout.PstProviderItemId ?? "(not a PST id)"));
                if (layout.Kind == FolderEntryIdKind.Pst)
                {
                    Assert.Equal(layout.PstProviderItemId, folderProviderItemId);

                    // End to end, with NO Outlook call: the folder id alone -> its store's index root (by the
                    // store UID its items' URLs carry) -> its own row (by System.ProviderItemID) -> the scope.
                    Stopwatch mapping = Stopwatch.StartNew();
                    string? mapped = FolderScopeFromIndexOnly(client, roots, folder.EntryId, out string how);
                    mapping.Stop();
                    int underMapped = mapped == null ? -1 : CountUnder(client, mapped);
                    _output.WriteLine(
                        "  id -> scope from the index alone: " + (mapped ?? "(none)") + " in " + mapping.ElapsedMilliseconds + " ms (" + how
                        + "); rows under it " + underMapped + " for " + folder.ItemCount + " item(s) in Outlook");
                    Assert.Equal(folderUrl.TrimEnd('/'), mapped?.TrimEnd('/'));
                    mappedEndToEnd++;
                }

                foldersCompared++;

                string? itemUrl = itemUrls.FirstOrDefault(u => MapiItemUrl.TryParse(u, out MapiItemUrl? p) && p!.EncodedItemSegment != null && !p.IsAttachment);
                if (itemUrl == null)
                {
                    continue;
                }

                Dictionary<string, object?>? itemRow = ReadRow(client, directoryForItems, itemUrl, properties, out int readItem);
                propertiesRead += readItem;
                Assert.NotNull(itemRow);
                ReportRow("item row", itemRow!);
                itemRowsCompared++;

                DecodedEntryId? decoded = null;
                Assert.True(MapiItemUrl.TryParse(itemUrl, out MapiItemUrl? parsed) && parsed!.TryDecodeEntryId(out decoded));
                string itemProviderItemId = FolderEntryIdLayout.ProviderItemIdOf(
                    BitConverter.ToUInt32(Convert.FromHexString(decoded!.NidHex), 0));
                string? itemRowProviderItemId = Text(itemRow!, "System.ProviderItemID");
                _output.WriteLine(
                    "  item row: carries the whole folder id in " + Carriers(itemRow!, folderId) + "; its URL id is 24 bytes, store UID equal to the folder id's "
                    + string.Equals(decoded.StoreUidHex, layout.ProviderUidHex, StringComparison.OrdinalIgnoreCase)
                    + ", node id type 0x" + decoded.NidLowFiveBits.ToString("X2", CultureInfo.InvariantCulture)
                    + "; System.ProviderItemID " + (itemRowProviderItemId ?? "(none)") + " is its OWN node id "
                    + string.Equals(itemRowProviderItemId, itemProviderItemId, StringComparison.Ordinal)
                    + " - an item row names its folder only by its URL path");
                if (layout.Kind == FolderEntryIdKind.Pst)
                {
                    Assert.Equal(itemProviderItemId, itemRowProviderItemId);
                }
            }
        }

        _output.WriteLine(
            "compared " + foldersCompared + " folder row(s) and " + itemRowsCompared + " item row(s), " + mappedEndToEnd
            + " PST folder id(s) mapped to their scope from the index alone; " + propertiesRead + " property values read in all");
        Assert.True(foldersCompared > 0, "no folder of any indexed store was compared - this run proves nothing");
    }

    // ------------------------------------------------------------------ 3. how fast the index follows a rename and a move

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "MailAccount")]
    public void IndexRefilesAFolderRenameAndMove_WithinTheCeiling()
    {
        LiveOutlookTestMailer.DeleteTestFolders(Hub);
        string parent = Prefix + "-Lag";
        string nameX = Prefix + "-LagX";
        string nameY = Prefix + "-LagY";
        string renamedX = Prefix + "-LagX-renamed";
        string term = "fidlag" + Marker;
        List<string> entryIds = new();
        try
        {
            IIndexClient client = IndexClientFactory.CreateAuto(out string providerReport);
            _output.WriteLine("index client: " + providerReport);
            StoreIdentity store = LiveOutlookTestMailer.ReadStoreIdentity(Hub);
            FileNewTaggedItem(parent + "/" + nameX, term, entryIds);
            FileNewTaggedItem(parent + "/" + nameY, term + "y", entryIds);
            string parentId = Child(store.RootFolderEntryId, parent);
            string idX = Child(parentId, nameX);
            string idY = Child(parentId, nameY);

            string? hubPrefix = WaitForItemPrefix(client, term, "/0/" + parent + "/" + nameX + "/");
            Assert.True(hubPrefix != null, "the seeded item never reached the index under its folder within " + IndexWaitSeconds + " s");
            string urlX = hubPrefix + "/0/" + parent + "/" + nameX;
            string urlRenamed = hubPrefix + "/0/" + parent + "/" + renamedX;
            string urlMoved = hubPrefix + "/0/" + parent + "/" + nameY + "/" + renamedX;

            // The folder's own row, and the node id it carries: what an id-to-scope lookup reads (Q115).
            FolderEntryIdLayout? layoutX = FolderEntryIdLayout.Parse(idX);
            string? pid = layoutX?.PstProviderItemId;
            string? rowBefore = pid == null ? null : UrlByProviderItemId(client, hubPrefix!, pid);
            _output.WriteLine("folder row by System.ProviderItemID " + (pid ?? "(not a PST id)") + ": " + (rowBefore ?? "(none)"));
            if (pid != null)
            {
                Assert.Equal(urlX, rowBefore);
            }

            LiveOutlookTestMailer.RenameTestFolder(Hub, idX, renamedX);
            Timeline rename = Follow(client, term, urlX, urlRenamed, hubPrefix!, pid);
            _output.WriteLine("rename: " + rename);

            TestFolderChange move = LiveOutlookTestMailer.MoveTestFolder(Hub, idX, idY);
            Assert.False(string.IsNullOrEmpty(move.EntryIdAfter));
            Timeline moved = Follow(client, term, urlRenamed, urlMoved, hubPrefix!, pid);
            _output.WriteLine("move: " + moved);

            Assert.True(rename.Settled, "after a rename the index did not settle on the new path within " + IndexWaitSeconds + " s: " + rename);
            Assert.True(moved.Settled, "after a move the index did not settle on the new path within " + IndexWaitSeconds + " s: " + moved);
        }
        finally
        {
            CleanUp(entryIds);
        }

        AssertHubClean();
    }

    // ------------------------------------------------------------------ helpers: the hub

    /// <summary>A tagged draft filed into <paramref name="path"/> by the product's own move_mail, creating the folders.</summary>
    private string FileNewTaggedItem(string path, string label, List<string> entryIds)
    {
        string subject = _fixture.TaggedSubject("folder identity " + label);
        DraftOutcome draft = Service.NewDraft(
            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft"), to: Hub, cc: null, subject: subject,
            body: "Folder identity probe " + label + " (Q114).", display: false);
        Assert.False(string.IsNullOrEmpty(draft.EntryId));
        entryIds.Add(draft.EntryId);
        MoveMailOutcome moved = Service.MoveMail(new[] { draft.EntryId }, path, createFolder: true);
        MoveItemView item = Assert.Single(moved.Items);
        Assert.True(item.Ok && item.NewEntryId != null, "move_mail did not file the probe into '" + path + "': " + item.Error);
        entryIds.Add(item.NewEntryId!);
        return item.NewEntryId!;
    }

    private string Child(string parentEntryId, string name)
    {
        string? id = LiveOutlookTestMailer.FindChildFolderEntryId(Hub, parentEntryId, name);
        Assert.False(string.IsNullOrEmpty(id), "folder '" + name + "' was not found under its parent");
        return id!;
    }

    private void DescribeStore(StoreIdentity store, FolderIdentity root)
    {
        _output.WriteLine(
            "store: StoreID " + (store.StoreId.Length / 2) + " bytes, file " + FileKind(store.FilePath)
            + ", ExchangeStoreType " + (store.ExchangeStoreType?.ToString(CultureInfo.InvariantCulture) ?? "-")
            + ", cached " + (store.IsCachedExchange?.ToString() ?? "-"));
        foreach (KeyValuePair<string, string?> p in store.Properties)
        {
            _output.WriteLine(
                "  store " + p.Key + ": " + (p.Value == null ? "absent " + store.PropertyErrors.GetValueOrDefault(p.Key) : (p.Value.Length / 2) + " bytes")
                + (p.Value != null && p.Key == "PR_ENTRYID" ? ", equal to StoreID " + FolderEntryIdLayout.SameBytes(p.Value, store.StoreId) : string.Empty));
        }

        Describe("root", root);
    }

    private void Describe(string label, FolderIdentity folder)
    {
        FolderEntryIdLayout? layout = FolderEntryIdLayout.Parse(folder.EntryId);
        _output.WriteLine(
            label + ": '" + folder.Name + "', " + (layout?.Describe() ?? "unparsable id") + ", opens without the store id "
            + (folder.OpensWithoutStoreId?.ToString() ?? "-") + ", from lower-case hex " + (folder.OpensFromLowercaseHex?.ToString() ?? "-"));
        foreach (KeyValuePair<string, string?> p in folder.Properties)
        {
            string shape = p.Value == null
                ? "absent " + folder.PropertyErrors.GetValueOrDefault(p.Key)
                : (p.Value.Length / 2) + " bytes" + (FolderEntryIdLayout.SameBytes(p.Value, folder.EntryId) ? " = the EntryID" : string.Empty)
                    + (p.Key == "PR_PARENT_ENTRYID" ? ", = the parent's EntryID " + FolderEntryIdLayout.SameBytes(p.Value, folder.ParentEntryId) : string.Empty)
                    + (p.Key == "PR_STORE_ENTRYID" ? ", = the StoreID " + FolderEntryIdLayout.SameBytes(p.Value, folder.StoreId) : string.Empty);
            _output.WriteLine("  " + p.Key + ": " + shape);
        }
    }

    /// <summary>The source key and record key, where the store keeps them, say what the entry id says.</summary>
    private void AssertSameCandidates(FolderIdentity before, FolderIdentity after, string change)
    {
        foreach (string key in new[] { "PR_SOURCE_KEY", "PR_RECORD_KEY" })
        {
            string? b = before.Properties.GetValueOrDefault(key);
            string? a = after.Properties.GetValueOrDefault(key);
            if (b != null || a != null)
            {
                Assert.True(FolderEntryIdLayout.SameBytes(b, a), change + " changed the folder's " + key);
            }
        }
    }

    private static string Same(string? candidate, string? reference)
    {
        if (string.IsNullOrEmpty(candidate))
        {
            return "(none)";
        }

        return FolderEntryIdLayout.SameBytes(candidate, reference) ? "unchanged" : "CHANGED";
    }

    private static string FileKind(string? path)
    {
        if (string.IsNullOrEmpty(path))
        {
            return "(none)";
        }

        return Path.GetExtension(path).ToLowerInvariant();
    }

    private void CleanUp(List<string> entryIds)
    {
        foreach (string entryId in entryIds.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                LiveOutlookTestMailer.DeleteItemByEntryId(Hub, entryId, Marker);
            }
            catch (Exception)
            {
                // The stable-zero sweep below is the authority.
            }
        }

        LiveOutlookTestMailer.DeleteTestFolders(Hub);
        LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(
            Hub, Marker, folderIds: LiveOutlookTestMailer.HubSweepFolderIdsWithArchive);
    }

    private void AssertHubClean()
    {
        LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(
            Hub, Marker, folderIds: LiveOutlookTestMailer.HubSweepFolderIdsWithArchive);
        int remaining = LiveOutlookTestMailer.CountTaggedArtifactsAfterPurgingStragglers(
            Hub, Marker, LiveOutlookTestMailer.HubSweepFolderIdsWithArchive, out int stragglersPurged);
        if (stragglersPurged > 0)
        {
            _output.WriteLine($"cleanup[{Hub}]: {stragglersPurged} late-materialized artifact(s) purged (documented lag)");
        }

        Assert.Equal(0, remaining);
        int liveTestFolders = LiveOutlookTestMailer.CountLiveTestFolders(Hub, out int wedgedEmpty);
        if (wedgedEmpty > 0)
        {
            _output.WriteLine(
                $"cleanup[{Hub}]: {wedgedEmpty} empty test folder(s) wedged in Deleted Items until Outlook restarts "
                + "(documented same-session limitation, no items involved)");
        }

        Assert.Equal(0, liveTestFolders);
        _output.WriteLine(_fixture.VerifyHubReconciled());
    }

    // ------------------------------------------------------------------ helpers: the index

    private sealed record IndexRoot(string Url, uint? Hash);

    private List<IndexRoot> ListIndexRoots(IIndexClient client)
    {
        string userRoot = IndexSearchService.CurrentUserMapiRoot();
        List<IndexRoot> roots = new();
        foreach (string url in DirectoryUrls(client, userRoot.TrimEnd('/') + "/", 200))
        {
            uint? hash = MapiItemUrl.TryParse(url, out MapiItemUrl? parsed) && StoreHash.TryParseUrlHash(parsed!.StoreUrlHash, out uint value)
                ? value
                : null;
            roots.Add(new IndexRoot(url, hash));
        }

        return roots;
    }

    /// <summary>The root that carries one of the store's documented hashes (<see cref="StoreHash.Candidates"/>) and that no other store's hash could explain.</summary>
    private static IndexRoot? MatchRoot(List<IndexRoot> roots, StoreIdentity store)
    {
        bool exchange = store.ExchangeStoreType is int t && t != 3;
        IReadOnlyList<StoreHashCandidate> candidates = StoreHash.Candidates(
            store.StoreId, exchange, publicStore: false, store.Properties.GetValueOrDefault("PR_MAPPING_SIGNATURE"), null, store.FilePath);
        List<IndexRoot> matches = roots.Where(r => r.Hash is uint h && candidates.Any(c => c.Hash == h)).ToList();
        return matches.Count == 1 ? matches[0] : null;
    }

    private List<string> DirectoryUrls(IIndexClient client, string directoryUrl, int top)
    {
        string sql = "SELECT TOP " + top.ToString(CultureInfo.InvariantCulture) + " System.ItemUrl FROM SystemIndex WHERE DIRECTORY='"
            + directoryUrl.Replace("'", "''", StringComparison.Ordinal) + "'";
        List<string> urls = new();
        foreach (IReadOnlyDictionary<string, object?> row in client.ExecuteRows(sql, top, StatementTimeoutSeconds))
        {
            if (row.TryGetValue("System.ItemUrl", out object? value) && value is string url)
            {
                urls.Add(url);
            }
        }

        return urls;
    }

    /// <summary>
    /// Every property the index will return for the row at <paramref name="rowUrl"/>, read through
    /// statements listing <paramref name="directoryUrl"/>; a column the index refuses is dropped. Null
    /// when the row is not under that directory.
    /// </summary>
    private Dictionary<string, object?>? ReadRow(
        IIndexClient client, string directoryUrl, string rowUrl, IReadOnlyList<string> properties, out int valuesRead)
    {
        valuesRead = 0;
        Dictionary<string, object?> values = new(StringComparer.OrdinalIgnoreCase);
        bool found = false;
        Queue<List<string>> batches = new(properties.Chunk(PropertyBatch).Select(c => c.ToList()));
        while (batches.Count > 0)
        {
            List<string> batch = batches.Dequeue();
            string sql = "SELECT System.ItemUrl, " + string.Join(", ", batch) + " FROM SystemIndex WHERE DIRECTORY='"
                + directoryUrl.Replace("'", "''", StringComparison.Ordinal) + "'";
            IReadOnlyList<IReadOnlyDictionary<string, object?>> rows;
            try
            {
                rows = client.ExecuteRows(sql, MaxItemsPerFolder + 100, StatementTimeoutSeconds);
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                if (batch.Count > 1)
                {
                    int half = batch.Count / 2;
                    batches.Enqueue(batch.Take(half).ToList());
                    batches.Enqueue(batch.Skip(half).ToList());
                }

                continue;
            }

            IReadOnlyDictionary<string, object?>? row = rows.FirstOrDefault(r =>
                r.TryGetValue("System.ItemUrl", out object? u) && u is string s
                && string.Equals(s.TrimEnd('/'), rowUrl.TrimEnd('/'), StringComparison.Ordinal));
            if (row == null)
            {
                continue;
            }

            found = true;
            foreach (KeyValuePair<string, object?> pair in row)
            {
                if (pair.Value != null && pair.Value is not DBNull)
                {
                    values[pair.Key] = pair.Value;
                    valuesRead++;
                }
            }
        }

        return found ? values : null;
    }

    /// <summary>The properties of a row that hold the folder's whole id in any spelling - reported, not asserted.</summary>
    private static string Carriers(Dictionary<string, object?> row, byte[] folderId)
    {
        List<string> carriers = row.Where(p => FolderEntryIdLayout.ValueCarries(p.Value, folderId)).Select(p => p.Key).ToList();
        return carriers.Count == 0 ? "no property" : string.Join(", ", carriers);
    }

    private static string? Text(IReadOnlyDictionary<string, object?> row, string column)
    {
        return row.TryGetValue(column, out object? value) && value is string s ? s : null;
    }

    /// <summary>
    /// A PST folder id's index scope, found WITHOUT Outlook (Q115): the root whose item URLs carry the
    /// id's store UID (bytes 4..19), then the one row under it whose <c>System.ProviderItemID</c> is the
    /// id's node id. Null when the id is not a PST id or nothing matches - never a guess by name.
    /// </summary>
    private string? FolderScopeFromIndexOnly(IIndexClient client, List<IndexRoot> roots, string folderEntryIdHex, out string how)
    {
        FolderEntryIdLayout? layout = FolderEntryIdLayout.Parse(folderEntryIdHex);
        if (layout?.Kind != FolderEntryIdKind.Pst)
        {
            how = "not a PST-format id";
            return null;
        }

        // A byte copy of a PST that was never attached beside its original keeps the original's UID, and the
        // index - one per Windows user, every profile's stores in it - may hold both: two roots with one UID
        // are refused, never picked between.
        List<string> matching = roots
            .Select(r => r.Url.TrimEnd('/'))
            .Where(u => string.Equals(StoreUidUnder(client, u), layout.ProviderUidHex, StringComparison.OrdinalIgnoreCase))
            .ToList();
        if (matching.Count != 1)
        {
            how = matching.Count + " index roots carry the id's store UID";
            return null;
        }

        string sql = "SELECT TOP 5 System.ItemUrl FROM SystemIndex WHERE SCOPE='" + matching[0].Replace("'", "''", StringComparison.Ordinal)
            + "/' AND System.ProviderItemID='" + layout.PstProviderItemId + "'";
        List<string> urls = client.ExecuteRows(sql, 5, StatementTimeoutSeconds)
            .Select(r => Text(r, "System.ItemUrl"))
            .Where(u => u != null)
            .Select(u => u!)
            .ToList();
        how = "the one root whose items carry the store UID, then " + urls.Count + " row(s) with System.ProviderItemID " + layout.PstProviderItemId;
        return urls.Count == 1 ? urls[0] : null;
    }

    /// <summary>The store UID (bytes 4..19 of an item id) the first item URL under a root carries, or null.</summary>
    private string? StoreUidUnder(IIndexClient client, string rootUrl)
    {
        string sql = "SELECT TOP 200 System.ItemUrl FROM SystemIndex WHERE SCOPE='" + rootUrl.Replace("'", "''", StringComparison.Ordinal) + "/'";
        foreach (IReadOnlyDictionary<string, object?> row in client.ExecuteRows(sql, 200, StatementTimeoutSeconds))
        {
            if (MapiItemUrl.TryParse(Text(row, "System.ItemUrl"), out MapiItemUrl? parsed) && parsed!.TryDecodeEntryId(out DecodedEntryId? decoded))
            {
                return decoded!.StoreUidHex;
            }
        }

        return null;
    }

    /// <summary>How many rows lie under a scope (TOP-bounded).</summary>
    private int CountUnder(IIndexClient client, string scopeUrl)
    {
        string sql = "SELECT TOP " + (MaxItemsPerFolder + 100).ToString(CultureInfo.InvariantCulture) + " System.ItemUrl FROM SystemIndex WHERE SCOPE='"
            + scopeUrl.TrimEnd('/').Replace("'", "''", StringComparison.Ordinal) + "/'";
        return client.ExecuteRows(sql, MaxItemsPerFolder + 100, StatementTimeoutSeconds).Count;
    }

    private void ReportRow(string label, Dictionary<string, object?> row)
    {
        _output.WriteLine("  " + label + ": " + row.Count + " non-empty properties");
        foreach (KeyValuePair<string, object?> p in row.OrderBy(p => p.Key, StringComparer.Ordinal))
        {
            string shown = p.Value switch
            {
                byte[] b => b.Length + " bytes",
                string s => s.Length > 160 ? s.Substring(0, 160) + "..." : s,
                System.Collections.IEnumerable e => "[" + string.Join("; ", e.Cast<object?>().Take(5)) + "]",
                _ => Convert.ToString(p.Value, CultureInfo.InvariantCulture) ?? string.Empty,
            };
            _output.WriteLine("    " + p.Key + " = " + shown);
        }
    }

    private string? WaitForItemPrefix(IIndexClient client, string term, string folderMarker)
    {
        string userRoot = IndexSearchService.CurrentUserMapiRoot();
        string sql = "SELECT TOP 50 System.ItemUrl FROM SystemIndex WHERE SCOPE='" + userRoot + "' AND CONTAINS(System.Subject, '\""
            + term + "\"')";
        Stopwatch waited = Stopwatch.StartNew();
        while (waited.Elapsed.TotalSeconds < IndexWaitSeconds)
        {
            try
            {
                foreach (IReadOnlyDictionary<string, object?> row in client.ExecuteRows(sql, 50, StatementTimeoutSeconds))
                {
                    if (row.TryGetValue("System.ItemUrl", out object? value) && value is string url
                        && url.Contains(folderMarker, StringComparison.Ordinal)
                        && MapiItemUrl.TryParse(url, out MapiItemUrl? parsed))
                    {
                        _output.WriteLine("seeded item indexed after " + waited.Elapsed.TotalSeconds.ToString("F0", CultureInfo.InvariantCulture) + " s");
                        return parsed!.StorePrefix;
                    }
                }
            }
            catch (System.Data.OleDb.OleDbException)
            {
                // A lost poll costs one poll.
            }

            Thread.Sleep(TimeSpan.FromSeconds(IndexPollSeconds));
        }

        return null;
    }

    /// <summary>
    /// When the item's row left the old path, reached the new one, whether the two ever overlapped or both
    /// stood empty - and, for a PST folder, when its own row (found by its node id) showed the new path.
    /// </summary>
    private Timeline Follow(IIndexClient client, string term, string oldUrl, string newUrl, string hubPrefix, string? pid)
    {
        Stopwatch clock = Stopwatch.StartNew();
        Timeline t = new();
        while (clock.Elapsed.TotalSeconds < IndexWaitSeconds)
        {
            int? before = CountUnder(client, oldUrl, term);
            int? after = CountUnder(client, newUrl, term);
            double now = clock.Elapsed.TotalSeconds;
            t.Polls++;
            if (before == 0 && t.OldGoneAt == null)
            {
                t.OldGoneAt = now;
            }

            if (after > 0 && t.NewSeenAt == null)
            {
                t.NewSeenAt = now;
            }

            if (before > 0 && after > 0)
            {
                t.Overlapped = true;
            }

            if (before == 0 && after == 0)
            {
                t.GapSeen = true;
            }

            if (pid != null && t.FolderRowAt == null
                && string.Equals(UrlByProviderItemId(client, hubPrefix, pid), newUrl, StringComparison.Ordinal))
            {
                t.FolderRowAt = clock.Elapsed.TotalSeconds;
            }

            if (before == 0 && after > 0 && (pid == null || t.FolderRowAt != null))
            {
                t.Settled = true;
                break;
            }

            Thread.Sleep(TimeSpan.FromSeconds(IndexPollSeconds));
        }

        t.Elapsed = clock.Elapsed.TotalSeconds;
        return t;
    }

    /// <summary>The URL of the one row under the hub whose System.ProviderItemID is <paramref name="pid"/>, or null.</summary>
    private string? UrlByProviderItemId(IIndexClient client, string hubPrefix, string pid)
    {
        string sql = "SELECT TOP 5 System.ItemUrl FROM SystemIndex WHERE SCOPE='" + hubPrefix.Replace("'", "''", StringComparison.Ordinal)
            + "/' AND System.ProviderItemID='" + pid + "'";
        try
        {
            List<string> urls = client.ExecuteRows(sql, 5, StatementTimeoutSeconds)
                .Select(r => Text(r, "System.ItemUrl"))
                .Where(u => u != null)
                .Select(u => u!)
                .ToList();
            return urls.Count == 1 ? urls[0] : null;
        }
        catch (System.Data.OleDb.OleDbException)
        {
            return null;
        }
    }

    private int? CountUnder(IIndexClient client, string scopeUrl, string term)
    {
        string sql = "SELECT TOP 20 System.ItemUrl FROM SystemIndex WHERE SCOPE='" + scopeUrl.Replace("'", "''", StringComparison.Ordinal)
            + "/' AND CONTAINS(System.Subject, '\"" + term + "\"')";
        try
        {
            return client.ExecuteRows(sql, 20, StatementTimeoutSeconds).Count;
        }
        catch (System.Data.OleDb.OleDbException)
        {
            return null;
        }
    }

    private sealed class Timeline
    {
        public int Polls { get; set; }

        public double? OldGoneAt { get; set; }

        public double? NewSeenAt { get; set; }

        public bool Overlapped { get; set; }

        public bool GapSeen { get; set; }

        public bool Settled { get; set; }

        public double? FolderRowAt { get; set; }

        public double Elapsed { get; set; }

        public override string ToString()
        {
            return "settled " + Settled + " after " + Elapsed.ToString("F0", CultureInfo.InvariantCulture) + " s (" + Polls + " polls); old path empty from "
                + (OldGoneAt?.ToString("F0", CultureInfo.InvariantCulture) ?? "never") + " s, new path seen from "
                + (NewSeenAt?.ToString("F0", CultureInfo.InvariantCulture) ?? "never") + " s; both at once " + Overlapped
                + ", neither " + GapSeen + "; the folder's own row at the new path (by its node id) from "
                + (FolderRowAt?.ToString("F0", CultureInfo.InvariantCulture) ?? "never/not a PST") + " s";
        }
    }

    // ------------------------------------------------------------------ helpers: names and paths

    /// <summary>A name with none of the characters the index or Outlook's FolderPath escapes, so the plain URL is the right one.</summary>
    private static bool Plain(string name)
    {
        return name.Length > 0 && name.Trim() == name && name.IndexOfAny(new[] { '/', '\\', '%', '*', '?' }) < 0;
    }

    /// <summary>A FolderPath's segments below the store's root folder path, or null when it does not start there.</summary>
    private static IReadOnlyList<string>? RelativeSegments(string? rootFolderPath, string folderPath)
    {
        if (string.IsNullOrEmpty(rootFolderPath) || !folderPath.StartsWith(rootFolderPath + "\\", StringComparison.Ordinal))
        {
            return null;
        }

        return folderPath.Substring(rootFolderPath.Length + 1).Split('\\');
    }

    /// <summary>
    /// The canonical name of every property the Windows property system describes - the column names a
    /// SystemIndex statement may select. Read through <c>PSEnumeratePropertyDescriptions</c>.
    /// </summary>
    private static class PropertySystemCatalog
    {
        private static readonly Guid IidPropertyDescriptionList = new("1F9FC1D0-C39B-4B26-817F-011967D3440E");
        private static readonly Guid IidPropertyDescription = new("6F79D558-3E96-4549-A1D1-7D75D2288814");

        public static IReadOnlyList<string> CanonicalNames()
        {
            Guid listIid = IidPropertyDescriptionList;
            int hr = PSEnumeratePropertyDescriptions(0 /* PDEF_ALL */, ref listIid, out IPropertyDescriptionList list);
            Marshal.ThrowExceptionForHR(hr);
            List<string> names = new();
            try
            {
                Marshal.ThrowExceptionForHR(list.GetCount(out uint count));
                for (uint i = 0; i < count; i++)
                {
                    Guid descriptionIid = IidPropertyDescription;
                    if (list.GetAt(i, ref descriptionIid, out IPropertyDescription description) != 0)
                    {
                        continue;
                    }

                    try
                    {
                        if (description.GetCanonicalName(out IntPtr name) == 0 && name != IntPtr.Zero)
                        {
                            string? text = Marshal.PtrToStringUni(name);
                            Marshal.FreeCoTaskMem(name);
                            if (!string.IsNullOrEmpty(text) && text.All(c => char.IsLetterOrDigit(c) || c == '.' || c == '_'))
                            {
                                names.Add(text);
                            }
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(description);
                    }
                }
            }
            finally
            {
                Marshal.ReleaseComObject(list);
            }

            // System.ItemUrl is selected on its own in every statement; a second copy would be a duplicate column.
            return names
                .Where(n => !string.Equals(n, "System.ItemUrl", StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n, StringComparer.Ordinal)
                .ToList();
        }

        [DllImport("propsys.dll")]
        private static extern int PSEnumeratePropertyDescriptions(
            int filterOn, ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out IPropertyDescriptionList ppv);

        [ComImport]
        [Guid("1F9FC1D0-C39B-4B26-817F-011967D3440E")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPropertyDescriptionList
        {
            [PreserveSig]
            int GetCount(out uint pcElem);

            [PreserveSig]
            int GetAt(uint iElem, ref Guid riid, [MarshalAs(UnmanagedType.Interface)] out IPropertyDescription ppv);
        }

        [ComImport]
        [Guid("6F79D558-3E96-4549-A1D1-7D75D2288814")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IPropertyDescription
        {
            [PreserveSig]
            int GetPropertyKey(IntPtr pkey);

            [PreserveSig]
            int GetCanonicalName(out IntPtr ppszName);
        }
    }
}
