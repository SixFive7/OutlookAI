using OutlookAI.Core.Mapi;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// T2, read-only, cached Exchange only: folders addressed by ID (Q131-Q136) on a real Exchange mailbox,
/// which the change could reach only through T1 and the measurement of 2026-11-17 (Q132) before it ran on
/// the Exchange VM. Every folder of the test account carries a 66-character folder ID of kind X that reads
/// back to an Exchange entry id; find_folder finds the Inbox by its name under the same ID; a search scoped
/// to that ID takes the Exchange folder path and answers; and an Exchange-shaped ID no folder has is
/// refused, never answered as an empty folder.
/// <para>
/// Nothing is created, moved or written. create_folder's refusals, and the refusal while Outlook does not
/// answer, stay with T1: reaching either here would take a write tool or a stopped Outlook.
/// </para>
/// </summary>
[Collection(LiveCollections.Phase2)]
[Trait("Category", "Live")]
public sealed class LiveExchangeFolderIdTests
{
    private readonly LivePhase2Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveExchangeFolderIdTests(LivePhase2Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    [Fact]
    [Trait("Requires", "CachedExchange")]
    [Trait("Writes", "Nothing")]
    public void EveryFolderOfTheTestAccount_HasA66CharacterKindXId_ThatReadsBackToAnExchangeEntryId()
    {
        IReadOnlyList<FolderView> folders = LiveIds.Folders(Service, Hub);
        Assert.NotEmpty(folders);
        foreach (FolderView folder in folders)
        {
            string id = folder.FolderId ?? throw new InvalidOperationException("list_folders shows no folderId for '" + folder.Path + "'.");
            Assert.True(id.Length == 66 && id[0] == 'X', "'" + folder.Path + "' has the folder ID '" + id + "', not 66 characters of kind X.");
            Assert.True(OutlookIds.TryDecodeFolderId(id, "folder_id", out string entryIdHex, out string? problem), problem);
            Assert.True(entryIdHex.Length == 92, "'" + folder.Path + "' reads back to " + (entryIdHex.Length / 2) + " bytes, not an Exchange folder's 46.");
        }

        _output.WriteLine("list_folders: " + folders.Count + " folders of the test account, every ID 66 characters of kind X, each reading back to a 46-byte entry id");
    }

    [Fact]
    [Trait("Requires", "CachedExchange")]
    [Trait("Writes", "Nothing")]
    public void EveryStoreOfTheProfile_HasThe28CharacterStoreId_AnExchangeStoreOutlookDoesNotCacheToo()
    {
        // Option (b) of the uncached Exchange store's ID (2026-11-18): every store's top folder is read, so every
        // store gets S and the provider UID, 28 characters - the online entry Exchange automapped included, which
        // used to get T and its whole StoreID. That entry is the test account's mailbox seen twice, so its ID is the
        // account's own; the mailbox is listed and resolved once, through the cached entry.
        IReadOnlyList<StoreView> stores = Service.ListAccounts().Stores;
        Assert.NotEmpty(stores);
        foreach (StoreView store in stores)
        {
            string id = store.StoreId ?? throw new InvalidOperationException("list_accounts shows no storeId for a store.");
            Assert.True(id.Length == 28 && id[0] == 'S', "a store has the ID '" + id + "', not 28 characters of kind S.");
        }

        string hubId = LiveIds.StoreId(Service, Hub);
        int sharing = stores.Count(s => s.StoreId == hubId);
        Assert.Single(Service.ListFolders(hubId).Stores);
        _output.WriteLine("list_accounts: " + stores.Count + " store(s), every ID 28 characters of kind S; the test account's ID is shown by "
            + sharing + " of them, and list_folders for it walks one store");
    }

    [Fact]
    [Trait("Requires", "CachedExchange")]
    [Trait("Writes", "Nothing")]
    public void FindFolder_FindsTheInboxByItsName_UnderTheIdListFoldersGives()
    {
        string listed = LiveIds.FolderId(Service, Hub, "Inbox");
        string found = LiveIds.FolderIdByName(Service, Hub, "Inbox");
        Assert.Equal(listed, found);
        _output.WriteLine("find_folder: the test account's Inbox by name, the same 66-character ID list_folders gives");
    }

    [Fact]
    [Trait("Requires", "CachedExchange")]
    [Trait("Writes", "Nothing")]
    public void ASearchScopedToAnExchangeFolderId_Answers_AndAnIdNoFolderHasIsRefused()
    {
        Service.ClearSweepCache();
        string inbox = LiveIds.FolderId(Service, Hub, "Inbox");
        SearchOutcome scoped = Service.Search(new SearchRequest
        {
            Query = "oaiexchangefolderprobe",
            FolderId = inbox,
            Top = 5,
            SnippetChars = 0,
        });
        Assert.NotNull(scoped);

        ArgumentException unknown = Assert.Throws<ArgumentException>(() => Service.Search(new SearchRequest
        {
            Query = "oaiexchangefolderprobe",
            FolderId = LiveIds.UnknownFolderIdIn(Service, Hub),
            Top = 5,
            SnippetChars = 0,
        }));
        Assert.Contains("No folder in this Outlook profile has the ID", unknown.Message, StringComparison.Ordinal);
        _output.WriteLine("search: scoped to the Inbox's Exchange folder ID it answered (" + scoped.Hits.Count
            + " hit(s) for a term nothing carries); an Exchange-shaped ID no folder has was refused");
    }
}
