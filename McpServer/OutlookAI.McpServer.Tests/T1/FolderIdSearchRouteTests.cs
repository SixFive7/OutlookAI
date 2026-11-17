using System.Reflection;
using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// How a search finds the folder its folder_id names (Q132 (i)): a local mail file's folder by its OWN
/// index row - the store's root by the provider UID its items carry, the folder's row by the node id
/// the ID carries - with no Outlook call at all; an Exchange folder only through Outlook, refused with
/// a clear message while Outlook is not answering; and a folder_id and store_id that contradict each
/// other refused before anything is searched. Every id and URL here is synthetic.
/// <para>
/// Marked for the machine's zone (Q144) as a class: every test here goes through
/// <c>MailService.Search</c>, whose staleness report converts a time with <c>DateTime.ToUniversalTime</c>,
/// and <c>Search</c> takes no zone a test could set - so these run a second time in W. Europe.
/// </para>
/// </summary>
[Trait("TimeZone", "Machine")]
public sealed class FolderIdSearchRouteTests
{
    private const string Store = "tier@vm.invalid";

    private static readonly string UserRoot = IndexSearchService.CurrentUserMapiRoot();

    private static readonly string StorePrefix = UserRoot + Store + "($1234abcd)";

    [Fact]
    public void ALocalMailFilesFolder_IsFoundByItsOwnIndexRow_WithNoOutlookCallAtAll()
    {
        string folderEntryId = TestFolders.FolderEntryId(Store, "Projects/2026");
        string folderUrl = StorePrefix + "/0/Projects/2026";
        RecordingIndex index = new RecordingIndex(folderEntryId, folderUrl);
        using MailService service = new MailService(new DirectGateway(RefusingSession.Create()), null, index);

        SearchOutcome outcome = service.Search(new SearchRequest
        {
            Query = "budget",
            FolderId = OutlookIds.FolderIdFromEntryId(folderEntryId),
            IndexOnly = true,
            Top = 5,
            SnippetChars = 0,
        });

        Assert.Equal("index", outcome.Scope!.FolderFoundVia);
        Assert.Equal("Projects/2026", outcome.Scope.Folder);
        Assert.Equal(TestFolders.StoreId(Store), outcome.Scope.StoreId);
        Assert.Equal(OutlookIds.FolderIdFromEntryId(folderEntryId), outcome.Scope.FolderId);
        Assert.Contains(index.Statements, sql => sql.Contains("System.ProviderItemID='" + OutlookIds.PstProviderItemId(Convert.FromHexString(folderEntryId)) + "'", StringComparison.Ordinal));
        Assert.Contains(index.Statements, sql => sql.Contains("CONTAINS", StringComparison.OrdinalIgnoreCase)
            && sql.Contains("SCOPE='" + folderUrl + "'", StringComparison.Ordinal));
    }

    [Fact]
    public void AnExchangeFolder_WhileOutlookIsNotAnswering_IsRefused_SayingWhy()
    {
        byte[] exchange = new byte[46];
        Array.Copy(TestFolders.Uid(Store), 0, exchange, 4, 16);
        exchange[20] = 1;
        exchange[43] = 9;
        RecordingIndex index = new RecordingIndex(null, null);
        using MailService service = new MailService(new DirectGateway(RefusingSession.Create()), null, index);

        ArgumentException refused = Assert.Throws<ArgumentException>(() => service.Search(new SearchRequest
        {
            Query = "budget",
            FolderId = OutlookIds.FolderIdFromEntryId(exchange),
            IndexOnly = true,
        }));

        Assert.Contains("needs Outlook", refused.Message, StringComparison.Ordinal);
        Assert.Contains("store_id", refused.Message, StringComparison.Ordinal);
        Assert.DoesNotContain(index.Statements, sql => sql.Contains("ProviderItemID", StringComparison.Ordinal));
    }

    [Fact]
    public void ALocalFolderTheIndexHasNoRowFor_WhileOutlookIsNotAnswering_IsRefused_NeverGuessed()
    {
        string folderEntryId = TestFolders.FolderEntryId(Store, "Brand new");
        RecordingIndex index = new RecordingIndex(folderEntryId, null);
        using MailService service = new MailService(new DirectGateway(RefusingSession.Create()), null, index);

        ArgumentException refused = Assert.Throws<ArgumentException>(() => service.Search(new SearchRequest
        {
            Query = "budget",
            FolderId = OutlookIds.FolderIdFromEntryId(folderEntryId),
            IndexOnly = true,
        }));

        Assert.Contains("no row for the folder itself", refused.Message, StringComparison.Ordinal);
        Assert.Contains("Outlook is not answering", refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AFolderIdAndAStoreIdOfAnotherStore_AreRefusedBeforeAnythingIsSearched()
    {
        RecordingIndex index = new RecordingIndex(null, null);
        using MailService service = new MailService(new DirectGateway(RefusingSession.Create()), null, index);

        ArgumentException refused = Assert.Throws<ArgumentException>(() => service.Search(new SearchRequest
        {
            Query = "budget",
            FolderId = TestFolders.FolderId(Store, "Inbox"),
            StoreId = TestFolders.StoreId("someone-else@vm.invalid"),
        }));

        Assert.Contains("another store", refused.Message, StringComparison.Ordinal);
        Assert.Empty(index.Statements);
    }

    // ------------------------------------------------------------------ fixtures

    /// <summary>The index: one store root, an item under it carrying the store's UID, and the folder's own row.</summary>
    private sealed class RecordingIndex : IIndexClient
    {
        private readonly string? _folderEntryId;
        private readonly string? _folderUrl;

        internal RecordingIndex(string? folderEntryId, string? folderUrl)
        {
            _folderEntryId = folderEntryId;
            _folderUrl = folderUrl;
        }

        internal List<string> Statements { get; } = new();

        public IndexProviderKind Provider => IndexProviderKind.OleDb;

        public IReadOnlyList<IReadOnlyDictionary<string, object?>> ExecuteRows(string sql, int maxRows, int? commandTimeoutSeconds = null)
        {
            Statements.Add(sql);
            if (sql.Contains("DIRECTORY='" + UserRoot + "'", StringComparison.Ordinal))
            {
                return Rows(StorePrefix);
            }

            if (sql.Contains("System.ProviderItemID=", StringComparison.Ordinal))
            {
                return _folderUrl == null ? Rows() : Rows(_folderUrl);
            }

            if (_folderEntryId != null && sql.StartsWith("SELECT TOP 200 System.ItemUrl FROM SystemIndex WHERE SCOPE='" + StorePrefix + "/'", StringComparison.Ordinal))
            {
                byte[] item = Convert.FromHexString(_folderEntryId);
                BitConverter.GetBytes(0x00200024u).CopyTo(item, 20); // a message node id in the same store
                return Rows(StorePrefix + "/0/Inbox/" + EntryIdCodec.EncodeBytes(item));
            }

            return Rows();
        }

        private static IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows(params string[] urls)
        {
            return urls.Select(u => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
            {
                ["System.ItemUrl"] = u,
            }).ToList();
        }
    }

    /// <summary>An Outlook that answers nothing - every call throws, as one that is not running does.</summary>
    private class RefusingSession : DispatchProxy
    {
        internal static IOutlookSession Create()
        {
            return Create<IOutlookSession, RefusingSession>()
                ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
        }

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            throw new TimeoutException("Outlook is not running; it is being started.");
        }
    }

    private sealed class DirectGateway : IComGateway
    {
        private readonly IOutlookSession _session;

        internal DirectGateway(IOutlookSession session)
        {
            _session = session;
        }

        public event Action? OutlookGone
        {
            add { }
            remove { }
        }

        public bool IsConnected => false;

        public bool? QuitSinkActive => null;

        public bool ProbeConnected() => false;

        public T Run<T>(Func<IOutlookSession, T> operation) => operation(_session);

        public T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery) => operation(_session);

        public T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false)
            => operation(_session);

        public ComHostDiagnostics GetDiagnostics() => new ComHostDiagnostics("in-process", "ready");

        public void Dispose()
        {
        }
    }
}
