using System.Text.Json;
using OutlookAI.ComHost.Protocol;
using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins how the index's stores are LISTED for the store map (Q92/Q99): one shallow
/// <c>DIRECTORY</c> traversal of the user's MAPI root, which returned exactly one row per store
/// root on the test guest - stores without a single mail item included - where the old 2000-row
/// mail sample saw one store of three. Also the small pieces the store map stands on: the
/// profile-section path the Exchange signature is read from, and the store snapshot carrying the
/// new inputs across the COM host's wire.
/// </summary>
public sealed class StoreRootListingTests
{
    private const string UserRoot = "mapi16://{S-1-5-21-1-2-3-1000}/";

    [Fact]
    public void TheListing_IsOneShallowTraversalOfTheUsersRoot()
    {
        Assert.Equal(
            "SELECT System.ItemUrl FROM SystemIndex WHERE DIRECTORY='" + UserRoot + "'",
            WsSqlBuilder.BuildStoreRootListing(UserRoot));
        Assert.Throws<ArgumentException>(() => WsSqlBuilder.BuildStoreRootListing("file:///C:/"));
    }

    [Fact]
    public void EveryStoreRootIsListed_AndNothingElse()
    {
        var client = new ListingClient(
            UserRoot + "identity@vm.invalid($be889d8b)",
            UserRoot + "Outlook Data File($23a27f0d)",
            UserRoot + "q99lz@vm.invalid($ce9d6e4)",
            UserRoot + "q99 50%25 off%2A%3Fx($5159380d)",
            UserRoot + "Outlook Data File($23a27f0d)",                // the same root twice
            UserRoot + "tier@vm.invalid($93f42b43)/0",                 // a row below a root
            UserRoot + "tier@vm.invalid($93f42b43)/0/Inbox/\uAC00\uAC01",
            "mapi16://{S-1-5-21-9-9-9-1001}/someone else($11111111)",  // another user's root
            "file:///C:/Users/u/notes.txt");

        IReadOnlyList<StoreScopeInfo> roots = new IndexSearchService(client).ListStoreRoots(UserRoot);

        Assert.Equal(
            new[] { "identity@vm.invalid($be889d8b)", "Outlook Data File($23a27f0d)", "q99lz@vm.invalid($ce9d6e4)", "q99 50%25 off%2A%3Fx($5159380d)" },
            roots.Select(r => r.StoreSegment));
        Assert.Equal(new uint?[] { 0xBE889D8Bu, 0x23A27F0Du, 0x0CE9D6E4u, 0x5159380Du }, roots.Select(r => r.StoreHash));
        Assert.Equal("identity@vm.invalid", roots[0].StoreDisplayName);
        Assert.Equal(UserRoot + "identity@vm.invalid($be889d8b)", roots[0].StorePrefix);

        // A name the index percent-encodes is listed under the store's own name, and keeps the
        // index's spelling in its segment and its prefix (McpServer/README.md fact 17).
        Assert.Equal("q99 50% off*?x", roots[3].StoreDisplayName);
        Assert.Equal(UserRoot + "q99 50%25 off%2A%3Fx($5159380d)", roots[3].StorePrefix);

        // ONE statement, however many roots: no probe per root rides along.
        string statement = Assert.Single(client.Statements);
        Assert.StartsWith("SELECT System.ItemUrl FROM SystemIndex WHERE DIRECTORY=", statement, StringComparison.Ordinal);
    }

    [Fact]
    public void ARowBelowARoot_NeverStandsInForTheRoot()
    {
        // Control for the shallow rule: a listing that returned only rows BELOW a root - which a
        // DIRECTORY traversal does not, but a wider statement would - lists no root at all,
        // rather than inventing one from the first segment of a deeper URL.
        var client = new ListingClient(
            UserRoot + "tier@vm.invalid($93f42b43)/0",
            UserRoot + "tier@vm.invalid($93f42b43)/0/Inbox/가각");
        Assert.Empty(new IndexSearchService(client).ListStoreRoots(UserRoot));
    }

    [Fact]
    public void AStoreRootUrlParses_AndAnythingBelowARootDoesNot()
    {
        StoreScopeInfo? root = StoreScopeInfo.FromStorePrefix(UserRoot + "tier@vm.invalid($93f42b43)");
        Assert.NotNull(root);
        Assert.Equal(0x93F42B43u, root!.StoreHash);
        Assert.Null(StoreScopeInfo.FromStorePrefix(UserRoot + "tier@vm.invalid($93f42b43)/0"));
        Assert.Null(StoreScopeInfo.FromStorePrefix("file:///C:/x"));
        Assert.Null(StoreScopeInfo.FromStorePrefix(null));
    }

    [Fact]
    public void TheUsersRoot_IsMapi16WithTheirSid()
    {
        string root = IndexSearchService.CurrentUserMapiRoot();
        Assert.Matches(@"^mapi16://\{S-1-[0-9-]+\}/$", root);
    }

    [Fact]
    public void AProfileSectionPath_CannotReachOutsideTheProfile()
    {
        Assert.Equal(
            @"Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\0a0b0c0d0e0f00112233445566778899",
            OutlookProfileSections.SectionKeyPath(@"Software\Microsoft\Office\16.0\Outlook", "Outlook", "0A0B0C0D0E0F00112233445566778899"));
        Assert.Null(OutlookProfileSections.SectionKeyPath(@"Software\X", @"..\Other", "0A0B0C0D0E0F00112233445566778899"));
        Assert.Null(OutlookProfileSections.SectionKeyPath(@"Software\X", "Outlook", "0A0B0C0D"));
        Assert.Null(OutlookProfileSections.SectionKeyPath(@"Software\X", "Outlook", "0A0B0C0D0E0F0011223344556677889Z"));
        Assert.Null(OutlookProfileSections.SectionKeyPath(@"Software\X", null, "0A0B0C0D0E0F00112233445566778899"));
        Assert.Equal("01020ff8", OutlookProfileSections.MappingSignatureValueName);
    }

    [Fact]
    public void AStoreSnapshot_CarriesTheExchangeHashInputsAcrossTheWire()
    {
        var original = new ComStoreDetail(
            "alice@example.com", "00AA", 0, true, false,
            filePath: @"C:\Users\a\alice.ost", mappingSignatureHex: "0102", exchangeProfileSectionHex: "0A0B0C0D0E0F00112233445566778899");

        string json = JsonSerializer.Serialize(original, ComHostProtocol.Json);
        ComStoreDetail? read = JsonSerializer.Deserialize<ComStoreDetail>(json, ComHostProtocol.Json);

        Assert.NotNull(read);
        Assert.Equal(@"C:\Users\a\alice.ost", read!.FilePath);
        Assert.Equal("0102", read.MappingSignatureHex);
        Assert.Equal("0A0B0C0D0E0F00112233445566778899", read.ExchangeProfileSectionHex);

        // A PST carries none of them, and an older snapshot without the fields still reads.
        ComStoreDetail? pst = JsonSerializer.Deserialize<ComStoreDetail>(
            "{\"displayName\":\"x\",\"storeId\":\"00\",\"exchangeStoreType\":3,\"isCachedExchange\":null}", ComHostProtocol.Json);
        Assert.NotNull(pst);
        Assert.Null(pst!.FilePath);
        Assert.Null(pst.MappingSignatureHex);
    }

    /// <summary>Answers the listing with the given URLs, and refuses every other statement.</summary>
    private sealed class ListingClient : IIndexClient
    {
        private readonly string[] _urls;

        public ListingClient(params string[] urls) => _urls = urls;

        public List<string> Statements { get; } = new();

        public IndexProviderKind Provider => IndexProviderKind.OleDb;

        public IReadOnlyList<IReadOnlyDictionary<string, object?>> ExecuteRows(string sql, int maxRows, int? commandTimeoutSeconds = null)
        {
            Statements.Add(sql);
            if (!sql.Contains("DIRECTORY=", StringComparison.Ordinal))
            {
                throw new InvalidOperationException("The listing stand-in only answers the listing: " + sql);
            }

            return _urls.Select(u => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
            {
                ["System.ItemUrl"] = u,
            }).ToList();
        }
    }
}
