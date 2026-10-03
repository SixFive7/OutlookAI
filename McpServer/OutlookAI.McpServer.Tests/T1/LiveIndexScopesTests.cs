using OutlookAI.Core.IndexSearch;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins how the live tier finds an indexed store's scope (<see cref="LiveIndexScopes"/>): the
/// 2000-row sample, then mail addressed to an address-named store, then the index's store-root
/// listing. The first live run with the 160,000-item Corpus A (OutlookAI-Indexed, 2026-10-03)
/// sampled another profile's store 1,941 times and Corpus A not once, and 13 tests failed before
/// measuring anything; the listing is what finds a store no sample reaches and no address names.
/// </summary>
public sealed class LiveIndexScopesTests
{
    private const string UserRoot = "mapi16://{S-1-5-21-1-2-3-1000}/";
    private const string CorpusA = UserRoot + "Corpus A($996dc7a9)";
    private const string OtherProfilesStore = UserRoot + "Outlook Data File($23a27f0d)";
    private const string Hub = UserRoot + "tier@vm.invalid($93f42b43)";

    [Fact]
    public void AStoreTheSampleHolds_IsTakenFromTheSample_WithoutTheListing()
    {
        var client = new IndexStandIn(
            sample: new[] { Item(CorpusA), Item(CorpusA), Item(Hub) },
            roots: new[] { CorpusA, Hub },
            withContent: new[] { CorpusA, Hub });

        StoreScopeInfo? scope = LiveIndexScopes.Find(new IndexSearchService(client), "Corpus A", userRoot: UserRoot);

        Assert.Equal(CorpusA, scope?.StorePrefix);
        Assert.DoesNotContain(client.Statements, s => s.Contains("DIRECTORY=", StringComparison.Ordinal));
    }

    [Fact]
    public void AStoreNoAddressNamesAndTheSampleMissed_IsFoundByTheRootListing()
    {
        // The guest-one shape: the sample is all another profile's store and the hub.
        var client = new IndexStandIn(
            sample: new[] { Item(OtherProfilesStore), Item(OtherProfilesStore), Item(OtherProfilesStore), Item(Hub) },
            roots: new[] { OtherProfilesStore, Hub, CorpusA },
            withContent: new[] { OtherProfilesStore, Hub, CorpusA });

        StoreScopeInfo? scope = LiveIndexScopes.Find(new IndexSearchService(client), "corpus a", userRoot: UserRoot);

        Assert.NotNull(scope);
        Assert.Equal(CorpusA, scope!.StorePrefix);
        Assert.Equal("Corpus A", scope.StoreDisplayName);
    }

    [Fact]
    public void AnAddressNamedStore_IsAskedByAddressBeforeTheListing()
    {
        string bystander = UserRoot + "bystander@vm.invalid($0badf00d)";
        var client = new IndexStandIn(
            sample: new[] { Item(OtherProfilesStore) },
            roots: new[] { OtherProfilesStore, bystander },
            withContent: new[] { OtherProfilesStore, bystander });

        StoreScopeInfo? scope = LiveIndexScopes.Find(new IndexSearchService(client), "bystander@vm.invalid", userRoot: UserRoot);

        Assert.Equal(bystander, scope?.StorePrefix);
        int byAddress = client.Statements.FindIndex(s => s.Contains("bystander@vm.invalid", StringComparison.OrdinalIgnoreCase)
            && !s.Contains("DIRECTORY=", StringComparison.Ordinal) && !s.Contains("SCOPE=", StringComparison.Ordinal));
        int listing = client.Statements.FindIndex(s => s.Contains("DIRECTORY=", StringComparison.Ordinal));
        Assert.True(byAddress >= 0, "the address was never asked");
        Assert.True(listing > byAddress, "the listing was read before the address was asked");
    }

    [Fact]
    public void TwoRootsOfOneName_AreNeverGuessedBetween()
    {
        var client = new IndexStandIn(
            sample: new[] { Item(Hub) },
            roots: new[] { Hub, OtherProfilesStore, UserRoot + "Outlook Data File($11111111)" },
            withContent: new[] { Hub, OtherProfilesStore, UserRoot + "Outlook Data File($11111111)" });

        Assert.Null(LiveIndexScopes.Find(new IndexSearchService(client), "Outlook Data File", userRoot: UserRoot));
    }

    [Fact]
    public void ARootWithNothingIndexedBelowIt_DoesNotCount()
    {
        // A catalog reset or a rename can bring the root back before anything below it.
        var client = new IndexStandIn(
            sample: new[] { Item(Hub) },
            roots: new[] { Hub, CorpusA },
            withContent: new[] { Hub });

        Assert.Null(LiveIndexScopes.Find(new IndexSearchService(client), "Corpus A", userRoot: UserRoot));
    }

    [Fact]
    public void AStoreTheIndexDoesNotHold_IsNull()
    {
        var client = new IndexStandIn(
            sample: new[] { Item(Hub) },
            roots: new[] { Hub },
            withContent: new[] { Hub });

        Assert.Null(LiveIndexScopes.Find(new IndexSearchService(client), "Corpus A", userRoot: UserRoot));
    }

    private static string Item(string storePrefix) => storePrefix + "/0/Inbox/가각";

    /// <summary>
    /// Answers the discovery sample, the root listing and SCOPE probes; every other statement -
    /// the address probes - finds nothing.
    /// </summary>
    private sealed class IndexStandIn : IIndexClient
    {
        private readonly string[] _sample;
        private readonly string[] _roots;
        private readonly HashSet<string> _withContent;

        public IndexStandIn(string[] sample, string[] roots, string[] withContent)
        {
            _sample = sample;
            _roots = roots;
            _withContent = new HashSet<string>(withContent, StringComparer.OrdinalIgnoreCase);
        }

        public List<string> Statements { get; } = new();

        public IndexProviderKind Provider => IndexProviderKind.OleDb;

        public IReadOnlyList<IReadOnlyDictionary<string, object?>> ExecuteRows(string sql, int maxRows, int? commandTimeoutSeconds = null)
        {
            Statements.Add(sql);
            if (sql.Contains("DIRECTORY=", StringComparison.Ordinal))
            {
                return Rows(_roots);
            }

            if (sql.Contains("System.Kind='email'", StringComparison.Ordinal) && !sql.Contains("CONTAINS", StringComparison.OrdinalIgnoreCase))
            {
                return Rows(_sample);
            }

            const string scopeMarker = "WHERE SCOPE='";
            int at = sql.IndexOf(scopeMarker, StringComparison.Ordinal);
            if (sql.StartsWith("SELECT TOP 1 ", StringComparison.Ordinal) && at >= 0)
            {
                string scope = sql.Substring(at + scopeMarker.Length).TrimEnd('\'');
                return _withContent.Contains(scope) ? Rows(new[] { Item(scope) }) : Rows(Array.Empty<string>());
            }

            return Rows(Array.Empty<string>());
        }

        private static IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows(IEnumerable<string> urls) =>
            urls.Select(u => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
            {
                ["System.ItemUrl"] = u,
            }).ToList();
    }
}
