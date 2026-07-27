using System.Collections.Concurrent;
using System.Reflection;
using System.Text.Json;
using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using Xunit;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Machine-local settings for the T2 live tier, loaded from the gitignored
/// live-fixtures/ folder: account/store identifiers must never be committed to this
/// PUBLIC repo (v3.MD S6/D13). The file is created once on the dev machine.
/// </summary>
public sealed class LiveTestSettings
{
    /// <summary>Display name of the designated test-hub store (v3.MD S2/D14).</summary>
    public string TestHubStoreDisplayName { get; set; } = string.Empty;

    /// <summary>Display names of the three primary account stores.</summary>
    public List<string> ExpectedStoreDisplayNames { get; set; } = new();

    /// <summary>Display names of the delegate/shared-mailbox cache stores (Phase-2 list_accounts exactness).</summary>
    public List<string> ExpectedDelegateStoreDisplayNames { get; set; } = new();

    /// <summary>The section-5 probe term (generic word; proven to hit on this machine).</summary>
    public string ProbeTerm { get; set; } = string.Empty;

    /// <summary>The SF-6 / D40 recall-regression population (see <see cref="SubjectOnlyProbeSettings"/>).</summary>
    public SubjectOnlyProbeSettings? SubjectOnlyProbe { get; set; }

    /// <summary>
    /// OPTIONAL coordinates of a delegate folder that is NESTED in Outlook but FLAT in the
    /// index (see <see cref="DelegateNestedFolderProbeSettings"/>). Absent = the locator
    /// test proves only the no-false-positive half.
    /// </summary>
    public DelegateNestedFolderProbeSettings? DelegateNestedFolderProbe { get; set; }

    /// <summary>Loads the settings file or throws with setup instructions.</summary>
    public static LiveTestSettings Load()
    {
        string testProjectDir =
            typeof(LiveTestSettings).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");

        string path = Path.Combine(testProjectDir, "live-fixtures", "live-test-settings.json");
        if (!File.Exists(path))
        {
            throw new InvalidOperationException(
                $"Live-test settings not found at '{path}'. The T2 live tier only runs on the dev machine; "
                + "create the gitignored file with testHubStoreDisplayName, expectedStoreDisplayNames and probeTerm "
                + "(account identifiers are never committed - v3.MD S6).");
        }

        LiveTestSettings settings = JsonSerializer.Deserialize<LiveTestSettings>(
                File.ReadAllText(path),
                new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
            ?? throw new InvalidOperationException("Live-test settings file deserialized to null.");

        if (string.IsNullOrWhiteSpace(settings.TestHubStoreDisplayName)
            || settings.ExpectedStoreDisplayNames.Count == 0
            || string.IsNullOrWhiteSpace(settings.ProbeTerm))
        {
            throw new InvalidOperationException("Live-test settings file is incomplete.");
        }

        if (settings.SubjectOnlyProbe == null
            || string.IsNullOrWhiteSpace(settings.SubjectOnlyProbe.StoreDisplayName)
            || string.IsNullOrWhiteSpace(settings.SubjectOnlyProbe.FolderPath)
            || string.IsNullOrWhiteSpace(settings.SubjectOnlyProbe.SubjectTerm)
            || string.IsNullOrWhiteSpace(settings.SubjectOnlyProbe.SenderFragment))
        {
            throw new InvalidOperationException(
                "Live-test settings are missing the 'subjectOnlyProbe' block (storeDisplayName, folderPath, "
                + "subjectTerm, senderFragment) required by the D40/SF-6 recall regression.");
        }

        return settings;
    }
}

/// <summary>
/// Coordinates of the SF-6 discovery-case population: a real, stable mail population in
/// a business store whose probe term appears in the SUBJECT (and sender address) but NOT
/// in the body stream - exactly the shape the pre-D40 unqualified
/// <c>CONTAINS('term')</c> predicate could never match. Kept in the gitignored settings
/// file because the store/folder/term triple identifies real mailbox content (S6/D13);
/// the tests only ever print counts and timings (S4).
/// </summary>
public sealed class SubjectOnlyProbeSettings
{
    /// <summary>Store display name holding the population.</summary>
    public string StoreDisplayName { get; set; } = string.Empty;

    /// <summary>Store-relative folder path holding the population.</summary>
    public string FolderPath { get; set; } = string.Empty;

    /// <summary>A term present in every member's subject and in no member's body.</summary>
    public string SubjectTerm { get; set; } = string.Empty;

    /// <summary>Sender-address fragment selecting exactly the same population (the independent expectation).</summary>
    public string SenderFragment { get; set; } = string.Empty;
}

/// <summary>
/// Coordinates of a delegate folder the index publishes FLAT while Outlook nests it - the
/// shape that made every delegate hit in a subfolder unopenable until soak fix 16.
/// Gitignored like every other machine coordinate (S6); the tests print counts, paths of
/// their own making and booleans only (S4) and write nothing.
/// </summary>
public sealed class DelegateNestedFolderProbeSettings
{
    /// <summary>Delegate store display name.</summary>
    public string StoreDisplayName { get; set; } = string.Empty;

    /// <summary>The flat leaf name the index serves (and the COM folder's own name).</summary>
    public string FolderName { get; set; } = string.Empty;

    /// <summary>The COM parent path the index drops (proves the folder really is nested).</summary>
    public string ComParentPath { get; set; } = string.Empty;
}

/// <summary>
/// Shared state for the T2 live tier: one index service (provider auto-selected), one
/// store-scope discovery, one lazy Outlook COM session (may START Outlook per S7/D17 -
/// never stops it), one lazy full walk of the tiny test-hub store, and the store-UID to
/// StoreID mapping learned from successful verify-on-open calls.
/// </summary>
public sealed class LivePhase1Fixture : IDisposable
{
    private readonly Lazy<OutlookComSession> _session;
    private readonly Lazy<IReadOnlyList<ComStoreInfo>> _comStores;
    private readonly Lazy<IReadOnlyList<ComWalkedItem>> _testHubWalk;

    public LivePhase1Fixture()
    {
        Settings = LiveTestSettings.Load();

        // Fail-closed per-store count tripwire: no census, no live tier. Cheap after
        // the first fixture (one process-wide baseline).
        LiveStoreCountTripwire.EnsureBaseline(Settings);
        Service = IndexSearchService.CreateDefault(out string providerReport);
        ProviderReport = providerReport;

        // Store discovery: broad sample first, then targeted per-address discovery for
        // stores the sample misses (Phase-1 finding: an unordered 30k sample never
        // surfaced the tiny idle store, and SCOPE needs the exact ($hash) segment).
        List<StoreScopeInfo> scopes = Service.DiscoverStoreScopes(2000).ToList();
        foreach (string expected in Settings.ExpectedStoreDisplayNames)
        {
            if (scopes.Any(s => string.Equals(s.StoreDisplayName, expected, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            StoreScopeInfo? targeted = Service.TryDiscoverStoreScopeByAddress(expected);
            if (targeted != null)
            {
                scopes.Add(targeted);
            }
        }

        StoreScopes = scopes;

        _session = new Lazy<OutlookComSession>(
            () => OutlookComSession.Connect(allowStartingOutlook: true),
            LazyThreadSafetyMode.ExecutionAndPublication);
        _comStores = new Lazy<IReadOnlyList<ComStoreInfo>>(
            () => Session.GetStores(),
            LazyThreadSafetyMode.ExecutionAndPublication);
        _testHubWalk = new Lazy<IReadOnlyList<ComWalkedItem>>(
            () => Session.WalkStoreMailItems(Settings.TestHubStoreDisplayName),
            LazyThreadSafetyMode.ExecutionAndPublication);
    }

    public LiveTestSettings Settings { get; }

    public IndexSearchService Service { get; }

    public string ProviderReport { get; }

    public IReadOnlyList<StoreScopeInfo> StoreScopes { get; }

    /// <summary>Store-UID hex -> StoreID learned from successful decode-verifies.</summary>
    public ConcurrentDictionary<string, string> UidToStoreId { get; } = new(StringComparer.OrdinalIgnoreCase);

    public OutlookComSession Session => _session.Value;

    public IReadOnlyList<ComStoreInfo> ComStores => _comStores.Value;

    /// <summary>Full COM walk of the test-hub store (ground truth for the oracle).</summary>
    public IReadOnlyList<ComWalkedItem> TestHubWalk => _testHubWalk.Value;

    public StoreScopeInfo GetScope(string storeDisplayName)
    {
        return StoreScopes.FirstOrDefault(s =>
                string.Equals(s.StoreDisplayName, storeDisplayName, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException(
                $"Store '{storeDisplayName}' not found among {StoreScopes.Count} discovered index scopes.");
    }

    public string GetComStoreId(string storeDisplayName)
    {
        return ComStores.FirstOrDefault(s =>
                string.Equals(s.DisplayName, storeDisplayName, StringComparison.OrdinalIgnoreCase))?.StoreId
            ?? throw new InvalidOperationException(
                $"Store '{storeDisplayName}' not found among {ComStores.Count} COM stores.");
    }

    public void Dispose()
    {
        if (_session.IsValueCreated)
        {
            // Releases COM references only - Outlook keeps running (S7: never kill/close).
            _session.Value.Dispose();
        }
    }
}

[CollectionDefinition("LivePhase1")]
public sealed class LivePhase1Collection : ICollectionFixture<LivePhase1Fixture>
{
}
