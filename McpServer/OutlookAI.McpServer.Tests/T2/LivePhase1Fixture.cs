using System.Collections.Concurrent;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using Xunit;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// What KIND of machine the live tier is running on. Declared by the settings file rather
/// than sniffed, because every way of guessing it is a way of guessing it wrong: a profile
/// that happens to have no delegate store today is not thereby a test machine.
/// </summary>
public enum LiveMachineProfile
{
    /// <summary>
    /// A real working profile - mail accounts, delegate/shared mailboxes, a populated
    /// Windows Search index. The default, so a settings file written before this field
    /// existed keeps the strict validation it was written under.
    /// </summary>
    Production = 0,

    /// <summary>
    /// A dedicated test machine: PST stores only, no mail accounts, no delegate mailboxes,
    /// nothing in the local search index. Tests that need any of those name it under
    /// <c>Requires</c> and must be filtered out on that; this value does not make them pass,
    /// it makes the settings file honest about what the machine can offer.
    /// </summary>
    Portable = 1,
}

/// <summary>
/// Machine-local settings for the T2 live tier, loaded from the gitignored
/// live-fixtures/ folder: account/store identifiers must never be committed to this
/// PUBLIC repo (v3.MD S6/D13). The file is created once per machine - see
/// <c>Docs/live-tier-on-the-vm.md</c> for the two shapes it can take.
/// </summary>
public sealed class LiveTestSettings
{
    /// <summary>
    /// What this machine is. Drives which of the blocks below are mandatory: a test machine
    /// with no accounts cannot supply a real index probe term or a subject-only population,
    /// and demanding them would only get them invented.
    /// </summary>
    public LiveMachineProfile MachineProfile { get; set; } = LiveMachineProfile.Production;

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

    /// <summary>
    /// OPTIONAL coordinates of the synthetic measurement corpus, when this machine has one.
    /// Present means the live tier proves the corpus can still answer the questions it exists
    /// for BEFORE it runs anything against it - see <see cref="LiveCorpusFreshness"/>.
    /// </summary>
    public CorpusSettings? Corpus { get; set; }

    /// <summary>
    /// OPTIONAL coordinates of the local mail sink, when this machine has no real transport.
    /// Absent means the profile really can send and receive - see <see cref="LiveMailSink"/>.
    /// </summary>
    public MailSinkSettings? MailSink { get; set; }

    /// <summary>
    /// How the settings file is read.
    /// <para>
    /// <b><see cref="JsonStringEnumConverter"/> is load-bearing, not tidiness.</b> Without it
    /// System.Text.Json accepts <c>machineProfile</c> only as a NUMBER, and the documented
    /// example - and every settings file a person would write - spells it
    /// <c>"Portable"</c>. That threw inside <see cref="Load"/>, before any test ran, so a
    /// correctly written settings file made the entire live tier refuse to start with an
    /// error about JSON rather than about the machine. Both spellings are accepted now, and
    /// the numeric one still works, so no existing file breaks.
    /// </para>
    /// </summary>
    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = true,
        Converters = { new JsonStringEnumConverter() },
    };

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

        return Parse(File.ReadAllText(path));
    }

    /// <summary>
    /// Reads and validates settings from JSON text. Split out from <see cref="Load"/> so the
    /// deserialization itself is pinned: every existing test builds the object in code, so
    /// nothing exercised the JSON path, and that is how a converter the documented example
    /// depends on went missing without a single test noticing.
    /// </summary>
    internal static LiveTestSettings Parse(string json)
    {
        LiveTestSettings settings = JsonSerializer.Deserialize<LiveTestSettings>(json, JsonOptions)
            ?? throw new InvalidOperationException("Live-test settings file deserialized to null.");

        Validate(settings);
        return settings;
    }

    /// <summary>
    /// The completeness rules, split by machine profile and separated from the file read so
    /// CI pins them without a settings file existing at all.
    /// <para>
    /// Two things are required everywhere, because without them nothing can run and nothing
    /// can be protected: the hub store to write in, and the list of stores the count tripwire
    /// watches. The index probe term and the subject-only population are required on a
    /// PRODUCTION profile only. They name real mail that a test machine does not have, and a
    /// blanket requirement does not conjure it - it just gets a plausible-looking value typed
    /// into the file, which is worse than an absent one because the tests that read it then
    /// fail somewhere far away from the mistake.
    /// </para>
    /// </summary>
    internal static void Validate(LiveTestSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        if (string.IsNullOrWhiteSpace(settings.TestHubStoreDisplayName)
            || settings.ExpectedStoreDisplayNames.Count == 0)
        {
            throw new InvalidOperationException(
                "Live-test settings file is incomplete: testHubStoreDisplayName and at least one entry in "
                + "expectedStoreDisplayNames are required on every machine (the hub is where the suite may "
                + "write; expectedStoreDisplayNames is what the count tripwire watches).");
        }

        // A partial block is a mistake on any machine: it reads as configured and behaves as
        // absent. Absent is allowed on a Portable machine; half-written never is.
        if (settings.SubjectOnlyProbe != null && !settings.SubjectOnlyProbe.IsComplete)
        {
            throw new InvalidOperationException(
                "Live-test settings have a partially filled 'subjectOnlyProbe' block. All four of "
                + "storeDisplayName, folderPath, subjectTerm and senderFragment are needed, or leave the "
                + "block out entirely.");
        }

        if (settings.MailSink != null && !settings.MailSink.IsComplete)
        {
            throw new InvalidOperationException(
                "Live-test settings have a partially filled 'mailSink' block. Both a submission host/port and a "
                + "retrieval host/port are needed - a sink that accepts mail and cannot hand it back leaves every "
                + "send in the Outbox, which is exactly what the block exists to prevent.");
        }

        if (settings.Corpus != null && !settings.Corpus.IsComplete)
        {
            throw new InvalidOperationException(
                "Live-test settings have a partially filled 'corpus' block. All of storeDisplayName, "
                + "manifestPath, corpusId, seed, anchorUtc and itemCount are needed, or leave the block out "
                + "entirely. A half-written block reads as configured and behaves as absent, which is exactly "
                + "the silence the freshness check exists to remove.");
        }

        if (settings.MachineProfile != LiveMachineProfile.Production)
        {
            return;
        }

        if (string.IsNullOrWhiteSpace(settings.ProbeTerm))
        {
            throw new InvalidOperationException(
                "Live-test settings are missing 'probeTerm' (a word proven to hit this machine's search "
                + "index). Required on a Production profile; set machineProfile to 'Portable' on a test "
                + "machine that has no index.");
        }

        if (settings.SubjectOnlyProbe == null)
        {
            throw new InvalidOperationException(
                "Live-test settings are missing the 'subjectOnlyProbe' block (storeDisplayName, folderPath, "
                + "subjectTerm, senderFragment) required by the D40/SF-6 recall regression. Required on a "
                + "Production profile; set machineProfile to 'Portable' on a test machine that has no such "
                + "population.");
        }
    }

    /// <summary>
    /// Refuses, on a Production profile, to let a test report success having proven nothing.
    /// <para>
    /// Several live tests discover their own population and return early when it is absent -
    /// a delegate folder that is nested in Outlook and flat in the index, a hub account row
    /// in the signature registry. On the machine those tests were written for, absent means
    /// something is wrong with the machine, and returning green hides it. On a Portable
    /// machine absent is simply the truth, and the test should not have been selected: it names
    /// what it needs under <c>Requires</c>. So this throws on the first and no-ops on the second.
    /// </para>
    /// </summary>
    /// <param name="what">The population that was not found, named as a reader would name it.</param>
    public void RequireProductionPopulation(string what)
    {
        if (MachineProfile != LiveMachineProfile.Production)
        {
            return;
        }

        throw new InvalidOperationException(
            "This machine declares machineProfile 'Production', where " + what + " is expected to exist. "
            + "It was not found, so this test can prove nothing and refuses to report success. Either the "
            + "machine or the live-test settings have drifted; a machine that genuinely lacks it should "
            + "declare machineProfile 'Portable' and filter out the tests whose Requires names "
            + "what it does not have.");
    }

    /// <summary>One line naming what this machine claims to be, printed at the start of a live run.</summary>
    public string Describe()
    {
        return "machineProfile=" + MachineProfile
            + ", stores=" + ExpectedStoreDisplayNames.Count
            + ", delegateStores=" + ExpectedDelegateStoreDisplayNames.Count
            + ", probeTerm=" + (string.IsNullOrWhiteSpace(ProbeTerm) ? "none" : "set")
            + ", subjectOnlyProbe=" + (SubjectOnlyProbe == null ? "none" : "set")
            + ", corpus=" + (Corpus == null ? "none" : Corpus.CorpusId)
            + ", mailSink=" + (MailSink == null
                ? "none (real transport)"
                : MailSink.SubmitHost + ":" + MailSink.SubmitPort);
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

    /// <summary>True when all four coordinates are present. A block with three of them is a typo, not a choice.</summary>
    public bool IsComplete =>
        !string.IsNullOrWhiteSpace(StoreDisplayName)
        && !string.IsNullOrWhiteSpace(FolderPath)
        && !string.IsNullOrWhiteSpace(SubjectTerm)
        && !string.IsNullOrWhiteSpace(SenderFragment);
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

        // Fail-closed corpus freshness: a corpus whose measurement windows have emptied is
        // worse than no corpus, because every test asking about those windows still passes.
        // Checked here rather than per test, and BEFORE anything reads the store, for the
        // same reason as the tripwire above: it is a property of the machine, and a machine
        // that cannot answer must not be measured.
        LiveCorpusFreshness.EnsureFresh(Settings);

        // And, on a machine whose transport is a local sink, that the sink is answering
        // before anything is sent into it. TCP only - it must be answerable before Outlook
        // has been started, because a sink that is down looks exactly like a code fault once
        // the mail is sitting in the Outbox.
        LiveMailSink.EnsureReachable(Settings);
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
            () =>
            {
                OutlookComSession connected = OutlookComSession.Connect(allowStartingOutlook: true);

                // The first moment COM is available is the first moment this can be asked,
                // and it has to be asked BEFORE anything sends: mail left queued by an
                // earlier run is indistinguishable, at teardown, from mail this run failed
                // to clean up. Checked here rather than in the constructor so it does not
                // force Outlook to start before a test that needs no Outlook at all.
                LiveMailSink.EnsureOutboxDrained(connected, Settings);
                return connected;
            },
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
        try
        {
            DisposeCore();
        }
        finally
        {
            // Outside the teardown on purpose, exactly like LiveLifecycleFixture's copy: a
            // tripwire that can be swallowed is not one. Whether this is where the run gets
            // verified depends on the FILTER, which LiveTierRunPlan knows and this fixture
            // does not.
            LiveStoreCountTripwire.CollectionFinished(LiveCollections.Phase1);
        }
    }

    /// <summary>This fixture's own teardown, separated so the tripwire signal cannot be skipped.</summary>
    private void DisposeCore()
    {
        if (_session.IsValueCreated)
        {
            // Releases COM references only - Outlook keeps running (S7: never kill/close).
            _session.Value.Dispose();
        }
    }
}

[CollectionDefinition(LiveCollections.Phase1)]
public sealed class LivePhase1Collection : ICollectionFixture<LivePhase1Fixture>
{
}
