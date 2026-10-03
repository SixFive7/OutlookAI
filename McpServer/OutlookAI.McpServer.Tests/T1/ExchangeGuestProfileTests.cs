using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins what <see cref="LiveMachineProfile.ExchangeGuest"/> - the Exchange test VM, <c>OutlookAI-Exchange</c>
/// (decided by the maintainer 2026-10-03, Q108 to Q111) - means to every gate that reads a machine
/// profile: it parses from a settings file, it refuses every write until the maintainer approves the
/// Phase 2 write-safety design, it gets the re-censuses a real Exchange mailbox needs and never the
/// re-run, nobody is at its keyboard, and - because its one mailbox is also its hub - the count tripwire
/// censuses that hub like any other store instead of exempting it.
/// <para>
/// Synthetic store names only - no real mailbox identifier belongs in this PUBLIC repo (S6). Nothing
/// here touches Outlook, a mailbox or a settings file.
/// </para>
/// </summary>
public sealed class ExchangeGuestProfileTests
{
    private const string Hub = "hub@example.test";

    private static LiveTestSettings ExchangeVm()
    {
        return new LiveTestSettings
        {
            MachineProfile = LiveMachineProfile.ExchangeGuest,
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub },
            IndexedStoreDisplayNames = new List<string> { Hub },
        };
    }

    [Fact]
    public void TheProfileParsesFromASettingsFile_AndNeedsNoProbeTermOrSubjectOnlyProbe()
    {
        // A real mailbox has no known probe word or subject-only population; finding one would mean
        // reading its owner's mail. Only Production demanded them, and this is not Production.
        string json = "{ \"machineProfile\": \"ExchangeGuest\", \"testHubStoreDisplayName\": \"" + Hub + "\", "
            + "\"expectedStoreDisplayNames\": [\"" + Hub + "\"], \"indexedStoreDisplayNames\": [\"" + Hub + "\"], "
            + "\"expectedDelegateStoreDisplayNames\": [], \"bystanderStoreDisplayNames\": [], \"probeTerm\": \"\" }";

        LiveTestSettings settings = LiveTestSettings.Parse(json);

        Assert.Equal(LiveMachineProfile.ExchangeGuest, settings.MachineProfile);
        Assert.Equal(2, (int)settings.MachineProfile);
    }

    [Fact]
    public void TheExchangeVm_RefusesEveryWrite_AndSaysWhyAndWhatLiftsIt()
    {
        Assert.True(LiveWriteAccess.RefusesEveryWrite(LiveMachineProfile.ExchangeGuest));
        Assert.True(ExchangeVm().RefusesEveryWrite);
        Assert.Contains("writes=NONE (read-only machine)", ExchangeVm().Describe(), StringComparison.Ordinal);

        string reason = LiveWriteAccess.ReadOnlyReason(LiveMachineProfile.ExchangeGuest);
        Assert.StartsWith(LiveWriteAccess.ReadOnlyMachine, reason, StringComparison.Ordinal);
        Assert.Contains("'ExchangeGuest'", reason, StringComparison.Ordinal);
        Assert.Contains("Phase 2", reason, StringComparison.Ordinal);
        Assert.Contains("Testbed/README.md section 4e", reason, StringComparison.Ordinal);
        Assert.Contains("Writes=Nothing", reason, StringComparison.Ordinal);

        // The in-process guard asks the same decision: every store refused, the hub included.
        StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(ExchangeVm());
        Assert.True(allowlist.RefusesEveryWrite);
        foreach (StoreWriteKind kind in Enum.GetValues<StoreWriteKind>())
        {
            Assert.False(allowlist.IsAllowed(Hub, kind), "the hub allowed " + kind);
        }

        // And so does the MCP client's posture, once a run is opted in.
        Assert.Equal(
            StdioWritePosture.ReadOnly,
            LiveWriteAccess.StdioPostureFor(LiveRunOptIn.Verdict.Open, ExchangeVm));
    }

    [Fact]
    public void TheExchangeVm_GetsTheReCensusesARealMailboxNeeds_AndNeverAReRun()
    {
        TripwireRetryPolicy policy = TripwireRetryPolicy.For(LiveMachineProfile.ExchangeGuest);

        Assert.Same(TripwireRetryPolicy.ExchangeGuest, policy);
        Assert.Equal(TripwireRetryLadder.MaxReCensuses, policy.MaxReCensuses);
        Assert.Equal(TripwireRetryLadder.ReCensusGapSeconds, policy.ReCensusGapSeconds);
        Assert.Equal(0, policy.MaxImplicatedReRuns);
        Assert.Equal(0, TripwireRetryPolicy.For(LiveMachineProfile.ExchangeGuest, isReRunChild: true).MaxImplicatedReRuns);
    }

    [Fact]
    public void NobodyIsAtTheExchangeVmsKeyboard()
    {
        Assert.False(ExchangeVm().AUserMayBeAtTheKeyboard);

        // Its missing populations do not refuse a run: the tests that need them are filtered out there.
        ExchangeVm().RequireProductionPopulation("a population the Exchange VM does not have");
    }

    [Fact]
    public void OnAReadOnlyMachine_TheHubIsNotExempt()
    {
        Assert.Equal(string.Empty, LiveStoreCountTripwire.ExemptHub(ExchangeVm()));
        Assert.Equal(
            string.Empty,
            LiveStoreCountTripwire.ExemptHub(new LiveTestSettings { TestHubStoreDisplayName = Hub, ExpectedStoreDisplayNames = new List<string> { Hub } }));

        // Where the suite may write in it, the hub stays exempt: its churn is tagged and swept there.
        Assert.Equal(
            Hub,
            LiveStoreCountTripwire.ExemptHub(new LiveTestSettings
            {
                MachineProfile = LiveMachineProfile.Portable,
                TestHubStoreDisplayName = Hub,
                ExpectedStoreDisplayNames = new List<string> { Hub },
            }));
    }

    [Fact]
    public void AHubThatIsTheMachinesOnlyStore_IsACensusThatCanFail_OnAReadOnlyMachine()
    {
        // The same one-store layout is refused on a machine that may write in its hub
        // (TripwireVacuousCensusTests.OneStoreThatIsAlsoTheHubIsRefusedAndToldHowToGetASecondStore).
        TripwireWatchReport report = TripwireWatchSoundness.Require(
            LiveStoreCountTripwire.WatchedStores(ExchangeVm()),
            LiveStoreWriteGuard.Build(ExchangeVm()),
            ExchangeVm().BystanderStoreDisplayNames);

        Assert.True(report.Usable);
        Assert.False(report.ProvesNothing);
        Assert.Equal(new[] { Hub }, report.Policed);
        Assert.Empty(report.Writable);
    }

    [Fact]
    public void WithTheHubPoliced_ADepartureFromItFails_AndAnArrivalIsOnlyNoted()
    {
        // The live mailbox receives real mail during a run: counts may rise. A departure that is not
        // a filing into another ordinary folder of the same store fails, the hub included.
        string exempt = LiveStoreCountTripwire.ExemptHub(ExchangeVm());

        Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> before = Census(
            ("Inbox", FolderCensus.WithItems(new[] { Item("id-1"), Item("id-2") })),
            ("Sent Items", FolderCensus.WithItems(new[] { Item("id-3") })));

        Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> arrived = Census(
            ("Inbox", FolderCensus.WithItems(new[] { Item("id-1"), Item("id-2"), Item("id-new") })),
            ("Sent Items", FolderCensus.WithItems(new[] { Item("id-3") })));
        TripwireVerdict arrival = StoreCountTripwire.Evaluate(before, arrived, exempt);
        Assert.False(arrival.Failed, string.Join(Environment.NewLine, arrival.Failures));

        Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> lost = Census(
            ("Inbox", FolderCensus.WithItems(new[] { Item("id-1") })),
            ("Sent Items", FolderCensus.WithItems(new[] { Item("id-3") })));
        Assert.True(StoreCountTripwire.Evaluate(before, lost, exempt).Failed);

        // The same departure with the hub exempt - a machine that may write there - does not fail here.
        Assert.False(StoreCountTripwire.Evaluate(before, lost, Hub).Failed);
    }

    private static Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> Census(
        params (string Folder, FolderCensus Census)[] folders)
    {
        Dictionary<string, FolderCensus> byFolder = new(StringComparer.OrdinalIgnoreCase);
        foreach ((string folder, FolderCensus census) in folders)
        {
            byFolder[folder] = census;
        }

        return new Dictionary<string, IReadOnlyDictionary<string, FolderCensus>>(StringComparer.OrdinalIgnoreCase)
        {
            [Hub] = byFolder,
        };
    }

    private static CensusItem Item(string id)
    {
        return new CensusItem(id, "fp-" + id, false);
    }
}
