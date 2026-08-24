using OutlookAI.Core.Com;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the refusal of a configuration the count tripwire can prove NOTHING from: no watched
/// store that is both non-hub and denied every write.
/// <para>
/// <b>What changed on 2026-08-24, and why.</b> That state used to print
/// <c>NO STORE THIS CENSUS WATCHES CAN PRODUCE A FAILURE</c> and carry on, on the argument that a
/// smoke run over one PST is legitimate. The argument answers the wrong question. In that state
/// the census still runs, still visits every folder, still identifies nothing and still prints
/// <c>0 failure(s)</c> - a line produced by arithmetic that could not have reached any other
/// answer, and which then appears in a run report indistinguishable from an earned one. A guard
/// that cannot fail is worse than no guard: no guard leaves a hole somebody can see.
/// </para>
/// <para>
/// <b>There is deliberately no opt-out.</b> Any flag by which a machine declared "I accept a
/// census that cannot fail" would be settable in the same file, with the same effort, as the
/// declaration that actually fixes the problem - and it would buy the run nothing except the
/// tripwire being off. The remedy is one settings key when the machine already has a second
/// store, and one empty PST plus two keys when it does not; the refusal text says which.
/// </para>
/// <para>
/// Synthetic store names only - no real mailbox identifier belongs in this PUBLIC repo (S6).
/// Nothing here touches Outlook, a mailbox or a settings file.
/// </para>
/// </summary>
public sealed class TripwireVacuousCensusTests
{
    private const string Hub = "hub@example.test";
    private const string Identity = "other@example.test";
    private const string Bystander = "OutlookAI Bystander";
    private const string DelegateStore = "Someone Else";
    private const string NoFailableStore = "NO STORE THIS CENSUS WATCHES CAN PRODUCE A FAILURE";

    private static TripwireWatchReport Assess(LiveTestSettings settings)
    {
        return TripwireWatchSoundness.Assess(
            LiveStoreCountTripwire.WatchedStores(settings),
            LiveStoreWriteGuard.Build(settings),
            settings.BystanderStoreDisplayNames);
    }

    private static InvalidOperationException Refused(LiveTestSettings settings)
    {
        return Assert.Throws<InvalidOperationException>(
            () => TripwireWatchSoundness.Require(
                LiveStoreCountTripwire.WatchedStores(settings),
                LiveStoreWriteGuard.Build(settings),
                settings.BystanderStoreDisplayNames));
    }

    // -------------------------------------------------------------- the refusal, both shapes

    [Fact]
    public void AHubWhoseOnlyCompanyIsTheIdentityGrantIsRefused()
    {
        // The shape a PST machine actually lands in: two stores, and the second one is inside
        // the identity-draft grant purely for being a non-hub primary. The census watches it and
        // the suite may write to it, so a change there is not evidence of anything - which
        // leaves the census with nothing at all it could fail on.
        LiveTestSettings settings = new()
        {
            MachineProfile = LiveMachineProfile.Portable,
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Identity },
        };

        TripwireWatchReport report = Assess(settings);
        Assert.True(report.Sound, "nothing is declared, so nothing is contradicted");
        Assert.True(report.ProvesNothing);
        Assert.False(report.Usable);
        Assert.Equal(new[] { Identity }, report.Writable);
        Assert.Empty(report.Policed);

        string message = Refused(settings).Message;
        Assert.Contains("REFUSING to run the live tier", message, StringComparison.Ordinal);
        Assert.Contains(NoFailableStore, message, StringComparison.Ordinal);

        // The remedy for THIS shape is one settings key, and the message must say so rather
        // than telling somebody who already has a second store to go and make one.
        Assert.Contains("bystanderStoreDisplayNames", message, StringComparison.Ordinal);
        Assert.Contains(Identity, message, StringComparison.Ordinal);
        Assert.DoesNotContain("add a store to the Outlook profile", message, StringComparison.Ordinal);
    }

    [Fact]
    public void OneStoreThatIsAlsoTheHubIsRefusedAndToldHowToGetASecondStore()
    {
        LiveTestSettings settings = new()
        {
            MachineProfile = LiveMachineProfile.Portable,
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub },
        };

        string message = Refused(settings).Message;
        Assert.Contains(NoFailableStore, message, StringComparison.Ordinal);

        // Whoever reads this is configuring a machine, not reading the source. All three steps,
        // both settings keys by name, and the two documents that say the same thing at length.
        Assert.Contains("add a store to the Outlook profile", message, StringComparison.Ordinal);
        Assert.Contains("expectedStoreDisplayNames", message, StringComparison.Ordinal);
        Assert.Contains("bystanderStoreDisplayNames", message, StringComparison.Ordinal);
        Assert.Contains("Docs/live-tier-on-the-vm.md section 1.3", message, StringComparison.Ordinal);
        Assert.Contains("Testbed/live-test-settings.example.json", message, StringComparison.Ordinal);
        Assert.Contains(Hub, message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheRefusalSaysOutrightThatNothingTurnsItOff()
    {
        // The decision, written where a future reader will look for it. A settings flag that
        // suppressed this refusal would be exactly as much typing as the declaration that fixes
        // the configuration, and would leave the tripwire off; if one is ever added, this line
        // has to be removed first, deliberately.
        LiveTestSettings settings = new()
        {
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub },
        };

        Assert.Contains(
            "no setting that turns this refusal off", Refused(settings).Message, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ what lifts the refusal

    [Fact]
    public void DeclaringAStoreTheMachineAlreadyHasLiftsTheRefusal()
    {
        // The one-line fix the message promises, executed. Same two stores as the refused case
        // above; the only difference is the declaration.
        LiveTestSettings settings = new()
        {
            MachineProfile = LiveMachineProfile.Portable,
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Identity },
            BystanderStoreDisplayNames = new List<string> { Identity },
        };

        TripwireWatchReport report = TripwireWatchSoundness.Require(
            LiveStoreCountTripwire.WatchedStores(settings),
            LiveStoreWriteGuard.Build(settings),
            settings.BystanderStoreDisplayNames);

        Assert.True(report.Usable);
        Assert.False(report.ProvesNothing);
        Assert.Equal(new[] { Identity }, report.Policed);
        Assert.Empty(report.Writable);
        Assert.Null(report.Refusal());
    }

    [Fact]
    public void AnEmptyDeclarationDoesNotLiftIt()
    {
        // Guards the cheapest wrong fix: a bystanderStoreDisplayNames key present but blank,
        // or full of whitespace, which reads as configured and denies nothing.
        LiveTestSettings settings = new()
        {
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Identity },
            BystanderStoreDisplayNames = new List<string> { "  ", string.Empty },
        };

        Assert.Contains(NoFailableStore, Refused(settings).Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AProductionProfileWithADelegateMailboxIsNotRefused()
    {
        // The regression that matters most: the maintainer's real profile must keep running. Its
        // delegate/shared mailboxes are denied every write and already censused, so they are
        // bystanders in fact and this refusal never reaches him.
        LiveTestSettings settings = new()
        {
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Identity },
            ExpectedDelegateStoreDisplayNames = new List<string> { DelegateStore },
        };

        TripwireWatchReport report = TripwireWatchSoundness.Require(
            LiveStoreCountTripwire.WatchedStores(settings),
            LiveStoreWriteGuard.Build(settings),
            settings.BystanderStoreDisplayNames);

        Assert.True(report.Usable);
        Assert.Equal(new[] { DelegateStore }, report.Policed);
    }

    [Fact]
    public void TheRunbooksThreeStoreLayoutIsNotRefused()
    {
        LiveTestSettings settings = new()
        {
            MachineProfile = LiveMachineProfile.Portable,
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Identity, Bystander },
            BystanderStoreDisplayNames = new List<string> { Bystander },
        };

        Assert.True(Assess(settings).Usable);
    }

    // ----------------------------------------------------- the gate cannot be made vacuous

    [Theory]
    [InlineData(true, true)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(false, false)]
    public void RefusalAndUsableAlwaysAgree(bool declareBystander, bool haveSecondStore)
    {
        // Refusal() and Usable are two spellings of one decision, and a mutation that makes
        // either of them unconditional shows up as a disagreement here rather than as a live
        // tier that runs when it should not.
        List<string> stores = haveSecondStore
            ? new List<string> { Hub, Identity }
            : new List<string> { Hub };
        LiveTestSettings settings = new()
        {
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = stores,
            BystanderStoreDisplayNames = declareBystander && haveSecondStore
                ? new List<string> { Identity }
                : new List<string>(),
        };

        TripwireWatchReport report = Assess(settings);

        Assert.Equal(report.Usable, report.Refusal() == null);
        Assert.Equal(report.Sound && !report.ProvesNothing, report.Usable);
        Assert.Equal(declareBystander && haveSecondStore, report.Usable);
    }

    [Fact]
    public void BothFaultsAtOnceAreReportedInOneMessage()
    {
        // A machine being configured has both faults together often enough - here the hub is
        // named as its own bystander, which contradicts the declaration AND leaves nothing
        // policed. Finding the second fault only after a live run fixed the first costs another
        // live run, so one message carries both.
        LiveTestSettings settings = new()
        {
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub },
            BystanderStoreDisplayNames = new List<string> { Hub },
        };

        TripwireWatchReport report = Assess(settings);
        Assert.False(report.Sound);
        Assert.True(report.ProvesNothing);

        string message = Refused(settings).Message;
        Assert.Contains("designated test hub", message, StringComparison.Ordinal);
        Assert.Contains(NoFailableStore, message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheOneLineSummarySaysRefusedRatherThanReadingLikeAWarning()
    {
        LiveTestSettings settings = new()
        {
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub },
        };

        string described = Assess(settings).Describe();

        Assert.Contains("0 store(s) this census can fail on", described, StringComparison.Ordinal);
        Assert.Contains("REFUSED", described, StringComparison.Ordinal);
        Assert.Contains(NoFailableStore, described, StringComparison.Ordinal);
    }

    // --------------------------------------------------- the gate on the live entry point

    [Fact]
    public void TheLiveTripwireRefusesAVacuousConfigurationBeforeItAsksAboutOutlook()
    {
        // Drives the real EnsureBaseline. Reachable from a runner with no Outlook and no
        // settings file only because the soundness gate runs AHEAD of the health gate and ahead
        // of every COM call, so this pins that ordering for the second refusal as well as the
        // first. The liveness override is forced to Hung: if the ordering is ever reversed this
        // fails on the preflight's message instead of reaching a real profile.
        LiveTestSettings settings = new()
        {
            MachineProfile = LiveMachineProfile.Portable,
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Identity },
        };

        string? saved = Environment.GetEnvironmentVariable(LiveOutlookPreflight.LivenessOverrideVariable);
        try
        {
            Environment.SetEnvironmentVariable(
                LiveOutlookPreflight.LivenessOverrideVariable, OutlookLivenessState.Hung.ToString());

            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
                () => LiveStoreCountTripwire.EnsureBaseline(settings));

            Assert.Contains(NoFailableStore, ex.Message, StringComparison.Ordinal);
            Assert.False(LiveStoreCountTripwire.HasBaseline);
        }
        finally
        {
            Environment.SetEnvironmentVariable(LiveOutlookPreflight.LivenessOverrideVariable, saved);
        }
    }
}
