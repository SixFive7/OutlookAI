using OutlookAI.Core.Com;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the bystander tier: the count tripwire's value rests entirely on there being a store
/// whose every change is evidence of a fault, and until this existed that property was an
/// accident of which list a store did NOT appear in.
/// <para>
/// The defect these tests close: <c>LiveStoreWriteGuard</c> passed
/// <c>expectedStoreDisplayNames</c> as the identity-draft grant, so every configured non-hub
/// store had draft-create and delete - including the VM's bystander, which has to be in that
/// list to be censused at all. "Nothing but the suite changes a mailbox here, so any decrease
/// is a fault" was therefore false in the one place the guard depends on it.
/// </para>
/// <para>
/// Synthetic store names only - no real mailbox identifier belongs in this PUBLIC repo (S6).
/// Nothing here touches Outlook, a mailbox or a settings file.
/// </para>
/// </summary>
public sealed class TripwireBystanderStoreTests
{
    private const string Hub = "hub@example.test";
    private const string Identity = "other@example.test";
    private const string Bystander = "OutlookAI Bystander";
    private const string DelegateStore = "Someone Else";

    /// <summary>The runbook's three-store VM layout: the bystander sits INSIDE the primary list.</summary>
    private static LiveTestSettings VmSettings()
    {
        return new LiveTestSettings
        {
            MachineProfile = LiveMachineProfile.Portable,
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Identity, Bystander },
            BystanderStoreDisplayNames = new List<string> { Bystander },
        };
    }

    // ------------------------------------------------------- the allowlist's bystander tier

    [Theory]
    [InlineData(StoreWriteKind.Send)]
    [InlineData(StoreWriteKind.Draft)]
    [InlineData(StoreWriteKind.Delete)]
    [InlineData(StoreWriteKind.Move)]
    [InlineData(StoreWriteKind.Folder)]
    public void ABystanderIsDeniedEveryKindOfWrite(StoreWriteKind kind)
    {
        StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(VmSettings());

        Assert.True(allowlist.IsBystander(Bystander));
        Assert.False(allowlist.IsAllowed(Bystander, kind));
        InvalidOperationException ex =
            Assert.Throws<InvalidOperationException>(() => allowlist.Assert(Bystander, kind, "unit"));
        Assert.Contains("REFUSING", ex.Message, StringComparison.Ordinal);
        Assert.Contains("BYSTANDER", ex.Message, StringComparison.Ordinal);
        Assert.Contains(Bystander, ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheBystanderDeclarationBeatsTheIdentityDraftGrant()
    {
        // THE defect, in one assertion. The bystander is a non-hub entry of
        // expectedStoreDisplayNames - it has to be, or the census never visits it - so it
        // arrives at the allowlist already inside the identity grant. A declaration checked
        // after that grant is a declaration that never applies to the store it was written for.
        StoreWriteAllowlist allowlist = new(
            Hub,
            identityDraftStores: new[] { Identity, Bystander },
            knownReadOnlyStores: null,
            bystanderStores: new[] { Bystander });

        Assert.False(allowlist.IsAllowed(Bystander, StoreWriteKind.Draft));
        Assert.False(allowlist.IsAllowed(Bystander, StoreWriteKind.Delete));

        // ... and the grant still works for the store that is only in the grant.
        Assert.True(allowlist.IsAllowed(Identity, StoreWriteKind.Draft));
        Assert.True(allowlist.IsAllowed(Identity, StoreWriteKind.Delete));
    }

    [Fact]
    public void TheGuardCarriesTheDeclarationFromTheSettingsFileIntoTheAllowlist()
    {
        StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(VmSettings());

        Assert.Contains(Bystander, allowlist.Bystanders);
        Assert.False(allowlist.IsAllowed(Bystander, StoreWriteKind.Draft));
        Assert.True(allowlist.IsAllowed(Identity, StoreWriteKind.Draft));
        Assert.True(allowlist.IsAllowed(Hub, StoreWriteKind.Folder));
    }

    [Fact]
    public void BystanderNamesMatchCaseInsensitivelyAndBlankEntriesAreIgnored()
    {
        StoreWriteAllowlist allowlist = new(
            Hub, identityDraftStores: new[] { Identity }, knownReadOnlyStores: null,
            bystanderStores: new[] { "OUTLOOKAI BYSTANDER", "  ", string.Empty });

        Assert.False(allowlist.IsAllowed(Bystander, StoreWriteKind.Draft));
        Assert.True(allowlist.IsBystander("outlookai bystander"));
        Assert.Single(allowlist.Bystanders);
    }

    // ------------------------------------------- which stores the identity tests may draft in

    [Fact]
    public void TheIdentityAccountsAreTheStoresTheAllowlistWouldActuallyLetADraftInto()
    {
        // LivePhase4Fixture.IdentityAccounts is this call. Derived any other way - "the
        // configured primaries that are not the hub", which is what it used to be - a denied
        // store still ends up in the list and the test throws at the write guard mid-run
        // instead of leaving that mailbox alone.
        StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(VmSettings());

        IReadOnlyList<string> accounts = allowlist.IdentityAccountsAmong(
            new[] { Hub, Identity, Bystander, Identity });

        Assert.Equal(new[] { Identity }, accounts);
    }

    [Fact]
    public void WithNoBystanderDeclaredTheIdentityAccountsAreUnchanged()
    {
        // The grant is NOT narrowed by this work: two live tests
        // (LiveDraftTests.IdentityDrafts_..., LiveDraftOptionsTests.NewDraft_BusinessAccounts_...)
        // create one tagged, never-displayed draft per business account and delete it, and a
        // Production profile declares no bystanders at all.
        LiveTestSettings production = new()
        {
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Identity, "third@example.test" },
            ExpectedDelegateStoreDisplayNames = new List<string> { DelegateStore },
        };

        IReadOnlyList<string> accounts = LiveStoreWriteGuard.Build(production)
            .IdentityAccountsAmong(production.ExpectedStoreDisplayNames);

        Assert.Equal(new[] { Identity, "third@example.test" }, accounts);
    }

    // ---------------------------------------------------------------- the watched set

    [Fact]
    public void TheCensusWatchesEveryDeclaredBystanderEvenOneNoOtherListNames()
    {
        LiveTestSettings settings = new()
        {
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Identity },
            ExpectedDelegateStoreDisplayNames = new List<string> { DelegateStore },
            BystanderStoreDisplayNames = new List<string> { Bystander },
        };

        IReadOnlyList<string> watched = LiveStoreCountTripwire.WatchedStores(settings);

        Assert.Equal(new[] { Hub, Identity, DelegateStore, Bystander }, watched);
    }

    [Fact]
    public void ABystanderAlreadyAmongThePrimariesIsWatchedOnce()
    {
        IReadOnlyList<string> watched = LiveStoreCountTripwire.WatchedStores(VmSettings());

        Assert.Equal(new[] { Hub, Identity, Bystander }, watched);
    }

    // ----------------------------------------------------------------- the soundness gate

    [Fact]
    public void AConsistentThreeStoreLayoutIsSound()
    {
        LiveTestSettings settings = VmSettings();

        TripwireWatchReport report = TripwireWatchSoundness.Require(
            LiveStoreCountTripwire.WatchedStores(settings),
            LiveStoreWriteGuard.Build(settings),
            settings.BystanderStoreDisplayNames);

        Assert.True(report.Sound);
        Assert.Equal(new[] { Bystander }, report.Bystanders);
        Assert.Equal(new[] { Bystander }, report.Policed);
        Assert.Equal(new[] { Identity }, report.Writable);
        Assert.False(report.ProvesNothing);
        Assert.Null(report.Refusal());
    }

    [Fact]
    public void AWatchedStoreTheSuiteMayWriteToIsRefusedWhenItWasDeclaredABystander()
    {
        // The pre-fix world, reconstructed: the allowlist built without the declaration, so
        // the bystander keeps the draft+delete it got for being a non-hub primary. This is the
        // assertion the whole change exists for, and it must FIRE.
        LiveTestSettings settings = VmSettings();
        StoreWriteAllowlist unfixed = new(
            settings.TestHubStoreDisplayName,
            settings.ExpectedStoreDisplayNames,
            settings.ExpectedDelegateStoreDisplayNames);

        TripwireWatchReport report = TripwireWatchSoundness.Assess(
            LiveStoreCountTripwire.WatchedStores(settings), unfixed, settings.BystanderStoreDisplayNames);

        Assert.False(report.Sound);
        Assert.Contains(Bystander, report.Writable);
        Assert.DoesNotContain(Bystander, report.Policed);

        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
            () => TripwireWatchSoundness.Require(
                LiveStoreCountTripwire.WatchedStores(settings), unfixed, settings.BystanderStoreDisplayNames));
        Assert.Contains("REFUSING to run the live tier", ex.Message, StringComparison.Ordinal);
        Assert.Contains(Bystander, ex.Message, StringComparison.Ordinal);
        Assert.Contains("draft, delete", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheHubDeclaredItsOwnBystanderIsRefusedRatherThanResolved()
    {
        // IsAllowed answers for the hub first and cannot be told otherwise - denying the hub
        // instead would fail a hundred tests far from the mistake. So this is the case the
        // gate exists to catch, and it is why the gate asks the allowlist rather than
        // restating the list it was built from.
        LiveTestSettings settings = new()
        {
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub },
            BystanderStoreDisplayNames = new List<string> { Hub },
        };
        StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(settings);

        Assert.True(allowlist.IsAllowed(Hub, StoreWriteKind.Send));
        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
            () => TripwireWatchSoundness.Require(
                LiveStoreCountTripwire.WatchedStores(settings), allowlist, settings.BystanderStoreDisplayNames));
        Assert.Contains("designated test hub", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ABystanderTheCensusDoesNotWatchIsRefused()
    {
        // A declaration nothing censuses guarantees nothing. Fires if the bystanders ever fall
        // out of WatchedStores, which is the other half of "watched and never written".
        LiveTestSettings settings = VmSettings();
        StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(settings);

        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
            () => TripwireWatchSoundness.Require(
                new[] { Hub, Identity }, allowlist, settings.BystanderStoreDisplayNames));

        Assert.Contains("the census does not watch it", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ADelegateMailboxIsABystanderInFactWithoutBeingDeclared()
    {
        // A Production profile needs no declaration: its delegate/shared mailboxes are already
        // denied and already watched, which is the whole property.
        LiveTestSettings production = new()
        {
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Identity },
            ExpectedDelegateStoreDisplayNames = new List<string> { DelegateStore },
        };

        TripwireWatchReport report = TripwireWatchSoundness.Require(
            LiveStoreCountTripwire.WatchedStores(production),
            LiveStoreWriteGuard.Build(production),
            production.BystanderStoreDisplayNames);

        Assert.Empty(report.Bystanders);
        Assert.Equal(new[] { DelegateStore }, report.Policed);
        Assert.Equal(new[] { Identity }, report.Writable);
        Assert.False(report.ProvesNothing);
    }

    [Fact]
    public void OnePstThatIsAlsoTheHubIsRefusedRatherThanWarnedAbout()
    {
        // The configuration Docs/live-tier-on-the-vm.md section 1.3 tells a rebuilder to avoid.
        // This USED to warn and proceed, on the argument that a one-store smoke run is
        // legitimate. It refuses as of 2026-08-24: the run it would allow prints '0 failure(s)'
        // by construction, and that line then reads as coverage. TripwireVacuousCensusTests
        // owns the detail; this asserts the reversal where the old expectation lived, so nobody
        // re-derives the warning from a stale sibling test.
        LiveTestSettings settings = new()
        {
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub },
        };
        TripwireWatchReport report = TripwireWatchSoundness.Assess(
            LiveStoreCountTripwire.WatchedStores(settings),
            LiveStoreWriteGuard.Build(settings),
            settings.BystanderStoreDisplayNames);

        Assert.True(report.Sound);
        Assert.True(report.ProvesNothing);
        Assert.False(report.Usable);
        Assert.Contains("NO STORE THIS CENSUS WATCHES CAN PRODUCE A FAILURE", report.Describe(), StringComparison.Ordinal);

        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
            () => TripwireWatchSoundness.Require(
                LiveStoreCountTripwire.WatchedStores(settings),
                LiveStoreWriteGuard.Build(settings),
                settings.BystanderStoreDisplayNames));
        Assert.Contains("REFUSING to run the live tier", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheSettingsSummaryCountsTheBystanders()
    {
        Assert.Contains("bystanders=1", VmSettings().Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheDeclarationSurvivesTheJsonRoundTrip()
    {
        // The loader, not a hand-built object: a settings key nothing deserializes reads as
        // configured and behaves as absent, which is how the JsonStringEnumConverter went
        // missing for machineProfile.
        LiveTestSettings settings = LiveTestSettings.Parse(
            """
            {
              "machineProfile": "Portable",
              "testHubStoreDisplayName": "hub@example.test",
              "expectedStoreDisplayNames": [ "hub@example.test", "OutlookAI Bystander" ],
              "bystanderStoreDisplayNames": [ "OutlookAI Bystander" ]
            }
            """);

        Assert.Equal(new[] { Bystander }, settings.BystanderStoreDisplayNames);
        Assert.False(LiveStoreWriteGuard.Build(settings).IsAllowed(Bystander, StoreWriteKind.Draft));
    }

    // --------------------------------------------------- the gate on the live entry point

    [Fact]
    public void TheLiveTripwireRefusesAContradictoryConfigurationBeforeItAsksAboutOutlook()
    {
        // Drives the real EnsureBaseline, which is the line a mutation would delete. It is
        // reachable here only because the soundness gate runs AHEAD of the health gate and
        // ahead of every COM call - so this test also pins that ordering.
        //
        // The liveness override is forced to Hung for the duration: if the ordering is ever
        // reversed this test fails on the preflight's message instead of touching Outlook,
        // which is the outcome that must not be left to luck on a machine with a real profile.
        LiveTestSettings settings = VmSettings();
        settings.BystanderStoreDisplayNames.Add(Hub);

        string? saved = Environment.GetEnvironmentVariable(LiveOutlookPreflight.LivenessOverrideVariable);
        try
        {
            Environment.SetEnvironmentVariable(
                LiveOutlookPreflight.LivenessOverrideVariable, OutlookLivenessState.Hung.ToString());

            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
                () => LiveStoreCountTripwire.EnsureBaseline(settings));

            Assert.Contains("REFUSING to run the live tier", ex.Message, StringComparison.Ordinal);
            Assert.Contains("designated test hub", ex.Message, StringComparison.Ordinal);
            Assert.False(LiveStoreCountTripwire.HasBaseline);
        }
        finally
        {
            Environment.SetEnvironmentVariable(LiveOutlookPreflight.LivenessOverrideVariable, saved);
        }
    }
}
