using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the code-enforced store allowlist (soak fix 16, part A2): a test write outside the
/// designated test mailbox must be impossible, not merely discouraged. Synthetic store
/// names only - no real mailbox identifier belongs in this PUBLIC repo (S6).
/// </summary>
public sealed class StoreWriteAllowlistTests
{
    private const string Hub = "hub@example.test";
    private const string Identity = "other@example.test";
    private const string DelegateStore = "Someone Else";

    private static StoreWriteAllowlist Build()
    {
        return new StoreWriteAllowlist(
            Hub,
            identityDraftStores: new[] { Hub, Identity },
            knownReadOnlyStores: new[] { DelegateStore, "Another Person" });
    }

    [Theory]
    [InlineData(StoreWriteKind.Send)]
    [InlineData(StoreWriteKind.Draft)]
    [InlineData(StoreWriteKind.Delete)]
    [InlineData(StoreWriteKind.Move)]
    [InlineData(StoreWriteKind.Folder)]
    public void Hub_PermitsEveryKindOfWrite(StoreWriteKind kind)
    {
        StoreWriteAllowlist allowlist = Build();

        Assert.True(allowlist.IsAllowed(Hub, kind));
        Assert.Equal(Hub, allowlist.Assert(Hub, kind, "unit"));
    }

    [Theory]
    [InlineData(StoreWriteKind.Send)]
    [InlineData(StoreWriteKind.Draft)]
    [InlineData(StoreWriteKind.Delete)]
    [InlineData(StoreWriteKind.Move)]
    [InlineData(StoreWriteKind.Folder)]
    public void DelegateStore_ThrowsForEveryKindOfWrite(StoreWriteKind kind)
    {
        StoreWriteAllowlist allowlist = Build();

        Assert.False(allowlist.IsAllowed(DelegateStore, kind));
        InvalidOperationException ex =
            Assert.Throws<InvalidOperationException>(() => allowlist.Assert(DelegateStore, kind, "unit"));
        Assert.Contains("REFUSING", ex.Message, StringComparison.Ordinal);
        Assert.Contains("READ-ONLY", ex.Message, StringComparison.Ordinal);
        Assert.Contains(DelegateStore, ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void UnknownStore_Throws()
    {
        StoreWriteAllowlist allowlist = Build();

        Assert.False(allowlist.IsAllowed("stranger@example.test", StoreWriteKind.Draft));
        Assert.False(allowlist.IsAllowed(null, StoreWriteKind.Draft));
        Assert.False(allowlist.IsAllowed("  ", StoreWriteKind.Draft));
        Assert.Throws<InvalidOperationException>(
            () => allowlist.Assert("stranger@example.test", StoreWriteKind.Draft, "unit"));
        Assert.Throws<InvalidOperationException>(() => allowlist.Assert(null, StoreWriteKind.Draft, "unit"));
    }

    [Fact]
    public void IdentityStores_MayDraftAndDelete_ButNeverSendMoveOrCreateFolders()
    {
        // The S2 exception is narrow on purpose: one tagged, never-displayed draft per
        // business account, created and deleted. Nothing else.
        StoreWriteAllowlist allowlist = Build();

        Assert.True(allowlist.IsAllowed(Identity, StoreWriteKind.Draft));
        Assert.True(allowlist.IsAllowed(Identity, StoreWriteKind.Delete));
        Assert.False(allowlist.IsAllowed(Identity, StoreWriteKind.Send));
        Assert.False(allowlist.IsAllowed(Identity, StoreWriteKind.Move));
        Assert.False(allowlist.IsAllowed(Identity, StoreWriteKind.Folder));

        InvalidOperationException ex =
            Assert.Throws<InvalidOperationException>(() => allowlist.Assert(Identity, StoreWriteKind.Send, "unit"));
        Assert.Contains("draft+delete only", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void StoreNamesMatchCaseInsensitively()
    {
        StoreWriteAllowlist allowlist = Build();

        Assert.True(allowlist.IsHub("HUB@EXAMPLE.TEST"));
        Assert.True(allowlist.IsAllowed("HUB@EXAMPLE.TEST", StoreWriteKind.Send));
        Assert.False(allowlist.IsAllowed("SOMEONE ELSE", StoreWriteKind.Draft));
    }

    [Fact]
    public void AContradictoryAllowlistIsRefusedAtConstruction()
    {
        // A read-only mailbox that also appears in the identity grant is a configuration
        // bug; resolving it silently is how a delegate store ends up writable.
        Assert.Throws<ArgumentException>(() => new StoreWriteAllowlist(
            Hub, identityDraftStores: new[] { DelegateStore }, knownReadOnlyStores: new[] { DelegateStore }));

        Assert.Throws<ArgumentException>(() => new StoreWriteAllowlist(" "));
    }

    [Fact]
    public void HubIsNeverDemotedToTheIdentityTier()
    {
        // The hub is passed in ExpectedStoreDisplayNames too; it must keep full rights.
        StoreWriteAllowlist allowlist = Build();

        Assert.DoesNotContain(Hub, allowlist.IdentityDraftStores);
        Assert.True(allowlist.IsAllowed(Hub, StoreWriteKind.Folder));
    }

    [Fact]
    public void GuardBuildsTheAllowlistFromTheLiveTestSettings()
    {
        // Derived, never hand-written: hub from the settings hub, identity grant from the
        // other configured primaries, delegates denied. On a machine that MAY write - a test
        // guest, which declares Portable; the workstation's shape is pinned below.
        LiveTestSettings settings = new()
        {
            MachineProfile = LiveMachineProfile.Portable,
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Identity },
            ExpectedDelegateStoreDisplayNames = new List<string> { DelegateStore },
        };

        StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(settings);

        Assert.False(allowlist.RefusesEveryWrite);
        Assert.True(allowlist.IsAllowed(Hub, StoreWriteKind.Send));
        Assert.True(allowlist.IsAllowed(Identity, StoreWriteKind.Draft));
        Assert.False(allowlist.IsAllowed(Identity, StoreWriteKind.Send));
        Assert.False(allowlist.IsAllowed(DelegateStore, StoreWriteKind.Delete));
    }

    // ------------------------------------------------------------------ Q74 layer 2: the read-only machine

    /// <summary>A second business mailbox on the workstation, beside <see cref="Identity"/>.</summary>
    private const string SecondPrimary = "third@example.test";

    /// <summary>
    /// The SHAPE of the maintainer's workstation settings, with synthetic names (S6): a hub and two
    /// other primary mailboxes, two delegate mailboxes, no bystander declared - and no machineProfile,
    /// which the loader reads as Production. The real file is gitignored and never read here.
    /// </summary>
    private static LiveTestSettings Workstation()
    {
        return new LiveTestSettings
        {
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Identity, SecondPrimary },
            ExpectedDelegateStoreDisplayNames = new List<string> { DelegateStore, "Another Person" },
            ProbeTerm = "term",
        };
    }

    public static IEnumerable<object[]> EveryKindOfWrite()
    {
        return Enum.GetValues<StoreWriteKind>().Select(kind => new object[] { kind });
    }

    [Theory]
    [MemberData(nameof(EveryKindOfWrite))]
    public void TheWorkstation_RefusesThisKindOfWrite_ToEveryStore_TheHubIncluded(StoreWriteKind kind)
    {
        // The rule (Q72) made code (Q74 layer 2): on the read-only machine an in-process write to
        // ANY store throws, the designated test mailbox first among them. Control: before Q74 the
        // same settings granted the hub every kind of write and the two other primaries draft and
        // delete, so the hub row and the two primary rows below failed.
        StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(Workstation());

        Assert.True(allowlist.RefusesEveryWrite);
        foreach (string store in new[] { Hub, Identity, SecondPrimary, DelegateStore, "Another Person", "stranger@example.test" })
        {
            Assert.False(allowlist.IsAllowed(store, kind), store + " still permits " + kind + " on the read-only machine");
            InvalidOperationException refused =
                Assert.Throws<InvalidOperationException>(() => allowlist.Assert(store, kind, "unit"));
            Assert.Contains(LiveWriteAccess.ReadOnlyMachine, refused.Message, StringComparison.Ordinal);
            Assert.Contains("'" + store + "'", refused.Message, StringComparison.Ordinal);

            // The one remedy that must never be followed is the one the ordinary refusal gives.
            Assert.Contains("Do NOT change machineProfile", refused.Message, StringComparison.Ordinal);
            Assert.DoesNotContain("widen the live-test settings", refused.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void TheWorkstation_GrantsItsOtherPrimaryMailboxes_NoIdentityDraftAnyMore()
    {
        // The finding the 2026-09-27 review recorded: with the workstation's settings the identity
        // grant opened draft AND delete on the two non-hub primary mailboxes - real business mail.
        // Closed by the read-only profile: the grant is empty and both kinds are refused.
        LiveTestSettings settings = Workstation();
        StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(settings);

        Assert.Empty(allowlist.IdentityAccountsAmong(settings.ExpectedStoreDisplayNames));
        foreach (string primary in new[] { Identity, SecondPrimary })
        {
            Assert.False(allowlist.IsAllowed(primary, StoreWriteKind.Draft));
            Assert.False(allowlist.IsAllowed(primary, StoreWriteKind.Delete));
        }
    }

    [Fact]
    public void ASettingsFileThatDeclaresNoProfile_IsReadOnly()
    {
        // The workstation's real file declares no machineProfile at all. Read through the real
        // parser, so the default the loader applies is the default this pins.
        LiveTestSettings parsed = LiveTestSettings.Parse(
            "{ \"testHubStoreDisplayName\": \"" + Hub + "\", "
            + "\"expectedStoreDisplayNames\": [\"" + Hub + "\", \"" + Identity + "\"], "
            + "\"probeTerm\": \"term\", "
            + "\"subjectOnlyProbe\": { \"storeDisplayName\": \"" + Hub + "\", \"folderPath\": \"Inbox\", "
            + "\"subjectTerm\": \"term\", \"senderFragment\": \"term\" } }");

        Assert.Equal(LiveMachineProfile.Production, parsed.MachineProfile);
        Assert.True(parsed.RefusesEveryWrite);
        Assert.False(LiveStoreWriteGuard.Build(parsed).IsAllowed(Hub, StoreWriteKind.Draft));
    }

    [Fact]
    public void ATestGuest_KeepsItsWriteAccess()
    {
        // The other half of the decision: the guests' settings, rendered Portable, keep the hub's
        // full rights and the identity grant - the live tier's write tests run there and only there.
        LiveTestSettings guest = new()
        {
            MachineProfile = LiveMachineProfile.Portable,
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Identity, "bystander@example.test" },
            BystanderStoreDisplayNames = new List<string> { "bystander@example.test" },
        };

        StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(guest);

        Assert.False(allowlist.RefusesEveryWrite);
        Assert.All(Enum.GetValues<StoreWriteKind>(), kind => Assert.True(allowlist.IsAllowed(Hub, kind)));
        Assert.True(allowlist.IsAllowed(Identity, StoreWriteKind.Draft));
        Assert.False(allowlist.IsAllowed("bystander@example.test", StoreWriteKind.Draft));
    }

    [Fact]
    public void AReadOnlyAllowlist_StillKnowsWhichStoreIsWhich_AndStillRefusesAContradiction()
    {
        // The tripwire's soundness check and the artifact sweep ask the allowlist WHICH kind of store
        // each one is; a read-only machine must still be able to answer, or their messages go blank.
        StoreWriteAllowlist allowlist = StoreWriteAllowlist.RefusingEveryWrite(
            "unit reason", Hub, new[] { Identity }, new[] { DelegateStore }, new[] { "bystander@example.test" });

        Assert.True(allowlist.IsHub(Hub));
        Assert.True(allowlist.IsKnownReadOnly(DelegateStore));
        Assert.True(allowlist.IsBystander("bystander@example.test"));
        Assert.Contains("unit reason", allowlist.Explain(Hub, StoreWriteKind.Send, "unit"), StringComparison.Ordinal);

        Assert.Throws<ArgumentException>(() => StoreWriteAllowlist.RefusingEveryWrite(
            "unit reason", Hub, identityDraftStores: new[] { DelegateStore }, knownReadOnlyStores: new[] { DelegateStore }));
        Assert.Throws<ArgumentException>(() => StoreWriteAllowlist.RefusingEveryWrite(" ", Hub));
    }

    [Fact]
    public void OnTheReadOnlyWorkstation_TheTripwirePolicesEveryStore_TheHubIncluded()
    {
        // A consequence worth pinning, because it is a strengthening rather than a side effect: the
        // census used to treat the two other primaries as stores the suite "may still write to", so
        // a change there could be the suite's own. Now nothing is writable, so they are policed - and
        // since 2026-10-03 the hub too, which on a read-only machine nothing exempts
        // (LiveStoreCountTripwire.ExemptHub). The workstation runs no live test since Q116 (a); the
        // rule is the read-only machine's, and the Exchange VM is where it now matters.
        LiveTestSettings settings = Workstation();
        TripwireWatchReport report = TripwireWatchSoundness.Assess(
            LiveStoreCountTripwire.WatchedStores(settings), LiveStoreWriteGuard.Build(settings), settings.BystanderStoreDisplayNames);

        Assert.Empty(report.Writable);
        Assert.Equal(new[] { Hub, Identity, SecondPrimary, DelegateStore, "Another Person" }, report.Policed);
        Assert.Null(report.Refusal());
    }
}
