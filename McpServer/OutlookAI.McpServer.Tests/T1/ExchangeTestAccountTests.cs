using OutlookAI.McpServer.Tests.T2;
using OutlookAI.McpServer.Tests.T3;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins <see cref="LiveMachineProfile.ExchangeTestAccount"/> (2026-10-08): the Exchange VM writing in its
/// dedicated test account and nowhere else - the hub writable, every other store a read-only bystander,
/// mail addressed to the hub alone. Pure; no Outlook, no settings file.
/// </summary>
public sealed class ExchangeTestAccountTests
{
    private const string Hub = "test-account@example.invalid";
    private const string RealMailbox = "real-mailbox@example.invalid";

    private static LiveTestSettings Settings() => new()
    {
        MachineProfile = LiveMachineProfile.ExchangeTestAccount,
        TestHubStoreDisplayName = Hub,
        ExpectedStoreDisplayNames = new List<string> { Hub, RealMailbox },
        IndexedStoreDisplayNames = new List<string> { Hub, RealMailbox },
        BystanderStoreDisplayNames = new List<string> { RealMailbox },
    };

    [Fact]
    public void ItMayWrite_ArmsTheRecipientAllowlist_AndItsMcpClientIsHubOnly()
    {
        Assert.False(LiveWriteAccess.RefusesEveryWrite(LiveMachineProfile.ExchangeTestAccount));
        Assert.True(LiveWriteAccess.ArmsRecipientAllowlist(LiveMachineProfile.ExchangeTestAccount));
        foreach (LiveMachineProfile other in new[] { LiveMachineProfile.Production, LiveMachineProfile.Portable, LiveMachineProfile.ExchangeGuest, (LiveMachineProfile)93 })
        {
            Assert.False(LiveWriteAccess.ArmsRecipientAllowlist(other), other.ToString());
        }

        Assert.Equal(
            StdioWritePosture.WritableHubOnly,
            LiveWriteAccess.StdioPostureFor(LiveRunOptIn.Verdict.Open, Settings));
        Assert.Equal(
            StdioWritePosture.Writable,
            LiveWriteAccess.StdioPostureFor(LiveRunOptIn.Verdict.Open, () => new LiveTestSettings { MachineProfile = LiveMachineProfile.Portable }));
        Assert.Equal(
            StdioWritePosture.ReadOnly,
            LiveWriteAccess.StdioPostureFor(LiveRunOptIn.Verdict.Open, () => new LiveTestSettings { MachineProfile = LiveMachineProfile.ExchangeGuest }));

        // The classification gate lets a write through under it: the hub-only gate is what judges it.
        Assert.Null(McpToolWriteClassification.RefusalFor(StdioWritePosture.WritableHubOnly, "send"));
    }

    [Fact]
    public void TheWriteGuard_GrantsTheHubAlone_AndNoIdentityDraftAnywhere()
    {
        StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(Settings());

        Assert.False(allowlist.RefusesEveryWrite);
        foreach (StoreWriteKind kind in Enum.GetValues<StoreWriteKind>())
        {
            Assert.True(allowlist.IsAllowed(Hub, kind), "hub " + kind);
            Assert.False(allowlist.IsAllowed(RealMailbox, kind), "real mailbox " + kind);
            Assert.False(allowlist.IsAllowed("anything-else@example.invalid", kind), "undeclared " + kind);
        }

        Assert.Empty(allowlist.IdentityDraftStores);

        // Even a primary left out of the bystander list - which Validate refuses - gets no draft grant here.
        LiveTestSettings careless = Settings();
        careless.BystanderStoreDisplayNames = new List<string>();
        StoreWriteAllowlist carelessAllowlist = LiveStoreWriteGuard.Build(careless);
        Assert.False(carelessAllowlist.IsAllowed(RealMailbox, StoreWriteKind.Draft));
        Assert.Empty(carelessAllowlist.IdentityDraftStores);
    }

    [Fact]
    public void TheSettingsMustDeclareEveryOtherStoreABystander_AndNameTheHubByItsAddress()
    {
        LiveTestSettings.Validate(Settings());

        LiveTestSettings undeclared = Settings();
        undeclared.BystanderStoreDisplayNames = new List<string>();
        Assert.Contains("bystanderStoreDisplayNames", Assert.Throws<InvalidOperationException>(() => LiveTestSettings.Validate(undeclared)).Message, StringComparison.Ordinal);

        LiveTestSettings hubIsBystander = Settings();
        hubIsBystander.BystanderStoreDisplayNames = new List<string> { RealMailbox, Hub };
        Assert.Contains("bystander", Assert.Throws<InvalidOperationException>(() => LiveTestSettings.Validate(hubIsBystander)).Message, StringComparison.Ordinal);

        LiveTestSettings notAnAddress = Settings();
        notAnAddress.TestHubStoreDisplayName = "Test Account";
        notAnAddress.ExpectedStoreDisplayNames = new List<string> { "Test Account", RealMailbox };
        notAnAddress.IndexedStoreDisplayNames = new List<string> { "Test Account", RealMailbox };
        Assert.Contains("not an address", Assert.Throws<InvalidOperationException>(() => LiveTestSettings.Validate(notAnAddress)).Message, StringComparison.Ordinal);

        LiveTestSettings withSink = Settings();
        withSink.MailSink = new MailSinkSettings { SubmitHost = "127.0.0.1", SubmitPort = 25, RetrieveHost = "127.0.0.1", RetrievePort = 110 };
        Assert.Contains("mailSink", Assert.Throws<InvalidOperationException>(() => LiveTestSettings.Validate(withSink)).Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheProfileParsesFromASettingsFile()
    {
        LiveTestSettings parsed = LiveTestSettings.Parse(
            "{ \"machineProfile\": \"ExchangeTestAccount\", \"testHubStoreDisplayName\": \"" + Hub + "\", "
            + "\"expectedStoreDisplayNames\": [\"" + Hub + "\", \"" + RealMailbox + "\"], "
            + "\"bystanderStoreDisplayNames\": [\"" + RealMailbox + "\"] }");

        Assert.Equal(LiveMachineProfile.ExchangeTestAccount, parsed.MachineProfile);
        Assert.False(parsed.RefusesEveryWrite);
        Assert.False(parsed.AUserMayBeAtTheKeyboard);
        Assert.Contains("writes=hub only, recipients the hub only", parsed.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheRealMailboxStaysCensused_TheHubIsExempt_AndARealMailboxGetsReCensusesNeverAReRun()
    {
        LiveTestSettings settings = Settings();

        Assert.Equal(Hub, LiveStoreCountTripwire.ExemptHub(settings));
        Assert.Contains(RealMailbox, LiveStoreCountTripwire.WatchedStores(settings));

        TripwireRetryPolicy policy = TripwireRetryPolicy.For(LiveMachineProfile.ExchangeTestAccount);
        Assert.Same(TripwireRetryPolicy.ExchangeGuest, policy);
        Assert.Equal(0, policy.MaxImplicatedReRuns);
        Assert.True(policy.MaxReCensuses > 0);
    }

    [Fact]
    public void TheRecipientGuard_IsArmedWithTheHubsAddress_AndOnlyThere()
    {
        RecipientAllowlist? armed = LiveRecipientGuard.For(Settings());
        Assert.NotNull(armed);
        Assert.Equal(new[] { Hub }, armed!.Allowed);

        Assert.Null(LiveRecipientGuard.For(new LiveTestSettings { MachineProfile = LiveMachineProfile.Portable, TestHubStoreDisplayName = "tier@vm.invalid" }));
        Assert.Null(LiveRecipientGuard.For(new LiveTestSettings { MachineProfile = LiveMachineProfile.ExchangeGuest, TestHubStoreDisplayName = Hub }));
    }
}
