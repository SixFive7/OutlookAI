using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins <see cref="LiveMailboxIdentity.Judge"/>: the Exchange VM's test mailbox by identity (2026-11-17). The
/// profile there holds that mailbox twice - its own cached store and an online entry automapped through the
/// bystander's Full Access, both with one <c>PR_STORE_RECORD_KEY</c> - and the tier may run only while the hub's
/// name is the test account's own store, no watched store or bystander is the same mailbox under another entry,
/// and no two watched stores are one mailbox.
/// </summary>
public sealed class MailboxIdentityTests
{
    private const string Hub = "telefonie-ai-test@xxlnet.nl";
    private const string Bystander = "telefonie@xxlnet.nl";
    private const string Alias = "Telefonie XXLNet";
    private const string HubKey = "0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C00000000";
    private const string BystanderKey = "0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C00000001";

    /// <summary>The VM as measured on 2026-11-17: the bystander's primary, the hub, and the hub automapped.</summary>
    private static MailboxIdentityReading Measured(
        string? hubEntryKey = HubKey,
        int? hubEntryType = 4,
        string? deliveryKey = HubKey,
        string? deliveryName = Hub,
        string aliasKey = HubKey)
    {
        return new MailboxIdentityReading(
            new[]
            {
                new MailboxStoreEntry(Bystander, 0, true, BystanderKey),
                new MailboxStoreEntry(Hub, hubEntryType, true, hubEntryKey),
                new MailboxStoreEntry(Alias, 1, false, aliasKey),
            },
            new[]
            {
                new MailboxAccountEntry(Bystander, Bystander, BystanderKey),
                new MailboxAccountEntry(Hub, deliveryName, deliveryKey),
            });
    }

    private static readonly string[] Watched = { Hub, Bystander };
    private static readonly string[] Bystanders = { Bystander };

    [Fact]
    public void TheVmAsMeasured_PassesWithTheAutomappedEntryAsAnAliasOfTheHub()
    {
        MailboxIdentityReport report = LiveMailboxIdentity.Judge(Measured(), Hub, Watched, Bystanders);

        Assert.Null(report.Refusal);
        MailboxStoreEntry alias = Assert.Single(report.HubAliases);
        Assert.Equal(Alias, alias.DisplayName);
        Assert.Equal(2, report.WatchedMailboxes);
        Assert.Equal(LiveMailboxIdentity.Short(HubKey), report.HubKeyShort);
        Assert.Contains("'" + Alias + "' (ExchangeStoreType 1, online)", report.Describe(), StringComparison.Ordinal);
        Assert.Contains("never censused as a second one, never written through", report.Describe(), StringComparison.Ordinal);
        Assert.DoesNotContain(HubKey, report.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void WatchingTheAutomappedEntry_IsRefused_OneMailboxWouldCountAsTwo()
    {
        MailboxIdentityReport report = LiveMailboxIdentity.Judge(Measured(), Hub, new[] { Hub, Bystander, Alias }, Bystanders);

        Assert.NotNull(report.Refusal);
        Assert.StartsWith(LiveMailboxIdentity.Refused, report.Refusal!, StringComparison.Ordinal);
        Assert.Contains("one mailbox counted as two stores", report.Refusal!, StringComparison.Ordinal);
    }

    [Fact]
    public void DeclaringTheAutomappedEntryABystander_IsRefused()
    {
        MailboxIdentityReport report = LiveMailboxIdentity.Judge(Measured(), Hub, Watched, new[] { Bystander, Alias });

        Assert.Contains("declared a bystander, but it is the hub's mailbox", report.Refusal ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void NamingTheAutomappedEntryAsTheHub_IsRefused_ItIsNoAccountsOwnStore()
    {
        MailboxIdentityReport report = LiveMailboxIdentity.Judge(Measured(), Alias, new[] { Alias, Bystander }, Bystanders);

        Assert.Contains("0 accounts of the profile are named", report.Refusal ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void AHubEntryThatIsNotTheMailboxTheAccountDeliversTo_IsRefused()
    {
        MailboxIdentityReport report = LiveMailboxIdentity.Judge(Measured(hubEntryKey: BystanderKey), Hub, Watched, Bystanders);

        Assert.Contains("is not the mailbox account", report.Refusal ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void AnAccountThatDeliversToAnotherEntry_IsRefused()
    {
        MailboxIdentityReport report = LiveMailboxIdentity.Judge(Measured(deliveryName: Alias), Hub, Watched, Bystanders);

        Assert.Contains("delivers to the store entry '" + Alias + "'", report.Refusal ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void AHubEntryThatIsADelegateEntry_IsRefused()
    {
        MailboxIdentityReport report = LiveMailboxIdentity.Judge(Measured(hubEntryType: LiveMailboxIdentity.DelegateEntry), Hub, Watched, Bystanders);

        Assert.Contains("is a delegate entry", report.Refusal ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void AnUnreadableDeliveryRecordKey_IsRefused_NeverReadAsAMatch()
    {
        MailboxIdentityReport report = LiveMailboxIdentity.Judge(Measured(deliveryKey: null), Hub, Watched, Bystanders);

        Assert.Contains("could not be read", report.Refusal ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void TwoWatchedStoresThatAreOneMailbox_AreRefused()
    {
        MailboxIdentityReading reading = new(
            new[]
            {
                new MailboxStoreEntry(Bystander, 0, true, BystanderKey),
                new MailboxStoreEntry("Archive copy", 1, false, BystanderKey),
                new MailboxStoreEntry(Hub, 4, true, HubKey),
            },
            new[] { new MailboxAccountEntry(Hub, Hub, HubKey) });

        MailboxIdentityReport report = LiveMailboxIdentity.Judge(reading, Hub, new[] { Hub, Bystander, "Archive copy" }, Bystanders);

        Assert.Contains("are one mailbox - the tripwire would count it twice", report.Refusal ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void AWatchedStoreTheProfileDoesNotHold_IsRefused()
    {
        MailboxIdentityReport report = LiveMailboxIdentity.Judge(Measured(), Hub, new[] { Hub, Bystander, "gone@example.invalid" }, Bystanders);

        Assert.Contains("0 store entries are named 'gone@example.invalid'", report.Refusal ?? string.Empty, StringComparison.Ordinal);
    }

    [Fact]
    public void TheNamesAreComparedWithoutCase_AsOutlookDoes()
    {
        MailboxIdentityReport report = LiveMailboxIdentity.Judge(Measured(), Hub.ToUpperInvariant(), new[] { Hub.ToUpperInvariant(), Bystander }, Bystanders);

        Assert.Null(report.Refusal);
    }

    [Theory]
    [InlineData(LiveMachineProfile.ExchangeTestAccount, true)]
    [InlineData(LiveMachineProfile.ExchangeGuest, true)]
    [InlineData(LiveMachineProfile.Portable, false)]
    [InlineData(LiveMachineProfile.Production, false)]
    public void TheCheckRunsOnTheExchangeVmsProfilesOnly(LiveMachineProfile profile, bool applies)
    {
        Assert.Equal(applies, LiveMailboxIdentity.AppliesTo(profile));
    }
}
