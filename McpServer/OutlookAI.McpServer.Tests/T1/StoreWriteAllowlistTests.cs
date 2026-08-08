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
        // other configured primaries, delegates denied.
        LiveTestSettings settings = new()
        {
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Identity },
            ExpectedDelegateStoreDisplayNames = new List<string> { DelegateStore },
        };

        StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(settings);

        Assert.True(allowlist.IsAllowed(Hub, StoreWriteKind.Send));
        Assert.True(allowlist.IsAllowed(Identity, StoreWriteKind.Draft));
        Assert.False(allowlist.IsAllowed(Identity, StoreWriteKind.Send));
        Assert.False(allowlist.IsAllowed(DelegateStore, StoreWriteKind.Delete));
    }
}
