using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 pins for the registry that makes discard_draft safe (v3.MD D46/C2, S1 v3). This
/// is the whole guardrail behind the product's only mail-deleting tool, so its rules are
/// pinned directly rather than inferred from the live tests: unknown ids are NOT
/// members, an EntryID change re-keys the entry instead of losing it, a discarded draft
/// is forgotten, and nothing leaks in through a blank id.
/// </summary>
public sealed class ServerDraftRegistryTests
{
    [Fact]
    public void UnknownEntryId_IsNotAMember_WhichIsWhatRefusesADiscard()
    {
        ServerDraftRegistry registry = new();

        Assert.False(registry.Contains("AAAA1111"));
        Assert.Equal(0, registry.Count);
    }

    [Fact]
    public void RegisteredDraft_IsAMember_CaseInsensitivelyAndTrimmed()
    {
        ServerDraftRegistry registry = new();
        registry.Register("ABC123");

        Assert.True(registry.Contains("ABC123"));
        Assert.True(registry.Contains("abc123"));
        Assert.True(registry.Contains("  ABC123  "));
        Assert.Equal(1, registry.Count);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void BlankIds_AreNeverRegisteredAndNeverMatch(string? id)
    {
        // A snapshot that could not read an EntryID must not widen the allowlist, and a
        // blank id must never be accepted as "in the registry".
        ServerDraftRegistry registry = new();
        registry.Register(id);

        Assert.Equal(0, registry.Count);
        Assert.False(registry.Contains(id));
    }

    [Fact]
    public void RegisteringTwice_KeepsOneEntry()
    {
        ServerDraftRegistry registry = new();
        registry.Register("ABC");
        registry.Register("abc");

        Assert.Equal(1, registry.Count);
    }

    [Fact]
    public void Replace_ReKeysTheEntry_BecauseEntryIdsChangeOnAnyMove()
    {
        ServerDraftRegistry registry = new();
        registry.Register("OLD");

        registry.Replace("OLD", "NEW");

        Assert.False(registry.Contains("OLD"));
        Assert.True(registry.Contains("NEW"));
        Assert.Equal(1, registry.Count);
    }

    [Fact]
    public void Replace_RegistersTheNewIdEvenWhenTheOldOneWasUnknown()
    {
        // An update must never LOSE the right to discard what it just rewrote.
        ServerDraftRegistry registry = new();

        registry.Replace("NEVER-SEEN", "NEW");

        Assert.True(registry.Contains("NEW"));
    }

    [Fact]
    public void Replace_WithAnUnchangedId_KeepsMembership()
    {
        ServerDraftRegistry registry = new();
        registry.Register("SAME");

        registry.Replace("SAME", "same");

        Assert.True(registry.Contains("SAME"));
        Assert.Equal(1, registry.Count);
    }

    [Fact]
    public void Forget_DropsMembership()
    {
        ServerDraftRegistry registry = new();
        registry.Register("ABC");

        registry.Forget("abc");

        Assert.False(registry.Contains("ABC"));
        Assert.Equal(0, registry.Count);
    }

    [Fact]
    public void OldestRegistrationIsEvictedBeyondCapacity_AndTheNewestSurvive()
    {
        ServerDraftRegistry registry = new();
        for (int i = 0; i < ServerDraftRegistry.Capacity + 10; i++)
        {
            registry.Register("draft-" + i);
        }

        Assert.Equal(ServerDraftRegistry.Capacity, registry.Count);
        Assert.False(registry.Contains("draft-0"));
        Assert.True(registry.Contains("draft-" + (ServerDraftRegistry.Capacity + 9)));
    }

    [Fact]
    public void ReRegisteringRefreshesRecency_SoAnActivelyEditedDraftIsNotEvicted()
    {
        ServerDraftRegistry registry = new();
        registry.Register("keeper");
        for (int i = 0; i < ServerDraftRegistry.Capacity - 1; i++)
        {
            registry.Register("filler-" + i);
        }

        registry.Register("keeper"); // touched again, e.g. by update_draft
        registry.Register("one-more-to-force-eviction");

        Assert.True(registry.Contains("keeper"));
        Assert.False(registry.Contains("filler-0"));
    }
}
