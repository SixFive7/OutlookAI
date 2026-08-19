using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the rule that decides WHERE a live run gets verified.
/// <para>
/// The defect it replaces: the store-count tripwire took its baseline in every live
/// collection fixture and compared it in one, on the strength of that collection being
/// forced last. Filtered runs - which is how the tier is meant to be used on a test machine,
/// and what the session log hands the maintainer for the two probes - contain no such
/// collection, so they paid for a census and never compared it, and reported green.
/// </para>
/// <para>
/// Pure logic with the run's collection list injected, so CI pins every branch without xunit
/// having to execute a live tier.
/// </para>
/// </summary>
public sealed class LiveTierRunPlanTests
{
    [Fact]
    public void TheLastGuardedCollectionInTheRun_IsWhereVerificationBelongs()
    {
        string[] ordered = { LiveCollections.Phase2, LiveCollections.Phase3, LiveCollections.Lifecycle };

        Assert.Equal(GuardedCollectionPosition.NotLast, LiveTierRunPlan.PositionIn(ordered, LiveCollections.Phase2));
        Assert.Equal(GuardedCollectionPosition.NotLast, LiveTierRunPlan.PositionIn(ordered, LiveCollections.Phase3));
        Assert.Equal(GuardedCollectionPosition.Last, LiveTierRunPlan.PositionIn(ordered, LiveCollections.Lifecycle));
    }

    [Fact]
    public void AFilteredRunOfOneCollection_VerifiesInThatCollection()
    {
        // This is the whole point: --filter "FullyQualifiedName~LiveTableSortProbeTests"
        // selects LivePhase3 alone, and LiveLifecycle - where the only Verify call used to
        // live - never runs at all.
        string[] ordered = { LiveCollections.Phase3 };

        Assert.Equal(GuardedCollectionPosition.Last, LiveTierRunPlan.PositionIn(ordered, LiveCollections.Phase3));
    }

    [Fact]
    public void UnguardedCollectionsAfterTheLastGuardedOne_DoNotDelayVerification()
    {
        // Non-live collections may sort after a live one. They take no baseline and dispose
        // nothing the tripwire hears about, so waiting for them would mean waiting forever.
        string[] ordered = { LiveCollections.Phase3, "Test collection for Something.Unrelated" };

        Assert.Equal(GuardedCollectionPosition.Last, LiveTierRunPlan.PositionIn(ordered, LiveCollections.Phase3));
    }

    [Fact]
    public void NoPublishedPlan_IsUnknownRatherThanAssumedLast()
    {
        // Unknown means "verify now and stay armed", which costs a census per collection and
        // is the deliberate trade: an unverified run is the outcome that must not happen.
        Assert.Equal(GuardedCollectionPosition.Unknown, LiveTierRunPlan.PositionIn(null, LiveCollections.Phase1));
        Assert.Equal(
            GuardedCollectionPosition.Unknown,
            LiveTierRunPlan.PositionIn(Array.Empty<string>(), LiveCollections.Phase1));
    }

    [Fact]
    public void ACollectionMissingFromThePlan_IsUnknownRatherThanLast()
    {
        string[] ordered = { LiveCollections.Phase2, LiveCollections.Lifecycle };

        Assert.Equal(GuardedCollectionPosition.Unknown, LiveTierRunPlan.PositionIn(ordered, LiveCollections.Phase4));
    }

    [Fact]
    public void PublishThenAsk_AnswersFromTheProcessWidePlan()
    {
        try
        {
            LiveTierRunPlan.Publish(new[] { LiveCollections.Phase1, LiveCollections.Phase5 });

            Assert.Equal(GuardedCollectionPosition.NotLast, LiveTierRunPlan.PositionOf(LiveCollections.Phase1));
            Assert.Equal(GuardedCollectionPosition.Last, LiveTierRunPlan.PositionOf(LiveCollections.Phase5));
        }
        finally
        {
            LiveTierRunPlan.ResetForTests();
        }
    }

    [Fact]
    public void EveryRegisteredCollectionName_IsRecognisedAsGuarded()
    {
        Assert.All(LiveCollections.All, name => Assert.True(LiveCollections.IsGuarded(name), name));
        Assert.False(LiveCollections.IsGuarded("Test collection for OutlookAI.McpServer.Tests.T3.Something"));
        Assert.False(LiveCollections.IsGuarded(null));
    }
}
