using OutlookAI.McpServer.Tests.T2;
using Xunit;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the collection orderer, which does two jobs that nothing else can do.
/// <para>
/// It decides the ORDER (the Outlook-bouncing collection last, and the stdio shape tests kept
/// where they already ran), and it PUBLISHES the shape of this run to
/// <see cref="LiveTierRunPlan"/>. The second is load-bearing: it is the only vantage point in
/// the process that sees which collections survived the command line's filter, and without it
/// a filtered live run takes a store census and never compares it.
/// </para>
/// </summary>
public sealed class SuiteCollectionOrdererTests
{
    /// <summary>
    /// Minimal <see cref="ITestCollection"/>: the orderer reads nothing but the display name,
    /// and a stub keeps this test free of xunit's discovery machinery.
    /// </summary>
    private sealed class Collection : LongLivedMarshalByRefObject, ITestCollection
    {
        public ITypeInfo? CollectionDefinition => null;

        // Settable, and the class needs a public parameterless constructor, because xunit's
        // analyzers require both of every ITestCollection - it is a type the framework may
        // serialise across an app-domain boundary.
        public string DisplayName { get; set; } = string.Empty;

        public ITestAssembly TestAssembly => null!;

        public Guid UniqueID { get; } = Guid.NewGuid();

        public void Deserialize(IXunitSerializationInfo info)
        {
        }

        public void Serialize(IXunitSerializationInfo info)
        {
        }
    }

    private static List<string> Order(params string[] displayNames)
    {
        return new SuiteCollectionOrderer()
            .OrderTestCollections(displayNames.Select(n => (ITestCollection)new Collection { DisplayName = n }))
            .Select(c => c.DisplayName)
            .ToList();
    }

    [Fact]
    public void TheOutlookBouncingCollectionRunsLast_AndTheStdioShapeTestsStayLate()
    {
        try
        {
            List<string> ordered = Order(
                LiveCollections.Lifecycle,
                LiveCollections.McpToolShape,
                LiveCollections.Phase2,
                LiveCollections.Phase1);

            Assert.Equal(
                new[]
                {
                    LiveCollections.Phase1,
                    LiveCollections.Phase2,
                    LiveCollections.McpToolShape,
                    LiveCollections.Lifecycle,
                },
                ordered);
        }
        finally
        {
            LiveTierRunPlan.ResetForTests();
        }
    }

    [Fact]
    public void OrderingPublishesTheRunPlan_SoAFilteredRunKnowsWhereItEnds()
    {
        try
        {
            LiveTierRunPlan.ResetForTests();
            Order(LiveCollections.Phase3);

            // Without the publish this is Unknown, and the tripwire falls back to verifying at
            // every collection boundary. With it, the one collection the filter left is the
            // one that verifies.
            Assert.Equal(GuardedCollectionPosition.Last, LiveTierRunPlan.PositionOf(LiveCollections.Phase3));
            Assert.Equal(new[] { LiveCollections.Phase3 }, LiveTierRunPlan.Current);
        }
        finally
        {
            LiveTierRunPlan.ResetForTests();
        }
    }

    [Fact]
    public void NonLiveCollections_KeepTheirAlphabeticalPlace()
    {
        try
        {
            List<string> ordered = Order("Zebra", "Alpha", LiveCollections.Lifecycle);

            Assert.Equal(new[] { "Alpha", "Zebra", LiveCollections.Lifecycle }, ordered);
        }
        finally
        {
            LiveTierRunPlan.ResetForTests();
        }
    }
}
