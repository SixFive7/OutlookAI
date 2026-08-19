using OutlookAI.McpServer.Tests.T2;
using Xunit;
using Xunit.Abstractions;

[assembly: TestCollectionOrderer("OutlookAI.McpServer.Tests.SuiteCollectionOrderer", "OutlookAI.McpServer.Tests")]

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// Deterministic collection order (collections already run sequentially per
/// xunit.runner.json): alphabetical, with the LiveLifecycle collection forced LAST.
/// LiveLifecycle's disconnect-recovery test closes and re-autostarts Outlook - doing
/// that mid-suite pauses mail sync long enough to widen the known sent-copy
/// materialization lag (Phase-4 fact 6 family): a fresh-proof Sent Items copy from an
/// earlier collection materialized AFTER that collection's stable-zero cleanup and
/// tripped a later artifact sweep (live-bitten in this batch's confirmation run).
/// Running the Outlook bounce after every other live collection has settled and swept
/// removes that interaction without weakening any S3 guarantee.
/// <para>
/// It is also the only place that knows the SHAPE of the run: xunit hands it the collections
/// that survived the command line's filter, before any fixture is constructed. That is
/// published to <see cref="LiveTierRunPlan"/> so the store-count tripwire can tell which
/// collection ends the run - without it, a filtered run takes a census and never compares it
/// (see <see cref="LiveTierRunPlan"/> for the defect in full).
/// </para>
/// </summary>
public sealed class SuiteCollectionOrderer : ITestCollectionOrderer
{
    /// <inheritdoc />
    public IEnumerable<ITestCollection> OrderTestCollections(IEnumerable<ITestCollection> testCollections)
    {
        // Materialised rather than returned lazily: the run plan has to be published now,
        // and enumerating the query twice would order the same collections twice.
        List<ITestCollection> ordered = testCollections
            .OrderBy(c => Rank(c.DisplayName))
            .ThenBy(c => c.DisplayName, StringComparer.Ordinal)
            .ToList();

        LiveTierRunPlan.Publish(ordered.Select(c => c.DisplayName));
        return ordered;
    }

    /// <summary>
    /// Late-running collections, latest last. Two of them, for different reasons.
    /// <para>
    /// <see cref="LiveCollections.McpToolShape"/> is late to PRESERVE the order it already
    /// had. Those three classes belonged to no collection until they were brought under the
    /// count tripwire, so xunit named their implicit collections "Test collection for ..."
    /// - which sorted after every "Live" name and ran them near the end. Naming a collection
    /// should not silently move nine live tests to the front of a suite whose ordering was
    /// arrived at by being bitten, so the rank puts them back where they were.
    /// </para>
    /// <para>
    /// <see cref="LiveCollections.Lifecycle"/> is last because it bounces Outlook - see the
    /// class remarks.
    /// </para>
    /// </summary>
    private static int Rank(string? displayName)
    {
        if (string.Equals(displayName, LiveCollections.Lifecycle, StringComparison.Ordinal))
        {
            return 2;
        }

        return string.Equals(displayName, LiveCollections.McpToolShape, StringComparison.Ordinal) ? 1 : 0;
    }
}
