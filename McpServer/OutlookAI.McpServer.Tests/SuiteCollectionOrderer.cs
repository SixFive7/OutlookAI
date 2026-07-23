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
/// </summary>
public sealed class SuiteCollectionOrderer : ITestCollectionOrderer
{
    /// <inheritdoc />
    public IEnumerable<ITestCollection> OrderTestCollections(IEnumerable<ITestCollection> testCollections)
    {
        return testCollections
            .OrderBy(c => string.Equals(c.DisplayName, "LiveLifecycle", StringComparison.Ordinal) ? 1 : 0)
            .ThenBy(c => c.DisplayName, StringComparer.Ordinal);
    }
}
