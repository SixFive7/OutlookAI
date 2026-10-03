using System.Globalization;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// T2, read-only, cached Exchange only: Q113 (b) - which documented store-hash input Outlook uses for a
/// CACHED EXCHANGE store, read the one way the product itself reports it: <c>outlook_health</c>'s
/// <c>index.perStore</c> row for the store, its <c>matchedBy</c> and <c>matchedInput</c>
/// (<c>McpServer/README.md</c>, load-bearing fact 16).
/// <para>
/// For a PST the input is measured (<c>entryId</c>). For a cached Exchange store Microsoft documents the
/// profile's <c>PR_MAPPING_SIGNATURE</c>, and the product also tries the store's own signature and the
/// entry-ID-plus-<c>.ost</c> variant - but no PST guest can say which one Outlook actually uses, and
/// agents never query the maintainer's Outlook. The Exchange VM can (decided 2026-10-03, Q113 (b)):
/// <c>storeHash</c> with an input settles it, and then the name fallback is not what finds an Exchange
/// store; <c>displayName</c> means no documented input matched and the store is still found by its name.
/// </para>
/// <para>
/// A DISCOVERY, recorded: the line this test prints is the measurement. It asserts only what any
/// answer must satisfy - the store is in the index and tied to it somehow, and a hash tie names one of
/// the Exchange inputs - so that the answer, whichever it is, comes back as a pass with its line.
/// Content-free (S4): flags, the input's name and the index root's own name.
/// </para>
/// </summary>
[Collection(LiveCollections.Phase2)]
[Trait("Category", "Live")]
public sealed class LiveExchangeStoreHashTests
{
    /// <summary>The wire spellings of the cached-Exchange inputs (MailService.StoreHashInputName).</summary>
    private static readonly string[] ExchangeInputs = { "profileMappingSignature", "mappingSignature", "entryIdAndPath" };

    private readonly LivePhase2Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveExchangeStoreHashTests(LivePhase2Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "CachedExchange")]
    [Trait("Writes", "Nothing")]
    public void OutlookHealth_TiesTheCachedExchangeStoreToItsIndexRoot_AndSaysByWhichInput()
    {
        // Outlook running and attached first: the per-store rows come from the store map only then.
        _ = _fixture.Service.ListAccounts();
        HealthOutcome report = _fixture.Service.Health();

        Assert.True(report.Outlook.Running && report.Outlook.ComConnected, "outlook_health did not see a running, attached Outlook");
        Assert.NotNull(report.Index.PerStore);

        string hub = _fixture.Settings.TestHubStoreDisplayName;
        StoreStaleness? row = report.Index.PerStore!.FirstOrDefault(
            r => string.Equals(r.Store, hub, StringComparison.OrdinalIgnoreCase));
        Assert.True(row != null, "outlook_health has no per-store row for the hub");

        _output.WriteLine("Q113 (b) cached Exchange store: matchedBy=" + (row!.MatchedBy ?? "(null)")
            + " matchedInput=" + (row.MatchedInput ?? "(null)")
            + " inLocalIndex=" + (row.InLocalIndex?.ToString() ?? "(null)")
            + " indexStore=" + (row.IndexStore ?? "(null)")
            + " matchNote=" + (row.MatchNote ?? "(none)")
            + " storesNotInProfile=" + (report.Index.StoresNotInProfile?.Count ?? 0).ToString(CultureInfo.InvariantCulture));

        Assert.True(row.InLocalIndex == true, "the hub is not in the local index");
        Assert.True(row.MatchedBy is "storeHash" or "displayName",
            "the cached Exchange store was tied by '" + row.MatchedBy + "' - expected a store hash or, failing that, its name");
        if (row.MatchedBy == "storeHash")
        {
            Assert.True(ExchangeInputs.Contains(row.MatchedInput),
                "a hash tie on a cached Exchange store named input '" + row.MatchedInput + "', which is not one of the Exchange inputs");
        }
    }
}
