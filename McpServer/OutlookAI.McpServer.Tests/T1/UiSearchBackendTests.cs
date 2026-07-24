using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// D35 (soak fix 5): the DisableServerAssistedSearch coupling made self-documenting.
/// Pins the effective-state mapping (policy hive authoritative, absent/0 =
/// server-assisted - Outlook's default), the never-throw registry read, and the
/// deliberately strong show_search_results advice wording: it must convey BOTH the
/// divergence (UI results server-capped/ranked vs the agent's local search) AND the
/// recommendation to re-enable the Search tuning group.
/// </summary>
public sealed class UiSearchBackendTests
{
    [Theory]
    // Neither hive set: Outlook's default is server-assisted search.
    [InlineData(null, null, "server-assisted")]
    // User hive alone decides when no policy value exists.
    [InlineData(null, 1, "local")]
    [InlineData(null, 0, "server-assisted")]
    // Policy hive present: AUTHORITATIVE, user hive ignored (both directions).
    [InlineData(1, 0, "local")]
    [InlineData(1, null, "local")]
    [InlineData(0, 1, "server-assisted")]
    [InlineData(0, null, "server-assisted")]
    // Any nonzero DWORD counts as enabled (registry-bool semantics).
    [InlineData(null, 2, "local")]
    [InlineData(2, 0, "local")]
    public void DescribeUiSearchBackend_PolicyHiveWins_AbsentMeansServerAssisted(
        int? policyValue, int? userValue, string expected)
    {
        Assert.Equal(expected, HealthReporting.DescribeUiSearchBackend(policyValue, userValue));
    }

    [Fact]
    public void BackendNames_ArePinned()
    {
        // The two wire values of health's tuning.uiSearchBackend (T3 asserts them over
        // stdio; this pins the constants the mapping emits).
        Assert.Equal("local", HealthReporting.UiSearchBackendLocal);
        Assert.Equal("server-assisted", HealthReporting.UiSearchBackendServerAssisted);
    }

    [Fact]
    public void ReadUiSearchBackendFromRegistry_NeverThrows_AndReturnsAKnownBackend()
    {
        // Machine-state agnostic (CI runners have neither hive -> server-assisted; this
        // machine has the tuning applied -> local). The contract here: no exception and
        // one of the two pinned values.
        string backend = HealthReporting.ReadUiSearchBackendFromRegistry();

        Assert.True(
            backend is "local" or "server-assisted",
            $"unexpected backend '{backend}'");
    }

    [Fact]
    public void ReadTuningStateFromRegistry_AlwaysStampsUiSearchBackend()
    {
        // The health tuning block must carry the effective backend even when the add-in
        // never initialized tuning on the machine (Managed=false on CI runners).
        TuningHealthView view = HealthReporting.ReadTuningStateFromRegistry();

        Assert.NotNull(view.UiSearchBackend);
        Assert.True(view.UiSearchBackend is "local" or "server-assisted");
    }

    [Fact]
    public void ServerAssistedAdvice_ConveysDivergenceAndRecommendation()
    {
        string advice = MailService.ServerAssistedUiSearchAdvice;

        // Divergence: the user-visible list is not the agent's list.
        Assert.Contains("server-assisted", advice, StringComparison.Ordinal);
        Assert.Contains("may differ from agent search results", advice, StringComparison.Ordinal);
        Assert.Contains("server-capped", advice, StringComparison.Ordinal);

        // Recommendation: strong wording plus the concrete fix location.
        Assert.Contains("RECOMMENDED", advice, StringComparison.Ordinal);
        Assert.Contains("Search tuning group", advice, StringComparison.Ordinal);
        Assert.Contains("OutlookAI Settings", advice, StringComparison.Ordinal);
    }
}
