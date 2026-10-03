using Microsoft.Win32;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// D35 (soak fix 5) live acceptance: flipping the user-hive DisableServerAssistedSearch
/// value exercises BOTH states of the show_search_results advice and health's
/// tuning.uiSearchBackend. The flip touches a PRODUCT-OWNED tuning value (the D24 Search
/// group writes this exact value; in-test modification is within the D24 scope) and the
/// original value is restored in a finally - and even if that failed, the add-in's
/// startup reconcile would re-write the desired value on the next Outlook boot. The UI
/// the show calls drive is parked on the test-hub store with a no-match query (S2/S5:
/// nothing but an empty result list ever appears).
/// <para>
/// <b>A policy value no longer passes it (Q101, 2026-10-03).</b> Where a policy-hive value
/// exists, the user-hive flip cannot reach both states, and the test used to print <c>SKIP:</c>
/// and return GREEN. It goes through <see cref="LivePopulationCoverage"/> now: a refusal on a
/// Production profile, a <c>PROVED NOTHING:</c> line on a Portable one. Pinned by
/// <c>T1/LiveEarlyReturnGuardTests</c>.
/// </para>
/// </summary>
[Collection(LiveCollections.Phase3)]
[Trait("Category", "Live")]
public sealed class LiveUiSearchBackendTests
{
    /// <summary>What the two-state flip needs, named as the Production refusal wraps it.</summary>
    internal const string UserHiveInControlPopulation =
        "a user-hive DisableServerAssistedSearch that no policy-hive value overrides";

    /// <summary>What a reader of a PROVED NOTHING line here is to do about it.</summary>
    internal const string UserHiveInControlRemedy =
        "A DisableServerAssistedSearch value under the Outlook Search key in the POLICY hive "
        + "(HKCU\\Software\\Policies\\Microsoft\\Office\\<version>\\Outlook\\Search) is authoritative over the user-hive "
        + "value this test flips, so the flip cannot reach both states. Nothing in the add-in writes it: find what did "
        + "(Group Policy, a setup script) and remove it, or run on a machine without it.";

    private const string NoMatchQuery = "OutlookAiMcpNoSuchTerm7391";

    private readonly LivePhase3Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveUiSearchBackendTests(LivePhase3Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    [Trait("Requires", "InteractiveDesktop")]
    [Trait("Requires", "AddInRegistry")]
    public void FlippingUserHiveValue_DrivesAdviceAndHealthField_BothStates()
    {
        // The user-hive flip only controls the EFFECTIVE state while no policy-hive
        // value exists (policy is authoritative by design). A policy value used to turn this
        // into a SKIP line and a green result; it is the Q57 pattern now (Q101): a refusal on a
        // Production profile, a PROVED NOTHING line on a Portable one.
        int? policyValue = ReadDword(HealthReporting.OutlookSearchPolicyKeyPath);
        if (policyValue.HasValue)
        {
            _output.WriteLine($"policy-hive DisableServerAssistedSearch={policyValue} is set and overrides the user hive.");
        }

        IReadOnlyList<string> flippable = LivePopulationCoverage.Require(
            _fixture.Settings,
            policyValue.HasValue ? Array.Empty<string>() : new[] { HealthReporting.OutlookSearchUserKeyPath },
            UserHiveInControlPopulation,
            "the two-state user-hive flip",
            UserHiveInControlRemedy,
            _output.WriteLine);
        if (flippable.Count == 0)
        {
            return;
        }

        // Park the Explorer on the hub store first so the driven search UI shows hub
        // content only (an empty list for the no-match query).
        _fixture.Service.GotoFolder(_fixture.Settings.TestHubStoreDisplayName);

        int? original = ReadDword(HealthReporting.OutlookSearchUserKeyPath);
        _output.WriteLine($"original user-hive DisableServerAssistedSearch: {(original.HasValue ? original.Value.ToString() : "(absent)")}");
        try
        {
            // --- State 1: server-assisted (value 0) -> advice present, health reports it.
            WriteDword(HealthReporting.OutlookSearchUserKeyPath, 0);
            Assert.Equal(
                HealthReporting.UiSearchBackendServerAssisted,
                HealthReporting.ReadUiSearchBackendFromRegistry());

            ShowSearchResultsOutcome shownServerAssisted = _fixture.Service.ShowSearchResults(
                NoMatchQuery, "current_folder", _fixture.Settings.TestHubStoreDisplayName);
            Assert.True(shownServerAssisted.Displayed);
            Assert.NotNull(shownServerAssisted.Advice);
            string note = Assert.Single(shownServerAssisted.Advice!);
            Assert.Equal(MailService.ServerAssistedUiSearchAdvice, note);

            HealthOutcome healthServerAssisted = _fixture.Service.Health();
            Assert.Equal("server-assisted", healthServerAssisted.Tuning.UiSearchBackend);
            _output.WriteLine("state server-assisted: advice present (exact wording) + health uiSearchBackend=server-assisted");

            _fixture.VerifySession.TryClearSearch(out _);

            // --- State 2: local (value 1) -> no advice, health reports local.
            WriteDword(HealthReporting.OutlookSearchUserKeyPath, 1);
            Assert.Equal(
                HealthReporting.UiSearchBackendLocal,
                HealthReporting.ReadUiSearchBackendFromRegistry());

            ShowSearchResultsOutcome shownLocal = _fixture.Service.ShowSearchResults(
                NoMatchQuery, "current_folder", _fixture.Settings.TestHubStoreDisplayName);
            Assert.True(shownLocal.Displayed);
            Assert.Null(shownLocal.Advice);

            HealthOutcome healthLocal = _fixture.Service.Health();
            Assert.Equal("local", healthLocal.Tuning.UiSearchBackend);
            _output.WriteLine("state local: no advice + health uiSearchBackend=local");
        }
        finally
        {
            RestoreDword(HealthReporting.OutlookSearchUserKeyPath, original);
            _fixture.VerifySession.TryClearSearch(out _);
            _output.WriteLine($"restored user-hive DisableServerAssistedSearch to {(original.HasValue ? original.Value.ToString() : "(absent)")}");
        }
    }

    private static int? ReadDword(string keyPath)
    {
        using RegistryKey? key = Registry.CurrentUser.OpenSubKey(keyPath, writable: false);
        return key?.GetValue(HealthReporting.DisableServerAssistedSearchValueName) as int?;
    }

    private static void WriteDword(string keyPath, int value)
    {
        using RegistryKey key = Registry.CurrentUser.CreateSubKey(keyPath)
            ?? throw new InvalidOperationException("cannot open " + keyPath);
        key.SetValue(HealthReporting.DisableServerAssistedSearchValueName, value, RegistryValueKind.DWord);
    }

    private static void RestoreDword(string keyPath, int? original)
    {
        if (original.HasValue)
        {
            WriteDword(keyPath, original.Value);
            return;
        }

        using RegistryKey? key = Registry.CurrentUser.OpenSubKey(keyPath, writable: true);
        key?.DeleteValue(HealthReporting.DisableServerAssistedSearchValueName, throwOnMissingValue: false);
    }
}
