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
/// </summary>
[Collection(LiveCollections.Phase3)]
[Trait("Category", "Live")]
[Trait("LiveTier", "Portable")]
[Trait("Requires", "InteractiveDesktop")]
public sealed class LiveUiSearchBackendTests
{
    private const string NoMatchQuery = "OutlookAiMcpNoSuchTerm7391";

    private readonly LivePhase3Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveUiSearchBackendTests(LivePhase3Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    public void FlippingUserHiveValue_DrivesAdviceAndHealthField_BothStates()
    {
        // The user-hive flip only controls the EFFECTIVE state while no policy-hive
        // value exists (policy is authoritative by design). No such policy exists on
        // this machine; guard so a future GPO turns this into a clear skip, not a red.
        int? policyValue = ReadDword(HealthReporting.OutlookSearchPolicyKeyPath);
        if (policyValue.HasValue)
        {
            _output.WriteLine($"SKIP: policy-hive DisableServerAssistedSearch={policyValue} exists - user-hive flips cannot exercise both states.");
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
        return key?.GetValue("DisableServerAssistedSearch") as int?;
    }

    private static void WriteDword(string keyPath, int value)
    {
        using RegistryKey key = Registry.CurrentUser.CreateSubKey(keyPath)
            ?? throw new InvalidOperationException("cannot open " + keyPath);
        key.SetValue("DisableServerAssistedSearch", value, RegistryValueKind.DWord);
    }

    private static void RestoreDword(string keyPath, int? original)
    {
        if (original.HasValue)
        {
            WriteDword(keyPath, original.Value);
            return;
        }

        using RegistryKey? key = Registry.CurrentUser.OpenSubKey(keyPath, writable: true);
        key?.DeleteValue("DisableServerAssistedSearch", throwOnMissingValue: false);
    }
}
