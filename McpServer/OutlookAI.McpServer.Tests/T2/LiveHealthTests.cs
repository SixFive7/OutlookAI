using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Phase-7 health tool against the real machine (rides the Phase-2 fixture: one
/// MailService, Outlook started once per S7/D17 and never stopped). Asserts the
/// concrete healthy-machine expectations: classic Outlook build present, all account
/// stores reachable, index provider live, WSearch automatic + indexer running, audit
/// writable, tuning managed by the Phase-6 add-in state (S4: assertions are
/// content-free - counts, flags and version strings only).
/// </summary>
[Collection("LivePhase2")]
[Trait("Category", "Live")]
public sealed class LiveHealthTests
{
    private readonly LivePhase2Fixture _fixture;

    public LiveHealthTests(LivePhase2Fixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void Health_OnThisMachine_ReportsOkWithFullDetail()
    {
        // Ensure the COM session exists so Outlook is definitely up (the fixture may
        // run this test first; ListAccounts attaches/starts per D17).
        _ = _fixture.Service.ListAccounts();

        HealthOutcome report = _fixture.Service.Health();

        // Outlook block: running, classic 16.x build, mutex free, session connected.
        Assert.True(report.Outlook.Running);
        Assert.NotNull(report.Outlook.Version);
        Assert.StartsWith("16.", report.Outlook.Version, StringComparison.Ordinal);
        Assert.False(report.Outlook.InstallerMutexHeld);
        Assert.True(report.Outlook.ComConnected);
        Assert.NotNull(report.Outlook.StoresReachable);
        Assert.True(report.Outlook.StoresReachable >= _fixture.Settings.ExpectedStoreDisplayNames.Count,
            $"expected at least {_fixture.Settings.ExpectedStoreDisplayNames.Count} stores, health saw {report.Outlook.StoresReachable}");
        Assert.NotNull(report.Outlook.Stores);
        foreach (string expected in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            Assert.Contains(expected, report.Outlook.Stores!, StringComparer.OrdinalIgnoreCase);
        }

        // Index block: live provider (not the unavailable marker), staleness populated,
        // WSearch service automatic with the indexer process up.
        Assert.DoesNotContain("unavailable", report.Index.Provider, StringComparison.OrdinalIgnoreCase);
        Assert.NotNull(report.Index.NewestIndexedUtc);
        Assert.NotNull(report.Index.AgeMinutes);
        Assert.Equal("automatic", report.Index.WSearchStartMode);
        Assert.True(report.Index.IndexerProcessRunning);

        // Audit block: writable at the shared-state path.
        Assert.True(report.Audit.Writable);
        Assert.EndsWith("audit.log", report.Audit.Path, StringComparison.OrdinalIgnoreCase);
        Assert.Null(report.Audit.Error);

        // Tuning block: Phase 6 initialized the registry state on this machine.
        Assert.True(report.Tuning.Managed);
        Assert.True(report.Tuning.Enabled);
        Assert.NotNull(report.Tuning.LastReconcileUtc);

        // Whole-machine verdict.
        Assert.True(report.Status == "ok",
            "expected status=ok, got '" + report.Status + "' with problems: "
            + string.Join(" | ", report.Problems ?? Array.Empty<string>()));
        Assert.Null(report.Problems);
    }
}
