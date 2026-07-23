using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pure health-report mapping logic (Phase 7): WSearch start-mode naming and the
/// tuning-state reader (fed by a fabricated value getter here; production feeds it
/// from HKCU\Software\OutlookAI\Tuning - registry layout per v3.MD section 0.8
/// Phase-6 fact 5).
/// </summary>
public sealed class HealthReportingTests
{
    [Theory]
    [InlineData(2, "automatic")]
    [InlineData(3, "manual")]
    [InlineData(4, "disabled")]
    [InlineData(null, "unknown")]
    [InlineData(1, "other(1)")]
    [InlineData(0, "other(0)")]
    public void DescribeServiceStartMode_MapsRegistryValues(int? start, string expected)
    {
        Assert.Equal(expected, HealthReporting.DescribeServiceStartMode(start));
    }

    [Fact]
    public void ReadTuningState_NoInitializedValue_ReportsNotManaged()
    {
        TuningHealthView view = HealthReporting.ReadTuningState(_ => null);

        Assert.False(view.Managed);
        Assert.Null(view.Enabled);
        Assert.Null(view.RestartNeeded);
        Assert.Null(view.PolicyConflicts);
    }

    [Fact]
    public void ReadTuningState_InitializedZero_ReportsNotManaged()
    {
        TuningHealthView view = HealthReporting.ReadTuningState(
            name => name == "Initialized" ? 0 : (object?)1);

        Assert.False(view.Managed);
    }

    [Fact]
    public void ReadTuningState_FullState_MapsAllFields()
    {
        var values = new Dictionary<string, object?>
        {
            ["Initialized"] = 1,
            ["Enabled"] = 1,
            ["SearchEnabled"] = 1,
            ["CachingEnabled"] = 0,
            ["OstEnabled"] = 1,
            ["RestartNeeded"] = 1,
            ["PolicyConflicts"] = "caching.policy.SyncWindowSetting",
            ["LastReconcileUtc"] = "2026-07-23T00:00:00.0000000Z",
        };

        TuningHealthView view = HealthReporting.ReadTuningState(
            name => values.TryGetValue(name, out object? v) ? v : null);

        Assert.True(view.Managed);
        Assert.True(view.Enabled);
        Assert.True(view.SearchEnabled);
        Assert.False(view.CachingEnabled);
        Assert.True(view.OstEnabled);
        Assert.True(view.RestartNeeded);
        Assert.Equal("caching.policy.SyncWindowSetting", view.PolicyConflicts);
        Assert.Equal("2026-07-23T00:00:00.0000000Z", view.LastReconcileUtc);
    }

    [Fact]
    public void ReadTuningState_EmptyPolicyConflicts_ReportsNull()
    {
        TuningHealthView view = HealthReporting.ReadTuningState(name => name switch
        {
            "Initialized" => 1,
            "PolicyConflicts" => "",
            _ => (object?)1,
        });

        Assert.True(view.Managed);
        Assert.Null(view.PolicyConflicts);
    }

    [Fact]
    public void ReadTuningStateFromRegistry_NeverThrows()
    {
        // Machine-state agnostic: on a machine without the add-in the key is absent
        // (Managed=false); with it, Managed=true and the toggles have values. Either
        // way the read must be exception-free (health always produces a report).
        TuningHealthView view = HealthReporting.ReadTuningStateFromRegistry();

        Assert.NotNull(view);
        if (view.Managed)
        {
            Assert.NotNull(view.Enabled);
        }
    }

    [Fact]
    public void MachineProbes_NeverThrow()
    {
        // CI-safe smoke of the impure probes: values are machine-dependent, absence of
        // exceptions is the contract.
        _ = HealthReporting.TryReadWSearchStartValue();
        _ = HealthReporting.TryIsProcessRunning("SearchIndexer");
        _ = HealthReporting.TryGetOutlookVersion();

        string probeDir = Path.Combine(Path.GetTempPath(), "outlookai-health-probe-" + Guid.NewGuid().ToString("N"));
        try
        {
            bool writable = OutlookAI.Core.Audit.AuditLog.TryProbeWritable(probeDir, out string? error);
            Assert.True(writable);
            Assert.Null(error);

            // The probe must not have APPENDED anything - it only opens the handle.
            Assert.Equal(0, new FileInfo(Path.Combine(probeDir, "audit.log")).Length);
        }
        finally
        {
            if (Directory.Exists(probeDir))
            {
                Directory.Delete(probeDir, recursive: true);
            }
        }
    }
}
