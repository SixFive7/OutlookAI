using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Shared state for the Outlook-lifecycle live tier (soak-fix batch 2026-07-23: D33
/// headless-first guarantee + SF-1/SF-2 disconnect handling). One MailService like
/// every other live collection; the disconnect-recovery test additionally quits a
/// HEADLESS Outlook through the graceful protocol (S7 v2 - safety counts first,
/// Application.Quit only, never kill) and proves the gateway reattaches. Hub-only
/// artifacts tagged with this run's marker, deleted on dispose (S3).
/// </summary>
public sealed class LiveLifecycleFixture : IDisposable
{
    public LiveLifecycleFixture()
    {
        Settings = LiveTestSettings.Load();

        // Fail-closed per-store count tripwire: no census, no live tier. Cheap after
        // the first fixture (one process-wide baseline).
        LiveStoreCountTripwire.EnsureBaseline(Settings);
        Service = MailService.CreateDefault();
        RunMarker = "lc" + Guid.NewGuid().ToString("N")[..12];
    }

    public LiveTestSettings Settings { get; }

    public MailService Service { get; }

    /// <summary>Unique per-run marker for the S3 double-match delete filter.</summary>
    public string RunMarker { get; }

    public string Hub => Settings.TestHubStoreDisplayName;

    public void Dispose()
    {
        try
        {
            // S3: only artifacts carrying tag + this run's marker; hub store only.
            // Stable-zero loop (Phase-4 fact 6): a just-saved draft can become
            // enumerable only seconds AFTER a one-shot delete pass ran - live-bitten
            // by this exact lag in this batch's first full-suite run.
            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(
                Hub, RunMarker, window: TimeSpan.FromSeconds(60), stableFor: TimeSpan.FromSeconds(10));
        }
        catch (Exception)
        {
            // Cleanup is re-attempted by the per-test paths; never mask a test failure.
        }

        try
        {
            Service.Dispose();
        }
        finally
        {
            // LAST collection in the suite (SuiteCollectionOrderer forces it): re-count
            // every watched store and fail loudly on any loss outside the hub. Outside the
            // try/catch above on purpose - a tripwire that can be swallowed is not one.
            LiveStoreCountTripwire.Verify();
        }
    }
}

[CollectionDefinition("LiveLifecycle")]
public sealed class LiveLifecycleCollection : ICollectionFixture<LiveLifecycleFixture>
{
}
