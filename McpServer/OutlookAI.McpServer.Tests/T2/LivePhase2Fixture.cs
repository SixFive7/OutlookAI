using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Shared state for the Phase-2 T2 live tier: ONE MailService (one ComGateway, one
/// pumped STA, one held-open Outlook session - may START Outlook per S7/D17, never
/// stops it) reused across all Phase-2 live tests so hit ids stay valid between
/// search and read within the collection.
/// </summary>
public sealed class LivePhase2Fixture : IDisposable
{
    public LivePhase2Fixture()
    {
        Settings = LiveTestSettings.Load();

        // Fail-closed per-store count tripwire: no census, no live tier. Cheap after
        // the first fixture (one process-wide baseline).
        LiveStoreCountTripwire.EnsureBaseline(Settings);
        Service = MailService.CreateDefault();
    }

    public LiveTestSettings Settings { get; }

    public MailService Service { get; }

    public void Dispose()
    {
        try
        {
            DisposeCore();
        }
        finally
        {
            // Outside the teardown on purpose, exactly like LiveLifecycleFixture's copy: a
            // tripwire that can be swallowed is not one. Whether this is where the run gets
            // verified depends on the FILTER, which LiveTierRunPlan knows and this fixture
            // does not.
            LiveStoreCountTripwire.CollectionFinished(LiveCollections.Phase2);
        }
    }

    /// <summary>This fixture's own teardown, separated so the tripwire signal cannot be skipped.</summary>
    private void DisposeCore()
    {
        // Releases COM references only - Outlook keeps running (S7: never kill/close).
        Service.Dispose();
    }
}

[CollectionDefinition(LiveCollections.Phase2)]
public sealed class LivePhase2Collection : ICollectionFixture<LivePhase2Fixture>
{
}
