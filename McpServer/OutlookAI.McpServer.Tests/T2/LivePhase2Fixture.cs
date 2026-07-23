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
        Service = MailService.CreateDefault();
    }

    public LiveTestSettings Settings { get; }

    public MailService Service { get; }

    public void Dispose()
    {
        // Releases COM references only - Outlook keeps running (S7: never kill/close).
        Service.Dispose();
    }
}

[CollectionDefinition("LivePhase2")]
public sealed class LivePhase2Collection : ICollectionFixture<LivePhase2Fixture>
{
}
