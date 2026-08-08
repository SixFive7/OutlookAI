using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Shared state for the Phase-5 T2/T3 live tier (high-friction send): the MailService
/// under test, an INDEPENDENT verify session (arrival sweeps, folder identities,
/// modification of this run's own drafts), and one unique run marker for the S3
/// double-match rule. ALL send-flow tests target the test hub only (S2/D20:
/// telefonie-to-telefonie; business accounts get NO send-flow artifacts at all).
/// Fixture disposal runs a final tag+marker cleanup on the hub.
/// </summary>
public sealed class LivePhase5Fixture : IDisposable
{
    private readonly Lazy<OutlookComSession> _verifySession;

    public LivePhase5Fixture()
    {
        Settings = LiveTestSettings.Load();

        // Fail-closed per-store count tripwire: no census, no live tier. Cheap after
        // the first fixture (one process-wide baseline).
        LiveStoreCountTripwire.EnsureBaseline(Settings);
        Service = MailService.CreateDefault();
        RunMarker = "p5" + Guid.NewGuid().ToString("N").Substring(0, 14);
        _verifySession = new Lazy<OutlookComSession>(
            () => OutlookComSession.Connect(allowStartingOutlook: true),
            LazyThreadSafetyMode.ExecutionAndPublication);
    }

    public LiveTestSettings Settings { get; }

    public MailService Service { get; }

    /// <summary>Per-run unique marker; every artifact subject carries tag + marker (S3).</summary>
    public string RunMarker { get; }

    /// <summary>Independent COM session for verification and test-side modifications.</summary>
    public OutlookComSession VerifySession => _verifySession.Value;

    /// <summary>Builds a tagged subject: [OutlookAI-McpTest] + run marker + label.</summary>
    public string TaggedSubject(string label)
    {
        return LiveOutlookTestMailer.SubjectTag + " " + RunMarker + " " + label;
    }

    /// <summary>StoreID of a store by display name (via the verify session).</summary>
    public string GetStoreId(string storeDisplayName)
    {
        return VerifySession.GetStores()
                .FirstOrDefault(s => string.Equals(s.DisplayName, storeDisplayName, StringComparison.OrdinalIgnoreCase))?.StoreId
            ?? throw new InvalidOperationException("Store not found by display name.");
    }

    public void Dispose()
    {
        try
        {
            // Final belt: purge anything of THIS run still tagged in the hub store.
            LiveOutlookTestMailer.DeleteTaggedArtifacts(Settings.TestHubStoreDisplayName, RunMarker);
        }
        catch (Exception)
        {
            // Best-effort - each test already cleaned up and asserted in finally.
        }

        if (_verifySession.IsValueCreated)
        {
            // Releases COM references only - Outlook keeps running (S7: never kill/close).
            _verifySession.Value.Dispose();
        }

        Service.Dispose();
    }
}

[CollectionDefinition("LivePhase5")]
public sealed class LivePhase5Collection : ICollectionFixture<LivePhase5Fixture>
{
}
