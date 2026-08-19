using Xunit;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The guard hook for the T3 stdio shape tests, which own no shared state of their own:
/// each spawns the built server exe for itself and talks to it over a pipe.
/// <para>
/// <b>Why they need a fixture anyway.</b> Three of them - <c>LiveMcpToolShapeTests</c>,
/// <c>Phase3LiveMcpToolShapeTests</c> and <c>Phase7LiveMcpToolShapeTests</c> - carried
/// <c>Category=Live</c> while belonging to no collection at all, so xunit gave each its own
/// implicit one and none of them passed through <see cref="LiveStoreCountTripwire"/>. They
/// ran against real mailboxes with no per-store census, no health preflight and no
/// verification: nine live tests, one of which drives Outlook's UI and one of which runs a
/// full exhaustive scan, entirely outside the safety envelope every other live test is held
/// to. A collection with a fixture is the only mechanism xunit offers for "run this before
/// the first test and after the last", so the fixture exists to be that hook and holds
/// nothing else.
/// </para>
/// <para>
/// Deliberately holds no <c>MailService</c> and no COM session. These tests assert things
/// about a server process they start themselves, and a session held open beside them would
/// be state they did not ask for.
/// </para>
/// </summary>
public sealed class LiveMcpToolShapeFixture : IDisposable
{
    public LiveMcpToolShapeFixture()
    {
        Settings = LiveTestSettings.Load();

        // Fail-closed per-store count tripwire: no census, no live tier. Also the health
        // preflight, which these tests previously skipped entirely.
        LiveStoreCountTripwire.EnsureBaseline(Settings);
    }

    /// <summary>The machine-local live-test settings, loaded once for the collection.</summary>
    public LiveTestSettings Settings { get; }

    public void Dispose()
    {
        LiveStoreCountTripwire.CollectionFinished(LiveCollections.McpToolShape);
    }
}

[CollectionDefinition(LiveCollections.McpToolShape)]
public sealed class LiveMcpToolShapeCollection : ICollectionFixture<LiveMcpToolShapeFixture>
{
}
