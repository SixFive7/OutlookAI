using OutlookAI.McpServer.Tools;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The shutdown release must not become a reason to START Outlook.
/// <para>
/// <c>ServerRuntime</c> holds its gateway in a <c>Lazy</c>, so the obvious spelling of
/// "dispose the gateway on the way out" - <c>Gateway.Dispose()</c> - CREATES one in a server
/// that never needed it. That is a real cost in the common case: a session that only ever
/// called <c>list_signatures</c>, or no tool at all, would build a supervisor and a dispatch
/// proxy at the last possible moment and then throw them away. The guard is one
/// <c>IsValueCreated</c> check, and an <c>IsValueCreated</c> check is exactly the kind of
/// line a later edit deletes as redundant, so it is pinned.
/// </para>
/// <para>
/// Whole thing in ONE test method on purpose. <c>ServerRuntime</c> is process-wide static
/// state and the release is deliberately once-only, so a second test method calling it would
/// be testing the second call whichever order xunit chose. Nothing else in this assembly
/// touches <c>ServerRuntime</c>, which is what makes the first assertion below meaningful.
/// </para>
/// <para>
/// In-process, and therefore silent about the thing that matters most - whether the shipped
/// <c>Program</c> actually calls this. That wiring is the whole defect, and it is pinned one
/// tier up, end to end, by <c>T3/ComHostSupervisionCiTests</c>.
/// </para>
/// </summary>
public sealed class ServerShutdownReleaseTests
{
    [Fact]
    public void ReleasingOnShutdown_NeverBuildsAGatewayThatNothingAskedFor()
    {
        Assert.False(
            ServerRuntime.HasGateway,
            "no test in this assembly uses ServerRuntime, so nothing should have built a gateway before this ran");

        ServerRuntime.ReleaseOutlookResources();

        Assert.False(
            ServerRuntime.HasGateway,
            "shutdown forced the Lazy and built a COM gateway the server never needed");

        // Idempotent: shutdown can be reached more than once (an exception escaping the host
        // unwinds through the same finally a normal stop does), and the second pass must be
        // a no-op rather than a second teardown of an already-disposed supervisor.
        ServerRuntime.ReleaseOutlookResources();
        Assert.False(ServerRuntime.HasGateway);
    }
}
