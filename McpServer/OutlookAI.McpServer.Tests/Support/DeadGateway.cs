using OutlookAI.Core.Com;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// A gateway with no Outlook behind it: never connected, and every operation throws
/// <see cref="OutlookUnavailableException"/>, as <c>ComGateway</c> does while Outlook is not running.
/// The one copy for the suites that need the product's "Outlook is not running" path (Q154).
/// </summary>
internal sealed class DeadGateway : IComGateway
{
    public event Action? OutlookGone
    {
        add { }
        remove { }
    }

    public bool IsConnected => false;

    public bool? QuitSinkActive => null;

    public bool ProbeConnected() => false;

    public T Run<T>(Func<IOutlookSession, T> operation) => throw Gone();

    public T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery) => throw Gone();

    public T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false)
        => throw Gone();

    public ComHostDiagnostics GetDiagnostics() => new ComHostDiagnostics("in-process", "down");

    public void Dispose()
    {
    }

    private static OutlookUnavailableException Gone()
    {
        return new OutlookUnavailableException("Outlook is not running.");
    }
}
