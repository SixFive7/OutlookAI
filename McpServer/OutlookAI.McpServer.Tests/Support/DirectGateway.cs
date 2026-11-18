using OutlookAI.Core.Com;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// Runs operations straight against a stand-in session - no COM host, no pipe, no budget layer -
/// where the product has <c>ComGateway</c>. The one copy for every suite that drives the product
/// over a fabricated <see cref="IOutlookSession"/> (Q154).
/// <para>
/// Two positions besides "connected": <c>connected: false</c> reports Outlook as not connected
/// while it still runs what it is given, for a path that must decide without asking Outlook; and
/// <see cref="Down"/> refuses every operation the way an Outlook that does not answer is refused.
/// </para>
/// </summary>
internal sealed class DirectGateway : IComGateway
{
    private readonly IOutlookSession _session;

    private readonly bool _connected;

    internal DirectGateway(IOutlookSession session, bool connected = true)
    {
        _session = session;
        _connected = connected;
    }

    public event Action? OutlookGone
    {
        add { }
        remove { }
    }

    /// <summary>While true, Outlook reads as not connected and every operation throws <see cref="OutlookUnavailableException"/>.</summary>
    public bool Down { get; set; }

    public bool IsConnected => _connected && !Down;

    public bool? QuitSinkActive => null;

    public bool ProbeConnected() => _connected && !Down;

    public T Run<T>(Func<IOutlookSession, T> operation) => Invoke(operation);

    public T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery) => Invoke(operation);

    public T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false)
        => Invoke(operation);

    public ComHostDiagnostics GetDiagnostics() => new ComHostDiagnostics("in-process", "ready");

    public void Dispose()
    {
    }

    private T Invoke<T>(Func<IOutlookSession, T> operation)
    {
        if (Down)
        {
            throw new OutlookUnavailableException("Outlook is not answering (T1 stand-in).");
        }

        return operation(_session);
    }
}
