using OutlookAI.Core.Com;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// Stands where <c>ComGateway</c> stands, runs the operation, and remembers what was asked of it and
/// what came back out of it - the recovery the caller requested, and the exception the operation
/// threw - so the position <c>ComGateway.Run</c>'s exception filter occupies can be observed without
/// an Outlook to connect to or disconnect from. The one copy (Q154).
/// </summary>
internal sealed class RecordingGateway : IComGateway
{
    private readonly IOutlookSession _session;

    internal RecordingGateway(IOutlookSession session)
    {
        _session = session;
    }

    public event Action? OutlookGone
    {
        add { }
        remove { }
    }

    /// <summary>The exception the last failing operation threw, rethrown unchanged to the caller.</summary>
    internal Exception? Observed { get; private set; }

    /// <summary>The recovery the last operation asked for; null before the first and after <see cref="Forget"/>.</summary>
    internal ComSessionRecovery? RequestedRecovery { get; private set; }

    public bool IsConnected => true;

    public bool? QuitSinkActive => null;

    public bool ProbeConnected() => true;

    internal void Forget() => RequestedRecovery = null;

    public T Run<T>(Func<IOutlookSession, T> operation) => Run(operation, ComSessionRecovery.None);

    public T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery)
    {
        RequestedRecovery = recovery;
        try
        {
            return operation(_session);
        }
        catch (Exception ex)
        {
            Observed = ex;
            throw;
        }
    }

    public T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false)
    {
        return Run(operation);
    }

    public ComHostDiagnostics GetDiagnostics() => new ComHostDiagnostics("in-process", "ready");

    public void Dispose()
    {
    }
}
