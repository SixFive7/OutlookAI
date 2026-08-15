using System;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// How the service layer reaches a live Outlook session, without knowing whether that
    /// session is in this process or in the killable COM host.
    /// <para>
    /// Two implementations exist. <see cref="ComGateway"/> owns a real in-process COM
    /// session and runs inside the COM host child. The MCP server uses a remote
    /// implementation whose <see cref="Run{T}"/> hands out a proxy session, so each
    /// method the operation calls becomes one bounded round trip across the pipe.
    /// </para>
    /// <para>
    /// Note what this implies for an operation that makes SEVERAL session calls: they are
    /// several round trips, not one atomic unit. That is deliberate and it is safe here -
    /// every such operation in the service layer is a retry-across-stores loop whose
    /// steps are independently idempotent - but it is a constraint on anything new.
    /// </para>
    /// </summary>
    public interface IComGateway : IDisposable
    {
        /// <summary>
        /// Raised when Outlook goes away underneath the session, so cached state derived
        /// from it can be dropped promptly rather than on next use.
        /// </summary>
        event Action? OutlookGone;

        /// <summary>True when a session is currently held. May be stale; see <see cref="ProbeConnected"/>.</summary>
        bool IsConnected { get; }

        /// <summary>Whether the held session's Quit sink is advised; null when no session is held.</summary>
        bool? QuitSinkActive { get; }

        /// <summary>Liveness with a real round trip: true only when Outlook actually answers.</summary>
        bool ProbeConnected();

        /// <summary>Runs <paramref name="operation"/> against a live session, connecting when necessary.</summary>
        T Run<T>(Func<IOutlookSession, T> operation);
    }
}
