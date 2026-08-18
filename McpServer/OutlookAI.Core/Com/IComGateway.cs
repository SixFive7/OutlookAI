using System;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Whether a failed operation may be run a SECOND time against a rebuilt session.
    /// <para>
    /// This is a safety decision, not a tuning knob. The rebuild fires on the
    /// RPC_E_DISCONNECTED family, which includes <c>RPC_S_CALL_FAILED</c> - an HRESULT that
    /// means the call may or may not already have executed - so a re-run is a possible
    /// SECOND execution, not a first one. <see cref="ComSessionOperations"/> holds the
    /// per-operation classification and the reasoning behind it.
    /// </para>
    /// </summary>
    public enum ComSessionRecovery
    {
        /// <summary>
        /// Never re-run. The default for anything whose effect the gateway cannot see -
        /// which includes every multi-call lambda, because replaying one replays all of its
        /// steps, including any that already succeeded.
        /// </summary>
        None = 0,

        /// <summary>
        /// Re-run once against a rebuilt session. Only ever asked for by a caller that knows
        /// the operation is a single read (see <c>GatewayRoutingProxy</c>).
        /// </summary>
        RebuildOnce = 1,
    }

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

        /// <summary>
        /// Runs <paramref name="operation"/> against a live session, connecting when
        /// necessary. The operation is never re-run: see the overload below.
        /// </summary>
        T Run<T>(Func<IOutlookSession, T> operation);

        /// <summary>
        /// Runs <paramref name="operation"/> and says whether it may be re-run once against
        /// a rebuilt session if Outlook disconnects underneath it.
        /// <para>
        /// Only a caller that knows the operation is exactly ONE read may ask for
        /// <see cref="ComSessionRecovery.RebuildOnce"/>. Two things have to hold, and only
        /// the COM host's routing proxy can promise both: the lambda makes a single contract
        /// call (so a replay does not redo earlier steps), and that call is classified
        /// read-only by <see cref="ComSessionOperations"/> (so a replay of a call that may
        /// already have executed cannot have an effect).
        /// </para>
        /// </summary>
        T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery);

        /// <summary>
        /// Runs <paramref name="operation"/> with an explicit time budget instead of the
        /// default one. The budget bounds the whole lambda, not each round trip inside it.
        /// <para>
        /// Exists for <c>outlook_health</c>. Health is asked precisely when Outlook may be
        /// unresponsive, so it must not spend the ordinary two-minute budget discovering
        /// that - it has to answer quickly and say so. The in-process implementation
        /// ignores the budget, having no way to enforce one.
        /// </para>
        /// </summary>
        /// <param name="allowConnectFloor">
        /// Whether the budget covers this operation's OWN work only, leaving the host free
        /// to add its cold-start connect allowance on the first call.
        /// <para>
        /// Default false, which is what health needs: an explicit short budget must not be
        /// widened by anything, because a wedged Outlook is exactly when health is asked.
        /// The freshness sweep needs the opposite - its budget is for the SWEEP, and
        /// charging the COM attach to it meant the first search on a fresh host had to fit
        /// both into 30 s, so on a machine where attaching to a large OST takes longer than
        /// that the sweep could never succeed at all.
        /// </para>
        /// </param>
        T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false);

        /// <summary>How Outlook is being reached, and the health of that path.</summary>
        ComHostDiagnostics GetDiagnostics();
    }
}
