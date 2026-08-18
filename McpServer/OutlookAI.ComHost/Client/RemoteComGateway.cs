using System.Runtime.Versioning;
using OutlookAI.ComHost.Supervision;
using OutlookAI.Core.Com;

namespace OutlookAI.ComHost.Client
{
    /// <summary>
    /// The MCP server's <see cref="IComGateway"/>: hands the service layer a session that
    /// lives in another process.
    /// <para>
    /// Note what is deliberately absent compared with the in-process
    /// <see cref="ComGateway"/>: there is no retry-the-whole-operation-on-disconnect here.
    /// That recovery still exists, but it belongs to the COM host, which retries a single
    /// call against a rebuilt session. Retrying at this level would replay an entire
    /// multi-call operation, and one of those operations is <c>TrySendDraft</c> - the
    /// existing design keeps double-send impossible by making the only escape route an
    /// exception thrown before any COM work happens, and a coarse retry here would quietly
    /// undo that.
    /// </para>
    /// </summary>
    [SupportedOSPlatform("windows")]
    public sealed class RemoteComGateway : IComGateway
    {
        private readonly ComHostSupervisor _supervisor;
        private readonly IOutlookSession _session;
        private bool _disposed;

        /// <summary>Creates the gateway and its supervisor. The child starts lazily, on first use.</summary>
        public RemoteComGateway(bool allowStartingOutlook = true)
        {
            _supervisor = new ComHostSupervisor(allowStartingOutlook);
            _session = RemoteSessionProxy.Create(_supervisor);
            _supervisor.OutlookGone += () => OutlookGone?.Invoke();
        }

        /// <inheritdoc />
        public event Action? OutlookGone;

        /// <inheritdoc />
        public bool IsConnected => _supervisor.State == ComHostState.Ready;

        /// <summary>
        /// Always null for a remote session. The flag is a child-side diagnostic that no
        /// production path reads; surfacing it would cost a round trip to answer a
        /// question only the live tests ask, and they use an in-process gateway.
        /// </summary>
        public bool? QuitSinkActive => null;

        /// <summary>Whether a COM host process is currently running, and its PID.</summary>
        public int? ChildProcessId => _supervisor.ChildProcessId;

        /// <summary>How many times the COM host has been replaced this process lifetime.</summary>
        public int RestartCount => _supervisor.RestartCount;

        /// <summary>The last supervision failure, for health reporting. Null when healthy.</summary>
        public string? LastFailureMessage => _supervisor.LastFailureMessage;

        /// <inheritdoc />
        public bool ProbeConnected()
        {
            if (_disposed)
            {
                return false;
            }

            // No host, no session - answer immediately. Starting one here would be wrong
            // twice over: it contradicts the in-process gateway's contract ("never
            // reconnects - probing must not start Outlook"), and it would let a liveness
            // PROBE cold-start Outlook and block for as long as that takes.
            if (_supervisor.State != ComHostState.Ready)
            {
                return false;
            }

            try
            {
                // Bounded by the health-probe budget: this is asked precisely when Outlook
                // may be unresponsive, so it must answer either way rather than join the
                // problem it is reporting on.
                using (ComHostRequestContext.Enter(CancellationToken.None, ComHostPolicy.HealthProbeDeadlineMilliseconds))
                {
                    _ = _session.GetProfileName();
                    return true;
                }
            }
            catch (Exception)
            {
                // Any failure - timeout, no child, dead Outlook - means "not connected".
                // The reason is reported separately through LastFailureMessage.
                return false;
            }
        }

        /// <inheritdoc />
        /// <remarks>
        /// The lambda gets an AGGREGATE budget, not just a per-call one. Each contract call
        /// inside it is one bounded round trip, and an operation that makes several - hit
        /// location makes 1 + up to 3 + N, the archive path walks every store - previously
        /// gave each of them a full budget and so had no bound of its own. The aggregate is
        /// the same ordinary operation deadline, measured from here across the whole lambda,
        /// so it follows <c>OUTLOOKAI_COMHOST_DEADLINE_MS</c> like everything else.
        /// </remarks>
        public T Run<T>(Func<IOutlookSession, T> operation)
        {
            ArgumentNullException.ThrowIfNull(operation);
            ObjectDisposedException.ThrowIf(_disposed, this);

            using (ComHostRequestContext.Enter(
                ComHostRequestContext.Token,
                deadlineOverrideMilliseconds: null,
                aggregateBudgetMilliseconds: ComHostPolicy.DeadlineFor(ComHostOperationClass.Operation, null)))
            {
                return operation(_session);
            }
        }

        /// <inheritdoc />
        /// <remarks>
        /// The recovery argument is accepted and ignored: this side has never had a re-run
        /// to suppress. The classification it carries is honoured one process away, by the
        /// COM host's routing proxy, where a single contract call is the unit being retried.
        /// </remarks>
        public T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery)
        {
            return Run(operation);
        }

        /// <inheritdoc />
        public T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false)
        {
            ArgumentNullException.ThrowIfNull(operation);
            ObjectDisposedException.ThrowIf(_disposed, this);

            // The explicit budget bounds the whole lambda as well as each call in it:
            // health asking for 5 s means 5 s of health, not 5 s per round trip.
            using (ComHostRequestContext.Enter(
                ComHostRequestContext.Token,
                deadlineOverrideMilliseconds: budgetMilliseconds,
                aggregateBudgetMilliseconds: budgetMilliseconds,
                allowConnectFloor: allowConnectFloor))
            {
                return operation(_session);
            }
        }

        /// <inheritdoc />
        public ComHostDiagnostics GetDiagnostics()
        {
            return new ComHostDiagnostics(
                mode: "child-process",
                state: _supervisor.State.ToString().ToLowerInvariant(),
                processId: _supervisor.ChildProcessId,
                restartCount: _supervisor.RestartCount,
                lastFailure: _supervisor.LastFailureMessage,
                injectedFault: Host.ComHostFaultInjection.IsActive ? Host.ComHostFaultInjection.Description : null,
                unresponsive: _supervisor.IsUnresponsive,
                consecutiveTimeouts: _supervisor.ConsecutiveTimeouts);
        }

        /// <inheritdoc />
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _supervisor.Dispose();
        }
    }
}
