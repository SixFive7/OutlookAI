using System;
using System.Runtime.InteropServices;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Thrown when a COM-requiring operation cannot proceed because Outlook is
    /// unavailable and may not be started right now (D17: the OutlookAISetup installer
    /// mutex is held, or autostart was disabled). The message is a clear retry-later
    /// instruction for the calling agent.
    /// </summary>
    public sealed class OutlookUnavailableException : Exception
    {
        /// <summary>Creates the exception.</summary>
        public OutlookUnavailableException(string message)
            : base(message)
        {
        }

        /// <summary>Creates the exception with an inner cause.</summary>
        public OutlookUnavailableException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }

    /// <summary>
    /// The single owner of Outlook COM access for a host process (v3.MD section 0.5):
    /// ONE dedicated pumped STA thread (inside <see cref="OutlookComSession"/>) with a
    /// serialized request queue, connected lazily on the first COM-requiring call and
    /// HELD OPEN for the gateway's lifetime - the standing session keeps a headless
    /// COM-started Outlook (and with it, index updating) alive between tool calls;
    /// measured 2026-07-23: after the LAST client releases, a headless Outlook keeps
    /// running for ~11.5 minutes and then exits on its own.
    ///
    /// SF-1/SF-2 lifecycle handling (soak fix): the session watches its OUTLOOK.EXE and
    /// releases every held ref the moment the process exits (user close, crash,
    /// logoff), <see cref="ProbeConnected"/> reports pinged - never stale - liveness,
    /// and the disconnect family of failures triggers a one-shot rebuild inside
    /// <see cref="Run{T}(Func{IOutlookSession, T}, ComSessionRecovery)"/> - for the
    /// operations that may be re-run, which is reads only
    /// (<see cref="ComSessionOperations"/>).
    ///
    /// NOT coverable from this side (probe-measured): an external
    /// client driving Application.Quit while sessions are attached parks Outlook
    /// indefinitely (no Quit event reaches out-of-process sinks, the process stays
    /// alive, COM keeps answering); the park self-heals ~6 s after our refs release
    /// (e.g. the server process ends). Protocol: release/stop server sessions BEFORE
    /// quitting Outlook programmatically.
    ///
    /// D17 rules enforced here: Outlook is started when needed unless the
    /// OutlookAISetup installer mutex is held (clear retry-later error instead); Outlook
    /// is NEVER killed, stopped or restarted by this process; everything runs
    /// non-elevated (S8 - an elevation mismatch breaks the COM attach).
    /// </summary>
    public sealed class ComGateway : IComGateway, IDisposable
    {
        private readonly object _lock = new object();
        private readonly bool _allowStartingOutlook;
        private OutlookComSession? _session;
        private bool _disposed;

        /// <summary>
        /// Raised (on a worker thread) when Outlook signalled Quit or its process exited.
        /// The COM host forwards this to the MCP server so the server can drop cached
        /// store details instead of discovering the loss on the next call.
        /// </summary>
        public event Action? OutlookGone;

        /// <summary>Creates the gateway; connection happens lazily on first use.</summary>
        public ComGateway(bool allowStartingOutlook = true)
        {
            _allowStartingOutlook = allowStartingOutlook;
        }

        /// <summary>True when an OUTLOOK.EXE process exists for this user.</summary>
        public static bool IsOutlookRunning()
        {
            return OutlookComSession.IsOutlookProcessRunning();
        }

        /// <summary>True while the add-in installer's SetupMutex is held (D17 window).</summary>
        public static bool IsInstallerMutexHeld()
        {
            return OutlookComSession.IsInstallerMutexHeld();
        }

        /// <summary>
        /// True when a COM session is currently held. NOTE: a held session can be stale
        /// (Outlook exited since) - use <see cref="ProbeConnected"/> for a liveness-true
        /// answer (SF-1). The SF-2 watchers release dead sessions within moments, so the
        /// two rarely disagree for long.
        /// </summary>
        public bool IsConnected
        {
            get
            {
                lock (_lock)
                {
                    return _session != null;
                }
            }
        }

        /// <summary>
        /// Whether the held session's Application Quit sink is advised (SF-2 proactive
        /// release); null when no session is held. Diagnostic surface for tests.
        /// </summary>
        public bool? QuitSinkActive
        {
            get
            {
                lock (_lock)
                {
                    return _session?.QuitSinkActive;
                }
            }
        }

        /// <summary>
        /// PROBED liveness (SF-1 fix): false when no session is held; otherwise pings
        /// the held session and returns true only when Outlook actually answers. A dead
        /// session is released immediately (never reconnects - probing must not start
        /// Outlook).
        /// </summary>
        public bool ProbeConnected()
        {
            if (_disposed)
            {
                return false;
            }

            OutlookComSession? session;
            lock (_lock)
            {
                session = _session;
            }

            if (session == null)
            {
                return false;
            }

            try
            {
                session.GetProfileName();
                return true;
            }
            catch (Exception ex) when (IsSessionUnusable(ex))
            {
                Invalidate(session);
                return false;
            }
        }

        /// <summary>
        /// Runs <paramref name="operation"/> against a live session, connecting (and
        /// starting Outlook per D17) when necessary. A dead HELD session is still replaced
        /// before the operation starts (<see cref="GetOrConnect"/> pings first), but the
        /// operation itself is never re-run - for that, name the recovery.
        /// </summary>
        public T Run<T>(Func<IOutlookSession, T> operation)
        {
            return Run(operation, ComSessionRecovery.None);
        }

        /// <summary>
        /// Runs <paramref name="operation"/> against a live session, connecting (and
        /// starting Outlook per D17) when necessary. If the session dies UNDER the call and
        /// <paramref name="recovery"/> allows it, the session is rebuilt and the operation
        /// runs exactly once more.
        /// <para>
        /// The re-run is a possible SECOND EXECUTION, not a first one:
        /// <see cref="IsDisconnectHResult"/> includes <c>RPC_S_CALL_FAILED</c> (0x800706BE),
        /// which means precisely that the call may or may not have reached Outlook. That is
        /// harmless for a read and unrecoverable for a send, so the caller has to say which
        /// it has, and the default is <see cref="ComSessionRecovery.None"/>.
        /// </para>
        /// <para>
        /// Why the default is no: this method cannot see what the lambda does. Before the
        /// COM host existed this recovery was reached only through whole SERVICE operations,
        /// several contract calls long, and replaying one replays the steps that already
        /// succeeded. The remote gateway refused to do that for exactly this reason
        /// (<c>RemoteComGateway</c>, "a coarse retry here would quietly undo that"); this
        /// side now refuses on the same grounds. The routing proxy inside the COM host is
        /// the one caller whose lambda is a single classified contract call, so it is the
        /// one caller that can safely opt in - which keeps every production READ covered,
        /// per call, and no write covered at all.
        /// </para>
        /// <para>
        /// A dead session is dropped either way. The difference is only whether the work is
        /// attempted again here or reported to the caller.
        /// </para>
        /// </summary>
        public T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery)
        {
            if (operation == null)
            {
                throw new ArgumentNullException(nameof(operation));
            }

            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(ComGateway));
            }

            OutlookComSession session = GetOrConnect();
            try
            {
                return operation(session);
            }
            catch (Exception ex) when (IsDisconnectException(ex))
            {
                Invalidate(session);
                if (recovery != ComSessionRecovery.RebuildOnce)
                {
                    throw;
                }

                OutlookComSession retrySession = GetOrConnect();
                return operation(retrySession);
            }
        }

        /// <summary>
        /// Runs with an explicit budget, enforced BETWEEN contract calls.
        /// <para>
        /// This implementation owns the COM session in its own process, so it cannot bound
        /// ONE call: a blocked outbound COM call is not cancellable, and killing the caller
        /// is not an option when the caller is us. Only the out-of-process gateway can do
        /// that, by ending the child. What this one can do - and now does - is bound the
        /// SEQUENCE: <see cref="BudgetedSessionProxy"/> checks the clock before each call
        /// and refuses to start another once the budget is spent.
        /// </para>
        /// <para>
        /// It used to be <c>{ return Run(operation); }</c>, accepting the budget and
        /// discarding it. That reads as harmless - inside the COM host child the parent's
        /// watchdog is the real bound - but <c>MailService.CreateDefault()</c> uses this
        /// gateway too, and every live (T2) fixture is built on <c>CreateDefault()</c>. So
        /// the entire live tier exercised a path with no budget, no aggregate, no breaker
        /// and no hang detector, which is both a coverage hole (anything sized against
        /// those budgets was unverified there by construction) and a hang the tier had
        /// nothing to stop.
        /// </para>
        /// <para>
        /// <paramref name="allowConnectFloor"/> is accepted and has no effect here, for a
        /// reason rather than by omission: the clock starts after the session is connected,
        /// so a cold start is never inside the caller's work budget in this implementation.
        /// The flag exists for the remote gateway, where the connect IS chargeable and the
        /// caller has to say whether to charge it.
        /// </para>
        /// </summary>
        public T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false)
        {
            if (operation == null)
            {
                throw new ArgumentNullException(nameof(operation));
            }

            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(ComGateway));
            }

            OutlookComSession session = GetOrConnect();
            IOutlookSession budgeted = BudgetedSessionProxy.Wrap(session, budgetMilliseconds);
            try
            {
                return operation(budgeted);
            }
            catch (Exception ex) when (IsDisconnectException(ex))
            {
                // Same rule as the no-recovery overload: the dead session is dropped, the
                // work is never re-run. A budgeted operation is a multi-call lambda by
                // assumption, and replaying one replays the steps that already succeeded.
                Invalidate(session);
                throw;
            }
        }

        /// <inheritdoc />
        public ComHostDiagnostics GetDiagnostics()
        {
            return new ComHostDiagnostics(
                mode: "in-process",
                state: IsConnected ? "ready" : "none",
                processId: System.Diagnostics.Process.GetCurrentProcess().Id);
        }

        /// <summary>
        /// Returns the live session, pinging a held one first and reconnecting when the
        /// ping fails. Serialized: concurrent callers queue here while a cold Outlook
        /// start is in progress.
        /// </summary>
        public OutlookComSession GetOrConnect()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(ComGateway));
            }

            lock (_lock)
            {
                if (_session != null)
                {
                    try
                    {
                        _session.GetProfileName();
                        return _session;
                    }
                    catch (Exception ex) when (IsSessionUnusable(ex))
                    {
                        // Outlook went away (it is never stopped by us - user action or
                        // crash). Drop the dead session and reconnect below.
                        try
                        {
                            _session.Dispose();
                        }
                        catch (Exception)
                        {
                            // Best-effort teardown of a dead session.
                        }

                        _session = null;
                    }
                }

                try
                {
                    _session = OutlookComSession.Connect(_allowStartingOutlook, HandleSessionGone);
                }
                catch (InvalidOperationException ex)
                {
                    throw new OutlookUnavailableException(ex.Message, ex);
                }

                return _session;
            }
        }

        /// <summary>
        /// SF-2 fix: invoked (on a worker thread) when the session's Outlook signals
        /// Quit or its process exits - releases ALL held COM refs immediately so (a) a
        /// dying Outlook is never kept alive by our references and (b) the next
        /// COM-needing call reconnects (and autostarts, D17) cleanly.
        /// </summary>
        private void HandleSessionGone(OutlookComSession session)
        {
            Invalidate(session);
            OutlookGone?.Invoke();
        }

        private void Invalidate(OutlookComSession session)
        {
            lock (_lock)
            {
                if (!ReferenceEquals(_session, session))
                {
                    return;
                }

                try
                {
                    _session!.Dispose();
                }
                catch (Exception)
                {
                    // Best-effort teardown of a dead session.
                }

                _session = null;
            }
        }

        /// <summary>
        /// The RPC_E_DISCONNECTED family: Outlook exited (or is exiting) under a live
        /// proxy. Includes the RCW-separated and disposed shapes those calls take after
        /// the SF-2 watcher released the refs mid-flight.
        /// </summary>
        private static bool IsDisconnectException(Exception ex)
        {
            if (ex is InvalidComObjectException || ex is ObjectDisposedException)
            {
                return true; // Session refs were force-released (watcher) while a call was staged.
            }

            return ex is COMException com && IsDisconnectHResult(com.HResult);
        }

        /// <summary>
        /// Anything that makes a HELD session worthless for the liveness ping: the
        /// disconnect family plus the late-bound failure shapes (Phase-2 fact 2 - the
        /// dynamic binder maps dead-proxy failures to plain .NET exception types).
        /// </summary>
        private static bool IsSessionUnusable(Exception ex)
        {
            return IsDisconnectException(ex)
                || ex is InvalidOperationException
                || OutlookComSession.IsComCallFailure(ex);
        }

        private static bool IsDisconnectHResult(int hresult)
        {
            const int RpcDisconnected = unchecked((int)0x80010108); // RPC_E_DISCONNECTED
            const int RpcServerUnavailable = unchecked((int)0x800706BA); // HRESULT_FROM_WIN32(RPC_S_SERVER_UNAVAILABLE)
            const int RpcCallFailed = unchecked((int)0x800706BE); // RPC_S_CALL_FAILED
            const int RpcServerDied = unchecked((int)0x80010007); // RPC_E_SERVER_DIED
            const int RpcServerDiedDne = unchecked((int)0x80010012); // RPC_E_SERVER_DIED_DNE
            return hresult == RpcDisconnected
                || hresult == RpcServerUnavailable
                || hresult == RpcCallFailed
                || hresult == RpcServerDied
                || hresult == RpcServerDiedDne;
        }

        /// <summary>
        /// Releases the COM session. Outlook keeps running (S7/D17: never stopped by
        /// this process) - but note a HEADLESS Outlook that this gateway itself started
        /// may then exit on its own once its last COM client disconnects.
        /// </summary>
        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            lock (_lock)
            {
                _session?.Dispose();
                _session = null;
            }
        }
    }
}
