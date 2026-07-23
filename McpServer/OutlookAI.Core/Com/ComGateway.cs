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
    /// HELD OPEN for the gateway's lifetime - a headless COM-started Outlook exits when
    /// its last client disconnects (Phase-1 execution fact), so the standing session is
    /// what keeps Outlook (and with it, index updating) alive between tool calls.
    ///
    /// D17 rules enforced here: Outlook is started when needed unless the
    /// OutlookAISetup installer mutex is held (clear retry-later error instead); Outlook
    /// is NEVER killed, stopped or restarted by this process; everything runs
    /// non-elevated (S8 - an elevation mismatch breaks the COM attach).
    /// </summary>
    public sealed class ComGateway : IDisposable
    {
        private readonly object _lock = new object();
        private readonly bool _allowStartingOutlook;
        private OutlookComSession? _session;
        private bool _disposed;

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

        /// <summary>True when a COM session is currently connected.</summary>
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
        /// Runs <paramref name="operation"/> against a live session, connecting (and
        /// starting Outlook per D17) when necessary. If the held session turns out dead
        /// (Outlook exited between calls), it is rebuilt exactly once.
        /// </summary>
        public T Run<T>(Func<OutlookComSession, T> operation)
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
            catch (COMException ex) when (IsDisconnectHResult(ex.HResult))
            {
                Invalidate(session);
                OutlookComSession retrySession = GetOrConnect();
                return operation(retrySession);
            }
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
                    catch (Exception ex) when (ex is COMException || ex is InvalidOperationException || ex is ObjectDisposedException)
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
                    _session = OutlookComSession.Connect(_allowStartingOutlook);
                }
                catch (InvalidOperationException ex)
                {
                    throw new OutlookUnavailableException(ex.Message, ex);
                }

                return _session;
            }
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

        private static bool IsDisconnectHResult(int hresult)
        {
            const int RpcDisconnected = unchecked((int)0x80010108); // RPC_E_DISCONNECTED
            const int RpcServerUnavailable = unchecked((int)0x800706BA); // HRESULT_FROM_WIN32(RPC_S_SERVER_UNAVAILABLE)
            const int RpcCallFailed = unchecked((int)0x800706BE); // RPC_S_CALL_FAILED
            return hresult == RpcDisconnected || hresult == RpcServerUnavailable || hresult == RpcCallFailed;
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
