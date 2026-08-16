using System.Runtime.InteropServices;
using OutlookAI.ComHost.Protocol;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

namespace OutlookAI.ComHost.Supervision
{
    /// <summary>
    /// Rebuilds a child-side exception on the parent side.
    /// <para>
    /// Faithfulness matters more than it looks. <c>OutlookTools.Guard</c> branches on
    /// exception TYPE to choose the error payload the agent sees, and <c>ComGateway</c>
    /// keys its disconnect-retry on <see cref="COMException"/> HRESULTs. A generic
    /// exception carrying a formatted string would silently downgrade both: refusals
    /// would lose their machine-readable reason, and a genuine RPC_E_DISCONNECTED would
    /// stop being recognised as one.
    /// </para>
    /// </summary>
    internal static class ComHostErrorMapper
    {
        /// <summary>Turns a wire error into the closest equivalent exception.</summary>
        internal static Exception ToException(ComHostError error)
        {
            ArgumentNullException.ThrowIfNull(error);

            string message = error.Message ?? string.Empty;

            switch (error.Type)
            {
                case nameof(SendRefusedException):
                    return new SendRefusedException(error.Reason ?? string.Empty, message);

                case nameof(DraftRefusedException):
                    return new DraftRefusedException(error.Reason ?? string.Empty, message);

                case nameof(OutlookUnavailableException):
                    return new OutlookUnavailableException(message);

                case nameof(COMException):
                    return error.HResult is int hr
                        ? new COMException(message, hr)
                        : new COMException(message);

                case nameof(InvalidComObjectException):
                    return new InvalidComObjectException(message);

                case nameof(ObjectDisposedException):
                    return new ObjectDisposedException(string.Empty, message);

                case nameof(ArgumentNullException):
                    return new ArgumentNullException(paramName: null, message: message);

                case nameof(ArgumentOutOfRangeException):
                    return new ArgumentOutOfRangeException(paramName: null, message: message);

                case nameof(ArgumentException):
                    return new ArgumentException(message);

                case nameof(InvalidOperationException):
                    return new InvalidOperationException(message);

                case nameof(NotSupportedException):
                    return new NotSupportedException(message);

                case nameof(UnauthorizedAccessException):
                    return new UnauthorizedAccessException(message);

                case nameof(IOException):
                    return new IOException(message);

                case nameof(TimeoutException):
                    return new TimeoutException(message);

                default:
                    // Unknown child-side type. Keep the original type name in the message
                    // so it still reaches the agent and the audit log, rather than being
                    // flattened into an anonymous failure.
                    return new ComHostRemoteException(error.Type, message);
            }
        }
    }

    /// <summary>
    /// A child-side failure whose type the parent does not model explicitly. Carries the
    /// original type name so nothing is lost in translation.
    /// </summary>
    public sealed class ComHostRemoteException : Exception
    {
        /// <summary>Creates the exception.</summary>
        public ComHostRemoteException(string remoteType, string message)
            : base(message)
        {
            RemoteType = remoteType;
        }

        /// <summary>The exception type name as it was on the child side.</summary>
        public string RemoteType { get; }
    }

    /// <summary>
    /// Raised when an operation exceeded its deadline and the COM host was killed to
    /// reclaim it. Carries the operation and the budget it breached so the tool layer can
    /// say plainly what happened and what the caller can still do.
    /// </summary>
    /// <remarks>
    /// Derives from <see cref="TimeoutException"/> so Core can recognise a deadline
    /// breach - and report it in plain words rather than as a type name - without taking
    /// a reference on this assembly.
    /// </remarks>
    public sealed class ComHostTimeoutException : TimeoutException
    {
        /// <summary>Creates the exception.</summary>
        public ComHostTimeoutException(string operation, long deadlineMilliseconds)
            : base($"Outlook did not respond to '{operation}' within {deadlineMilliseconds} ms. "
                 + "The COM host was restarted to recover; Outlook itself may be busy or not responding.")
        {
            Operation = operation;
            DeadlineMilliseconds = deadlineMilliseconds;
        }

        /// <summary>The contract operation that timed out.</summary>
        public string Operation { get; }

        /// <summary>The budget it exceeded.</summary>
        public long DeadlineMilliseconds { get; }
    }

    /// <summary>
    /// Raised immediately, without contacting Outlook, while Outlook is known to be
    /// unresponsive.
    /// <para>
    /// Derives from <see cref="TimeoutException"/> so every layer that already degrades
    /// gracefully on a deadline breach - notably the freshness sweep, which falls back to
    /// index-only results - treats this identically, and reports its message rather than a
    /// type name.
    /// </para>
    /// </summary>
    public sealed class ComHostUnresponsiveException : TimeoutException
    {
        /// <summary>Creates the exception.</summary>
        public ComHostUnresponsiveException(int consecutiveTimeouts)
            : base($"Outlook has failed to answer {consecutiveTimeouts} request(s) in a row, so requests that need it "
                 + "are being refused immediately instead of waiting. Outlook is being re-checked periodically and this "
                 + "clears by itself once it responds; restarting Outlook fixes it immediately.")
        {
            ConsecutiveTimeouts = consecutiveTimeouts;
        }

        /// <summary>How many consecutive timeouts led to this.</summary>
        public int ConsecutiveTimeouts { get; }
    }

    /// <summary>
    /// Raised when the COM host could not be started, or died and is in start backoff.
    /// </summary>
    public sealed class ComHostUnavailableException : Exception
    {
        /// <summary>Creates the exception.</summary>
        public ComHostUnavailableException(string message)
            : base(message)
        {
        }

        /// <summary>Creates the exception with an inner cause.</summary>
        public ComHostUnavailableException(string message, Exception inner)
            : base(message, inner)
        {
        }
    }
}
