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
        /// <summary>
        /// Turns a wire error into the closest equivalent exception.
        /// </summary>
        /// <param name="error">The failure as the child described it.</param>
        /// <param name="operation">
        /// The contract operation the parent sent, when it is known. The wire error carries
        /// no operation name of its own, and one failure needs it: an answer too large to
        /// frame is reported over an operation that RAN, so whether the caller may repeat it
        /// depends entirely on whether that operation changes mail.
        /// </param>
        internal static Exception ToException(ComHostError error, string? operation = null)
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

                case nameof(ComHostResponseTooLargeException):
                    // Not raised by a `throw` in the COM layer, so invariant 10 cannot see
                    // it: the child stamps this name onto the wire error directly, because
                    // the failure happens while ENCODING the reply - past the point where
                    // throwing could still produce one.
                    return new ComHostResponseTooLargeException(message, operation);

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
                    // A child-side type this parent does not model. The name is carried on
                    // the exception rather than folded into the message, and the tool layer
                    // reports THAT as the error type - so the agent is told what actually
                    // failed instead of being told the name of the pipe it crossed. An
                    // earlier comment here claimed the name went into the message; it never
                    // did, and while that was believed the name reached nothing at all.
                    //
                    // Landing here is not free - OutlookTools.GuardAsync branches on
                    // exception TYPE to choose its advice, and this branch has none to
                    // choose from - so the set of types that reach it is held down by
                    // invariant 10 in .github/scripts/check-pinned-constants.ps1, which
                    // fails the build when the COM layer starts raising a type the switch
                    // above does not name.
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
            RemoteType = string.IsNullOrWhiteSpace(remoteType) ? nameof(ComHostRemoteException) : remoteType;
        }

        /// <summary>
        /// The exception type name as it was on the child side. Never blank: this is what
        /// the tool layer reports as the error type, and an empty <c>type</c> field would
        /// be a worse answer than the transport's own name.
        /// </summary>
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
        /// <summary>Creates the exception after repeated timeouts.</summary>
        public ComHostUnresponsiveException(int consecutiveTimeouts, int retryAfterSeconds)
            : base($"Outlook has failed to answer {consecutiveTimeouts} request(s) in a row, so requests that need it "
                 + "are being refused immediately instead of waiting. Outlook is being re-checked periodically and this "
                 + "clears by itself once it responds; restarting Outlook fixes it immediately.")
        {
            ConsecutiveTimeouts = consecutiveTimeouts;
            RetryAfterSeconds = retryAfterSeconds;
        }

        /// <summary>
        /// Creates the exception from a direct observation - Windows itself reporting
        /// Outlook's windows as hung - rather than from accumulated timeouts.
        /// </summary>
        public ComHostUnresponsiveException(string observation, int retryAfterSeconds)
            : base($"Outlook is running but not responding ({observation}), so requests that need it are being refused "
                 + "immediately rather than waiting for a call that would not return. This is detected directly from "
                 + "Windows, not guessed. It clears by itself once Outlook responds; restarting Outlook fixes it now.")
        {
            RetryAfterSeconds = retryAfterSeconds;
        }

        /// <summary>How many consecutive timeouts led to this, when that is what led to it.</summary>
        public int ConsecutiveTimeouts { get; }

        /// <summary>How long the caller should wait before retrying.</summary>
        public int RetryAfterSeconds { get; }
    }

    /// <summary>
    /// Raised immediately when Outlook is starting up, instead of making the caller wait
    /// for it.
    /// <para>
    /// A cold Outlook start can take tens of seconds. Blocking a tool call for that long
    /// is indistinguishable, from the caller's side, from the hang this whole design
    /// exists to remove - so the caller is told what is happening and roughly how long to
    /// wait, and can do something useful meanwhile. Derives from
    /// <see cref="TimeoutException"/> so the freshness sweep degrades to index-only
    /// results rather than failing the search.
    /// </para>
    /// </summary>
    public sealed class ComHostStartingException : TimeoutException
    {
        /// <summary>Creates the exception.</summary>
        public ComHostStartingException(int retryAfterSeconds, string reason)
            : base($"Outlook is starting up ({reason}); this answered straight away rather than making you wait for it. "
                 + $"Retry in about {retryAfterSeconds} seconds.")
        {
            RetryAfterSeconds = retryAfterSeconds;
        }

        /// <summary>How long the caller should wait before retrying.</summary>
        public int RetryAfterSeconds { get; }
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

        /// <summary>Creates the exception for a request the host was serving when it died.</summary>
        /// <param name="message">What the caller is told, from <c>ComHostSupervisor.DescribeInterruption</c>.</param>
        /// <param name="outcome">
        /// The machine-readable half of that same sentence - see
        /// <see cref="OutlookAI.Core.Com.MutationOutcome"/>. Carried separately because the
        /// prose is what a model reads and the field is what it can branch on, and only the
        /// site that knew the operation name can fill it in.
        /// </param>
        public ComHostUnavailableException(string message, string? outcome)
            : base(message)
        {
            Outcome = outcome;
        }

        /// <summary>Creates the exception with an inner cause.</summary>
        public ComHostUnavailableException(string message, Exception inner)
            : base(message, inner)
        {
        }

        /// <summary>
        /// Whether the interrupted request took effect, when the raising site could say.
        /// Null for the states where nothing was ever dispatched (no host, start backoff),
        /// which is deliberately not the same as claiming nothing changed.
        /// </summary>
        public string? Outcome { get; }
    }
}
