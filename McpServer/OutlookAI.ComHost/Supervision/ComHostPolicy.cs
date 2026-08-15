namespace OutlookAI.ComHost.Supervision
{
    /// <summary>Lifecycle state of the COM child as the parent understands it.</summary>
    internal enum ComHostState
    {
        /// <summary>No child process exists.</summary>
        None = 0,

        /// <summary>Spawned, handshake not yet complete.</summary>
        Starting = 1,

        /// <summary>Connected and serving operations.</summary>
        Ready = 2,

        /// <summary>Died, was killed, or failed its handshake. Awaiting a restart decision.</summary>
        Faulted = 3,
    }

    /// <summary>What the parent should do when a tool call needs the COM child.</summary>
    internal enum DispatchVerdict
    {
        /// <summary>A ready child exists - send the request.</summary>
        Dispatch = 0,

        /// <summary>No usable child - start one, then send.</summary>
        StartThenDispatch = 1,

        /// <summary>Too many recent start failures; refuse fast instead of thrashing.</summary>
        RefuseBackoff = 2,

        /// <summary>Starting Outlook is forbidden right now (installer mutex held, or autostart disabled).</summary>
        RefuseUnavailable = 3,
    }

    /// <summary>What the parent should do about a request that has not yet answered.</summary>
    internal enum InFlightVerdict
    {
        /// <summary>Still within budget - keep waiting.</summary>
        KeepWaiting = 0,

        /// <summary>Budget exhausted - kill the child and fail this request as a timeout.</summary>
        TimeoutKillChild = 1,

        /// <summary>The child process is gone - fail this request without a kill.</summary>
        FailChildDied = 2,

        /// <summary>The MCP client cancelled - stop waiting; the SDK suppresses the response.</summary>
        AbandonClientCancelled = 3,
    }

    /// <summary>Inputs to <see cref="ComHostPolicy.DecideDispatch"/>. A plain value bag so tests can enumerate it.</summary>
    internal readonly record struct DispatchInput(
        ComHostState State,
        int ConsecutiveStartFailures,
        long MillisecondsSinceLastStartFailure,
        bool StartingOutlookAllowed);

    /// <summary>Inputs to <see cref="ComHostPolicy.DecideInFlight"/>.</summary>
    internal readonly record struct InFlightInput(
        long ElapsedMilliseconds,
        long DeadlineMilliseconds,
        bool ChildAlive,
        bool ClientCancelled);

    /// <summary>
    /// The supervision decisions, as pure total functions of their inputs.
    /// <para>
    /// Written this way deliberately. The behaviour that matters most here - "a wedged
    /// Outlook call becomes a bounded, structured failure instead of silence" - is
    /// otherwise only observable by wedging a real Outlook, which is neither
    /// reproducible nor safe in CI. Keeping the decision pure lets T1 pin every branch
    /// with a synthetic clock, exactly as SweepWalkBoundsTests does for the sweep
    /// bounds, and leaves only the mechanical act of killing a process to the live tier.
    /// </para>
    /// </summary>
    internal static class ComHostPolicy
    {
        /// <summary>Consecutive failed starts before the parent stops retrying for a while.</summary>
        internal const int StartFailureBackoffThreshold = 3;

        /// <summary>How long to refuse further start attempts once the threshold is reached.</summary>
        internal const long StartBackoffMilliseconds = 30_000;

        /// <summary>
        /// Budget for an ordinary COM operation. Generous - a first search against a cold
        /// multi-store profile legitimately takes seconds, and a full-profile freshness
        /// sweep can take longer - but finite, which is the entire point.
        /// </summary>
        internal const long DefaultOperationDeadlineMilliseconds = 120_000;

        /// <summary>
        /// Budget for establishing the COM session (which may cold-start OUTLOOK.EXE).
        /// Measured on a healthy machine 2026-08-15: attach + health 1.0 s, cold search
        /// 6.2 s. A large OST on a slow disk is far slower, hence the wide margin.
        /// </summary>
        internal const long ConnectDeadlineMilliseconds = 90_000;

        /// <summary>
        /// Budget for the health probe. Short on purpose: outlook_health must answer even
        /// when Outlook is wedged, because that is precisely when it is asked. Exceeding
        /// this degrades the report, it never fails it.
        /// </summary>
        internal const long HealthProbeDeadlineMilliseconds = 5_000;

        /// <summary>Decides whether a tool call can be dispatched to the child.</summary>
        internal static DispatchVerdict DecideDispatch(DispatchInput input)
        {
            if (input.State == ComHostState.Ready)
            {
                return DispatchVerdict.Dispatch;
            }

            // Starting Outlook is gated by D17 (installer mutex / autostart disabled).
            // That gate applies only when we would have to START something; a child that
            // is already up is served regardless, which is why this sits below the Ready
            // check rather than above it.
            if (!input.StartingOutlookAllowed)
            {
                return DispatchVerdict.RefuseUnavailable;
            }

            if (IsInStartBackoff(input.ConsecutiveStartFailures, input.MillisecondsSinceLastStartFailure))
            {
                return DispatchVerdict.RefuseBackoff;
            }

            return DispatchVerdict.StartThenDispatch;
        }

        /// <summary>
        /// True while repeated start failures should suppress further attempts. Exposed so
        /// health reporting can explain the refusal rather than merely emitting it.
        /// </summary>
        internal static bool IsInStartBackoff(int consecutiveStartFailures, long millisecondsSinceLastStartFailure)
        {
            return consecutiveStartFailures >= StartFailureBackoffThreshold
                && millisecondsSinceLastStartFailure < StartBackoffMilliseconds;
        }

        /// <summary>Decides the fate of a request that has not yet answered.</summary>
        internal static InFlightVerdict DecideInFlight(InFlightInput input)
        {
            // Client cancellation outranks everything: the response is going to be
            // suppressed by the SDK regardless, so there is nothing to be gained by
            // killing a child that may still be perfectly healthy.
            if (input.ClientCancelled)
            {
                return InFlightVerdict.AbandonClientCancelled;
            }

            if (!input.ChildAlive)
            {
                return InFlightVerdict.FailChildDied;
            }

            if (input.ElapsedMilliseconds >= input.DeadlineMilliseconds)
            {
                return InFlightVerdict.TimeoutKillChild;
            }

            return InFlightVerdict.KeepWaiting;
        }

        /// <summary>
        /// The deadline for an operation, honouring an explicit per-call override.
        /// Non-positive overrides fall back to the default rather than meaning "instant"
        /// - a zero deadline would make every call fail before it began.
        /// </summary>
        internal static long DeadlineFor(ComHostOperationClass operationClass, long? overrideMilliseconds)
        {
            if (overrideMilliseconds is > 0)
            {
                return overrideMilliseconds.Value;
            }

            return operationClass switch
            {
                ComHostOperationClass.Connect => ConnectDeadlineMilliseconds,
                ComHostOperationClass.HealthProbe => HealthProbeDeadlineMilliseconds,
                _ => DefaultOperationDeadlineMilliseconds,
            };
        }
    }

    /// <summary>Deadline class of an operation.</summary>
    internal enum ComHostOperationClass
    {
        /// <summary>An ordinary mailbox operation.</summary>
        Operation = 0,

        /// <summary>Establishing the session, possibly cold-starting Outlook.</summary>
        Connect = 1,

        /// <summary>A health probe, which must degrade rather than block.</summary>
        HealthProbe = 2,
    }
}
