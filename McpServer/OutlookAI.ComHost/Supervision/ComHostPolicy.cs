using OutlookAI.Core.Com;

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

    /// <summary>What to do about an Outlook that has repeatedly failed to answer.</summary>
    internal enum BreakerVerdict
    {
        /// <summary>Outlook is behaving; send the request.</summary>
        Closed = 0,

        /// <summary>Outlook is known unresponsive and the cooldown has not elapsed; fail immediately.</summary>
        Open = 1,

        /// <summary>Cooldown elapsed; run one cheap liveness probe before committing a full request.</summary>
        HalfOpen = 2,
    }

    /// <summary>Inputs to <see cref="ComHostPolicy.DecideBreaker"/>.</summary>
    internal readonly record struct BreakerInput(
        int ConsecutiveTimeouts,
        long MillisecondsSinceLastTimeout);

    /// <summary>What Outlook's externally observable state implies for a request.</summary>
    internal enum LivenessVerdict
    {
        /// <summary>Outlook is up and pumping - go ahead.</summary>
        Proceed = 0,

        /// <summary>Outlook is up but not answering. Fail immediately; a COM call would not return.</summary>
        Hung = 1,

        /// <summary>Outlook is coming up. Tell the caller to retry shortly rather than blocking it.</summary>
        Starting = 2,

        /// <summary>Outlook is not running and may be started now.</summary>
        MayStart = 3,

        /// <summary>
        /// Outlook is not running, but we tried to start one very recently. Starting again
        /// now is the churn that appears to wedge it - wait instead.
        /// </summary>
        StartSuppressed = 4,
    }

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
        /// <para>
        /// Derived from <see cref="ComOperationBudgets"/> rather than written here, because
        /// the service layer sizes its INNER budgets against this number and cannot see
        /// this assembly (Core has no dependency on ComHost). Two literals that happen to
        /// agree is what produced the exhaustive-scan composition defect.
        /// </para>
        /// </summary>
        internal const long DefaultOperationDeadlineMilliseconds = ComOperationBudgets.OperationDeadlineMs;

        /// <summary>
        /// Budget for establishing the COM session (which may cold-start OUTLOOK.EXE).
        /// Measured on a healthy machine 2026-08-15: attach + health 1.0 s, cold search
        /// 6.2 s. A large OST on a slow disk is far slower, hence the wide margin.
        /// </summary>
        internal const long ConnectDeadlineMilliseconds = ComOperationBudgets.ConnectDeadlineMs;

        /// <summary>
        /// Budget for the health probe. Short on purpose: outlook_health must answer even
        /// when Outlook is wedged, because that is precisely when it is asked. Exceeding
        /// this degrades the report, it never fails it.
        /// <para>
        /// The service layer's <c>MailService.HealthProbeBudgetMs</c> is the same constant,
        /// not a second copy of 5 000.
        /// </para>
        /// </summary>
        internal const long HealthProbeDeadlineMilliseconds = ComOperationBudgets.HealthProbeDeadlineMs;

        /// <summary>
        /// Ceiling on the COM host pipe handshake - the parent's wait for the child to
        /// connect AND report ready, shared with the child's own wait for the parent's pipe
        /// (<c>Program.ConnectTimeoutMs</c>). One handshake, one number.
        /// </summary>
        internal const long HandshakeBudgetMilliseconds = ComOperationBudgets.HandshakeBudgetMs;

        /// <summary>
        /// Floor under the handshake budget when the operation's own deadline is shorter.
        /// <para>
        /// The handshake used to sit entirely outside the deadline system: it was consumed
        /// TWICE (once waiting for the pipe, once waiting for the ready event), so a slow
        /// child start could cost 60 s before any budget applied, and
        /// <see cref="DeadlineVariable"/> could not shorten it. It is now one shared budget
        /// that follows the operation deadline - but never below this floor, because
        /// starting a fresh .NET child on a loaded box legitimately takes seconds and the
        /// only caller that sets a shorter deadline is the test suite, which is testing the
        /// timeout path rather than the start path.
        /// </para>
        /// </summary>
        internal const long HandshakeFloorMilliseconds = 10_000;

        /// <summary>
        /// Floor under any deadline actually sent to the child. A remaining aggregate of a
        /// few milliseconds must not turn into a deadline that kills a perfectly healthy
        /// host before it can answer; below this the caller is told the aggregate is
        /// exhausted instead.
        /// </summary>
        internal const long MinimumDispatchDeadlineMilliseconds = 1_000;

        /// <summary>
        /// How long the child start handshake may take, given the deadline of the operation
        /// that triggered it. Pure so T1 can pin the boundaries.
        /// </summary>
        internal static long HandshakeBudgetFor(long operationDeadlineMilliseconds)
        {
            if (operationDeadlineMilliseconds <= HandshakeFloorMilliseconds)
            {
                return HandshakeFloorMilliseconds;
            }

            return operationDeadlineMilliseconds < HandshakeBudgetMilliseconds
                ? operationDeadlineMilliseconds
                : HandshakeBudgetMilliseconds;
        }

        /// <summary>
        /// The deadline one contract call may actually use, given its own budget and how
        /// much of the enclosing operation's AGGREGATE budget is left.
        /// <para>
        /// The per-call deadline bounds ONE round trip. A gateway operation is a lambda
        /// that may make many - hit location makes 1 + up to 3 + N, the archive path walks
        /// every store - and before this each of those independently got a full budget, so
        /// the operation as a whole had no bound at all. The aggregate is measured from the
        /// start of the lambda and shrinks every call after it.
        /// </para>
        /// <para>
        /// Returns 0 to mean "the aggregate is spent, do not dispatch": a sub-second
        /// deadline would kill a healthy host rather than bound a wedged one, so anything
        /// below <see cref="MinimumDispatchDeadlineMilliseconds"/> is reported as exhausted.
        /// A null aggregate means no enclosing operation declared one.
        /// </para>
        /// <para>
        /// The surviving aggregate is rounded UP to a whole second before it clamps.
        /// Every budget in this system is expressed in whole seconds, and the timeout
        /// message quotes the deadline back to a human ("exceeded its 4000 ms budget"): a
        /// first call must not be told 3 999 merely because a millisecond of its own
        /// aggregate has elapsed since the lambda started. The slack this concedes is under
        /// one second per call, against a defect measured in whole extra budgets.
        /// </para>
        /// </summary>
        internal static long EffectiveDeadlineMilliseconds(long callDeadlineMilliseconds, long? remainingAggregateMilliseconds)
        {
            if (remainingAggregateMilliseconds is not { } remaining)
            {
                return callDeadlineMilliseconds;
            }

            if (remaining < MinimumDispatchDeadlineMilliseconds)
            {
                return 0;
            }

            long wholeSeconds = ((remaining + 999) / 1000) * 1000;
            return wholeSeconds < callDeadlineMilliseconds ? wholeSeconds : callDeadlineMilliseconds;
        }

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

        /// <summary>
        /// Minimum gap between attempts to start Outlook.
        /// <para>
        /// This is the anti-churn guard, and it exists because of a specific root-cause
        /// finding on 2026-08-16. The wedged Outlook was one WE started
        /// (<c>OUTLOOK.EXE -Embedding</c>, parent svchost - i.e. COM activation), and the
        /// event log showed two Outlook starts 39 seconds apart, the second of which hung
        /// after loading add-ins and never made a single network call.
        /// </para>
        /// <para>
        /// The suspected mechanism is our own doing: a timeout kills the COM host, that
        /// was the last COM client of an Outlook we had started headlessly, so Outlook
        /// begins shutting down - and the very next request activates it again while the
        /// previous instance is still exiting. Starting Outlook on top of an exiting one
        /// is a well-known way to get a half-initialised, wedged process.
        /// </para>
        /// </summary>
        internal const long AutostartCooldownMilliseconds = 20_000;

        /// <summary>How long to tell a caller to wait while Outlook is starting.</summary>
        internal const int StartingRetryAfterSeconds = 15;

        /// <summary>How long to tell a caller to wait while Outlook is unresponsive.</summary>
        internal const int UnresponsiveRetryAfterSeconds = 30;

        /// <summary>
        /// Decides what Outlook's externally observed state means for a request, before any
        /// COM is attempted.
        /// <para>
        /// Asking Windows first is close to free and replaces a 30-120 s discovery with a
        /// microsecond one: it already knows whether a window's thread is servicing its
        /// message queue.
        /// </para>
        /// </summary>
        internal static LivenessVerdict DecideLiveness(
            OutlookLivenessState liveness,
            long millisecondsSinceLastStartAttempt,
            bool startingOutlookAllowed)
        {
            switch (liveness)
            {
                case OutlookLivenessState.Responsive:
                    return LivenessVerdict.Proceed;

                case OutlookLivenessState.Hung:
                    return LivenessVerdict.Hung;

                case OutlookLivenessState.Starting:
                    return LivenessVerdict.Starting;

                case OutlookLivenessState.NotRunning:
                default:
                    if (!startingOutlookAllowed)
                    {
                        // Caller maps this to the existing "may not start Outlook" refusal.
                        return LivenessVerdict.StartSuppressed;
                    }

                    return millisecondsSinceLastStartAttempt < AutostartCooldownMilliseconds
                        ? LivenessVerdict.StartSuppressed
                        : LivenessVerdict.MayStart;
            }
        }

        /// <summary>Retry guidance, in seconds, for a verdict that asks the caller to come back.</summary>
        internal static int RetryAfterSecondsFor(LivenessVerdict verdict, long millisecondsSinceLastStartAttempt)
        {
            switch (verdict)
            {
                case LivenessVerdict.Starting:
                    return StartingRetryAfterSeconds;

                case LivenessVerdict.StartSuppressed:
                    long remaining = AutostartCooldownMilliseconds - millisecondsSinceLastStartAttempt;
                    if (remaining < 1_000)
                    {
                        remaining = 1_000;
                    }

                    return (int)((remaining + 999) / 1000);

                case LivenessVerdict.Hung:
                    return UnresponsiveRetryAfterSeconds;

                default:
                    return 0;
            }
        }

        /// <summary>Consecutive operation timeouts before the parent stops sending full requests.</summary>
        internal const int UnresponsiveTimeoutThreshold = 2;

        /// <summary>How long to keep failing fast before re-probing Outlook.</summary>
        internal const long UnresponsiveCooldownMilliseconds = 30_000;

        /// <summary>
        /// Decides whether to send a request at all, given how Outlook has been behaving.
        /// <para>
        /// Bounding each call individually is necessary but not sufficient. Against an
        /// Outlook that is persistently wedged, every request independently pays its full
        /// budget - measured on this machine at 120 s for search, list_accounts and
        /// list_folders - and each one spawns and kills a child to learn what the previous
        /// one already established. The tenth search in a row should not cost two minutes
        /// to rediscover that Outlook is not answering.
        /// </para>
        /// <para>
        /// So the supervisor remembers. After
        /// <see cref="UnresponsiveTimeoutThreshold"/> consecutive timeouts it fails COM
        /// requests immediately for <see cref="UnresponsiveCooldownMilliseconds"/>, then
        /// allows one cheap liveness probe. Any success closes it again, so a user who
        /// restarts Outlook is picked up within one cooldown rather than having to wait
        /// for a full-budget request to succeed.
        /// </para>
        /// <para>
        /// Crucially this makes SEARCH good rather than merely survivable: with the
        /// breaker open the freshness sweep fails in microseconds, so search returns its
        /// indexed results immediately with advice, instead of stalling two minutes first.
        /// </para>
        /// </summary>
        internal static BreakerVerdict DecideBreaker(BreakerInput input)
        {
            if (input.ConsecutiveTimeouts < UnresponsiveTimeoutThreshold)
            {
                return BreakerVerdict.Closed;
            }

            return input.MillisecondsSinceLastTimeout >= UnresponsiveCooldownMilliseconds
                ? BreakerVerdict.HalfOpen
                : BreakerVerdict.Open;
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

            // When the test override is set it governs EVERY class, including connect.
            // Otherwise the cold-start floor below would silently reimpose the 90 s
            // allowance and any test of the timeout path would take 90 s to observe it.
            if (ConfiguredDefaultDeadline is > 0)
            {
                return ConfiguredDefaultDeadline.Value;
            }

            return operationClass switch
            {
                ComHostOperationClass.Connect => ConnectDeadlineMilliseconds,
                ComHostOperationClass.HealthProbe => HealthProbeDeadlineMilliseconds,
                _ => DefaultOperationDeadlineMilliseconds,
            };
        }

        /// <summary>
        /// Environment override for the ordinary operation budget. Its purpose is testing:
        /// the timeout path is only observable by waiting out a deadline, and waiting out
        /// the real two-minute budget in every such test would make the suite unusable.
        /// Unset in production.
        /// </summary>
        internal const string DeadlineVariable = "OUTLOOKAI_COMHOST_DEADLINE_MS";

        private static readonly long? ConfiguredDefaultDeadline = ReadConfiguredDeadline();

        /// <summary>
        /// Floor applied to the first operation on a fresh child, which also pays for
        /// establishing the COM session and possibly cold-starting Outlook. Follows the
        /// test override when one is set.
        /// </summary>
        internal static long ConnectFloorMilliseconds =>
            ConfiguredDefaultDeadline is > 0 ? ConfiguredDefaultDeadline.Value : ConnectDeadlineMilliseconds;

        private static long? ReadConfiguredDeadline()
        {
            string? raw = Environment.GetEnvironmentVariable(DeadlineVariable);
            if (long.TryParse(raw, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out long value)
                && value > 0)
            {
                return value;
            }

            return null;
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
