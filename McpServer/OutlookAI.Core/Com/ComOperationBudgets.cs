namespace OutlookAI.Core.Com
{
    /// <summary>
    /// The wall-clock budgets that BOTH sides of the COM-host split have to agree on, in
    /// one place.
    /// <para>
    /// They live in Core rather than beside the supervisor because the dependency only
    /// runs one way: <c>OutlookAI.ComHost</c> references <c>OutlookAI.Core</c>, never the
    /// reverse. So the service layer (which owns the INNER budgets - the exhaustive scan,
    /// the freshness sweep) cannot see <c>ComHostPolicy</c>, and before this type existed
    /// the two tiers each wrote their own literal and happened to agree. That is exactly
    /// how the composition defects below were introduced:
    /// </para>
    /// <para>
    /// 1. The exhaustive scan's soft budget was written as its own <c>120_000</c> and so
    /// EQUALLED the outer hard deadline. An inner budget that equals its outer one can
    /// never degrade gracefully: the scan stops once elapsed has PASSED the budget and
    /// then still has to serialize its result set back across the pipe, while the
    /// watchdog fires at <c>&gt;=</c>. The documented "results are partial" outcome was
    /// therefore unreachable whenever the scan actually ran long - the caller got a
    /// Timeout, the host was killed, and two of those open the breaker. The inner budget
    /// is now DERIVED as the outer one minus <see cref="ResultReturnHeadroomMs"/>.
    /// </para>
    /// <para>
    /// 2. The 5 s health budget and the 30 s pipe handshake were each declared
    /// independently at both ends. Nothing related them, so either half could be raised
    /// alone.
    /// </para>
    /// <para>
    /// Nothing here is a tuning knob to be edited casually: every value is either measured
    /// (see <c>ComHostPolicy</c> for the measurements) or derived from one that is.
    /// </para>
    /// </summary>
    public static class ComOperationBudgets
    {
        /// <summary>
        /// Hard deadline for one ordinary COM contract call, after which the host is
        /// killed. Generous but finite; <c>ComHostPolicy.DefaultOperationDeadlineMilliseconds</c>
        /// is this value and carries the full rationale.
        /// </summary>
        public const int OperationDeadlineMs = 120_000;

        /// <summary>
        /// Deadline for establishing the COM session, which may cold-start OUTLOOK.EXE.
        /// Measured 2026-08-15: attach + health 1.0 s, cold search 6.2 s; a large OST on a
        /// slow disk is far slower, hence the margin.
        /// </summary>
        public const int ConnectDeadlineMs = 90_000;

        /// <summary>
        /// Deadline for a health probe. Short on purpose: outlook_health must answer even
        /// when Outlook is wedged, because that is precisely when it is asked.
        /// </summary>
        public const int HealthProbeDeadlineMs = 5_000;

        /// <summary>
        /// The COM host pipe handshake, as seen from BOTH ends: the parent waits this long
        /// for the child to connect and report ready, and the child waits this long for
        /// the parent's pipe. One constant because it is one handshake - two independent
        /// declarations of "30 s" is how the two ends of a protocol drift apart.
        /// </summary>
        public const int HandshakeBudgetMs = 30_000;

        /// <summary>
        /// How much of an operation's deadline is reserved for handing the result BACK -
        /// serializing the result set, writing the frame, and the overshoot of an inner
        /// loop that can only check its budget between units of work.
        /// <para>
        /// Sized against the largest thing that crosses the pipe on this path: an
        /// exhaustive scan returning up to <c>MailService.SearchTopCap</c> briefs, plus one
        /// more folder's <c>Restrict</c> after the last budget check. It is deliberately
        /// far larger than the measured serialize cost - the point is that the caller sees
        /// the tool's own documented partial-results answer instead of a host kill, and
        /// buying that certainty with a few seconds of scan time is the right trade.
        /// </para>
        /// </summary>
        public const int ResultReturnHeadroomMs = 15_000;

        /// <summary>
        /// The most wall clock a single unit of work INSIDE the child may spend before it
        /// must stop and hand back whatever it has. Derived, never written as a literal:
        /// it is the outer deadline less the return trip, so a scan that runs long
        /// degrades to "results are partial" instead of to a timeout and a host kill.
        /// </summary>
        public const int ChildWorkBudgetMs = OperationDeadlineMs - ResultReturnHeadroomMs;
    }
}
