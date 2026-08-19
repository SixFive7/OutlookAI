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
    /// WHY THESE NUMBERS ARE LARGE (2026-08-19). The maintainer's instruction is the
    /// premise: "Outlook can sometimes be very slow. Also, I have +/- 50 GB of data, so
    /// searches can be slow. It doesn't matter all that much if operations are slow. If
    /// you use AI and thus our MCP server you have delegated this work and can perfectly
    /// wait 15 min." A tool call that FINISHES slowly is a result; a tool call that gives
    /// up is a coverage hole the caller has to work around. Every value here was re-sized
    /// against that, with one deliberate exception - <see cref="HealthProbeDeadlineMs"/>,
    /// which is the diagnostic run precisely when Outlook is wedged and therefore must
    /// stay short.
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
        /// <para>
        /// Raised from 120 000 to 300 000 on 2026-08-19. It is the HANG DETECTOR for every
        /// tool other than the exhaustive scan, so its job is to be unreachable by slow but
        /// working work and reachable by a wedge. Measured on the maintainer's real profile
        /// (5 stores, one of them 108 144 items): a whole-store 7-day sweep is 36.6 s and an
        /// Inbox-with-subfolders exhaustive scan is 66.5 s - so 120 s was under 2x the
        /// slowest healthy operation observed, which is not a hang detector, it is a second
        /// work limit wearing one's clothes. 300 s is roughly 4.5x that, and it also has to
        /// cover the composed search shape (<c>MailService.SearchBudgetMs</c>) with room to
        /// spare.
        /// </para>
        /// <para>
        /// The cost of raising it is real and is the reason it is not higher: a genuinely
        /// wedged Outlook now takes 5 min to be recognised, and
        /// <c>ComHostPolicy.UnresponsiveTimeoutThreshold</c> consecutive timeouts to open
        /// the breaker, so up to 10 min to fail fast. That cost is paid only by a wedge -
        /// and since 2026-08-19 an expiring caller-declared WORK budget no longer counts
        /// toward the breaker at all (<c>ComHostPolicy.TimeoutIndicatesUnresponsiveness</c>),
        /// so ordinary slow work cannot open it.
        /// </para>
        /// </summary>
        public const int OperationDeadlineMs = 300_000;

        /// <summary>
        /// Deadline for establishing the COM session, which may cold-start OUTLOOK.EXE.
        /// Measured 2026-08-15: attach + health 1.0 s, cold search 6.2 s; a large OST on a
        /// slow disk is far slower, hence the margin.
        /// <para>
        /// Raised from 90 000 to 180 000 on 2026-08-19, keeping it at 60% of
        /// <see cref="OperationDeadlineMs"/> as before. The measurement did not change; the
        /// margin over it did, for the same reason the operation deadline moved - a cold
        /// start on a 50 GB profile is the case nobody has timed, and paying two extra
        /// minutes once beats reporting a start failure that was really a slow disk.
        /// </para>
        /// </summary>
        public const int ConnectDeadlineMs = 180_000;

        /// <summary>
        /// Deadline for a health probe. Short on purpose: outlook_health must answer even
        /// when Outlook is wedged, because that is precisely when it is asked.
        /// <para>
        /// DELIBERATELY UNCHANGED while everything around it was raised (2026-08-19). It is
        /// the one budget whose expiry is the answer rather than a failure: a health check
        /// that also takes minutes turns every generous budget elsewhere into an unbounded
        /// wait with no way to find out why. It is the instrument, not the work.
        /// </para>
        /// </summary>
        public const int HealthProbeDeadlineMs = 5_000;

        /// <summary>
        /// The COM host pipe handshake, as seen from BOTH ends: the parent waits this long
        /// for the child to connect and report ready, and the child waits this long for
        /// the parent's pipe. One constant because it is one handshake - two independent
        /// declarations of "30 s" is how the two ends of a protocol drift apart.
        /// <para>
        /// Unchanged by the 2026-08-19 widening: this bounds starting a .NET child process
        /// and exchanging two frames, which has nothing to do with the size of a mailbox.
        /// </para>
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
        /// <para>
        /// It did NOT scale with the 2026-08-19 widening, and that is deliberate. What it
        /// covers is the size of one answer and the cost of one more unit of work, neither
        /// of which is a function of how long the work was allowed to run: the result set is
        /// still capped at <c>SearchTopCap</c>, and the sweep is still capped per folder and
        /// by <c>OutlookComSession.SweepBodyBytesBudget</c>. Scaling it "proportionally"
        /// would have reserved 77 s of a 615 s scan for a frame measured at 432 KB.
        /// </para>
        /// </summary>
        public const int ResultReturnHeadroomMs = 15_000;

        /// <summary>
        /// Hard deadline for the exhaustive scan, which has a deadline class of its own
        /// (<c>ComHostOperationClass.ExhaustiveScan</c>) rather than sharing
        /// <see cref="OperationDeadlineMs"/>.
        /// <para>
        /// WHY A CLASS AND NOT A BIGGER SHARED NUMBER. <c>exhaustive: true</c> is the one
        /// mode a caller picks BECAUSE completeness matters more than speed, and it is the
        /// one operation whose expiry is a normal documented answer rather than an incident.
        /// Dragging the shared deadline up to hold it would have made every other tool -
        /// read, new_draft, list_folders, move_mail - wait ten minutes before a wedged
        /// Outlook is reclaimed, and doubled that again before the breaker opened. The
        /// per-class mechanism already existed for <c>Connect</c> and <c>HealthProbe</c>;
        /// this is the third member and the reason the mechanism was worth having.
        /// </para>
        /// <para>
        /// The number is the answer to a measurement: on the maintainer's real profile a
        /// 60-day whole-store scan reached 3 folders of 32 before the old 105 s budget
        /// stopped it, so the budget was the binding constraint on completeness rather than
        /// a backstop. Ten minutes of scanning is roughly 6x that and is the figure the
        /// maintainer named.
        /// </para>
        /// </summary>
        public const int ExhaustiveScanDeadlineMs = 615_000;

        /// <summary>
        /// The most wall clock the exhaustive scan INSIDE the child may spend before it
        /// must stop and hand back whatever it has. Derived, never written as a literal:
        /// it is the scan's own hard deadline less the return trip, so a scan that runs
        /// long degrades to "results are partial" instead of to a timeout and a host kill.
        /// </summary>
        public const int ExhaustiveScanWorkBudgetMs = ExhaustiveScanDeadlineMs - ResultReturnHeadroomMs;
    }
}
