using System;

namespace OutlookAI.Services
{
    /// <summary>
    /// WHEN THE UPDATER IS ALLOWED TO TOUCH THE NETWORK: the first-check delay and the
    /// failure backoff, as pure arithmetic over a consecutive-failure count.
    ///
    /// <para>
    /// It lives in its own file, holding no state and touching nothing, for one reason: this
    /// is the only part of <see cref="UpdateService"/> that can be tested. The rest of that
    /// class is an <c>HttpClient</c>, a <c>System.Threading.Timer</c> and Outlook's process
    /// lifetime, none of which a test host has. This file is LINKED into
    /// <c>OutlookAI.McpServer.Tests</c> (T1 <c>UpdatePollScheduleTests</c>) the same way
    /// <c>AddInServerContract.cs</c> and <c>McpConfigEditor.cs</c> are, so the curve the
    /// add-in actually ships is the curve the suite pins - not a copy of it.
    /// </para>
    ///
    /// <para>
    /// FRAMEWORK-NEUTRAL and INTERNAL, under the same rules as the other linked files: it
    /// compiles as net48 (the add-in, C# 7.3) and as net10 (the test host, nullable-enabled
    /// with warnings as errors), so no nullable annotations and nothing newer than C# 7.3. A
    /// PUBLIC type in a file compiled into two assemblies that can see each other is CS0436,
    /// which is an error here.
    /// </para>
    /// </summary>
    internal static class UpdatePollSchedule
    {
        /// <summary>
        /// How often a healthy machine looks for an update. Every sentence in the product that
        /// says how often OutlookAI checks builds itself from this
        /// (<see cref="UpdateService.PollIntervalDescription"/>) rather than restating it.
        /// </summary>
        internal static readonly TimeSpan BaseInterval = TimeSpan.FromMinutes(10);

        /// <summary>
        /// HOW LONG THE UPDATER WAITS AFTER A DISRUPTIVE EVENT BEFORE TOUCHING THE NETWORK,
        /// and it is deliberately ONE number used at both of the two moments that qualify:
        /// the add-in loading, and the network coming back.
        ///
        /// <para>
        /// The first is what it was added for. The check used to fire on a
        /// <c>TimeSpan.Zero</c> due time, i.e. inside add-in load, which is the single busiest
        /// moment Outlook has: it is opening the user's mailbox, mounting delegate stores and
        /// running every other add-in's startup at the same time. The cost of one conditional
        /// GET is trivial in isolation and is not trivial there, and add-in startup delay is
        /// the thing users actually notice and actually blame - Outlook itself measures it and
        /// will disable a slow add-in.
        /// </para>
        ///
        /// <para>
        /// The second reuses it for the same reason rather than a different one: a network
        /// that has just come up cannot always resolve a name yet, so an immediate retry on
        /// the reconnect edge would usually just manufacture one more failure. Thirty seconds
        /// is long enough to be out of both storms and short enough that nobody waiting for an
        /// update notices it at all - the poll behind it is ten minutes wide.
        /// </para>
        /// </summary>
        internal static readonly TimeSpan SettleDelay = TimeSpan.FromSeconds(30);

        /// <summary>
        /// HOW MANY CONSECUTIVE FAILURES ARE TOLERATED AT FULL RATE before the interval starts
        /// doubling. Three, because one failure is noise and three is a fact: a Wi-Fi handoff,
        /// a VPN reconnect, a captive portal, a GitHub 5xx or a rate-limit answer all produce
        /// one or two, and all of them clear inside the ~20 minutes three checks span. Keeping
        /// those at full rate means a blip never costs the user a late update, and it is the
        /// only cost of not backing off from the first failure.
        /// </summary>
        internal const int FailuresBeforeBackoff = 3;

        /// <summary>
        /// THE CEILING ON THE BACKOFF, and therefore the worst case for how long a machine
        /// that quietly regains connectivity can sit before it notices - so it is chosen
        /// against recovery, not against bandwidth.
        ///
        /// <para>
        /// Two hours, twelve times <see cref="BaseInterval"/>. It bounds the daily traffic of
        /// a permanently offline machine at 12 conditional GETs instead of 144, which is the
        /// point of the whole exercise: the bandwidth was never the problem, a machine
        /// endlessly failing to reach an external host on a managed network is, and so is the
        /// log it fills. And two hours is short enough that the worst case is an inconvenience
        /// rather than a stall.
        /// </para>
        ///
        /// <para>
        /// The worst case is rarer than the cap suggests, because the cap is the SECOND
        /// recovery path. The first is <c>UpdateService</c>'s network-availability hook, which
        /// clears the failure count the moment Windows says the machine is back on a network
        /// and checks one <see cref="SettleDelay"/> later. What is left for this cap to cover
        /// is recovery WITHOUT a local network change: a proxy that starts answering, a
        /// firewall rule that lands, a captive portal that gets paid for, GitHub coming back
        /// up. And the user has an immediate third path at any time - "Check for updates" in
        /// the sidebar and in Settings runs now and restarts the poll clock from there.
        /// </para>
        /// </summary>
        internal static readonly TimeSpan MaxInterval = TimeSpan.FromHours(2);

        /// <summary>
        /// Guards <c>2^doublings</c> against a failure count that has been climbing for months.
        /// At <see cref="BaseInterval"/> = 10 minutes the cap is reached at 4 doublings, so
        /// anything past this is already clamped; the bound exists so the intermediate
        /// <c>Math.Pow</c> cannot reach infinity and hand <c>TimeSpan.FromMinutes</c> a value
        /// it throws on.
        /// </summary>
        private const int MaxDoublings = 16;

        /// <summary>
        /// How long to wait before the next check, given how many checks in a row have failed
        /// to reach the update server. Zero failures - the healthy case, and the case a
        /// successful check resets to - is <see cref="BaseInterval"/>.
        ///
        /// <para>
        /// The curve: full rate for the first <see cref="FailuresBeforeBackoff"/> failures,
        /// then double per further failure, clamped at <see cref="MaxInterval"/>. With the
        /// shipped numbers that is 10, 10, 20, 40, 80, 120, 120, … minutes, so a machine that
        /// cannot reach GitHub at all makes six attempts in its first three hours and one
        /// every two hours after that. Doubling rather than a fixed penalty interval because
        /// the two failure shapes want different answers and this is the only curve that
        /// serves both: a five-minute outage should barely be noticed, and a fortnight
        /// offline should cost almost nothing.
        /// </para>
        /// </summary>
        /// <param name="consecutiveFailures">
        /// Checks that have failed to reach the server since the last one that did. Negative
        /// values are treated as zero rather than rejected - this decides a timer's due time,
        /// and there is no caller for whom throwing here would be better than polling.
        /// </param>
        internal static TimeSpan DelayAfter(int consecutiveFailures)
        {
            if (consecutiveFailures < FailuresBeforeBackoff)
                return BaseInterval;

            int doublings = consecutiveFailures - FailuresBeforeBackoff + 1;
            if (doublings > MaxDoublings)
                doublings = MaxDoublings;

            double minutes = BaseInterval.TotalMinutes * Math.Pow(2, doublings);
            return minutes >= MaxInterval.TotalMinutes ? MaxInterval : TimeSpan.FromMinutes(minutes);
        }

        /// <summary>
        /// Whether the backoff is engaged at this failure count - i.e. whether the interval in
        /// force is longer than <see cref="BaseInterval"/>. Exists so a caller can say so
        /// without re-deriving the threshold.
        /// </summary>
        internal static bool IsBackingOff(int consecutiveFailures)
        {
            return consecutiveFailures >= FailuresBeforeBackoff;
        }
    }
}
