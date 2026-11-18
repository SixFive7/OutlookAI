using System;
using System.Globalization;

namespace OutlookAI.Core.Com
{
    /// <summary>What Outlook's start looks like from outside, for a request that needs Outlook.</summary>
    public enum OutlookStartState
    {
        /// <summary>No OUTLOOK.EXE runs, and no OutlookAI connection is starting one.</summary>
        NotRunning = 0,

        /// <summary>Outlook is starting, or has not settled since it started: wait, and open no mailbox yet.</summary>
        Settling = 1,

        /// <summary>Outlook has settled - or started long enough ago that it counts as settled.</summary>
        Ready = 2,

        /// <summary>Outlook is not responding, or has not finished starting within the limit.</summary>
        NotResponding = 3,
    }

    /// <summary>What kind of "not responding" an Outlook is in.</summary>
    public enum OutlookStuckKind
    {
        /// <summary>Not stuck.</summary>
        None = 0,

        /// <summary>Windows reports every Outlook UI window as not responding.</summary>
        Hung = 1,

        /// <summary>Running past the starting limit and still without its windows: its start never finished.</summary>
        NoWindows = 2,

        /// <summary>An OutlookAI connection has waited on Outlook in one start-up step past the starting limit.</summary>
        ConnectStuck = 3,
    }

    /// <summary>
    /// The facts one start decision rests on, as plain values so T1 can enumerate them. Gathered without
    /// COM by <see cref="OutlookStartObserver"/>.
    /// </summary>
    public sealed class OutlookStartFacts
    {
        /// <summary>Creates the facts.</summary>
        /// <param name="liveness">Windows' view of Outlook (<see cref="OutlookLiveness"/>).</param>
        /// <param name="ageMilliseconds">How long ago OUTLOOK.EXE started; null when not running or unreadable.</param>
        /// <param name="addInStamped">
        /// Whether the OutlookAI add-in has finished starting inside this Outlook; null when no add-in is
        /// set to load, so no stamp is coming.
        /// </param>
        /// <param name="mainThreadIdle">
        /// Whether Outlook's main thread was idle when measured; null when not measured. The gate in front
        /// of every request does not measure it - it takes a sample of a second - and leaves it to the
        /// connection, which measures it before it opens a mailbox.
        /// </param>
        /// <param name="connectLockHeldElsewhere">Whether another connection holds the start lock (<see cref="OutlookStartLock"/>).</param>
        /// <param name="connectStepMilliseconds">
        /// How long that connection has been in its current start-up step, from the start record; null
        /// when no connection holds the lock or the record does not say.
        /// </param>
        /// <param name="finishedStartingBefore">
        /// Whether this Outlook got through its start once already - its add-in stamped it, a window of it
        /// was seen, or a connection got past the settle wait in it. One with no windows then is closing,
        /// not stuck starting.
        /// </param>
        public OutlookStartFacts(
            OutlookLivenessState liveness,
            long? ageMilliseconds,
            bool? addInStamped,
            bool? mainThreadIdle,
            bool connectLockHeldElsewhere,
            long? connectStepMilliseconds,
            bool finishedStartingBefore = false)
        {
            Liveness = liveness;
            AgeMilliseconds = ageMilliseconds;
            AddInStamped = addInStamped;
            MainThreadIdle = mainThreadIdle;
            ConnectLockHeldElsewhere = connectLockHeldElsewhere;
            ConnectStepMilliseconds = connectStepMilliseconds;
            FinishedStartingBefore = finishedStartingBefore;
        }

        /// <summary>Windows' view of Outlook.</summary>
        public OutlookLivenessState Liveness { get; }

        /// <summary>How long ago OUTLOOK.EXE started; null when not running or unreadable.</summary>
        public long? AgeMilliseconds { get; }

        /// <summary>Whether the add-in has stamped this Outlook's start; null when no stamp is coming.</summary>
        public bool? AddInStamped { get; }

        /// <summary>Whether Outlook's main thread was idle; null when not measured.</summary>
        public bool? MainThreadIdle { get; }

        /// <summary>Whether another connection holds the start lock.</summary>
        public bool ConnectLockHeldElsewhere { get; }

        /// <summary>How long that connection has been in its current step; null when unknown.</summary>
        public long? ConnectStepMilliseconds { get; }

        /// <summary>Whether this Outlook got through its start once already.</summary>
        public bool FinishedStartingBefore { get; }
    }

    /// <summary>One start decision: the state, why, and when a caller should come back.</summary>
    public sealed class OutlookStartVerdict
    {
        internal OutlookStartVerdict(OutlookStartState state, OutlookStuckKind stuck, string reasonCode, string reason, int retryAfterSeconds)
        {
            State = state;
            Stuck = stuck;
            ReasonCode = reasonCode;
            Reason = reason;
            RetryAfterSeconds = retryAfterSeconds;
        }

        /// <summary>The state.</summary>
        public OutlookStartState State { get; }

        /// <summary>What kind of stuck, when <see cref="State"/> is <see cref="OutlookStartState.NotResponding"/>.</summary>
        public OutlookStuckKind Stuck { get; }

        /// <summary>A short token for the start log - never a sentence, so the log stays machine-readable.</summary>
        public string ReasonCode { get; }

        /// <summary>The same, in words, for a caller and for health.</summary>
        public string Reason { get; }

        /// <summary>When a caller should retry, in seconds; 0 for <see cref="OutlookStartState.Ready"/>.</summary>
        public int RetryAfterSeconds { get; }

        /// <summary>True when Outlook counts as settled only because it started long enough ago.</summary>
        public bool SettledByAge => State == OutlookStartState.Ready
            && string.Equals(ReasonCode, OutlookStartPolicy.SettledByAgeCode, StringComparison.Ordinal);
    }

    /// <summary>
    /// When OutlookAI may open a mailbox in an Outlook that has just started, and when an Outlook that is
    /// starting has become one that is stuck. Pure and total, so T1 pins every branch with plain values.
    /// <para>
    /// <b>Why it exists - the hang of 2026-11-15.</b> OutlookAI's server started Outlook over COM, which
    /// starts it with no window, and opened mailboxes within the first second, while Outlook's add-ins were
    /// still starting. The Teams Meeting add-in called back into Outlook from a background thread while
    /// Outlook was opening an Exchange store for OutlookAI, and Outlook's main thread deadlocked on itself.
    /// The user could neither see that Outlook nor close it, and every launch of Outlook was handed to it.
    /// The maintainer decided on 2026-11-18 (Q175 (a)) that OutlookAI waits for Outlook to settle before it
    /// opens a mailbox: Outlook at least about 20 s old, the add-in's start-up stamp present, and Outlook's
    /// main thread idle. That narrows the race; it cannot close it, because another add-in's own calls are
    /// outside OutlookAI's reach.
    /// </para>
    /// <para>
    /// <b>And when starting has gone on too long</b> (Q176 (c)): an Outlook that has not created its windows
    /// <see cref="StartingLimitMilliseconds"/> after it started, or an OutlookAI connection that has waited
    /// that long in one step, is reported as not responding, where it used to read as "starting" for ever.
    /// An Outlook that responds but has not met every settle condition by then - its add-in disabled, say -
    /// is used anyway: the conditions narrow a start-up race, and there is no start-up left to race.
    /// </para>
    /// </summary>
    public static class OutlookStartPolicy
    {
        /// <summary>How old Outlook must be before a mailbox is opened in it (Q175 (a): "at least about 20 s old").</summary>
        public const int MinimumAgeMilliseconds = 20_000;

        /// <summary>
        /// How long "starting" may last (Q176 (c): "about 120 s"). The connect budget is 180 s
        /// (<see cref="ComOperationBudgets.ConnectDeadlineMs"/>), so a connection that gives up at this limit
        /// still answers with a report of its own before the server would reclaim it as a timeout.
        /// </summary>
        public const int StartingLimitMilliseconds = 120_000;

        /// <summary>
        /// How long a connection waits for another one to finish starting or attaching to Outlook (Q175 (e):
        /// one starter at a time, the others wait). Longer than <see cref="StartingLimitMilliseconds"/> -
        /// the holder may spend that long settling - and shorter than the 180 s connect budget.
        /// </summary>
        public const int ConnectLockWaitMilliseconds = 150_000;

        /// <summary>How long Outlook's main thread is watched for being idle.</summary>
        public const int IdleSampleMilliseconds = 500;

        /// <summary>
        /// The most CPU time the main thread may use over <see cref="IdleSampleMilliseconds"/> and still count
        /// as idle: a tenth of the sample. A thread waiting for messages uses none; one loading add-ins or
        /// opening stores uses most of it.
        /// </summary>
        public const int IdleCpuMilliseconds = 50;

        /// <summary>How long the main thread is given to answer a message before it counts as busy.</summary>
        public const int MainThreadReplyMilliseconds = 1_000;

        /// <summary>How often a connection looks again while it waits for Outlook to settle.</summary>
        public const int SettlePollMilliseconds = 500;

        /// <summary>
        /// The retry guidance for a request that has just started Outlook: the settle wait plus the
        /// connection after it. It was 15 s before the settle wait existed.
        /// </summary>
        public const int StartRetryAfterSeconds = 25;

        /// <summary>The retry guidance while Outlook is not responding.</summary>
        public const int NotRespondingRetryAfterSeconds = 30;

        /// <summary>The reason code of an Outlook that counts as settled because it started long enough ago.</summary>
        public const string SettledByAgeCode = "settled-by-age";

        /// <summary>The reason code of an Outlook that met every settle condition.</summary>
        public const string SettledCode = "settled";

        /// <summary>Decides what <paramref name="facts"/> mean for a request that needs Outlook.</summary>
        public static OutlookStartVerdict Decide(OutlookStartFacts facts)
        {
            if (facts == null)
            {
                throw new ArgumentNullException(nameof(facts));
            }

            bool lockStuck = facts.ConnectLockHeldElsewhere
                && facts.ConnectStepMilliseconds is long step
                && step >= StartingLimitMilliseconds;

            if (facts.Liveness == OutlookLivenessState.NotRunning)
            {
                if (lockStuck)
                {
                    return Stuck(OutlookStuckKind.ConnectStuck, "connect-stuck",
                        "an OutlookAI connection has been trying to start Outlook for " + Seconds(facts.ConnectStepMilliseconds)
                        + " s and no OUTLOOK.EXE is running");
                }

                return facts.ConnectLockHeldElsewhere
                    ? Settling("starting-elsewhere", "another OutlookAI server is starting Outlook", StartRetryAfterSeconds)
                    : new OutlookStartVerdict(OutlookStartState.NotRunning, OutlookStuckKind.None, "not-running", "Outlook is not running", 0);
            }

            if (facts.Liveness == OutlookLivenessState.Hung)
            {
                return Stuck(OutlookStuckKind.Hung, "hung", "Windows reports every Outlook window as not responding");
            }

            if (lockStuck)
            {
                return Stuck(OutlookStuckKind.ConnectStuck, "connect-stuck",
                    "an OutlookAI connection has been waiting on Outlook in one start-up step for "
                    + Seconds(facts.ConnectStepMilliseconds) + " s");
            }

            if (facts.Liveness == OutlookLivenessState.Starting && facts.FinishedStartingBefore)
            {
                // Windows gone from an Outlook that had them: it is shutting down, as after the user closes
                // it. Wait for it to go; the request after that starts or finds the next one.
                return Settling("closing", "Outlook has closed its windows and is shutting down", 10);
            }

            if (facts.AgeMilliseconds is long age && age >= StartingLimitMilliseconds)
            {
                if (facts.Liveness == OutlookLivenessState.Starting)
                {
                    return Stuck(OutlookStuckKind.NoWindows, "no-windows",
                        "Outlook started " + Seconds(age) + " s ago and has still not created its windows");
                }

                return new OutlookStartVerdict(OutlookStartState.Ready, OutlookStuckKind.None, SettledByAgeCode,
                    "Outlook started " + Seconds(age) + " s ago", 0);
            }

            if (facts.Liveness == OutlookLivenessState.Starting)
            {
                return Settling("no-windows-yet", "Outlook has not created its windows yet", SettlingRetryAfterSeconds(facts.AgeMilliseconds));
            }

            if (facts.ConnectLockHeldElsewhere)
            {
                return Settling("connect-elsewhere", "another OutlookAI connection is starting or attaching to Outlook",
                    SettlingRetryAfterSeconds(facts.AgeMilliseconds));
            }

            if (facts.AgeMilliseconds is long young && young < MinimumAgeMilliseconds)
            {
                return Settling("young", "Outlook started " + Seconds(young) + " s ago; a mailbox is opened once it is "
                    + Seconds(MinimumAgeMilliseconds) + " s old", SettlingRetryAfterSeconds(young));
            }

            if (facts.AddInStamped == false)
            {
                return Settling("add-in-starting", "the OutlookAI add-in has not finished starting inside Outlook",
                    SettlingRetryAfterSeconds(facts.AgeMilliseconds));
            }

            if (facts.MainThreadIdle == false)
            {
                return Settling("main-thread-busy", "Outlook's main thread is still busy",
                    SettlingRetryAfterSeconds(facts.AgeMilliseconds));
            }

            return new OutlookStartVerdict(OutlookStartState.Ready, OutlookStuckKind.None, SettledCode, "Outlook has settled", 0);
        }

        /// <summary>
        /// When to come back to an Outlook that is settling: the rest of the minimum age plus a few seconds
        /// for the connection, or 10 s once the age is not what is being waited for. Never under 5 s or over
        /// 30 s - a number an agent will actually wait out.
        /// </summary>
        public static int SettlingRetryAfterSeconds(long? ageMilliseconds)
        {
            int seconds = 10;
            if (ageMilliseconds is long age && age < MinimumAgeMilliseconds)
            {
                seconds = (int)((MinimumAgeMilliseconds - age + 999) / 1000) + 5;
            }

            return Math.Max(5, Math.Min(30, seconds));
        }

        private static OutlookStartVerdict Settling(string code, string reason, int retryAfterSeconds)
        {
            return new OutlookStartVerdict(OutlookStartState.Settling, OutlookStuckKind.None, code, reason, retryAfterSeconds);
        }

        private static OutlookStartVerdict Stuck(OutlookStuckKind kind, string code, string reason)
        {
            return new OutlookStartVerdict(OutlookStartState.NotResponding, kind, code, reason, NotRespondingRetryAfterSeconds);
        }

        private static string Seconds(long? milliseconds)
        {
            return ((milliseconds ?? 0) / 1000).ToString(CultureInfo.InvariantCulture);
        }
    }
}
