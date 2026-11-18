using System;
using System.Diagnostics;
using System.Globalization;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// How one connection to Outlook went: whether it started Outlook, how long it waited for another
    /// connection and for Outlook to settle, and why the wait ended. What the live tier reads to measure
    /// what the settle wait costs (Q175, Q181 (d)), and what the start log records.
    /// </summary>
    public sealed class OutlookConnectTrace
    {
        internal OutlookConnectTrace(
            bool startedOutlook,
            int? outlookProcessId,
            long lockWaitMilliseconds,
            long settleWaitMilliseconds,
            string settleOutcome,
            long? ageAtSettleMilliseconds,
            string? addInStampSource,
            long totalMilliseconds)
        {
            StartedOutlook = startedOutlook;
            OutlookProcessId = outlookProcessId;
            LockWaitMilliseconds = lockWaitMilliseconds;
            SettleWaitMilliseconds = settleWaitMilliseconds;
            SettleOutcome = settleOutcome;
            AgeAtSettleMilliseconds = ageAtSettleMilliseconds;
            AddInStampSource = addInStampSource;
            TotalMilliseconds = totalMilliseconds;
        }

        /// <summary>Whether this connection started Outlook rather than attaching to it.</summary>
        public bool StartedOutlook { get; }

        /// <summary>The OUTLOOK.EXE it connected to.</summary>
        public int? OutlookProcessId { get; }

        /// <summary>How long it waited for another connection to finish (Q175 (e)).</summary>
        public long LockWaitMilliseconds { get; }

        /// <summary>How long it waited for Outlook to settle before opening a mailbox (Q175 (a)).</summary>
        public long SettleWaitMilliseconds { get; }

        /// <summary>Why the settle wait ended: <c>settled</c>, <c>settled-by-age</c>, or a code naming what was still unmet at the limit.</summary>
        public string SettleOutcome { get; }

        /// <summary>How old Outlook was when the wait ended.</summary>
        public long? AgeAtSettleMilliseconds { get; }

        /// <summary>Which add-in stamp said the add-in had started, when one did.</summary>
        public string? AddInStampSource { get; }

        /// <summary>The whole connection, lock wait to connected.</summary>
        public long TotalMilliseconds { get; }

        /// <summary>One line for a test's output and for a report.</summary>
        public string Describe()
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                "{0} Outlook pid {1}; waited {2} ms for another connection, {3} ms for Outlook to settle ({4}, Outlook {5} s old, add-in stamp: {6}); {7} ms in all",
                StartedOutlook ? "started" : "attached to",
                OutlookProcessId?.ToString(CultureInfo.InvariantCulture) ?? "?",
                LockWaitMilliseconds,
                SettleWaitMilliseconds,
                SettleOutcome,
                AgeAtSettleMilliseconds is long age ? (age / 1000.0).ToString("0.0", CultureInfo.InvariantCulture) : "?",
                AddInStampSource ?? "none",
                TotalMilliseconds);
        }
    }

    /// <summary>
    /// One connection's way through Outlook's start, on the connection's own STA thread: the start lock
    /// (one connection at a time, Q175 (e)), the start record (which step, since when, Q176 (c)), the
    /// start log (Q178 (e)), and the wait for Outlook to settle before any mailbox is opened (Q175 (a)).
    /// <para>
    /// <b>Every wait keeps the thread's messages flowing</b> (<see cref="PumpedStaRunner.PumpedWait"/>) and
    /// keeps looking at Outlook: an Outlook that stops responding, or a start that runs past
    /// <see cref="OutlookStartPolicy.StartingLimitMilliseconds"/>, ends the wait with an
    /// <see cref="OutlookNotRespondingException"/> carrying the report, inside the connect budget - so the
    /// caller is told what is stuck rather than timed out.
    /// </para>
    /// <para>
    /// Used only by <see cref="OutlookComSession"/>'s connect, inside one work item, so the lock is taken
    /// and released on one thread as a mutex requires.
    /// </para>
    /// </summary>
    internal sealed class OutlookStartGuard : IDisposable
    {
        /// <summary>How often a connection waiting for the start lock tries it again.</summary>
        private const int LockPollMilliseconds = 250;

        private readonly OutlookStartLock _lock;
        private readonly Stopwatch _clock;
        private readonly long _lockWaitMilliseconds;
        private OutlookConnectRecord _record;
        private int? _outlookProcessId;
        private DateTime? _outlookStartUtc;
        private bool _startedOutlook;
        private long _settleWaitMilliseconds;
        private string _settleOutcome = "not-waited";
        private long? _ageAtSettle;
        private string? _stampSource;
        private bool _finished;

        private OutlookStartGuard(OutlookStartLock heldLock, Stopwatch clock, long lockWaitMilliseconds)
        {
            _lock = heldLock;
            _clock = clock;
            _lockWaitMilliseconds = lockWaitMilliseconds;
            DateTime now = DateTime.UtcNow;
            _record = new OutlookConnectRecord(
                null, null, false, ComHostIdentity.ServerProcessId, ComHostIdentity.HostProcessId,
                OutlookStartSteps.CreateApplication, now, now);
        }

        /// <summary>How the connection went, once it is connected; null before.</summary>
        internal OutlookConnectTrace? Trace { get; private set; }

        /// <summary>
        /// Takes the start lock, waiting for another connection to finish if one holds it - and looking at
        /// Outlook while it waits, so a hung Outlook or a holder stuck past the limit ends the wait with a
        /// report instead of a timeout. Runs on the connection's STA thread.
        /// </summary>
        internal static OutlookStartGuard Enter(bool outlookWasRunning)
        {
            Stopwatch clock = Stopwatch.StartNew();
            OutlookStartLog.Write("connect.begin", OutlookStartLog.Field("outlookRunning", outlookWasRunning));

            OutlookStartLock? held = OutlookStartLock.TryAcquire();
            if (held == null)
            {
                OutlookConnectRecord? holder = OutlookStartRecord.ReadConnect();
                OutlookStartLog.Write(
                    "lock.wait",
                    OutlookStartLog.Field("holderServer", holder?.ServerProcessId),
                    OutlookStartLog.Field("holderHost", holder?.HostProcessId),
                    OutlookStartLog.Field("holderStep", holder?.Step));
            }

            while (held == null)
            {
                OutlookStartObservation observation = OutlookStartObserver.Observe(connectLockHeldElsewhere: true);
                OutlookStartVerdict verdict = OutlookStartPolicy.Decide(observation.ToFacts());
                if (verdict.State == OutlookStartState.NotResponding)
                {
                    throw NotResponding("lock.wait", verdict, observation, clock);
                }

                if (clock.ElapsedMilliseconds >= OutlookStartPolicy.ConnectLockWaitMilliseconds)
                {
                    OutlookStartVerdict waitedOut = OutlookStartPolicy.Decide(new OutlookStartFacts(
                        observation.Windows.State,
                        observation.AgeMilliseconds,
                        observation.AddIn.AsFact,
                        null,
                        connectLockHeldElsewhere: true,
                        connectStepMilliseconds: Math.Max(clock.ElapsedMilliseconds, OutlookStartPolicy.StartingLimitMilliseconds),
                        finishedStartingBefore: observation.FinishedStartingBefore));
                    throw NotResponding("lock.timeout", waitedOut, observation, clock);
                }

                PumpedStaRunner.PumpedWait(LockPollMilliseconds);
                held = OutlookStartLock.TryAcquire();
            }

            OutlookStartLog.Write(
                "lock.acquired",
                OutlookStartLog.Field("waitedMs", clock.ElapsedMilliseconds),
                OutlookStartLog.Field("previousHolderDied", held.PreviousHolderDied ? true : (bool?)null),
                OutlookStartLog.Field("lockAvailable", held.Held ? (bool?)null : false));
            return new OutlookStartGuard(held, clock, clock.ElapsedMilliseconds);
        }

        /// <summary>Records that the connection has moved to <paramref name="step"/>.</summary>
        internal void BeginStep(string step)
        {
            _record = _record.WithStep(step, DateTime.UtcNow, _outlookProcessId, _outlookStartUtc, _startedOutlook);
            Save();
            OutlookStartLog.Write(
                "connect.step",
                OutlookStartLog.Field("step", step),
                OutlookStartLog.Field("outlook", _outlookProcessId),
                OutlookStartLog.Field("elapsedMs", _clock.ElapsedMilliseconds));
        }

        /// <summary>Records whether this connection is starting Outlook or attaching to it.</summary>
        internal void SetStartingOutlook(bool starting)
        {
            _startedOutlook = starting;
        }

        /// <summary>
        /// Waits until Outlook has settled (Q175 (a)): at least <see cref="OutlookStartPolicy.MinimumAgeMilliseconds"/>
        /// old, the add-in's start-up stamp present when one is coming, and its main thread idle. An Outlook
        /// that started long enough ago counts as settled at once, so a connection to it pays one look.
        /// </summary>
        /// <param name="startedHere">
        /// True after this connection started Outlook, which notes it as OutlookAI's in the start record; false
        /// before an attach - where the wait comes BEFORE asking for Outlook.Application, because that request
        /// itself blocks with no limit on a hung Outlook.
        /// </param>
        /// <exception cref="OutlookNotRespondingException">Outlook stopped responding, or did not finish starting in time.</exception>
        /// <exception cref="InvalidOperationException">Outlook exited while it was waited for.</exception>
        internal void AwaitSettled(bool startedHere)
        {
            BeginStep(OutlookStartSteps.Settle);
            Stopwatch waited = Stopwatch.StartNew();
            string? lastCode = null;
            bool notedStart = false;
            while (true)
            {
                OutlookStartObservation observation = OutlookStartObserver.Observe(connectLockHeldElsewhere: false);
                AdoptOutlook(observation);
                if (startedHere && !notedStart && observation.Windows.ProcessId is int pid)
                {
                    OutlookStartRecord.WriteNote(
                        OutlookStartRecord.StartedValueName,
                        new OutlookProcessNote(pid, observation.ProcessStartUtc, DateTime.UtcNow, ComHostIdentity.ServerProcessId));
                    notedStart = true;
                }

                OutlookStartVerdict verdict = OutlookStartPolicy.Decide(observation.ToFacts());
                long cpuMilliseconds = -1;
                long replyMilliseconds = -1;
                if (verdict.State == OutlookStartState.Ready && !verdict.SettledByAge)
                {
                    // Every condition that costs nothing holds; now the one that takes a sample.
                    bool? idle = MeasureMainThread(observation, out cpuMilliseconds, out replyMilliseconds);
                    verdict = OutlookStartPolicy.Decide(observation.ToFacts(idle));
                }

                if (!string.Equals(lastCode, verdict.ReasonCode, StringComparison.Ordinal))
                {
                    lastCode = verdict.ReasonCode;
                    OutlookStartLog.Write(
                        "settle.look",
                        OutlookStartLog.Field("outlook", observation.Windows.ProcessId),
                        OutlookStartLog.Field("state", verdict.ReasonCode),
                        OutlookStartLog.Field("ageMs", observation.AgeMilliseconds),
                        OutlookStartLog.Field("liveness", observation.Windows.State.ToString()),
                        OutlookStartLog.Field("addIn", observation.AddIn.Present ? "stamped" : observation.AddIn.Expected ? "expected" : "not-expected"),
                        OutlookStartLog.Field("mainThreadCpuMs", cpuMilliseconds >= 0 ? cpuMilliseconds : (long?)null),
                        OutlookStartLog.Field("mainThreadReplyMs", replyMilliseconds >= 0 ? replyMilliseconds : (long?)null),
                        OutlookStartLog.Field("waitedMs", waited.ElapsedMilliseconds));
                }

                switch (verdict.State)
                {
                    case OutlookStartState.Ready:
                        Settled(verdict.ReasonCode, waited.ElapsedMilliseconds, observation);
                        return;

                    case OutlookStartState.NotRunning:
                        OutlookStartLog.Write("settle.end", OutlookStartLog.Field("outcome", "outlook-exited"),
                            OutlookStartLog.Field("waitedMs", waited.ElapsedMilliseconds));
                        throw new InvalidOperationException(
                            "Outlook exited while OutlookAI was waiting for it to finish starting. Nothing was opened in it; "
                            + "retry, and OutlookAI starts or attaches to Outlook again.");

                    case OutlookStartState.NotResponding:
                        throw NotResponding("settle.end", verdict, observation, waited);

                    default:
                        break;
                }

                if (waited.ElapsedMilliseconds >= OutlookStartPolicy.StartingLimitMilliseconds)
                {
                    // The policy's limit runs on Outlook's age; this one on the wait itself, for the Outlook
                    // whose start time could not be read. A responding Outlook is used: the conditions narrow
                    // a start-up race, and no start-up is left to race.
                    if (observation.Windows.State == OutlookLivenessState.Responsive)
                    {
                        Settled("waived-" + verdict.ReasonCode, waited.ElapsedMilliseconds, observation);
                        return;
                    }

                    OutlookStartVerdict expired = OutlookStartPolicy.Decide(new OutlookStartFacts(
                        observation.Windows.State,
                        OutlookStartPolicy.StartingLimitMilliseconds,
                        observation.AddIn.AsFact,
                        null,
                        connectLockHeldElsewhere: false,
                        connectStepMilliseconds: null,
                        finishedStartingBefore: false));
                    throw NotResponding("settle.end", expired, observation, waited);
                }

                PumpedStaRunner.PumpedWait(OutlookStartPolicy.SettlePollMilliseconds);
            }
        }

        /// <summary>Records the connection as connected. The lock is released by <see cref="Dispose"/>.</summary>
        internal void Connected()
        {
            _record = _record.WithStep(OutlookStartSteps.Connected, DateTime.UtcNow, _outlookProcessId, _outlookStartUtc, _startedOutlook);
            Save();
            Trace = new OutlookConnectTrace(
                _startedOutlook,
                _outlookProcessId,
                _lockWaitMilliseconds,
                _settleWaitMilliseconds,
                _settleOutcome,
                _ageAtSettle,
                _stampSource,
                _clock.ElapsedMilliseconds);
            OutlookStartLog.Write(
                "connect.end",
                OutlookStartLog.Field("outcome", "connected"),
                OutlookStartLog.Field("outlook", _outlookProcessId),
                OutlookStartLog.Field("started", _startedOutlook),
                OutlookStartLog.Field("lockWaitMs", _lockWaitMilliseconds),
                OutlookStartLog.Field("settleWaitMs", _settleWaitMilliseconds),
                OutlookStartLog.Field("settle", _settleOutcome),
                OutlookStartLog.Field("totalMs", _clock.ElapsedMilliseconds));
            _finished = true;
        }

        /// <summary>
        /// Records the connection as failed with <paramref name="failure"/> - its type and HRESULT, never its
        /// message, which can name folders and stores. Null when it ended without saying why.
        /// </summary>
        internal void Failed(Exception? failure)
        {
            if (_finished)
            {
                return;
            }

            _finished = true;
            string failedIn = _record.Step;
            _record = _record.WithStep(OutlookStartSteps.Failed, DateTime.UtcNow, _outlookProcessId, _outlookStartUtc, _startedOutlook);
            Save();
            OutlookStartLog.Write(
                "connect.end",
                OutlookStartLog.Field("outcome", "failed"),
                OutlookStartLog.Field("failedIn", failedIn),
                OutlookStartLog.Field("error", failure?.GetType().Name ?? "unfinished"),
                OutlookStartLog.Field("hresult", failure == null ? null : "0x" + failure.HResult.ToString("X8", CultureInfo.InvariantCulture)),
                OutlookStartLog.Field("outlook", _outlookProcessId),
                OutlookStartLog.Field("totalMs", _clock.ElapsedMilliseconds));
        }

        /// <summary>Releases the start lock; a connection that never said how it ended is recorded as failed.</summary>
        public void Dispose()
        {
            if (!_finished)
            {
                Failed(null);
            }

            _lock.Dispose();
        }

        private void Settled(string outcome, long waitedMilliseconds, OutlookStartObservation observation)
        {
            _settleOutcome = outcome;
            _settleWaitMilliseconds = waitedMilliseconds;
            _ageAtSettle = observation.AgeMilliseconds;
            _stampSource = observation.AddIn.Source;
            OutlookStartLog.Write(
                "settle.end",
                OutlookStartLog.Field("outcome", outcome),
                OutlookStartLog.Field("outlook", observation.Windows.ProcessId),
                OutlookStartLog.Field("waitedMs", waitedMilliseconds),
                OutlookStartLog.Field("ageMs", observation.AgeMilliseconds),
                OutlookStartLog.Field("addInStamp", observation.AddIn.Source));
        }

        private void AdoptOutlook(OutlookStartObservation observation)
        {
            if (observation.Windows.ProcessId is int pid && (_outlookProcessId != pid || _outlookStartUtc != observation.ProcessStartUtc))
            {
                _outlookProcessId = pid;
                _outlookStartUtc = observation.ProcessStartUtc;
                _record = _record.WithStep(_record.Step, _record.StepSinceUtc, pid, observation.ProcessStartUtc, _startedOutlook);
                Save();
            }
        }

        private void Save()
        {
            OutlookStartRecord.WriteConnect(_record);
        }

        /// <summary>
        /// Whether Outlook's main thread is idle: it uses almost no CPU over a short sample and answers a
        /// message promptly. Null when its CPU time cannot be read and it answered - the reply is then the
        /// only evidence, and a thread that answers is not stuck.
        /// </summary>
        private static bool? MeasureMainThread(OutlookStartObservation observation, out long cpuMilliseconds, out long replyMilliseconds)
        {
            cpuMilliseconds = -1;
            replyMilliseconds = -1;
            if (observation.Windows.MainThreadId is not int thread)
            {
                return false;
            }

            long? before = OutlookStartObserver.TryGetThreadCpuMilliseconds(thread);
            PumpedStaRunner.PumpedWait(OutlookStartPolicy.IdleSampleMilliseconds);
            long? after = OutlookStartObserver.TryGetThreadCpuMilliseconds(thread);
            bool replied = OutlookStartObserver.TryPingMainThread(
                observation.Windows.MainWindow, OutlookStartPolicy.MainThreadReplyMilliseconds, out replyMilliseconds);
            if (before == null || after == null)
            {
                return replied ? (bool?)null : false;
            }

            cpuMilliseconds = after.Value - before.Value;
            return replied && cpuMilliseconds <= OutlookStartPolicy.IdleCpuMilliseconds;
        }

        /// <summary>The report for what was seen, logged and noted, as the exception that ends the wait.</summary>
        private static OutlookNotRespondingException NotResponding(
            string logEvent, OutlookStartVerdict verdict, OutlookStartObservation observation, Stopwatch waited)
        {
            OutlookStartRecord.NoteWhatWasSeen(observation, verdict);
            OutlookStartLog.Write(
                logEvent,
                OutlookStartLog.Field("outcome", "not-responding"),
                OutlookStartLog.Field("stuck", verdict.ReasonCode),
                OutlookStartLog.Field("outlook", observation.Windows.ProcessId),
                OutlookStartLog.Field("ageMs", observation.AgeMilliseconds),
                OutlookStartLog.Field("hasWindow", observation.Windows.HasVisibleWindow),
                OutlookStartLog.Field("waitedMs", waited.ElapsedMilliseconds));
            return new OutlookNotRespondingException(
                OutlookHangReport.DescribeNow(verdict, observation), verdict.RetryAfterSeconds, observation.Windows.ProcessId);
        }
    }
}
