using System;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using AddInServerContract = OutlookAI.Services.AddInServerContract;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Whether the OutlookAI add-in has finished starting inside a given Outlook (Q175 (a): "the add-in's
    /// start-up stamp present (build the stamp if it doesn't exist)").
    /// <para>
    /// Two stamps count. The add-in writes its own at the end of every start - this Outlook's process id
    /// and the time (<c>HKCU\Software\OutlookAI\AddIn</c>). An add-in from before that stamp is recognised
    /// by its tuning reconcile's time, which every start of it writes (<c>Tuning\LastReconcileUtc</c>):
    /// the live-tier guests run the add-in their frozen checkpoint holds, and a user may run a server newer
    /// than the add-in Outlook last loaded. Either one counts only when it is no older than the Outlook it
    /// is read for, so a stamp from an Outlook that has gone never stands for the next one.
    /// </para>
    /// </summary>
    public sealed class AddInStartupStamp
    {
        /// <summary>The stamp source: the add-in's own start-up stamp.</summary>
        public const string StartupStampSource = "add-in start-up stamp";

        /// <summary>The stamp source: the add-in's tuning reconcile, which an older add-in writes at each start.</summary>
        public const string TuningReconcileSource = "add-in tuning reconcile";

        /// <summary>
        /// How far before the process's start a stamp may be dated and still count: the two clocks are one
        /// clock, read at different resolutions.
        /// </summary>
        internal static readonly TimeSpan ClockTolerance = TimeSpan.FromSeconds(1);

        private AddInStartupStamp(bool expected, bool present, string? source, DateTime? stampUtc)
        {
            Expected = expected;
            Present = present;
            Source = source;
            StampUtc = stampUtc;
        }

        /// <summary>Whether the add-in is set to load at every Outlook start, so a stamp is coming.</summary>
        public bool Expected { get; }

        /// <summary>Whether the add-in has stamped this Outlook's start.</summary>
        public bool Present { get; }

        /// <summary>Which stamp said so, when one did.</summary>
        public string? Source { get; }

        /// <summary>When the add-in finished starting (UTC), when a stamp says.</summary>
        public DateTime? StampUtc { get; }

        /// <summary>
        /// The settle fact: true when stamped, false when a stamp is expected and missing, null when none
        /// is coming - or when the Outlook's start time is unknown, which leaves no stamp checkable.
        /// </summary>
        public bool? AsFact => Present ? true : Expected ? false : (bool?)null;

        /// <summary>
        /// The stamp for the Outlook with <paramref name="processId"/> that started at
        /// <paramref name="processStartUtc"/>, from the raw registry values. Pure, so T1 pins it.
        /// </summary>
        /// <param name="processId">The OUTLOOK.EXE.</param>
        /// <param name="processStartUtc">Its start; null when unreadable, which makes the stamp unknowable.</param>
        /// <param name="loadBehavior">The add-in registration's <c>LoadBehavior</c>.</param>
        /// <param name="stampProcessId">The add-in's <c>StartupProcessId</c>.</param>
        /// <param name="stampUtc">The add-in's <c>StartupCompleteUtc</c>.</param>
        /// <param name="tuningReconcileUtc">The add-in's <c>Tuning\LastReconcileUtc</c>.</param>
        public static AddInStartupStamp Evaluate(
            int? processId,
            DateTime? processStartUtc,
            object? loadBehavior,
            object? stampProcessId,
            object? stampUtc,
            object? tuningReconcileUtc)
        {
            bool expected = loadBehavior is int behavior && behavior == AddInServerContract.OutlookAddInLoadAtStartup;
            if (processId == null || processStartUtc == null)
            {
                return new AddInStartupStamp(expected: false, present: false, source: null, stampUtc: null);
            }

            DateTime earliest = processStartUtc.Value - ClockTolerance;
            DateTime? own = ParseUtc(stampUtc as string);
            if (stampProcessId is int stampedPid && stampedPid == processId.Value && own is DateTime ownTime && ownTime >= earliest)
            {
                return new AddInStartupStamp(expected, present: true, StartupStampSource, ownTime);
            }

            DateTime? tuning = ParseUtc(tuningReconcileUtc as string);
            if (tuning is DateTime tuningTime && tuningTime >= earliest)
            {
                return new AddInStartupStamp(expected, present: true, TuningReconcileSource, tuningTime);
            }

            return new AddInStartupStamp(expected, present: false, source: null, stampUtc: null);
        }

        /// <summary>Reads the stamp for one Outlook from the registry. Never throws.</summary>
        internal static AddInStartupStamp Read(int? processId, DateTime? processStartUtc)
        {
            return Evaluate(
                processId,
                processStartUtc,
                ReadValue(AddInServerContract.OutlookAddInRegistrationKeyPath, AddInServerContract.OutlookAddInLoadBehaviorValueName),
                ReadValue(AddInServerContract.AddInKeyPath, AddInServerContract.AddInStartupProcessIdValueName),
                ReadValue(AddInServerContract.AddInKeyPath, AddInServerContract.AddInStartupCompleteUtcValueName),
                ReadValue(AddInServerContract.TuningKeyPath, AddInServerContract.TuningLastReconcileUtcValueName));
        }

        private static object? ReadValue(string keyPath, string valueName)
        {
            try
            {
                using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(keyPath, writable: false))
                {
                    return key?.GetValue(valueName);
                }
            }
            catch (Exception ex) when (ex is System.Security.SecurityException || ex is UnauthorizedAccessException
                || ex is System.IO.IOException || ex is ObjectDisposedException)
            {
                return null;
            }
        }

        private static DateTime? ParseUtc(string? text)
        {
            return !string.IsNullOrEmpty(text)
                && DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out DateTime value)
                    ? DateTime.SpecifyKind(value, DateTimeKind.Utc)
                    : (DateTime?)null;
        }
    }

    /// <summary>One look at Outlook's start from outside - no COM - and the facts a decision takes from it.</summary>
    public sealed class OutlookStartObservation
    {
        internal OutlookStartObservation(
            DateTime observedUtc,
            OutlookLivenessSnapshot windows,
            DateTime? processStartUtc,
            AddInStartupStamp addIn,
            bool connectLockHeldElsewhere,
            bool connectLockAbandoned,
            OutlookConnectRecord? connect,
            OutlookProcessNote? windowSeen = null)
        {
            ObservedUtc = observedUtc;
            Windows = windows;
            ProcessStartUtc = processStartUtc;
            AddIn = addIn;
            ConnectLockHeldElsewhere = connectLockHeldElsewhere;
            ConnectLockAbandoned = connectLockAbandoned;
            Connect = connect;
            WindowSeen = windowSeen;
        }

        /// <summary>When the look was taken (UTC).</summary>
        public DateTime ObservedUtc { get; }

        /// <summary>What the window walk saw.</summary>
        public OutlookLivenessSnapshot Windows { get; }

        /// <summary>When the OUTLOOK.EXE started (UTC); null when none runs or it could not be read.</summary>
        public DateTime? ProcessStartUtc { get; }

        /// <summary>How long ago it started; null when unknown.</summary>
        public long? AgeMilliseconds => ProcessStartUtc is DateTime start
            ? Math.Max(0, (long)(ObservedUtc - start).TotalMilliseconds)
            : (long?)null;

        /// <summary>Whether the add-in has finished starting inside it.</summary>
        public AddInStartupStamp AddIn { get; }

        /// <summary>Whether another connection holds the start lock.</summary>
        public bool ConnectLockHeldElsewhere { get; }

        /// <summary>Whether the look found the lock abandoned: the connection holding it had died.</summary>
        public bool ConnectLockAbandoned { get; }

        /// <summary>The start record's connection, whatever Outlook it is about; null when there is none.</summary>
        public OutlookConnectRecord? Connect { get; }

        /// <summary>The start record's note of a window seen on screen, whatever Outlook it is about; null when there is none.</summary>
        public OutlookProcessNote? WindowSeen { get; }

        /// <summary>
        /// Whether this Outlook got through its start once already: its add-in stamped it, a window of it was
        /// seen, or an OutlookAI connection got past the settle wait in it. With no windows now it is closing.
        /// </summary>
        public bool FinishedStartingBefore
        {
            get
            {
                int? pid = Windows.ProcessId;
                if (pid == null)
                {
                    return false;
                }

                if (AddIn.Present || (WindowSeen != null && WindowSeen.IsAbout(pid, ProcessStartUtc)))
                {
                    return true;
                }

                OutlookConnectRecord? record = Connect;
                return record != null
                    && record.OutlookProcessId == pid
                    && new OutlookProcessNote(pid.Value, record.OutlookStartUtc, record.StepSinceUtc, null).IsAbout(pid, ProcessStartUtc)
                    && OutlookStartSteps.IsPastSettle(record.Step);
            }
        }

        /// <summary>
        /// How long the connection holding the lock has been in its step, when the record says; null when
        /// nobody holds it.
        /// </summary>
        public long? ConnectStepMilliseconds => ConnectLockHeldElsewhere && Connect is { InProgress: true } record
            ? Math.Max(0, (long)(ObservedUtc - record.StepSinceUtc).TotalMilliseconds)
            : (long?)null;

        /// <summary>The decision's facts, with the main thread measured or not.</summary>
        public OutlookStartFacts ToFacts(bool? mainThreadIdle = null)
        {
            return new OutlookStartFacts(
                Windows.State, AgeMilliseconds, AddIn.AsFact, mainThreadIdle, ConnectLockHeldElsewhere, ConnectStepMilliseconds,
                FinishedStartingBefore);
        }

        /// <summary>A few words on what was seen, for health and the supervisor's last-failure line.</summary>
        public string Describe()
        {
            return Windows.Detail
                + (AgeMilliseconds is long age ? "; started " + (age / 1000).ToString(CultureInfo.InvariantCulture) + " s ago" : string.Empty)
                + (Windows.State == OutlookLivenessState.NotRunning ? string.Empty : Windows.HasVisibleWindow ? "; has a window" : "; no window on screen");
        }
    }

    /// <summary>
    /// Looks at Outlook's start without COM: the window walk, the process's start time, the add-in's
    /// stamp, the start lock and the start record. Each costs microseconds; nothing here can block on
    /// Outlook, so the gate in front of every request can afford it.
    /// </summary>
    public static class OutlookStartObserver
    {
        /// <summary>Takes one look.</summary>
        /// <param name="connectLockHeldElsewhere">
        /// Null to ask whether another connection holds the start lock; the answer, when the caller already
        /// knows it - false for the connection that holds it, true for one that has just failed to take it.
        /// </param>
        public static OutlookStartObservation Observe(bool? connectLockHeldElsewhere = null)
        {
            DateTime now = DateTime.UtcNow;
            OutlookLivenessSnapshot windows = OutlookLiveness.Snapshot();
            DateTime? started = windows.ProcessId is int pid ? TryGetProcessStartUtc(pid) : null;
            AddInStartupStamp addIn = windows.ProcessId == null
                ? AddInStartupStamp.Evaluate(null, null, null, null, null, null)
                : AddInStartupStamp.Read(windows.ProcessId, started);

            bool abandoned = false;
            bool held = connectLockHeldElsewhere ?? OutlookStartLock.IsHeldElsewhere(out abandoned);
            return new OutlookStartObservation(
                now,
                windows,
                started,
                addIn,
                held,
                abandoned,
                OutlookStartRecord.ReadConnect(),
                windows.ProcessId == null ? null : OutlookStartRecord.ReadNote(OutlookStartRecord.WindowSeenValueName));
        }

        /// <summary>When the process started (UTC), read from Windows in UTC; null when it cannot be read.</summary>
        public static DateTime? TryGetProcessStartUtc(int processId)
        {
            IntPtr process = NativeMethods.OpenProcess(NativeMethods.ProcessQueryLimitedInformation, false, (uint)processId);
            if (process == IntPtr.Zero)
            {
                return null;
            }

            try
            {
                return NativeMethods.GetProcessTimes(process, out long created, out _, out _, out _) && created > 0
                    ? DateTime.FromFileTimeUtc(created)
                    : (DateTime?)null;
            }
            finally
            {
                _ = NativeMethods.CloseHandle(process);
            }
        }

        /// <summary>The CPU time a thread has used, in milliseconds; null when it cannot be read.</summary>
        internal static long? TryGetThreadCpuMilliseconds(int threadId)
        {
            IntPtr thread = NativeMethods.OpenThread(NativeMethods.ThreadQueryLimitedInformation, false, (uint)threadId);
            if (thread == IntPtr.Zero)
            {
                return null;
            }

            try
            {
                return NativeMethods.GetThreadTimes(thread, out _, out _, out long kernel, out long user)
                    ? (kernel + user) / 10_000
                    : (long?)null;
            }
            finally
            {
                _ = NativeMethods.CloseHandle(thread);
            }
        }

        /// <summary>
        /// Whether the thread that owns <paramref name="window"/> answers a message within
        /// <paramref name="timeoutMilliseconds"/>, and how fast. WM_NULL does nothing; it only has to be
        /// taken off the queue. Never waits on a hung thread: Windows answers for it at once.
        /// </summary>
        internal static bool TryPingMainThread(IntPtr window, int timeoutMilliseconds, out long replyMilliseconds)
        {
            replyMilliseconds = 0;
            if (window == IntPtr.Zero)
            {
                return false;
            }

            Stopwatch clock = Stopwatch.StartNew();
            IntPtr answered = NativeMethods.SendMessageTimeoutW(
                window,
                NativeMethods.WmNull,
                IntPtr.Zero,
                IntPtr.Zero,
                NativeMethods.SmtoAbortIfHung | NativeMethods.SmtoErrorOnExit,
                (uint)timeoutMilliseconds,
                out _);
            replyMilliseconds = clock.ElapsedMilliseconds;
            return answered != IntPtr.Zero;
        }

        private static class NativeMethods
        {
            internal const uint ProcessQueryLimitedInformation = 0x1000;
            internal const uint ThreadQueryLimitedInformation = 0x0800;
            internal const uint WmNull = 0x0000;
            internal const uint SmtoAbortIfHung = 0x0002;
            internal const uint SmtoErrorOnExit = 0x0020;

            [DllImport("kernel32.dll", SetLastError = true)]
            internal static extern IntPtr OpenProcess(uint desiredAccess, [MarshalAs(UnmanagedType.Bool)] bool inheritHandle, uint processId);

            [DllImport("kernel32.dll", SetLastError = true)]
            internal static extern IntPtr OpenThread(uint desiredAccess, [MarshalAs(UnmanagedType.Bool)] bool inheritHandle, uint threadId);

            [DllImport("kernel32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool GetProcessTimes(IntPtr process, out long creation, out long exit, out long kernel, out long user);

            [DllImport("kernel32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool GetThreadTimes(IntPtr thread, out long creation, out long exit, out long kernel, out long user);

            [DllImport("kernel32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool CloseHandle(IntPtr handle);

            [DllImport("user32.dll", SetLastError = true)]
            internal static extern IntPtr SendMessageTimeoutW(
                IntPtr window, uint message, IntPtr wParam, IntPtr lParam, uint flags, uint timeout, out UIntPtr result);
        }
    }
}
