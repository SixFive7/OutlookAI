using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// A process, told apart from a later one that reuses its id by its start time (UTC ticks).
    /// </summary>
    public readonly struct ProcessStamp : IEquatable<ProcessStamp>
    {
        /// <summary>Creates the stamp.</summary>
        public ProcessStamp(int pid, long startedUtcTicks)
        {
            Pid = pid;
            StartedUtcTicks = startedUtcTicks;
        }

        /// <summary>The process id.</summary>
        public int Pid { get; }

        /// <summary>When the process started, as UTC ticks.</summary>
        public long StartedUtcTicks { get; }

        /// <summary>True for the default value, which names no process.</summary>
        public bool IsEmpty => Pid == 0 && StartedUtcTicks == 0;

        /// <summary>The stamp of <paramref name="process"/>, or null when its start time cannot be read.</summary>
        public static ProcessStamp? Of(Process process)
        {
            if (process == null)
            {
                return null;
            }

            try
            {
                return new ProcessStamp(process.Id, process.StartTime.ToUniversalTime().Ticks);
            }
            catch (Exception ex) when (ex is InvalidOperationException || ex is System.ComponentModel.Win32Exception || ex is NotSupportedException)
            {
                return null;
            }
        }

        private static readonly Lazy<ProcessStamp> ThisProcess = new Lazy<ProcessStamp>(() =>
        {
            using (Process self = Process.GetCurrentProcess())
            {
                return Of(self) ?? new ProcessStamp(self.Id, 0);
            }
        });

        /// <summary>The stamp of this process.</summary>
        public static ProcessStamp Current()
        {
            return ThisProcess.Value;
        }

        /// <summary>
        /// Whether the process this stamp names is still running: a process with its id that started at its time.
        /// A process this one may not read counts as running - nothing is ever closed on a guess.
        /// </summary>
        public static bool IsRunning(ProcessStamp stamp)
        {
            if (stamp.Pid <= 0)
            {
                return false;
            }

            Process? process = null;
            try
            {
                process = Process.GetProcessById(stamp.Pid);
                if (process.HasExited)
                {
                    return false;
                }

                ProcessStamp? now = Of(process);
                return now == null || stamp.StartedUtcTicks == 0 || now.Value.StartedUtcTicks == stamp.StartedUtcTicks;
            }
            catch (ArgumentException)
            {
                // No process with that id.
                return false;
            }
            catch (Exception ex) when (ex is InvalidOperationException || ex is System.ComponentModel.Win32Exception)
            {
                return true;
            }
            finally
            {
                process?.Dispose();
            }
        }

        /// <inheritdoc />
        public bool Equals(ProcessStamp other)
        {
            return Pid == other.Pid && StartedUtcTicks == other.StartedUtcTicks;
        }

        /// <inheritdoc />
        public override bool Equals(object? obj)
        {
            return obj is ProcessStamp other && Equals(other);
        }

        /// <inheritdoc />
        public override int GetHashCode()
        {
            unchecked
            {
                return (Pid * 397) ^ StartedUtcTicks.GetHashCode();
            }
        }

        /// <summary>"pid@ticks" - the form the registry file writes.</summary>
        public override string ToString()
        {
            return Pid.ToString(CultureInfo.InvariantCulture) + "@" + StartedUtcTicks.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>Reads the "pid@ticks" form back; null when it is not one.</summary>
        public static ProcessStamp? TryParse(string? text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return null;
            }

            int at = text!.IndexOf('@');
            if (at <= 0
                || !int.TryParse(text.Substring(0, at), NumberStyles.None, CultureInfo.InvariantCulture, out int pid)
                || !long.TryParse(text.Substring(at + 1), NumberStyles.None, CultureInfo.InvariantCulture, out long ticks)
                || pid <= 0)
            {
                return null;
            }

            return new ProcessStamp(pid, ticks);
        }

        /// <summary>Equality.</summary>
        public static bool operator ==(ProcessStamp left, ProcessStamp right)
        {
            return left.Equals(right);
        }

        /// <summary>Inequality.</summary>
        public static bool operator !=(ProcessStamp left, ProcessStamp right)
        {
            return !left.Equals(right);
        }
    }

    /// <summary>
    /// One Outlook process as the registry knows it: who started it, whether OutlookAI has promoted a compose
    /// window in it, and the hidden lifetime pin that keeps it up while it has no window of its own.
    /// </summary>
    public sealed class RegisteredOutlook
    {
        /// <summary>Creates the record.</summary>
        public RegisteredOutlook(ProcessStamp outlook)
        {
            Outlook = outlook;
        }

        /// <summary>The OUTLOOK.EXE this record is about.</summary>
        public ProcessStamp Outlook { get; }

        /// <summary>The key of the OutlookAI session that started this Outlook, or null when something else did.</summary>
        public string? StartedBy { get; set; }

        /// <summary>When OutlookAI first promoted a compose window in this process (UTC ticks), 0 while it has not.</summary>
        public long PromotedUtcTicks { get; set; }

        /// <summary>True when a draft's editor was refused without a promotion since the last one: the next draft promotes.</summary>
        public bool DirectRefused { get; set; }

        /// <summary>The window handle of this Outlook's lifetime pin, 0 when it has none.</summary>
        public long PinWindow { get; set; }

        /// <summary>The key of the session that made the pin.</summary>
        public string? PinMadeBy { get; set; }

        /// <summary>When the pin was made (UTC ticks).</summary>
        public long PinSinceUtcTicks { get; set; }

        /// <summary>A copy, so a rule can be tested without sharing state.</summary>
        public RegisteredOutlook Copy()
        {
            return new RegisteredOutlook(Outlook)
            {
                StartedBy = StartedBy,
                PromotedUtcTicks = PromotedUtcTicks,
                DirectRefused = DirectRefused,
                PinWindow = PinWindow,
                PinMadeBy = PinMadeBy,
                PinSinceUtcTicks = PinSinceUtcTicks,
            };
        }
    }

    /// <summary>One OutlookAI COM session attached to an Outlook.</summary>
    public sealed class RegisteredSession
    {
        /// <summary>Creates the record.</summary>
        public RegisteredSession(string key, ProcessStamp process, ProcessStamp outlook, long sinceUtcTicks, bool startedOutlook)
        {
            Key = key ?? throw new ArgumentNullException(nameof(key));
            Process = process;
            Outlook = outlook;
            SinceUtcTicks = sinceUtcTicks;
            StartedOutlook = startedOutlook;
        }

        /// <summary>"pid.processStartTicks.sequence" - unique per session, across processes and PID reuse.</summary>
        public string Key { get; }

        /// <summary>The process the session lives in: an OutlookAI COM host, or a test process.</summary>
        public ProcessStamp Process { get; }

        /// <summary>The Outlook it is attached to.</summary>
        public ProcessStamp Outlook { get; }

        /// <summary>When it attached (UTC ticks).</summary>
        public long SinceUtcTicks { get; }

        /// <summary>True when this session's connect started Outlook.</summary>
        public bool StartedOutlook { get; }
    }

    /// <summary>The whole registry: the Outlooks it knows and the sessions attached to them.</summary>
    public sealed class SessionRegistryState
    {
        /// <summary>The Outlook processes, at most one per process stamp.</summary>
        public List<RegisteredOutlook> Outlooks { get; } = new List<RegisteredOutlook>();

        /// <summary>The live sessions, at most one per key.</summary>
        public List<RegisteredSession> Sessions { get; } = new List<RegisteredSession>();

        /// <summary>The record of <paramref name="outlook"/>, or null.</summary>
        public RegisteredOutlook? Find(ProcessStamp outlook)
        {
            return Outlooks.FirstOrDefault(o => o.Outlook == outlook);
        }

        /// <summary>The record of <paramref name="outlook"/>, made when there is none.</summary>
        public RegisteredOutlook FindOrAdd(ProcessStamp outlook)
        {
            RegisteredOutlook? found = Find(outlook);
            if (found != null)
            {
                return found;
            }

            RegisteredOutlook added = new RegisteredOutlook(outlook);
            Outlooks.Add(added);
            return added;
        }

        /// <summary>The sessions attached to <paramref name="outlook"/>.</summary>
        public IReadOnlyList<RegisteredSession> SessionsOf(ProcessStamp outlook)
        {
            return Sessions.Where(s => s.Outlook == outlook).ToList();
        }
    }

    /// <summary>The route by which a draft's Word editor is had (Q149 (d), decided as Q208 B on 2026-11-18).</summary>
    public enum DraftEditorRoute
    {
        /// <summary>
        /// Read straight from the inspector, never activated: <c>ModifiedFormPages.Add("Message")</c> first, the
        /// edit committed by <c>Close(olSave)</c> (cells B3 and W7, 100 of 100 on Office LTSC 2024).
        /// </summary>
        Direct = 0,

        /// <summary>
        /// The promotion: the inspector parked, activated, hidden again; the edit committed by
        /// <c>CurrentItem.Save()</c> and <c>Close(olDiscard)</c>. Once per Outlook process (cells W6 and W7).
        /// </summary>
        Promote = 1,
    }

    /// <summary>What a session leaving must do about its Outlook's lifetime pin.</summary>
    public sealed class LeaveVerdict
    {
        internal LeaveVerdict(long closePinWindow, string reason, int othersAttached)
        {
            ClosePinWindow = closePinWindow;
            Reason = reason;
            OthersAttached = othersAttached;
        }

        /// <summary>The pin's window to close, 0 when the pin stays (or there is none).</summary>
        public long ClosePinWindow { get; }

        /// <summary>Why, in a few words - for the diagnostics.</summary>
        public string Reason { get; }

        /// <summary>How many other live sessions are still attached to the same Outlook.</summary>
        public int OthersAttached { get; }
    }

    /// <summary>
    /// THE RULES of the per-user session registry (Q118 (a), approved 2026-10-11; Q208 B and Q175 (b), decided
    /// 2026-11-18). Pure: every function takes the state and returns a decision, so T1 pins each rule without a
    /// file, a lock or an Outlook.
    /// <list type="bullet">
    /// <item>Every OutlookAI COM session JOINS when it attaches and LEAVES when it ends. A session whose process
    /// is gone is pruned by the next one that takes the lock.</item>
    /// <item>An Outlook with no window has ONE lifetime pin, shared by every session, made only when a draft
    /// first needs it (Q175 (b)) and recorded by its window handle.</item>
    /// <item>The LAST session out closes the pin, whoever made it and whoever started Outlook; while another
    /// session lives, the pin stays - a dead session's pin is adopted, never closed under the living.</item>
    /// <item>A session that hears Outlook's Quit closes the pin whatever else is attached: the user's exit wins
    /// for every session.</item>
    /// <item>One promotion per Outlook process: after it, and whenever Outlook shows a window, a draft's editor
    /// is read directly (Q208 B). A direct read that is refused sends the next draft back to the promotion.</item>
    /// </list>
    /// </summary>
    public static class SessionRegistryRules
    {
        /// <summary>The first line of the registry file.</summary>
        public const string Header = "# OutlookAI sessions attached to Outlook, and its lifetime pin (Q118 (a), Q208 B). Rewritten under a lock; safe to delete while no OutlookAI runs.";

        /// <summary>
        /// Drops every session whose process is gone and every Outlook that is gone, and returns how many sessions
        /// went. A session of THIS process counts as gone when <paramref name="liveInThisProcess"/> does not know
        /// its key - one that was never left, because its owner never disposed it.
        /// </summary>
        public static int Prune(
            SessionRegistryState state,
            Func<ProcessStamp, bool> isRunning,
            ProcessStamp thisProcess,
            Func<string, bool>? liveInThisProcess)
        {
            if (state == null)
            {
                throw new ArgumentNullException(nameof(state));
            }

            if (isRunning == null)
            {
                throw new ArgumentNullException(nameof(isRunning));
            }

            Dictionary<ProcessStamp, bool> running = new Dictionary<ProcessStamp, bool>();
            bool Running(ProcessStamp stamp)
            {
                if (!running.TryGetValue(stamp, out bool alive))
                {
                    alive = isRunning(stamp);
                    running[stamp] = alive;
                }

                return alive;
            }

            int before = state.Sessions.Count;
            state.Sessions.RemoveAll(s =>
                !Running(s.Process)
                || !Running(s.Outlook)
                || (s.Process == thisProcess && liveInThisProcess != null && !liveInThisProcess(s.Key)));
            state.Outlooks.RemoveAll(o => !Running(o.Outlook));
            return before - state.Sessions.Count;
        }

        /// <summary>Adds <paramref name="session"/> (replacing one with its key) and the record of its Outlook.</summary>
        public static void Join(SessionRegistryState state, RegisteredSession session)
        {
            if (state == null)
            {
                throw new ArgumentNullException(nameof(state));
            }

            if (session == null)
            {
                throw new ArgumentNullException(nameof(session));
            }

            state.Sessions.RemoveAll(s => string.Equals(s.Key, session.Key, StringComparison.Ordinal));
            state.Sessions.Add(session);
            RegisteredOutlook outlook = state.FindOrAdd(session.Outlook);
            if (session.StartedOutlook && outlook.StartedBy == null)
            {
                outlook.StartedBy = session.Key;
            }
        }

        /// <summary>
        /// Takes the session with <paramref name="key"/> out and says whether its Outlook's pin goes with it: when
        /// <paramref name="heardQuit"/>, always (the user's exit wins); otherwise only when no other live session is
        /// attached to the same Outlook. The pin record is NOT cleared here - <see cref="PinClosed"/> does that once
        /// the close has happened, so a close that failed leaves the record for the next session to retry.
        /// </summary>
        public static LeaveVerdict Leave(SessionRegistryState state, string key, ProcessStamp outlook, bool heardQuit)
        {
            if (state == null)
            {
                throw new ArgumentNullException(nameof(state));
            }

            state.Sessions.RemoveAll(s => string.Equals(s.Key, key, StringComparison.Ordinal));
            int others = state.Sessions.Count(s => s.Outlook == outlook);
            RegisteredOutlook? record = state.Find(outlook);
            if (record == null || record.PinWindow == 0)
            {
                return new LeaveVerdict(0, "no pin", others);
            }

            if (heardQuit)
            {
                return new LeaveVerdict(record.PinWindow, "Outlook was told to quit", others);
            }

            return others == 0
                ? new LeaveVerdict(record.PinWindow, "last session out", 0)
                : new LeaveVerdict(0, "other sessions still attached", others);
        }

        /// <summary>Records that the pin with <paramref name="window"/> is gone (closed, or its window no longer exists).</summary>
        public static void PinClosed(SessionRegistryState state, ProcessStamp outlook, long window)
        {
            RegisteredOutlook? record = state?.Find(outlook);
            if (record != null && (window == 0 || record.PinWindow == window))
            {
                record.PinWindow = 0;
                record.PinMadeBy = null;
                record.PinSinceUtcTicks = 0;
            }
        }

        /// <summary>Records a new pin of <paramref name="outlook"/>.</summary>
        public static void PinMade(SessionRegistryState state, ProcessStamp outlook, long window, string madeBy, long utcTicks)
        {
            if (state == null)
            {
                throw new ArgumentNullException(nameof(state));
            }

            RegisteredOutlook record = state.FindOrAdd(outlook);
            record.PinWindow = window;
            record.PinMadeBy = madeBy;
            record.PinSinceUtcTicks = utcTicks;
        }

        /// <summary>
        /// The route to a draft's Word editor (Q208 B). Direct whenever Outlook shows a window a person could see -
        /// no per-draft activation at all then (cell B3, 40 of 40) - and in a window-less Outlook once this process
        /// has been promoted (cells W6 and W7, 120 of 120). The promotion otherwise, and after a direct read was
        /// refused: that is what puts a window on screen, so it happens once per Outlook process unless a direct
        /// read fails.
        /// </summary>
        public static DraftEditorRoute DecideEditorRoute(bool outlookShowsAWindow, RegisteredOutlook? record)
        {
            if (record != null && record.DirectRefused)
            {
                return DraftEditorRoute.Promote;
            }

            if (outlookShowsAWindow || (record != null && record.PromotedUtcTicks != 0))
            {
                return DraftEditorRoute.Direct;
            }

            return DraftEditorRoute.Promote;
        }

        /// <summary>Records what happened to a draft's editor: a promotion that worked, or a direct read that was refused.</summary>
        public static void EditorOutcome(SessionRegistryState state, ProcessStamp outlook, DraftEditorRoute route, bool editorObtained, long utcTicks)
        {
            if (state == null)
            {
                throw new ArgumentNullException(nameof(state));
            }

            RegisteredOutlook record = state.FindOrAdd(outlook);
            if (route == DraftEditorRoute.Promote && editorObtained)
            {
                if (record.PromotedUtcTicks == 0)
                {
                    record.PromotedUtcTicks = utcTicks;
                }

                record.DirectRefused = false;
            }
            else if (route == DraftEditorRoute.Direct && !editorObtained)
            {
                record.DirectRefused = true;
            }
        }

        /// <summary>
        /// One sentence for outlook_health: what keeps this Outlook running and what will end it. Pure.
        /// </summary>
        public static string DescribeWhatEndsOutlook(bool outlookRunning, bool outlookShowsAWindow, RegisteredOutlook? record, int sessionsAttached)
        {
            if (!outlookRunning)
            {
                return "Outlook is not running.";
            }

            if (outlookShowsAWindow)
            {
                return "Outlook has a window open: it ends when the user closes its last window or chooses Exit.";
            }

            if (record != null && record.PinWindow != 0)
            {
                return sessionsAttached > 0
                    ? "Outlook has no window; OutlookAI's hidden window keeps it running for the drafts, and the last OutlookAI session to end closes it, which ends Outlook."
                    : "Outlook has no window and an OutlookAI hidden window is left with no session attached; the next OutlookAI session to start adopts it and closes it when it ends, and File > Exit ends Outlook now.";
            }

            return sessionsAttached > 0
                ? "Outlook has no window and no OutlookAI hidden window: it keeps running while programs use it, and ends on its own some minutes after the last one lets go."
                : "Outlook has no window and no OutlookAI session is attached: it ends on its own some minutes after the last program lets go.";
        }

        /// <summary>The registry file's text. Pure.</summary>
        public static string Format(SessionRegistryState state)
        {
            if (state == null)
            {
                throw new ArgumentNullException(nameof(state));
            }

            StringBuilder text = new StringBuilder();
            text.Append(Header).Append('\n');
            foreach (RegisteredOutlook o in state.Outlooks)
            {
                text.Append("outlook stamp=").Append(o.Outlook.ToString());
                AppendIf(text, "startedBy", o.StartedBy);
                if (o.PromotedUtcTicks != 0)
                {
                    text.Append(" promoted=").Append(o.PromotedUtcTicks.ToString(CultureInfo.InvariantCulture));
                }

                if (o.DirectRefused)
                {
                    text.Append(" directRefused=1");
                }

                if (o.PinWindow != 0)
                {
                    text.Append(" pin=").Append(o.PinWindow.ToString(CultureInfo.InvariantCulture));
                    AppendIf(text, "pinBy", o.PinMadeBy);
                    text.Append(" pinSince=").Append(o.PinSinceUtcTicks.ToString(CultureInfo.InvariantCulture));
                }

                text.Append('\n');
            }

            foreach (RegisteredSession s in state.Sessions)
            {
                text.Append("session key=").Append(s.Key)
                    .Append(" process=").Append(s.Process.ToString())
                    .Append(" outlook=").Append(s.Outlook.ToString())
                    .Append(" since=").Append(s.SinceUtcTicks.ToString(CultureInfo.InvariantCulture))
                    .Append(" startedOutlook=").Append(s.StartedOutlook ? "1" : "0")
                    .Append('\n');
            }

            return text.ToString();
        }

        /// <summary>
        /// Reads the registry file's text back. Pure and forgiving: a line it does not understand is dropped, never
        /// thrown on, so a damaged file costs the records on its damaged lines and nothing else.
        /// </summary>
        public static SessionRegistryState Parse(string? text)
        {
            SessionRegistryState state = new SessionRegistryState();
            if (string.IsNullOrEmpty(text))
            {
                return state;
            }

            foreach (string raw in text!.Split('\n'))
            {
                string line = raw.Trim();
                if (line.Length == 0 || line[0] == '#')
                {
                    continue;
                }

                string[] parts = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                Dictionary<string, string> fields = new Dictionary<string, string>(StringComparer.Ordinal);
                for (int i = 1; i < parts.Length; i++)
                {
                    int eq = parts[i].IndexOf('=');
                    if (eq > 0)
                    {
                        fields[parts[i].Substring(0, eq)] = parts[i].Substring(eq + 1);
                    }
                }

                if (parts[0] == "outlook")
                {
                    ProcessStamp? stamp = ProcessStamp.TryParse(Field(fields, "stamp"));
                    if (stamp == null || state.Find(stamp.Value) != null)
                    {
                        continue;
                    }

                    RegisteredOutlook o = new RegisteredOutlook(stamp.Value)
                    {
                        StartedBy = Field(fields, "startedBy"),
                        PromotedUtcTicks = Long(fields, "promoted"),
                        DirectRefused = Field(fields, "directRefused") == "1",
                        PinWindow = Long(fields, "pin"),
                        PinSinceUtcTicks = Long(fields, "pinSince"),
                    };
                    o.PinMadeBy = o.PinWindow == 0 ? null : Field(fields, "pinBy");
                    state.Outlooks.Add(o);
                }
                else if (parts[0] == "session")
                {
                    string? key = Field(fields, "key");
                    ProcessStamp? process = ProcessStamp.TryParse(Field(fields, "process"));
                    ProcessStamp? outlook = ProcessStamp.TryParse(Field(fields, "outlook"));
                    if (string.IsNullOrEmpty(key) || process == null || outlook == null
                        || state.Sessions.Any(s => string.Equals(s.Key, key, StringComparison.Ordinal)))
                    {
                        continue;
                    }

                    state.Sessions.Add(new RegisteredSession(key!, process.Value, outlook.Value, Long(fields, "since"), Field(fields, "startedOutlook") == "1"));
                }
            }

            return state;
        }

        private static void AppendIf(StringBuilder text, string name, string? value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                text.Append(' ').Append(name).Append('=').Append(value);
            }
        }

        private static string? Field(Dictionary<string, string> fields, string name)
        {
            return fields.TryGetValue(name, out string? value) && value.Length > 0 ? value : null;
        }

        private static long Long(Dictionary<string, string> fields, string name)
        {
            return fields.TryGetValue(name, out string? value)
                && long.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out long parsed)
                ? parsed
                : 0;
        }
    }

    /// <summary>How an update of the registry went.</summary>
    public enum RegistryAccess
    {
        /// <summary>Read and written under the lock.</summary>
        Locked = 0,

        /// <summary>The lock was not had within its wait; the change was made to this process's copy only.</summary>
        LockTimedOut = 1,

        /// <summary>The lock or the file could not be used at all; the change was made to this process's copy only.</summary>
        Unavailable = 2,
    }

    /// <summary>
    /// THE PER-USER SESSION REGISTRY (Q118 (a)): a small text file under <c>%LOCALAPPDATA%\OutlookAI</c>, read and
    /// rewritten only under a named mutex, so every OutlookAI process of the user - one COM host per agent
    /// session - sees the same sessions, the same lifetime pin and the same promotion record. The rules are
    /// <see cref="SessionRegistryRules"/>; this class only locks, loads, applies and saves.
    /// <para>
    /// Best-effort by design: a registry that cannot be locked or read never stops a draft. The change is then
    /// applied to this process's own copy, which keeps the promotion once per process here, and the caller is
    /// told (<see cref="RegistryAccess"/>).
    /// </para>
    /// </summary>
    public sealed class OutlookSessionRegistry
    {
        /// <summary>The file's name under the OutlookAI state folder.</summary>
        public const string FileName = "outlook-sessions.txt";

        /// <summary>The lock every OutlookAI process of this logon session shares.</summary>
        public const string DefaultMutexName = @"Local\OutlookAI.OutlookSessions";

        /// <summary>How long a change waits for the lock: well past a promotion (under a second) and a cold pin (a few).</summary>
        public const int LockWaitMilliseconds = 30_000;

        private static long _sequence;
        private static readonly object LocalLock = new object();
        private static readonly HashSet<string> LiveSessionsOfThisProcess = new HashSet<string>(StringComparer.Ordinal);

        private readonly string _path;
        private readonly string _mutexName;
        private readonly Func<ProcessStamp, bool> _isRunning;
        private readonly int _lockWaitMilliseconds;
        private SessionRegistryState _localCopy = new SessionRegistryState();

        /// <summary>Creates a registry over <paramref name="path"/>, locked by <paramref name="mutexName"/>.</summary>
        public OutlookSessionRegistry(string path, string mutexName, Func<ProcessStamp, bool>? isRunning = null, int lockWaitMilliseconds = LockWaitMilliseconds)
        {
            _path = path ?? throw new ArgumentNullException(nameof(path));
            _mutexName = mutexName ?? throw new ArgumentNullException(nameof(mutexName));
            _isRunning = isRunning ?? ProcessStamp.IsRunning;
            _lockWaitMilliseconds = lockWaitMilliseconds;
        }

        /// <summary>The registry every product session uses: <c>%LOCALAPPDATA%\OutlookAI\outlook-sessions.txt</c>.</summary>
        public static OutlookSessionRegistry Default { get; } = new OutlookSessionRegistry(
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookAI", FileName),
            DefaultMutexName);

        /// <summary>The file this registry reads and writes.</summary>
        public string FilePath => _path;

        /// <summary>A key no other session of any process will ever have: this process's stamp and a sequence number.</summary>
        public static string NewSessionKey(ProcessStamp thisProcess)
        {
            long sequence = Interlocked.Increment(ref _sequence);
            return thisProcess.Pid.ToString(CultureInfo.InvariantCulture) + "."
                + thisProcess.StartedUtcTicks.ToString(CultureInfo.InvariantCulture) + "."
                + sequence.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>Marks <paramref name="key"/> as a session of this process that is still open (or no longer, when false).</summary>
        public static void MarkLiveInThisProcess(string key, bool live)
        {
            lock (LocalLock)
            {
                if (live)
                {
                    _ = LiveSessionsOfThisProcess.Add(key);
                }
                else
                {
                    _ = LiveSessionsOfThisProcess.Remove(key);
                }
            }
        }

        private static bool IsLiveInThisProcess(string key)
        {
            lock (LocalLock)
            {
                return LiveSessionsOfThisProcess.Contains(key);
            }
        }

        /// <summary>
        /// Locks, loads, prunes, applies <paramref name="change"/>, saves and unlocks - or, when the lock or the file
        /// cannot be had, applies the change to this process's own copy. Runs <paramref name="change"/> exactly once
        /// either way, on the calling thread, so a change may make COM calls (the lock is a mutex: the thread that
        /// takes it releases it).
        /// </summary>
        public T Update<T>(Func<SessionRegistryState, T> change, out RegistryAccess access)
        {
            if (change == null)
            {
                throw new ArgumentNullException(nameof(change));
            }

            ProcessStamp thisProcess = ProcessStamp.Current();
            Mutex? mutex = null;
            bool owned = false;
            try
            {
                try
                {
                    mutex = new Mutex(initiallyOwned: false, _mutexName);
                    try
                    {
                        owned = mutex.WaitOne(_lockWaitMilliseconds);
                    }
                    catch (AbandonedMutexException)
                    {
                        // The holder died holding it (a COM host killed mid-change); ownership passed to this thread.
                        owned = true;
                    }
                }
                catch (Exception ex) when (ex is UnauthorizedAccessException || ex is IOException
                    || ex is WaitHandleCannotBeOpenedException || ex is ArgumentException)
                {
                    owned = false;
                }

                if (owned)
                {
                    SessionRegistryState? loaded = TryLoad();
                    if (loaded != null)
                    {
                        SessionRegistryRules.Prune(loaded, _isRunning, thisProcess, IsLiveInThisProcess);
                        T result = change(loaded);
                        if (TrySave(loaded))
                        {
                            lock (LocalLock)
                            {
                                _localCopy = loaded;
                            }

                            access = RegistryAccess.Locked;
                            return result;
                        }

                        access = RegistryAccess.Unavailable;
                        lock (LocalLock)
                        {
                            _localCopy = loaded;
                        }

                        return result;
                    }
                }

                access = owned ? RegistryAccess.Unavailable : (mutex == null ? RegistryAccess.Unavailable : RegistryAccess.LockTimedOut);
                SessionRegistryState local;
                lock (LocalLock)
                {
                    local = _localCopy;
                }

                SessionRegistryRules.Prune(local, _isRunning, thisProcess, IsLiveInThisProcess);
                return change(local);
            }
            finally
            {
                if (owned && mutex != null)
                {
                    try
                    {
                        mutex.ReleaseMutex();
                    }
                    catch (ApplicationException)
                    {
                        // Not owned by this thread after all; nothing to release.
                    }
                }

                mutex?.Dispose();
            }
        }

        /// <summary>
        /// The registry as the file holds it now, without the lock and without pruning - for outlook_health, which
        /// must never wait. Null when there is no file or it cannot be read.
        /// </summary>
        public SessionRegistryState? TryRead()
        {
            return TryLoad(missingIsEmpty: false);
        }

        private SessionRegistryState? TryLoad(bool missingIsEmpty = true)
        {
            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    if (!File.Exists(_path))
                    {
                        return missingIsEmpty ? new SessionRegistryState() : null;
                    }

                    using (FileStream stream = new FileStream(_path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                    using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        return SessionRegistryRules.Parse(reader.ReadToEnd());
                    }
                }
                catch (FileNotFoundException)
                {
                    return missingIsEmpty ? new SessionRegistryState() : null;
                }
                catch (DirectoryNotFoundException)
                {
                    return missingIsEmpty ? new SessionRegistryState() : null;
                }
                catch (IOException)
                {
                    // Being replaced this instant: once more.
                    Thread.Sleep(20);
                }
                catch (UnauthorizedAccessException)
                {
                    return null;
                }
            }

            return null;
        }

        private bool TrySave(SessionRegistryState state)
        {
            string? directory = Path.GetDirectoryName(_path);
            string temporary = _path + "." + ProcessStamp.Current().Pid.ToString(CultureInfo.InvariantCulture) + ".tmp";
            try
            {
                if (!string.IsNullOrEmpty(directory))
                {
                    _ = Directory.CreateDirectory(directory);
                }

                File.WriteAllText(temporary, SessionRegistryRules.Format(state), new UTF8Encoding(false));
                if (File.Exists(_path))
                {
                    File.Replace(temporary, _path, null);
                }
                else
                {
                    File.Move(temporary, _path);
                }

                return true;
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                try
                {
                    File.Delete(temporary);
                }
                catch (Exception cleanup) when (cleanup is IOException || cleanup is UnauthorizedAccessException)
                {
                }

                return false;
            }
        }
    }
}
