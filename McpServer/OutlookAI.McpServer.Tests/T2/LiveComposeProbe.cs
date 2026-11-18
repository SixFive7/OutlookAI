using System.ComponentModel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using OutlookAI.Core.AddIn;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The live half of the Q149 (d) measurement: the COM client each trial runs in, the window-less staging,
/// the process and window watches, and every write the measurement makes to the test mailbox.
/// <para>
/// <b>Mailbox safety (AGENTS.md rules 1-8).</b> Every item written here is a draft in the TEST HUB's own
/// folders (<see cref="LiveStoreWriteGuard"/> refuses any other store before Outlook is asked), carries
/// <see cref="LiveOutlookTestMailer.SubjectTag"/> and the run's marker in its subject from the moment it is
/// first saved (<see cref="Q149ProbeGuards.RefuseSubject"/>), and is recorded in the run's EntryID
/// allowlist. A change to an existing item checks the tag, the marker AND the allowlist first; a delete
/// goes through <see cref="LiveOutlookTestMailer.DeleteItemByEntryId"/> (EntryID plus ordinal tag and
/// marker), and the collection's fixture ends with the tested sweep and the count tripwire. Nothing is sent
/// to anyone but the hub itself (<see cref="Q149ProbeGuards.RefuseRecipients"/>), and only through the
/// product's own two-step send.
/// </para>
/// <para>
/// <b>Outlook's lifecycle (rule 7).</b> OUTLOOK.EXE is never killed. It is ended and started through
/// <see cref="LiveOutlookRestart"/>, the one stop the live tests share: closing its windows (WM_CLOSE, what a
/// user's close sends), closing its Explorers once no window is visible, or <c>Application.Quit()</c> from a
/// client that holds nothing else - each only after the S7 counts say no Inspector is open and every Outbox
/// is empty - and a stop that does not end it fails the cell instead of escalating. No inspector is ever opened on an object <c>Items.Add</c>, <c>Reply</c> or
/// <c>Forward</c> returned: that close crashed Outlook about once in 40 (runbook 4.6), so every draft is
/// saved, released and re-opened by EntryID first.
/// </para>
/// </summary>
public static class LiveComposeProbe
{
    /// <summary>What the measurement's stops name as what did not run (<see cref="LiveOutlookRestart.Stop"/>).</summary>
    internal const string Scenario = "the Q149 measurement, which ends and restarts Outlook";

    /// <summary>How long one probe STA step may block before the trial fails naming it.</summary>
    internal static readonly TimeSpan StepBudget = TimeSpan.FromMinutes(3);

    /// <summary>olFolderDrafts, olFolderInbox.</summary>
    internal const int Drafts = 16;

    internal const int Inbox = 6;

    /// <summary>OlInspectorClose: olSave 0, olDiscard 1.</summary>
    internal const int OlSave = 0;

    internal const int OlDiscard = 1;

    /// <summary>The attachment properties the HTML-only draft sets (probe row 12; PR_ATTACH_FLAGS is not writable and not needed).</summary>
    internal const string PrAttachContentId = "http://schemas.microsoft.com/mapi/proptag/0x3712001F";

    internal const string PrAttachmentHidden = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B";

    internal const string PrAttachMimeTag = "http://schemas.microsoft.com/mapi/proptag/0x370E001F";

    /// <summary>The form-definition properties B3 checks are absent (PSETID_Common).</summary>
    internal static readonly IReadOnlyList<(string Name, string Tag)> FormProperties = new[]
    {
        ("PidLidFormPropStream", "http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/851B0102"),
        ("PidLidPageDirStream", "http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/85130102"),
        ("PidLidPropertyDefinitionStream", "http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/85400102"),
        ("PidLidScriptStream", "http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/85410102"),
    };

    /// <summary>
    /// Which of <see cref="FormProperties"/> an item carries - none on an ordinary draft. The form-page commit (B3) adds
    /// Outlook's "Message" page to the inspector's form, which could turn an item into a one-off form; these are the
    /// properties such an item gains.
    /// </summary>
    internal static IReadOnlyList<string> FormPropertiesPresent(object item)
    {
        List<string> present = new();
        dynamic accessor = ((dynamic)item).PropertyAccessor;
        try
        {
            foreach ((string name, string tag) in FormProperties)
            {
                try
                {
                    object? value = accessor.GetProperty(tag);
                    if (value != null)
                    {
                        present.Add(name);
                    }
                }
                catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                {
                    // Absent: the property is not there to read.
                }
            }
        }
        finally
        {
            Q149Com.Release((object)accessor);
        }

        return present;
    }

    /// <summary>Runs <paramref name="work"/> on its own thread and gives up after <paramref name="budget"/>, naming the step.</summary>
    internal static T Bounded<T>(string step, TimeSpan budget, Func<T> work)
    {
        return LiveOutlookRestart.Bounded("Q149: '" + step + "'", budget, work);
    }

    /// <summary>
    /// Ends Outlook for the measurement through <see cref="LiveOutlookRestart.Stop"/>, the one graceful stop the live
    /// tests share (rule 7); the measurement does not run on past a stop that stood aside for a person at the keyboard.
    /// </summary>
    internal static void StopOutlook(LiveTestSettings settings, Action<string> log)
    {
        if (!LiveOutlookRestart.Stop(settings, Scenario, log))
        {
            throw new InvalidOperationException("Q149: stood aside for a user at the keyboard - the measurement did not run.");
        }
    }

    /// <summary>
    /// Starts Outlook the way a user does - OUTLOOK.EXE on the default profile, so with its main window -
    /// and waits until that window is up and the object model answers. Stops a running window-less Outlook
    /// first, so the measurement's "windowed" state is a normally started Outlook and not a promoted one.
    /// </summary>
    internal static void EnsureWindowedOutlook(LiveTestSettings settings, Action<string> log)
    {
        EnsureWindowedOutlook(settings, log, coldStart: null);
    }

    /// <summary>
    /// <see cref="EnsureWindowedOutlook(LiveTestSettings, Action{string})"/>, telling <paramref name="coldStart"/> - when
    /// given - when OUTLOOK.EXE was started, when its window showed and when its object model first answered (Q180 (b)).
    /// </summary>
    internal static void EnsureWindowedOutlook(LiveTestSettings settings, Action<string> log, Q180ColdStartWatch? coldStart)
    {
        if (!LiveOutlookRestart.EnsureWindowed(settings, Scenario, log, windowBudget: null, coldStart))
        {
            throw new InvalidOperationException("Q149: stood aside for a user at the keyboard - the measurement did not run.");
        }
    }

    /// <summary>Top-level windows of OUTLOOK.EXE with class OpusApp - Word's frame, inside Outlook when Word is its editor.</summary>
    internal static int CountOpusAppWindows()
    {
        int count = 0;
        foreach (Q149Native.TopWindow w in Q149Native.OutlookTopWindows())
        {
            if (string.Equals(w.ClassName, "OpusApp", StringComparison.Ordinal))
            {
                count++;
            }
        }

        return count;
    }

    /// <summary>The frames (class rctrl_renwnd32) of every OUTLOOK.EXE now - taken before an Inspector or Explorer is made.</summary>
    internal static HashSet<IntPtr> FrameSnapshot()
    {
        return Q149Native.OutlookTopWindows()
            .Where(w => string.Equals(w.ClassName, "rctrl_renwnd32", StringComparison.Ordinal))
            .Select(w => w.Handle)
            .ToHashSet();
    }

    /// <summary>
    /// The frame of an Inspector or Explorer just made: through IOleWindow when Outlook answers it, else the
    /// one frame that was not there before. Says which way it was found.
    /// </summary>
    internal static (IntPtr Handle, string FoundBy) FrameOf(object outlookWindowObject, HashSet<IntPtr> framesBefore)
    {
        IntPtr hwnd = WindowOf(outlookWindowObject);
        if (hwnd != IntPtr.Zero)
        {
            return (hwnd, "IOleWindow");
        }

        List<IntPtr> added = FrameSnapshot().Where(h => !framesBefore.Contains(h)).ToList();
        return added.Count == 1 ? (added[0], "snapshot") : (IntPtr.Zero, "none (" + added.Count.ToString(CultureInfo.InvariantCulture) + " new frames)");
    }

    /// <summary>The handle of an Inspector's or Explorer's frame, through IOleWindow; zero when Outlook will not say.</summary>
    internal static IntPtr WindowOf(object outlookWindowObject)
    {
        try
        {
            if (outlookWindowObject is Q149Native.IOleWindow window && window.GetWindow(out IntPtr hwnd) == 0)
            {
                return hwnd;
            }
        }
        catch (Exception ex) when (ex is InvalidCastException or COMException or InvalidComObjectException)
        {
        }

        return IntPtr.Zero;
    }
}

/// <summary>
/// Q180 (b), decided 2026-11-18: one cold start of Outlook, logged - the time until it is ready, every liveness verdict
/// Windows gives on the way (<see cref="OutlookLiveness"/>, sampled every 100 ms from the start request until two
/// seconds after ready), the milestones the caller reports, and any hang: a step that did not return within its
/// budget, or verdicts of "not responding". Written as one <see cref="Q149Cells.Q180"/> attempt line by
/// <see cref="Finish"/>. It only watches: nothing here touches COM or Outlook's windows.
/// </summary>
internal sealed class Q180ColdStartWatch : IDisposable
{
    private const int SampleMs = 100;
    private const int AfterReadyMs = 2000;

    private readonly Stopwatch _clock = Stopwatch.StartNew();
    private readonly object _gate = new();
    private readonly List<(long Ms, string Verdict)> _samples = new();
    private readonly List<string> _details = new();
    private readonly SortedDictionary<string, long> _milestones = new(StringComparer.Ordinal);
    private readonly List<string> _hangs = new();
    private readonly Thread _sampler;
    private readonly string _kind;
    private readonly string _during;
    private readonly int _iteration;
    private volatile bool _stop;
    private long _readyMs = -1;
    private long _processSeenMs = -1;
    private bool _finished;

    /// <param name="kind">"com" for a start through Outlook.Application, "windowed" for OUTLOOK.EXE started with its window.</param>
    /// <param name="during">The cell whose start this is.</param>
    /// <param name="iteration">Which of its starts.</param>
    public Q180ColdStartWatch(string kind, string during, int iteration)
    {
        _kind = kind;
        _during = during;
        _iteration = iteration;
        _sampler = new Thread(Sample) { IsBackground = true, Name = "Q180.ColdStartWatch" };
        _sampler.Start();
    }

    /// <summary>Milliseconds since the start was requested.</summary>
    public long ElapsedMs => _clock.ElapsedMilliseconds;

    /// <summary>Records a named moment of the start (the first one of each name counts).</summary>
    public void Milestone(string name)
    {
        lock (_gate)
        {
            if (!_milestones.ContainsKey(name))
            {
                _milestones[name] = _clock.ElapsedMilliseconds;
            }
        }
    }

    /// <summary>Outlook answers: the object model gave its first answer. Sampling goes on for two seconds more.</summary>
    public void Ready()
    {
        Interlocked.CompareExchange(ref _readyMs, _clock.ElapsedMilliseconds, -1);
        Milestone("ready");
    }

    /// <summary>Records a hang: a step of the start that did not return within its budget.</summary>
    public void Hang(string step)
    {
        lock (_gate)
        {
            _hangs.Add(step + " at " + _clock.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + " ms");
        }
    }

    /// <summary>Stops sampling (after the grace past ready) and writes the line. Safe to call once; later calls do nothing.</summary>
    public void Finish(Q149Run run)
    {
        ArgumentNullException.ThrowIfNull(run);
        if (_finished)
        {
            return;
        }

        _finished = true;
        long ready = Interlocked.Read(ref _readyMs);
        if (ready >= 0)
        {
            long wait = ready + AfterReadyMs - _clock.ElapsedMilliseconds;
            if (wait > 0)
            {
                Thread.Sleep((int)Math.Min(wait, AfterReadyMs));
            }
        }

        _stop = true;
        _sampler.Join(TimeSpan.FromSeconds(5));
        Q149Attempt attempt = new(Q149Cells.Q180, _iteration);
        lock (_gate)
        {
            (IReadOnlyList<object[]> changes, int hungSamples, long longestHungMs, long? firstResponsive) = Q180Verdicts.Summarize(_samples);
            attempt.Set("kind", _kind)
                .Set("during", _during)
                .Set("readyMs", ready < 0 ? null : ready)
                .Set("processSeenMs", Interlocked.Read(ref _processSeenMs) < 0 ? null : Interlocked.Read(ref _processSeenMs))
                .Set("firstResponsiveMs", firstResponsive)
                .Set("verdicts", changes.ToArray())
                .Set("verdictDetails", _details.Distinct(StringComparer.Ordinal).Take(12).ToArray())
                .Set("samples", _samples.Count)
                .Set("hungSamples", hungSamples)
                .Set("longestHungMs", longestHungMs)
                .Set("milestones", new SortedDictionary<string, long>(_milestones, StringComparer.Ordinal))
                .Set("hangs", _hangs.Count == 0 ? null : _hangs.ToArray())
                .Set("hung", _hangs.Count > 0 || hungSamples > 0);
        }

        run.Write(attempt);
    }

    public void Dispose()
    {
        _stop = true;
        _sampler.Join(TimeSpan.FromSeconds(5));
    }

    private void Sample()
    {
        while (!_stop)
        {
            long now = _clock.ElapsedMilliseconds;
            long ready = Interlocked.Read(ref _readyMs);
            if (ready >= 0 && now > ready + AfterReadyMs)
            {
                return;
            }

            OutlookLivenessState state = OutlookLiveness.Probe(out string detail);
            if (state != OutlookLivenessState.NotRunning)
            {
                Interlocked.CompareExchange(ref _processSeenMs, now, -1);
            }

            lock (_gate)
            {
                _samples.Add((now, state.ToString()));
                if (_details.Count < 64 && (_details.Count == 0 || !string.Equals(_details[_details.Count - 1], state + ": " + detail, StringComparison.Ordinal)))
                {
                    _details.Add(state + ": " + detail);
                }
            }

            Thread.Sleep(SampleMs);
        }
    }
}

/// <summary>
/// Q180 (b): EVERY start of OUTLOOK.EXE while the measurement runs, whoever caused it - a cell's cold start, the
/// client that re-reads the drafts, the fixture's teardown - watched from outside as it comes up: from the moment it
/// is first seen until two seconds after Windows first calls it responsive, until it exits, or for three minutes, with
/// <see cref="OutlookLiveness"/> sampled every 100 ms. One <see cref="Q149Cells.Q180"/> line per start, kind
/// "observed", written to the console capture only - it runs on its own thread, past the end of any one test. A start
/// a cell makes itself is also logged by that cell's <see cref="Q180ColdStartWatch"/>, with the moment the object
/// model first answered.
/// </summary>
internal sealed class Q180StartObserver : IDisposable
{
    private const int SampleMs = 100;
    private const int AfterResponsiveMs = 2000;
    private const int WatchLimitMs = 180_000;

    private readonly HashSet<int> _known = new();
    private readonly Thread _thread;
    private volatile bool _stop;
    private int _starts;

    public Q180StartObserver()
    {
        foreach (int pid in OutlookPids())
        {
            _ = _known.Add(pid);
        }

        _thread = new Thread(Watch) { IsBackground = true, Name = "Q180.StartObserver" };
        _thread.Start();
    }

    /// <summary>How many starts it has logged.</summary>
    public int Starts => Volatile.Read(ref _starts);

    public void Dispose()
    {
        _stop = true;
        _thread.Join(TimeSpan.FromSeconds(10));
    }

    private void Watch()
    {
        while (!_stop)
        {
            try
            {
                foreach (int pid in OutlookPids())
                {
                    if (_known.Add(pid))
                    {
                        Observe(pid, Interlocked.Increment(ref _starts));
                    }
                }
            }
            catch (Exception ex)
            {
                // An observer must never take the run down; the gap is logged instead.
                Console.WriteLine(new Q149Attempt(Q149Cells.Q180, 999).Set("kind", "observed").Set("observerFailure", ex.GetType().Name).ToLine());
            }

            Thread.Sleep(SampleMs);
        }
    }

    private void Observe(int pid, int start)
    {
        Q149Attempt attempt = new(Q149Cells.Q180, 1000 + start);
        attempt.Set("kind", "observed").Set("pid", pid);
        // The kernel's creation time, as UTC: no machine time zone is read (Q144).
        long? createdTicks = ProcessStamp.CreationUtcTicks(pid);
        DateTime? startedUtc = createdTicks == null ? null : new DateTime(createdTicks.Value, DateTimeKind.Utc);
        using (Process? process = TryOpen(pid))
        {

            string? commandLine = Q149Native.TryReadCommandLine(pid);
            attempt.Set("embedding", commandLine == null ? null : commandLine.Contains("Embedding", StringComparison.OrdinalIgnoreCase))
                .Set("firstSeenMsAfterStart", startedUtc == null ? null : (long)(DateTime.UtcNow - startedUtc.Value).TotalMilliseconds);
            List<(long Ms, string Verdict)> samples = new();
            List<string> details = new();
            long? responsiveAt = null;
            bool exited = false;
            Stopwatch watched = Stopwatch.StartNew();
            while (!_stop && watched.ElapsedMilliseconds < WatchLimitMs)
            {
                if (process == null || HasExited(process))
                {
                    exited = true;
                    break;
                }

                long ms = startedUtc == null ? watched.ElapsedMilliseconds : (long)(DateTime.UtcNow - startedUtc.Value).TotalMilliseconds;
                OutlookLivenessState state = OutlookLiveness.Probe(out string detail);
                samples.Add((ms, state.ToString()));
                if (details.Count < 32 && (details.Count == 0 || details[details.Count - 1] != state + ": " + detail))
                {
                    details.Add(state + ": " + detail);
                }

                if (state == OutlookLivenessState.Responsive && responsiveAt == null)
                {
                    responsiveAt = watched.ElapsedMilliseconds;
                }

                if (responsiveAt != null && watched.ElapsedMilliseconds > responsiveAt.Value + AfterResponsiveMs)
                {
                    break;
                }

                Thread.Sleep(SampleMs);
            }

            (IReadOnlyList<object[]> changes, int hungSamples, long longestHungMs, long? firstResponsive) = Q180Verdicts.Summarize(samples);
            attempt.Set("readyMs", firstResponsive)
                .Set("verdicts", changes.ToArray())
                .Set("verdictDetails", details.ToArray())
                .Set("samples", samples.Count)
                .Set("hungSamples", hungSamples)
                .Set("longestHungMs", longestHungMs)
                .Set("exitedBeforeResponsive", exited && firstResponsive == null)
                .Set("neverResponsiveWithin", firstResponsive == null && !exited ? WatchLimitMs : null)
                .Set("hung", hungSamples > 0 || (firstResponsive == null && !exited));
        }

        Console.WriteLine(attempt.ToLine());
    }

    private static Process? TryOpen(int pid)
    {
        try
        {
            return Process.GetProcessById(pid);
        }
        catch (ArgumentException)
        {
            return null;
        }
    }

    private static bool HasExited(Process process)
    {
        try
        {
            return process.HasExited;
        }
        catch (Exception ex) when (ex is InvalidOperationException or Win32Exception)
        {
            return true;
        }
    }

    private static List<int> OutlookPids()
    {
        Process[] processes = Process.GetProcessesByName("OUTLOOK");
        try
        {
            return processes.Select(p => p.Id).ToList();
        }
        finally
        {
            foreach (Process p in processes)
            {
                p.Dispose();
            }
        }
    }
}

/// <summary>
/// A small visible window of the test's own that holds the foreground, standing in for the program a person is typing
/// in while an agent drafts: before each draft it is brought to the foreground, and after it the cells record whether
/// it still is. Whether a draft TAKES THE FOCUS is that question - a window shown off-screen or hidden that took the
/// foreground would take the person's keystrokes with it. It shows nothing but its own empty window, and sends input
/// only when a cell has it type (<see cref="StartTyping"/>): one character no draft holds, into its own window or an
/// Outlook window that took the foreground from it, never into any other program.
/// </summary>
internal sealed class Q149ForegroundHolder : IDisposable
{
    private const uint WsOverlapped = 0x00000000;
    private const uint WsCaption = 0x00C00000;
    private const uint WsVisible = 0x10000000;
    private const uint WmClose = 0x0010;
    private const uint WmDestroy = 0x0002;
    private const uint WmChar = 0x0102;
    private const uint InputKeyboard = 1;
    private const uint KeyEventKeyUp = 0x0002;
    private const uint KeyEventUnicode = 0x0004;

    private readonly Thread _thread;
    private readonly ManualResetEventSlim _ready = new(false);
    private readonly string _className = "Q149ForegroundHolder" + Guid.NewGuid().ToString("N");
    private readonly Stopwatch _since = Stopwatch.StartNew();
    private readonly object _gate = new();
    private readonly List<(long Ms, bool Holder, string What)> _foreground = new();
    private readonly List<(long Ms, bool Holder, string What)> _sampled = new();
    private Thread? _sampler;
    private volatile bool _closing;
    private WndProc? _proc;
    private Q149Native.WinEventProc? _onForeground;
    private IntPtr _hwnd;
    private uint _threadId;
    private Thread? _typist;
    private volatile bool _typing;
    private int _received;
    private int _sent;
    private int _skipped;
    private int _failed;
    private int _sentWhileOutlookForeground;
    private int _sentWhileOutlookFocus;
    private string? _skippedTo;

    public Q149ForegroundHolder()
    {
        _thread = new Thread(Loop) { IsBackground = true, Name = "Q149.ForegroundHolder" };
        _thread.SetApartmentState(ApartmentState.STA);
        _thread.Start();
        _ready.Wait(TimeSpan.FromSeconds(10));
        if (_hwnd != IntPtr.Zero)
        {
            _sampler = new Thread(SampleForeground) { IsBackground = true, Name = "Q149.ForegroundSampler" };
            _sampler.Start();
        }
    }

    private delegate IntPtr WndProc(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);

    /// <summary>
    /// The character the typist sends: the currency sign, which no token, old body, signature or quoted original of the
    /// cells holds - so one found in a stored draft came from the typist.
    /// </summary>
    public const char TypedKey = '\u00A4';

    /// <summary>True when the holder's window exists.</summary>
    public bool Created => _hwnd != IntPtr.Zero;

    /// <summary>True when the holder's window is the foreground window now.</summary>
    public bool HasForeground => _hwnd != IntPtr.Zero && TestWin32.GetForegroundWindow() == _hwnd;

    /// <summary>Where the foreground timelines stand now, in their milliseconds - pass it to <see cref="AwaySince"/> after a draft.</summary>
    public long Mark()
    {
        return _since.ElapsedMilliseconds;
    }

    /// <summary>
    /// The foreground since <paramref name="mark"/>, two ways. As Windows announced it: every EVENT_SYSTEM_FOREGROUND of
    /// the desktop, for every process, and how long in all from one that left the holder to the next that came back to
    /// it (to now, when none came back). And as Windows answered when asked: GetForegroundWindow every 10 ms, its
    /// changes, and how long in all it named another window than the holder's - the time a person's keystrokes would
    /// have gone elsewhere. Process names, window classes and places only.
    /// </summary>
    public (long AwayMs, string[] Changes, long SampledAwayMs, string[] Samples) AwaySince(long mark)
    {
        lock (_gate)
        {
            (long announced, string[] changes) = Away(_foreground, mark, startsOnHolder: true);
            bool sampledStart = true;
            foreach ((long ms, bool holder, _) in _sampled)
            {
                if (ms < mark)
                {
                    sampledStart = holder;
                }
            }

            (long sampled, string[] samples) = Away(_sampled, mark, sampledStart);
            return (announced, changes, sampled, samples);
        }
    }

    private (long AwayMs, string[] Changes) Away(List<(long Ms, bool Holder, string What)> timeline, long mark, bool startsOnHolder)
    {
        long away = 0;
        long? leftAt = startsOnHolder ? null : mark;
        List<string> changes = new();
        foreach ((long ms, bool holder, string what) in timeline)
        {
            if (ms < mark)
            {
                continue;
            }

            changes.Add("+" + ms.ToString(CultureInfo.InvariantCulture) + " " + what);
            if (!holder && leftAt == null)
            {
                leftAt = ms;
            }
            else if (holder && leftAt != null)
            {
                away += ms - leftAt.Value;
                leftAt = null;
            }
        }

        if (leftAt != null)
        {
            away += _since.ElapsedMilliseconds - leftAt.Value;
        }

        return (away, changes.Take(40).ToArray());
    }

    /// <summary>
    /// Brings the holder to the foreground the way a program a person clicked would have it - attached to the input of
    /// whatever has the foreground, so Windows' foreground lock lets it - and says whether it has it now.
    /// </summary>
    public bool TakeForeground()
    {
        if (_hwnd == IntPtr.Zero)
        {
            return false;
        }

        IntPtr current = TestWin32.GetForegroundWindow();
        if (current == _hwnd)
        {
            return true;
        }

        uint mine = Q149Native.GetCurrentThreadId();
        uint theirs = current == IntPtr.Zero ? 0 : NativeWindows.GetWindowThreadProcessId(current, out _);
        bool attached = theirs != 0 && theirs != mine && AttachThreadInput(mine, theirs, true);
        try
        {
            _ = BringWindowToTop(_hwnd);
            _ = TestWin32.SetForegroundWindow(_hwnd);
        }
        finally
        {
            if (attached)
            {
                _ = AttachThreadInput(mine, theirs, false);
            }
        }

        PumpedStaRunner.PumpedWait(100);
        return HasForeground;
    }

    /// <summary>
    /// Starts typing as a person typing in the holder would: one <see cref="TypedKey"/> through SendInput every
    /// <paramref name="intervalMs"/> ms, which also keeps Windows' foreground lock engaged as a person's input does - an
    /// idle desktop lets any program take the foreground once the lock's time-out has passed. A key goes out only while
    /// the holder or an OUTLOOK.EXE window has the foreground, so none reaches another program on the guest.
    /// <see cref="StopTyping"/> ends it and says where the keys went.
    /// </summary>
    public void StartTyping(int intervalMs)
    {
        _ = StopTyping();
        Interlocked.Exchange(ref _received, 0);
        _sent = 0;
        _skipped = 0;
        _failed = 0;
        _sentWhileOutlookForeground = 0;
        _sentWhileOutlookFocus = 0;
        _skippedTo = null;
        _typing = true;
        _typist = new Thread(() => Type(intervalMs)) { IsBackground = true, Name = "Q149.Typist" };
        _typist.Start();
    }

    /// <summary>
    /// Ends the typing <see cref="StartTyping"/> began: the keys sent, how many reached the holder's own window (the
    /// rest went to whatever had the foreground when they arrived), how many were held back because some other program
    /// had the foreground and what it was, and how many SendInput refused. Null when no typing ran.
    /// </summary>
    public (int Sent, int ToHolder, int Skipped, int Failed, string? SkippedTo)? StopTyping()
    {
        return StopTypingCounted()?.Keys;
    }

    /// <summary>
    /// <see cref="StopTyping"/>, and of the keys sent, how many went out while GetForegroundWindow named an Outlook window,
    /// and while the foreground thread's keyboard focus (GetGUIThreadInfo) was in one - where Windows would deliver them.
    /// </summary>
    public ((int Sent, int ToHolder, int Skipped, int Failed, string? SkippedTo) Keys, int WhileOutlookForeground, int WhileOutlookFocus)? StopTypingCounted()
    {
        if (_typist == null)
        {
            return null;
        }

        _typing = false;
        _typist.Join(TimeSpan.FromSeconds(5));
        _typist = null;

        // The last keys reach the holder's message loop a moment after SendInput returns.
        Thread.Sleep(200);
        return ((_sent, Volatile.Read(ref _received), _skipped, _failed, _skippedTo), _sentWhileOutlookForeground, _sentWhileOutlookFocus);
    }

    public void Dispose()
    {
        _ = StopTyping();
        _closing = true;
        _sampler?.Join(TimeSpan.FromSeconds(2));
        if (_hwnd != IntPtr.Zero)
        {
            _ = TestWin32.PostMessageW(_hwnd, WmClose, IntPtr.Zero, IntPtr.Zero);
        }
        else if (_threadId != 0)
        {
            _ = Q149Native.PostThreadMessage(_threadId, Q149Native.WmQuit, IntPtr.Zero, IntPtr.Zero);
        }

        _thread.Join(TimeSpan.FromSeconds(5));
    }

    private void Loop()
    {
        _threadId = Q149Native.GetCurrentThreadId();
        _proc = Proc;
        WNDCLASSEX windowClass = new()
        {
            Size = (uint)Marshal.SizeOf<WNDCLASSEX>(),
            WndProc = Marshal.GetFunctionPointerForDelegate(_proc),
            Instance = GetModuleHandle(null),
            ClassName = _className,
        };
        if (RegisterClassEx(ref windowClass) != 0)
        {
            _hwnd = CreateWindowEx(0, _className, "Q149 foreground holder", WsOverlapped | WsCaption | WsVisible, 40, 40, 320, 120, IntPtr.Zero, IntPtr.Zero, windowClass.Instance, IntPtr.Zero);
        }

        // Every foreground change on the desktop, whichever process it goes to - out of context, on this thread's loop.
        _onForeground = OnForeground;
        IntPtr hook = _hwnd == IntPtr.Zero
            ? IntPtr.Zero
            : Q149Native.SetWinEventHook(Q149Native.EventSystemForeground, Q149Native.EventSystemForeground, IntPtr.Zero, _onForeground, 0, 0, Q149Native.WinEventOutOfContext);
        _ready.Set();
        if (_hwnd == IntPtr.Zero)
        {
            return;
        }

        while (Q149Native.GetMessage(out Q149Native.MSG msg, IntPtr.Zero, 0, 0) > 0)
        {
            Q149Native.TranslateMessage(ref msg);
            Q149Native.DispatchMessage(ref msg);
        }

        if (hook != IntPtr.Zero)
        {
            _ = Q149Native.UnhookWinEvent(hook);
        }
    }

    private void OnForeground(IntPtr hook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint thread, uint time)
    {
        if (hwnd == IntPtr.Zero)
        {
            return;
        }

        long at = _since.ElapsedMilliseconds;
        string what = Q149Native.Describe(hwnd);
        lock (_gate)
        {
            if (_foreground.Count < 4000)
            {
                _foreground.Add((at, hwnd == _hwnd, what));
            }
        }
    }

    private void Type(int intervalMs)
    {
        INPUT[] key =
        {
            new() { Type = InputKeyboard, Data = new InputUnion { Keyboard = new KEYBDINPUT { Scan = TypedKey, Flags = KeyEventUnicode } } },
            new() { Type = InputKeyboard, Data = new InputUnion { Keyboard = new KEYBDINPUT { Scan = TypedKey, Flags = KeyEventUnicode | KeyEventKeyUp } } },
        };
        while (_typing)
        {
            IntPtr foreground = TestWin32.GetForegroundWindow();
            if (foreground == IntPtr.Zero || (foreground != _hwnd && Q149Native.OutlookClassOf(foreground) == null))
            {
                _skipped++;
                _skippedTo ??= foreground == IntPtr.Zero ? "none" : Q149Native.Describe(foreground);
            }
            else if (SendInput((uint)key.Length, key, Marshal.SizeOf<INPUT>()) == key.Length)
            {
                _sent++;
                _sentWhileOutlookForeground += foreground != _hwnd ? 1 : 0;
                _sentWhileOutlookFocus += FocusIsOutlooks() ? 1 : 0;
            }
            else
            {
                _failed++;
            }

            Thread.Sleep(intervalMs);
        }
    }

    /// <summary>Whether the foreground thread's keyboard focus is in a window of an OUTLOOK.EXE - where a key would go.</summary>
    private static bool FocusIsOutlooks()
    {
        GUITHREADINFO info = new() { Size = (uint)Marshal.SizeOf<GUITHREADINFO>() };
        return GetGUIThreadInfo(0, ref info) && info.Focus != IntPtr.Zero && Q149Native.OutlookClassOf(info.Focus) != null;
    }

    /// <summary>GetForegroundWindow every 10 ms, its changes kept - what Windows answers, beside what it announces.</summary>
    private void SampleForeground()
    {
        IntPtr last = new(-1);
        while (!_closing)
        {
            IntPtr now = TestWin32.GetForegroundWindow();
            if (now != last)
            {
                last = now;
                long at = _since.ElapsedMilliseconds;
                string what = now == IntPtr.Zero ? "none" : Q149Native.Describe(now);
                lock (_gate)
                {
                    if (_sampled.Count < 4000)
                    {
                        _sampled.Add((at, now == _hwnd, what));
                    }
                }
            }

            Thread.Sleep(10);
        }
    }

    private IntPtr Proc(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam)
    {
        if (message == WmChar && wParam == (IntPtr)TypedKey)
        {
            Interlocked.Increment(ref _received);
            return IntPtr.Zero;
        }

        if (message == WmClose)
        {
            _ = DestroyWindow(hwnd);
            return IntPtr.Zero;
        }

        if (message == WmDestroy)
        {
            PostQuitMessage(0);
            return IntPtr.Zero;
        }

        return DefWindowProc(hwnd, message, wParam, lParam);
    }

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    private struct WNDCLASSEX
    {
        public uint Size;
        public uint Style;
        public IntPtr WndProc;
        public int ClassExtra;
        public int WindowExtra;
        public IntPtr Instance;
        public IntPtr Icon;
        public IntPtr Cursor;
        public IntPtr Background;
        public string? MenuName;
        public string ClassName;
        public IntPtr SmallIcon;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct INPUT
    {
        public uint Type;
        public InputUnion Data;
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct InputUnion
    {
        [FieldOffset(0)]
        public MOUSEINPUT Mouse;

        [FieldOffset(0)]
        public KEYBDINPUT Keyboard;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct MOUSEINPUT
    {
        public int Dx;
        public int Dy;
        public uint MouseData;
        public uint Flags;
        public uint Time;
        public IntPtr ExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct KEYBDINPUT
    {
        public ushort Vk;
        public ushort Scan;
        public uint Flags;
        public uint Time;
        public IntPtr ExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct GUITHREADINFO
    {
        public uint Size;
        public uint Flags;
        public IntPtr Active;
        public IntPtr Focus;
        public IntPtr Capture;
        public IntPtr MenuOwner;
        public IntPtr MoveSize;
        public IntPtr Caret;
        public int CaretLeft;
        public int CaretTop;
        public int CaretRight;
        public int CaretBottom;
    }

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint SendInput(uint count, INPUT[] inputs, int size);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetGUIThreadInfo(uint thread, ref GUITHREADINFO info);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern ushort RegisterClassEx(ref WNDCLASSEX windowClass);

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr CreateWindowEx(uint exStyle, string className, string windowName, uint style, int x, int y, int width, int height, IntPtr parent, IntPtr menu, IntPtr instance, IntPtr param);

    [DllImport("user32.dll")]
    private static extern IntPtr DefWindowProc(IntPtr hwnd, uint message, IntPtr wParam, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool DestroyWindow(IntPtr hwnd);

    [DllImport("user32.dll")]
    private static extern void PostQuitMessage(int exitCode);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool BringWindowToTop(IntPtr hwnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AttachThreadInput(uint attach, uint attachTo, [MarshalAs(UnmanagedType.Bool)] bool doAttach);

    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
    private static extern IntPtr GetModuleHandle(string? moduleName);
}

/// <summary>Release, and the failure description every attempt line uses.</summary>
internal static class Q149Com
{
    internal static void Release(object? comObject)
    {
        if (comObject != null && Marshal.IsComObject(comObject))
        {
            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch (Exception ex) when (ex is InvalidComObjectException or ArgumentException)
            {
            }
        }
    }

    /// <summary>Type, HRESULT and Outlook's own message - generic texts, never item content.</summary>
    internal static Dictionary<string, object> Failure(Exception ex)
    {
        Exception inner = ex is System.Reflection.TargetInvocationException { InnerException: not null } tie ? tie.InnerException! : ex;
        return new Dictionary<string, object>(StringComparer.Ordinal)
        {
            ["type"] = inner.GetType().Name,
            ["hresult"] = "0x" + inner.HResult.ToString("X8", CultureInfo.InvariantCulture),
            ["message"] = inner.Message.Length > 160 ? inner.Message.Substring(0, 160) : inner.Message,
        };
    }
}

/// <summary>
/// One COM client on its own pumped STA - the measurement's equivalent of an agent session, with nothing of
/// the product's around it: no pin, no retry, no fallback. It hears Outlook's Quit event and says when.
/// </summary>
internal sealed class Q149OutlookClient : IDisposable
{
    private readonly PumpedStaRunner _runner;
    private readonly Stopwatch _clock;
    private object? _app;
    private object? _ns;
    private OutlookQuitSinkRegistration? _registration;
    private long _quitAtMs = -1;
    private bool _disposed;

    private Q149OutlookClient(Stopwatch clock)
    {
        _clock = clock;
        _runner = new PumpedStaRunner("Q149.Probe.Sta");
    }

    /// <summary>When Outlook raised Quit, on the trial's clock; null if it has not.</summary>
    public long? QuitEventAtMs => Interlocked.Read(ref _quitAtMs) < 0 ? null : Interlocked.Read(ref _quitAtMs);

    public bool QuitSinkActive { get; private set; }

    /// <summary>
    /// Creates <c>Outlook.Application</c> - which starts OUTLOOK.EXE with -Embedding when none runs - and
    /// logs on to the default profile. <paramref name="allowStart"/> false refuses to start one.
    /// </summary>
    public static Q149OutlookClient Connect(Stopwatch clock, bool allowStart)
    {
        if (!allowStart && LiveOutlookRestart.OutlookProcessCount() == 0)
        {
            throw new InvalidOperationException("Q149: Outlook is not running and this step may not start it.");
        }

        Q149OutlookClient client = new(clock);
        try
        {
            LiveComposeProbe.Bounded<object?>("connect", LiveComposeProbe.StepBudget, () =>
            {
                client._runner.Run(() =>
                {
                    Type type = Type.GetTypeFromProgID("Outlook.Application")
                        ?? throw new InvalidOperationException("Outlook.Application ProgID is not registered.");
                    client._app = Activator.CreateInstance(type) ?? throw new InvalidOperationException("Outlook.Application did not create.");
                    dynamic ns = ((dynamic)client._app).GetNamespace("MAPI");
                    client._ns = (object)ns;
                    try
                    {
                        ns.Logon(Type.Missing, Type.Missing, false, false);
                    }
                    catch (COMException)
                    {
                    }

                    try
                    {
                        OutlookQuitSink sink = new(() => Interlocked.CompareExchange(ref client._quitAtMs, client._clock.ElapsedMilliseconds, -1));
                        client._registration = OutlookQuitSink.TryAdvise(client._app, sink);
                        client.QuitSinkActive = client._registration != null;
                    }
                    catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                    {
                        client.QuitSinkActive = false;
                    }
                });
                return null;
            });
            return client;
        }
        catch
        {
            client.Dispose();
            throw;
        }
    }

    /// <summary>Runs <paramref name="work"/> on this client's STA with its Application and NameSpace, bounded.</summary>
    public T Run<T>(Func<dynamic, dynamic, T> work, TimeSpan? budget = null)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        return LiveComposeProbe.Bounded("probe step", budget ?? LiveComposeProbe.StepBudget, () => _runner.Run(() => work((dynamic)_app!, (dynamic)_ns!)));
    }

    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _disposed = true;
        try
        {
            LiveComposeProbe.Bounded<object?>("release the probe client", TimeSpan.FromSeconds(60), () =>
            {
                _runner.Run(() =>
                {
                    _registration?.Unadvise();
                    _registration = null;
                    Q149Com.Release(_ns);
                    Q149Com.Release(_app);
                    _ns = null;
                    _app = null;
                });
                return null;
            });
        }
        catch (Exception ex) when (ex is TimeoutException or ObjectDisposedException or InvalidOperationException)
        {
        }

        _runner.Dispose();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }
}

/// <summary>Watches one OUTLOOK.EXE: its start, its command line, and the exact moment it exits and with what code.</summary>
internal sealed class Q149ProcessWatch : IDisposable
{
    private readonly Process _process;
    private long _exitedAtMs = -1;

    private Q149ProcessWatch(Process process, Stopwatch clock)
    {
        _process = process;
        Pid = process.Id;
        try
        {
            StartTimeUtc = process.StartTime.ToUniversalTime();
        }
        catch (Exception ex) when (ex is Win32Exception or InvalidOperationException or NotSupportedException)
        {
            StartTimeUtc = DateTime.UtcNow;
        }

        CommandLine = Q149Native.TryReadCommandLine(process.Id);
        process.EnableRaisingEvents = true;
        process.Exited += (_, _) => Interlocked.CompareExchange(ref _exitedAtMs, clock.ElapsedMilliseconds, -1);
        if (process.HasExited)
        {
            Interlocked.CompareExchange(ref _exitedAtMs, clock.ElapsedMilliseconds, -1);
        }
    }

    public int Pid { get; }

    public DateTime StartTimeUtc { get; }

    public string? CommandLine { get; }

    /// <summary>True when OUTLOOK.EXE was started by COM (an automation start), which is what -Embedding says.</summary>
    public bool Embedding => CommandLine?.Contains("-Embedding", StringComparison.OrdinalIgnoreCase) == true
        || CommandLine?.Contains("/Embedding", StringComparison.OrdinalIgnoreCase) == true;

    public long? ExitedAtMs => Interlocked.Read(ref _exitedAtMs) < 0 ? null : Interlocked.Read(ref _exitedAtMs);

    public bool HasExited
    {
        get
        {
            try
            {
                return _process.HasExited;
            }
            catch (Exception ex) when (ex is InvalidOperationException or Win32Exception)
            {
                return true;
            }
        }
    }

    public int? ExitCode
    {
        get
        {
            try
            {
                return _process.HasExited ? _process.ExitCode : null;
            }
            catch (Exception ex) when (ex is InvalidOperationException or Win32Exception)
            {
                return null;
            }
        }
    }

    public double MsSinceStart => (DateTime.UtcNow - StartTimeUtc).TotalMilliseconds;

    /// <summary>Attaches to the one OUTLOOK.EXE that runs; null when there is not exactly one.</summary>
    public static Q149ProcessWatch? AttachToTheOnlyOutlook(Stopwatch clock)
    {
        Process[] processes = Process.GetProcessesByName("OUTLOOK");
        if (processes.Length != 1)
        {
            foreach (Process p in processes)
            {
                p.Dispose();
            }

            return null;
        }

        return new Q149ProcessWatch(processes[0], clock);
    }

    public bool WaitForExit(int milliseconds)
    {
        try
        {
            return _process.WaitForExit(milliseconds);
        }
        catch (Exception ex) when (ex is InvalidOperationException or Win32Exception)
        {
            return true;
        }
    }

    public void Dispose()
    {
        _process.Dispose();
    }
}

/// <summary>
/// Every top-level window of one OUTLOOK.EXE that Windows shows while it watches (EVENT_OBJECT_SHOW, out of
/// context, on its own message loop), whether each was where a person could see it, and every time the
/// foreground moved to one of them. A flash of one frame is a show like any other, which is what a poll
/// cannot promise.
/// </summary>
internal sealed class Q149WindowWatch : IDisposable
{
    private readonly uint _pid;
    private readonly object _gate = new();
    private readonly List<string> _shown = new();
    private readonly Thread _thread;
    private readonly ManualResetEventSlim _ready = new(false);
    private Q149Native.WinEventProc? _callback;
    private uint _threadId;
    private readonly List<string> _events = new();
    private readonly System.Diagnostics.Stopwatch _since = System.Diagnostics.Stopwatch.StartNew();
    private int _shows;
    private int _onScreenShows;
    private int _foregroundToOutlook;

    public Q149WindowWatch(int pid)
    {
        _pid = (uint)pid;
        _thread = new Thread(Loop) { IsBackground = true, Name = "Q149.WindowWatch" };
        _thread.Start();
        _ready.Wait(TimeSpan.FromSeconds(10));
    }

    public int Shows
    {
        get
        {
            lock (_gate)
            {
                return _shows;
            }
        }
    }

    /// <summary>Shows of a window that, at that moment, was visible, unparked and of real size - one a person could see.</summary>
    public int OnScreenShows
    {
        get
        {
            lock (_gate)
            {
                return _onScreenShows;
            }
        }
    }

    public int ForegroundToOutlook
    {
        get
        {
            lock (_gate)
            {
                return _foregroundToOutlook;
            }
        }
    }

    /// <summary>
    /// Every show, hide and move of an Outlook top-level window, in order, at most 60: milliseconds since the
    /// watch began, the event, the class, and where the window was - "+120 show rctrl_renwnd32 on 140,90 1200x800".
    /// Enough to say how long a window was where a person could see it, and nothing about what it showed.
    /// </summary>
    public IReadOnlyList<string> Events
    {
        get
        {
            lock (_gate)
            {
                return _events.ToList();
            }
        }
    }

    public IReadOnlyList<string> ShownClasses
    {
        get
        {
            lock (_gate)
            {
                return _shown.ToList();
            }
        }
    }

    private void Loop()
    {
        _threadId = Q149Native.GetCurrentThreadId();
        _callback = OnEvent;
        IntPtr show = Q149Native.SetWinEventHook(Q149Native.EventObjectShow, Q149Native.EventObjectHide, IntPtr.Zero, _callback, _pid, 0, Q149Native.WinEventOutOfContext);
        IntPtr moved = Q149Native.SetWinEventHook(Q149Native.EventObjectLocationChange, Q149Native.EventObjectLocationChange, IntPtr.Zero, _callback, _pid, 0, Q149Native.WinEventOutOfContext);
        IntPtr foreground = Q149Native.SetWinEventHook(Q149Native.EventSystemForeground, Q149Native.EventSystemForeground, IntPtr.Zero, _callback, _pid, 0, Q149Native.WinEventOutOfContext);
        _ready.Set();
        while (Q149Native.GetMessage(out Q149Native.MSG msg, IntPtr.Zero, 0, 0) > 0)
        {
            Q149Native.TranslateMessage(ref msg);
            Q149Native.DispatchMessage(ref msg);
        }

        if (show != IntPtr.Zero)
        {
            Q149Native.UnhookWinEvent(show);
        }

        if (moved != IntPtr.Zero)
        {
            Q149Native.UnhookWinEvent(moved);
        }

        if (foreground != IntPtr.Zero)
        {
            Q149Native.UnhookWinEvent(foreground);
        }
    }

    private void OnEvent(IntPtr hook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint thread, uint time)
    {
        if (hwnd == IntPtr.Zero || idObject != 0 || idChild != 0)
        {
            return;
        }

        if (Q149Native.GetAncestor(hwnd, Q149Native.GaRoot) != hwnd)
        {
            return;
        }

        lock (_gate)
        {
            string what = eventType switch
            {
                Q149Native.EventSystemForeground => "foreground",
                Q149Native.EventObjectShow => "show",
                Q149Native.EventObjectHide => "hide",
                Q149Native.EventObjectLocationChange => "move",
                _ => "other",
            };

            if (what == "move" && !NativeWindows.IsWindowVisible(hwnd))
            {
                // A hidden window moving - the product's own parking - is nothing anyone could see.
                return;
            }

            if (_events.Count < 400)
            {
                _events.Add("+" + _since.ElapsedMilliseconds.ToString(System.Globalization.CultureInfo.InvariantCulture) + " " + what + " "
                    + Q149Native.ClassOf(hwnd) + " " + Q149Native.DescribePlace(hwnd));
            }

            if (eventType == Q149Native.EventSystemForeground)
            {
                _foregroundToOutlook++;
                return;
            }

            if (eventType != Q149Native.EventObjectShow)
            {
                return;
            }

            _shows++;
            _shown.Add(Q149Native.ClassOf(hwnd));
            if (Q149Native.IsOnScreen(hwnd))
            {
                _onScreenShows++;
            }
        }
    }

    public void Dispose()
    {
        if (_threadId != 0)
        {
            _ = Q149Native.PostThreadMessage(_threadId, Q149Native.WmQuit, IntPtr.Zero, IntPtr.Zero);
        }

        _ = _thread.Join(TimeSpan.FromSeconds(5));
        _ready.Dispose();
    }
}

/// <summary>The Windows calls the measurement makes.</summary>
internal static class Q149Native
{
    internal const uint EventObjectShow = 0x8002;
    internal const uint EventObjectHide = 0x8003;
    internal const uint EventObjectLocationChange = 0x800B;
    internal const uint EventSystemForeground = 0x0003;
    internal const uint WinEventOutOfContext = 0x0000;
    internal const uint GaRoot = 2;
    internal const uint WmQuit = 0x0012;
    internal const int SwShowNoActivate = 4;
    internal delegate void WinEventProc(IntPtr hook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint thread, uint time);

    /// <summary>One top-level window of OUTLOOK.EXE.</summary>
    internal sealed record TopWindow(IntPtr Handle, string ClassName, bool Visible, bool OnScreen);

    [ComImport]
    [Guid("00000114-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleWindow
    {
        [PreserveSig]
        int GetWindow(out IntPtr phwnd);

        [PreserveSig]
        int ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool enterMode);
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct MSG
    {
        public IntPtr Hwnd;
        public uint Message;
        public IntPtr WParam;
        public IntPtr LParam;
        public uint Time;
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct UnicodeString
    {
        public ushort Length;
        public ushort MaximumLength;
        public IntPtr Buffer;
    }

    /// <summary>Every top-level window of every OUTLOOK.EXE.</summary>
    internal static IReadOnlyList<TopWindow> OutlookTopWindows()
    {
        HashSet<uint> pids = new();
        foreach (Process p in Process.GetProcessesByName("OUTLOOK"))
        {
            using (p)
            {
                pids.Add((uint)p.Id);
            }
        }

        List<TopWindow> windows = new();
        if (pids.Count == 0)
        {
            return windows;
        }

        _ = NativeWindows.EnumWindows((hwnd, lParam) =>
        {
            _ = NativeWindows.GetWindowThreadProcessId(hwnd, out uint pid);
            if (pids.Contains(pid))
            {
                windows.Add(new TopWindow(hwnd, ClassOf(hwnd), NativeWindows.IsWindowVisible(hwnd), IsOnScreen(hwnd)));
            }

            return true;
        }, IntPtr.Zero);
        return windows;
    }

    internal static string ClassOf(IntPtr hwnd)
    {
        StringBuilder name = new(64);
        _ = NativeWindows.GetClassNameW(hwnd, name, name.Capacity);
        return name.ToString();
    }

    /// <summary>
    /// Every visible top-level window of every OUTLOOK.EXE, one line each: its class and place, and for a dialog the
    /// labels of its buttons, which name it without quoting its text. What a blocked step finds Outlook showing.
    /// </summary>
    internal static string[] DescribeVisibleOutlookWindows()
    {
        return OutlookTopWindows()
            .Where(w => w.Visible)
            .Select(w => w.ClassName + " " + DescribePlace(w.Handle)
                + (IsDialogClass(w.ClassName) ? " buttons " + string.Join("|", ButtonLabels(w.Handle)) : string.Empty))
            .ToArray();
    }

    /// <summary>
    /// Asks every visible dialog of every OUTLOOK.EXE to close (WM_CLOSE: a dialog's Cancel, never its OK) and returns
    /// how many it asked - for a step blocked on a window meant for a person, on a test guest.
    /// </summary>
    internal static int CloseVisibleOutlookDialogs()
    {
        int asked = 0;
        foreach (TopWindow w in OutlookTopWindows())
        {
            if (w.Visible && IsDialogClass(w.ClassName))
            {
                WindowProbe.PostClose(w.Handle);
                asked++;
            }
        }

        return asked;
    }

    /// <summary>Windows' own dialog class, Office's (NUIDialog) and Word's (bosa_sdm_...).</summary>
    internal static bool IsDialogClass(string className)
    {
        return className == "#32770" || className == "NUIDialog" || className.StartsWith("bosa_sdm_", StringComparison.Ordinal);
    }

    private static List<string> ButtonLabels(IntPtr dialog)
    {
        const uint WmGetText = 0x000D;
        const uint SmtoAbortIfHung = 0x0002;
        List<string> labels = new();
        _ = TestWin32.EnumChildWindows(dialog, (child, _) =>
        {
            if (labels.Count < 8 && string.Equals(ClassOf(child), "Button", StringComparison.Ordinal))
            {
                StringBuilder text = new(64);
                if (SendMessageTimeout(child, WmGetText, (IntPtr)text.Capacity, text, SmtoAbortIfHung, 500, out _) != IntPtr.Zero)
                {
                    labels.Add(text.ToString().Replace("&", string.Empty, StringComparison.Ordinal));
                }
            }

            return true;
        }, IntPtr.Zero);
        return labels;
    }

    [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint message, IntPtr wParam, StringBuilder lParam, uint flags, uint timeoutMs, out IntPtr result);

    /// <summary>Visible, of real size, and not parked far off the screen - where a person could see it.</summary>
    internal static bool IsOnScreen(IntPtr hwnd)
    {
        if (!NativeWindows.IsWindowVisible(hwnd) || !NativeWindows.GetWindowRect(hwnd, out NativeWindows.RECT r))
        {
            return false;
        }

        if (r.Right <= r.Left || r.Bottom <= r.Top)
        {
            return false;
        }

        return !(r.Left <= ComposeSurface.ParkX / 2 && r.Top <= ComposeSurface.ParkY / 2);
    }

    /// <summary>"on|off|hidden left,top widthxheight" - where a window is, and whether a person could see it there.</summary>
    internal static string DescribePlace(IntPtr hwnd)
    {
        if (!NativeWindows.GetWindowRect(hwnd, out NativeWindows.RECT r))
        {
            return "unknown";
        }

        string state = !NativeWindows.IsWindowVisible(hwnd) ? "hidden" : IsOnScreen(hwnd) ? "on" : "off";
        return string.Format(System.Globalization.CultureInfo.InvariantCulture, "{0} {1},{2} {3}x{4}", state, r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top);
    }

    /// <summary>
    /// The foreground window as one line: the name of the process that owns it, its class and its place ("hidden", "on"
    /// or "off" the screen) - window classes and process names only, never a title.
    /// </summary>
    internal static string DescribeForeground()
    {
        IntPtr foreground = TestWin32.GetForegroundWindow();
        return foreground == IntPtr.Zero ? "none" : Describe(foreground);
    }

    /// <summary><paramref name="window"/> as one line, as <see cref="DescribeForeground"/> gives the foreground window.</summary>
    internal static string Describe(IntPtr window)
    {
        _ = NativeWindows.GetWindowThreadProcessId(window, out uint pid);
        string owner;
        try
        {
            using Process process = Process.GetProcessById((int)pid);
            owner = process.ProcessName;
        }
        catch (Exception ex) when (ex is ArgumentException || ex is InvalidOperationException)
        {
            // No process with that id any more, or one that has just exited.
            owner = "?";
        }

        return owner + " " + ClassOf(window) + " " + DescribePlace(window);
    }

    /// <summary>Whether the foreground window belongs to an OUTLOOK.EXE, and its class - null when there is none.</summary>
    internal static string? ForegroundOutlookClass()
    {
        IntPtr foreground = TestWin32.GetForegroundWindow();
        return foreground == IntPtr.Zero ? null : OutlookClassOf(foreground);
    }

    /// <summary>The class of <paramref name="window"/> when an OUTLOOK.EXE owns it; null when another process does.</summary>
    internal static string? OutlookClassOf(IntPtr window)
    {
        _ = NativeWindows.GetWindowThreadProcessId(window, out uint pid);
        try
        {
            using Process owner = Process.GetProcessById((int)pid);
            return string.Equals(owner.ProcessName, "OUTLOOK", StringComparison.OrdinalIgnoreCase) ? ClassOf(window) : null;
        }
        catch (Exception ex) when (ex is ArgumentException || ex is InvalidOperationException)
        {
            return null;
        }
    }

    /// <summary>
    /// How long after a person's input Windows keeps other programs from taking the foreground
    /// (SPI_GETFOREGROUNDLOCKTIMEOUT), in milliseconds; null when it cannot be read.
    /// </summary>
    internal static long? ForegroundLockTimeoutMs()
    {
        const uint GetForegroundLockTimeout = 0x2000;
        uint timeout = 0;
        return SystemParametersInfo(GetForegroundLockTimeout, 0, ref timeout, 0) ? timeout : null;
    }

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SystemParametersInfo(uint action, uint param, ref uint value, uint winIni);

    internal static void Park(IntPtr hwnd)
    {
        _ = NativeWindows.SetWindowPos(hwnd, IntPtr.Zero, ComposeSurface.ParkX, ComposeSurface.ParkY, 0, 0, NativeWindows.SWP_NOSIZE | NativeWindows.SWP_NOZORDER | NativeWindows.SWP_NOACTIVATE);
    }

    /// <summary>
    /// A process's command line, read the way Task Manager reads it (ProcessCommandLineInformation), or null.
    /// Needed to tell an automation start (-Embedding) from a user's.
    /// </summary>
    internal static string? TryReadCommandLine(int pid)
    {
        const uint QueryLimitedInformation = 0x1000;
        const int ProcessCommandLineInformation = 60;
        IntPtr process = OpenProcess(QueryLimitedInformation, false, (uint)pid);
        if (process == IntPtr.Zero)
        {
            return null;
        }

        try
        {
            _ = NtQueryInformationProcess(process, ProcessCommandLineInformation, IntPtr.Zero, 0, out int needed);
            if (needed <= 0 || needed > 1 << 20)
            {
                return null;
            }

            IntPtr buffer = Marshal.AllocHGlobal(needed);
            try
            {
                if (NtQueryInformationProcess(process, ProcessCommandLineInformation, buffer, needed, out _) != 0)
                {
                    return null;
                }

                UnicodeString text = Marshal.PtrToStructure<UnicodeString>(buffer);
                return text.Buffer == IntPtr.Zero ? null : Marshal.PtrToStringUni(text.Buffer, text.Length / 2);
            }
            finally
            {
                Marshal.FreeHGlobal(buffer);
            }
        }
        finally
        {
            _ = CloseHandle(process);
        }
    }

    [DllImport("user32.dll")]
    internal static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventProc callback, uint idProcess, uint idThread, uint flags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool UnhookWinEvent(IntPtr hook);

    [DllImport("user32.dll")]
    internal static extern int GetMessage(out MSG msg, IntPtr hwnd, uint filterMin, uint filterMax);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool TranslateMessage(ref MSG msg);

    [DllImport("user32.dll")]
    internal static extern IntPtr DispatchMessage(ref MSG msg);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool PostThreadMessage(uint threadId, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll")]
    internal static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    internal static extern IntPtr GetAncestor(IntPtr hwnd, uint flags);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr OpenProcess(uint access, [MarshalAs(UnmanagedType.Bool)] bool inherit, uint processId);

    [DllImport("kernel32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseHandle(IntPtr handle);

    [DllImport("ntdll.dll")]
    private static extern int NtQueryInformationProcess(IntPtr process, int informationClass, IntPtr information, int length, out int returnLength);
}
