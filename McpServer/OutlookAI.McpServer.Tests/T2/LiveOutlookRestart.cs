using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32;
using OutlookAI.Core.AddIn;
using OutlookAI.Core.Com;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The one way a live test ends Outlook and starts it again (stream K2, 2026-11-18). It is the fold of two copies:
/// the stop the Q149 measurement kept in <see cref="LiveComposeProbe"/> - with its <c>Application.Quit()</c> from a
/// client that holds nothing else, which was private there - and the quit protocol
/// <see cref="LiveDisconnectRecoveryTests"/> carries: the S7 counts first, a person's windows left alone, the pin
/// released only on a session that is still alive, every wait bounded, nothing killed. Every test that bounces
/// Outlook keeps AGENTS.md's mailbox-safety rule 7 through it.
/// <para>
/// <b>Rule 7, as kept here.</b> OUTLOOK.EXE is never killed, and a stop that does not end it fails instead of
/// escalating. Before anything that can end it come the two S7 counts, whatever is on screen (Q101 3(a)): an open
/// Inspector - an unsent compose would be thrown away - is a user-protection stop
/// (<see cref="LivePopulationCoverage.StandAsideForAUser"/>), and an Outbox that is not provably empty fails on every
/// profile (<see cref="OutboxRefusal"/>, Q76). Then each rung is tried only when the one before did not end Outlook,
/// and every rung is one a user's own Outlook accepts:
/// </para>
/// <list type="number">
/// <item>WM_CLOSE to its visible windows - a user's close, and only once nobody has touched the keyboard or mouse
/// for <see cref="IdleSecondsBeforeClosing"/> seconds;</item>
/// <item>closing its Explorers once none is visible: a hidden Explorer is a pin, and a pin keeps an Outlook up
/// after its last window closed (D49);</item>
/// <item>for a window-less Outlook that a client still holds - a <c>MailService</c> of the suite, which has no
/// way to let go - ONE Explorer on the hub, shown through the product's goto surface and closed like the first
/// rung: closing the last window ends Outlook while clients hold it, where <c>Application.Quit()</c> parks it
/// until they let go (both measured, <see cref="LiveDisconnectRecoveryTests"/>);</item>
/// <item><c>Application.Quit()</c> from a client that holds nothing but the Application object, released at
/// once.</item>
/// </list>
/// <para>
/// A start opens Outlook the way the guest runner does (<c>Testbed/guest/Start-OutlookUnelevated.ps1</c>):
/// OUTLOOK.EXE on the default profile, with its window, unelevated as the test itself runs - and is done once the
/// window is up and the object model answers.
/// </para>
/// </summary>
public static class LiveOutlookRestart
{
    /// <summary>How long the last rung of a stop may take before the stop fails rather than escalating.</summary>
    internal static readonly TimeSpan StopBudget = TimeSpan.FromSeconds(120);

    /// <summary>How long one rung's close may take to end Outlook before the next rung is tried.</summary>
    internal static readonly TimeSpan RungBudget = TimeSpan.FromSeconds(30);

    /// <summary>The pause after OUTLOOK.EXE has gone before the next start, so no start lands on an exiting instance.</summary>
    internal const int SettleAfterExitMs = 5000;

    /// <summary>How long one COM step of a stop may block before it fails naming the step.</summary>
    internal static readonly TimeSpan StepBudget = TimeSpan.FromMinutes(3);

    /// <summary>How long a start waits for Outlook's window unless the caller gives its own budget.</summary>
    internal static readonly TimeSpan DefaultWindowBudget = TimeSpan.FromMinutes(3);

    /// <summary>RPC_S_SERVER_UNAVAILABLE as an HRESULT: the COM server is gone or going.</summary>
    private const int RpcServerUnavailable = unchecked((int)0x800706BA);

    /// <summary>RPC_E_DISCONNECTED: the object invoked has disconnected from its clients.</summary>
    private const int RpcDisconnected = unchecked((int)0x80010108);

    /// <summary>
    /// Seconds without keyboard or mouse input before a window may be closed or shown: less means somebody may be at
    /// the machine, and their Outlook is not this class's to close.
    /// </summary>
    internal const double IdleSecondsBeforeClosing = 180;

    /// <summary>What a stop for recent input means on a machine nobody uses during a run.</summary>
    internal const string RecentInputOnAnUnattendedMachine =
        "Nothing in the suite synthesizes input, so on a test guest recent keyboard or mouse input means a person, or "
        + "a console session left connected, was using the guest during the run - which the live tier assumes never "
        + "happens. Disconnect from the guest's console and re-run.";

    /// <summary>What a stop for an open Inspector means on a machine nobody uses during a run.</summary>
    internal const string OpenInspectorOnAnUnattendedMachine =
        "Nobody composes mail on a test guest, so an open Inspector is one an earlier test or step left open - and "
        + "making Outlook exit would throw away whatever it holds. Find what opened it before re-running.";

    /// <summary>
    /// Runs <paramref name="work"/> on its own thread and gives up after <paramref name="budget"/>, naming
    /// <paramref name="what"/>. A step that fails throws its own failure, as a direct call would. Not
    /// <c>Task.Wait</c>: it hands a failed step's failure on inside an <see cref="AggregateException"/>, which no
    /// catch around a bounded step looks for - the stop's and the Q149 cells' own never matched (the Q149 smoke on
    /// OutlookAI-Indexed, 2026-11-18). The disconnect test's bounded step waits the same way.
    /// </summary>
    internal static T Bounded<T>(string what, TimeSpan budget, Func<T> work)
    {
        Task<T> task = Task.Run(work);
        if (!ReferenceEquals(Task.WhenAny(task, Task.Delay(budget)).GetAwaiter().GetResult(), task))
        {
            throw new TimeoutException(what + " did not finish within " + budget.TotalSeconds.ToString("0", CultureInfo.InvariantCulture)
                + " s - Outlook is probably wedged; nothing was killed.");
        }

        return task.GetAwaiter().GetResult();
    }

    /// <summary>The OUTLOOK.EXE processes now running.</summary>
    internal static int OutlookProcessCount()
    {
        return OutlookProcessIds().Count;
    }

    /// <summary>
    /// The ids of the OUTLOOK.EXE processes now running in this Windows session - the product's own rule for which
    /// Outlook is this user's (<see cref="OutlookProcesses"/>) - lowest first.
    /// </summary>
    internal static IReadOnlyList<int> OutlookProcessIds()
    {
        Process[] processes = OutlookProcesses.InThisSession();
        List<int> ids = new(processes.Length);
        foreach (Process p in processes)
        {
            ids.Add(p.Id);
            p.Dispose();
        }

        ids.Sort();
        return ids;
    }

    /// <summary>Waits until no OUTLOOK.EXE runs; true when that happened inside <paramref name="timeout"/>.</summary>
    internal static bool WaitForNoOutlook(TimeSpan timeout)
    {
        LiveWaitBudget wait = LiveWaitBudget.Of(timeout);
        while (wait.HasTimeLeft)
        {
            if (OutlookProcessCount() == 0)
            {
                return true;
            }

            Thread.Sleep(200);
        }

        return OutlookProcessCount() == 0;
    }

    /// <summary>
    /// Ends Outlook the S7 way, or fails naming why it could not; a no-op when none runs. True once no OUTLOOK.EXE
    /// runs. False only where it stood aside for a person at the keyboard - which
    /// <see cref="LivePopulationCoverage.StandAsideForAUser"/> allows on a Production profile alone, every other
    /// profile failing there instead - and then nothing was closed.
    /// </summary>
    /// <param name="settings">The run's settings: the profile the user-protection stops read, and the hub the third rung shows.</param>
    /// <param name="scenario">What ends Outlook, as a stop names what did not run: "the Q149 measurement, which ends and restarts Outlook".</param>
    /// <param name="log">Where each rung is reported.</param>
    internal static bool Stop(LiveTestSettings settings, string scenario, Action<string> log)
    {
        ArgumentNullException.ThrowIfNull(settings);
        ArgumentException.ThrowIfNullOrWhiteSpace(scenario);
        ArgumentNullException.ThrowIfNull(log);

        // An Outlook already on its way out is given a moment to finish, so nothing below attaches to - or, in a
        // race, starts a new Outlook on top of - an instance that is ending by itself. One that does end is settled
        // like any rung's, so the caller's next start does not land on its heels: in the Q149 smoke of 2026-11-18 a
        // start less than a second after an exit was followed by a crash in mso40uiwin32client.dll - whether it
        // caused it is not known.
        bool wasRunning = OutlookProcessCount() > 0;
        if (WaitForNoOutlook(TimeSpan.FromSeconds(5)))
        {
            if (wasRunning)
            {
                log("stop: OUTLOOK.EXE ended by itself");
                Thread.Sleep(SettleAfterExitMs);
            }

            return true;
        }

        // The S7 counts first, whatever is on screen: every rung below ends Outlook. An Outlook that ended by itself
        // between the wait above and the count leaves nothing to protect: the count found none to attach to, or the
        // one it attached to exited while the product's start guard waited for it to settle. The Q149 smoke met the
        // second on OutlookAI-Indexed on 2026-11-18: the Outlook a cell's own client had started to read its draft,
        // still ending when the next cell's stop counted.
        int inspectors;
        int outbox;
        try
        {
            (inspectors, outbox) = ReadSafetyCounts();
        }
        catch (InvalidOperationException ex)
        {
            if (!WaitForNoOutlook(RungBudget))
            {
                throw;
            }

            log("stop: Outlook ended by itself while its S7 counts were read (" + ex.Message + ")");
            Thread.Sleep(SettleAfterExitMs);
            return true;
        }

        log("stop: S7 counts read - " + inspectors.ToString(CultureInfo.InvariantCulture) + " Inspector(s), "
            + outbox.ToString(CultureInfo.InvariantCulture) + " item(s) in the Outboxes");
        if (inspectors != 0)
        {
            LivePopulationCoverage.StandAsideForAUser(
                settings,
                inspectors.ToString(CultureInfo.InvariantCulture) + " Inspector window(s) are open (possibly an unsent compose)",
                scenario,
                OpenInspectorOnAnUnattendedMachine,
                log);
            return false;
        }

        string? outboxRefusal = OutboxRefusal(outbox, scenario);
        if (outboxRefusal != null)
        {
            throw new InvalidOperationException(outboxRefusal);
        }

        // 1. A user's close of its visible windows.
        IReadOnlyList<IntPtr> visible = WindowProbe.VisibleOutlookWindows();
        if (visible.Count > 0)
        {
            if (!NobodyAtTheKeyboard(settings, "Outlook windows are open", scenario, log))
            {
                return false;
            }

            log("stop: WM_CLOSE to " + visible.Count.ToString(CultureInfo.InvariantCulture) + " visible window(s)");
            foreach (IntPtr window in visible)
            {
                WindowProbe.PostClose(window);
            }

            if (Ended(RungBudget, "WM_CLOSE to its " + visible.Count.ToString(CultureInfo.InvariantCulture) + " visible window(s)", log))
            {
                return true;
            }
        }

        if (WindowProbe.VisibleOutlookWindows().Count == 0)
        {
            // 2. Its hidden Explorers - a pin - closed once nothing is on screen.
            log("stop: closing hidden Explorers");
            int closed = CloseHiddenExplorers(log);
            if (closed > 0 && Ended(RungBudget, "its " + closed.ToString(CultureInfo.InvariantCulture) + " hidden Explorer(s) were closed", log))
            {
                return true;
            }

            // 3. A window-less Outlook a client still holds: one Explorer shown and closed like a user's.
            if (OutlookProcessCount() > 0)
            {
                if (!NobodyAtTheKeyboard(settings, "Outlook is to show a window so that it can be closed", scenario, log))
                {
                    return false;
                }

                if (ShowOneExplorer(settings, scenario, log) is IntPtr shown)
                {
                    log("stop: WM_CLOSE to the one Explorer shown on the hub");
                    WindowProbe.PostClose(shown);
                    if (Ended(RungBudget, "WM_CLOSE to the one Explorer shown on the hub", log))
                    {
                        return true;
                    }
                }
            }
        }

        // 4. Application.Quit() from a client that holds nothing else.
        log("stop: Application.Quit() from a client holding nothing else");
        QuitFromAnEmptyClient();
        if (!Ended(StopBudget, "Application.Quit()", log))
        {
            throw new InvalidOperationException(
                scenario + ": Outlook did not end after WM_CLOSE to its windows, closing its Explorers, closing one Explorer "
                + "shown for it and Application.Quit() - it was NOT killed, and nothing that needs it ended can go on.");
        }

        return true;
    }

    /// <summary>
    /// Starts Outlook the way a user does - OUTLOOK.EXE on the default profile, so with its main window - and waits until
    /// that window is up and the object model answers. Refuses while an OUTLOOK.EXE still runs: a start on top of one
    /// would only hand its window to the running instance.
    /// </summary>
    /// <param name="scenario">What starts Outlook, for the failure messages.</param>
    /// <param name="log">Where the start is reported.</param>
    /// <param name="windowBudget">How long the window may take; <see cref="DefaultWindowBudget"/> when null.</param>
    /// <param name="coldStart">Told when OUTLOOK.EXE was started, its window showed and its object model first answered (Q180 (b)).</param>
    internal static void StartWindowed(string scenario, Action<string> log, TimeSpan? windowBudget = null, Q180ColdStartWatch? coldStart = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(scenario);
        ArgumentNullException.ThrowIfNull(log);
        if (OutlookProcessCount() > 0)
        {
            throw new InvalidOperationException(scenario + ": OUTLOOK.EXE still runs, so it is not started again - stop it first (LiveOutlookRestart.Stop).");
        }

        // THROUGH THE SHELL, as Start-Process starts it in Start-OutlookUnelevated.ps1. A child the test host creates
        // itself, without the shell, inherits every inheritable handle the test host has, the suite's output pipe
        // among them, and an Outlook left running then holds that pipe open after the suite has finished: the guest
        // task never reads end-of-file and never writes exit.txt, and the run waits out its whole time limit. Measured
        // on the Exchange VM on 2026-11-18: the suite done at 09:51:32 UTC, the run idle for 50 minutes after it.
        string exe = OutlookExePath(scenario);
        ProcessStartInfo start = new(exe) { UseShellExecute = true };
        string? profile = DefaultProfileName();
        if (!string.IsNullOrWhiteSpace(profile))
        {
            // As the guest runner starts it: the profile named, so no picker can appear.
            start.Arguments = "/profile " + QuotedArgument(profile!);
        }

        coldStart?.Milestone("startRequested");
        using (Process? started = Process.Start(start))
        {
            log("start: OUTLOOK.EXE started with its window (pid " + (started?.Id.ToString(CultureInfo.InvariantCulture) ?? "?") + ")");
        }

        TimeSpan budget = windowBudget ?? DefaultWindowBudget;
        LiveWaitBudget wait = LiveWaitBudget.Of(budget);
        while (wait.HasTimeLeft && WindowProbe.VisibleOutlookWindows().Count == 0)
        {
            Thread.Sleep(coldStart == null ? 500 : 100);
        }

        string waited = budget.TotalSeconds % 60 == 0
            ? (budget.TotalSeconds / 60).ToString("0", CultureInfo.InvariantCulture) + " minutes"
            : budget.TotalSeconds.ToString("0", CultureInfo.InvariantCulture) + " s";
        if (WindowProbe.VisibleOutlookWindows().Count == 0)
        {
            coldStart?.Hang("no window within " + waited);
            throw new InvalidOperationException(scenario + ": OUTLOOK.EXE started but showed no window within " + waited + ".");
        }

        coldStart?.Milestone("windowShown");
        log("start: Outlook's window is up after " + wait.Elapsed.TotalSeconds.ToString("0", CultureInfo.InvariantCulture) + " s");

        // The window is up before the stores are; one store read through a short-lived client proves the object model answers.
        _ = ReadSafetyCounts();
        coldStart?.Ready();
        Thread.Sleep(3000);
    }

    /// <summary>
    /// Leaves Outlook running with its window: as it is when it already shows one, else ended (<see cref="Stop"/>) and
    /// started with its window (<see cref="StartWindowed"/>), so the windowed state is a normally started Outlook and not
    /// a promoted one. False where the stop stood aside for a person at the keyboard, and nothing was ended.
    /// </summary>
    internal static bool EnsureWindowed(
        LiveTestSettings settings, string scenario, Action<string> log, TimeSpan? windowBudget = null, Q180ColdStartWatch? coldStart = null)
    {
        if (OutlookProcessCount() > 0 && WindowProbe.VisibleOutlookWindows().Count > 0)
        {
            coldStart?.Milestone("alreadyWindowed");
            return true;
        }

        if (!Stop(settings, scenario, log))
        {
            return false;
        }

        StartWindowed(scenario, log, windowBudget, coldStart);
        return true;
    }

    /// <summary>
    /// Ends Outlook (<see cref="Stop"/>) and starts it again with its window (<see cref="StartWindowed"/>): a new
    /// OUTLOOK.EXE, so a new Outlook session, which is what a person's restart is. Null where the stop stood aside for a
    /// person at the keyboard, and nothing was ended.
    /// </summary>
    internal static OutlookRestart? Restart(LiveTestSettings settings, string scenario, Action<string> log, TimeSpan? windowBudget = null)
    {
        IReadOnlyList<int> before = OutlookProcessIds();
        if (!Stop(settings, scenario, log))
        {
            return null;
        }

        StartWindowed(scenario, log, windowBudget);
        OutlookRestart restart = new(before, OutlookProcessIds());
        log("restart: " + restart.Describe());
        return restart;
    }

    /// <summary>The two S7 counts through a short-lived session that makes no pin: open Inspectors, items in every Outbox.</summary>
    internal static (int Inspectors, int Outbox) ReadSafetyCounts()
    {
        return Bounded("the S7 counts", StepBudget, () =>
        {
            using OutlookComSession session = OutlookComSession.Connect(allowStartingOutlook: false, onOutlookGone: null, lifetimePin: false);
            return (session.GetOpenInspectors().Count, session.CountOutboxItems());
        });
    }

    /// <summary>
    /// What a test says when the S7 Outbox count is not a clean zero, or null when it is and Outlook may be ended.
    /// <para>
    /// <b>A FAILURE on every machine profile - not a skip, and not a Portable announcement (Q76, 2026-09-27).</b>
    /// <see cref="LiveDisconnectRecoveryTests"/> used to print <c>SKIP:</c> here and return GREEN, having proved nothing.
    /// The usual answer to "this test could not prove anything here" is the Production/Portable split of
    /// <see cref="LivePopulationCoverage"/>, and it is deliberately NOT used, because its premise does not hold for an
    /// Outbox. Its Portable half announces instead of failing because a test machine legitimately LACKS some
    /// population and names it under <c>Requires</c>; an empty Outbox is not something any machine lacks - every
    /// machine can have one, and the suite demands it. On a test guest nothing but a test ever queues mail, so a count
    /// above zero there is residue the zero-artifact rule forbids (a send the local sink never delivered, an earlier
    /// run's leftovers, a corpus build's strays), and <see cref="LiveMailSink.EnsureOutboxDrained"/> already refuses a
    /// whole run for it. Announcing it would turn a stuck Outbox into a permanent <c>PROVED NOTHING:</c>.
    /// </para>
    /// <para>
    /// An unreadable count (-1) fails the same way: unknown is unsafe under S7. A count, never a subject (S4) - on a
    /// working profile the Outbox is the user's own unsent mail.
    /// </para>
    /// </summary>
    /// <param name="outboxItems">Items in every Outbox of the profile; -1 when the count could not be read.</param>
    /// <param name="provedNothingAbout">What the stop leaves unproved: "disconnect recovery".</param>
    internal static string? OutboxRefusal(int outboxItems, string provedNothingAbout)
    {
        if (outboxItems == 0)
        {
            return null;
        }

        string found = outboxItems < 0
            ? "The Outbox item count could not be read, so an empty Outbox cannot be told from a full one"
            : "The profile's Outbox holds " + outboxItems.ToString(CultureInfo.InvariantCulture) + " item(s)";
        return found + ". This test closes Outlook, and the quit-when-safe rule (S7) forbids that while anything may "
            + "be queued, so it closed nothing and proved nothing about " + provedNothingAbout + ". That fails on every machine "
            + "profile instead of skipping (Q76). On a test guest nothing but a test ever queues mail, so this is residue "
            + "the zero-artifact rule forbids - a send the local sink never delivered, an earlier run's leftovers or a "
            + "corpus build's strays; LiveMailSink.EnsureOutboxDrained refuses a whole run for the same state. On a working "
            + "profile it may be your own unsent mail: send or clear it, then re-run.";
    }

    /// <summary>
    /// One command-line argument for the shell's start: in double quotes, so a name with a space stays one argument.
    /// A name holding a double quote cannot be passed that way, and is refused rather than split. Pure.
    /// </summary>
    internal static string QuotedArgument(string value)
    {
        ArgumentNullException.ThrowIfNull(value);
        if (value.Contains('"', StringComparison.Ordinal))
        {
            throw new InvalidOperationException("A profile name holding a double quote cannot be passed to OUTLOOK.EXE on its command line.");
        }

        return "\"" + value + "\"";
    }

    /// <summary>The profile Outlook opens by default (DefaultProfile under HKCU Software\Microsoft\Office\16.0\Outlook), or null.</summary>
    internal static string? DefaultProfileName()
    {
        using RegistryKey? key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Outlook");
        return key?.GetValue("DefaultProfile") as string;
    }

    /// <summary>OUTLOOK.EXE's path from App Paths, else Office's default location.</summary>
    internal static string OutlookExePath(string scenario)
    {
        foreach (RegistryKey root in new[] { Registry.CurrentUser, Registry.LocalMachine })
        {
            using RegistryKey? key = root.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE");
            if (key?.GetValue(null) is string path && File.Exists(path))
            {
                return path;
            }
        }

        const string Default = @"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE";
        return File.Exists(Default)
            ? Default
            : throw new InvalidOperationException(scenario + ": OUTLOOK.EXE was not found under App Paths or at its default location.");
    }

    /// <summary>
    /// Whether a window may be closed or shown now: true once nobody has touched the keyboard or mouse for
    /// <see cref="IdleSecondsBeforeClosing"/> seconds. Otherwise the user-protection stop, which fails on every profile
    /// but Production and there says so in one line and leaves Outlook alone (false).
    /// </summary>
    private static bool NobodyAtTheKeyboard(LiveTestSettings settings, string observed, string scenario, Action<string> log)
    {
        double idle = WindowProbe.UserIdleSeconds();
        if (idle >= IdleSecondsBeforeClosing)
        {
            return true;
        }

        LivePopulationCoverage.StandAsideForAUser(
            settings,
            observed + " and the user was active " + idle.ToString("0", CultureInfo.InvariantCulture) + " s ago",
            scenario,
            RecentInputOnAnUnattendedMachine,
            log);
        return false;
    }

    /// <summary>Waits for OUTLOOK.EXE to end after a rung, and settles when it did.</summary>
    private static bool Ended(TimeSpan budget, string after, Action<string> log)
    {
        if (!WaitForNoOutlook(budget))
        {
            return false;
        }

        log("stop: OUTLOOK.EXE ended after " + after);
        Thread.Sleep(SettleAfterExitMs);
        return true;
    }

    /// <summary>
    /// Closes every Explorer of an Outlook that shows none - the product's own <c>TryCloseInvisibleExplorers</c>, which
    /// refuses while any window is visible - through a short-lived session that makes no pin of its own. A session that
    /// dies under the call is Outlook already going: on the Exchange VM (run 8, 2026-11-18) a quit took the session with
    /// it as an <c>ObjectDisposedException</c>, not as an RPC error.
    /// </summary>
    private static int CloseHiddenExplorers(Action<string> log)
    {
        try
        {
            return Bounded("closing the hidden Explorers", StepBudget, () =>
            {
                using OutlookComSession session = OutlookComSession.Connect(allowStartingOutlook: false, onOutlookGone: null, lifetimePin: false);
                return session.TryCloseInvisibleExplorers();
            });
        }
        catch (Exception ex) when (IsOutlookGoing(ex))
        {
            log("stop: the session died under the Explorer close (" + DescribeGone(ex) + ") - Outlook was already going");
            return 0;
        }
    }

    /// <summary>
    /// Shows ONE Explorer on the hub's root through the product's goto surface - the sanctioned window maker - from a
    /// short-lived session that is released before the window is closed, and returns the window, or null when none
    /// could be shown.
    /// </summary>
    private static IntPtr? ShowOneExplorer(LiveTestSettings settings, string scenario, Action<string> log)
    {
        IReadOnlyList<IntPtr> before = WindowProbe.VisibleOutlookWindows();
        string? refused;
        try
        {
            refused = Bounded(scenario + ": showing one Explorer on the hub", StepBudget, () =>
            {
                using OutlookComSession session = OutlookComSession.Connect(allowStartingOutlook: false, onOutlookGone: null, lifetimePin: false);
                ComStoreDetail? hub = session.GetStoreDetails().FirstOrDefault(
                    d => string.Equals(d.DisplayName, settings.TestHubStoreDisplayName, StringComparison.OrdinalIgnoreCase));
                if (hub == null)
                {
                    return "the hub is not a store of the running profile";
                }

                ComExplorerState? state = session.TryGotoFolder(null, hub.StoreId, out string? error);
                return state == null ? (error ?? "the goto surface showed no Explorer") : null;
            });
        }
        catch (Exception ex) when (IsOutlookGoing(ex))
        {
            log("stop: the session died while showing an Explorer (" + DescribeGone(ex) + ") - Outlook was already going");
            return null;
        }

        if (refused != null)
        {
            log("stop: no Explorer could be shown on the hub - " + refused);
            return null;
        }

        LiveWaitBudget wait = LiveWaitBudget.OfSeconds(15);
        while (wait.HasTimeLeft)
        {
            IntPtr shown = WindowProbe.VisibleOutlookWindows().FirstOrDefault(h => !before.Contains(h));
            if (shown != IntPtr.Zero)
            {
                return shown;
            }

            Thread.Sleep(250);
        }

        log("stop: the Explorer shown on the hub did not become visible within 15 s");
        return null;
    }

    /// <summary>
    /// A call whose session's Outlook quit under it, in every shape <see cref="ComGateway"/> counts as that
    /// disconnect: a COM disconnect (RPC_S_SERVER_UNAVAILABLE, RPC_E_DISCONNECTED), or a session the exit watcher
    /// already disposed or separated - the <c>ObjectDisposedException</c> the disconnect test's pin release met on the
    /// Exchange VM (runbook section 4.4a, run 8). The one copy: the stop's rungs and that test's pin release use it.
    /// </summary>
    internal static bool IsSessionGone(Exception ex)
    {
        return ex is ObjectDisposedException
            || ex is InvalidComObjectException
            || (ex is COMException com && (com.HResult == RpcServerUnavailable || com.HResult == RpcDisconnected));
    }

    /// <summary>A session-gone failure for the log: the HRESULT of a COM disconnect, else the exception's type.</summary>
    internal static string DescribeGone(Exception ex)
    {
        return ex is COMException com ? "0x" + com.HResult.ToString("X8", CultureInfo.InvariantCulture) : ex.GetType().Name;
    }

    /// <summary>A rung's call that failed because Outlook is going away under it - or is already gone.</summary>
    private static bool IsOutlookGoing(Exception ex)
    {
        return IsSessionGone(ex) || (ex is InvalidOperationException && OutlookProcessCount() == 0);
    }

    private static void QuitFromAnEmptyClient()
    {
        // Rule 7: the client calling Quit holds nothing but the Application object, and releases it at once.
        Bounded<object?>("Application.Quit", StepBudget, () =>
        {
            Thread sta = new Thread(() =>
            {
                object? app = null;
                try
                {
                    Type type = Type.GetTypeFromProgID("Outlook.Application")!;
                    if (OutlookProcessCount() == 0)
                    {
                        return;
                    }

                    app = Activator.CreateInstance(type);
                    ((dynamic)app!).Quit();
                }
                catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                {
                }
                finally
                {
                    Q149Com.Release(app);
                }
            })
            {
                IsBackground = true,
                Name = "LiveOutlookRestart.Quit",
            };
            sta.SetApartmentState(ApartmentState.STA);
            sta.Start();
            sta.Join();
            return null;
        });
    }
}

/// <summary>
/// What one <see cref="LiveOutlookRestart.Restart"/> did, by OUTLOOK.EXE process id: the ones running before the stop
/// and the ones running once the start was done.
/// </summary>
/// <param name="ProcessIdsBefore">OUTLOOK.EXE ids before the stop.</param>
/// <param name="ProcessIdsAfter">OUTLOOK.EXE ids after the start.</param>
internal sealed record OutlookRestart(IReadOnlyList<int> ProcessIdsBefore, IReadOnlyList<int> ProcessIdsAfter)
{
    /// <summary>True when an OUTLOOK.EXE runs and none of them ran before: the restart made a new Outlook. Pure.</summary>
    public bool IsNewOutlook => ProcessIdsAfter.Count > 0 && !ProcessIdsAfter.Any(ProcessIdsBefore.Contains);

    /// <summary>One line for the log.</summary>
    public string Describe()
    {
        return "OUTLOOK.EXE before " + Ids(ProcessIdsBefore) + ", after " + Ids(ProcessIdsAfter)
            + (IsNewOutlook ? " - a new Outlook" : " - NOT a new Outlook");
    }

    private static string Ids(IReadOnlyList<int> ids)
    {
        return ids.Count == 0 ? "none" : "pid " + string.Join(", ", ids.Select(i => i.ToString(CultureInfo.InvariantCulture)));
    }
}

/// <summary>Visible top-level Outlook windows (class rctrl_renwnd32), WM_CLOSE to one, and how long the user has been idle.</summary>
internal static class WindowProbe
{
    internal static IReadOnlyList<IntPtr> VisibleOutlookWindows()
    {
        HashSet<int> outlookPids = new(LiveOutlookRestart.OutlookProcessIds());

        List<IntPtr> result = new();
        if (outlookPids.Count == 0)
        {
            return result;
        }

        NativeWindows.EnumWindows((hwnd, _) =>
        {
            if (!NativeWindows.IsWindowVisible(hwnd))
            {
                return true;
            }

            NativeWindows.GetWindowThreadProcessId(hwnd, out uint pid);
            if (!outlookPids.Contains((int)pid))
            {
                return true;
            }

            StringBuilder className = new(64);
            _ = NativeWindows.GetClassNameW(hwnd, className, className.Capacity);
            if (string.Equals(className.ToString(), "rctrl_renwnd32", StringComparison.Ordinal))
            {
                result.Add(hwnd);
            }

            return true;
        }, IntPtr.Zero);
        return result;
    }

    internal static void PostClose(IntPtr hwnd)
    {
        const uint WmClose = 0x0010;
        _ = TestWin32.PostMessageW(hwnd, WmClose, IntPtr.Zero, IntPtr.Zero);
    }

    /// <summary>Seconds since the interactive user's last keyboard/mouse input.</summary>
    internal static double UserIdleSeconds()
    {
        NativeMethods.LASTINPUTINFO info = new()
        {
            cbSize = (uint)Marshal.SizeOf<NativeMethods.LASTINPUTINFO>(),
        };
        if (!NativeMethods.GetLastInputInfo(ref info))
        {
            return 0; // Unknown - treat as "user just active" (the conservative direction).
        }

        return (Environment.TickCount - (int)info.dwTime) / 1000.0;
    }

    private static class NativeMethods
    {
        [StructLayout(LayoutKind.Sequential)]
        internal struct LASTINPUTINFO
        {
            internal uint cbSize;
            internal uint dwTime;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
    }
}
