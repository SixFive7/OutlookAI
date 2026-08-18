using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// SF-1/SF-2 live proof (soak-fix batch 2026-07-23, probe-measured): when the Outlook
/// this gateway is attached to goes away, the background process-exit watcher must
/// release every held COM ref promptly, health must report PROBED (never stale)
/// connectivity, and the next COM call must re-autostart headless (D17/D33).
///
/// The Outlook exit is driven through the USER-CLOSE path: promote the headless
/// Outlook with ONE Explorer window (goto surface - the sanctioned window creator),
/// then WM_CLOSE exactly that window. Measured on this machine: a windowed Outlook
/// fully exits within ~1-2 s of a window close even with COM clients attached
/// (Outlook 2007 SP2+ forced shutdown) - so this path can never zombie.
/// A programmatic Application.Quit-while-attached instead PARKS Outlook indefinitely
/// (probe: no Quit event reaches out-of-process sinks, the process never exits, COM
/// pings still answer; it unsticks ~6 s after the refs are released) - that scenario
/// is a documented protocol rule (release/stop server sessions BEFORE driving Quit),
/// deliberately not reproduced here because inducing the park would wedge the suite.
///
/// Safety: pre-existing windows (normally the show-me tests' parked hub Explorer -
/// this collection runs last) are closed gracefully ONLY under the S7 quit-when-safe
/// counts (user idle >= 3 min, zero open Inspectors, every Outbox empty) - otherwise
/// the test skips; never kill. Side benefit: a full-suite run now ENDS with Outlook
/// headless (D33) instead of leaving the show-me Explorer open.
///
/// EVERY WAIT IN HERE IS BOUNDED, and that is a safety property rather than a tidiness
/// one. On 2026-08-18 this test ran 22.5 minutes and was still going: its polls were all
/// bounded, but the service and gateway calls BETWEEN them were not, and in the live tier
/// they cannot be - the in-process ComGateway accepts a budget and documents that it
/// cannot enforce one, because a blocked COM call is not cancellable and killing the
/// caller is not an option when the caller is us. So the run was stopped by hand, which
/// skipped the fixture teardown, which is what left 7 tagged items sitting in a real
/// mailbox. An unbounded live test is worse than a failing one: it burns the run AND it
/// loses the artifact sweep. Every blocking step now goes through ScenarioClock.Step and
/// every condition through ScenarioClock.WaitUntil, so the failure names the step, its
/// budget, and the state observed when it expired.
/// </summary>
[Collection("LiveLifecycle")]
[Trait("Category", "Live")]
public sealed class LiveDisconnectRecoveryTests
{
    private readonly LiveLifecycleFixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveDisconnectRecoveryTests(LiveLifecycleFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    public void OutlookExit_ReleasesHeldRefsInBackground_HealthProbes_GatewayReattaches()
    {
        MailService service = _fixture.Service;

        // Two independent ref holders, like the day-1 incident shape.
        using ComGateway independentGateway = new();

        string lastHealth = "not probed yet";
        ScenarioClock clock = new ScenarioClock(
            _output,
            () => "OUTLOOK.EXE processes=" + Process.GetProcessesByName("OUTLOOK").Length
                + ", visible Outlook windows=" + WindowProbe.VisibleOutlookWindows().Count
                + ", independent gateway IsConnected=" + independentGateway.IsConnected
                + ", last health reading: " + lastHealth);

        _ = clock.Step("first service call (list_accounts)", service.ListAccounts);
        int stores = clock.Step(
            "count stores through the independent gateway",
            () => independentGateway.Run(s => ((OutlookComSession)s).GetStores().Count));
        Assert.True(stores > 0);

        // Wiring pin: the Quit-sink advise must keep succeeding (defense-in-depth; the
        // process-exit watcher is the empirically load-bearing signal).
        Assert.True(independentGateway.QuitSinkActive == true,
            "Application Quit sink failed to advise on the pumped STA");

        HealthOutcome before = clock.Step("health with a live session", service.Health);
        lastHealth = Describe(before);
        Assert.True(before.Outlook.Running);
        Assert.True(before.Outlook.ComConnected, "probed comConnected must be true with a live session");

        // Guard chain (S7 v2 graceful protocol + user protection): pre-existing visible
        // windows are usually the show-me tests' parked hub-store Explorer (Phase-3
        // fact 3 - every full-suite run leaves one; this collection runs last). Those
        // may be closed gracefully ONLY when the user is not recently active, no
        // Inspector (potential compose) window is open, and every Outbox is empty.
        IReadOnlyList<IntPtr> baselineWindows = WindowProbe.VisibleOutlookWindows();
        if (baselineWindows.Count > 0)
        {
            double idleSeconds = WindowProbe.UserIdleSeconds();
            if (idleSeconds < 180)
            {
                _output.WriteLine($"SKIP: Outlook windows exist and the user was active {idleSeconds:F0} s ago - not closing anything.");
                return;
            }

            IReadOnlyList<ComInspectorInfo> inspectors = clock.Step(
                "count open Inspector windows (S7 safety count)",
                () => independentGateway.Run(s => ((OutlookComSession)s).GetOpenInspectors()));
            if (inspectors.Count > 0)
            {
                _output.WriteLine($"SKIP: {inspectors.Count} open Inspector window(s) (possible unsent compose) - not closing anything.");
                return;
            }

            int outboxItems = clock.Step(
                "count Outbox items (S7 safety count)",
                () => independentGateway.Run(s => ((OutlookComSession)s).CountOutboxItems()));
            if (outboxItems != 0)
            {
                _output.WriteLine($"SKIP: {outboxItems} Outbox item(s) (or count unavailable) - not closing anything.");
                return;
            }

            _output.WriteLine($"closing {baselineWindows.Count} parked Explorer window(s) gracefully (idle {idleSeconds:F0} s, no inspectors, outbox empty)");
            foreach (IntPtr hwnd in baselineWindows)
            {
                WindowProbe.PostClose(hwnd);
            }

            // D49: the parked windows are no longer the only thing holding Outlook up -
            // a live session pins it with an invisible Explorer. Wait for the windows to
            // go, then relinquish the pin, or Outlook will (correctly) refuse to exit.
            // Advisory: proceed either way, the pin release below is what matters.
            _ = clock.TryWaitUntil(
                "the pre-existing Outlook windows have closed",
                () => WindowProbe.VisibleOutlookWindows().Count == 0,
                TimeSpan.FromSeconds(60));
            _ = clock.Step(
                "release the lifetime pin on the pre-existing instance",
                () => independentGateway.Run(s => ((OutlookComSession)s).TryCloseInvisibleExplorers()));

            clock.WaitUntil(
                "the pre-existing Outlook has exited after its parked windows were closed gracefully "
                + "(the S7 safety counts were clean)",
                () => Process.GetProcessesByName("OUTLOOK").Length == 0,
                TimeSpan.FromSeconds(120));
            _output.WriteLine("windowed Outlook exited after graceful close; re-autostarting headless for the scenario");

            // Fresh headless Outlook for the actual scenario (D17 autostart).
            _ = clock.Step(
                "re-autostart Outlook headless for the scenario",
                () => independentGateway.Run(s => ((OutlookComSession)s).GetStores().Count));
            baselineWindows = WindowProbe.VisibleOutlookWindows();
            if (baselineWindows.Count != 0)
            {
                _output.WriteLine("SKIP: a window appeared during re-autostart (user activity?) - stopping here.");
                return;
            }
        }

        // Promote with ONE window of our own via the sanctioned goto surface (hub store).
        ComExplorerState? explorerState = clock.Step(
            "promote Outlook with one Explorer window (goto hub)",
            () => independentGateway.Run(s =>
            {
                ComExplorerState? state = s.TryGotoFolder(_fixture.Hub, null, out string? error);
                Assert.True(state != null, "TryGotoFolder failed: " + (error ?? "unknown"));
                return state;
            }));
        _output.WriteLine($"promoted: explorer on '{explorerState!.CurrentFolderPath}'");

        IntPtr ourWindow = IntPtr.Zero;
        IReadOnlyList<IntPtr> baseline = baselineWindows;
        clock.WaitUntil(
            "the Explorer window we just promoted to become visible",
            () =>
            {
                ourWindow = WindowProbe.VisibleOutlookWindows().FirstOrDefault(h => !baseline.Contains(h));
                return ourWindow != IntPtr.Zero;
            },
            TimeSpan.FromSeconds(15));

        IReadOnlyList<IntPtr> beforeClose = WindowProbe.VisibleOutlookWindows();
        if (beforeClose.Count != 1)
        {
            _output.WriteLine($"SKIP: expected exactly our window before close, saw {beforeClose.Count} (user activity?) - closing ours and stopping.");
            WindowProbe.PostClose(ourWindow);
            return;
        }

        // Graceful user-close-equivalent on exactly our window.
        _output.WriteLine("posting WM_CLOSE to our window");
        WindowProbe.PostClose(ourWindow);

        // (0) D49 - THE NEW CONTRACT, asserted before the old scenario can even be
        // staged. Losing its last window used to kill Outlook within ~1-2 s even with
        // COM clients attached, and that is exactly what broke the compose path: closing
        // the Inspector a draft was written in took the whole instance down, which is
        // where update_draft's com_failure and the three RPC_S_SERVER_UNAVAILABLE suite
        // failures came from. A live session now holds an invisible Explorer, so Outlook
        // returns to the window-less state the product prefers (D33) instead of dying.
        clock.WaitUntil(
            "the Explorer window we posted WM_CLOSE to has gone",
            () => WindowProbe.VisibleOutlookWindows().Count == 0,
            TimeSpan.FromSeconds(30));
        Thread.Sleep(3000); // well past the measured ~1-2 s forced-shutdown window
        Assert.True(
            Process.GetProcessesByName("OUTLOOK").Length > 0,
            "D49 regression: Outlook exited when its last window closed - the compose-surface pin is not holding it");
        _output.WriteLine(
            "D49: Outlook survived losing its last window and is headless again "
            + $"(session IsConnected={independentGateway.IsConnected} - passive flag, healed on the next call)");

        // Now relinquish the pin, which is the ONLY thing still keeping Outlook alive -
        // otherwise the disconnect scenario below cannot be staged at all any more.
        int closedExplorers = clock.Step(
            "release the lifetime pin so Outlook can exit",
            () => independentGateway.Run(s => ((OutlookComSession)s).TryCloseInvisibleExplorers()));
        _output.WriteLine($"released the lifetime pin ({closedExplorers} invisible Explorer(s) closed)");

        // (1) Background release: the independent gateway receives NO calls - only the
        // process-exit watcher can flip IsConnected (the sharp SF-2 assert).
        clock.WaitUntil(
            "the independent gateway released its held COM refs by itself (SF-2 process-exit watcher)",
            () => !independentGateway.IsConnected,
            TimeSpan.FromSeconds(45));
        _output.WriteLine("independent gateway released its session (background watcher)");

        // (2) Full exit, no zombie (measured ~1-2 s on this machine; generous cap).
        clock.WaitUntil(
            "OUTLOOK.EXE has exited after losing its last window and its pin",
            () => Process.GetProcessesByName("OUTLOOK").Length == 0,
            TimeSpan.FromSeconds(120));
        _output.WriteLine("OUTLOOK.EXE exited cleanly");

        // (3) SF-1 shape on the service surface: probed comConnected never reports a
        // dead held session; headless is omitted when not running.
        HealthOutcome after = clock.Step("health with Outlook stopped", service.Health);
        lastHealth = Describe(after);
        Assert.False(after.Outlook.Running);
        Assert.False(after.Outlook.ComConnected, "comConnected reported a dead session as connected (SF-1 regression)");
        Assert.Null(after.Outlook.Headless);

        // (3b) D34 graceful-degradation proof: with Outlook stopped AND the
        // OutlookAISetup installer mutex held (acquired by this test - the D17
        // autostart guard), a search must still SUCCEED with index results, carry the
        // mutex reason in the sweep error + freshness advice, and must NOT start
        // Outlook. The staleness block must report the post-sweep reality
        // (outlookRunning=false - the D34 snapshot fix).
        using (var installerMutex = new System.Threading.Mutex(initiallyOwned: true, "OutlookAISetup", out bool createdNew))
        {
            try
            {
                if (!createdNew)
                {
                    _output.WriteLine("SKIP(3b): a real OutlookAISetup mutex already exists (installer running?) - not simulating.");
                }
                else
                {
                    service.ClearSweepCache(); // A <10 s-old cached sweep would mask the degradation path.
                    SearchOutcome degraded = clock.Step(
                        "degraded search while the installer mutex is held (D34)",
                        () => service.Search(new SearchRequest
                        {
                            Query = "oaimcpDegradationProbe" + _fixture.RunMarker,
                            Store = _fixture.Hub,
                            Top = 5,
                            SnippetChars = 0,
                        }));

                    Assert.NotNull(degraded.Sweep);
                    Assert.False(degraded.Sweep!.Performed, "the sweep must degrade while the installer mutex is held");
                    Assert.NotNull(degraded.Sweep.Error);
                    Assert.Contains("mutex", degraded.Sweep.Error!, StringComparison.OrdinalIgnoreCase);
                    Assert.NotNull(degraded.Advice);
                    // Pins the CONTRACT of the not-run case - the advice must shout, and must name
                    // the sweep as the thing that could not run - rather than a phrase. It used to
                    // assert "Freshness sweep unavailable", which the shipped advice stopped saying
                    // long before this line was last read, so the live tier carried a failure that
                    // had nothing to do with the behaviour under test.
                    Assert.Contains(degraded.Advice!, a => a.Contains("INCOMPLETE RESULTS - TELL THE USER", StringComparison.Ordinal));
                    Assert.Contains(degraded.Advice!, a => a.Contains("live check against Outlook could not", StringComparison.OrdinalIgnoreCase));
                    Assert.Contains(degraded.Advice!, a => a.Contains("add-in update", StringComparison.OrdinalIgnoreCase));
                    Assert.False(degraded.Staleness.OutlookRunning, "staleness must reflect post-sweep reality (D34)");
                    Assert.Empty(Process.GetProcessesByName("OUTLOOK"));
                    _output.WriteLine("degradation proven: search returned index results with mutex-reason advice, no autostart");
                }
            }
            finally
            {
                if (createdNew)
                {
                    installerMutex.ReleaseMutex();
                }
            }
        }

        // (4) Recovery: the next COM-needing call reattaches via D17 autostart - headless (D33).
        // This is the step the 2026-08-18 run was inside when it was stopped, and the one
        // most likely to block without ever returning: it cold-starts OUTLOOK.EXE through
        // Activator.CreateInstance on the Outlook.Application ProgID, which has no timeout
        // of its own and cannot be given one from this side.
        AccountsOutcome reattached = clock.Step("reattach: list_accounts must re-autostart Outlook", service.ListAccounts);
        Assert.True(reattached.Accounts.Count > 0);
        HealthOutcome healed = clock.Step("health after the reattach", service.Health);
        lastHealth = Describe(healed);
        Assert.True(healed.Outlook.Running);
        Assert.True(healed.Outlook.ComConnected);
        Assert.True(healed.Outlook.Headless == true, "D17 re-autostart must come up headless (D33)");
        _output.WriteLine($"gateway reattached: Outlook re-autostarted headless ({clock.Spent})");
    }

    /// <summary>One health reading, short enough to sit inside a failure message.</summary>
    private static string Describe(HealthOutcome health)
    {
        return $"running={health.Outlook.Running} comConnected={health.Outlook.ComConnected} "
            + $"headless={health.Outlook.Headless?.ToString() ?? "null"}";
    }

    /// <summary>
    /// The scenario's clock, its budget, and the diagnosis it produces when either runs out.
    /// <para>
    /// Two numbers, both derived from figures the project already committed to rather than
    /// invented here.
    /// </para>
    /// <para>
    /// <b>Per step: 180 s</b> - one <see cref="LiveInboxArrival.DeadlineSeconds"/>, which is
    /// the live tier's own allowance for the slowest real thing it waits on (a mail crossing
    /// a real mail server). It is also 1.5x <c>ComOperationBudgets.OperationDeadlineMs</c>
    /// (120 s), the point at which the product itself declares a single Outlook operation
    /// wedged and reclaims the COM host. So no step here is bounded more tightly than either
    /// rule the codebase already lives by, and a slow-but-working machine cannot trip it:
    /// nothing this test asks for is a mail round trip, and everything it asks for is one
    /// operation the shipped product would have abandoned an entire minute earlier.
    /// </para>
    /// <para>
    /// <b>Whole scenario: 900 s</b> - five of those. The test's own condition waits already
    /// account for 393 s in the worst case (60 + 120 + 15 + 30 + 3 + 45 + 120), and on top of
    /// that it drives two full Outlook exits, two cold starts, a degraded search and three
    /// health reports. 900 s leaves well over 250 s of headroom above a plausible slow-machine
    /// run, and still fails in two thirds of the time the 2026-08-18 run had already burned
    /// when it was stopped by hand.
    /// </para>
    /// <para>
    /// A step that expires ABANDONS its worker thread, deliberately. The thread is inside a
    /// COM call that nothing in this process can cancel - that is precisely why the product
    /// moved COM into a killable child, and precisely what the in-process gateway used here
    /// documents that it cannot do. Abandoning one pool thread is the cheaper of the two
    /// available outcomes; the other one is the whole run hanging and the artifact sweep
    /// never running.
    /// </para>
    /// </summary>
    private sealed class ScenarioClock
    {
        private static readonly TimeSpan StepBudget = TimeSpan.FromSeconds(LiveInboxArrival.DeadlineSeconds);
        private static readonly TimeSpan ScenarioBudget = TimeSpan.FromSeconds(5 * LiveInboxArrival.DeadlineSeconds);

        private readonly Stopwatch _clock = Stopwatch.StartNew();
        private readonly ITestOutputHelper _output;
        private readonly Func<string> _observe;

        internal ScenarioClock(ITestOutputHelper output, Func<string> observe)
        {
            _output = output;
            _observe = observe;
        }

        /// <summary>Elapsed scenario time, for the success log.</summary>
        internal string Spent => $"{_clock.Elapsed.TotalSeconds:F0} s of the {ScenarioBudget.TotalSeconds:F0} s budget";

        private TimeSpan Remaining
        {
            get
            {
                TimeSpan left = ScenarioBudget - _clock.Elapsed;
                return left > TimeSpan.Zero ? left : TimeSpan.Zero;
            }
        }

        /// <summary>
        /// Runs one blocking COM or service call under a bound, and fails naming it when it
        /// does not come back.
        /// <para>
        /// The body runs on a pool thread, which is where the shipped server runs these calls
        /// too (<c>OutlookTools.GuardAsync</c> wraps every tool in <c>Task.Run</c>), so this
        /// changes no apartment behaviour: the COM work funnels onto the session's own pumped
        /// STA thread either way.
        /// </para>
        /// </summary>
        internal T Step<T>(string what, Func<T> body)
        {
            TimeSpan budget = StepBudget < Remaining ? StepBudget : Remaining;
            _output.WriteLine($"[{_clock.Elapsed.TotalSeconds,5:F0}s] {what} (up to {budget.TotalSeconds:F0} s)");
            if (budget <= TimeSpan.Zero)
            {
                throw new TimeoutException(Expired(what, TimeSpan.Zero, "could not even be started - the scenario budget was already spent"));
            }

            Stopwatch step = Stopwatch.StartNew();
            Task<T> work = Task.Run(body);
            Task finished = Task.WhenAny(work, Task.Delay(budget)).GetAwaiter().GetResult();
            if (!ReferenceEquals(finished, work))
            {
                throw new TimeoutException(Expired(what, step.Elapsed, $"never returned (bound {budget.TotalSeconds:F0} s)"));
            }

            // Not Task.Wait: that would re-wrap a genuine assertion failure inside an
            // AggregateException and bury the message the test meant to report.
            return work.GetAwaiter().GetResult();
        }

        /// <summary>Polls until <paramref name="probe"/> is true, or fails naming the condition.</summary>
        internal void WaitUntil(string condition, Func<bool> probe, TimeSpan timeout)
        {
            if (!Poll(condition, probe, timeout, out TimeSpan budget, out TimeSpan waited))
            {
                throw new TimeoutException(
                    Expired(condition, waited, $"was still not true after {budget.TotalSeconds:F0} s"));
            }
        }

        /// <summary>
        /// The same wait, reported rather than asserted, for the one condition this test
        /// proceeds past either way.
        /// </summary>
        internal bool TryWaitUntil(string condition, Func<bool> probe, TimeSpan timeout)
        {
            return Poll(condition, probe, timeout, out _, out _);
        }

        /// <summary>
        /// The poll itself. It reports the budget it actually used rather than letting the
        /// caller recompute one - by the time a wait has failed, Remaining has shrunk by
        /// exactly the amount that wait took, so a recomputed number would understate the
        /// bound in the very message written to explain it.
        /// <para>
        /// Every probe passed in here reads Windows or a local flag, never COM, so the poll
        /// cannot itself become the thing that hangs.
        /// </para>
        /// </summary>
        private bool Poll(string condition, Func<bool> probe, TimeSpan timeout, out TimeSpan budget, out TimeSpan waited)
        {
            budget = timeout < Remaining ? timeout : Remaining;
            _output.WriteLine($"[{_clock.Elapsed.TotalSeconds,5:F0}s] waiting until {condition} (up to {budget.TotalSeconds:F0} s)");

            Stopwatch poll = Stopwatch.StartNew();
            while (poll.Elapsed < budget)
            {
                if (probe())
                {
                    waited = poll.Elapsed;
                    return true;
                }

                Thread.Sleep(500);
            }

            bool last = probe();
            waited = poll.Elapsed;
            return last;
        }

        /// <summary>
        /// The failure text. It names the condition that was still unmet, what was spent
        /// reaching it, and the state observed at that moment - because a bound whose message
        /// is "timed out" has not removed the debugging cost, only moved it to whoever reads
        /// the next run.
        /// </summary>
        private string Expired(string what, TimeSpan spent, string why)
        {
            string message =
                $"live disconnect scenario stalled waiting for: {what}. It {why}; "
                + $"{spent.TotalSeconds:F0} s spent here, {_clock.Elapsed.TotalSeconds:F0} s of the "
                + $"{ScenarioBudget.TotalSeconds:F0} s scenario budget. Observed: {_observe()}.";
            _output.WriteLine(message);
            return message;
        }
    }

    /// <summary>Visible top-level Outlook windows (class rctrl_renwnd32) + WM_CLOSE poster.</summary>
    internal static class WindowProbe
    {
        internal static IReadOnlyList<IntPtr> VisibleOutlookWindows()
        {
            HashSet<int> outlookPids = new();
            foreach (Process p in Process.GetProcessesByName("OUTLOOK"))
            {
                using (p)
                {
                    outlookPids.Add(p.Id);
                }
            }

            List<IntPtr> result = new();
            if (outlookPids.Count == 0)
            {
                return result;
            }

            NativeMethods.EnumWindows((hwnd, _) =>
            {
                if (!NativeMethods.IsWindowVisible(hwnd))
                {
                    return true;
                }

                NativeMethods.GetWindowThreadProcessId(hwnd, out uint pid);
                if (!outlookPids.Contains((int)pid))
                {
                    return true;
                }

                StringBuilder className = new(64);
                _ = NativeMethods.GetClassName(hwnd, className, className.Capacity);
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
            _ = NativeMethods.PostMessage(hwnd, WmClose, IntPtr.Zero, IntPtr.Zero);
        }

        /// <summary>Seconds since the interactive user's last keyboard/mouse input.</summary>
        internal static double UserIdleSeconds()
        {
            NativeMethods.LASTINPUTINFO info = new()
            {
                cbSize = (uint)System.Runtime.InteropServices.Marshal.SizeOf<NativeMethods.LASTINPUTINFO>(),
            };
            if (!NativeMethods.GetLastInputInfo(ref info))
            {
                return 0; // Unknown - treat as "user just active" (the conservative direction).
            }

            return (Environment.TickCount - (int)info.dwTime) / 1000.0;
        }

        private static class NativeMethods
        {
            internal delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

            [StructLayout(LayoutKind.Sequential)]
            internal struct LASTINPUTINFO
            {
                internal uint cbSize;
                internal uint dwTime;
            }

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool IsWindowVisible(IntPtr hWnd);

            [DllImport("user32.dll")]
            internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            internal static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
        }
    }
}
