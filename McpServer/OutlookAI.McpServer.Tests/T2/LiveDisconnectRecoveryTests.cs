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
        _ = service.ListAccounts();
        int stores = independentGateway.Run(s => s.GetStores().Count);
        Assert.True(stores > 0);

        // Wiring pin: the Quit-sink advise must keep succeeding (defense-in-depth; the
        // process-exit watcher is the empirically load-bearing signal).
        Assert.True(independentGateway.QuitSinkActive == true,
            "Application Quit sink failed to advise on the pumped STA");

        HealthOutcome before = service.Health();
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

            IReadOnlyList<ComInspectorInfo> inspectors = independentGateway.Run(s => s.GetOpenInspectors());
            if (inspectors.Count > 0)
            {
                _output.WriteLine($"SKIP: {inspectors.Count} open Inspector window(s) (possible unsent compose) - not closing anything.");
                return;
            }

            int outboxItems = independentGateway.Run(s => s.CountOutboxItems());
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
            _ = PollUntil(() => WindowProbe.VisibleOutlookWindows().Count == 0, TimeSpan.FromSeconds(60));
            _ = independentGateway.Run(s => s.TryCloseInvisibleExplorers());

            bool preExited = PollUntil(() => Process.GetProcessesByName("OUTLOOK").Length == 0, TimeSpan.FromSeconds(120));
            Assert.True(preExited, "Outlook did not exit within 120 s of closing its parked windows (safety counts were clean)");
            _output.WriteLine("windowed Outlook exited after graceful close; re-autostarting headless for the scenario");

            // Fresh headless Outlook for the actual scenario (D17 autostart).
            _ = independentGateway.Run(s => s.GetStores().Count);
            baselineWindows = WindowProbe.VisibleOutlookWindows();
            if (baselineWindows.Count != 0)
            {
                _output.WriteLine("SKIP: a window appeared during re-autostart (user activity?) - stopping here.");
                return;
            }
        }

        // Promote with ONE window of our own via the sanctioned goto surface (hub store).
        ComExplorerState? explorerState = independentGateway.Run(s =>
        {
            ComExplorerState? state = s.TryGotoFolder(_fixture.Hub, null, out string? error);
            Assert.True(state != null, "TryGotoFolder failed: " + (error ?? "unknown"));
            return state;
        });
        _output.WriteLine($"promoted: explorer on '{explorerState!.CurrentFolderPath}'");

        IntPtr ourWindow = IntPtr.Zero;
        DateTime windowDeadline = DateTime.UtcNow.AddSeconds(15);
        while (DateTime.UtcNow < windowDeadline && ourWindow == IntPtr.Zero)
        {
            IReadOnlyList<IntPtr> now = WindowProbe.VisibleOutlookWindows();
            IntPtr fresh = now.FirstOrDefault(h => !baselineWindows.Contains(h));
            if (fresh != IntPtr.Zero)
            {
                ourWindow = fresh;
                break;
            }

            Thread.Sleep(250);
        }

        Assert.True(ourWindow != IntPtr.Zero, "the promoted Explorer window never became visible");

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
        bool windowGone = PollUntil(() => WindowProbe.VisibleOutlookWindows().Count == 0, TimeSpan.FromSeconds(30));
        Assert.True(windowGone, "our window did not close");
        Thread.Sleep(3000); // well past the measured ~1-2 s forced-shutdown window
        Assert.True(
            Process.GetProcessesByName("OUTLOOK").Length > 0,
            "D49 regression: Outlook exited when its last window closed - the compose-surface pin is not holding it");
        Assert.True(independentGateway.IsConnected, "the session must stay connected across a window close");
        _output.WriteLine("D49: Outlook survived losing its last window and is headless again; session still connected");

        // Now relinquish the pin, which is the ONLY thing still keeping Outlook alive -
        // otherwise the disconnect scenario below cannot be staged at all any more.
        int closedExplorers = independentGateway.Run(s => s.TryCloseInvisibleExplorers());
        _output.WriteLine($"released the lifetime pin ({closedExplorers} invisible Explorer(s) closed)");

        // (1) Background release: the independent gateway receives NO calls - only the
        // process-exit watcher can flip IsConnected (the sharp SF-2 assert).
        bool refsReleased = PollUntil(() => !independentGateway.IsConnected, TimeSpan.FromSeconds(45));
        Assert.True(refsReleased, "held COM refs were not released within 45 s of the Outlook exit (SF-2 watcher regression)");
        _output.WriteLine("independent gateway released its session (background watcher)");

        // (2) Full exit, no zombie (measured ~1-2 s on this machine; generous cap).
        bool exited = PollUntil(() => Process.GetProcessesByName("OUTLOOK").Length == 0, TimeSpan.FromSeconds(120));
        Assert.True(exited, "OUTLOOK.EXE did not exit within 120 s of its last window closing");
        _output.WriteLine("OUTLOOK.EXE exited cleanly");

        // (3) SF-1 shape on the service surface: probed comConnected never reports a
        // dead held session; headless is omitted when not running.
        HealthOutcome after = service.Health();
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
                    SearchOutcome degraded = service.Search(new SearchRequest
                    {
                        Query = "oaimcpDegradationProbe" + _fixture.RunMarker,
                        Store = _fixture.Hub,
                        Top = 5,
                        SnippetChars = 0,
                    });

                    Assert.NotNull(degraded.Sweep);
                    Assert.False(degraded.Sweep!.Performed, "the sweep must degrade while the installer mutex is held");
                    Assert.NotNull(degraded.Sweep.Error);
                    Assert.Contains("mutex", degraded.Sweep.Error!, StringComparison.OrdinalIgnoreCase);
                    Assert.NotNull(degraded.Advice);
                    Assert.Contains(degraded.Advice!, a => a.Contains("Freshness sweep unavailable", StringComparison.OrdinalIgnoreCase));
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
        AccountsOutcome reattached = service.ListAccounts();
        Assert.True(reattached.Accounts.Count > 0);
        HealthOutcome healed = service.Health();
        Assert.True(healed.Outlook.Running);
        Assert.True(healed.Outlook.ComConnected);
        Assert.True(healed.Outlook.Headless == true, "D17 re-autostart must come up headless (D33)");
        _output.WriteLine("gateway reattached: Outlook re-autostarted headless");
    }

    private static bool PollUntil(Func<bool> condition, TimeSpan timeout)
    {
        DateTime deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            if (condition())
            {
                return true;
            }

            Thread.Sleep(500);
        }

        return condition();
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
