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
/// Safety: only a window-LESS Outlook is promoted (a pre-existing window means a user
/// session - skip); only the window this test itself created is closed; never kill.
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
        if (before.Outlook.Headless != true)
        {
            _output.WriteLine("SKIP: Outlook has a visible window (user session) - this test only drives its own window.");
            return;
        }

        IReadOnlyList<IntPtr> baselineWindows = WindowProbe.VisibleOutlookWindows();
        if (baselineWindows.Count != 0)
        {
            _output.WriteLine($"SKIP: {baselineWindows.Count} visible Outlook window(s) already exist.");
            return;
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

        private static class NativeMethods
        {
            internal delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

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
