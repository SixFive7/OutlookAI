using System.Runtime.InteropServices;
using System.Text;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// D33 live proof (soak-fix batch 2026-07-23): NO tool outside the show-me set and
/// draft display=true may create an Outlook window. Strategy: count VISIBLE Outlook
/// top-level windows of class rctrl_renwnd32 (Explorer + Inspector share it; the
/// class filter keeps Outlook-owned toasts/reminder dialogs from flaking the delta),
/// run every COM-touching non-show-me operation, and assert the count NEVER grows.
/// Delta-based, so it holds both against a headless autostarted Outlook (count 0) and
/// a user session with open windows.
/// <para>
/// <b>The read and thread checks say when they could not run (Q101, 2026-10-03).</b> They need
/// a hub hit that still opens, and a conversation on it; without one they used to log a line, or
/// nothing, and the test passed on the rest. Each now goes through
/// <see cref="LivePopulationCoverage"/>: a refusal on a Production profile, a <c>PROVED NOTHING:</c>
/// line on a Portable one. Pinned by <c>T1/LiveEarlyReturnGuardTests</c>.
/// </para>
/// </summary>
[Collection(LiveCollections.Lifecycle)]
[Trait("Category", "Live")]
public sealed class LiveHeadlessGuaranteeTests
{
    /// <summary>What the read and thread checks need, named as the Production refusal wraps it.</summary>
    internal const string OpenableHubHitPopulation =
        "a hit from a hub-scoped search that opens as a live item";

    /// <summary>What a reader of a PROVED NOTHING line about it is to do about it.</summary>
    internal const string OpenableHubHitRemedy =
        "The read and thread window-delta checks open one hit of a hub-scoped search. The hub is where every test "
        + "writes, so an index row can outlive the item it names, and a hub with no hit at all holds no mail the search "
        + "can see. On a test guest the hub is the population Testbed/guest/Reset-HubPopulation.ps1 rebuilds at the start "
        + "of the run; on a working profile it is real mail.";

    /// <summary>What the thread check needs on top, named as the Production refusal wraps it.</summary>
    internal const string ConversationPopulation =
        "a conversation id on the hub item the read check opened";

    /// <summary>What a reader of a PROVED NOTHING line about it is to do about it.</summary>
    internal const string ConversationRemedy =
        "The thread window-delta check follows the conversation of the item the read check opened, and read returned "
        + "no conversation id for it, so there was no conversation to follow. Give the hub mail that belongs to a "
        + "conversation among its newest items, which is where a hub-scoped search of three looks.";

    private readonly LiveLifecycleFixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveHeadlessGuaranteeTests(LiveLifecycleFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    public void NonShowMeOperations_NeverCreateAnOutlookWindow()
    {
        MailService service = _fixture.Service;
        string hub = _fixture.Hub;

        // First COM op establishes the session (D17 autostart when Outlook is off -
        // itself a non-show-me path that must produce no window).
        AccountsOutcome accounts = service.ListAccounts();
        Assert.True(accounts.Accounts.Count > 0);

        // Rolling per-op baselines: each op is judged on the window count right before
        // it, so user activity BETWEEN ops (this suite runs on a live daily-use
        // machine) cannot flake the delta - only growth DURING one of our operations
        // fails, which is exactly the D33 violation shape.
        int baseline = CountVisibleOutlookWindows();
        _output.WriteLine($"baseline visible Outlook windows: {baseline}");

        AssertNoNewWindow(baseline, "list_accounts (repeat)", () => service.ListAccounts());
        AssertNoNewWindow(baseline, "list_folders", () => service.ListFolders(hub));
        AssertNoNewWindow(baseline, "outlook_health", () => service.Health());
        AssertNoNewWindow(baseline, "list_signatures", () => service.ListSignatures());

        SearchOutcome search = AssertNoNewWindow(baseline, "search (hub-scoped, always-fresh)", () => service.Search(new SearchRequest
        {
            Store = hub,
            Top = 3,
            SnippetChars = 0,
        }));

        // Any hub hit will do - but an index row can outlive the item it names (the
        // documented test-tier interference rule: rows of deleted test artifacts survive
        // the delete, and the hub is where every test writes). That is a locate failure,
        // not a window-delta failure, so walk the hits until one opens.
        ReadOutcome? read = null;
        foreach (HitSummary candidate in search.Hits)
        {
            try
            {
                read = AssertNoNewWindow(
                    baseline, "read", () => service.Read(candidate.Id, maxBodyChars: 500, includeHeaders: false));
                break;
            }
            catch (InvalidOperationException ex)
            {
                _output.WriteLine($"hub hit not openable (stale index row), trying the next: {ex.Message}");
            }
        }

        // Q101 - the Q57 pattern for the two checks above that need a population: a hub hit that
        // opens, and a conversation on it. Each used to end in a log line, or - the thread check,
        // when the item carried no conversation id - in nothing at all, and the test then passed on
        // the draft check below. Now a refusal on a Production profile and a PROVED NOTHING line on a
        // Portable one; the draft check still runs on Portable either way.
        _output.WriteLine($"hub search: {search.Hits.Count} hit(s), {(read == null ? "none" : "one")} opened as a live item.");
        IReadOnlyList<ReadOutcome> opened = LivePopulationCoverage.Require(
            _fixture.Settings,
            read == null ? Array.Empty<ReadOutcome>() : new[] { read },
            OpenableHubHitPopulation,
            "the read and thread window-delta checks",
            OpenableHubHitRemedy,
            _output.WriteLine);
        if (opened.Count > 0)
        {
            IReadOnlyList<string> conversations = LivePopulationCoverage.Require(
                _fixture.Settings,
                opened[0].ConversationId == null ? Array.Empty<string>() : new[] { opened[0].ConversationId! },
                ConversationPopulation,
                "the thread window-delta check",
                ConversationRemedy,
                _output.WriteLine);
            if (conversations.Count > 0)
            {
                AssertNoNewWindow(baseline, "thread", () => service.Thread(conversations[0], id: null, store: hub, top: 5));
            }
        }

        // Draft with display:false - the signature GetInspector touch must stay a
        // HIDDEN inspector (Phase-4 fact 5) and no window may become visible.
        string subject = $"{LiveOutlookTestMailer.SubjectTag} headless-guarantee {_fixture.RunMarker}";
        DraftOutcome draft = AssertNoNewWindow(baseline, "new_draft display:false", () => service.NewDraft(
            LiveStoreWriteGuard.Writable(hub, StoreWriteKind.Draft, "new_draft"), hub, cc: null, subject,
            "Agent-authored body (D33 window-delta check). " + _fixture.RunMarker, display: false));
        Assert.False(draft.Displayed);
        try
        {
            // Deterministic since the SnapshotDraft store fallback (this batch): the
            // Parent.Store probe can transiently null on a fresh Outlook.
            Assert.Equal(hub, draft.Store, ignoreCase: true);
        }
        finally
        {
            // Stable-zero loop, not a one-shot: a just-saved draft can materialize in
            // the folder seconds after a single delete pass (Phase-4 fact 6 - this
            // exact lag left a stray draft in the first full-suite run of this batch).
            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(
                hub, _fixture.RunMarker, window: TimeSpan.FromSeconds(60), stableFor: TimeSpan.FromSeconds(10));
        }

        _output.WriteLine("all non-show-me operations left the visible window count unchanged");
    }

    private T AssertNoNewWindow<T>(int _, string label, Func<T> operation)
    {
        int before = CountVisibleOutlookWindows();
        T result = operation();
        int after = CountVisibleOutlookWindows();
        _output.WriteLine($"{label}: visible windows {before} -> {after}");
        Assert.True(after <= before,
            $"'{label}' grew the visible Outlook window count {before} -> {after} (D33 violation)");
        return result;
    }

    // ------------------------------------------------------------------ window probe

    private static int CountVisibleOutlookWindows()
    {
        HashSet<int> outlookPids = new();
        foreach (System.Diagnostics.Process p in System.Diagnostics.Process.GetProcessesByName("OUTLOOK"))
        {
            using (p)
            {
                outlookPids.Add(p.Id);
            }
        }

        if (outlookPids.Count == 0)
        {
            return 0;
        }

        int count = 0;
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
                count++;
            }

            return true;
        }, IntPtr.Zero);
        return count;
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
    }
}
