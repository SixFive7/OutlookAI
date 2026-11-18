using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the shared way the live tests end Outlook and start it again (<see cref="LiveOutlookRestart"/>, stream K2,
/// 2026-11-18) - AGENTS.md's mailbox-safety rule 7 in one place. The helper runs only against a live Outlook, so what
/// T1 can hold is its pure halves and the order of its rungs, read out of the source as
/// <c>LiveEarlyReturnGuardTests</c> reads the disconnect test's. Nothing here touches Outlook or a mailbox.
/// </summary>
public sealed class LiveOutlookRestartTests
{
    // ------------------------------------------------------------------ what a restart left

    [Theory]
    [InlineData("10", "20", true)]
    [InlineData("", "20", true)]
    [InlineData("10", "10", false)]
    [InlineData("10", "10,20", false)]
    [InlineData("10", "", false)]
    [InlineData("10,11", "11,30", false)]
    public void ARestartIsANewOutlookOnlyWhenNothingThatRanBeforeRunsAfter(string before, string after, bool isNew)
    {
        OutlookRestart restart = new(Ids(before), Ids(after));

        Assert.Equal(isNew, restart.IsNewOutlook);
        Assert.EndsWith(isNew ? " - a new Outlook" : " - NOT a new Outlook", restart.Describe(), StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ the Outbox, on every profile

    [Fact]
    public void AnEmptyOutboxLetsTheStopGoOn()
    {
        Assert.Null(LiveOutlookRestart.OutboxRefusal(0, "the scenario"));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(17)]
    public void AnOutboxHoldingMailRefusesTheStop_NamingTheCountAndWhatWasNotProved(int queued)
    {
        string? refusal = LiveOutlookRestart.OutboxRefusal(queued, "the compose measurement");

        Assert.NotNull(refusal);
        Assert.Contains("holds " + queued.ToString(System.Globalization.CultureInfo.InvariantCulture) + " item(s)", refusal, StringComparison.Ordinal);
        Assert.Contains("proved nothing about the compose measurement", refusal, StringComparison.Ordinal);
        Assert.Contains("fails on every machine profile", refusal, StringComparison.Ordinal);

        // It is the failure message, so nothing in it may read as a pass.
        Assert.DoesNotContain("SKIP", refusal, StringComparison.Ordinal);
        Assert.DoesNotContain("PROVED NOTHING", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void AnUnreadableOutboxCountRefusesTheSameWay()
    {
        string? refusal = LiveOutlookRestart.OutboxRefusal(-1, "the scenario");

        Assert.NotNull(refusal);
        Assert.Contains("could not be read", refusal, StringComparison.Ordinal);
        Assert.Contains("fails on every machine profile", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void TheDisconnectTestRefusesItsOutboxThroughTheSharedRefusal()
    {
        Assert.Equal(
            LiveOutlookRestart.OutboxRefusal(3, "disconnect recovery"),
            LiveDisconnectRecoveryTests.OutboxRefusal(3));
        Assert.Equal(LiveOutlookRestart.RecentInputOnAnUnattendedMachine, LiveDisconnectRecoveryTests.RecentInputOnAnUnattendedMachine);
        Assert.Equal(LiveOutlookRestart.OpenInspectorOnAnUnattendedMachine, LiveDisconnectRecoveryTests.OpenInspectorOnAnUnattendedMachine);
    }

    // ------------------------------------------------------------------ a session that died under a call

    [Fact]
    public void ASessionThatDiedUnderACall_IsToldFromEveryOtherFailure()
    {
        Assert.True(LiveOutlookRestart.IsSessionGone(new ObjectDisposedException("session")));
        Assert.True(LiveOutlookRestart.IsSessionGone(new System.Runtime.InteropServices.InvalidComObjectException("detached")));
        Assert.True(LiveOutlookRestart.IsSessionGone(new System.Runtime.InteropServices.COMException("gone", unchecked((int)0x800706BA))));
        Assert.True(LiveOutlookRestart.IsSessionGone(new System.Runtime.InteropServices.COMException("gone", unchecked((int)0x80010108))));

        // Any other failure is a failure of the step, and the stop says so instead of moving on.
        Assert.False(LiveOutlookRestart.IsSessionGone(new System.Runtime.InteropServices.COMException("refused", unchecked((int)0x80004005))));
        Assert.False(LiveOutlookRestart.IsSessionGone(new InvalidOperationException("no")));
        Assert.False(LiveOutlookRestart.IsSessionGone(new TimeoutException("wedged")));

        Assert.Equal("0x800706BA", LiveOutlookRestart.DescribeGone(new System.Runtime.InteropServices.COMException("gone", unchecked((int)0x800706BA))));
        Assert.Equal("ObjectDisposedException", LiveOutlookRestart.DescribeGone(new ObjectDisposedException("session")));
    }

    [Fact]
    public void TheDisconnectTestsPinReleaseUsesTheSharedTest()
    {
        string source = SourceText.LiveSource("LiveDisconnectRecoveryTests.cs");

        Assert.Contains("catch (Exception ex) when (LiveOutlookRestart.IsSessionGone(ex))", source, StringComparison.Ordinal);
        Assert.DoesNotContain("static bool IsSessionGone(", source, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ rule 7, in the order of the rungs

    [Fact]
    public void TheStopCountsInspectorsAndOutboxesBeforeItLooksAtAWindow()
    {
        string[] stop = StopBody();

        int counts = Index(stop, "ReadSafetyCounts()");
        int inspectorStop = Index(stop, "OpenInspectorOnAnUnattendedMachine,");
        int outbox = Index(stop, "OutboxRefusal(outbox, scenario)");
        int firstWindow = Index(stop, "WindowProbe.VisibleOutlookWindows()");

        Assert.True(
            counts < inspectorStop && inspectorStop < outbox && outbox < firstWindow,
            $"expected the counts, the Inspector stop and the Outbox refusal before the first look at a window; found lines {counts}, {inspectorStop}, {outbox}, {firstWindow}");
    }

    [Fact]
    public void EveryRungIsTriedInOrder_AndQuitComesLast()
    {
        string[] stop = StopBody();

        int close = Index(stop, "WindowProbe.PostClose(window);");
        int hidden = Index(stop, "CloseHiddenExplorers(log)");
        int shown = Index(stop, "ShowOneExplorer(settings, scenario, log)");
        int quit = Index(stop, "QuitFromAnEmptyClient();");

        Assert.True(
            close < hidden && hidden < shown && shown < quit,
            $"expected WM_CLOSE, the hidden Explorers, one Explorer shown and closed, then Quit; found lines {close}, {hidden}, {shown}, {quit}");
        Assert.Single(stop, l => l.Contains("QuitFromAnEmptyClient();", StringComparison.Ordinal));
    }

    [Fact]
    public void NoWindowIsClosedOrShownBeforeTheKeyboardWasProvedIdle()
    {
        string[] stop = StopBody();

        int[] idleChecks = Indexes(stop, "NobodyAtTheKeyboard(settings, ");
        int[] closes = Indexes(stop, "WindowProbe.PostClose(");
        int show = Index(stop, "ShowOneExplorer(settings, scenario, log)");

        Assert.Equal(2, idleChecks.Length);
        Assert.Equal(2, closes.Length);
        Assert.True(idleChecks[0] < closes[0], "the visible windows are closed before the keyboard was proved idle");
        Assert.True(idleChecks[1] < show && show < closes[1], "the Explorer is shown and closed before the keyboard was proved idle");
    }

    [Fact]
    public void NothingInTheSharedStopEverKillsOutlook()
    {
        string source = SourceText.LiveSource("LiveOutlookRestart.cs");

        Assert.DoesNotContain(".Kill(", source, StringComparison.Ordinal);
        Assert.DoesNotContain("taskkill", source, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("TerminateProcess", source, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ the start

    [Fact]
    public void TheStartGoesThroughTheShell_SoOutlookInheritsNoHandleOfTheTestHost()
    {
        // An Outlook the test host created itself inherited the suite's output pipe and held the guest task open
        // for 50 minutes after the suite had finished (Exchange VM, 2026-11-18).
        string start = SourceText.MemberBody(
            SourceText.Lines(SourceText.LiveSource("LiveOutlookRestart.cs")), "internal static void StartWindowed(", "    ", "LiveOutlookRestart.cs");

        Assert.Contains("UseShellExecute = true", start, StringComparison.Ordinal);
        Assert.DoesNotContain("UseShellExecute = false", start, StringComparison.Ordinal);
        Assert.DoesNotContain("ArgumentList", start, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("Outlook", "\"Outlook\"")]
    [InlineData("Test Profile", "\"Test Profile\"")]
    public void AProfileNameIsOneQuotedArgument(string name, string expected)
    {
        Assert.Equal(expected, LiveOutlookRestart.QuotedArgument(name));
    }

    [Fact]
    public void AProfileNameHoldingAQuoteIsRefused_NotSplit()
    {
        Assert.Throws<InvalidOperationException>(() => LiveOutlookRestart.QuotedArgument("a\"b"));
    }

    // ------------------------------------------------------------------ one copy (Q154)

    [Fact]
    public void TheMeasurementStopsAndStartsOutlookThroughTheSharedHelper()
    {
        string[] lines = SourceText.Lines(SourceText.LiveSource("LiveComposeProbe.cs"));
        string stop = SourceText.MemberBody(lines, "internal static void StopOutlook(", "    ", "LiveComposeProbe.cs");
        string windowed = SourceText.MemberBody(
            lines, "internal static void EnsureWindowedOutlook(LiveTestSettings settings, Action<string> log, Q180ColdStartWatch? coldStart)", "    ", "LiveComposeProbe.cs");

        Assert.Contains("LiveOutlookRestart.Stop(", stop, StringComparison.Ordinal);
        Assert.Contains("LiveOutlookRestart.EnsureWindowed(", windowed, StringComparison.Ordinal);

        // And no second copy of a rung is left behind there: no Quit, no Explorer closed.
        string source = string.Join("\n", lines);
        Assert.DoesNotContain("QuitFromAnEmptyClient", source, StringComparison.Ordinal);
        Assert.DoesNotContain("CloseHiddenExplorers", source, StringComparison.Ordinal);
        Assert.DoesNotContain(".Quit();", source, StringComparison.Ordinal);
        Assert.DoesNotContain("explorer.Close();", source, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ helpers

    private static string[] StopBody()
    {
        return SourceText.Lines(SourceText.MemberBody(
            SourceText.Lines(SourceText.LiveSource("LiveOutlookRestart.cs")),
            "internal static bool Stop(LiveTestSettings settings, string scenario, Action<string> log)",
            "    ",
            "LiveOutlookRestart.cs"));
    }

    private static int Index(string[] lines, string text)
    {
        int index = Array.FindIndex(lines, l => l.Contains(text, StringComparison.Ordinal));
        Assert.True(index >= 0, "the shared stop no longer contains: " + text);
        return index;
    }

    private static int[] Indexes(string[] lines, string text)
    {
        return Enumerable.Range(0, lines.Length).Where(i => lines[i].Contains(text, StringComparison.Ordinal)).ToArray();
    }

    private static IReadOnlyList<int> Ids(string list)
    {
        return list.Length == 0
            ? Array.Empty<int>()
            : list.Split(',').Select(s => int.Parse(s, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
    }
}
