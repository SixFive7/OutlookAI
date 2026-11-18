using System.Text.Json;
using ModelContextProtocol.Protocol;
using OutlookAI.ComHost.Protocol;
using OutlookAI.ComHost.Supervision;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins what OutlookAI tells an agent - and through it the user - about an Outlook that is not responding,
/// or never finished starting (Q177 (a) and Q178 (e), decided by the maintainer 2026-11-18).
/// <para>
/// <b>The incident behind it.</b> On 2026-11-15 an Outlook OutlookAI had started with no window deadlocked
/// while it started. The user could not see it or close it, every start of Outlook was handed to it, and
/// what OutlookAI would have said - "restarting Outlook fixes it now", "this also clears by itself" - was
/// true of none of that. The maintainer: OutlookAI never ends Outlook; it reports what is stuck and why,
/// the process, how to end it, and - in his words - "the risks of killing it".
/// </para>
/// </summary>
public sealed class OutlookHangReportTests
{
    private static readonly DateTime Now = new(2026, 11, 15, 16, 2, 0, DateTimeKind.Utc);
    private static readonly DateTime Started = new(2026, 11, 15, 15, 27, 39, DateTimeKind.Utc);

    private static OutlookHangFacts Facts(
        OutlookStuckKind kind = OutlookStuckKind.Hung,
        bool hasWindow = false,
        bool everShowedWindow = false,
        int? startedBy = 68964,
        DateTime? notRespondingSince = null,
        OutlookConnectRecord? connect = null,
        bool connectInProgress = false,
        string reason = "Windows reports every Outlook window as not responding")
    {
        return new OutlookHangFacts(
            kind, reason, 76888, Started, hasWindow, everShowedWindow, startedBy, notRespondingSince, connect, connectInProgress, Now,
            @"C:\Users\<user>\AppData\Local\OutlookAI\outlook-start.log");
    }

    // ------------------------------------------------------------------ an Outlook with no window

    [Fact]
    public void AWindowlessHang_SaysTheUserCannotSeeOrCloseIt_AndHowToEndIt()
    {
        string report = OutlookHangReport.Describe(Facts(notRespondingSince: Started.AddSeconds(5)));

        Assert.Contains("running without a window and has stopped responding", report, StringComparison.Ordinal);
        Assert.Contains("cannot see or close it", report, StringComparison.Ordinal);
        Assert.Contains("starting Outlook does nothing until it is ended", report, StringComparison.Ordinal);
        Assert.Contains("Task Manager (Ctrl+Shift+Esc)", report, StringComparison.Ordinal);
        Assert.Contains("OUTLOOK.EXE with PID 76888", report, StringComparison.Ordinal);
        Assert.Contains("End task", report, StringComparison.Ordinal);
    }

    [Fact]
    public void AWindowlessHang_NeverPromisesThatItClearsOrThatRestartingIsEnough()
    {
        // The two sentences that were wrong on 2026-11-15.
        string report = OutlookHangReport.Describe(Facts());

        Assert.DoesNotContain("clears by itself", report, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("restarting Outlook fixes", report, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("may recover", report, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void EveryReport_CarriesTheRisksOfEndingIt_AndThatOutlookAINeverDoes()
    {
        foreach (OutlookStuckKind kind in new[] { OutlookStuckKind.Hung, OutlookStuckKind.NoWindows, OutlookStuckKind.ConnectStuck })
        {
            foreach (bool hasWindow in new[] { false, true })
            {
                string report = OutlookHangReport.Describe(Facts(kind, hasWindow: hasWindow, everShowedWindow: hasWindow));
                Assert.Contains(OutlookHangReport.EndingRisks, report, StringComparison.Ordinal);
                Assert.Contains("OutlookAI never ends Outlook itself.", report, StringComparison.Ordinal);
                Assert.Contains("PID 76888", report, StringComparison.Ordinal);
                Assert.Contains("outlook-start.log", report, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public void TheRisks_NameWhatEndingOutlookCanCost()
    {
        // The maintainer's examples: unsaved compose windows, mail still in the Outbox, a data file being written.
        Assert.Contains("compose window", OutlookHangReport.EndingRisks, StringComparison.Ordinal);
        Assert.Contains("Outbox", OutlookHangReport.EndingRisks, StringComparison.Ordinal);
        Assert.Contains(".ost or .pst", OutlookHangReport.EndingRisks, StringComparison.Ordinal);
        Assert.Contains("safe mode", OutlookHangReport.EndingRisks, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ an Outlook with a window

    [Fact]
    public void AHangWithAWindow_MayBeOnlyBusy_AndTheUserCanCloseIt()
    {
        string report = OutlookHangReport.Describe(Facts(hasWindow: true, everShowedWindow: true));

        Assert.StartsWith("Outlook has stopped responding:", report, StringComparison.Ordinal);
        Assert.Contains("It may recover if it is only busy.", report, StringComparison.Ordinal);
        Assert.Contains("close it when Windows offers to", report, StringComparison.Ordinal);
        Assert.DoesNotContain("without a window", report, StringComparison.Ordinal);
        Assert.DoesNotContain("clears by itself", report, StringComparison.OrdinalIgnoreCase);
    }

    // ------------------------------------------------------------------ the other kinds of stuck

    [Fact]
    public void AStartThatNeverFinished_SaysSo_AndIsTreatedAsWindowless()
    {
        string report = OutlookHangReport.Describe(Facts(
            OutlookStuckKind.NoWindows, reason: "Outlook started 130 s ago and has still not created its windows"));

        Assert.StartsWith("Outlook has not finished starting: Outlook started 130 s ago", report, StringComparison.Ordinal);
        Assert.Contains("never shown a window", report, StringComparison.Ordinal);
        Assert.Contains("Task Manager", report, StringComparison.Ordinal);
    }

    [Fact]
    public void AConnectionStuckInAStep_NamesItsServerAndStep_AndThatADialogMayBeWaiting()
    {
        OutlookConnectRecord connect = new(76888, Started, true, 68964, 4242, OutlookStartSteps.Logon, Started.AddSeconds(21), Started);
        string report = OutlookHangReport.Describe(Facts(
            OutlookStuckKind.ConnectStuck,
            connect: connect,
            connectInProgress: true,
            reason: "an OutlookAI connection has been waiting on Outlook in one start-up step for 2000 s"));

        Assert.StartsWith("Outlook is not answering OutlookAI:", report, StringComparison.Ordinal);
        Assert.Contains("server process 68964 has been waiting in start-up step 'logon' since 2026-11-15 15:28:00 UTC", report, StringComparison.Ordinal);
        Assert.Contains("dialog", report, StringComparison.Ordinal);
    }

    [Fact]
    public void AConnectionThatDiedInAStep_IsReportedAsStopped_NotAsWaiting()
    {
        OutlookConnectRecord connect = new(76888, Started, true, 68964, 4242, OutlookStartSteps.Pin, Started.AddSeconds(22), Started);
        string report = OutlookHangReport.Describe(Facts(connect: connect, connectInProgress: false));

        Assert.Contains("stopped in start-up step 'pin' at 2026-11-15 15:28:01 UTC without finishing", report, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ who started it, and since when

    [Fact]
    public void WhoStartedIt_AndSinceWhenItIsStuck_AreInTheReport_InUtcAndAsAges()
    {
        string report = OutlookHangReport.Describe(Facts(notRespondingSince: Started.AddSeconds(5)));

        Assert.Contains("started 34 min ago (2026-11-15 15:27:39 UTC), by OutlookAI for its server process 68964", report, StringComparison.Ordinal);
        Assert.Contains("not responding since 2026-11-15 15:27:44 UTC, 34 min ago", report, StringComparison.Ordinal);
    }

    [Fact]
    public void AnOutlookOutlookAIDidNotStart_IsNotClaimed()
    {
        string report = OutlookHangReport.Describe(Facts(startedBy: null));
        Assert.DoesNotContain("by OutlookAI", report, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(5, "5 s")]
    [InlineData(119, "119 s")]
    [InlineData(120, "2 min")]
    [InlineData(7199, "119 min")]
    [InlineData(7200, "2 h")]
    public void Ages_AreWrittenTheWayAPersonReadsThem(int seconds, string expected)
    {
        Assert.Equal(expected, OutlookHangReport.Ago(TimeSpan.FromSeconds(seconds)));
    }

    // ------------------------------------------------------------------ the facts, from a look and the record's notes

    [Fact]
    public void TheRecordsNotes_CountOnlyForTheOutlookTheyAreAbout()
    {
        OutlookStartObservation look = Look(OutlookLivenessState.Hung, hasWindow: false);
        OutlookStartVerdict verdict = OutlookStartPolicy.Decide(look.ToFacts());

        OutlookHangFacts mine = OutlookHangReport.FactsFor(
            verdict,
            look,
            new OutlookProcessNote(76888, Started, Started.AddSeconds(1), 68964),
            new OutlookProcessNote(76888, Started, Started.AddSeconds(5), null),
            new OutlookProcessNote(76888, Started, Started.AddSeconds(9), null),
            "log");
        Assert.Equal(68964, mine.StartedByServerProcessId);
        Assert.Equal(Started.AddSeconds(5), mine.NotRespondingSinceUtc);
        Assert.True(mine.EverShowedWindow);

        // The same notes about the Outlook before it, under a reused process id.
        OutlookHangFacts stale = OutlookHangReport.FactsFor(
            verdict,
            look,
            new OutlookProcessNote(76888, Started.AddHours(-2), Started, 1),
            new OutlookProcessNote(76888, Started.AddHours(-2), Started, null),
            new OutlookProcessNote(60368, Started, Started, null),
            "log");
        Assert.Null(stale.StartedByServerProcessId);
        Assert.Null(stale.NotRespondingSinceUtc);
        Assert.False(stale.EverShowedWindow);
    }

    [Fact]
    public void HealthsStartBlock_ComesFromTheLook()
    {
        OutlookStartObservation look = Look(OutlookLivenessState.Hung, hasWindow: false);
        OutlookStartVerdict verdict = OutlookStartPolicy.Decide(look.ToFacts());

        OutlookStartHealthView view = MailService.DescribeStart(look, verdict);

        Assert.Equal("not-responding", view.State);
        Assert.Equal(76888, view.ProcessId);
        Assert.Equal(Started, view.StartedUtc);
        Assert.False(view.HasWindow);
        Assert.True(view.AddInStarted);
        Assert.EndsWith(OutlookStartLog.FileName, view.LogPath, StringComparison.Ordinal);
        Assert.Contains(view.StartedBy, new[] { "outlookai", "unknown" });
    }

    // ------------------------------------------------------------------ what reaches the agent

    [Fact]
    public void TheAgent_IsToldToTellTheUser_AndToLeaveTheDecisionToThem()
    {
        // Q177 (a) and the maintainer's addition: report back so the user can report to us, or ask the
        // agent to fix it - never end it on the agent's own initiative.
        Assert.Contains("TELL THE USER", OutlookNotRespondingException.AgentAdvice, StringComparison.Ordinal);
        Assert.Contains("only if the user asks you to", OutlookNotRespondingException.AgentAdvice, StringComparison.Ordinal);
        Assert.Contains("never ends", OutlookNotRespondingException.AgentAdvice, StringComparison.Ordinal);
    }

    [Fact]
    public async Task TheToolGuard_ReturnsTheReport_AsOutlookUnresponsive_WithRetryGuidance()
    {
        CallToolResult result = await OutlookTools.GuardAsync<int>(
            CancellationToken.None, () => throw new OutlookNotRespondingException("the report", 30, 76888));

        Assert.True(result.IsError);
        string text = Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text;
        using JsonDocument payload = JsonDocument.Parse(text);
        JsonElement error = payload.RootElement.GetProperty("error");
        Assert.Equal("OutlookUnresponsive", error.GetProperty("type").GetString());
        Assert.Equal("the report", error.GetProperty("message").GetString());
        Assert.Equal(OutlookNotRespondingException.AgentAdvice, error.GetProperty("advice").GetString());
        Assert.Equal(MutationOutcome.Unchanged, error.GetProperty("outcome").GetString());
        Assert.Equal(30, error.GetProperty("retryAfterSeconds").GetInt32());
    }

    [Fact]
    public void TheReport_CrossesTheComHostsPipe_AsItself()
    {
        Exception rebuilt = ComHostErrorMapper.ToException(new ComHostError
        {
            Type = nameof(OutlookNotRespondingException),
            Message = "the report",
        });

        OutlookNotRespondingException report = Assert.IsType<OutlookNotRespondingException>(rebuilt);
        Assert.Equal("the report", report.Message);
        Assert.Equal(OutlookStartPolicy.NotRespondingRetryAfterSeconds, report.RetryAfterSeconds);

        // A TimeoutException, so the freshness sweep degrades on it as on every deadline.
        Assert.IsAssignableFrom<TimeoutException>(rebuilt);
    }

    private static OutlookStartObservation Look(OutlookLivenessState liveness, bool hasWindow)
    {
        OutlookLivenessSnapshot windows = new(liveness, "2 of 2 UI windows hung", 76888, 4321, IntPtr.Zero, hasWindow);
        AddInStartupStamp stamp = AddInStartupStamp.Evaluate(
            76888, Started, 3, 76888, Started.AddSeconds(1).ToString("o", System.Globalization.CultureInfo.InvariantCulture), null);
        return new OutlookStartObservation(Now, windows, Started, stamp, false, false, null);
    }
}
