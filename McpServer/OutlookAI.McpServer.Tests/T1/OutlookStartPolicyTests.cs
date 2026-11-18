using OutlookAI.ComHost.Supervision;
using OutlookAI.Core.Com;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins when OutlookAI may open a mailbox in an Outlook that has just started, and when "starting" has
/// become "stuck" (Q175 (a), Q175 (e) and Q176 (c), decided by the maintainer 2026-11-18, after the hang
/// of 2026-11-15: an Outlook OutlookAI had started over COM, with no window, deadlocked while its add-ins
/// were still loading, because OutlookAI opened mailboxes within its first second).
/// <para>
/// Every branch of <see cref="OutlookStartPolicy.Decide"/> with plain values: the waits themselves run on
/// a real Outlook only, in the live tier, which is where their cost is measured.
/// </para>
/// </summary>
public sealed class OutlookStartPolicyTests
{
    private const long Young = 5_000;
    private const long Settled = OutlookStartPolicy.MinimumAgeMilliseconds;
    private const long Old = OutlookStartPolicy.StartingLimitMilliseconds;

    private static OutlookStartVerdict Decide(
        OutlookLivenessState liveness = OutlookLivenessState.Responsive,
        long? age = Settled,
        bool? stamped = true,
        bool? idle = null,
        bool lockElsewhere = false,
        long? stepMs = null,
        bool finishedBefore = false)
    {
        return OutlookStartPolicy.Decide(new OutlookStartFacts(liveness, age, stamped, idle, lockElsewhere, stepMs, finishedBefore));
    }

    // ------------------------------------------------------------------ the decided numbers

    [Fact]
    public void TheNumbers_AreTheOnesTheMaintainerDecided()
    {
        // Q175 (a): "Outlook at least about 20 s old". Q176 (c): "A time limit on 'starting', about 120 s".
        Assert.Equal(20_000, OutlookStartPolicy.MinimumAgeMilliseconds);
        Assert.Equal(120_000, OutlookStartPolicy.StartingLimitMilliseconds);
    }

    [Fact]
    public void EveryWait_EndsInsideTheConnectBudget_SoTheCallerGetsAReportNotATimeout()
    {
        // A connection that gives up at a limit answers with the report; past the 180 s connect budget the
        // server would have reclaimed it as a bare timeout instead.
        Assert.True(OutlookStartPolicy.StartingLimitMilliseconds < ComOperationBudgets.ConnectDeadlineMs);
        Assert.True(OutlookStartPolicy.ConnectLockWaitMilliseconds < ComOperationBudgets.ConnectDeadlineMs);

        // ...and a waiting connection outlasts a holder that is settling for as long as it may.
        Assert.True(OutlookStartPolicy.ConnectLockWaitMilliseconds > OutlookStartPolicy.StartingLimitMilliseconds);
    }

    [Fact]
    public void TheIdleThreshold_IsAFractionOfItsSample()
    {
        Assert.InRange(OutlookStartPolicy.IdleCpuMilliseconds, 1, OutlookStartPolicy.IdleSampleMilliseconds / 4);
    }

    // ------------------------------------------------------------------ not running

    [Fact]
    public void NotRunning_WithNobodyStartingIt_IsNotRunning()
    {
        OutlookStartVerdict verdict = Decide(OutlookLivenessState.NotRunning, age: null, stamped: null);
        Assert.Equal(OutlookStartState.NotRunning, verdict.State);
    }

    [Fact]
    public void NotRunning_WhileAnotherConnectionStartsIt_Settles_RatherThanStartingASecond()
    {
        // Q175 (e): one starter at a time, the others wait.
        OutlookStartVerdict verdict = Decide(OutlookLivenessState.NotRunning, age: null, stamped: null, lockElsewhere: true, stepMs: 1_000);
        Assert.Equal(OutlookStartState.Settling, verdict.State);
        Assert.Equal(OutlookStartPolicy.StartRetryAfterSeconds, verdict.RetryAfterSeconds);
    }

    [Fact]
    public void NotRunning_WhileAnotherConnectionHasTriedToStartItPastTheLimit_IsStuck()
    {
        OutlookStartVerdict verdict = Decide(
            OutlookLivenessState.NotRunning, age: null, stamped: null, lockElsewhere: true, stepMs: OutlookStartPolicy.StartingLimitMilliseconds);
        Assert.Equal(OutlookStartState.NotResponding, verdict.State);
        Assert.Equal(OutlookStuckKind.ConnectStuck, verdict.Stuck);
    }

    // ------------------------------------------------------------------ hung and stuck

    [Theory]
    [InlineData(Young)]
    [InlineData(Settled)]
    [InlineData(Old)]
    public void Hung_IsNotResponding_AtAnyAge(long age)
    {
        OutlookStartVerdict verdict = Decide(OutlookLivenessState.Hung, age: age);
        Assert.Equal(OutlookStartState.NotResponding, verdict.State);
        Assert.Equal(OutlookStuckKind.Hung, verdict.Stuck);
        Assert.Equal(OutlookStartPolicy.NotRespondingRetryAfterSeconds, verdict.RetryAfterSeconds);
    }

    [Fact]
    public void NoWindows_IsStartingWhileYoung_AndStuckPastTheLimit()
    {
        // The hole Q176 closes: an Outlook that hangs before it has created its windows used to read as
        // "starting" for ever, and the agent was told nothing was wrong.
        Assert.Equal(OutlookStartState.Settling, Decide(OutlookLivenessState.Starting, age: Old - 1).State);

        OutlookStartVerdict stuck = Decide(OutlookLivenessState.Starting, age: Old);
        Assert.Equal(OutlookStartState.NotResponding, stuck.State);
        Assert.Equal(OutlookStuckKind.NoWindows, stuck.Stuck);
    }

    [Theory]
    [InlineData(Young)]
    [InlineData(Old)]
    [InlineData(Old * 100)]
    public void NoWindows_InAnOutlookThatHadThem_IsClosing_NotStuckStarting(long age)
    {
        // The user closed Outlook and it is shutting down: no windows, an old process. Reporting that as a
        // start that never finished would send the user to Task Manager over a normal exit.
        OutlookStartVerdict verdict = Decide(OutlookLivenessState.Starting, age: age, finishedBefore: true);
        Assert.Equal(OutlookStartState.Settling, verdict.State);
        Assert.Equal("closing", verdict.ReasonCode);
    }

    [Theory]
    [InlineData(OutlookStartSteps.Namespace, true)]
    [InlineData(OutlookStartSteps.Logon, true)]
    [InlineData(OutlookStartSteps.QuitSink, true)]
    [InlineData(OutlookStartSteps.Pin, true)]
    [InlineData(OutlookStartSteps.Connected, true)]
    [InlineData(OutlookStartSteps.CreateApplication, false)]
    [InlineData(OutlookStartSteps.Settle, false)]
    [InlineData(OutlookStartSteps.Failed, false)]
    public void OnlyAStepAfterTheSettleWait_SaysOutlookHadFinishedStarting(string step, bool pastSettle)
    {
        Assert.Equal(pastSettle, OutlookStartSteps.IsPastSettle(step));
    }

    [Fact]
    public void AConnectionWaitingInOneStepPastTheLimit_IsStuck_HoweverOldOutlookIs()
    {
        foreach (long age in new[] { Young, Settled, Old * 10 })
        {
            OutlookStartVerdict verdict = Decide(age: age, lockElsewhere: true, stepMs: OutlookStartPolicy.StartingLimitMilliseconds);
            Assert.Equal(OutlookStartState.NotResponding, verdict.State);
            Assert.Equal(OutlookStuckKind.ConnectStuck, verdict.Stuck);
        }
    }

    [Fact]
    public void AConnectionInItsStep_IsNotStuck_JustBelowTheLimit()
    {
        OutlookStartVerdict verdict = Decide(age: Old * 10, lockElsewhere: true, stepMs: OutlookStartPolicy.StartingLimitMilliseconds - 1);
        Assert.Equal(OutlookStartState.Ready, verdict.State);
    }

    // ------------------------------------------------------------------ settling (Q175 (a))

    [Fact]
    public void AYoungOutlook_Settles_EvenWithEveryOtherConditionMet()
    {
        OutlookStartVerdict verdict = Decide(age: Settled - 1, stamped: true, idle: true);
        Assert.Equal(OutlookStartState.Settling, verdict.State);
        Assert.Equal("young", verdict.ReasonCode);
    }

    [Fact]
    public void AnOutlookOfTwentySeconds_WithItsStampAndAnIdleThread_IsReady()
    {
        OutlookStartVerdict verdict = Decide(age: Settled, stamped: true, idle: true);
        Assert.Equal(OutlookStartState.Ready, verdict.State);
        Assert.Equal(OutlookStartPolicy.SettledCode, verdict.ReasonCode);
        Assert.False(verdict.SettledByAge);
    }

    [Fact]
    public void AMissingStamp_IsWaitedFor_OnlyWhenTheAddInIsSetToLoad()
    {
        Assert.Equal(OutlookStartState.Settling, Decide(stamped: false).State);
        Assert.Equal("add-in-starting", Decide(stamped: false).ReasonCode);

        // No add-in set to load: no stamp is coming, and none is waited for.
        Assert.Equal(OutlookStartState.Ready, Decide(stamped: null).State);
    }

    [Fact]
    public void ABusyMainThread_IsWaitedFor_AndAnUnmeasuredOne_IsLeftToTheConnection()
    {
        Assert.Equal(OutlookStartState.Settling, Decide(idle: false).State);
        Assert.Equal("main-thread-busy", Decide(idle: false).ReasonCode);
        Assert.Equal(OutlookStartState.Ready, Decide(idle: true).State);

        // The gate in front of each request does not take the sample; the connection does, before it
        // opens a mailbox.
        Assert.Equal(OutlookStartState.Ready, Decide(idle: null).State);
    }

    [Fact]
    public void AnotherConnectionInProgress_IsWaitedFor_WhileOutlookIsYoung()
    {
        OutlookStartVerdict verdict = Decide(age: Settled, lockElsewhere: true, stepMs: 2_000);
        Assert.Equal(OutlookStartState.Settling, verdict.State);
        Assert.Equal("connect-elsewhere", verdict.ReasonCode);
    }

    [Fact]
    public void AnOutlookPastTheStartingLimit_IsReady_WhateverIsStillUnmet()
    {
        // The conditions narrow a start-up race; past the limit there is no start-up left to race, and an
        // Outlook whose add-in is disabled must not be refused for ever.
        OutlookStartVerdict verdict = Decide(age: Old, stamped: false, idle: false, lockElsewhere: true, stepMs: 1_000);
        Assert.Equal(OutlookStartState.Ready, verdict.State);
        Assert.True(verdict.SettledByAge);
    }

    [Fact]
    public void AnUnknownAge_DecidesOnTheOtherConditions()
    {
        Assert.Equal(OutlookStartState.Ready, Decide(age: null, stamped: null, idle: true).State);
        Assert.Equal(OutlookStartState.Settling, Decide(age: null, stamped: false).State);
        Assert.Equal(OutlookStartState.Settling, Decide(OutlookLivenessState.Starting, age: null).State);
    }

    // ------------------------------------------------------------------ retry guidance

    [Theory]
    [InlineData(0L)]
    [InlineData(1_000L)]
    [InlineData(19_999L)]
    [InlineData(20_000L)]
    [InlineData(500_000L)]
    public void SettlingRetryGuidance_IsAlwaysANumberAnAgentWillWaitOut(long age)
    {
        Assert.InRange(OutlookStartPolicy.SettlingRetryAfterSeconds(age), 5, 30);
    }

    [Fact]
    public void SettlingRetryGuidance_CountsDownTheMinimumAge()
    {
        Assert.Equal(20, OutlookStartPolicy.SettlingRetryAfterSeconds(5_000));
        Assert.Equal(6, OutlookStartPolicy.SettlingRetryAfterSeconds(19_500));
        Assert.Equal(10, OutlookStartPolicy.SettlingRetryAfterSeconds(null));
    }

    [Fact]
    public void AStartRequest_IsToldToComeBackAfterTheSettleWait()
    {
        // 15 s before the settle wait existed; the wait alone is 20 s now.
        Assert.True(ComHostPolicy.StartingRetryAfterSeconds * 1000 > OutlookStartPolicy.MinimumAgeMilliseconds);
        Assert.Equal(OutlookStartPolicy.StartRetryAfterSeconds, ComHostPolicy.StartingRetryAfterSeconds);
    }

    // ------------------------------------------------------------------ every verdict explains itself

    [Fact]
    public void EveryVerdict_HasAReasonInWords_AndACodeTheLogCanCarry()
    {
        OutlookLivenessState[] states =
        {
            OutlookLivenessState.NotRunning, OutlookLivenessState.Starting, OutlookLivenessState.Responsive, OutlookLivenessState.Hung,
        };
        IEnumerable<OutlookStartFacts> everyCombination =
            from liveness in states
            from age in new long?[] { null, 0, Young, Settled, Old }
            from stamped in new bool?[] { null, false, true }
            from idle in new bool?[] { null, false, true }
            from lockElsewhere in new[] { false, true }
            from step in new long?[] { null, 0, Old }
            from finishedBefore in new[] { false, true }
            select new OutlookStartFacts(liveness, age, stamped, idle, lockElsewhere, step, finishedBefore);

        int decided = 0;
        foreach (OutlookStartFacts facts in everyCombination)
        {
            OutlookStartVerdict verdict = OutlookStartPolicy.Decide(facts);
            Assert.False(string.IsNullOrWhiteSpace(verdict.Reason));
            Assert.Equal(OutlookStartLog.Token(verdict.ReasonCode), verdict.ReasonCode);
            Assert.Equal(verdict.State == OutlookStartState.NotResponding, verdict.Stuck != OutlookStuckKind.None);
            bool comeBack = verdict.State == OutlookStartState.Settling || verdict.State == OutlookStartState.NotResponding;
            Assert.Equal(comeBack, verdict.RetryAfterSeconds > 0);
            decided++;
        }

        Assert.Equal(4 * 5 * 3 * 3 * 2 * 3 * 2, decided);
    }

    // ------------------------------------------------------------------ the supervisor's reading of it

    [Fact]
    public void TheSupervisor_MapsEveryStartState()
    {
        Assert.Equal(LivenessVerdict.Proceed, ComHostPolicy.DecideLiveness(OutlookStartState.Ready, long.MaxValue, true));
        Assert.Equal(LivenessVerdict.Starting, ComHostPolicy.DecideLiveness(OutlookStartState.Settling, long.MaxValue, true));
        Assert.Equal(LivenessVerdict.Hung, ComHostPolicy.DecideLiveness(OutlookStartState.NotResponding, long.MaxValue, true));
        Assert.Equal(LivenessVerdict.MayStart, ComHostPolicy.DecideLiveness(OutlookStartState.NotRunning, long.MaxValue, true));
        Assert.Equal(
            LivenessVerdict.StartSuppressed,
            ComHostPolicy.DecideLiveness(OutlookStartState.NotRunning, ComHostPolicy.AutostartCooldownMilliseconds - 1, true));
        Assert.Equal(LivenessVerdict.StartSuppressed, ComHostPolicy.DecideLiveness(OutlookStartState.NotRunning, long.MaxValue, false));
    }

    [Fact]
    public void TheOldLivenessReading_StillMeansWhatItMeant()
    {
        // The test override and every existing caller still decide from Windows' plain verdict.
        Assert.Equal(LivenessVerdict.Proceed, ComHostPolicy.DecideLiveness(OutlookLivenessState.Responsive, long.MaxValue, true));
        Assert.Equal(LivenessVerdict.Starting, ComHostPolicy.DecideLiveness(OutlookLivenessState.Starting, long.MaxValue, true));
        Assert.Equal(LivenessVerdict.Hung, ComHostPolicy.DecideLiveness(OutlookLivenessState.Hung, long.MaxValue, true));
        Assert.Equal(LivenessVerdict.MayStart, ComHostPolicy.DecideLiveness(OutlookLivenessState.NotRunning, long.MaxValue, true));
    }

    // ------------------------------------------------------------------ the look itself

    [Fact]
    public void TheLookAtOutlook_IsTotalAndFast_WhateverTheMachine()
    {
        // Runs on the build VM, where no Outlook exists; it must answer anywhere, never throw and never
        // block - the gate in front of every request takes this look.
        System.Diagnostics.Stopwatch clock = System.Diagnostics.Stopwatch.StartNew();
        OutlookStartObservation observation = OutlookStartObserver.Observe();
        OutlookStartVerdict verdict = OutlookStartPolicy.Decide(observation.ToFacts());
        clock.Stop();

        Assert.False(string.IsNullOrWhiteSpace(observation.Describe()));
        Assert.False(string.IsNullOrWhiteSpace(verdict.Reason));
        Assert.Equal(observation.Windows.State == OutlookLivenessState.NotRunning, observation.Windows.ProcessId == null);
        Assert.True(clock.Elapsed < System.TimeSpan.FromSeconds(5), $"the look took {clock.ElapsedMilliseconds} ms");
    }
}
