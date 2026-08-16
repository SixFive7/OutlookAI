using OutlookAI.ComHost.Supervision;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins every branch of the COM-host supervision policy.
/// <para>
/// These decisions are the load-bearing part of the fix for the 2026-08-15 hang, and
/// they are deliberately pure so they can be pinned here. The alternative - observing
/// them only through a genuinely wedged Outlook - is not reproducible, not safe on a
/// machine holding real mail, and impossible in CI. Same idiom as
/// <c>SweepWalkBoundsTests</c>: a total function over a synthetic clock.
/// </para>
/// </summary>
public sealed class ComHostPolicyTests
{
    // ---------------------------------------------------------------- DecideDispatch

    [Fact]
    public void Dispatch_ReadyHost_IsUsedDirectly()
    {
        DispatchVerdict verdict = ComHostPolicy.DecideDispatch(
            new DispatchInput(ComHostState.Ready, 0, long.MaxValue, StartingOutlookAllowed: true));

        Assert.Equal(DispatchVerdict.Dispatch, verdict);
    }

    // The state arrives as an int because ComHostState is internal - the supervision types
    // are not a supported surface - and a public [Theory] parameter cannot be less
    // accessible than the method carrying it.
    [Theory]
    [InlineData((int)ComHostState.None)]
    [InlineData((int)ComHostState.Starting)]
    [InlineData((int)ComHostState.Faulted)]
    public void Dispatch_WithoutReadyHost_StartsOne(int state)
    {
        DispatchVerdict verdict = ComHostPolicy.DecideDispatch(
            new DispatchInput((ComHostState)state, 0, long.MaxValue, StartingOutlookAllowed: true));

        Assert.Equal(DispatchVerdict.StartThenDispatch, verdict);
    }

    [Fact]
    public void Dispatch_ReadyHost_IsServedEvenWhenStartingOutlookIsForbidden()
    {
        // The D17 gate is about STARTING Outlook. A host that is already up requires
        // starting nothing, so refusing here would disable the tools during an add-in
        // update for no reason at all.
        DispatchVerdict verdict = ComHostPolicy.DecideDispatch(
            new DispatchInput(ComHostState.Ready, 0, long.MaxValue, StartingOutlookAllowed: false));

        Assert.Equal(DispatchVerdict.Dispatch, verdict);
    }

    [Fact]
    public void Dispatch_WhenStartingIsForbidden_RefusesRatherThanStarting()
    {
        DispatchVerdict verdict = ComHostPolicy.DecideDispatch(
            new DispatchInput(ComHostState.None, 0, long.MaxValue, StartingOutlookAllowed: false));

        Assert.Equal(DispatchVerdict.RefuseUnavailable, verdict);
    }

    [Fact]
    public void Dispatch_AfterRepeatedStartFailures_BacksOffInsteadOfThrashing()
    {
        DispatchVerdict verdict = ComHostPolicy.DecideDispatch(new DispatchInput(
            ComHostState.Faulted,
            ComHostPolicy.StartFailureBackoffThreshold,
            MillisecondsSinceLastStartFailure: 0,
            StartingOutlookAllowed: true));

        Assert.Equal(DispatchVerdict.RefuseBackoff, verdict);
    }

    [Fact]
    public void Dispatch_OnceTheBackoffWindowHasPassed_TriesAgain()
    {
        DispatchVerdict verdict = ComHostPolicy.DecideDispatch(new DispatchInput(
            ComHostState.Faulted,
            ComHostPolicy.StartFailureBackoffThreshold,
            ComHostPolicy.StartBackoffMilliseconds,
            StartingOutlookAllowed: true));

        Assert.Equal(DispatchVerdict.StartThenDispatch, verdict);
    }

    [Fact]
    public void Dispatch_BelowTheFailureThreshold_KeepsTrying()
    {
        DispatchVerdict verdict = ComHostPolicy.DecideDispatch(new DispatchInput(
            ComHostState.Faulted,
            ComHostPolicy.StartFailureBackoffThreshold - 1,
            MillisecondsSinceLastStartFailure: 0,
            StartingOutlookAllowed: true));

        Assert.Equal(DispatchVerdict.StartThenDispatch, verdict);
    }

    [Fact]
    public void Dispatch_ForbiddenOutranksBackoff()
    {
        // Both refusals apply; the "not allowed to start" reason is the more accurate one
        // to report, because it names something the user can wait out deterministically.
        DispatchVerdict verdict = ComHostPolicy.DecideDispatch(new DispatchInput(
            ComHostState.Faulted,
            ComHostPolicy.StartFailureBackoffThreshold,
            MillisecondsSinceLastStartFailure: 0,
            StartingOutlookAllowed: false));

        Assert.Equal(DispatchVerdict.RefuseUnavailable, verdict);
    }

    [Theory]
    [InlineData(0, 0, false)]
    [InlineData(2, 0, false)]
    [InlineData(3, 0, true)]
    [InlineData(3, 29_999, true)]
    [InlineData(3, 30_000, false)]
    [InlineData(9, 30_001, false)]
    public void StartBackoff_WindowBoundaries(int failures, long sinceMs, bool expected)
    {
        Assert.Equal(expected, ComHostPolicy.IsInStartBackoff(failures, sinceMs));
    }

    // ------------------------------------------------------------------ DecideBreaker

    [Fact]
    public void Breaker_StartsClosed()
    {
        Assert.Equal(BreakerVerdict.Closed, ComHostPolicy.DecideBreaker(new BreakerInput(0, long.MaxValue)));
    }

    [Fact]
    public void Breaker_ToleratesASingleTimeout()
    {
        // One timeout is not evidence of a wedged Outlook - a single slow operation, a
        // cold start behind a big OST, a transient stall. Opening on the first would make
        // the server refuse work it could have done.
        Assert.Equal(BreakerVerdict.Closed, ComHostPolicy.DecideBreaker(new BreakerInput(1, 0)));
    }

    [Fact]
    public void Breaker_OpensOnRepeatedTimeouts()
    {
        BreakerVerdict verdict = ComHostPolicy.DecideBreaker(
            new BreakerInput(ComHostPolicy.UnresponsiveTimeoutThreshold, 0));

        // Measured on a genuinely wedged Outlook: search, list_accounts and list_folders
        // each burned their full 120 s budget, independently, every time. Bounding each
        // call is necessary but not sufficient - the server has to remember.
        Assert.Equal(BreakerVerdict.Open, verdict);
    }

    [Fact]
    public void Breaker_StaysOpenForTheWholeCooldown()
    {
        Assert.Equal(
            BreakerVerdict.Open,
            ComHostPolicy.DecideBreaker(new BreakerInput(5, ComHostPolicy.UnresponsiveCooldownMilliseconds - 1)));
    }

    [Fact]
    public void Breaker_GoesHalfOpenWhenTheCooldownElapses()
    {
        // It must re-probe rather than latch: a user who restarts Outlook has to be picked
        // up automatically, without restarting the server.
        Assert.Equal(
            BreakerVerdict.HalfOpen,
            ComHostPolicy.DecideBreaker(new BreakerInput(5, ComHostPolicy.UnresponsiveCooldownMilliseconds)));
    }

    [Fact]
    public void Breaker_StaysHalfOpenIndefinitelyUntilSomethingSucceeds()
    {
        // Long after the cooldown it is still HalfOpen, never Open again by the passage of
        // time alone. Only an actual probe outcome moves it: success resets the count to
        // zero, failure re-stamps the clock.
        Assert.Equal(
            BreakerVerdict.HalfOpen,
            ComHostPolicy.DecideBreaker(new BreakerInput(99, ComHostPolicy.UnresponsiveCooldownMilliseconds * 100)));
    }

    [Fact]
    public void Breaker_CooldownIsShortEnoughToBeUnnoticeable()
    {
        // The cost of a stale-open breaker is a user waiting after fixing Outlook. Keep it
        // well under a minute, and far below the operation budget it exists to avoid.
        Assert.True(ComHostPolicy.UnresponsiveCooldownMilliseconds <= 60_000);
        Assert.True(ComHostPolicy.UnresponsiveCooldownMilliseconds < ComHostPolicy.DefaultOperationDeadlineMilliseconds);
    }

    // ---------------------------------------------------------------- DecideInFlight

    [Fact]
    public void InFlight_WithinBudget_KeepsWaiting()
    {
        InFlightVerdict verdict = ComHostPolicy.DecideInFlight(
            new InFlightInput(ElapsedMilliseconds: 500, DeadlineMilliseconds: 5_000, ChildAlive: true, ClientCancelled: false));

        Assert.Equal(InFlightVerdict.KeepWaiting, verdict);
    }

    [Fact]
    public void InFlight_AtTheDeadline_KillsTheHost()
    {
        // Exactly at the budget counts as exceeded: a deadline that only fires strictly
        // after would never fire at all for an operation that hangs precisely on it.
        InFlightVerdict verdict = ComHostPolicy.DecideInFlight(
            new InFlightInput(5_000, 5_000, ChildAlive: true, ClientCancelled: false));

        Assert.Equal(InFlightVerdict.TimeoutKillChild, verdict);
    }

    [Fact]
    public void InFlight_WhenTheHostIsGone_FailsWithoutKilling()
    {
        InFlightVerdict verdict = ComHostPolicy.DecideInFlight(
            new InFlightInput(10, 5_000, ChildAlive: false, ClientCancelled: false));

        Assert.Equal(InFlightVerdict.FailChildDied, verdict);
    }

    [Fact]
    public void InFlight_ClientCancellation_OutranksEverythingElse()
    {
        // Even past the deadline, and even with the host apparently gone: the response is
        // going to be suppressed by the SDK regardless, so there is nothing to gain by
        // killing a host that may be perfectly healthy.
        InFlightVerdict verdict = ComHostPolicy.DecideInFlight(
            new InFlightInput(999_999, 5_000, ChildAlive: false, ClientCancelled: true));

        Assert.Equal(InFlightVerdict.AbandonClientCancelled, verdict);
    }

    // ---------------------------------------------------------------- DeadlineFor

    [Fact]
    public void Deadline_ExplicitOverride_Wins()
    {
        Assert.Equal(1_234, ComHostPolicy.DeadlineFor(ComHostOperationClass.Operation, 1_234));
    }

    [Theory]
    [InlineData(0L)]
    [InlineData(-1L)]
    public void Deadline_NonPositiveOverride_FallsBackToTheDefault(long invalid)
    {
        // A zero deadline would mean "fail before starting", which is never what a caller
        // passing a bad value meant.
        Assert.Equal(
            ComHostPolicy.DefaultOperationDeadlineMilliseconds,
            ComHostPolicy.DeadlineFor(ComHostOperationClass.Operation, invalid));
    }

    [Fact]
    public void Deadline_HealthProbe_IsMuchShorterThanAnOperation()
    {
        long health = ComHostPolicy.DeadlineFor(ComHostOperationClass.HealthProbe, null);
        long operation = ComHostPolicy.DeadlineFor(ComHostOperationClass.Operation, null);

        // outlook_health is asked precisely when Outlook may be unresponsive. If it waited
        // the ordinary budget it would reproduce the very symptom it exists to report.
        Assert.True(health < operation, $"health probe {health} ms must be shorter than an operation {operation} ms");
        Assert.Equal(ComHostPolicy.HealthProbeDeadlineMilliseconds, health);
    }

    [Fact]
    public void Deadline_Connect_AllowsForAColdOutlookStart()
    {
        long connect = ComHostPolicy.DeadlineFor(ComHostOperationClass.Connect, null);

        Assert.Equal(ComHostPolicy.ConnectDeadlineMilliseconds, connect);
        Assert.True(connect > ComHostPolicy.HealthProbeDeadlineMilliseconds);
    }

    [Fact]
    public void Budgets_AreFiniteAndOrdered()
    {
        // The single property that matters most: every budget is finite. An infinite one
        // anywhere here is the 2026-08-15 hang coming back.
        Assert.True(ComHostPolicy.HealthProbeDeadlineMilliseconds > 0);
        Assert.True(ComHostPolicy.ConnectDeadlineMilliseconds > 0);
        Assert.True(ComHostPolicy.DefaultOperationDeadlineMilliseconds > 0);
        Assert.True(ComHostPolicy.StartBackoffMilliseconds > 0);
        Assert.True(ComHostPolicy.StartFailureBackoffThreshold > 0);
    }
}
