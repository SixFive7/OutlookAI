using System.Diagnostics;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// How long a live-tier poll may keep waiting, measured on a clock that only goes forward.
/// <para>
/// Every wait in the live tier used to be <c>DateTime deadline = DateTime.UtcNow + timeout;
/// while (DateTime.UtcNow &lt; deadline)</c>. That is a DURATION measured on the wall clock,
/// and the wall clock jumps: an NTP correction, a person setting the time, a VM resuming
/// from a snapshot. A backwards jump extends every live wait by the size of the jump - on a
/// suite whose longest waits are three minutes each and which has already lost a night to a
/// test that "was still going" after 22.5 minutes, that is the exact failure nobody could
/// tell apart from a hang. A forwards jump does the opposite and ends a wait for real mail
/// before the mail could possibly have arrived, which reads as a flaky test.
/// </para>
/// <para>
/// The same rule as <c>OutlookAI.Core.Services.MonotonicClock</c>, one layer up: a wall
/// clock is right for RECORDING OR COMPARING AN ABSOLUTE INSTANT - the send time a sweep
/// window is built from, a screenshot filename - and wrong for measuring how long something
/// has been going. Those absolute uses stay on <see cref="DateTime.UtcNow"/> deliberately.
/// </para>
/// </summary>
public readonly struct LiveWaitBudget
{
    private readonly long _startTimestamp;

    private LiveWaitBudget(TimeSpan budget)
    {
        Budget = budget;
        _startTimestamp = Stopwatch.GetTimestamp();
    }

    /// <summary>How long this wait was given.</summary>
    public TimeSpan Budget { get; }

    /// <summary>How long it has been waiting. Monotonic; never negative, never jumps.</summary>
    public TimeSpan Elapsed => Stopwatch.GetElapsedTime(_startTimestamp);

    /// <summary>True while the budget has not run out - the loop condition.</summary>
    public bool HasTimeLeft => Elapsed < Budget;

    /// <summary>Starts a wait of <paramref name="budget"/>.</summary>
    public static LiveWaitBudget Of(TimeSpan budget)
    {
        return new LiveWaitBudget(budget);
    }

    /// <summary>Starts a wait of <paramref name="seconds"/> seconds.</summary>
    public static LiveWaitBudget OfSeconds(double seconds)
    {
        return new LiveWaitBudget(TimeSpan.FromSeconds(seconds));
    }
}
