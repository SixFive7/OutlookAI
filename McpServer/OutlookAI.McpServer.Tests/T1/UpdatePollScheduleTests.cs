using System;

using OutlookAI.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The add-in updater's poll schedule: how long it waits before its first check, and how far
/// it backs off while it cannot reach GitHub.
///
/// <para>
/// WHY THIS IS TESTABLE AT ALL. <c>UpdateService</c> is an <c>HttpClient</c>, a
/// <c>System.Threading.Timer</c> and Outlook's process lifetime, none of which this host has,
/// and it is net48/VSTO so it cannot be referenced from here. The schedule was split into
/// <c>Services\UpdatePollSchedule.cs</c> precisely so the one part with a right answer could
/// be pinned; that file is LINKED into this project (see the csproj), so what runs below is
/// the code the add-in ships and not a re-implementation of it.
/// </para>
///
/// <para>
/// WHAT WOULD BE LOST WITHOUT IT. The old behaviour was a <c>TimeSpan.Zero</c> due time and no
/// backoff: an offline machine asked GitHub 144 times a day, for ever. Both halves of the fix
/// are numbers nobody can check by reading - a curve that is one doubling too shallow still
/// looks like a backoff, and a cap two hours too high still looks bounded - and neither has a
/// visible symptom on a healthy machine, which is every machine anyone tests on.
/// </para>
/// </summary>
public sealed class UpdatePollScheduleTests
{
    /// <summary>
    /// The first check is delayed rather than fired inside add-in load, and the delay is a real
    /// interval rather than a token one. The upper bound matters as much as the lower: this
    /// delays every user's first update check, so it may not creep into minutes.
    /// </summary>
    [Fact]
    public void FirstCheckIsDelayedByAboutHalfAMinute()
    {
        Assert.True(UpdatePollSchedule.SettleDelay > TimeSpan.Zero,
            "a zero due time is the defect: the check then runs inside Outlook's own startup.");
        Assert.Equal(TimeSpan.FromSeconds(30), UpdatePollSchedule.SettleDelay);
        Assert.True(UpdatePollSchedule.SettleDelay < UpdatePollSchedule.BaseInterval,
            "the settle delay is a pause before the first check, not a second poll interval.");
    }

    /// <summary>
    /// A healthy machine polls at the base interval - the case that must not change, because
    /// the tooltip, the README and <c>PollIntervalDescription</c> all describe it.
    /// </summary>
    [Fact]
    public void SuccessPollsAtTheBaseInterval()
    {
        Assert.Equal(TimeSpan.FromMinutes(10), UpdatePollSchedule.BaseInterval);
        Assert.Equal(UpdatePollSchedule.BaseInterval, UpdatePollSchedule.DelayAfter(0));
        Assert.False(UpdatePollSchedule.IsBackingOff(0));
    }

    /// <summary>
    /// The tolerance window. A blip - a Wi-Fi handoff, a VPN reconnect, one GitHub 5xx - must
    /// not slow the poll down at all, or a transient failure costs the user a late update.
    /// </summary>
    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    public void FailuresBelowTheThresholdDoNotSlowThePollDown(int failures)
    {
        Assert.Equal(UpdatePollSchedule.BaseInterval, UpdatePollSchedule.DelayAfter(failures));
        Assert.False(UpdatePollSchedule.IsBackingOff(failures));
    }

    /// <summary>
    /// THE CURVE ITSELF, spelled out rather than recomputed from the constants - a test that
    /// re-derives the answer from the same expression the code uses cannot fail. These are the
    /// minutes an unreachable machine actually waits: 10, 10, 20, 40, 80, then the two-hour
    /// ceiling for ever.
    /// </summary>
    [Theory]
    [InlineData(3, 20)]
    [InlineData(4, 40)]
    [InlineData(5, 80)]
    [InlineData(6, 120)]
    [InlineData(7, 120)]
    [InlineData(50, 120)]
    public void BackoffDoublesThenStopsAtTheCeiling(int failures, int expectedMinutes)
    {
        Assert.Equal(TimeSpan.FromMinutes(expectedMinutes), UpdatePollSchedule.DelayAfter(failures));
        Assert.True(UpdatePollSchedule.IsBackingOff(failures));
    }

    /// <summary>
    /// Backoff starts at the third consecutive failure and not before - the threshold itself,
    /// asserted at its two adjacent counts so moving it by one fails here.
    /// </summary>
    [Fact]
    public void BackoffStartsAtTheThirdConsecutiveFailure()
    {
        Assert.Equal(3, UpdatePollSchedule.FailuresBeforeBackoff);
        Assert.Equal(UpdatePollSchedule.BaseInterval,
            UpdatePollSchedule.DelayAfter(UpdatePollSchedule.FailuresBeforeBackoff - 1));
        Assert.True(UpdatePollSchedule.DelayAfter(UpdatePollSchedule.FailuresBeforeBackoff)
            > UpdatePollSchedule.BaseInterval);
    }

    /// <summary>
    /// The curve never goes backwards and never exceeds the ceiling, at any count. Monotonicity
    /// is what makes "consecutive failures" a meaningful input at all, and the ceiling is the
    /// bound on how long a machine that has quietly come back can sit unaware.
    /// </summary>
    [Fact]
    public void TheCurveIsMonotonicAndBounded()
    {
        var previous = TimeSpan.Zero;
        for (int failures = 0; failures <= 200; failures++)
        {
            var delay = UpdatePollSchedule.DelayAfter(failures);
            Assert.True(delay >= previous, $"delay went backwards at {failures} failures.");
            Assert.True(delay <= UpdatePollSchedule.MaxInterval,
                $"delay exceeded the ceiling at {failures} failures.");
            previous = delay;
        }

        Assert.Equal(TimeSpan.FromHours(2), UpdatePollSchedule.MaxInterval);
        Assert.Equal(UpdatePollSchedule.MaxInterval, UpdatePollSchedule.DelayAfter(int.MaxValue));
    }

    /// <summary>
    /// A negative count decides a timer's due time, so it may not throw and it may not produce
    /// a negative or zero interval - <c>Timer.Change</c> treats those as "fire immediately" or
    /// throws, and either would turn a bookkeeping slip into a hot loop against GitHub.
    /// </summary>
    [Fact]
    public void ANegativeCountFallsBackToTheBaseInterval()
    {
        Assert.Equal(UpdatePollSchedule.BaseInterval, UpdatePollSchedule.DelayAfter(-1));
        Assert.Equal(UpdatePollSchedule.BaseInterval, UpdatePollSchedule.DelayAfter(int.MinValue));
    }

    /// <summary>
    /// THE ACTUAL POINT OF THE BACKOFF, stated as the quantity it was meant to change: how many
    /// requests a permanently unreachable machine makes in a day. It was 144. The bound below is
    /// deliberately loose - it is checking an order of magnitude, not the curve, which the cases
    /// above already pin.
    /// </summary>
    [Fact]
    public void AnOfflineMachineStopsHammeringTheServer()
    {
        var elapsed = UpdatePollSchedule.SettleDelay;
        int checks = 0;
        while (elapsed < TimeSpan.FromHours(24))
        {
            checks++;
            elapsed += UpdatePollSchedule.DelayAfter(checks);
        }

        Assert.InRange(checks, 5, 20);
    }
}
