using OutlookAI.ComHost.Supervision;
using OutlookAI.Core.Com;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins how Outlook's externally observed state governs a request.
/// <para>
/// This is the cheapest and most valuable decision in the server: Windows already knows
/// whether Outlook is servicing its message queue, and answers in microseconds. Consulting
/// it first turned a 30-120 s discovery (measured against a genuinely wedged Outlook on
/// 2026-08-16) into a free one, and it is what lets a cold start return retry guidance
/// instead of blocking a caller for a minute and a half.
/// </para>
/// </summary>
public sealed class OutlookLivenessPolicyTests
{
    private const bool MayStartOutlook = true;
    private const bool MayNotStartOutlook = false;

    [Fact]
    public void Responsive_Proceeds()
    {
        Assert.Equal(
            LivenessVerdict.Proceed,
            ComHostPolicy.DecideLiveness(OutlookLivenessState.Responsive, long.MaxValue, MayStartOutlook));
    }

    [Fact]
    public void Hung_IsRefusedWithoutTryingCom()
    {
        // The whole point: no COM call, no budget spent, no child spawned.
        Assert.Equal(
            LivenessVerdict.Hung,
            ComHostPolicy.DecideLiveness(OutlookLivenessState.Hung, long.MaxValue, MayStartOutlook));
    }

    [Fact]
    public void Hung_IsRefusedEvenWhenStartingIsAllowed_AndWhenItIsNot()
    {
        // A hung Outlook is hung regardless of whether we would be permitted to start one;
        // starting a second Outlook is never the answer.
        Assert.Equal(
            LivenessVerdict.Hung,
            ComHostPolicy.DecideLiveness(OutlookLivenessState.Hung, 0, MayNotStartOutlook));
    }

    [Fact]
    public void Starting_AsksTheCallerToComeBack()
    {
        Assert.Equal(
            LivenessVerdict.Starting,
            ComHostPolicy.DecideLiveness(OutlookLivenessState.Starting, long.MaxValue, MayStartOutlook));
    }

    [Fact]
    public void NotRunning_MayStartWhenTheCooldownHasPassed()
    {
        Assert.Equal(
            LivenessVerdict.MayStart,
            ComHostPolicy.DecideLiveness(
                OutlookLivenessState.NotRunning, ComHostPolicy.AutostartCooldownMilliseconds, MayStartOutlook));
    }

    [Fact]
    public void NotRunning_StartIsSuppressedInsideTheCooldown()
    {
        // The anti-churn guard, and the reason it exists: the 2026-08-16 RCA found the
        // wedged Outlook was one WE started, 39 s after starting a previous one - almost
        // certainly activating Outlook while the prior instance was still exiting.
        Assert.Equal(
            LivenessVerdict.StartSuppressed,
            ComHostPolicy.DecideLiveness(
                OutlookLivenessState.NotRunning, ComHostPolicy.AutostartCooldownMilliseconds - 1, MayStartOutlook));
    }

    [Fact]
    public void NotRunning_IsNeverStartedWhenStartingIsForbidden()
    {
        Assert.Equal(
            LivenessVerdict.StartSuppressed,
            ComHostPolicy.DecideLiveness(OutlookLivenessState.NotRunning, long.MaxValue, MayNotStartOutlook));
    }

    [Fact]
    public void AutostartCooldown_IsLongEnoughToOutlastAnExitingOutlook()
    {
        // Too short and it does not cover the exiting-instance race it exists for; too
        // long and a legitimately closed Outlook takes ages to come back.
        Assert.True(ComHostPolicy.AutostartCooldownMilliseconds >= 10_000);
        Assert.True(ComHostPolicy.AutostartCooldownMilliseconds <= 60_000);
    }

    // -------------------------------------------------------------- retry guidance

    [Fact]
    public void Starting_CarriesUsableRetryGuidance()
    {
        int seconds = ComHostPolicy.RetryAfterSecondsFor(LivenessVerdict.Starting, 0);

        // Must be a real, actionable number: long enough that a retry has a chance, short
        // enough that an agent will actually wait rather than give up on the tool.
        Assert.InRange(seconds, 5, 60);
    }

    [Fact]
    public void StartSuppressed_CountsDownWithTheCooldown()
    {
        int early = ComHostPolicy.RetryAfterSecondsFor(LivenessVerdict.StartSuppressed, 0);
        int late = ComHostPolicy.RetryAfterSecondsFor(
            LivenessVerdict.StartSuppressed, ComHostPolicy.AutostartCooldownMilliseconds / 2);

        Assert.True(late < early, $"guidance must shrink as the wait elapses; {early}s -> {late}s");
        Assert.True(late >= 1, "never advise retrying in zero seconds - that invites a hot loop");
    }

    [Fact]
    public void StartSuppressed_NeverAdvisesZeroEvenAtTheBoundary()
    {
        int atBoundary = ComHostPolicy.RetryAfterSecondsFor(
            LivenessVerdict.StartSuppressed, ComHostPolicy.AutostartCooldownMilliseconds);

        Assert.True(atBoundary >= 1);
    }

    [Fact]
    public void Hung_AdvisesWaitingRatherThanHammering()
    {
        Assert.Equal(
            ComHostPolicy.UnresponsiveRetryAfterSeconds,
            ComHostPolicy.RetryAfterSecondsFor(LivenessVerdict.Hung, 0));
        Assert.True(ComHostPolicy.UnresponsiveRetryAfterSeconds >= 10);
    }

    [Fact]
    public void Proceed_CarriesNoRetryGuidance()
    {
        Assert.Equal(0, ComHostPolicy.RetryAfterSecondsFor(LivenessVerdict.Proceed, 0));
    }

    // -------------------------------------------------------------- the probe itself

    [Fact]
    public void Probe_IsTotalAndFastRegardlessOfMachineState()
    {
        // Runs in CI where Outlook does not exist, and on a dev box where it may be
        // running, starting or wedged. It must answer in all of them, never throw, and
        // never block - it is consulted on the hot path of every COM-needing request.
        System.Diagnostics.Stopwatch clock = System.Diagnostics.Stopwatch.StartNew();
        OutlookLivenessState state = OutlookLiveness.Probe(out string detail);
        clock.Stop();

        Assert.Contains(state, new[]
        {
            OutlookLivenessState.NotRunning,
            OutlookLivenessState.Starting,
            OutlookLivenessState.Responsive,
            OutlookLivenessState.Hung,
        });
        Assert.False(string.IsNullOrWhiteSpace(detail), "the probe must always explain what it saw");
        Assert.True(clock.Elapsed < System.TimeSpan.FromSeconds(5), $"probe took {clock.ElapsedMilliseconds} ms");
    }

    [Fact]
    public void Describe_CoversEveryState()
    {
        foreach (OutlookLivenessState state in System.Enum.GetValues<OutlookLivenessState>())
        {
            string described = OutlookLiveness.Describe(state);
            Assert.False(string.IsNullOrWhiteSpace(described));
            Assert.NotEqual("unknown", described);
        }
    }
}
