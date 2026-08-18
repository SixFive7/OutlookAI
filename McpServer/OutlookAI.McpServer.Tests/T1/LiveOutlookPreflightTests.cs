using OutlookAI.Core.Com;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the live tier's health gate: with Outlook wedged, the tier must refuse in
/// milliseconds instead of hanging.
/// <para>
/// This is the regression test for the 2026-08-18 incident, where fixture setup sat for 10
/// and then 15 minutes against an Outlook that could be neither started nor killed, and
/// the hand-abort that followed skipped the teardown sweep and left 7 tagged items in a
/// real mailbox.
/// </para>
/// <para>
/// Every case here is CI-safe and machine-independent, which is the point: the probe and
/// the clock are injected, so the refusal path, the settle window and the second opinion
/// are all provable without an unresponsive Outlook and without waiting out real time. The
/// two cases that do exercise the real entry point drive it through the same
/// <c>OUTLOOKAI_COMHOST_LIVENESS</c> override the COM host supervision tests use to force
/// an observed state.
/// </para>
/// </summary>
public sealed class LiveOutlookPreflightTests
{
    /// <summary>A probe that reads a fixed script, remembering how often it was asked.</summary>
    private sealed class ScriptedProbe
    {
        private readonly OutlookLivenessState[] _script;

        public ScriptedProbe(params OutlookLivenessState[] script)
        {
            _script = script;
        }

        public int Calls { get; private set; }

        public List<int> Waits { get; } = new();

        public int TotalWaitMilliseconds => Waits.Sum();

        public (OutlookLivenessState State, string Detail) Next()
        {
            // The last entry repeats, so "hung forever" is one element rather than a
            // guess at how many times the gate will look.
            OutlookLivenessState state = _script[Math.Min(Calls, _script.Length - 1)];
            Calls++;
            return (state, "scripted:" + state);
        }

        public void Record(int milliseconds)
        {
            Waits.Add(milliseconds);
        }
    }

    // ---------------------------------------------------------------- the decision

    [Fact]
    public void Responsive_Proceeds()
    {
        Assert.Equal(LivePreflightVerdict.Proceed, LiveOutlookPreflight.Decide(OutlookLivenessState.Responsive));
    }

    [Fact]
    public void Hung_Refuses()
    {
        // The whole point: no COM session opened, no fixture constructed, nothing waited on.
        Assert.Equal(LivePreflightVerdict.Refuse, LiveOutlookPreflight.Decide(OutlookLivenessState.Hung));
    }

    [Fact]
    public void Starting_Settles_RatherThanRefusing()
    {
        Assert.Equal(LivePreflightVerdict.Settle, LiveOutlookPreflight.Decide(OutlookLivenessState.Starting));
    }

    [Fact]
    public void NotRunning_Proceeds_BecauseTheFixturesAreAllowedToStartOutlook()
    {
        // The fixtures connect with allowStartingOutlook (S7/D17). Refusing here would
        // break every run on a machine where Outlook simply is not up yet.
        Assert.Equal(LivePreflightVerdict.Proceed, LiveOutlookPreflight.Decide(OutlookLivenessState.NotRunning));
    }

    [Fact]
    public void Decide_IsTotal()
    {
        foreach (OutlookLivenessState state in Enum.GetValues<OutlookLivenessState>())
        {
            Assert.Contains(
                LiveOutlookPreflight.Decide(state),
                new[] { LivePreflightVerdict.Proceed, LivePreflightVerdict.Settle, LivePreflightVerdict.Refuse });
        }
    }

    // ---------------------------------------------------------------- the gate

    [Fact]
    public void HealthyOutlook_CostsOneProbeAndNoWaiting()
    {
        // The gate runs before every live collection, so the healthy path has to be free.
        ScriptedProbe probe = new(OutlookLivenessState.Responsive);

        LiveOutlookPreflight.Require(probe.Next, probe.Record);

        Assert.Equal(1, probe.Calls);
        Assert.Empty(probe.Waits);
    }

    [Fact]
    public void WedgedOutlook_RefusesTheTier()
    {
        ScriptedProbe probe = new(OutlookLivenessState.Hung);

        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
            () => LiveOutlookPreflight.Require(probe.Next, probe.Record));

        // Two looks, one short pause between them, and nothing else spent.
        Assert.Equal(2, probe.Calls);
        Assert.Equal(new[] { LiveOutlookPreflight.ConfirmationDelayMilliseconds }, probe.Waits);
        Assert.Contains("REFUSING to run the live tier", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheRefusalSaysWhatIsWrong_WhyItMatters_AndWhatToDo()
    {
        // A quiet gate would be a worse outcome than the hang: whoever asked for a live run
        // has to learn from this message alone why they got no tests.
        string message = LiveOutlookPreflight.RefusalMessage("2 of 2 UI windows hung");

        Assert.Contains("REFUSING to run the live tier", message, StringComparison.Ordinal);
        Assert.Contains("2 of 2 UI windows hung", message, StringComparison.Ordinal);
        Assert.Contains("not responding", message, StringComparison.OrdinalIgnoreCase);

        // The consequence, so nobody reads this as a flaky gate and reruns it.
        Assert.Contains("sweep", message, StringComparison.OrdinalIgnoreCase);

        // The remedy, and the escape hatch.
        Assert.Contains("Task Manager", message, StringComparison.Ordinal);
        Assert.Contains(LiveOutlookPreflight.LivenessOverrideVariable, message, StringComparison.Ordinal);

        // House style for every user-visible string in this product: " - ", never an em
        // dash. Spelled as a code point so this file does not contain what it forbids.
        Assert.DoesNotContain((char)0x2014, message);
    }

    [Fact]
    public void AHungReadingThatClearsOnTheSecondLook_DoesNotRefuse()
    {
        // Confirm before crying, the same discipline the count tripwire applies to a
        // suspected loss: one long synchronous operation on the UI thread looks exactly
        // like a wedge for a moment, and a false refusal would make the gate untrustworthy.
        ScriptedProbe probe = new(OutlookLivenessState.Hung, OutlookLivenessState.Responsive);

        LiveOutlookPreflight.Require(probe.Next, probe.Record);

        Assert.Equal(2, probe.Calls);
    }

    [Fact]
    public void AStartingOutlookIsWaitedFor_ThenProceeds()
    {
        ScriptedProbe probe = new(
            OutlookLivenessState.Starting,
            OutlookLivenessState.Starting,
            OutlookLivenessState.Responsive);

        LiveOutlookPreflight.Require(probe.Next, probe.Record);

        Assert.Equal(3, probe.Calls);
        Assert.Equal(2 * LiveOutlookPreflight.SettlePollMilliseconds, probe.TotalWaitMilliseconds);
    }

    [Fact]
    public void AStartingOutlookThatWedges_IsCaughtWhenItSettles()
    {
        // The dangerous ordering: it looked like an ordinary cold start and turned out not
        // to be one. The settle window must not launder that into a proceed.
        ScriptedProbe probe = new(OutlookLivenessState.Starting, OutlookLivenessState.Hung);

        _ = Assert.Throws<InvalidOperationException>(
            () => LiveOutlookPreflight.Require(probe.Next, probe.Record));
    }

    [Fact]
    public void AnOutlookThatNeverFinishesStarting_ProceedsRatherThanRefusing_AndDoesNotSpin()
    {
        // A slow cold start is not a wedge, so this is not a refusal - but it must also not
        // poll forever, or the gate becomes the hang it was written to prevent.
        ScriptedProbe probe = new(OutlookLivenessState.Starting);

        LiveOutlookPreflight.Require(probe.Next, probe.Record);

        Assert.Equal(LiveOutlookPreflight.SettleBudgetMilliseconds, probe.TotalWaitMilliseconds);
        Assert.Equal(
            (LiveOutlookPreflight.SettleBudgetMilliseconds / LiveOutlookPreflight.SettlePollMilliseconds) + 1,
            probe.Calls);
    }

    [Fact]
    public void WithTheSettleWindowAlreadySpent_AStartingOutlookIsNotWaitedForAgain()
    {
        // Eight collections ask this gate, and each one asks the free question. Only the
        // first is allowed to spend the settle window, or the gate itself becomes a
        // multi-minute delay that a human would reasonably mistake for the hang.
        ScriptedProbe probe = new(OutlookLivenessState.Starting);

        LiveOutlookPreflight.Require(probe.Next, probe.Record, settleBudgetMilliseconds: 0);

        Assert.Equal(1, probe.Calls);
        Assert.Empty(probe.Waits);
    }

    [Fact]
    public void WithTheSettleWindowAlreadySpent_AWedgeIsStillRefused()
    {
        // The cheap half of the gate never lapses: Outlook can wedge between collections.
        ScriptedProbe probe = new(OutlookLivenessState.Hung);

        _ = Assert.Throws<InvalidOperationException>(
            () => LiveOutlookPreflight.Require(probe.Next, probe.Record, settleBudgetMilliseconds: 0));
    }

    [Fact]
    public void TheBudgetsAreProportionate()
    {
        // Long enough that a real cold start finishes inside it, short enough that a human
        // watching the run does not mistake the gate itself for the hang.
        Assert.InRange(LiveOutlookPreflight.SettleBudgetMilliseconds, 10_000, 120_000);
        Assert.InRange(LiveOutlookPreflight.SettlePollMilliseconds, 100, 5_000);

        // The second opinion must cost a moment, not a coffee break.
        Assert.InRange(LiveOutlookPreflight.ConfirmationDelayMilliseconds, 500, 10_000);
    }

    // ---------------------------------------------------------------- the override seam

    [Fact]
    public void ForcingHung_RefusesThroughTheRealProbe()
    {
        // Proves the shipped entry point - real probe, real clock - refuses, on a machine
        // whose own Outlook may be perfectly healthy or absent. Same seam the T3
        // supervision tests use to force an observed state.
        string? saved = Environment.GetEnvironmentVariable(LiveOutlookPreflight.LivenessOverrideVariable);
        try
        {
            Environment.SetEnvironmentVariable(
                LiveOutlookPreflight.LivenessOverrideVariable, OutlookLivenessState.Hung.ToString());

            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(LiveOutlookPreflight.Require);
            Assert.Contains("REFUSING to run the live tier", ex.Message, StringComparison.Ordinal);
            Assert.Contains(LiveOutlookPreflight.LivenessOverrideVariable, ex.Message, StringComparison.Ordinal);
        }
        finally
        {
            Environment.SetEnvironmentVariable(LiveOutlookPreflight.LivenessOverrideVariable, saved);
        }
    }

    [Fact]
    public void ForcingResponsive_LetsTheTierRunWhateverTheMachineLooksLike()
    {
        // The documented escape hatch: the message tells the operator to set exactly this,
        // so it has to work even where the real probe would refuse.
        string? saved = Environment.GetEnvironmentVariable(LiveOutlookPreflight.LivenessOverrideVariable);
        try
        {
            Environment.SetEnvironmentVariable(
                LiveOutlookPreflight.LivenessOverrideVariable, OutlookLivenessState.Responsive.ToString());

            LiveOutlookPreflight.Require();
        }
        finally
        {
            Environment.SetEnvironmentVariable(LiveOutlookPreflight.LivenessOverrideVariable, saved);
        }
    }

    [Theory]
    [InlineData("Hung", OutlookLivenessState.Hung)]
    [InlineData("hung", OutlookLivenessState.Hung)]
    [InlineData("RESPONSIVE", OutlookLivenessState.Responsive)]
    public void TheOverrideIsCaseInsensitive(string raw, OutlookLivenessState expected)
    {
        Assert.True(LiveOutlookPreflight.TryReadOverride(raw, out OutlookLivenessState state));
        Assert.Equal(expected, state);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("banana")]
    [InlineData("99")]
    public void AnUnrecognisedOverrideIsIgnored_RatherThanBecomingAState(string? raw)
    {
        // "99" is the one that matters: Enum.TryParse accepts any integer, so without the
        // defined-value check a typo would silently become an undefined liveness state and
        // the gate would be deciding on nonsense.
        Assert.False(LiveOutlookPreflight.TryReadOverride(raw, out _));
    }
}
