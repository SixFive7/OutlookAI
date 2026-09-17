using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the STA cancellation PROTOCOL - the half of the corpus tool's threading that can be
/// tested without Outlook.
/// <para>
/// <b>What went wrong, and what these tests hold shut.</b> The old runner threw
/// <c>TimeoutException</c> the moment <c>Thread.Join(timeout)</c> returned false and returned
/// to the caller while the thread kept running. It was a background thread, so nothing joined
/// it and nothing cancelled it: it went on holding its Outlook references and driving the
/// store until the process exited. For the placement and date probes - which CREATE AND
/// DELETE MAIL - that is an abandoned COM session mutating a store the caller has been told
/// it stopped touching, and a caller that retries then has two writers against one store. It
/// fired twice on the measurement guest on 2026-09-16, both times during a cold Outlook
/// start, which is the case that is slow rather than stuck.
/// </para>
/// <para>
/// <b>The COM half cannot be tested here and is not pretended to be.</b> What is pinned is
/// everything that decides behaviour: that a signalled run stops at a loop boundary and not
/// mid-item, that the runner WAITS and reports acknowledged and not-acknowledged as different
/// outcomes, that a stopped write says how far it got, and that a cold start is judged by its
/// own allowance rather than by the work bound.
/// </para>
/// <para>
/// <b><see cref="Stop_NotAcknowledgedWhenTheBodyNeverChecks"/> is the control.</b> It runs a
/// body with no cancellation check - the defect, reproduced deliberately - and requires the
/// runner to report NOT ACKNOWLEDGED. Its sibling
/// <see cref="Stop_AcknowledgedAtTheLoopBoundaryAndSaysHowFarItGot"/> differs only by having
/// the check, so neither can pass vacuously: delete the check from
/// <see cref="ComStaCheckpoint.Steps{T}"/> and the sibling flips to this one's verdict and
/// fails.
/// </para>
/// </summary>
public class ComStaCancellationTests
{
    /// <summary>
    /// A bound wide enough that no scheduling hiccup can trip it, for the tests that are not
    /// about expiry.
    /// </summary>
    private static readonly TimeSpan Generous = TimeSpan.FromSeconds(30);

    /// <summary>
    /// The silence a test deliberately lets happen. One second, because the runner polls every
    /// 250 ms and the bodies below report their first steps in microseconds: no realistic
    /// pause can get between them and the bound.
    /// </summary>
    private static readonly TimeSpan ShortSilence = TimeSpan.FromSeconds(1);

    private static ComStaBudget Budget(TimeSpan? startup, TimeSpan? work, TimeSpan? grace = null)
        => new(startup, work, grace ?? TimeSpan.FromSeconds(30));

    // ------------------------------------------------------------------ the happy path

    [Fact]
    public void Run_ReturnsWhatTheWorkReturned()
    {
        int answer = ComStaRunner.Run(
            "test operation",
            Budget(Generous, Generous),
            checkpoint =>
            {
                checkpoint.OutlookBound("Test Profile");
                int seen = 0;
                foreach (int item in checkpoint.Steps(Enumerable.Range(1, 10), "item"))
                {
                    seen += item;
                }

                return seen;
            });

        Assert.Equal(55, answer);
    }

    [Fact]
    public void Run_RunsTheWorkOnAnStaThread()
    {
        // The whole reason this class exists rather than a Task: Outlook's object model is
        // apartment-threaded, and a runner that quietly stopped setting STA would work in the
        // tests and fail on the one machine that matters.
        ApartmentState state = ComStaRunner.Run(
            "test operation",
            Budget(Generous, Generous),
            _ => Thread.CurrentThread.GetApartmentState());

        Assert.Equal(ApartmentState.STA, state);
    }

    [Fact]
    public void Run_SurfacesTheWorkFailureWithItsCause()
    {
        var thrown = new InvalidTimeZoneException("the store said no");
        InvalidOperationException wrapper = Assert.Throws<InvalidOperationException>(
            () => ComStaRunner.Run<int>(
                "test operation",
                Budget(Generous, Generous),
                _ => throw thrown));

        Assert.Same(thrown, wrapper.InnerException);
        Assert.Contains("test operation", wrapper.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Run_RecordsTheProfileTheRunBound()
    {
        // The name is carried on the checkpoint because ONE place reads it - the same place
        // that ends the cold-start phase - and two things need it: this message, and the
        // corpus refusal that used to name only the store.
        string? bound = ComStaRunner.Run(
            "test operation",
            Budget(Generous, Generous),
            checkpoint =>
            {
                checkpoint.OutlookBound("Outlook Measurement Profile");
                return checkpoint.ProfileName;
            });

        Assert.Equal("Outlook Measurement Profile", bound);
    }

    // ------------------------------------------------------------------ stopping

    [Fact]
    public void Stop_AcknowledgedAtTheLoopBoundaryAndSaysHowFarItGot()
    {
        // The body writes three items, then simulates a COM call that does not return
        // promptly - no steps are reported while it runs, which is exactly the silence the
        // work bound is looking for. When it does return, the loop's own enumerator refuses
        // the fourth item.
        int written = 0;
        var stalled = new ManualResetEventSlim(false);

        ComStaTimeoutException expiry = Assert.Throws<ComStaTimeoutException>(
            () => ComStaRunner.Run(
                "corpus build",
                Budget(Generous, ShortSilence),
                checkpoint =>
                {
                    checkpoint.OutlookBound("Test Profile");
                    foreach (int item in checkpoint.Steps(Enumerable.Range(1, 1000), "build item"))
                    {
                        written++;
                        if (item == 3)
                        {
                            stalled.Set();
                            SpinUntilStopping(checkpoint);
                        }
                    }

                    return written;
                }));

        Assert.True(stalled.IsSet, "the body never reached the simulated stall");
        Assert.True(expiry.Acknowledged, expiry.Message);
        Assert.Equal(ComStaPhase.Working, expiry.Phase);

        // How far it got. THREE, not 1 000: the loop stopped at the boundary after the item it
        // had finished, which for the corpus build is the point the manifest already describes.
        Assert.Equal(3, written);
        Assert.Contains("3 'build item' step(s)", expiry.Progress, StringComparison.Ordinal);
        Assert.Contains("3 'build item' step(s)", expiry.Message, StringComparison.Ordinal);
        Assert.Contains("ACKNOWLEDGED", expiry.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("DID NOT ACKNOWLEDGE", expiry.Message, StringComparison.Ordinal);
        Assert.Contains("corpus build", expiry.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Stop_NotAcknowledgedWhenTheBodyNeverChecks()
    {
        // THE CONTROL. This body is the old defect: it takes the work, goes quiet, and never
        // asks whether it may continue. The runner must not describe that as a clean stop -
        // it is the state where a COM session really is still live against the store.
        var release = new ManualResetEventSlim(false);
        try
        {
            ComStaTimeoutException expiry = Assert.Throws<ComStaTimeoutException>(
                () => ComStaRunner.Run(
                    "corpus placement probe",
                    Budget(Generous, ShortSilence, grace: TimeSpan.FromMilliseconds(400)),
                    checkpoint =>
                    {
                        checkpoint.OutlookBound("Test Profile");
                        checkpoint.Step("placement rung");
                        release.Wait();
                        return 0;
                    }));

            Assert.False(expiry.Acknowledged, expiry.Message);
            Assert.Contains("DID NOT ACKNOWLEDGE", expiry.Message, StringComparison.Ordinal);
            Assert.Contains("STILL RUNNING", expiry.Message, StringComparison.Ordinal);
            Assert.Contains("STILL HOLDS OUTLOOK REFERENCES", expiry.Message, StringComparison.Ordinal);

            // The operator instruction is part of the contract, not decoration: a retry while
            // the abandoned session lives is two writers against one store.
            Assert.Contains("Do NOT re-run", expiry.Message, StringComparison.Ordinal);
            Assert.Contains("1 'placement rung' step(s)", expiry.Message, StringComparison.Ordinal);
        }
        finally
        {
            // Never leave the thread parked: it is a background thread, so it would outlive
            // this test exactly the way the defect outlived its caller.
            release.Set();
        }
    }

    [Fact]
    public void Stop_ReportsTheBoundsAndHowToWidenThem()
    {
        var release = new ManualResetEventSlim(false);
        try
        {
            ComStaTimeoutException expiry = Assert.Throws<ComStaTimeoutException>(
                () => ComStaRunner.Run<int>(
                    "corpus scan",
                    Budget(Generous, ShortSilence, grace: TimeSpan.FromMilliseconds(400)),
                    checkpoint =>
                    {
                        checkpoint.OutlookBound("Test Profile");
                        release.Wait();
                        return 0;
                    }));

            Assert.Contains(ComStaBudget.StartupVariable, expiry.Message, StringComparison.Ordinal);
            Assert.Contains(ComStaBudget.WorkVariable, expiry.Message, StringComparison.Ordinal);
            Assert.Contains("0 removes the bound", expiry.Message, StringComparison.Ordinal);
            Assert.Contains("work silence 1s", expiry.Message, StringComparison.Ordinal);
        }
        finally
        {
            release.Set();
        }
    }

    // ------------------------------------------------------------------ the cold start

    [Fact]
    public void ColdStart_IsJudgedByItsOwnAllowanceAndSaysSo()
    {
        // A run that never reaches OutlookBound is still starting Outlook. The old flat bound
        // could not say that, and "Corpus STA operation timed out" sent an operator looking
        // for a stuck write that had never begun.
        var release = new ManualResetEventSlim(false);
        try
        {
            ComStaTimeoutException expiry = Assert.Throws<ComStaTimeoutException>(
                () => ComStaRunner.Run<int>(
                    "corpus store facts",
                    Budget(ShortSilence, Generous, grace: TimeSpan.FromMilliseconds(400)),
                    _ =>
                    {
                        release.Wait();
                        return 0;
                    }));

            Assert.Equal(ComStaPhase.StartingOutlook, expiry.Phase);
            Assert.Contains("still starting Outlook", expiry.Message, StringComparison.Ordinal);
            Assert.Contains("never got past starting Outlook", expiry.Message, StringComparison.Ordinal);
        }
        finally
        {
            release.Set();
        }
    }

    [Fact]
    public void ColdStart_IsNotChargedToTheWorkBound()
    {
        // The direction the maintainer accepted: a cold start is not a hang. Here the start
        // takes five times the work bound and the run still succeeds, because the work clock
        // does not begin until Outlook has answered - the same rule the product side keeps in
        // BudgetedSessionProxy.
        TimeSpan coldStart = TimeSpan.FromMilliseconds(1000);
        TimeSpan workBound = TimeSpan.FromMilliseconds(200);

        int answer = ComStaRunner.Run(
            "corpus store facts",
            Budget(Generous, workBound),
            checkpoint =>
            {
                Thread.Sleep(coldStart);
                checkpoint.OutlookBound("Test Profile");
                foreach (int item in checkpoint.Steps(Enumerable.Range(1, 3), "store fact"))
                {
                    _ = item;
                }

                return 42;
            });

        Assert.Equal(42, answer);
    }

    // ------------------------------------------------------------------ the checkpoint itself

    [Fact]
    public void Steps_YieldsEverythingWhileTheRunIsNotStopping()
    {
        List<int> seen = ComStaRunner.Run(
            "test operation",
            Budget(Generous, Generous),
            checkpoint =>
            {
                checkpoint.OutlookBound(null);
                return checkpoint.Steps(Enumerable.Range(1, 7), "item").ToList();
            });

        Assert.Equal(new[] { 1, 2, 3, 4, 5, 6, 7 }, seen);
    }

    [Fact]
    public void Step_AnswersYesWhileTheRunIsNotStopping()
    {
        bool mayContinue = ComStaRunner.Run(
            "test operation",
            Budget(Generous, Generous),
            checkpoint => checkpoint.Step("row") && !checkpoint.Stopping);

        Assert.True(mayContinue);
    }

    [Fact]
    public void Checkpoint_RefusesAnUnlabelledStep()
    {
        // The label is what the progress report is made of. A blank one would turn "it got as
        // far as 1 234 'build item' steps" into a number with nothing attached.
        Assert.Throws<InvalidOperationException>(
            () => ComStaRunner.Run(
                "test operation",
                Budget(Generous, Generous),
                checkpoint => checkpoint.Step("  ")));
    }

    // ------------------------------------------------------------------ the budget

    [Fact]
    public void Budget_UsesTheDocumentedDefaultsWhenNothingIsConfigured()
    {
        ComStaBudget budget = ComStaBudget.Resolve(TimeSpan.FromMinutes(10), null, null);
        Assert.Equal(ComStaBudget.DefaultStartup, budget.Startup);
        Assert.Equal(TimeSpan.FromMinutes(10), budget.Work);
        Assert.Equal(ComStaBudget.DefaultGrace, budget.Grace);

        // Ten minutes, and the number is not arbitrary - see ComStaBudget.DefaultStartup for
        // where it comes from. Pinned because a silently shrunk cold-start allowance would
        // reintroduce exactly the failure this work removed.
        Assert.Equal(TimeSpan.FromMinutes(10), ComStaBudget.DefaultStartup);
        Assert.Equal(TimeSpan.FromSeconds(60), ComStaBudget.DefaultGrace);
    }

    [Fact]
    public void Budget_KeepsAnUnboundedWorkBoundUnbounded()
    {
        // The corpus BUILD passes null on purpose: a build of tens of thousands of items runs
        // for hours and must not be abandoned mid-write. It still gets the cold-start bound.
        ComStaBudget budget = ComStaBudget.Resolve(null, null, null);
        Assert.Null(budget.Work);
        Assert.Equal(ComStaBudget.DefaultStartup, budget.Startup);
    }

    [Theory]
    [InlineData("60000", 60_000)]
    [InlineData("1", 1)]
    public void Budget_TakesAConfiguredValueAsMilliseconds(string raw, int expectedMs)
    {
        ComStaBudget budget = ComStaBudget.Resolve(TimeSpan.FromMinutes(10), raw, raw);
        Assert.Equal(TimeSpan.FromMilliseconds(expectedMs), budget.Startup);
        Assert.Equal(TimeSpan.FromMilliseconds(expectedMs), budget.Work);
    }

    [Fact]
    public void Budget_ZeroMeansUnbounded()
    {
        // The escape hatch for a machine where a cold start really does take longer than
        // anyone has budgeted for. It has to be reachable without a rebuild: the alternative
        // an operator reaches for otherwise is to stop using the bound at all.
        ComStaBudget budget = ComStaBudget.Resolve(TimeSpan.FromMinutes(10), "0", "0");
        Assert.Null(budget.Startup);
        Assert.Null(budget.Work);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("later")]
    [InlineData("-1")]
    [InlineData("10.5")]
    public void Budget_IgnoresAnUnusableOverrideAndKeepsTheDefault(string raw)
    {
        ComStaBudget budget = ComStaBudget.Resolve(TimeSpan.FromMinutes(3), raw, raw);
        Assert.Equal(ComStaBudget.DefaultStartup, budget.Startup);
        Assert.Equal(TimeSpan.FromMinutes(3), budget.Work);
    }

    [Theory]
    [InlineData(null, "none")]
    [InlineData(500, "0.5s")]
    [InlineData(45_000, "45s")]
    [InlineData(600_000, "10m")]
    [InlineData(90_000, "1m30s")]
    public void Budget_FormatsABoundTheWayAnOperatorReadsIt(int? milliseconds, string expected)
    {
        TimeSpan? bound = milliseconds == null ? null : TimeSpan.FromMilliseconds(milliseconds.Value);
        Assert.Equal(expected, ComStaBudget.Format(bound));
    }

    [Fact]
    public void Budget_DescribesAllThreeNumbers()
    {
        string described = new ComStaBudget(
            TimeSpan.FromMinutes(10), TimeSpan.FromMinutes(3), TimeSpan.FromSeconds(60)).Describe();
        Assert.Equal("cold start 10m, work silence 3m, grace 1m", described);
    }

    /// <summary>
    /// Stands in for a COM call that does not return promptly: it reports nothing while it
    /// runs, which is what the work bound is watching for, and it gives up eventually so a
    /// failing test cannot hang the suite.
    /// </summary>
    private static void SpinUntilStopping(ComStaCheckpoint checkpoint)
    {
        DateTime giveUp = DateTime.UtcNow.AddSeconds(30);
        while (!checkpoint.Stopping && DateTime.UtcNow < giveUp)
        {
            Thread.Sleep(5);
        }
    }
}
