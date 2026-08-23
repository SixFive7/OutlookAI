using System.Reflection;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins what the count tripwire does about a SUSPECTED loss: re-census first, re-run only if
/// it persists, both bounded, and every path says out loud what it did.
/// <para>
/// These tests exist because the retry is the dangerous half of the guard. The tripwire stands
/// between the suite and the incident that once destroyed real mail, and <b>every retry is a
/// chance to convert a real loss into a pass</b>. So the assertions are not only about the
/// verdict: they pin the number of censuses actually taken, the number of re-runs, and the
/// text the run reports - because a retry nobody can see in the output is the same thing as no
/// guard at all.
/// </para>
/// <para>
/// The census source is a fake. No COM, no Outlook, no mailbox, no wall clock: the ladder is a
/// pure policy and this is the whole point of its shape.
/// </para>
/// </summary>
public sealed class TripwireRetryLadderTests
{
    private const string LostFromInbox = "items-removed|other@example.test|Inbox";
    private const string LostFromArchive = "items-removed|other@example.test|Archive";
    private const string FolderGone = "folder-removed|other@example.test|2019";

    /// <summary>One comparison's worth of failures, keyed the way the tripwire keys them.</summary>
    private static TripwireVerdict Verdict(params string[] keys)
    {
        List<TripwireFailure> failures = keys
            .Select(key => new TripwireFailure(key, "  " + key + ": items left a mailbox the suite may not touch."))
            .ToList();
        return new TripwireVerdict(
            failures,
            new[] { "  churn: mail arrived somewhere else." },
            failures.Count > 0 ? "ATTRIBUTION: undecidable." : null);
    }

    [Fact]
    public void ADeltaThatIsGoneOnTheFirstReCensus_PassesAndSaysItRetried()
    {
        // Two censuses agree that nothing persisted, so the ladder stops there: the second
        // re-census is never taken, and the fake would throw if it were.
        FakeCensusSource source = new(Verdict());

        TripwireRetryReport report = TripwireRetryLadder.Resolve(
            Verdict(LostFromInbox, FolderGone), source);

        Assert.False(report.Failed);
        Assert.Empty(report.Confirmed);
        Assert.Equal(1, source.ReCensuses);
        Assert.Equal(0, source.ReRuns);
        Assert.False(report.SurvivedARecensus);
        Assert.False(report.BoundReached);

        // ...and it is not a silent pass. A run that needed a second reading must never look
        // like one that did not.
        Assert.True(report.Entered);
        Assert.Equal("retry: 1 re-census(es), 0 re-run(s), verdict PASSED", report.Summary);
        Assert.Contains("RETRIED AND PASSED", report.Describe(), StringComparison.Ordinal);
        Assert.Contains("post-run census: 2 suspected failure(s)", report.Describe(), StringComparison.Ordinal);
        Assert.Contains("re-census 1 of 2", report.Describe(), StringComparison.Ordinal);
        Assert.Contains("cleared: " + LostFromInbox, report.Describe(), StringComparison.Ordinal);
        Assert.Contains("two censuses agree", report.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void ADeltaThatSurvivesOneReCensusAndThenClears_ReadsAsAPassAndStillFailsTheRun()
    {
        // The bounds are ceilings, not targets - but a reading that survived one census is no
        // longer noise. The ladder's VERDICT is a pass and the RUN fails, which is the whole
        // point of PassedWithASurvivedDelta: the items really were gone at two readings.
        FakeCensusSource source = new(Verdict(LostFromInbox), Verdict());

        TripwireRetryReport report = TripwireRetryLadder.Resolve(
            Verdict(LostFromInbox, FolderGone), source);

        Assert.Equal(TripwireRunOutcome.PassedWithASurvivedDelta, report.Outcome);
        Assert.True(report.Failed);
        Assert.Empty(report.Confirmed);
        Assert.Equal(2, source.ReCensuses);
        Assert.Equal(0, source.ReRuns);
        Assert.True(report.SurvivedARecensus);
        Assert.Equal(
            "retry: 2 re-census(es), 0 re-run(s), verdict PASSED WITH A SURVIVED DELTA (fails the run)",
            report.Summary);
        Assert.Contains("PASSED WITH A SURVIVED DELTA", report.Describe(), StringComparison.Ordinal);
        Assert.Contains("STILL EXITS NON-ZERO", report.Describe(), StringComparison.Ordinal);
        Assert.Contains("PERSISTED 1 of 2", report.Describe(), StringComparison.Ordinal);
        Assert.Contains("cleared: " + FolderGone, report.Describe(), StringComparison.Ordinal);
        Assert.Contains("persisted 0 of 1", report.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void ADeltaThatSurvivesBothReCensuses_SpendsExactlyTwoAndEscalates()
    {
        FakeCensusSource source = new(Verdict(LostFromInbox), Verdict(LostFromInbox));

        TripwireRetryReport report = TripwireRetryLadder.Resolve(Verdict(LostFromInbox), source);

        // Exactly two, then the re-run rung. Not three, not one.
        Assert.Equal(2, source.ReCensuses);
        Assert.Equal(new[] { 1, 2 }, source.ReCensusAttempts);
        Assert.Equal(1, source.ReRuns);
        Assert.True(report.BoundReached);
        Assert.True(report.Failed);
        Assert.Equal(LostFromInbox, Assert.Single(report.Confirmed).Key);
        Assert.Contains("BOUND REACHED", report.Describe(), StringComparison.Ordinal);
        Assert.Contains("escalation, not a give-up", report.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheCappedReRun_ClearsTheDeltaWhenItDoesNotReproduce_AndTheRunStillFails()
    {
        // The one exoneration the ladder can actually make: a person's one-off deletion does
        // not come back when the implicated tests run again, while a test that deletes does.
        // THE MAINTAINER'S DECISION: it clears the ladder's verdict and not the run. This is
        // the exact shape where a real loss could otherwise end green - the delta survived two
        // censuses, so the items are gone, and the re-run only says the suite is not who took
        // them.
        FakeCensusSource source = new(Verdict(LostFromInbox), Verdict(LostFromInbox))
        {
            ReRunOutcome = TripwireReRunOutcome.NotReproduced,
            Implicated = new[] { "LiveMoveArchive", "LivePhase4" },
        };

        TripwireRetryReport report = TripwireRetryLadder.Resolve(Verdict(LostFromInbox), source);

        Assert.Equal(TripwireRunOutcome.PassedWithASurvivedDelta, report.Outcome);
        Assert.True(report.Failed);
        Assert.Empty(report.Confirmed);
        Assert.Equal(1, source.ReRuns);
        Assert.True(report.BoundReached);
        Assert.Equal(
            "retry: 2 re-census(es), 1 re-run(s), verdict PASSED WITH A SURVIVED DELTA (fails the run)",
            report.Summary);

        // A pass that took a re-run to reach carries the whole record with it.
        Assert.Contains(
            "re-run 1 of 1 over 2 implicated selection(s) (LiveMoveArchive, LivePhase4): not reproduced",
            report.Describe(),
            StringComparison.Ordinal);
        Assert.Contains("pass ON THIS RECORD", report.Describe(), StringComparison.Ordinal);
        Assert.Contains("the run STILL FAILS", report.Describe(), StringComparison.Ordinal);
        Assert.Contains(LostFromInbox, report.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheCappedReRun_FailsLoudlyWhenItReproduces()
    {
        FakeCensusSource source = new(Verdict(LostFromInbox), Verdict(LostFromInbox))
        {
            ReRunOutcome = TripwireReRunOutcome.Reproduced,
            Implicated = new[] { "LiveMoveArchive" },
        };

        TripwireRetryReport report = TripwireRetryLadder.Resolve(Verdict(LostFromInbox), source);

        Assert.True(report.Failed);
        Assert.Equal(LostFromInbox, Assert.Single(report.Confirmed).Key);
        Assert.Equal(1, source.ReRuns);
        Assert.Equal("retry: 2 re-census(es), 1 re-run(s), verdict FAILED", report.Summary);
        Assert.Contains("RETRIED AND STILL FAILING", report.Describe(), StringComparison.Ordinal);
        Assert.Contains("): REPRODUCED.", report.Describe(), StringComparison.Ordinal);
        Assert.Contains("THIS IS THE SUITE REMOVING MAIL", report.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void AReRunThatCouldNotBeCarriedOut_FailsRatherThanExonerating()
    {
        // The live tier cannot start a second xunit run from inside its own teardown, so this
        // is the branch it actually takes. An experiment nobody performed exonerates nothing.
        FakeCensusSource source = new(Verdict(LostFromInbox), Verdict(LostFromInbox))
        {
            ReRunOutcome = TripwireReRunOutcome.Inconclusive,
        };

        TripwireRetryReport report = TripwireRetryLadder.Resolve(Verdict(LostFromInbox), source);

        Assert.True(report.Failed);
        Assert.Equal(1, source.ReRuns);
        Assert.Contains("INCONCLUSIVE", report.Describe(), StringComparison.Ordinal);
        Assert.Contains("exonerates nothing", report.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void NothingPlausiblyImplicated_FailsInsteadOfPassingForWantOfAnExperiment()
    {
        FakeCensusSource source = new(Verdict(LostFromInbox), Verdict(LostFromInbox))
        {
            ReRunOutcome = TripwireReRunOutcome.NotReproduced,
            Implicated = Array.Empty<string>(),
        };

        TripwireRetryReport report = TripwireRetryLadder.Resolve(Verdict(LostFromInbox), source);

        Assert.True(report.Failed);
        Assert.Equal(0, source.ReRuns);
        Assert.True(report.BoundReached);
        Assert.Contains("nothing to re-run", report.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheBoundsAreTwoReCensusesThirtySecondsApartAndOneReRun()
    {
        // Written as literals on purpose. Reading the constant back through itself would pass
        // for any value, and these three numbers are the whole cost model: two censuses of
        // ~16.9 s each is ~34 s of work against a tier run of about 27 minutes.
        Assert.Equal(2, TripwireRetryLadder.MaxReCensuses);
        Assert.Equal(30, TripwireRetryLadder.ReCensusGapSeconds);
        Assert.Equal(1, TripwireRetryLadder.MaxImplicatedReRuns);
        Assert.Equal(TimeSpan.FromSeconds(30), TripwireRetryLadder.ReCensusGap);

        FakeCensusSource source = new(Verdict(LostFromInbox), Verdict(LostFromInbox))
        {
            ReRunOutcome = TripwireReRunOutcome.Reproduced,
        };

        TripwireRetryLadder.Resolve(Verdict(LostFromInbox), source);

        // Spent in full, and never over: the fake throws rather than answering a third census.
        Assert.Equal(2, source.ReCensuses);
        Assert.Equal(1, source.ReRuns);
        Assert.Equal(new[] { TimeSpan.FromSeconds(30), TimeSpan.FromSeconds(30) }, source.Waits);
    }

    [Fact]
    public void AFailureThatAppearsONLYDuringAReCensus_IsNamedButNeverConfirmed()
    {
        // Confirmation is an INTERSECTION, so a delta seen once cannot fail the run however
        // alarming it looks - but it must not vanish from the record either.
        FakeCensusSource source = new(Verdict(LostFromArchive));

        TripwireRetryReport report = TripwireRetryLadder.Resolve(Verdict(LostFromInbox), source);

        Assert.False(report.Failed);
        Assert.Contains("new since the post-run census: " + LostFromArchive, report.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void TwoFailureLinesWithTheSameKey_AreOneThingToConfirm()
    {
        // The bound has to mean the same amount of work on every run, and a store that
        // produced the same failure twice must not consume it twice over.
        TripwireVerdict suspected = new(
            new[]
            {
                new TripwireFailure(LostFromInbox, "  first line."),
                new TripwireFailure(LostFromInbox, "  second line, same key."),
            },
            Array.Empty<string>(),
            "ATTRIBUTION: undecidable.");
        FakeCensusSource source = new(Verdict(LostFromInbox), Verdict(LostFromInbox))
        {
            ReRunOutcome = TripwireReRunOutcome.Reproduced,
        };

        TripwireRetryReport report = TripwireRetryLadder.Resolve(suspected, source);

        Assert.True(report.Failed);
        Assert.Single(report.Confirmed);
        Assert.Contains("post-run census: 1 suspected failure(s)", report.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void ACleanVerdictIsNeverRetried()
    {
        // Re-censusing a pass can only ever turn it into a failure, and it would spend the
        // bounds on nothing.
        FakeCensusSource source = new();

        Assert.Throws<ArgumentException>(() => TripwireRetryLadder.Resolve(Verdict(), source));
        Assert.Equal(0, source.ReCensuses);
    }

    [Fact]
    public void ARunThatNeverSuspectedAnything_ReportsThatToo()
    {
        TripwireRetryReport report = TripwireRetryReport.NotNeeded();

        Assert.False(report.Entered);
        Assert.False(report.Failed);
        Assert.Equal(0, report.ReCensuses);
        Assert.Equal(0, report.ReRuns);
        Assert.Equal("retry: none needed", report.Summary);
        Assert.Contains("no retry", report.Describe(), StringComparison.Ordinal);
        Assert.DoesNotContain("RETRIED", report.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheLiveTripwireReallyGoesThroughTheLadder_ReadOutOfTheCompiledIl()
    {
        // The wiring itself sits behind a COM census no CI test can execute, so it is read
        // out of the compiled method instead of trusted. A policy nothing calls is the same
        // thing as no policy, and the failure would be silent: the run would simply report a
        // suspected loss the way it did before this existed.
        MethodInfo verify = typeof(LiveStoreCountTripwire).GetMethod(
            "Verify",
            BindingFlags.NonPublic | BindingFlags.Static,
            binder: null,
            types: new[] { typeof(bool) },
            modifiers: null)!;

        Assert.Contains(
            MethodsCalledBy(verify),
            called => called.DeclaringType == typeof(TripwireRetryLadder)
                && called.Name == nameof(TripwireRetryLadder.Resolve));
    }

    /// <summary>
    /// Every method one method CALLS, resolved from its IL. <c>call</c> is 0x28 and
    /// <c>callvirt</c> 0x6F, each followed by a 4-byte metadata token; the token is resolved
    /// rather than pattern-matched, so a stray operand byte cannot pass for an instruction.
    /// </summary>
    private static List<MethodBase> MethodsCalledBy(MethodInfo method)
    {
        byte[] il = method.GetMethodBody()!.GetILAsByteArray()!;
        List<MethodBase> called = new();
        for (int i = 0; i + 4 < il.Length; i++)
        {
            if (il[i] != 0x28 && il[i] != 0x6F)
            {
                continue;
            }

            int token = il[i + 1] | (il[i + 2] << 8) | (il[i + 3] << 16) | (il[i + 4] << 24);
            try
            {
                MethodBase? resolved = method.Module.ResolveMethod(token);
                if (resolved != null)
                {
                    called.Add(resolved);
                }
            }
            catch (ArgumentException)
            {
                // Not a real call - the bytes happened to look like one.
            }
        }

        return called;
    }

    /// <summary>
    /// A scripted census source. It answers with the verdicts it was given, in order, and
    /// THROWS when the ladder asks for one more than it was scripted for - which turns a
    /// ladder that lost its bound into a fast, named test failure instead of a hang.
    /// </summary>
    private sealed class FakeCensusSource : ITripwireRetrySource
    {
        private readonly Queue<TripwireVerdict> _scripted;

        internal FakeCensusSource(params TripwireVerdict[] censuses)
        {
            _scripted = new Queue<TripwireVerdict>(censuses);
        }

        internal TripwireReRunOutcome ReRunOutcome { get; init; } = TripwireReRunOutcome.Inconclusive;

        internal IReadOnlyList<string> Implicated { get; init; } = new[] { "LivePhase1" };

        internal int ReCensuses { get; private set; }

        internal int ReRuns { get; private set; }

        internal List<int> ReCensusAttempts { get; } = new();

        internal List<TimeSpan> Waits { get; } = new();

        public void Wait(TimeSpan gap)
        {
            Waits.Add(gap);
        }

        public TripwireVerdict ReCensus(int attempt)
        {
            ReCensuses++;
            ReCensusAttempts.Add(attempt);
            if (_scripted.Count == 0)
            {
                throw new InvalidOperationException(
                    "The ladder asked for re-census " + attempt + ", which is more than this case scripted. "
                    + "Either the bound was raised or the stop-when-two-censuses-agree rule stopped working.");
            }

            return _scripted.Dequeue();
        }

        public IReadOnlyList<string> ImplicatedBy(IReadOnlyList<TripwireFailure> persisting)
        {
            Assert.NotEmpty(persisting);
            return Implicated;
        }

        public TripwireReRunOutcome ReRun(IReadOnlyList<string> implicated, int attempt)
        {
            ReRuns++;
            if (ReRuns > 1)
            {
                throw new InvalidOperationException(
                    "The ladder ran the implicated tests " + ReRuns + " times. One re-run is the bound; an "
                    + "experiment repeated until it gives the answer you wanted is not evidence.");
            }

            Assert.NotEmpty(implicated);
            Assert.Equal(ReRuns, attempt);
            return ReRunOutcome;
        }
    }
}
