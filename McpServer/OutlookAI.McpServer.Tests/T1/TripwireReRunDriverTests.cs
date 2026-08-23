using System.Reflection;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the two maintainer decisions the retry ladder now carries: that a re-run which clears a
/// delta reads as a PASS and still fails the run, and that the bounds a machine may spend come
/// from what that machine IS.
/// <para>
/// <b>Why these are worth their own file.</b> Both decisions are about the direction a guard
/// fails in, and both act at a line no CI test can execute - the throw at the end of
/// <c>LiveStoreCountTripwire.Verify</c>, behind a COM census, and the process launch behind it.
/// So the shape here is deliberate: every DECISION is a pure function pinned by value, and
/// every CALL to one of those decisions is read back out of the compiled IL. A policy nothing
/// calls is the same thing as no policy, and its absence would be silent - the run would simply
/// look the way it looked before.
/// </para>
/// <para>
/// No COM, no Outlook, no mailbox, no process is started by anything in this file.
/// </para>
/// </summary>
public sealed class TripwireReRunDriverTests
{
    private static readonly IReadOnlyList<string> SomeSteps = new[] { "a step." };

    private static readonly IReadOnlyList<TripwireFailure> SomeFailures = new[]
    {
        new TripwireFailure("items-removed|other@example.test|Inbox", "  items left a mailbox."),
    };

    // ---------------------------------------------------------------------------------------
    // The structural guarantee: a survived delta can never leave a zero exit code.
    // ---------------------------------------------------------------------------------------

    /// <summary>
    /// Every outcome, and which of them lets a run finish quietly. Exactly one does.
    /// </summary>
    public static TheoryData<TripwireRunOutcome> AllOutcomes()
    {
        TheoryData<TripwireRunOutcome> data = new();
        foreach (TripwireRunOutcome outcome in Enum.GetValues<TripwireRunOutcome>())
        {
            data.Add(outcome);
        }

        return data;
    }

    [Theory]
    [MemberData(nameof(AllOutcomes))]
    public void OnlyPassed_LetsARunFinishQuietly(TripwireRunOutcome outcome)
    {
        // Failed is DERIVED from Outcome rather than stored beside it, so this is the whole
        // rule and there is no second place it can disagree with itself. ReportWith throws for
        // an outcome nobody has classified, which is what makes a fourth value fail here rather
        // than silently default to "passes".
        TripwireRetryReport report = ReportWith(outcome);

        Assert.Equal(outcome, report.Outcome);
        Assert.Equal(outcome != TripwireRunOutcome.Passed, report.Failed);
    }

    [Fact]
    public void TheOutcomeVocabularyIsThreeValues_AndPassedIsTheZeroOne()
    {
        // Written as literals: reading the enum back through itself would pass for any shape.
        // Passed is zero so that a default-constructed or forgotten outcome reads as the
        // permissive one ONLY where every construction path is a named factory - which is why
        // the constructor invariant below exists as well.
        Assert.Equal(
            new[]
            {
                TripwireRunOutcome.Passed,
                TripwireRunOutcome.PassedWithASurvivedDelta,
                TripwireRunOutcome.Failed,
            },
            Enum.GetValues<TripwireRunOutcome>());
        Assert.Equal(0, (int)TripwireRunOutcome.Passed);
    }

    [Fact]
    public void AReportCannotSayADeltaSurvivedAndAlsoReportACleanPass()
    {
        // THE mutation this file exists to kill: make a survived-delta run exit zero. The
        // invariant is in the constructor, so it does not matter which factory or which caller
        // tries it - such a report cannot be built at all. Reached by reflection because every
        // public path derives the outcome correctly, which is exactly why the guard is needed:
        // a mutation would change one of those derivations, and this catches it there.
        ConstructorInfo ctor = typeof(TripwireRetryReport).GetConstructors(
            BindingFlags.NonPublic | BindingFlags.Instance).Single();

        TargetInvocationException thrown = Assert.Throws<TargetInvocationException>(
            () => ctor.Invoke(new object?[]
            {
                /* entered */ true,
                /* outcome */ TripwireRunOutcome.Passed,
                /* confirmed */ Array.Empty<TripwireFailure>(),
                /* reCensuses */ 2,
                /* reRuns */ 1,
                /* survivedARecensus */ true,
                /* boundReached */ true,
                /* steps */ SomeSteps,
            }));

        ArgumentException inner = Assert.IsType<ArgumentException>(thrown.InnerException);
        Assert.Contains("SURVIVED a re-census", inner.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AReportCannotClaimASurvivedDeltaThatNeverSurvivedAnything()
    {
        // The other direction of the same invariant. A headline telling a reader to go looking
        // for items that were never missing twice is its own kind of false alarm.
        ConstructorInfo ctor = typeof(TripwireRetryReport).GetConstructors(
            BindingFlags.NonPublic | BindingFlags.Instance).Single();

        TargetInvocationException thrown = Assert.Throws<TargetInvocationException>(
            () => ctor.Invoke(new object?[]
            {
                true,
                TripwireRunOutcome.PassedWithASurvivedDelta,
                Array.Empty<TripwireFailure>(),
                1,
                0,
                /* survivedARecensus */ false,
                false,
                SomeSteps,
            }));

        Assert.IsType<ArgumentException>(thrown.InnerException);
    }

    // ---------------------------------------------------------------------------------------
    // The refusal: the one decision that says whether a verified run may finish quietly.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void ACleanVerdictAndNoRetry_RefusesNothing()
    {
        Assert.Null(TripwireRunVerdict.RefusalFor(Verdict(), TripwireRetryReport.NotNeeded()));
    }

    [Fact]
    public void ACleanVerdictAndASurvivedDelta_STILLREFUSES()
    {
        // The whole point. Nothing was CONFIRMED as lost - Confirmed is empty and the rebuilt
        // verdict is clean - and the run must not finish quietly all the same.
        TripwireRetryReport retry = TripwireRetryReport.Cleared(
            reCensuses: 2, reRuns: 1, survivedARecensus: true, boundReached: true, SomeSteps);

        string refusal = Assert.IsType<string>(TripwireRunVerdict.RefusalFor(Verdict(), retry));

        Assert.Contains(TripwireRetryReport.SurvivedDeltaHeadline, refusal, StringComparison.Ordinal);
        Assert.DoesNotContain("STORE COUNT TRIPWIRE", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void AConfirmedLoss_RefusesWithBothTheCensusAndTheRetryRecord()
    {
        TripwireRetryReport retry = TripwireRetryReport.Confirms(
            SomeFailures, reCensuses: 2, reRuns: 1, survivedARecensus: true, boundReached: true, SomeSteps);

        string refusal = Assert.IsType<string>(TripwireRunVerdict.RefusalFor(Verdict(SomeFailures[0]), retry));

        Assert.Contains("STORE COUNT TRIPWIRE", refusal, StringComparison.Ordinal);
        Assert.Contains("RETRIED AND STILL FAILING", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void ACleanRunWithNoRetryIsTheONLYWayOut()
    {
        // Stated as a closed set rather than as three separate cases, because the danger is a
        // fourth way out being added later without anybody noticing it exists.
        TripwireRetryReport[] reports =
        {
            TripwireRetryReport.NotNeeded(),
            TripwireRetryReport.Cleared(1, 0, survivedARecensus: false, boundReached: false, SomeSteps),
            TripwireRetryReport.Cleared(2, 1, survivedARecensus: true, boundReached: true, SomeSteps),
            TripwireRetryReport.Confirms(SomeFailures, 2, 1, true, true, SomeSteps),
        };

        List<string> quiet = new();
        foreach (TripwireRetryReport report in reports)
        {
            foreach (TripwireVerdict verdict in new[] { Verdict(), Verdict(SomeFailures[0]) })
            {
                if (TripwireRunVerdict.RefusalFor(verdict, report) == null)
                {
                    quiet.Add(verdict.Failed ? "failing verdict" : "clean verdict");
                }
            }
        }

        Assert.Equal(new[] { "clean verdict", "clean verdict" }, quiet);
    }

    // ---------------------------------------------------------------------------------------
    // The per-machine bounds.
    // ---------------------------------------------------------------------------------------

    [Fact]
    public void TheProductionPolicyIsTheThreeMeasuredNumbers()
    {
        // Literals on purpose: reading the constants back through themselves would pass for any
        // value, and these three are the cost model - two censuses of ~16.9 s each is ~34 s of
        // work against a tier run of about 27 minutes.
        Assert.Equal(2, TripwireRetryPolicy.Production.MaxReCensuses);
        Assert.Equal(30, TripwireRetryPolicy.Production.ReCensusGapSeconds);
        Assert.Equal(1, TripwireRetryPolicy.Production.MaxImplicatedReRuns);
        Assert.Equal(TimeSpan.FromSeconds(30), TripwireRetryPolicy.Production.ReCensusGap);
        Assert.True(TripwireRetryPolicy.Production.RetriesAtAll);
    }

    [Fact]
    public void TheNonePolicySpendsNothing()
    {
        Assert.Equal(0, TripwireRetryPolicy.None.MaxReCensuses);
        Assert.Equal(0, TripwireRetryPolicy.None.ReCensusGapSeconds);
        Assert.Equal(0, TripwireRetryPolicy.None.MaxImplicatedReRuns);
        Assert.False(TripwireRetryPolicy.None.RetriesAtAll);
        Assert.Contains("NO RETRIES", TripwireRetryPolicy.None.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void OnlyADeclaredProductionProfileBuysRetries()
    {
        // The direction that fails safe. Every retry is a chance to convert a real loss into a
        // pass, so a machine gets the accommodation only by SAYING it is the kind of machine
        // that needs it - a real mailbox with real people, real rules and a real server in it.
        Assert.Same(TripwireRetryPolicy.Production, TripwireRetryPolicy.For(LiveMachineProfile.Production));
        Assert.Same(TripwireRetryPolicy.None, TripwireRetryPolicy.For(LiveMachineProfile.Portable));
    }

    [Fact]
    public void AProfileNobodyHasThoughtAbout_GetsNoRetries()
    {
        // A value added to LiveMachineProfile later, before anybody decides what it means. The
        // undecided answer must be the strict one: the tier gets noisier, never weaker.
        Assert.Same(TripwireRetryPolicy.None, TripwireRetryPolicy.For((LiveMachineProfile)93));
    }

    [Fact]
    public void EveryDeclaredProfileIsClassified_AndOnlyOneOfThemRetries()
    {
        // Guards the shape rather than the two values: a third profile added later shows up
        // here as a second retrying profile, or as a profile nobody classified.
        List<LiveMachineProfile> retrying = Enum.GetValues<LiveMachineProfile>()
            .Where(profile => TripwireRetryPolicy.For(profile).RetriesAtAll)
            .ToList();

        Assert.Equal(new[] { LiveMachineProfile.Production }, retrying);
    }

    [Theory]
    [InlineData(LiveMachineProfile.Production)]
    [InlineData(LiveMachineProfile.Portable)]
    public void AReRunChildKeepsItsCensusesAndLosesItsReRun(LiveMachineProfile profile)
    {
        // The recursion guard, and the reason it is not simply "the child gets None": the
        // censuses are cheap and they stop a child crying wolf over a transient, while the
        // re-run rung is the one that would recurse a whole live tier at a time.
        TripwireRetryPolicy parent = TripwireRetryPolicy.For(profile, isReRunChild: false);
        TripwireRetryPolicy child = TripwireRetryPolicy.For(profile, isReRunChild: true);

        Assert.Equal(0, child.MaxImplicatedReRuns);
        Assert.Equal(parent.MaxReCensuses, child.MaxReCensuses);
        Assert.Equal(parent.ReCensusGapSeconds, child.ReCensusGapSeconds);
    }

    [Fact]
    public void UnderTheNonePolicy_ASuspectedLossFailsOnTheFirstReading()
    {
        // On a machine where nothing but the suite changes a mailbox, the first reading IS the
        // verdict: no census is asked for, no re-run is attempted, and the run fails. The fake
        // throws if anything is asked of it, so a policy that leaked a retry fails by name.
        RefusingSource source = new();

        TripwireRetryReport report = TripwireRetryLadder.Resolve(
            Verdict(SomeFailures[0]), source, TripwireRetryPolicy.None);

        Assert.True(report.Failed);
        Assert.Equal(TripwireRunOutcome.Failed, report.Outcome);
        Assert.Equal(0, report.ReCensuses);
        Assert.Equal(0, report.ReRuns);
        Assert.False(report.SurvivedARecensus);
        Assert.Single(report.Confirmed);
        Assert.Contains("NO RE-CENSUS IS CONFIGURED", report.Describe(), StringComparison.Ordinal);
        Assert.Contains("no re-run is permitted", report.Describe(), StringComparison.Ordinal);
    }

    // ---------------------------------------------------------------------------------------
    // The out-of-process driver.
    // ---------------------------------------------------------------------------------------

    [Theory]
    // Never started, or still running past the budget: no experiment, no answer.
    [InlineData(null, TripwireReRunMarker.Absent, TripwireReRunOutcome.Inconclusive)]
    [InlineData(null, TripwireReRunMarker.Clean, TripwireReRunOutcome.Inconclusive)]
    [InlineData(null, TripwireReRunMarker.TripwireFailed, TripwireReRunOutcome.Inconclusive)]
    // The child's own tripwire fired. The ONLY reading that earns "the suite is removing mail".
    [InlineData(1, TripwireReRunMarker.TripwireFailed, TripwireReRunOutcome.Reproduced)]
    [InlineData(0, TripwireReRunMarker.TripwireFailed, TripwireReRunOutcome.Reproduced)]
    // It ran and never reached its own verification: a refused preflight, a crash. Proves nothing.
    [InlineData(1, TripwireReRunMarker.Absent, TripwireReRunOutcome.Inconclusive)]
    [InlineData(0, TripwireReRunMarker.Absent, TripwireReRunOutcome.Inconclusive)]
    // Clean census AND a clean run: the one exoneration - which still reports a survived delta.
    [InlineData(0, TripwireReRunMarker.Clean, TripwireReRunOutcome.NotReproduced)]
    // Clean census but something else failed: the experiment did not complete on its own terms.
    [InlineData(1, TripwireReRunMarker.Clean, TripwireReRunOutcome.Inconclusive)]
    public void TheChildsAnswerIsReadFromItsOwnTripwire_NotFromItsExitCodeAlone(
        int? exitCode, TripwireReRunMarker marker, TripwireReRunOutcome expected)
    {
        Assert.Equal(expected, TripwireReRunDriver.Classify(exitCode, marker));
    }

    [Fact]
    public void ExactlyOneCombinationClearsADelta()
    {
        // Stated as a closed set: NotReproduced is the only outcome the ladder can act on to
        // clear anything, so anything that widens the set of ways to reach it widens the only
        // path to a pass in this whole design.
        List<string> clearing = new();
        foreach (int? exitCode in new int?[] { null, 0, 1 })
        {
            foreach (TripwireReRunMarker marker in Enum.GetValues<TripwireReRunMarker>())
            {
                if (TripwireReRunDriver.Classify(exitCode, marker) == TripwireReRunOutcome.NotReproduced)
                {
                    clearing.Add(exitCode + "/" + marker);
                }
            }
        }

        Assert.Equal(new[] { "0/Clean" }, clearing);
    }

    [Theory]
    [InlineData("Clean", TripwireReRunMarker.Clean)]
    [InlineData("TripwireFailed", TripwireReRunMarker.TripwireFailed)]
    [InlineData(" TripwireFailed\r\n", TripwireReRunMarker.TripwireFailed)]
    [InlineData("clean", TripwireReRunMarker.Absent)]
    [InlineData("", TripwireReRunMarker.Absent)]
    [InlineData(null, TripwireReRunMarker.Absent)]
    [InlineData("TripwireFail", TripwireReRunMarker.Absent)]
    public void AMarkerThatIsNotOneOfTheTwoSpellings_IsNoAnswerAtAll(string? text, TripwireReRunMarker expected)
    {
        // Fail-closed on a half-written or corrupt file: Absent is Inconclusive, which fails.
        Assert.Equal(expected, TripwireReRunDriver.ReadMarker(text));
        Assert.Equal(0, (int)TripwireReRunMarker.Absent);
    }

    [Fact]
    public void ThePlanRunsANEWPROCESS_WithTheSameBinariesAndNothingElse()
    {
        TripwireReRunPlan plan = TripwireReRunDriver.PlanFor(
            @"C:\dotnet\dotnet.exe", @"C:\repo\Tests.csproj", "FullyQualifiedName~A.", "Debug", @"C:\tmp\v.txt");

        Assert.Equal(@"C:\dotnet\dotnet.exe", plan.FileName);
        Assert.Equal(
            new[] { "test", @"C:\repo\Tests.csproj", "--no-build", "--filter", "FullyQualifiedName~A.", "-c", "Debug" },
            plan.Arguments);
        Assert.Equal(@"C:\tmp\v.txt", plan.MarkerPath);
    }

    [Fact]
    public void ThePlanFallsBackToDotnetOnPath_RatherThanRefusing()
    {
        // DOTNET_HOST_PATH is set by the SDK for anything it launches, so it is normally there;
        // when it is not, the muxer on PATH is a better answer than no experiment at all.
        TripwireReRunPlan plan = TripwireReRunDriver.PlanFor(
            null, @"C:\repo\Tests.csproj", "FullyQualifiedName~A.", string.Empty, @"C:\tmp\v.txt");

        Assert.Equal("dotnet", plan.FileName);
        Assert.DoesNotContain("-c", plan.Arguments);
    }

    [Fact]
    public void ThePlanRefusesToReRunTheWholeAssembly()
    {
        // An unfiltered re-run would schedule live collections this run never executed, against
        // real mailboxes, which is a different experiment from the one the ladder asked for.
        Assert.Throws<ArgumentException>(
            () => TripwireReRunDriver.PlanFor("dotnet", @"C:\repo\Tests.csproj", " ", "Debug", @"C:\tmp\v.txt"));
    }

    [Fact]
    public void ThePlanRefusesAChildThatCouldNotAnswer()
    {
        Assert.Throws<ArgumentException>(
            () => TripwireReRunDriver.PlanFor("dotnet", @"C:\repo\Tests.csproj", "F~A.", "Debug", " "));
        Assert.Throws<ArgumentException>(
            () => TripwireReRunDriver.PlanFor("dotnet", " ", "F~A.", "Debug", @"C:\tmp\v.txt"));
    }

    [Fact]
    public void TheFilterNamesTheImplicatedClasses_AndNothingElse()
    {
        // Deduplicated and ordered, so the same implicated set always produces the same command
        // - a filter that varied run to run would make a re-run unreproducible by hand. The
        // trailing dot keeps `Ns.Alpha` from also selecting `Ns.Alpha2`.
        Assert.Equal(
            "FullyQualifiedName~Ns.Alpha.|FullyQualifiedName~Ns.Beta.",
            TripwireReRunDriver.FilterFor(new[] { "Ns.Beta", "Ns.Alpha", "Ns.Beta" }));
        Assert.Null(TripwireReRunDriver.FilterFor(Array.Empty<string>()));
        Assert.Null(TripwireReRunDriver.FilterFor(new[] { " " }));
    }

    [Fact]
    public void CollectionsAreTranslatedToRealClassesOfThisAssembly()
    {
        // The translation exists because an xunit COLLECTION is not something the VSTest filter
        // language can express. Read off the [Collection] attributes by reflection, so it cannot
        // drift from them.
        IReadOnlyList<string> classes = TripwireReRunDriver.ClassesIn(
            typeof(LiveCollections).Assembly, new[] { LiveCollections.Lifecycle });

        Assert.NotEmpty(classes);
        Assert.All(classes, name => Assert.StartsWith("OutlookAI.McpServer.Tests.", name, StringComparison.Ordinal));
        Assert.Empty(TripwireReRunDriver.ClassesIn(
            typeof(LiveCollections).Assembly, new[] { "NoSuchCollection" }));

        // And every guarded collection really does have classes to re-run - a collection that
        // translated to nothing would make the rung Inconclusive for a reason nobody could see.
        List<string> empty = LiveCollections.All
            .Where(name => TripwireReRunDriver.ClassesIn(typeof(LiveCollections).Assembly, new[] { name }).Count == 0)
            .ToList();
        Assert.Empty(empty);
    }

    [Fact]
    public void TheReRunBudgetIsSixtyMinutes()
    {
        // A literal, and a CEILING: the re-run may cover a whole tier run (26.8 min measured,
        // 2026-08-18) plus the child's own censuses and re-census gaps. Too low kills a working
        // experiment; too high costs nothing when the child finishes early.
        Assert.Equal(60, TripwireReRunDriver.ReRunBudgetMinutes);
        Assert.Equal(TimeSpan.FromMinutes(60), TripwireReRunDriver.ReRunBudget);
    }

    [Fact]
    public void TheMarkerVariableIsWhatMakesAProcessAChild()
    {
        // One variable saying two things - "you are a re-run" and "write your verdict here" -
        // because two variables could disagree and one cannot.
        Assert.Equal("OUTLOOKAI_TRIPWIRE_RERUN_MARKER", TripwireReRunDriver.MarkerVariable);

        string? was = Environment.GetEnvironmentVariable(TripwireReRunDriver.MarkerVariable);
        try
        {
            Environment.SetEnvironmentVariable(TripwireReRunDriver.MarkerVariable, null);
            Assert.False(TripwireReRunDriver.IsReRunChild);
            Assert.Null(TripwireReRunDriver.MarkerPathForThisProcess());

            Environment.SetEnvironmentVariable(TripwireReRunDriver.MarkerVariable, "   ");
            Assert.False(TripwireReRunDriver.IsReRunChild);

            Environment.SetEnvironmentVariable(TripwireReRunDriver.MarkerVariable, @"C:\tmp\v.txt");
            Assert.True(TripwireReRunDriver.IsReRunChild);
            Assert.Equal(@"C:\tmp\v.txt", TripwireReRunDriver.MarkerPathForThisProcess());
        }
        finally
        {
            Environment.SetEnvironmentVariable(TripwireReRunDriver.MarkerVariable, was);
        }
    }

    [Fact]
    public void AnOrdinaryRunWritesNoMarker_AndAChildWritesOneThatReadsBack()
    {
        string? was = Environment.GetEnvironmentVariable(TripwireReRunDriver.MarkerVariable);
        string path = Path.Combine(Path.GetTempPath(), "outlookai-tripwire-marker-test", Guid.NewGuid() + ".txt");
        try
        {
            Environment.SetEnvironmentVariable(TripwireReRunDriver.MarkerVariable, null);
            TripwireReRunDriver.RecordOwnVerdict(tripwireFailed: true);
            Assert.False(File.Exists(path));

            Environment.SetEnvironmentVariable(TripwireReRunDriver.MarkerVariable, path);
            TripwireReRunDriver.RecordOwnVerdict(tripwireFailed: true);
            Assert.Equal(TripwireReRunMarker.TripwireFailed, TripwireReRunDriver.ReadMarker(File.ReadAllText(path)));

            TripwireReRunDriver.RecordOwnVerdict(tripwireFailed: false);
            Assert.Equal(TripwireReRunMarker.Clean, TripwireReRunDriver.ReadMarker(File.ReadAllText(path)));
        }
        finally
        {
            Environment.SetEnvironmentVariable(TripwireReRunDriver.MarkerVariable, was);
            try
            {
                Directory.Delete(Path.GetDirectoryName(path)!, recursive: true);
            }
            catch (DirectoryNotFoundException)
            {
            }
        }
    }

    // ---------------------------------------------------------------------------------------
    // The wiring, read out of the compiled IL because no CI test can execute these lines.
    // ---------------------------------------------------------------------------------------

    [Theory]
    [InlineData(false, TripwireRunOutcome.Passed, false)]
    [InlineData(false, TripwireRunOutcome.PassedWithASurvivedDelta, true)]
    [InlineData(false, TripwireRunOutcome.Failed, true)]
    [InlineData(true, TripwireRunOutcome.Failed, true)]
    public void EnforceThrowsForEverythingExceptACleanPass(
        bool verdictFailed, TripwireRunOutcome outcome, bool expectThrow)
    {
        // Enforce is the DECISION AND THE THROW in one call, and it exists in that shape
        // because its caller cannot be executed in CI. MEASURED: with the branch up there
        // instead, changing `if (refusal != null)` to `if (refusal != null && false)` passed
        // the entire 2,125-test suite. Down here both sides of it are ordinary CI.
        TripwireVerdict verdict = verdictFailed ? Verdict(SomeFailures[0]) : Verdict();
        TripwireRetryReport retry = ReportWith(outcome);

        if (expectThrow)
        {
            InvalidOperationException thrown = Assert.Throws<InvalidOperationException>(
                () => TripwireRunVerdict.Enforce(verdict, retry));
            Assert.Equal(TripwireRunVerdict.RefusalFor(verdict, retry), thrown.Message);
        }
        else
        {
            TripwireRunVerdict.Enforce(verdict, retry);
        }
    }

    [Fact]
    public void TheLiveTripwireRefusesThroughTheONEDecision_AndHasNoBranchOfItsOwn()
    {
        // Two reads, and then a third assertion that is the real point: Verify must CALL the
        // decision, and it must not contain the throw itself. A throw up there needs a
        // condition, a condition up there is behind a COM census, and a condition nothing can
        // execute is a condition a mutation can invert for free - which is exactly what
        // happened before this was restructured.
        MethodInfo verify = Verify();

        Assert.Contains(
            CalleesOf(verify),
            called => called.DeclaringType == typeof(TripwireRunVerdict)
                && called.Name == nameof(TripwireRunVerdict.Enforce));
        Assert.Contains(
            CalleesOf(verify),
            called => called.DeclaringType == typeof(TripwireRunVerdict)
                && called.Name == nameof(TripwireRunVerdict.RefusalFor));
        Assert.DoesNotContain(
            CalleesOf(verify),
            called => called.DeclaringType == typeof(InvalidOperationException) && called.IsConstructor);
    }

    [Fact]
    public void TheChildIsMarkedAsAChild_WhichIsWhatStopsTheLadderRecursing()
    {
        // MEASURED: writing the marker path to a DIFFERENT variable name passed the whole
        // suite, because the assignment was inside the launch and nothing in CI can launch
        // anything. The content is decided in a pure method now, pinned here by value, and the
        // launch is read back out of the IL for the call to it.
        TripwireReRunPlan plan = TripwireReRunDriver.PlanFor(
            "dotnet", @"C:\repo\Tests.csproj", "FullyQualifiedName~A.", "Debug", @"C:\tmp\v.txt");

        IReadOnlyDictionary<string, string> environment = TripwireReRunDriver.ChildEnvironment(plan);

        Assert.Equal(@"C:\tmp\v.txt", Assert.Contains(TripwireReRunDriver.MarkerVariable, environment));
        Assert.Single(environment);

        MethodInfo launch = typeof(TripwireReRunDriver).GetMethod(
            "Launch", BindingFlags.NonPublic | BindingFlags.Static)!;
        Assert.Contains(
            CalleesOf(launch),
            called => called.DeclaringType == typeof(TripwireReRunDriver)
                && called.Name == nameof(TripwireReRunDriver.ChildEnvironment));
    }

    [Fact]
    public void TheLiveTripwireTellsItsParentWhatItsOwnGuardDecided()
    {
        // Without this the parent reads Absent, which is Inconclusive, which fails - so the
        // failure direction is already safe and the loss is only in accuracy: every re-run
        // would report "no answer" instead of naming what happened.
        Assert.Contains(
            CalleesOf(Verify()),
            called => called.DeclaringType == typeof(TripwireReRunDriver)
                && called.Name == nameof(TripwireReRunDriver.RecordOwnVerdict));
    }

    [Fact]
    public void TheLiveTripwireTakesItsBoundsFromTheMACHINE()
    {
        // A policy nothing consults is the same thing as no policy, and the failure would be
        // silent: the VM would keep retrying exactly as it does today.
        MethodInfo ensureBaseline = typeof(LiveStoreCountTripwire).GetMethod(
            nameof(LiveStoreCountTripwire.EnsureBaseline),
            BindingFlags.Public | BindingFlags.Static)!;

        Assert.Contains(
            CalleesOf(ensureBaseline),
            called => called.DeclaringType == typeof(TripwireRetryPolicy)
                && called.Name == nameof(TripwireRetryPolicy.For));
    }

    [Fact]
    public void TheReRunRungReallyStartsANEWPROCESS()
    {
        // The out-of-process guarantee, and the only thing that makes the rung safe at all: an
        // in-process re-run would re-enter fixtures that are disposing and write to a mailbox
        // with nothing left to sweep it. Read in two steps because that is where the two
        // mutations are - the live source could stop calling the driver, or the driver could
        // stop starting a process.
        Type source = typeof(LiveStoreCountTripwire)
            .GetNestedTypes(BindingFlags.NonPublic)
            .Single(t => t.Name == "LiveRetrySource");
        MethodInfo reRun = source.GetMethod(nameof(ITripwireRetrySource.ReRun))!;

        Assert.Contains(
            CalleesOf(reRun),
            called => called.DeclaringType == typeof(TripwireReRunDriver)
                && called.Name == nameof(TripwireReRunDriver.Run));

        MethodInfo launch = typeof(TripwireReRunDriver).GetMethod(
            "Launch", BindingFlags.NonPublic | BindingFlags.Static)!;
        Assert.Contains(
            CalleesOf(launch),
            called => called.DeclaringType == typeof(System.Diagnostics.Process)
                && called.Name == nameof(System.Diagnostics.Process.Start));
    }

    [Fact]
    public void TheReRunChildIsNeverKilled()
    {
        // Killing a live run mid-flight leaves tagged artifacts in a real mailbox with nothing
        // left to sweep them - the exact hazard that kept this rung unimplemented. An abandoned
        // child finishes its own teardown; the parent just stops waiting and reports no answer.
        MethodInfo launch = typeof(TripwireReRunDriver).GetMethod(
            "Launch", BindingFlags.NonPublic | BindingFlags.Static)!;

        Assert.DoesNotContain(
            CalleesOf(launch),
            called => called.DeclaringType == typeof(System.Diagnostics.Process)
                && called.Name == nameof(System.Diagnostics.Process.Kill));
    }

    // ---------------------------------------------------------------------------------------

    private static MethodInfo Verify()
    {
        return typeof(LiveStoreCountTripwire).GetMethod(
            "Verify",
            BindingFlags.NonPublic | BindingFlags.Static,
            binder: null,
            types: new[] { typeof(bool) },
            modifiers: null)!;
    }

    /// <summary>
    /// A report carrying <paramref name="outcome"/>, built through the factory that produces it.
    /// An outcome nobody has classified throws rather than defaulting, so a fourth enum value
    /// fails <see cref="OnlyPassed_LetsARunFinishQuietly"/> instead of silently reading as a pass.
    /// </summary>
    private static TripwireRetryReport ReportWith(TripwireRunOutcome outcome)
    {
        return outcome switch
        {
            TripwireRunOutcome.Passed =>
                TripwireRetryReport.Cleared(1, 0, survivedARecensus: false, boundReached: false, SomeSteps),
            TripwireRunOutcome.PassedWithASurvivedDelta =>
                TripwireRetryReport.Cleared(2, 1, survivedARecensus: true, boundReached: true, SomeSteps),
            TripwireRunOutcome.Failed =>
                TripwireRetryReport.Confirms(SomeFailures, 2, 1, true, true, SomeSteps),
            _ => throw new ArgumentOutOfRangeException(
                nameof(outcome),
                outcome,
                "A new TripwireRunOutcome exists and nothing here says whether it fails a run. Classify it: "
                + "everything except Passed must fail, or the guard has a way out nobody chose."),
        };
    }

    private static TripwireVerdict Verdict(params TripwireFailure[] failures)
    {
        return new TripwireVerdict(
            failures,
            Array.Empty<string>(),
            failures.Length > 0 ? "ATTRIBUTION: undecidable." : null);
    }

    /// <summary>
    /// Every method and constructor one method calls, resolved from its IL. <c>call</c> is 0x28,
    /// <c>callvirt</c> 0x6F and <c>newobj</c> 0x73, each followed by a 4-byte metadata token; the
    /// token is RESOLVED rather than pattern-matched, so a stray operand byte cannot pass for an
    /// instruction.
    /// </summary>
    private static List<MethodBase> CalleesOf(MethodInfo method)
    {
        byte[] il = method.GetMethodBody()!.GetILAsByteArray()!;
        List<MethodBase> called = new();
        for (int i = 0; i + 4 < il.Length; i++)
        {
            if (il[i] != 0x28 && il[i] != 0x6F && il[i] != 0x73)
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
    /// A census source that answers nothing and throws if asked. Used to prove that a
    /// zero-retry policy asks for NOTHING rather than asking and ignoring the answer: a
    /// silent extra census on a machine that is not supposed to take one would show up as
    /// wasted minutes and as one more chance to convert a real loss into a pass.
    /// </summary>
    private sealed class RefusingSource : ITripwireRetrySource
    {
        public void Wait(TimeSpan gap)
        {
            throw new InvalidOperationException("A zero-retry policy waited. It has nothing to wait for.");
        }

        public TripwireVerdict ReCensus(int attempt)
        {
            throw new InvalidOperationException("A zero-retry policy asked for re-census " + attempt + ".");
        }

        public IReadOnlyList<string> ImplicatedBy(IReadOnlyList<TripwireFailure> persisting)
        {
            throw new InvalidOperationException(
                "A zero-retry policy asked which tests are implicated, which it can do nothing about.");
        }

        public TripwireReRunOutcome ReRun(IReadOnlyList<string> implicated, int attempt)
        {
            throw new InvalidOperationException("A zero-retry policy started a re-run.");
        }
    }
}
