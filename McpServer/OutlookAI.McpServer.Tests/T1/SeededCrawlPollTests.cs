using System.Reflection;
using OutlookAI.Core.IndexSearch;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the seeded attachment-crawl probe's wait: the bound it puts on one statement, what an
/// expired statement costs, and what a run is entitled to say afterwards.
///
/// <para>
/// <b>The defect.</b> <c>T2/LiveAttachmentKindRecallTests.SeededMixedAttachments_*</c> polls the
/// search index for 90 s, every 5 s, for an item it has just created. It passed no
/// <c>commandTimeoutSeconds</c>, so every statement ran on
/// <see cref="OleDbIndexClient.DefaultCommandTimeoutSeconds"/> = 60 s - two thirds of the whole
/// wait. One slow statement therefore consumed most of the budget, and the probe then reported
/// "the gatherer had not crawled it", which named the indexer for something the indexer had not
/// done. The number and the sentence were about different things.
/// </para>
///
/// <para>
/// <b>What is pinned here, and why here.</b> The decision - bound the statement, treat an expired
/// one as ONE LOST POLL, and never attribute a lost poll to the gatherer - is
/// <see cref="SeededCrawlPoll"/>, which is pure. Its consumer is <c>Category=Live</c> and needs a
/// real Windows Search index, so CI can never run it; a cleverer live test would be a promise
/// nobody checks. <see cref="AFastProviderErrorIsAFaultAndMustReachTheTest"/> is the one that
/// matters most: "treat an expired statement as ordinary" becomes "swallow everything" in one
/// careless edit, and a swallowed fault reports as the very sentence this change removes.
/// </para>
///
/// <para>
/// <b>The second defect, found in the first one's shadow (decision 57, 2026-09-17).</b> Bounding
/// the poll said what ONE lost poll costs and left the case where every poll is lost still
/// reporting green - a live test passing on the strength of a wait that never asked the index
/// anything. The fix is the profile split this repository already uses for exactly this ambiguity:
/// refuse on Production, say so and pass on Portable. It is pinned on both sides here, and
/// <see cref="TheTwoProfilesCarryTheSameFindingAndDifferOnlyInWhetherTheyStopTheRun"/> is the
/// control - each side alone still passes if the profile is never read at all, which is the old
/// always-green behaviour wearing the new wording.
/// </para>
///
/// <para>
/// Synthetic numbers and a fabricated settings object. Nothing here touches Outlook, an index, a
/// machine-local settings file or a mailbox, and no real store or account name appears (S6).
/// </para>
/// </summary>
public sealed class SeededCrawlPollTests
{
    // ------------------------------------------------------------------ the arithmetic

    [Fact]
    public void TheBudgetAffordsFourPollsEvenWhenEveryStatementRunsToTheBound()
    {
        // The three literals only mean anything together. A bound raised without raising the wait
        // silently turns a poll loop back into one bounded statement, which is the defect wearing
        // a different number.
        Assert.True(
            SeededCrawlPoll.WaitSeconds
                >= SeededCrawlPoll.MinimumWorstCasePolls
                    * (SeededCrawlPoll.StatementTimeoutSeconds + SeededCrawlPoll.PollSeconds),
            $"a {SeededCrawlPoll.WaitSeconds}s wait with a {SeededCrawlPoll.StatementTimeoutSeconds}s "
                + $"statement bound and a {SeededCrawlPoll.PollSeconds}s gap affords "
                + $"{SeededCrawlPoll.WorstCasePolls} poll(s) in the worst case, below the "
                + $"{SeededCrawlPoll.MinimumWorstCasePolls} this probe needs to be a poll loop at all");

        Assert.True(SeededCrawlPoll.WorstCasePolls >= SeededCrawlPoll.MinimumWorstCasePolls);
    }

    [Fact]
    public void TheBoundIsStrictlyNARROWERThanTheClientsOwnDefault()
    {
        // Equal to the default is the state before this change: passing it explicitly would then
        // be decoration, and the probe would be back to one statement per two thirds of the wait.
        Assert.True(
            SeededCrawlPoll.StatementTimeoutSeconds < OleDbIndexClient.DefaultCommandTimeoutSeconds,
            "passing a bound equal to the client's own default changes nothing - that IS the defect");
    }

    [Fact]
    public void NoSingleStatementMayTakeHalfTheWait()
    {
        // The defect stated as a ratio rather than as two numbers: 60 of 90 was two thirds.
        Assert.True(SeededCrawlPoll.StatementTimeoutSeconds * 2 < SeededCrawlPoll.WaitSeconds);
    }

    // ------------------------------------------------------------------ what counts as expired

    [Fact]
    public void AStatementThatRanForItsBoundIsOneLostPoll()
    {
        Assert.True(SeededCrawlPoll.IsExpiredStatement(
            IndexProviderKind.OleDb, TimeSpan.FromSeconds(SeededCrawlPoll.StatementTimeoutSeconds)));

        // And longer still - the caller's stopwatch starts before the connection opens, so the
        // measured elapsed is normally longer than the statement the provider timed.
        Assert.True(SeededCrawlPoll.IsExpiredStatement(
            IndexProviderKind.OleDb, TimeSpan.FromSeconds(SeededCrawlPoll.StatementTimeoutSeconds + 4)));
    }

    [Fact]
    public void AFastProviderErrorIsAFaultAndMustReachTheTest()
    {
        // THE test. Everything else here is about what a run SAYS; this is what keeps the catch
        // from becoming a swallow. A statement that failed in a second did not run out of a
        // fifteen-second bound, whatever it failed with.
        Assert.False(SeededCrawlPoll.IsExpiredStatement(IndexProviderKind.OleDb, TimeSpan.FromSeconds(1)));
        Assert.False(SeededCrawlPoll.IsExpiredStatement(IndexProviderKind.OleDb, TimeSpan.Zero));
        Assert.False(SeededCrawlPoll.IsExpiredStatement(
            IndexProviderKind.OleDb, TimeSpan.FromSeconds(SeededCrawlPoll.StatementTimeoutSeconds / 2)));
    }

    [Fact]
    public void TheToleranceIsOneSecondAndTheEdgeIsPinnedBothWays()
    {
        // Two clocks and a rounded comparison, not a licence to round down. Stated as an exact
        // edge so that widening the slack has to be a deliberate edit to this number.
        TimeSpan justInside = TimeSpan.FromMilliseconds(
            (SeededCrawlPoll.StatementTimeoutSeconds * 1000) - SeededCrawlPoll.BoundToleranceMs);
        TimeSpan justOutside = justInside - TimeSpan.FromMilliseconds(1);

        Assert.True(SeededCrawlPoll.IsExpiredStatement(IndexProviderKind.OleDb, justInside));
        Assert.False(SeededCrawlPoll.IsExpiredStatement(IndexProviderKind.OleDb, justOutside));
        Assert.Equal(1000, SeededCrawlPoll.BoundToleranceMs);
    }

    [Fact]
    public void NothingIsEverALostPollOnTheClientThatIgnoresTheBound()
    {
        // AdodbIndexClient takes commandTimeoutSeconds and never reads it, so on that provider
        // there is no bound to expire - and an error from it is therefore always a fault. This is
        // the half that answers "the exception type depends on which client was selected": it does
        // not, because only one of the two clients implements the thing being caught.
        Assert.False(SeededCrawlPoll.BoundIsHonoured(IndexProviderKind.AdodbCom));
        Assert.True(SeededCrawlPoll.BoundIsHonoured(IndexProviderKind.OleDb));

        Assert.False(SeededCrawlPoll.IsExpiredStatement(
            IndexProviderKind.AdodbCom, TimeSpan.FromSeconds(SeededCrawlPoll.StatementTimeoutSeconds * 10)));
    }

    [Fact]
    public void AnAbsentBoundClassifiesNothing()
    {
        // Defensive, and cheap: "expired" is meaningless without something to expire against, and
        // zero is OLE DB's own spelling of "no limit".
        Assert.False(SeededCrawlPoll.IsExpiredStatement(IndexProviderKind.OleDb, TimeSpan.FromHours(1), 0));
        Assert.False(SeededCrawlPoll.IsExpiredStatement(IndexProviderKind.OleDb, TimeSpan.FromHours(1), -5));
    }

    // ------------------------------------------------------------------ what the run says

    [Fact]
    public void RowsFoundIsTheGathererHavingCrawledIt_WhateverWasLostAlongTheWay()
    {
        Assert.Equal(SeededCrawlVerdict.Crawled, SeededCrawlPoll.Decide(attachmentRows: 3, pollsCompleted: 1));
        Assert.Equal(SeededCrawlVerdict.Crawled, SeededCrawlPoll.Decide(attachmentRows: 1, pollsCompleted: 6));
    }

    [Fact]
    public void AnAnsweredWaitWithNoRowsIsTheOnlyOutcomeEntitledToNameTheGatherer()
    {
        Assert.Equal(SeededCrawlVerdict.NotCrawled, SeededCrawlPoll.Decide(attachmentRows: 0, pollsCompleted: 6));

        string line = SeededCrawlPoll.Explain(SeededCrawlVerdict.NotCrawled, pollsCompleted: 6, pollsLost: 0);
        Assert.Contains("the gatherer had not crawled the seeded item inside the budget", line, StringComparison.Ordinal);
        Assert.DoesNotContain(SeededCrawlPoll.NothingAsked, line, StringComparison.Ordinal);
    }

    [Fact]
    public void AWaitThatLostPollsSaysTheBudgetWasNotAllTheGatherers()
    {
        // The half-measured case. It still reports the gatherer - the index WAS asked and answered
        // - but a reader must not take the crawl latency off a run that spent part of the budget
        // elsewhere, which is exactly the misreading the old single sentence invited.
        string line = SeededCrawlPoll.Explain(SeededCrawlVerdict.NotCrawled, pollsCompleted: 3, pollsLost: 2);

        Assert.Contains("2 of 5 poll(s)", line, StringComparison.Ordinal);
        Assert.Contains("rather than the gatherer's", line, StringComparison.Ordinal);
        Assert.Contains("ran out of its " + SeededCrawlPoll.StatementTimeoutSeconds + "s bound", line, StringComparison.Ordinal);
    }

    /// <summary>
    /// The outcome the defect used to print as a statement about the indexer. Nothing was asked,
    /// so nothing may be concluded - and the line carries the repository's own greppable idiom
    /// rather than a new one.
    /// <para>
    /// <b>What this test used to be, kept because it is the evidence.</b> Until 2026-09-17 these
    /// four assertions were the WHOLE of what the repository said about this case: the verdict was
    /// classified, the sentence was written - and the live test then <b>passed</b>, on every
    /// machine, having established nothing whatever about the crawl. Nothing was deleted to close
    /// that; every assertion below is the one that was here. What was missing is underneath, in the
    /// three methods that pin what a run now DOES with the sentence. This one goes on pinning the
    /// sentence itself, because the sentence is the finding and the profile only decides whether
    /// the finding stops the run.
    /// </para>
    /// </summary>
    [Fact]
    public void AWaitThatCompletedNoPollAtAllProvesNothingAndSaysSo()
    {
        Assert.Equal(SeededCrawlVerdict.NothingAsked, SeededCrawlPoll.Decide(attachmentRows: 0, pollsCompleted: 0));

        string line = SeededCrawlPoll.Explain(SeededCrawlVerdict.NothingAsked, pollsCompleted: 0, pollsLost: 5);

        Assert.StartsWith(SeededCrawlPoll.NothingAsked + ":", line, StringComparison.Ordinal);
        Assert.Contains("the index was never actually asked", line, StringComparison.Ordinal);
        Assert.Contains("NOT evidence that the gatherer was slow", line, StringComparison.Ordinal);
        Assert.Contains("what ran out was the statement", line, StringComparison.Ordinal);
    }

    // ------------------------------------------- and what the run DOES about it, by profile

    [Fact]
    public void AProductionProfileRefusesAWaitThatAskedTheIndexNothing()
    {
        // Same idiom, same call, as LivePopulationCoverage and IdentityDraftCoverage: on the
        // profile this probe was written for, a wait that could not get one statement through
        // means the machine or its search client has drifted, and a green test hides it.
        List<string> lines = new();

        InvalidOperationException refusal = Assert.Throws<InvalidOperationException>(
            () => SeededCrawlPoll.Report(
                Settings(LiveMachineProfile.Production),
                attachmentRows: 0, pollsCompleted: 0, pollsLost: 5, lines.Add));

        // WHAT was not established comes first - a bare "a population was expected and is missing"
        // reads here as a flaky wait, which is the misreading this whole file exists to remove.
        Assert.Contains("the index was never actually asked", refusal.Message, StringComparison.Ordinal);
        Assert.Contains("NOT evidence that the gatherer was slow", refusal.Message, StringComparison.Ordinal);

        // ...including what a saturated indexer would have looked like INSTEAD, because "this is
        // not the gatherer" is only useful to a reader who knows what the gatherer does look like.
        Assert.Contains("looks the OTHER way round", refusal.Message, StringComparison.Ordinal);
        Assert.Contains("the statements COMPLETE", refusal.Message, StringComparison.Ordinal);

        // ...and what to do about it, on both halves of the message.
        Assert.Contains("Windows Search service is running", refusal.Message, StringComparison.Ordinal);
        Assert.Contains(SeededCrawlPoll.Population, refusal.Message, StringComparison.Ordinal);
        Assert.Contains("Production", refusal.Message, StringComparison.Ordinal);
        Assert.Contains("Portable", refusal.Message, StringComparison.Ordinal);

        // It refused BEFORE announcing anything - nothing on a Production machine may read as an
        // acceptable emptiness - so the PROVED NOTHING line never reaches the log at all.
        Assert.Empty(lines);
    }

    [Fact]
    public void APortableProfileSaysSoJustAsLoudlyAndLetsTheRunContinue()
    {
        // The other half of the same emptiness. On a machine with no working search index the
        // absence is simply true, and this test names Requires=SearchIndex, so it should not have
        // been selected here - failing it would fail a machine for a property of the machine.
        List<string> lines = new();

        SeededCrawlVerdict verdict = SeededCrawlPoll.Report(
            Settings(LiveMachineProfile.Portable),
            attachmentRows: 0, pollsCompleted: 0, pollsLost: 5, lines.Add);

        Assert.Equal(SeededCrawlVerdict.NothingAsked, verdict);

        string line = Assert.Single(lines);
        Assert.StartsWith(SeededCrawlPoll.NothingAsked + ":", line, StringComparison.Ordinal);
        Assert.Contains("the index was never actually asked", line, StringComparison.Ordinal);
        Assert.Contains("looks the OTHER way round", line, StringComparison.Ordinal);
        Assert.Contains("Windows Search service is running", line, StringComparison.Ordinal);
    }

    [Fact]
    public void TheTwoProfilesCarryTheSameFindingAndDifferOnlyInWhetherTheyStopTheRun()
    {
        // THE control, and it is needed in both directions. Delete the profile check and the
        // Portable half above still warns and still passes - which IS the always-green behaviour
        // this decision removed, in new words. So the assertion is the DIFFERENCE: one wait, the
        // same five lost polls, refused on one profile and merely reported on the other.
        List<string> portableLines = new();
        SeededCrawlVerdict portable = SeededCrawlPoll.Report(
            Settings(LiveMachineProfile.Portable),
            attachmentRows: 0, pollsCompleted: 0, pollsLost: 5, portableLines.Add);

        List<string> productionLines = new();
        InvalidOperationException refusal = Assert.Throws<InvalidOperationException>(
            () => SeededCrawlPoll.Report(
                Settings(LiveMachineProfile.Production),
                attachmentRows: 0, pollsCompleted: 0, pollsLost: 5, productionLines.Add));

        Assert.Equal(SeededCrawlVerdict.NothingAsked, portable);
        Assert.Empty(productionLines);

        // And that the Portable machine is told exactly as much: the announced line is the
        // greppable label in front of the finding, and the refusal is the same finding with the
        // shared drift sentence behind it. Word for word, so neither side can quietly become the
        // louder one - a warning that says less than the failure is a warning nobody acts on.
        string finding = Assert.Single(portableLines)[(SeededCrawlPoll.NothingAsked.Length + 2)..];
        Assert.StartsWith(finding, refusal.Message, StringComparison.Ordinal);
        Assert.True(
            refusal.Message.Length > finding.Length,
            "the Production refusal must add the drift sentence to the finding, not replace it");
    }

    [Fact]
    public void AProductionProfileIsNotRefusedForAWaitThatSIMPLYDidNotFindTheRow()
    {
        // The second control, and the reason NothingAsked is a separate verdict at all. An index
        // that ANSWERED and did not have the item is the gatherer's business and the machine's,
        // not drift - refusing there would fail a Production run for a slow indexer, which is
        // exactly the option decision 55 rejected. Reported, never refused, on either profile.
        List<string> lines = new();

        SeededCrawlVerdict verdict = SeededCrawlPoll.Report(
            Settings(LiveMachineProfile.Production),
            attachmentRows: 0, pollsCompleted: 6, pollsLost: 0, lines.Add);

        Assert.Equal(SeededCrawlVerdict.NotCrawled, verdict);
        Assert.Contains("the gatherer had not crawled", Assert.Single(lines), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(LiveMachineProfile.Portable)]
    [InlineData(LiveMachineProfile.Production)]
    public void AWaitThatFoundTheRowSaysNothingAndRefusesNothingOnEitherProfile(LiveMachineProfile profile)
    {
        // Rows found is rows found, however many polls were lost getting to them: the seed WAS
        // crawled, so there is nothing to announce and nothing to refuse.
        List<string> lines = new();

        Assert.Equal(
            SeededCrawlVerdict.Crawled,
            SeededCrawlPoll.Report(
                Settings(profile), attachmentRows: 2, pollsCompleted: 1, pollsLost: 3, lines.Add));

        Assert.Empty(lines);
    }

    [Fact]
    public void TheSettingsAndTheSinkAreBothRequired()
    {
        Assert.Throws<ArgumentNullException>(
            () => SeededCrawlPoll.Report(null!, 0, 0, 1, _ => { }));
        Assert.Throws<ArgumentNullException>(
            () => SeededCrawlPoll.Report(Settings(LiveMachineProfile.Portable), 0, 0, 1, null!));
    }

    [Fact]
    public void EveryNonCrawlOutcomeSaysAdmissionIsStillProvenElsewhere()
    {
        // A green run must never leave a reader thinking the class proved nothing: the corpus
        // tests in it assert admission unconditionally and do not depend on this seed.
        foreach (SeededCrawlVerdict verdict in new[] { SeededCrawlVerdict.NotCrawled, SeededCrawlVerdict.NothingAsked })
        {
            Assert.Contains(
                "Admission is proven unconditionally",
                SeededCrawlPoll.Explain(verdict, pollsCompleted: 1, pollsLost: 1),
                StringComparison.Ordinal);
        }
    }

    [Fact]
    public void TheAccountingIsPrintedOnACleanRunToo()
    {
        // "0 lost" on a good run is what makes "2 lost" on a bad one readable. A line that only
        // appears when something went wrong teaches a reader nothing about the normal case.
        string clean = SeededCrawlPoll.Accounting(IndexProviderKind.OleDb, pollsCompleted: 7, pollsLost: 0);

        Assert.Contains("7 completed poll(s), 0 lost", clean, StringComparison.Ordinal);
        Assert.DoesNotContain("IGNORES", clean, StringComparison.Ordinal);
    }

    [Fact]
    public void TheAccountingSaysSoWhenTheProviderIgnoresTheBoundEntirely()
    {
        // A run on the ADODB fallback is keeping no budget at all, and the line has to say that
        // rather than print a bound it is not applying.
        string fallback = SeededCrawlPoll.Accounting(IndexProviderKind.AdodbCom, pollsCompleted: 2, pollsLost: 0);

        Assert.Contains("IGNORES the bound", fallback, StringComparison.Ordinal);
        Assert.Contains("AdodbCom", fallback, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ the call site

    [Fact]
    public void TheLiveProbeStillPassesTheBoundAndStillCatchesNarrowly()
    {
        // The one thing a pure function cannot pin: that the live test still USES it. Dropping the
        // third argument compiles - it is optional on IIndexClient.ExecuteRows - and restores the
        // 60 s default exactly; widening the catch to Exception compiles too, and turns every
        // provider fault into a quiet lost poll.
        string source = LiveProbeSource();

        Assert.Contains(
            "client.ExecuteRows(sql, 200, SeededCrawlPoll.StatementTimeoutSeconds)", source, StringComparison.Ordinal);
        Assert.Contains("catch (OleDbException) when (", source, StringComparison.Ordinal);
        Assert.Contains("SeededCrawlPoll.IsExpiredStatement(client.Provider", source, StringComparison.Ordinal);
        Assert.DoesNotContain("catch (Exception", source, StringComparison.Ordinal);

        // And the three numbers are not re-typed beside the call that uses them.
        Assert.DoesNotContain("SeededCrawlWaitSeconds = ", source, StringComparison.Ordinal);
        Assert.DoesNotContain("SeededCrawlPollSeconds = ", source, StringComparison.Ordinal);
    }

    [Fact]
    public void TheLiveProbeGoesThroughReportRatherThanRoundIt()
    {
        // Report is the only member that reads the machine profile, so the profile check is only
        // in force while the live probe actually calls it. Calling Decide and printing Explain
        // beside it compiles, reads almost identically, prints the identical line - and is exactly
        // the always-green code this decision replaced. So the call is read out of the source,
        // which is the substitute the method above already uses for the same reason.
        string source = LiveProbeSource();

        Assert.Contains(
            "_fixture.Settings, attachmentRows.Count, pollsCompleted, pollsLost, _output.WriteLine)",
            source,
            StringComparison.Ordinal);
        Assert.Contains("SeededCrawlPoll.Report(", source, StringComparison.Ordinal);
        Assert.DoesNotContain("SeededCrawlPoll.Decide(", source, StringComparison.Ordinal);
        Assert.DoesNotContain("SeededCrawlPoll.Explain(", source, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ helpers

    /// <summary>
    /// A fabricated settings object carrying nothing but the profile and the two fields
    /// <c>LiveTestSettings</c> will not be constructed without. Synthetic names only (S6) - the
    /// same shape <c>LivePopulationCoverageTests</c> uses.
    /// </summary>
    private static LiveTestSettings Settings(LiveMachineProfile profile)
    {
        return new LiveTestSettings
        {
            MachineProfile = profile,
            TestHubStoreDisplayName = "hub@example.test",
            ExpectedStoreDisplayNames = new List<string> { "hub@example.test" },
        };
    }

    private static string LiveProbeSource()
    {
        string path = Path.Combine(TestProjectDir(), "T2", "LiveAttachmentKindRecallTests.cs");
        Assert.True(File.Exists(path), "the seeded attachment probe's source is missing: " + path);
        return File.ReadAllText(path);
    }

    private static string TestProjectDir()
    {
        return typeof(LiveTestSettings).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
    }
}
