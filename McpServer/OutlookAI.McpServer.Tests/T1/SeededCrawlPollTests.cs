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
/// Synthetic numbers only. Nothing here touches Outlook, an index, a settings file or a mailbox.
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

    [Fact]
    public void AWaitThatCompletedNoPollAtAllProvesNothingAndSaysSo()
    {
        // The outcome the defect used to print as a statement about the indexer. Nothing was
        // asked, so nothing may be concluded - and the line carries the repository's own greppable
        // idiom rather than a new one.
        Assert.Equal(SeededCrawlVerdict.NothingAsked, SeededCrawlPoll.Decide(attachmentRows: 0, pollsCompleted: 0));

        string line = SeededCrawlPoll.Explain(SeededCrawlVerdict.NothingAsked, pollsCompleted: 0, pollsLost: 5);

        Assert.StartsWith(SeededCrawlPoll.NothingAsked + ":", line, StringComparison.Ordinal);
        Assert.Contains("the index was never actually asked", line, StringComparison.Ordinal);
        Assert.Contains("NOT evidence that the gatherer was slow", line, StringComparison.Ordinal);
        Assert.Contains("what ran out was the statement", line, StringComparison.Ordinal);
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
        string path = Path.Combine(TestProjectDir(), "T2", "LiveAttachmentKindRecallTests.cs");
        Assert.True(File.Exists(path), "the seeded attachment probe's source is missing: " + path);
        string source = File.ReadAllText(path);

        Assert.Contains(
            "client.ExecuteRows(sql, 200, SeededCrawlPoll.StatementTimeoutSeconds)", source, StringComparison.Ordinal);
        Assert.Contains("catch (OleDbException) when (", source, StringComparison.Ordinal);
        Assert.Contains("SeededCrawlPoll.IsExpiredStatement(client.Provider", source, StringComparison.Ordinal);
        Assert.DoesNotContain("catch (Exception", source, StringComparison.Ordinal);

        // And the three numbers are not re-typed beside the call that uses them.
        Assert.DoesNotContain("SeededCrawlWaitSeconds = ", source, StringComparison.Ordinal);
        Assert.DoesNotContain("SeededCrawlPollSeconds = ", source, StringComparison.Ordinal);
    }

    private static string TestProjectDir()
    {
        return typeof(LiveTestSettings).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
    }
}
