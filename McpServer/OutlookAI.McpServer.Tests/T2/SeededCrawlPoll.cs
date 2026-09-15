using OutlookAI.Core.IndexSearch;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>What a whole seeded-crawl wait actually established.</summary>
public enum SeededCrawlVerdict
{
    /// <summary>The gatherer crawled the seeded item inside the budget; the probe's rows are real.</summary>
    Crawled = 0,

    /// <summary>
    /// The index was asked at least once and did not have the item. A statement about the
    /// GATHERER, and the only outcome entitled to be one.
    /// </summary>
    NotCrawled = 1,

    /// <summary>
    /// Not one poll completed - every statement ran out of its bound. A statement about the
    /// SEARCH CLIENT, which says nothing whatever about the gatherer.
    /// </summary>
    NothingAsked = 2,
}

/// <summary>
/// The seeded attachment-crawl probe's wait, as arithmetic and a classification rather than as
/// three literals and a <c>catch</c>.
///
/// <para>
/// <b>The defect.</b> <c>LiveAttachmentKindRecallTests.SeededMixedAttachments_*</c> polls the
/// Windows Search index for an item it has just created, for a total of
/// <see cref="WaitSeconds"/>, every <see cref="PollSeconds"/>. Until 2026-09-15 it passed no
/// <c>commandTimeoutSeconds</c>, so each statement ran on
/// <see cref="OleDbIndexClient.DefaultCommandTimeoutSeconds"/> - <b>60 s against a 90 s total</b>.
/// One slow statement therefore spent two thirds of the budget, and the probe then printed
/// "the gatherer had not crawled the seeded item inside the budget": it blamed the indexer for
/// what was the statement running out. The measurement and the sentence reporting it were about
/// different things, and nothing in the output distinguished them.
/// </para>
///
/// <para>
/// <b>What was chosen, and over what.</b> Three ways out were written down: (a) pass the bound and
/// let an expired statement throw - honest, but it fails a live run for a slow indexer, which is
/// the machine's property and not the product's; (b) pass the bound and treat an expired statement
/// as ONE LOST POLL; (c) leave the default and accept that "90 s" really means "90 s, of which one
/// statement may take 60". <b>(b) chosen.</b> The loop already treats "no rows yet" as an ordinary
/// outcome, and a statement that ran out is indistinguishable from it - both mean "the index did
/// not hand us the row this time round".
/// </para>
///
/// <para>
/// <b>What (b) must not become, and this is the whole reason the classification lives here.</b>
/// "Treat an expired statement as a lost poll" turns into "swallow everything" in one careless
/// edit, and a swallowed fault reports as the same sentence the defect produced - the gatherer
/// blamed for somebody else's failure, only now with a real error hidden behind it. So the catch
/// is narrow twice over. It catches <see cref="System.Data.OleDb.OleDbException"/> and nothing
/// else, and it only classifies one as a lost poll when the statement RAN FOR THE BOUND before
/// it failed. A provider error that arrives in a second is a fault and still fails the test.
/// </para>
///
/// <para>
/// <b>Why <c>OleDbException</c> is the exception type, established from the source rather than
/// caught broadly.</b> The probe's client is <c>IndexClientFactory.CreateAuto</c>, and the product
/// has exactly two <see cref="IIndexClient"/> implementations:
/// <see cref="OleDbIndexClient"/>, which sets <c>OleDbCommand.CommandTimeout</c> from the
/// <c>commandTimeoutSeconds</c> argument, and <see cref="AdodbIndexClient"/>, which takes the same
/// argument and <b>never reads it</b>. So a bound can only exist, and only expire, on the OleDb
/// path - where <c>System.Data.OleDb</c> surfaces provider failures as <c>OleDbException</c>,
/// which is what <c>IndexClientFactory.CreateAuto</c>'s own filter names first around the very
/// same <c>ExecuteRows</c> call. The late-bound <c>COMException</c> the earlier note worried about
/// cannot be a bound expiring, because on that client there is no bound;
/// <see cref="BoundIsHonoured"/> says so in code and <see cref="IsExpiredStatement"/> refuses to
/// classify anything on that provider.
/// </para>
///
/// <para>
/// <b>And the reporting half, which is the bug as a person meets it.</b> A wait that lost polls
/// did not give the gatherer the time the sentence claims, and a wait that lost EVERY poll never
/// asked the index at all. <see cref="Decide"/> separates those three outcomes and
/// <see cref="Explain"/> writes each one in its own words, so a green run can no longer say
/// something about the indexer that the run did not measure.
/// </para>
///
/// <para>
/// Pure: counts, a <see cref="TimeSpan"/> and a provider in, strings out. No COM, no index, no
/// settings file. The consumer is <c>Category=Live</c> and CI can never run it - the same reason
/// <see cref="ArtifactSweepPolicy"/>, <see cref="LivePopulationCoverage"/> and
/// <see cref="TripwireWatchSoundness"/> are shaped this way - so the decision lives where every
/// branch of it is reachable from a runner with no Outlook and no search index.
/// </para>
/// </summary>
public static class SeededCrawlPoll
{
    /// <summary>
    /// How long the seeded probe waits in total for the Windows Search gatherer to crawl the item
    /// it just created. A CEILING, not an expectation: crawl latency is the gatherer's business and
    /// the probe reports it rather than asserting it.
    /// </summary>
    public const int WaitSeconds = 90;

    /// <summary>Gap between polls of the index while waiting for that crawl.</summary>
    public const int PollSeconds = 5;

    /// <summary>
    /// The per-statement bound the probe passes to <see cref="IIndexClient.ExecuteRows"/>.
    /// <para>
    /// Deliberately far below <see cref="OleDbIndexClient.DefaultCommandTimeoutSeconds"/>, and for
    /// the opposite reason to the product's: a product search has ONE statement and must wait for
    /// the real answer, whereas this is a poll loop where a lost poll costs nothing and another
    /// one follows five seconds later. The number is chosen so the worst case - every statement
    /// running to the bound - still asks the index <see cref="MinimumWorstCasePolls"/> times inside
    /// <see cref="WaitSeconds"/>; <c>TheBudgetAffordsFourPollsEvenWhenEveryStatementRunsToTheBound</c>
    /// is the pin that keeps all three numbers honest together.
    /// </para>
    /// </summary>
    public const int StatementTimeoutSeconds = 15;

    /// <summary>
    /// How many polls the budget must still afford when every single statement runs to the bound.
    /// The floor under <see cref="StatementTimeoutSeconds"/>: one poll would make the wait a
    /// single bounded statement wearing a loop's clothing.
    /// </summary>
    public const int MinimumWorstCasePolls = 4;

    /// <summary>
    /// How far short of the bound a statement may stop and still count as the bound firing.
    /// <para>
    /// The caller's stopwatch starts before <c>ExecuteRows</c> opens its connection, so a measured
    /// elapsed is normally LONGER than the statement the provider timed - this is slack for the
    /// other direction only: two clocks and a rounded comparison. It is one second because a real
    /// fault arriving one second inside a fifteen-second bound is still a fault worth failing on.
    /// </para>
    /// </summary>
    public const int BoundToleranceMs = 1000;

    /// <summary>
    /// The phrase a run log is grepped for when a wait asked the index nothing at all. The
    /// repository's own idiom (<see cref="LivePopulationCoverage"/>), in one place because the
    /// T1 pin greps for it.
    /// </summary>
    public const string NothingAsked = "PROVED NOTHING";

    /// <summary>
    /// How many polls <see cref="WaitSeconds"/> affords when every statement runs to the bound and
    /// every gap is slept in full. The arithmetic the three literals have to satisfy together.
    /// </summary>
    public static int WorstCasePolls => WaitSeconds / (StatementTimeoutSeconds + PollSeconds);

    /// <summary>
    /// True when <paramref name="provider"/> actually applies a <c>commandTimeoutSeconds</c>.
    /// <para>
    /// Only <see cref="IndexProviderKind.OleDb"/> does: <see cref="AdodbIndexClient"/> accepts the
    /// argument and discards it. On that provider the bound is not a shorter wait, it is no wait
    /// at all - so a run has to SAY it fell back rather than print a budget it is not keeping.
    /// </para>
    /// </summary>
    public static bool BoundIsHonoured(IndexProviderKind provider)
    {
        return provider == IndexProviderKind.OleDb;
    }

    /// <summary>
    /// True when a failed statement is the bound firing - one lost poll - rather than a fault.
    /// <para>
    /// Time, not an HRESULT. A statement the caller bounded at <paramref name="boundSeconds"/>
    /// which failed only after running that long is the bound firing, whatever number the
    /// provider chose to fail with; one that failed in a second is something else and must reach
    /// the test. Classifying on the error code instead would mean guessing which of
    /// <c>DB_E_ABORTLIMITREACHED</c> and <c>DB_E_CANCELED</c> Search.CollatorDSO returns, and a
    /// guess that is wrong fails OPEN - the swallow-everything shape this exists to avoid.
    /// </para>
    /// <para>
    /// A provider that ignores the bound can never produce one: see <see cref="BoundIsHonoured"/>.
    /// </para>
    /// </summary>
    /// <param name="provider">The client that ran the statement.</param>
    /// <param name="statementRan">How long it ran before it failed.</param>
    /// <param name="boundSeconds">The bound that was passed to it.</param>
    public static bool IsExpiredStatement(
        IndexProviderKind provider, TimeSpan statementRan, int boundSeconds = StatementTimeoutSeconds)
    {
        if (!BoundIsHonoured(provider) || boundSeconds <= 0)
        {
            return false;
        }

        return statementRan.TotalMilliseconds >= ((long)boundSeconds * 1000) - BoundToleranceMs;
    }

    /// <summary>
    /// What the wait established, from what came back and how many polls actually ran.
    /// <para>
    /// <paramref name="pollsCompleted"/> is the load-bearing one and it is NOT the same as the
    /// elapsed time: a wait that spent its whole budget on statements that expired asked the index
    /// nothing, and the sentence it used to print - about the gatherer - was about something it
    /// had not measured.
    /// </para>
    /// </summary>
    /// <param name="attachmentRows">Attachment-content rows the last completed poll returned.</param>
    /// <param name="pollsCompleted">Polls whose statement returned rows or an empty result.</param>
    public static SeededCrawlVerdict Decide(int attachmentRows, int pollsCompleted)
    {
        if (attachmentRows > 0)
        {
            return SeededCrawlVerdict.Crawled;
        }

        return pollsCompleted > 0 ? SeededCrawlVerdict.NotCrawled : SeededCrawlVerdict.NothingAsked;
    }

    /// <summary>
    /// The poll accounting, appended to the probe's summary line on EVERY run - including the
    /// clean one, because "0 lost" is what makes the other lines mean anything.
    /// </summary>
    /// <param name="provider">The client the probe used.</param>
    /// <param name="pollsCompleted">Polls that asked the index and got an answer.</param>
    /// <param name="pollsLost">Polls whose statement ran out of its bound.</param>
    public static string Accounting(IndexProviderKind provider, int pollsCompleted, int pollsLost)
    {
        string line = pollsCompleted + " completed poll(s), " + pollsLost + " lost to the "
            + StatementTimeoutSeconds + "s statement bound";
        return BoundIsHonoured(provider)
            ? line
            : line + " (provider=" + provider + ", which IGNORES the bound - a slow statement here "
                + "is unbounded and can still spend the whole wait)";
    }

    /// <summary>
    /// What the run says when the probe found no attachment row. One sentence per verdict, and
    /// never the gatherer's name on an outcome the gatherer did not cause.
    /// </summary>
    /// <param name="verdict">From <see cref="Decide"/>.</param>
    /// <param name="pollsCompleted">Polls that asked the index and got an answer.</param>
    /// <param name="pollsLost">Polls whose statement ran out of its bound.</param>
    public static string Explain(SeededCrawlVerdict verdict, int pollsCompleted, int pollsLost)
    {
        const string admission = "Admission is proven unconditionally by the corpus tests in "
            + "this class, which do not depend on this seed.";

        switch (verdict)
        {
            case SeededCrawlVerdict.Crawled:
                return "the gatherer crawled the seeded item inside the budget after " + pollsCompleted
                    + " completed poll(s).";

            case SeededCrawlVerdict.NotCrawled when pollsLost == 0:
                return "the gatherer had not crawled the seeded item inside the budget - the index "
                    + "answered all " + pollsCompleted + " poll(s) and did not have it. " + admission;

            case SeededCrawlVerdict.NotCrawled:
                return "the gatherer had not crawled the seeded item in what was LEFT of the budget: "
                    + pollsLost + " of " + (pollsCompleted + pollsLost) + " poll(s) ended in a statement "
                    + "that ran out of its " + StatementTimeoutSeconds + "s bound, so part of this wait "
                    + "was the search client's rather than the gatherer's. Read the crawl latency off a "
                    + "run with no lost polls, not off this one. " + admission;

            default:
                return NothingAsked + ": not one of the " + pollsLost + " poll(s) completed - every "
                    + "statement ran out of its " + StatementTimeoutSeconds + "s bound, so the index was "
                    + "never actually asked whether it had the seeded item. This is NOT evidence that "
                    + "the gatherer was slow: what ran out was the statement. Treat the crawl half of "
                    + "this test as not having run. " + admission;
        }
    }
}
