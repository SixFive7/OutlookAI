using System.Globalization;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>What one bounded re-run of the plausibly-implicated tests decided.</summary>
public enum TripwireReRunOutcome
{
    /// <summary>
    /// The re-run could not be carried out at all. Deliberately the ZERO value, so a source
    /// that forgets to answer fails closed: an experiment nobody performed exonerates nothing.
    /// </summary>
    Inconclusive = 0,

    /// <summary>The implicated tests ran again and the same failure(s) did not come back.</summary>
    NotReproduced = 1,

    /// <summary>The same failure(s) came back. That is the suite, not the person using the mailbox.</summary>
    Reproduced = 2,
}

/// <summary>
/// Everything the retry ladder needs from the outside world, so the policy itself is pure and
/// CI can drive every branch of it without an Outlook, a mailbox or a wall clock.
/// </summary>
public interface ITripwireRetrySource
{
    /// <summary>Waits between two censuses. Injected so a self-test costs no wall clock.</summary>
    void Wait(TimeSpan gap);

    /// <summary>
    /// One more census of every watched store, compared against the SAME baseline, numbered
    /// from 1. Throwing is allowed and is fail-closed: a census that cannot be taken must
    /// refuse the run rather than clear a suspected loss.
    /// </summary>
    TripwireVerdict ReCensus(int attempt);

    /// <summary>
    /// Which tests could plausibly have produced <paramref name="persisting"/>. An empty
    /// answer means nothing can be re-run, which the ladder treats as a failure rather than
    /// as an exoneration.
    /// </summary>
    IReadOnlyList<string> ImplicatedBy(IReadOnlyList<TripwireFailure> persisting);

    /// <summary>Runs exactly those tests again and says whether the delta came back.</summary>
    TripwireReRunOutcome ReRun(IReadOnlyList<string> implicated, int attempt);
}

/// <summary>
/// The record of one trip through <see cref="TripwireRetryLadder"/>: the verdict, the bounds
/// that were spent, and what changed between attempts.
/// <para>
/// It exists because the retry is the dangerous half of the feature. Every retry is a chance
/// to convert a real loss into a pass, so a run that passed on the second census must never be
/// indistinguishable from one that passed on the first - and the only way to guarantee that is
/// for the passing path to carry its own evidence.
/// </para>
/// </summary>
public sealed class TripwireRetryReport
{
    private TripwireRetryReport(
        bool entered,
        bool failed,
        IReadOnlyList<TripwireFailure> confirmed,
        int reCensuses,
        int reRuns,
        bool survivedARecensus,
        bool boundReached,
        IReadOnlyList<string> steps)
    {
        Entered = entered;
        Failed = failed;
        Confirmed = confirmed;
        ReCensuses = reCensuses;
        ReRuns = reRuns;
        SurvivedARecensus = survivedARecensus;
        BoundReached = boundReached;
        Steps = steps;
    }

    /// <summary>False when nothing was suspected and the ladder was never entered.</summary>
    public bool Entered { get; }

    /// <summary>True when the run must fail.</summary>
    public bool Failed { get; }

    /// <summary>The failures that survived everything. Empty on a pass.</summary>
    public IReadOnlyList<TripwireFailure> Confirmed { get; }

    /// <summary>
    /// How many re-censuses were spent, never more than <see cref="TripwireRetryLadder.MaxReCensuses"/>.
    /// </summary>
    public int ReCensuses { get; }

    /// <summary>
    /// How many re-runs were spent, never more than <see cref="TripwireRetryLadder.MaxImplicatedReRuns"/>.
    /// </summary>
    public int ReRuns { get; }

    /// <summary>
    /// True when a delta was still there after at least one re-census - loud even when the
    /// final verdict is a pass, because that is the reading that stopped being noise.
    /// </summary>
    public bool SurvivedARecensus { get; }

    /// <summary>True when a bound was spent in full rather than the ladder stopping early.</summary>
    public bool BoundReached { get; }

    /// <summary>Attempt by attempt: what was suspected, what persisted, what cleared, what was new.</summary>
    public IReadOnlyList<string> Steps { get; }

    /// <summary>The headline a reader sees first, and the only part that is ever shouted.</summary>
    public string Headline
    {
        get
        {
            if (!Entered)
            {
                return "[tripwire] no retry: the post-run census found nothing to confirm.";
            }

            if (Failed)
            {
                return "[tripwire] RETRIED AND STILL FAILING: a suspected loss survived "
                    + Plural(ReCensuses, "re-census", "re-censuses") + " and "
                    + Plural(ReRuns, "re-run", "re-runs") + ".";
            }

            return SurvivedARecensus
                ? "[tripwire] RETRIED AND PASSED, BUT NOT ON THE FIRST READING: a suspected loss survived at "
                    + "least one re-census before it cleared. Read the attempts below before trusting this pass."
                : "[tripwire] RETRIED AND PASSED: the suspected loss was not reproducible on a second census.";
        }
    }

    /// <summary>One line for the run summary, so a clean run and a retried one never read alike.</summary>
    public string Summary
    {
        get
        {
            if (!Entered)
            {
                return "retry: none needed";
            }

            return "retry: " + Count(ReCensuses) + " re-census(es), " + Count(ReRuns) + " re-run(s), verdict "
                + (Failed ? "FAILED" : SurvivedARecensus ? "PASSED (survived a re-census)" : "PASSED");
        }
    }

    /// <summary>A run where the post-run census found nothing, so no retry was owed.</summary>
    public static TripwireRetryReport NotNeeded()
    {
        return new TripwireRetryReport(
            entered: false,
            failed: false,
            confirmed: Array.Empty<TripwireFailure>(),
            reCensuses: 0,
            reRuns: 0,
            survivedARecensus: false,
            boundReached: false,
            steps: Array.Empty<string>());
    }

    /// <summary>The whole record: headline, then every attempt, one per line.</summary>
    public string Describe()
    {
        return Entered
            ? Headline + Environment.NewLine + string.Join(Environment.NewLine, Steps.Select(step => "  " + step))
            : Headline;
    }

    internal static TripwireRetryReport Cleared(
        int reCensuses, int reRuns, bool survivedARecensus, bool boundReached, IReadOnlyList<string> steps)
    {
        return new TripwireRetryReport(
            entered: true,
            failed: false,
            confirmed: Array.Empty<TripwireFailure>(),
            reCensuses,
            reRuns,
            survivedARecensus,
            boundReached,
            steps);
    }

    internal static TripwireRetryReport Confirms(
        IReadOnlyList<TripwireFailure> confirmed,
        int reCensuses,
        int reRuns,
        bool boundReached,
        IReadOnlyList<string> steps)
    {
        return new TripwireRetryReport(
            entered: true,
            failed: true,
            confirmed,
            reCensuses,
            reRuns,
            survivedARecensus: true,
            boundReached,
            steps);
    }

    private static string Plural(int value, string one, string many)
    {
        return Count(value) + " " + (value == 1 ? one : many);
    }

    private static string Count(int value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }
}

/// <summary>
/// What the tripwire does about a SUSPECTED loss, bounded so it can neither cry wolf nor retry
/// its way to a green run.
/// <para>
/// <b>Two questions, two experiments.</b> A census that reports items missing is answering
/// "is this reading real?" and "did the suite do it?" at once, and they are not the same
/// question. Re-censusing answers the first: COM enumeration under a busy Outlook is not
/// perfectly repeatable, and a failure that is gone when the same baseline is compared again
/// seconds later was describing enumeration noise. Re-RUNNING the implicated tests answers the
/// second: a person reading their own mail during a 27-minute run produces a one-off delta,
/// while a test that deletes something reproduces it on demand.
/// </para>
/// <para>
/// <b>Why both are bounded, and bounded tightly.</b> Every retry is a chance to convert a real
/// loss into a pass. The ladder is therefore a ceiling rather than a budget to be spent: it
/// stops the moment two censuses agree, it says out loud that it retried and what changed, and
/// reaching a bound is an escalation with its own message rather than a quiet give-up. What
/// makes it affordable is measured - a whole-profile census is 16.9 s for 5 stores / 159
/// folders / 2,044 items, so two of them are ~34 s of work (plus two
/// <see cref="ReCensusGapSeconds"/>-second gaps) against a tier run of about 27 minutes.
/// </para>
/// <para>
/// <b>The ladder.</b> The post-run census fails -&gt; up to <see cref="MaxReCensuses"/>
/// re-censuses, <see cref="ReCensusGapSeconds"/> s apart, each intersecting the survivors by
/// <see cref="TripwireFailure.Key"/> -&gt; if anything survives all of them, up to
/// <see cref="MaxImplicatedReRuns"/> re-run of the plausibly-implicated tests -&gt; if it is
/// still there, fail loudly. There is no silent exit: the only two ways out that pass are a
/// survivor set that emptied and a re-run that did not reproduce it, and both are recorded in
/// <see cref="TripwireRetryReport.Steps"/>.
/// </para>
/// </summary>
public static class TripwireRetryLadder
{
    /// <summary>
    /// How many extra censuses a suspected loss may be given. Two, because the question a
    /// re-census answers is "does this reading repeat?", and a third reading of the same answer
    /// buys nothing while costing one more chance to pass by accident.
    /// </summary>
    public const int MaxReCensuses = 2;

    /// <summary>
    /// Gap between censuses. Long enough for an item caught mid-move, a folder mid-sync or a
    /// rule mid-file to settle; short enough to be noise against a 27-minute run.
    /// </summary>
    public const int ReCensusGapSeconds = 30;

    /// <summary>
    /// How many times the plausibly-implicated tests may be run again. One: the re-run is an
    /// experiment, and an experiment repeated until it gives the answer you wanted is not
    /// evidence.
    /// </summary>
    public const int MaxImplicatedReRuns = 1;

    /// <summary>The gap as a <see cref="TimeSpan"/>, for the source that has to wait it out.</summary>
    public static TimeSpan ReCensusGap => TimeSpan.FromSeconds(ReCensusGapSeconds);

    /// <summary>
    /// Runs the ladder over one suspected loss and returns what it decided, together with
    /// everything it did to decide it.
    /// </summary>
    /// <param name="suspected">The failing post-run comparison. A clean verdict is rejected.</param>
    /// <param name="source">Where the extra censuses and the re-run come from.</param>
    public static TripwireRetryReport Resolve(TripwireVerdict suspected, ITripwireRetrySource source)
    {
        ArgumentNullException.ThrowIfNull(suspected);
        ArgumentNullException.ThrowIfNull(source);
        if (!suspected.Failed)
        {
            throw new ArgumentException(
                "The retry ladder is only entered on a suspected loss. Re-censusing a clean verdict would spend "
                + "the bounds on nothing and could only ever turn a pass into a failure.",
                nameof(suspected));
        }

        List<TripwireFailure> survivors = Distinct(suspected.FailureRecords);
        HashSet<string> everSeen = new(survivors.Select(f => f.Key), StringComparer.Ordinal);
        List<string> steps = new()
        {
            "post-run census: " + Count(survivors.Count) + " suspected failure(s) - " + Keys(survivors) + ".",
        };

        int reCensuses = 0;
        bool survivedARecensus = false;
        while (survivors.Count > 0 && reCensuses < MaxReCensuses)
        {
            source.Wait(ReCensusGap);
            reCensuses++;
            TripwireVerdict again = source.ReCensus(reCensuses);
            HashSet<string> nowKeys = new(again.FailureRecords.Select(f => f.Key), StringComparer.Ordinal);
            List<TripwireFailure> persisted = survivors.Where(f => nowKeys.Contains(f.Key)).ToList();
            List<TripwireFailure> cleared = survivors.Where(f => !nowKeys.Contains(f.Key)).ToList();
            List<string> appeared = new();
            foreach (TripwireFailure failure in again.FailureRecords)
            {
                if (everSeen.Add(failure.Key))
                {
                    appeared.Add(failure.Key);
                }
            }

            steps.Add(
                "re-census " + Count(reCensuses) + " of " + Count(MaxReCensuses) + " (~"
                + Count(ReCensusGapSeconds) + " s later): "
                + (persisted.Count > 0 ? "PERSISTED " : "persisted ") + Count(persisted.Count) + " of "
                + Count(survivors.Count) + "; cleared: " + (cleared.Count == 0 ? "none" : Keys(cleared))
                + "; new since the post-run census: "
                + (appeared.Count == 0 ? "none" : string.Join(", ", appeared)) + ".");

            survivors = persisted;
            survivedARecensus = survivedARecensus || survivors.Count > 0;
        }

        if (survivors.Count == 0)
        {
            steps.Add(
                "two censuses agree: nothing that failed the post-run census was still failing "
                + Count(reCensuses) + " reading(s) later, so this was enumeration noise and the run passes.");
            return TripwireRetryReport.Cleared(
                reCensuses, reRuns: 0, survivedARecensus, boundReached: false, steps);
        }

        steps.Add(
            "BOUND REACHED: all " + Count(MaxReCensuses) + " re-census(es) are spent and " + Count(survivors.Count)
            + " failure(s) survived every one - " + Keys(survivors)
            + ". Escalating to a bounded re-run; this is an escalation, not a give-up.");

        IReadOnlyList<string> implicated = source.ImplicatedBy(survivors);
        if (implicated.Count == 0)
        {
            steps.Add(
                "no test could be named as plausibly implicated, so there is nothing to re-run and nothing that "
                + "could exonerate the delta - FAILING.");
            return TripwireRetryReport.Confirms(survivors, reCensuses, reRuns: 0, boundReached: true, steps);
        }

        int reRuns = 0;
        TripwireReRunOutcome outcome = TripwireReRunOutcome.Inconclusive;
        while (reRuns < MaxImplicatedReRuns)
        {
            reRuns++;
            outcome = source.ReRun(implicated, reRuns);
            steps.Add(
                "re-run " + Count(reRuns) + " of " + Count(MaxImplicatedReRuns) + " over "
                + Count(implicated.Count) + " implicated selection(s) (" + string.Join(", ", implicated) + "): "
                + Describe(outcome) + ".");
            if (outcome != TripwireReRunOutcome.Inconclusive)
            {
                break;
            }
        }

        if (outcome == TripwireReRunOutcome.NotReproduced)
        {
            steps.Add(
                "the delta did not come back when the implicated tests ran again, so the suite is not what removed "
                + "it - the run passes ON THIS RECORD, and the items named above are still worth a look.");
            return TripwireRetryReport.Cleared(
                reCensuses, reRuns, survivedARecensus: true, boundReached: true, steps);
        }

        steps.Add(
            outcome == TripwireReRunOutcome.Reproduced
                ? "the same failure(s) came back when the implicated tests ran again - THIS IS THE SUITE REMOVING "
                    + "MAIL. Stop and investigate."
                : "the " + Count(MaxImplicatedReRuns) + " permitted re-run(s) produced no answer, and an experiment "
                    + "nobody could carry out exonerates nothing - FAILING.");
        return TripwireRetryReport.Confirms(survivors, reCensuses, reRuns, boundReached: true, steps);
    }

    private static string Describe(TripwireReRunOutcome outcome)
    {
        return outcome switch
        {
            TripwireReRunOutcome.Reproduced => "REPRODUCED",
            TripwireReRunOutcome.NotReproduced => "not reproduced",
            _ => "INCONCLUSIVE (the re-run could not be carried out)",
        };
    }

    /// <summary>
    /// One entry per failure KEY, first occurrence kept. Two failure lines naming the same
    /// kind, store and folder are one thing to confirm, and counting them twice would make the
    /// bounds mean different amounts of work on different runs.
    /// </summary>
    private static List<TripwireFailure> Distinct(IReadOnlyList<TripwireFailure> failures)
    {
        HashSet<string> seen = new(StringComparer.Ordinal);
        List<TripwireFailure> distinct = new(failures.Count);
        foreach (TripwireFailure failure in failures)
        {
            if (seen.Add(failure.Key))
            {
                distinct.Add(failure);
            }
        }

        return distinct;
    }

    private static string Keys(IReadOnlyList<TripwireFailure> failures)
    {
        return string.Join(", ", failures.Select(f => f.Key));
    }

    private static string Count(int value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }
}
