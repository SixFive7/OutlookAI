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
/// What one whole trip through the ladder means for the RUN, as opposed to what it means for
/// the reader. Three values rather than a boolean, because two of them read as a pass and only
/// one of them may leave the process with a zero exit code.
/// </summary>
public enum TripwireRunOutcome
{
    /// <summary>
    /// Nothing was suspected, or a suspected delta was gone on the very first re-reading. The
    /// only value that lets a run exit zero.
    /// </summary>
    Passed = 0,

    /// <summary>
    /// A suspected loss SURVIVED at least one re-census - so the items really were gone at two
    /// separate readings - and then cleared, either on a later census or because the bounded
    /// re-run did not reproduce it.
    /// <para>
    /// <b>A pass for the reader and a failure for automation, deliberately.</b> This is the one
    /// place in the design where a real loss of mail could end as a green run: the delta is
    /// real (two readings saw it), the re-run cannot pin it on the suite, and every other check
    /// in the tier is happy. The maintainer's standing rule is to fail aggressively rather than
    /// leak a slow degradation, so the headline says PASSED WITH A SURVIVED DELTA and the run
    /// still exits non-zero.
    /// </para>
    /// </summary>
    PassedWithASurvivedDelta = 1,

    /// <summary>The delta survived everything the ladder could spend on it.</summary>
    Failed = 2,
}

/// <summary>
/// The three bounds the ladder is allowed to spend, as one value, so a machine can be given
/// fewer of them than another WITHOUT any of the ladder's own rules moving.
/// <para>
/// <b>Why this is per machine and not per run.</b> The ladder's two rungs answer questions that
/// only exist on a mailbox something OTHER than the suite can change. Re-censusing asks "did the
/// mailbox settle?", and re-running asks "was it the suite or was it the person?" - and on a
/// dedicated test machine there is no person, no server-side rule, no retention policy and no
/// Exchange sync: the stores are PSTs, nobody is at the keyboard during a run, and the only
/// transport is a loopback sink. There is nothing for those two questions to distinguish, so
/// the honest bound there is ZERO and the first reading is the verdict.
/// </para>
/// <para>
/// <b>An unknown profile gets NO retries.</b> Every retry is a chance to convert a real loss
/// into a pass, so the direction that fails safe is the strict one: only the profile that
/// explicitly declares itself <see cref="LiveMachineProfile.Production"/> - a real working
/// mailbox with real people and a real server in it - buys the accommodation.
/// </para>
/// </summary>
public sealed class TripwireRetryPolicy
{
    private TripwireRetryPolicy(string name, int maxReCensuses, int reCensusGapSeconds, int maxImplicatedReRuns)
    {
        if (maxReCensuses < 0 || reCensusGapSeconds < 0 || maxImplicatedReRuns < 0)
        {
            throw new ArgumentException("A retry policy cannot spend a negative number of anything.", nameof(name));
        }

        Name = name;
        MaxReCensuses = maxReCensuses;
        ReCensusGapSeconds = reCensusGapSeconds;
        MaxImplicatedReRuns = maxImplicatedReRuns;
    }

    /// <summary>
    /// The bounds a real working mailbox needs: 2 re-censuses ~30 s apart, then at most 1
    /// bounded re-run. The numbers themselves stay on <see cref="TripwireRetryLadder"/>, which
    /// is where they are documented and pinned.
    /// </summary>
    public static TripwireRetryPolicy Production { get; } = new(
        "Production",
        TripwireRetryLadder.MaxReCensuses,
        TripwireRetryLadder.ReCensusGapSeconds,
        TripwireRetryLadder.MaxImplicatedReRuns);

    /// <summary>
    /// No retries at all: the post-run census is the verdict. What a machine gets when nothing
    /// but the suite can change a mailbox on it, and what an UNDECLARED machine gets too.
    /// </summary>
    public static TripwireRetryPolicy None { get; } = new("None", 0, 0, 0);

    /// <summary>How this policy is named in the run's own output.</summary>
    public string Name { get; }

    /// <summary>Extra censuses a suspected loss may be given.</summary>
    public int MaxReCensuses { get; }

    /// <summary>Seconds between two censuses.</summary>
    public int ReCensusGapSeconds { get; }

    /// <summary>Times the plausibly-implicated tests may be run again.</summary>
    public int MaxImplicatedReRuns { get; }

    /// <summary>The gap as a <see cref="TimeSpan"/>, for the source that has to wait it out.</summary>
    public TimeSpan ReCensusGap => TimeSpan.FromSeconds(ReCensusGapSeconds);

    /// <summary>True when this policy would spend anything at all on a suspected loss.</summary>
    public bool RetriesAtAll => MaxReCensuses > 0 || MaxImplicatedReRuns > 0;

    /// <summary>
    /// The policy for one machine. Only a declared <see cref="LiveMachineProfile.Production"/>
    /// profile gets retries; everything else - <see cref="LiveMachineProfile.Portable"/>, and any
    /// value added later that nobody has thought about - gets <see cref="None"/>.
    /// </summary>
    public static TripwireRetryPolicy For(LiveMachineProfile profile)
    {
        return profile == LiveMachineProfile.Production ? Production : None;
    }

    /// <summary>
    /// The policy for one machine, with the recursion guard applied.
    /// <para>
    /// A process that IS the bounded re-run must not start a bounded re-run of its own, or the
    /// ladder would recurse one live tier run at a time until something ran out. It keeps its
    /// censuses - they are cheap and they stop the child crying wolf over a transient - and
    /// loses the re-run rung, which is strictly stricter: a child with no re-run available
    /// reports a delta it cannot clear, and the parent reads that as the delta coming back.
    /// </para>
    /// </summary>
    public static TripwireRetryPolicy For(LiveMachineProfile profile, bool isReRunChild)
    {
        TripwireRetryPolicy policy = For(profile);
        return isReRunChild ? policy.WithoutReRuns() : policy;
    }

    /// <summary>The same censuses, no re-runs.</summary>
    public TripwireRetryPolicy WithoutReRuns()
    {
        return MaxImplicatedReRuns == 0
            ? this
            : new TripwireRetryPolicy(Name + " (re-run child)", MaxReCensuses, ReCensusGapSeconds, 0);
    }

    /// <summary>One line for the run's output, so a run says which bounds it was given.</summary>
    public string Describe()
    {
        return RetriesAtAll
            ? "policy " + Name + ": at most " + MaxReCensuses + " re-census(es) ~" + ReCensusGapSeconds
                + " s apart, then at most " + MaxImplicatedReRuns + " bounded re-run(s)"
            : "policy " + Name + ": NO RETRIES - the post-run census is the verdict";
    }
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
        TripwireRunOutcome outcome,
        IReadOnlyList<TripwireFailure> confirmed,
        int reCensuses,
        int reRuns,
        bool survivedARecensus,
        bool boundReached,
        IReadOnlyList<string> steps)
    {
        // THE STRUCTURAL GUARANTEE, and the reason it lives in the constructor rather than in
        // the code that builds reports. A report is the only thing the live tripwire consults
        // when it decides whether to throw, and Failed is DERIVED from Outcome rather than
        // stored beside it - so the only way to produce a run that exits zero is to produce a
        // report whose Outcome is Passed. This invariant then makes that impossible whenever a
        // delta survived a reading: not "we remembered to fail it" but "such a report cannot be
        // constructed". Both directions, because either lie is the same lie.
        if (survivedARecensus && outcome == TripwireRunOutcome.Passed)
        {
            throw new ArgumentException(
                "A tripwire retry report cannot say a suspected loss SURVIVED a re-census and also report a "
                + "clean pass. A delta seen at two separate readings is real, and the one shape in this design "
                + "where real mail loss could end as a green run is exactly this one. Use "
                + nameof(TripwireRunOutcome.PassedWithASurvivedDelta) + ", which reads as a pass and still "
                + "fails the run.",
                nameof(outcome));
        }

        if (!survivedARecensus && outcome == TripwireRunOutcome.PassedWithASurvivedDelta)
        {
            throw new ArgumentException(
                "A tripwire retry report cannot claim a survived delta when nothing survived a re-census. "
                + "The headline would tell a reader to go looking for items that were never missing twice.",
                nameof(outcome));
        }

        Entered = entered;
        Outcome = outcome;
        Confirmed = confirmed;
        ReCensuses = reCensuses;
        ReRuns = reRuns;
        SurvivedARecensus = survivedARecensus;
        BoundReached = boundReached;
        Steps = steps;
    }

    /// <summary>False when nothing was suspected and the ladder was never entered.</summary>
    public bool Entered { get; }

    /// <summary>What this trip means for the run. See <see cref="TripwireRunOutcome"/>.</summary>
    public TripwireRunOutcome Outcome { get; }

    /// <summary>
    /// True when the run must fail. DERIVED, never stored: exactly one of the three outcomes
    /// lets a run exit zero, so adding a fourth outcome fails the run until somebody decides
    /// otherwise on purpose, and no report can carry a verdict that disagrees with itself.
    /// </summary>
    public bool Failed => Outcome != TripwireRunOutcome.Passed;

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

            if (Outcome == TripwireRunOutcome.Failed)
            {
                return "[tripwire] RETRIED AND STILL FAILING: a suspected loss survived "
                    + Plural(ReCensuses, "re-census", "re-censuses") + " and "
                    + Plural(ReRuns, "re-run", "re-runs") + ".";
            }

            return Outcome == TripwireRunOutcome.PassedWithASurvivedDelta
                ? "[tripwire] " + SurvivedDeltaHeadline + ": a suspected loss survived at least one re-census "
                    + "before it cleared, so the items really were gone at two separate readings. The retry "
                    + "ladder's verdict is a PASS and THE RUN STILL EXITS NON-ZERO - a delta that survived a "
                    + "reading is the one shape in this design where real mail loss could end as a green run. "
                    + "Read the attempts below; the items named there are still worth a look."
                : "[tripwire] RETRIED AND PASSED: the suspected loss was not reproducible on a second census.";
        }
    }

    /// <summary>
    /// The phrase a human reads for <see cref="TripwireRunOutcome.PassedWithASurvivedDelta"/>,
    /// named once so the report, the refusal message and the pin cannot spell it differently.
    /// </summary>
    public const string SurvivedDeltaHeadline = "PASSED WITH A SURVIVED DELTA";

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
                + Outcome switch
                {
                    TripwireRunOutcome.Failed => "FAILED",
                    TripwireRunOutcome.PassedWithASurvivedDelta => SurvivedDeltaHeadline + " (fails the run)",
                    _ => "PASSED",
                };
        }
    }

    /// <summary>A run where the post-run census found nothing, so no retry was owed.</summary>
    public static TripwireRetryReport NotNeeded()
    {
        return new TripwireRetryReport(
            entered: false,
            outcome: TripwireRunOutcome.Passed,
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

    /// <summary>
    /// A trip that ended without a confirmed loss. The OUTCOME is not a parameter: it is
    /// derived from whether anything survived a re-census, so "cleared" and "clean" are one
    /// decision made in one place rather than two flags that can drift apart.
    /// </summary>
    internal static TripwireRetryReport Cleared(
        int reCensuses, int reRuns, bool survivedARecensus, bool boundReached, IReadOnlyList<string> steps)
    {
        return new TripwireRetryReport(
            entered: true,
            outcome: survivedARecensus
                ? TripwireRunOutcome.PassedWithASurvivedDelta
                : TripwireRunOutcome.Passed,
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
        bool survivedARecensus,
        bool boundReached,
        IReadOnlyList<string> steps)
    {
        return new TripwireRetryReport(
            entered: true,
            outcome: TripwireRunOutcome.Failed,
            confirmed,
            reCensuses,
            reRuns,
            survivedARecensus,
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
/// The single place that decides whether a verified live run may finish quietly, and what it
/// says when it may not.
/// <para>
/// It exists as a pure function for one reason: the line that acts on it - the throw at the end
/// of <c>LiveStoreCountTripwire.Verify</c> - sits behind a COM census that no CI test can
/// execute, so the DECISION is moved somewhere CI can drive every combination of it and only
/// the throw is left behind the census. T1 pins the decision by value and reads the call back
/// out of the compiled method.
/// </para>
/// </summary>
public static class TripwireRunVerdict
{
    /// <summary>
    /// The message a run must refuse with, or null when it may finish quietly.
    /// <para>
    /// Both halves are consulted, and the second is the one that matters here: a retry report
    /// can fail a run that the rebuilt verdict calls clean. That is exactly the
    /// <see cref="TripwireRunOutcome.PassedWithASurvivedDelta"/> case - nothing was CONFIRMED
    /// as lost, and a delta was nevertheless real at two separate readings.
    /// </para>
    /// </summary>
    public static string? RefusalFor(TripwireVerdict verdict, TripwireRetryReport retry)
    {
        ArgumentNullException.ThrowIfNull(verdict);
        ArgumentNullException.ThrowIfNull(retry);
        if (!verdict.Failed && !retry.Failed)
        {
            return null;
        }

        return verdict.Failed
            ? verdict.Describe() + Environment.NewLine + retry.Describe()
            : retry.Describe();
    }

    /// <summary>
    /// Refuses the run, or returns quietly.
    /// <para>
    /// <b>The decision and the acting on it are one call on purpose.</b> The caller is
    /// <c>LiveStoreCountTripwire.Verify</c>, which sits behind a COM census no CI test can
    /// execute, so any <c>if</c> left in it is a condition a mutation can invert where nothing
    /// would notice - measured: inverting exactly that condition passed the whole suite. There
    /// is now no such condition there; the branch lives here, where CI drives both sides of it,
    /// and all a mutation can do up there is remove the call, which is read out of the IL.
    /// </para>
    /// </summary>
    public static void Enforce(TripwireVerdict verdict, TripwireRetryReport retry)
    {
        string? refusal = RefusalFor(verdict, retry);
        if (refusal != null)
        {
            throw new InvalidOperationException(refusal);
        }
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
/// <b>Both rungs are a PRODUCTION accommodation, and the bounds say so.</b> Re-censusing and
/// re-running exist to separate the suite from everything ELSE that can change a real mailbox
/// during a 27-minute run: a person reading their own mail, a server-side rule, an Exchange
/// sync, a retention policy, a delegate store whose folder hierarchy syncs lazily. A dedicated
/// test machine has none of those - PST stores, nobody at the keyboard, a loopback sink for
/// transport - so there is nothing for the two experiments to tell apart and the honest bound
/// there is zero. <see cref="TripwireRetryPolicy"/> carries the numbers per machine profile;
/// this class carries the rules, and they do not change with the numbers.
/// </para>
/// <para>
/// <b>A pass that took a retry is not a green run.</b> A delta that survives even one re-census
/// was real at two separate readings, so clearing it later reports
/// <see cref="TripwireRunOutcome.PassedWithASurvivedDelta"/> - which reads as a pass and still
/// fails the run. The invariant is enforced in <see cref="TripwireRetryReport"/>'s constructor
/// rather than by remembering to set a flag: a report that says a delta survived and also
/// reports a clean pass cannot be built.
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
    /// How many extra censuses a suspected loss may be given ON A PRODUCTION PROFILE. Two,
    /// because the question a re-census answers is "does this reading repeat?", and a third
    /// reading of the same answer buys nothing while costing one more chance to pass by
    /// accident. A machine that declares any other profile gets <see cref="TripwireRetryPolicy.None"/>.
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
        return Resolve(suspected, source, TripwireRetryPolicy.Production);
    }

    /// <summary>
    /// The same ladder under a stated set of bounds. Every rule is unchanged; only how many
    /// times each rung may be climbed comes from <paramref name="policy"/>, so a machine where
    /// nothing but the suite can change a mailbox can be given zero of both and have its first
    /// reading be the verdict.
    /// </summary>
    /// <param name="suspected">The failing post-run comparison. A clean verdict is rejected.</param>
    /// <param name="source">Where the extra censuses and the re-run come from.</param>
    /// <param name="policy">The bounds this machine is allowed to spend.</param>
    public static TripwireRetryReport Resolve(
        TripwireVerdict suspected, ITripwireRetrySource source, TripwireRetryPolicy policy)
    {
        ArgumentNullException.ThrowIfNull(suspected);
        ArgumentNullException.ThrowIfNull(source);
        ArgumentNullException.ThrowIfNull(policy);
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
            "post-run census: " + Count(survivors.Count) + " suspected failure(s) - " + Keys(survivors) + ". "
                + policy.Describe() + ".",
        };

        int reCensuses = 0;
        bool survivedARecensus = false;
        while (survivors.Count > 0 && reCensuses < policy.MaxReCensuses)
        {
            source.Wait(policy.ReCensusGap);
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
                "re-census " + Count(reCensuses) + " of " + Count(policy.MaxReCensuses) + " (~"
                + Count(policy.ReCensusGapSeconds) + " s later): "
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
            policy.MaxReCensuses == 0
                ? "NO RE-CENSUS IS CONFIGURED on this machine profile (" + policy.Name + "), so the post-run "
                    + "census is the verdict: " + Count(survivors.Count) + " failure(s) - " + Keys(survivors)
                    + ". Nothing but the suite changes a mailbox here, so a count change outside the hub is a "
                    + "fault by definition and there is no second reading that could make it not one."
                : "BOUND REACHED: all " + Count(policy.MaxReCensuses) + " re-census(es) are spent and "
                    + Count(survivors.Count) + " failure(s) survived every one - " + Keys(survivors)
                    + ". Escalating to a bounded re-run; this is an escalation, not a give-up.");

        if (policy.MaxImplicatedReRuns == 0)
        {
            // Short-circuited BEFORE ImplicatedBy, so a machine with no re-run rung never asks
            // a question it has no way of answering. Strictly stricter than the alternative:
            // there is no path from here that does not fail.
            steps.Add(
                "no re-run is permitted under this policy, and an experiment nobody performed exonerates "
                + "nothing - FAILING.");
            return TripwireRetryReport.Confirms(
                survivors, reCensuses, reRuns: 0, survivedARecensus, boundReached: true, steps);
        }

        IReadOnlyList<string> implicated = source.ImplicatedBy(survivors);
        if (implicated.Count == 0)
        {
            steps.Add(
                "no test could be named as plausibly implicated, so there is nothing to re-run and nothing that "
                + "could exonerate the delta - FAILING.");
            return TripwireRetryReport.Confirms(
                survivors, reCensuses, reRuns: 0, survivedARecensus, boundReached: true, steps);
        }

        int reRuns = 0;
        TripwireReRunOutcome outcome = TripwireReRunOutcome.Inconclusive;
        while (reRuns < policy.MaxImplicatedReRuns)
        {
            reRuns++;
            outcome = source.ReRun(implicated, reRuns);
            steps.Add(
                "re-run " + Count(reRuns) + " of " + Count(policy.MaxImplicatedReRuns) + " over "
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
                + "it - the ladder's verdict is a pass ON THIS RECORD. It survived a re-census, so the run STILL "
                + "FAILS and the items named above are the thing to look at.");
            return TripwireRetryReport.Cleared(
                reCensuses, reRuns, survivedARecensus: true, boundReached: true, steps);
        }

        steps.Add(
            outcome == TripwireReRunOutcome.Reproduced
                ? "the same failure(s) came back when the implicated tests ran again - THIS IS THE SUITE REMOVING "
                    + "MAIL. Stop and investigate."
                : "the " + Count(policy.MaxImplicatedReRuns)
                    + " permitted re-run(s) produced no answer, and an experiment "
                    + "nobody could carry out exonerates nothing - FAILING.");
        return TripwireRetryReport.Confirms(
            survivors, reCensuses, reRuns, survivedARecensus, boundReached: true, steps);
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
