namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// One live test's answer to "was there anything on this machine to prove anything with?".
/// <para>
/// <b>The shape this exists to stop.</b> A live test whose body is a <c>foreach</c> over a
/// machine-dependent list, or an <c>Assert.All</c> over one, asserts NOTHING when that list is
/// empty and reports GREEN - indistinguishable, in every report it appears in, from a run that
/// exercised the path. xunit 2.9.3 has no assertion-counting hook, so the empty case really is a
/// pass and not a skip; that was confirmed rather than assumed when the same defect was found in
/// the identity-draft pair (<see cref="IdentityDraftCoverage"/>) on 2026-08-25, and reading every
/// <c>foreach</c> in the live tier the same day found three more.
/// </para>
/// <para>
/// <b>The idiom, which is this repository's own and not a new one:</b>
/// <see cref="LiveTestSettings.RequireProductionPopulation"/> plus a <c>PROVED NOTHING:</c> line.
/// On a Production profile an empty population means the machine or the settings have drifted from
/// what the tests were written for, and the run refuses. On a Portable one it is simply true of the
/// machine - the test names what it needs under <c>Requires</c> and should not have been selected -
/// so the run says so in a line no reader can mistake for a pass. Making it a hard failure
/// everywhere was considered and rejected when the identity pair was fixed: it fails machines these
/// tests were never meant to run on, for a property of the machine.
/// </para>
/// <para>
/// <b>Why it is generic and <see cref="IdentityDraftCoverage"/> is not.</b> That one has a real
/// decision inside it - which stores the write allowlist grants a draft in, and why each other one
/// was withheld - and the answer has to come from the allowlist itself rather than from a second
/// opinion about it. This one has no such decision: the population is handed in, already computed
/// by whoever owns it, and all that is decided here is what the run SAYS about its being empty.
/// One place to decide that is the point; a second copy of "if the list is empty, print something"
/// is how the first three copies came to disagree.
/// </para>
/// <para>
/// Pure: strings and a count in, strings out, no COM, no settings file, no mailbox. That is
/// deliberate - the callers are <c>Category=Live</c> and CI can never run them, so the decision
/// lives here where every branch of it is reachable from a runner with no Outlook, and the live
/// side is one call.
/// </para>
/// </summary>
public static class LivePopulationCoverage
{
    /// <summary>
    /// Classifies a population without reporting anything or refusing anything - the pure half,
    /// for callers that want the lines and will decide for themselves what to do with them.
    /// </summary>
    /// <param name="population">
    /// What was looked for, named as a reader would name it and as a NOUN PHRASE: the Production
    /// refusal wraps it ("...where &lt;this&gt; is expected to exist").
    /// </param>
    /// <param name="whatWouldNotRun">What the caller was about to do with it.</param>
    /// <param name="found">How many there were. Never negative.</param>
    /// <param name="remedy">
    /// One sentence telling whoever reads the log how to give this machine the population. Required
    /// rather than optional: a PROVED NOTHING line that does not say what to do about it is a line
    /// people learn to skip.
    /// </param>
    public static LivePopulationReport Assess(
        string population, string whatWouldNotRun, int found, string remedy)
    {
        return new LivePopulationReport(population, whatWouldNotRun, found, remedy);
    }

    /// <summary>
    /// The coverage line, the refusal and the announcement, in the one order they may happen -
    /// and the population back, so this is the whole of what a caller has to write.
    /// </summary>
    /// <param name="settings">The machine's live-test settings; only its profile is read.</param>
    /// <param name="population">The list the caller was about to iterate. Null counts as empty.</param>
    /// <param name="populationName">What that list is, as <see cref="Assess"/> documents it.</param>
    /// <param name="whatWouldNotRun">What the caller was about to do.</param>
    /// <param name="remedy">How to give this machine the population.</param>
    /// <param name="report">Where the lines go - normally <c>ITestOutputHelper.WriteLine</c>.</param>
    /// <returns>
    /// The population, unchanged and never null. Empty only after the run has been told so in
    /// writing, and on a Production profile never at all.
    /// </returns>
    public static IReadOnlyList<T> Require<T>(
        LiveTestSettings settings,
        IReadOnlyList<T>? population,
        string populationName,
        string whatWouldNotRun,
        string remedy,
        Action<string> report)
    {
        ArgumentNullException.ThrowIfNull(settings);
        ArgumentNullException.ThrowIfNull(report);

        IReadOnlyList<T> found = population ?? Array.Empty<T>();
        LivePopulationReport coverage = Assess(populationName, whatWouldNotRun, found.Count, remedy);

        // Printed on EVERY run and not only the empty ones: a reader should be able to tell from
        // the log how much a passing test actually visited, rather than inferring it from the
        // test's name. That inference is exactly what was wrong before.
        report(coverage.Describe());
        if (!coverage.ProvesNothing)
        {
            return found;
        }

        // Throws on Production - an empty population there means drift, and a green test hides it.
        // No-ops on Portable, where the line below is the whole point. The order matters: nothing
        // on a Production machine may read as an acceptable emptiness, so the refusal happens
        // BEFORE anything is announced.
        settings.RequireProductionPopulation(populationName);
        report(coverage.ProvedNothing());
        return found;
    }
}

/// <summary>
/// One machine's answer to "how much of this test could possibly have run" - see
/// <see cref="LivePopulationCoverage"/> for why it is asked at all.
/// </summary>
public sealed class LivePopulationReport
{
    internal LivePopulationReport(string population, string whatWouldNotRun, int found, string remedy)
    {
        // Named, not defaulted. A guard that reads as coverage in every report it appears in must
        // say WHICH population and WHICH test, or the line it prints is worth less than nothing;
        // and an unnamed population would reach RequireProductionPopulation's message as a blank.
        Require(population, nameof(population));
        Require(whatWouldNotRun, nameof(whatWouldNotRun));
        Require(remedy, nameof(remedy));
        if (found < 0)
        {
            throw new ArgumentOutOfRangeException(
                nameof(found), found, "a population cannot have a negative size.");
        }

        Population = population;
        WhatWouldNotRun = whatWouldNotRun;
        Found = found;
        Remedy = remedy;
    }

    /// <summary>What was looked for.</summary>
    public string Population { get; }

    /// <summary>What the caller was about to do with it.</summary>
    public string WhatWouldNotRun { get; }

    /// <summary>How many there were.</summary>
    public int Found { get; }

    /// <summary>How to give this machine the population.</summary>
    public string Remedy { get; }

    /// <summary>
    /// True when the caller would iterate nothing - the state that used to report green.
    /// </summary>
    public bool ProvesNothing => Found == 0;

    /// <summary>The coverage line, printed on every run.</summary>
    public string Describe()
    {
        return "coverage: " + Found + " " + Population + " on this machine - that is what "
            + WhatWouldNotRun + " iterates.";
    }

    /// <summary>
    /// The line a run prints instead of quietly passing. Written for whoever reads the log
    /// afterwards: what did not run, why there was nothing to run it against, what a green result
    /// here does and does not mean, and what would change it.
    /// </summary>
    public string ProvedNothing()
    {
        return "PROVED NOTHING: " + WhatWouldNotRun + " iterated nothing - this machine has no "
            + Population + ". Nothing this test asserts about that path was verified here, so a "
            + "green result on this machine says only that there was nothing to verify it "
            + "against. " + Remedy;
    }

    private static void Require(string value, string name)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new ArgumentException("a coverage line needs " + name + " spelled out.", name);
        }
    }
}
