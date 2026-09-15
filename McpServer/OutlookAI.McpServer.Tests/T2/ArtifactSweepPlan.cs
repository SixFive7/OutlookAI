namespace OutlookAI.McpServer.Tests.T2;

/// <summary>What the post-run artifact sweep is entitled to do to ONE store.</summary>
public enum ArtifactSweepAction
{
    /// <summary>
    /// Count, purge whatever is found, count again, and require zero. The ordinary case: a store
    /// the suite writes to, swept because the suite is what put the artifacts there.
    /// </summary>
    Sweep = 0,

    /// <summary>
    /// Count, and nothing else. A store the write allowlist refuses - so a tagged item in it is
    /// not a cleanup job the sweep should finish, it is a finding the sweep should report.
    /// </summary>
    CountOnly = 1,
}

/// <summary>
/// The post-run artifact sweep, decided store by store before any COM call.
///
/// <para>
/// <b>What the sweep is.</b> Mailbox-safety rule 4: every live run must end with zero items
/// carrying the live-tier subject tag, proven by walking every store the census watches
/// (<c>expectedStoreDisplayNames</c>) and counting. Self-send copies materialise with lag, so a
/// store the suite may write to gets one purge pass before its count is believed.
/// </para>
///
/// <para>
/// <b>The defect this closes.</b> That walk used to purge unconditionally: any store whose count
/// came back non-zero was handed to <c>DeleteTaggedArtifactsUntilStableZero</c>. But
/// <c>expectedStoreDisplayNames</c> is also where the measurement corpus is declared - it has to
/// be, or the count tripwire never censuses it - and the corpus store is a declared BYSTANDER,
/// the tier the write allowlist refuses every kind of write to. Until the subject tags were split
/// on 2026-08-25 the corpus carried the same tag as the live tier's artifacts, and that sweep was
/// one non-zero count away from deleting twenty thousand real items. The tag split removed the
/// match. It did not remove the AIM: the code still pointed a delete at a store no test may write
/// to, and what stopped it was a property of the data rather than of the code.
/// </para>
///
/// <para>
/// <b>What was chosen, and over what.</b> Not a skip - skipping the bystanders gives up the one
/// thing worth knowing, which is whether an artifact ever reached one. Not "leave it to the count
/// tripwire" either, and that argument is worth writing down because it is the obvious one: the
/// tripwire fires on a per-store item-count DECREASE, and an artifact appearing in a bystander is
/// an INCREASE. It would pass in silence. So the sweep walks every store including the
/// bystanders, COUNTS every one of them, deletes from none it may not write to, and FAILS -
/// naming the store and the count - if a count it cannot act on is non-zero.
/// </para>
///
/// <para>
/// <b>Why it is phrased as "may the allowlist delete from this store" rather than "is it
/// declared a bystander".</b> Both answers are the same one today: <c>LiveStoreWriteGuard.Build</c>
/// derives the identity-draft grant from <c>expectedStoreDisplayNames</c> itself, so every store
/// the sweep visits is the hub, a declared bystander, or a store granted delete. Asking the
/// allowlist keeps them the same answer if that ever stops being true, and it means this class can
/// never hand the purge a store <c>LiveOutlookTestMailer</c> would refuse anyway - which today is
/// what actually stops the corpus being deleted, and is a refusal thrown from inside the loop that
/// names the guard rather than the finding. The declaration is still carried separately, because
/// the two failures read differently to whoever hits them.
/// </para>
///
/// <para>
/// Pure: store names in, strings out. No COM, no settings file, no mailbox. That is deliberate and
/// it is this repository's established shape for live-tier decisions - see
/// <see cref="TripwireWatchSoundness"/> and <see cref="IdentityDraftCoverage"/>. CI can never run
/// the tier that consumes this, so the decision lives where CI can pin every branch of it and the
/// live side is one call.
/// </para>
/// </summary>
public static class ArtifactSweepPolicy
{
    /// <summary>
    /// The phrase to grep a run log for when the sweep finds a tagged item somewhere nothing was
    /// entitled to put one. In one place because the refusal text wraps it and the T1 pin greps
    /// for it.
    /// </summary>
    public const string Residue = "TAGGED ARTIFACT IN A STORE NO TEST MAY WRITE TO";

    /// <summary>The phrase for a sweep that was handed nothing to walk.</summary>
    public const string NothingToVisit = "THE ARTIFACT SWEEP WAS GIVEN NO STORE TO VISIT";

    /// <summary>Plans the sweep over the stores <paramref name="settings"/> has the census watch.</summary>
    public static ArtifactSweepPlan Assess(LiveTestSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        return Assess(settings.ExpectedStoreDisplayNames, LiveStoreWriteGuard.Build(settings));
    }

    /// <summary>
    /// Plans the sweep over <paramref name="watchedStores"/> under <paramref name="allowlist"/>:
    /// one step per distinct store, in the order given, each carrying whether this sweep may
    /// delete from it and why not when it may not.
    /// </summary>
    public static ArtifactSweepPlan Assess(
        IEnumerable<string>? watchedStores, StoreWriteAllowlist allowlist)
    {
        ArgumentNullException.ThrowIfNull(allowlist);

        List<ArtifactSweepStep> steps = new();
        HashSet<string> seen = new(StringComparer.OrdinalIgnoreCase);
        foreach (string store in watchedStores ?? [])
        {
            if (string.IsNullOrWhiteSpace(store) || !seen.Add(store))
            {
                continue;
            }

            bool mayDelete = allowlist.IsAllowed(store, StoreWriteKind.Delete);
            steps.Add(new ArtifactSweepStep(
                store,
                mayDelete ? ArtifactSweepAction.Sweep : ArtifactSweepAction.CountOnly,
                allowlist.IsBystander(store)));
        }

        return new ArtifactSweepPlan(steps);
    }

    /// <summary>
    /// Plans and runs the sweep for <paramref name="settings"/>. The whole of what the two live
    /// call sites do.
    /// </summary>
    public static ArtifactSweepPlan Run(
        LiveTestSettings settings,
        Func<string, int> countTagged,
        Action<string> purgeUntilStableZero,
        Action<string> report)
    {
        return Run(Assess(settings), countTagged, purgeUntilStableZero, report);
    }

    /// <summary>
    /// Walks <paramref name="plan"/>: counts every store, purges only the ones the plan says may
    /// be purged, re-counts those, reports one line per store either way, and throws once at the
    /// end if anything is left over anywhere.
    /// <para>
    /// The COM half arrives as delegates so this loop - the part that decides what is handed to a
    /// delete - runs in CI against fakes. <paramref name="purgeUntilStableZero"/> is called for a
    /// <see cref="ArtifactSweepAction.CountOnly"/> store under no circumstances, and
    /// <c>NeverPurgesAStoreTheAllowlistRefuses</c> is the pin that says so.
    /// </para>
    /// <para>
    /// Every refusal is collected and thrown together rather than one per run: a live run costs
    /// minutes and a mailbox, and finding the second problem only after fixing the first costs
    /// another one.
    /// </para>
    /// </summary>
    /// <param name="plan">The per-store decision, from <see cref="Assess(LiveTestSettings)"/>.</param>
    /// <param name="countTagged">Counts tagged items in one store. Read-only.</param>
    /// <param name="purgeUntilStableZero">Deletes tagged items from one store until the count holds at zero.</param>
    /// <param name="report">Where the per-store lines go - normally <c>ITestOutputHelper.WriteLine</c>.</param>
    /// <returns>The plan that was walked, so a caller may assert on its shape.</returns>
    public static ArtifactSweepPlan Run(
        ArtifactSweepPlan plan,
        Func<string, int> countTagged,
        Action<string> purgeUntilStableZero,
        Action<string> report)
    {
        ArgumentNullException.ThrowIfNull(plan);
        ArgumentNullException.ThrowIfNull(countTagged);
        ArgumentNullException.ThrowIfNull(purgeUntilStableZero);
        ArgumentNullException.ThrowIfNull(report);

        report(plan.Describe());

        // Before the walk, not after it: a sweep over no store at all counts nothing, finds
        // nothing and asserts nothing, and the line it prints is indistinguishable from a clean
        // run. Refusing it here means the emptiness is never reported as a pass.
        string? vacuous = plan.Refusal();
        if (vacuous != null)
        {
            throw new InvalidOperationException(vacuous);
        }

        List<string> refusals = new();
        foreach (ArtifactSweepStep step in plan.Steps)
        {
            int count = countTagged(step.Store);
            if (count > 0 && step.MayDelete)
            {
                report(step.Purging(count));
                purgeUntilStableZero(step.Store);
                count = countTagged(step.Store);
            }

            report(step.Announce(count));
            string? refusal = step.Refusal(count);
            if (refusal != null)
            {
                refusals.Add(refusal);
            }
        }

        if (refusals.Count > 0)
        {
            throw new InvalidOperationException(string.Join(Environment.NewLine, refusals));
        }

        return plan;
    }
}

/// <summary>One store's place in the sweep: whether it is swept or only counted, and why.</summary>
public sealed class ArtifactSweepStep
{
    internal ArtifactSweepStep(string store, ArtifactSweepAction action, bool declaredBystander)
    {
        Store = store;
        Action = action;
        DeclaredBystander = declaredBystander;
    }

    /// <summary>The store's display name, as the settings spell it.</summary>
    public string Store { get; }

    /// <summary>What the sweep may do here.</summary>
    public ArtifactSweepAction Action { get; }

    /// <summary>
    /// True when this store is in <c>bystanderStoreDisplayNames</c> - watched precisely because
    /// nothing writes to it. Carried separately from <see cref="Action"/> because it is the
    /// difference between "somebody declared this store off limits" and "the allowlist happens
    /// not to grant it", and the two mean different things to whoever reads the failure.
    /// </summary>
    public bool DeclaredBystander { get; }

    /// <summary>True when the sweep may delete from this store.</summary>
    public bool MayDelete => Action == ArtifactSweepAction.Sweep;

    /// <summary>The line printed before a purge pass, naming what is about to be deleted.</summary>
    /// <param name="found">The count that triggered the purge.</param>
    public string Purging(int found)
    {
        return "sweep[" + Store + "]: " + found + " late-materialized tagged artifact(s) found - "
            + "purging (documented sent-copy lag)";
    }

    /// <summary>
    /// The per-store line, printed for EVERY store and not only the interesting ones. A
    /// count-only store says so in the line, so a run states what it did not do rather than
    /// leaving a reader to infer from a zero that the store was swept clean.
    /// </summary>
    /// <param name="count">The final count for this store.</param>
    public string Announce(int count)
    {
        string line = "sweep[" + Store + "]: taggedArtifacts=" + count;
        return MayDelete ? line : line + " - COUNTED, NOT SWEPT (" + Because() + ")";
    }

    /// <summary>
    /// The refusal for this store's final count, or null when it is zero. Two different
    /// failures: a swept store still holding artifacts is a cleanup that did not converge; a
    /// counted store holding any at all is something arriving where nothing should be able to
    /// put it, which no other guard would report - the count tripwire fires on a DECREASE, and
    /// this is an increase.
    /// </summary>
    /// <param name="count">The final count for this store.</param>
    public string? Refusal(int count)
    {
        if (count <= 0)
        {
            return null;
        }

        if (MayDelete)
        {
            return "sweep[" + Store + "]: " + count + " tagged artifact(s) still present after a "
                + "stable-zero purge. The live tier must end with none anywhere (mailbox-safety "
                + "rule 4) - check the store by hand before running again.";
        }

        return Residue(count);
    }

    /// <summary>The refusal text for a count the sweep is not allowed to act on.</summary>
    private string Residue(int count)
    {
        return ArtifactSweepPolicy.Residue + ": store '" + Store + "' holds " + count
            + " item(s) carrying the live-tier subject tag, and " + Because() + "." + Environment.NewLine
            + "  NOTHING HERE WILL DELETE THEM, and that is on purpose: a delete aimed at this "
            + "store is the shape that nearly destroyed the measurement corpus, so the sweep "
            + "counts it and stops." + Environment.NewLine
            + "  Nor would any other guard have told you: the per-store count tripwire fires on "
            + "an item-count DECREASE, and this is an increase." + Environment.NewLine
            + "  Either a test wrote where the allowlist says it cannot - in which case find it "
            + "before running the tier again - or a real item's subject contains the live-tier "
            + "tag, in which case it is a person's mail and must be left alone. Inspect it by "
            + "hand. Do not widen the guard and do not sweep this store.";
    }

    /// <summary>Why this store is counted rather than swept, in the words its reader needs.</summary>
    private string Because()
    {
        return DeclaredBystander
            ? "declared BYSTANDER - the count tripwire watches it precisely because nothing "
                + "writes to it, so no test may write to it"
            : "the write allowlist grants no delete on it";
    }
}

/// <summary>
/// One run's sweep, store by store - see <see cref="ArtifactSweepPolicy"/> for why the walk is
/// planned before it is executed.
/// </summary>
public sealed class ArtifactSweepPlan
{
    internal ArtifactSweepPlan(IReadOnlyList<ArtifactSweepStep> steps)
    {
        Steps = steps;
        Swept = steps.Where(s => s.MayDelete).Select(s => s.Store).ToList();
        CountedOnly = steps.Where(s => !s.MayDelete).Select(s => s.Store).ToList();
    }

    /// <summary>Every store to visit, once each, in the order the settings name them.</summary>
    public IReadOnlyList<ArtifactSweepStep> Steps { get; }

    /// <summary>The stores this sweep may delete from.</summary>
    public IReadOnlyList<string> Swept { get; }

    /// <summary>The stores it counts and leaves alone.</summary>
    public IReadOnlyList<string> CountedOnly { get; }

    /// <summary>
    /// True when the sweep would visit nothing - the one shape whose green result means
    /// nothing at all. <b>A refusal</b>, for the same reason the vacuous census is one.
    /// </summary>
    public bool ProvesNothing => Steps.Count == 0;

    /// <summary>The one-line summary printed before the walk.</summary>
    public string Describe()
    {
        string line = "artifact sweep: " + Steps.Count + " store(s) - " + Swept.Count
            + " swept, " + CountedOnly.Count + " counted and left alone";
        if (CountedOnly.Count == 0)
        {
            return line;
        }

        return line + " (" + string.Join(", ", CountedOnly.Select(s => "'" + s + "'")) + ")";
    }

    /// <summary>The refusal for the plan itself, or null when there is something to walk.</summary>
    public string? Refusal()
    {
        if (!ProvesNothing)
        {
            return null;
        }

        return "REFUSING to report a clean sweep: " + ArtifactSweepPolicy.NothingToVisit
            + ". It would count nothing, find nothing and report zero tagged artifacts, which is "
            + "the same output a genuinely clean run produces - so it cannot tell a swept profile "
            + "from an unswept one." + Environment.NewLine
            + "  'expectedStoreDisplayNames' in the live-test settings is what this walks; it is "
            + "empty, or every entry in it is blank.";
    }
}
