using System.Diagnostics;
using System.Globalization;
using OutlookAI.Core.Com;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The live-tier half of <see cref="StoreCountTripwire"/>: one per-store, per-folder
/// census before the first live collection runs and one after the last, over ALL
/// configured stores - the primary accounts AND the delegate/shared mailboxes that are
/// read-only for tests.
/// <para>
/// Every folder is counted; folders within <see cref="CensusIdentityPlan"/>'s budget are
/// also walked item by item, so a firing can say WHICH items left rather than only how
/// many. The post-run passes repeat the baseline's identity choices rather than re-deciding
/// them, because a folder walked at one end and counted at the other cannot be compared
/// item by item at all.
/// </para>
/// <para>
/// Fail-closed, like <c>SignatureDirectorySnapshot</c>: if the baseline cannot be taken
/// the live tier REFUSES to run, because an unmeasured mailbox cannot be proven
/// untouched. Every live collection fixture calls <see cref="EnsureBaseline"/> in its
/// constructor (a throw there fails the whole collection) and
/// <see cref="CollectionFinished"/> in its Dispose; the comparison happens when the last
/// guarded collection OF THIS RUN finishes, which <see cref="LiveTierRunPlan"/> works out
/// from the filtered collection list rather than assuming a whole-tier run.
/// </para>
/// <para>
/// Being that single funnel, it is also where <see cref="LiveOutlookPreflight"/> gates the
/// tier on Outlook actually responding - and it needs the gate for its own sake, since the
/// <c>OutlookComSession.Connect</c> below is the call that hung for 10 and then 15 minutes
/// on 2026-08-18.
/// </para>
/// </summary>
public static class LiveStoreCountTripwire
{
    private static readonly object Gate = new();
    private static Dictionary<string, IReadOnlyDictionary<string, FolderCensus>>? _baseline;
    private static string? _hub;
    private static IReadOnlyList<string> _lazyHierarchyStores = Array.Empty<string>();
    private static bool _verified;

    /// <summary>
    /// One COM session held for the whole live tier. The census itself opens and releases
    /// short-lived Outlook references; if the tests STARTED Outlook, releasing the last one
    /// arms its idle self-exit (~11.5 min - the measured headless lifetime), which then
    /// fires in the middle of the run and turns every later COM call into "RPC server is
    /// unavailable". Holding one reference keeps the instance alive exactly like a real
    /// agent session does. Released in <see cref="Verify"/> - never quits Outlook (S7).
    /// </summary>
    private static OutlookComSession? _keepAlive;

    /// <summary>True once a baseline exists for this process.</summary>
    public static bool HasBaseline
    {
        get
        {
            lock (Gate)
            {
                return _baseline != null;
            }
        }
    }

    /// <summary>Stores the tripwire watches: every configured primary AND delegate store.</summary>
    public static IReadOnlyList<string> WatchedStores(LiveTestSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        List<string> stores = new(settings.ExpectedStoreDisplayNames);
        foreach (string delegateStore in settings.ExpectedDelegateStoreDisplayNames)
        {
            if (!stores.Any(s => string.Equals(s, delegateStore, StringComparison.OrdinalIgnoreCase)))
            {
                stores.Add(delegateStore);
            }
        }

        return stores;
    }

    /// <summary>
    /// Takes the baseline once per process. Throws (refusing the live tier) when any
    /// watched store cannot be censused.
    /// </summary>
    public static void EnsureBaseline(LiveTestSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        lock (Gate)
        {
            // Health gate first, ahead of the early return and ahead of every COM call.
            // This method is the single funnel all eight live collection fixtures pass
            // through, and the OutlookComSession.Connect below is the exact line that sat
            // for 10 and then 15 minutes against a wedged Outlook on 2026-08-18. Asked per
            // collection rather than once per process because the probe costs microseconds
            // and Outlook can wedge mid-suite as easily as before it.
            LiveOutlookPreflight.Require();

            if (_baseline != null)
            {
                return;
            }

            _hub = settings.TestHubStoreDisplayName;
            _lazyHierarchyStores = settings.ExpectedDelegateStoreDisplayNames.ToList();

            // Printed before the first COM call: a run that turns out to have been pointed
            // at the wrong machine's settings should say so at the top of the log, not be
            // inferred afterwards from which tests behaved oddly.
            Console.WriteLine("[tripwire] live-test settings: " + settings.Describe() + ".");
            try
            {
                _keepAlive = OutlookComSession.Connect(allowStartingOutlook: true);
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                throw new InvalidOperationException(
                    "REFUSING to run the live tier: Outlook could not be reached for the count tripwire ("
                    + ex.GetType().Name + ").", ex);
            }

            Stopwatch stopwatch = Stopwatch.StartNew();
            CensusPass pass = Capture(WatchedStores(settings), "baseline", settings.TestHubStoreDisplayName, null);
            stopwatch.Stop();
            _baseline = pass.Census;
            Console.WriteLine(
                $"[tripwire] baseline: {_baseline.Count} stores, "
                + $"{_baseline.Values.Sum(f => f.Count)} mail folders, {pass.Describe()}, "
                + $"{stopwatch.ElapsedMilliseconds} ms.");
        }
    }

    /// <summary>
    /// Signals that one guarded collection has finished, and verifies the run when nothing
    /// guarded comes after it.
    /// <para>
    /// Every live collection fixture calls this from its dispose, because which collection
    /// ends the run is a property of the FILTER, not of the suite. A run selecting one test
    /// class ends at that class's collection; a whole-tier run ends at
    /// <see cref="LiveCollections.Lifecycle"/>, which the collection orderer forces last.
    /// Before this existed only the second case was verified, so every filtered run - which
    /// is how the tier is used on a test machine - paid for a baseline and threw it away.
    /// </para>
    /// <para>
    /// When the run plan is <see cref="GuardedCollectionPosition.Unknown"/> (nothing
    /// published, so the collection orderer did not run) this verifies and stays ARMED, so
    /// each later collection boundary is checked too. That costs a census per collection and
    /// is the deliberate trade: an unverified run is the only outcome that must not happen.
    /// </para>
    /// </summary>
    public static void CollectionFinished(string collectionName)
    {
        ArgumentNullException.ThrowIfNull(collectionName);
        GuardedCollectionPosition position = LiveTierRunPlan.PositionOf(collectionName);
        if (position == GuardedCollectionPosition.NotLast)
        {
            return;
        }

        Verify(final: position == GuardedCollectionPosition.Last);
    }

    /// <summary>
    /// Re-censuses and compares. Throws naming the store, the folder, the items and where
    /// they went when anything was removed outside the hub, and ends with an attribution
    /// line saying how far the evidence actually goes. Runs once; later calls are no-ops.
    /// <para>
    /// A suspected loss is not reported straight away: it goes through
    /// <see cref="TripwireRetryLadder"/>, which is bounded (2 re-censuses, then 1 re-run) and
    /// which reports whatever it did in both directions - a run that passed on the second
    /// census says so, in the same summary line a clean run uses.
    /// </para>
    public static void Verify()
    {
        Verify(final: true);
    }

    /// <summary>
    /// The comparison, with <paramref name="final"/> saying whether this is the last word on
    /// the run. A final verification latches (later calls are no-ops) and releases the
    /// keep-alive COM reference; a non-final one does neither, so the baseline survives to be
    /// compared again at the next collection boundary.
    /// </summary>
    private static void Verify(bool final)
    {
        Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> baseline;
        string hub;
        IReadOnlyList<string> lazyStores;
        lock (Gate)
        {
            if (_baseline == null || _hub == null || _verified)
            {
                return;
            }

            _verified = final;
            baseline = _baseline;
            hub = _hub;
            lazyStores = _lazyHierarchyStores;
        }

        Stopwatch stopwatch = Stopwatch.StartNew();
        CensusPass after = Capture(baseline.Keys.ToList(), "post-run", hub, baseline);
        stopwatch.Stop();

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(baseline, after.Census, hub, lazyStores);
        TripwireRetryReport retry = TripwireRetryReport.NotNeeded();
        if (verdict.Failed)
        {
            // Confirm before crying, but BOUNDED and on the record: every retry is a chance
            // to convert a real loss into a pass, so the ladder stops the moment two censuses
            // agree, and whatever it did is printed whether it ends in a pass or a failure.
            // TripwireRetryLadder owns the policy; this only supplies the censuses.
            Console.WriteLine(
                "[tripwire] suspected loss - re-censusing to confirm (at most "
                + TripwireRetryLadder.MaxReCensuses + ", ~" + TripwireRetryLadder.ReCensusGapSeconds
                + " s apart; then at most " + TripwireRetryLadder.MaxImplicatedReRuns
                + " bounded re-run of the implicated tests).");
            retry = TripwireRetryLadder.Resolve(verdict, new LiveRetrySource(baseline, hub, lazyStores));
            verdict = new TripwireVerdict(
                retry.Confirmed, verdict.Notes, retry.Failed ? verdict.Attribution : null);
            Console.WriteLine(retry.Describe());
        }

        foreach (string note in verdict.Notes)
        {
            Console.WriteLine("[tripwire] note:" + note);
        }

        Console.WriteLine(
            $"[tripwire] post-run census in {stopwatch.ElapsedMilliseconds} ms ({after.Describe()}); "
            + $"{verdict.Failures.Count} failure(s), {verdict.Notes.Count} note(s); {retry.Summary}.");

        // Releases COM references only - Outlook keeps running (S7: never kill/close). Held
        // on a non-final pass: releasing the last reference to an Outlook the tests started
        // arms its idle self-exit, and the run is not over yet.
        if (final)
        {
            OutlookComSession? keepAlive;
            lock (Gate)
            {
                keepAlive = _keepAlive;
                _keepAlive = null;
            }

            keepAlive?.Dispose();
        }

        if (verdict.Failed)
        {
            throw new InvalidOperationException(
                verdict.Describe() + Environment.NewLine + retry.Describe());
        }
    }

    /// <summary>
    /// The live half of <see cref="TripwireRetryLadder"/>: real censuses, a real wait, and an
    /// honest refusal to pretend it can drive the re-run rung.
    /// </summary>
    private sealed class LiveRetrySource : ITripwireRetrySource
    {
        private readonly Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> _baseline;
        private readonly string _hub;
        private readonly IReadOnlyList<string> _lazyStores;

        internal LiveRetrySource(
            Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> baseline,
            string hub,
            IReadOnlyList<string> lazyStores)
        {
            _baseline = baseline;
            _hub = hub;
            _lazyStores = lazyStores;
        }

        /// <summary>
        /// Blocks the teardown for the gap. Deliberate: the alternative is comparing two
        /// censuses taken in the same instant, which would confirm every transient reading
        /// it was supposed to filter out.
        /// </summary>
        public void Wait(TimeSpan gap)
        {
            Console.WriteLine(
                "[tripwire] waiting " + gap.TotalSeconds.ToString("0", CultureInfo.InvariantCulture)
                + " s before the next census.");
            Thread.Sleep(gap);
        }

        /// <summary>
        /// One more census against the SAME baseline. Throws exactly as the first pass does
        /// when a store cannot be censused, which refuses the run rather than clearing it.
        /// </summary>
        public TripwireVerdict ReCensus(int attempt)
        {
            CensusPass again = Capture(
                _baseline.Keys.ToList(),
                "re-census " + attempt.ToString(CultureInfo.InvariantCulture),
                _hub,
                _baseline);
            return StoreCountTripwire.Evaluate(_baseline, again.Census, _hub, _lazyStores);
        }

        /// <summary>
        /// The guarded collections this run actually executed, which is as far as the evidence
        /// goes. A before/after census cannot name an actor at all (that is why the attribution
        /// line says so), and the write allowlist confines the suite to the hub, so a failure
        /// OUTSIDE the hub points at no collection in particular - only at this run. On a
        /// filtered run that is already a short list, which is the whole reason it is worth
        /// naming rather than saying "the tier".
        /// </summary>
        public IReadOnlyList<string> ImplicatedBy(IReadOnlyList<TripwireFailure> persisting)
        {
            IReadOnlyList<string>? ordered = LiveTierRunPlan.Current;
            if (ordered == null)
            {
                return LiveCollections.All;
            }

            List<string> guarded = ordered.Where(LiveCollections.IsGuarded).ToList();
            return guarded.Count > 0 ? guarded : LiveCollections.All;
        }

        /// <summary>
        /// NOT ATTEMPTED, and that is a deliberate refusal rather than an omission. A re-run
        /// means starting a second xunit run of this assembly from inside the first one's
        /// teardown: it would re-enter the fixtures that are currently disposing, re-take a
        /// baseline over a profile mid-teardown, and write to a mailbox at a moment when
        /// nothing is left to sweep the artifacts away. So the rung is reported as
        /// <see cref="TripwireReRunOutcome.Inconclusive"/> - which fails the run - and the
        /// command the maintainer should run is printed instead. An unperformed experiment
        /// exonerates nothing, so this can only ever make the tier fail, never pass.
        /// </summary>
        public TripwireReRunOutcome ReRun(IReadOnlyList<string> implicated, int attempt)
        {
            Console.WriteLine(
                "[tripwire] the bounded re-run cannot be driven from inside the suite's own teardown - it would "
                + "re-enter the fixtures that are disposing right now. Run it by hand, once, over the "
                + implicated.Count.ToString(CultureInfo.InvariantCulture) + " collection(s) that ran: "
                + string.Join(", ", implicated) + ". If the same failure(s) come back, the suite is removing "
                + "mail; if they do not, the delta was ambient.");
            return TripwireReRunOutcome.Inconclusive;
        }
    }

    /// <summary>
    /// Test hook: forgets the baseline AND the run plan so a self-test can drive the guard.
    /// The two are reset together because a plan left over from one case would decide
    /// whether the next one verifies.
    /// </summary>
    internal static void ResetForTests()
    {
        lock (Gate)
        {
            _keepAlive?.Dispose();
            _keepAlive = null;
            _baseline = null;
            _hub = null;
            _verified = false;
        }

        LiveTierRunPlan.ResetForTests();
    }

    /// <summary>
    /// One census over every watched store, plus how much of it was walked item by item.
    /// <para>
    /// <paramref name="repeatOf"/> is null for the baseline and the baseline census for every
    /// later pass: a folder identified at one end and only counted at the other cannot be
    /// compared item by item, so later passes repeat the baseline's choices instead of
    /// re-deciding them against a budget that has since moved.
    /// </para>
    /// </summary>
    private static CensusPass Capture(
        IReadOnlyList<string> stores,
        string phase,
        string hubStoreDisplayName,
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, FolderCensus>>? repeatOf)
    {
        Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> census =
            new(StringComparer.OrdinalIgnoreCase);
        int folders = 0;
        int items = 0;
        foreach (string store in stores)
        {
            CensusIdentityPlan plan = PlanFor(store, hubStoreDisplayName, repeatOf);
            Stopwatch storeClock = Stopwatch.StartNew();
            try
            {
                census[store] = LiveOutlookTestMailer.CaptureMailFolderCensus(store, plan);
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                // The plan doubles as the census's progress record, and it is the only thing
                // still readable when the STA call TIMED OUT rather than failed - the census
                // thread may even still be running. Without it a timeout says which store
                // was too slow and nothing about WHY, which is exactly the position the
                // 2026-08-20 refusal left the maintainer in: no way to tell a slow folder
                // tree from a slow item walk. A count here may be a moment stale; it is a
                // diagnostic, and nothing decides anything from it.
                throw new InvalidOperationException(
                    "REFUSING to run the live tier: the " + phase + " per-store census for '" + store
                    + "' could not be taken (" + ex.GetType().Name + ": " + ex.Message
                    + ") after " + storeClock.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture)
                    + " ms, having reached " + plan.Describe()
                    + ". An unmeasured mailbox cannot be proven untouched.",
                    ex);
            }

            storeClock.Stop();

            // Per store, not just per pass: a profile where ONE mailbox costs minutes and
            // the other four cost milliseconds is invisible in a single total, and that is
            // the shape this census actually has on an Exchange profile with delegate
            // mailboxes that may not be cached locally.
            Console.WriteLine(
                "[tripwire] " + phase + " census of '" + store + "': "
                + census[store].Count.ToString(CultureInfo.InvariantCulture) + " mail folder(s), "
                + plan.Describe() + ", "
                + storeClock.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + " ms.");

            folders += plan.IdentifiedFolders;
            items += plan.IdentifiedItems;
        }

        return new CensusPass(census, folders, items);
    }

    /// <summary>
    /// How much of one store to walk. The hub is counted only - this guard exempts it
    /// anyway, its churn is tagged, and it is the busiest store in the run.
    /// </summary>
    private static CensusIdentityPlan PlanFor(
        string store,
        string hubStoreDisplayName,
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, FolderCensus>>? repeatOf)
    {
        if (string.Equals(store, hubStoreDisplayName, StringComparison.OrdinalIgnoreCase))
        {
            return CensusIdentityPlan.CountOnly();
        }

        if (repeatOf == null)
        {
            return CensusIdentityPlan.Baseline();
        }

        return repeatOf.TryGetValue(store, out IReadOnlyDictionary<string, FolderCensus>? was)
            ? CensusIdentityPlan.Repeating(was)
            : CensusIdentityPlan.CountOnly();
    }

    /// <summary>One census and what identifying it cost, so a run reports its own overhead.</summary>
    private sealed class CensusPass
    {
        internal CensusPass(
            Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> census, int folders, int items)
        {
            Census = census;
            IdentifiedFolders = folders;
            IdentifiedItems = items;
        }

        internal Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> Census { get; }

        internal int IdentifiedFolders { get; }

        internal int IdentifiedItems { get; }

        internal string Describe()
        {
            return $"identified {IdentifiedFolders} folder(s)/{IdentifiedItems} item(s)";
        }
    }
}
