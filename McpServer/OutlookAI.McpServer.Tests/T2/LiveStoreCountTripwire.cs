using System.Diagnostics;
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
    /// </summary>
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
        if (verdict.Failed)
        {
            // Confirm before crying: one more census, and only what fails BOTH times is
            // reported. COM enumeration under a busy Outlook is not perfectly repeatable.
            //
            // Failure lines name EntryIDs now, so this intersection is stricter than it was:
            // the SAME items must be missing both times, not merely the same tally. That is
            // the right direction - a line that changes between two censuses seconds apart
            // was describing enumeration noise, not a deletion.
            Console.WriteLine("[tripwire] suspected loss - re-censusing to confirm.");
            CensusPass recheck = Capture(baseline.Keys.ToList(), "confirmation", hub, baseline);
            TripwireVerdict second = StoreCountTripwire.Evaluate(baseline, recheck.Census, hub, lazyStores);
            HashSet<string> secondKeys = new(second.FailureRecords.Select(f => f.Key), StringComparer.Ordinal);
            List<TripwireFailure> confirmed =
                verdict.FailureRecords.Where(f => secondKeys.Contains(f.Key)).ToList();
            if (confirmed.Count == 0)
            {
                Console.WriteLine("[tripwire] not reproducible on the second census - treating as enumeration noise.");
            }

            verdict = new TripwireVerdict(
                confirmed, verdict.Notes, confirmed.Count > 0 ? verdict.Attribution : null);
        }

        foreach (string note in verdict.Notes)
        {
            Console.WriteLine("[tripwire] note:" + note);
        }

        Console.WriteLine(
            $"[tripwire] post-run census in {stopwatch.ElapsedMilliseconds} ms ({after.Describe()}); "
            + $"{verdict.Failures.Count} failure(s), {verdict.Notes.Count} note(s).");

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
            throw new InvalidOperationException(verdict.Describe());
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
            try
            {
                census[store] = LiveOutlookTestMailer.CaptureMailFolderCensus(store, plan);
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                throw new InvalidOperationException(
                    "REFUSING to run the live tier: the " + phase + " per-store census for '" + store
                    + "' could not be taken (" + ex.GetType().Name + ": " + ex.Message
                    + "). An unmeasured mailbox cannot be proven untouched.",
                    ex);
            }

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
