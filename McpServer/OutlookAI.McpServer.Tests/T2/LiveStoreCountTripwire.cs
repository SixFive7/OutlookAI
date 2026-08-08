using System.Diagnostics;
using OutlookAI.Core.Com;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The live-tier half of <see cref="StoreCountTripwire"/>: one per-store, per-folder
/// census before the first live collection runs and one after the last, over ALL
/// configured stores - the primary accounts AND the delegate/shared mailboxes that are
/// read-only for tests.
/// <para>
/// Fail-closed, like <c>SignatureDirectorySnapshot</c>: if the baseline cannot be taken
/// the live tier REFUSES to run, because an unmeasured mailbox cannot be proven
/// untouched. Every live collection fixture calls <see cref="EnsureBaseline"/> in its
/// constructor (a throw there fails the whole collection), and the last-ordered fixture
/// calls <see cref="Verify"/> in Dispose.
/// </para>
/// </summary>
public static class LiveStoreCountTripwire
{
    private static readonly object Gate = new();
    private static Dictionary<string, IReadOnlyDictionary<string, int>>? _baseline;
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
    /// watched store cannot be counted.
    /// </summary>
    public static void EnsureBaseline(LiveTestSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        lock (Gate)
        {
            if (_baseline != null)
            {
                return;
            }

            _hub = settings.TestHubStoreDisplayName;
            _lazyHierarchyStores = settings.ExpectedDelegateStoreDisplayNames.ToList();
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
            _baseline = Capture(WatchedStores(settings), "baseline");
            stopwatch.Stop();
            Console.WriteLine(
                $"[tripwire] baseline: {_baseline.Count} stores, "
                + $"{_baseline.Values.Sum(f => f.Count)} mail folders, {stopwatch.ElapsedMilliseconds} ms.");
        }
    }

    /// <summary>
    /// Re-counts and compares. Throws naming store/folder/delta when anything was lost
    /// outside the hub. Runs once; later calls are no-ops.
    /// </summary>
    public static void Verify()
    {
        Dictionary<string, IReadOnlyDictionary<string, int>> baseline;
        string hub;
        IReadOnlyList<string> lazyStores;
        lock (Gate)
        {
            if (_baseline == null || _hub == null || _verified)
            {
                return;
            }

            _verified = true;
            baseline = _baseline;
            hub = _hub;
            lazyStores = _lazyHierarchyStores;
        }

        Stopwatch stopwatch = Stopwatch.StartNew();
        Dictionary<string, IReadOnlyDictionary<string, int>> after = Capture(baseline.Keys.ToList(), "post-run");
        stopwatch.Stop();

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(baseline, after, hub, lazyStores);
        if (verdict.Failed)
        {
            // Confirm before crying: one more census, and only what fails BOTH times is
            // reported. COM enumeration under a busy Outlook is not perfectly repeatable.
            Console.WriteLine("[tripwire] suspected loss - re-counting to confirm.");
            Dictionary<string, IReadOnlyDictionary<string, int>> recheck =
                Capture(baseline.Keys.ToList(), "confirmation");
            TripwireVerdict second = StoreCountTripwire.Evaluate(baseline, recheck, hub, lazyStores);
            List<string> confirmed = verdict.Failures.Intersect(second.Failures, StringComparer.Ordinal).ToList();
            if (confirmed.Count == 0)
            {
                Console.WriteLine("[tripwire] not reproducible on the second census - treating as enumeration noise.");
            }

            verdict = new TripwireVerdict(confirmed, verdict.Notes);
        }

        foreach (string note in verdict.Notes)
        {
            Console.WriteLine("[tripwire] note:" + note);
        }

        Console.WriteLine(
            $"[tripwire] post-run census in {stopwatch.ElapsedMilliseconds} ms; "
            + $"{verdict.Failures.Count} failure(s), {verdict.Notes.Count} note(s).");

        // Releases COM references only - Outlook keeps running (S7: never kill/close).
        OutlookComSession? keepAlive;
        lock (Gate)
        {
            keepAlive = _keepAlive;
            _keepAlive = null;
        }

        keepAlive?.Dispose();

        if (verdict.Failed)
        {
            throw new InvalidOperationException(verdict.Describe());
        }
    }

    /// <summary>Test hook: forgets the baseline so a self-test can drive the guard.</summary>
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
    }

    private static Dictionary<string, IReadOnlyDictionary<string, int>> Capture(
        IReadOnlyList<string> stores, string phase)
    {
        Dictionary<string, IReadOnlyDictionary<string, int>> census = new(StringComparer.OrdinalIgnoreCase);
        foreach (string store in stores)
        {
            try
            {
                census[store] = LiveOutlookTestMailer.CountMailFolderItems(store);
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                throw new InvalidOperationException(
                    "REFUSING to run the live tier: the " + phase + " per-store count for '" + store
                    + "' could not be taken (" + ex.GetType().Name + ": " + ex.Message
                    + "). An unmeasured mailbox cannot be proven untouched.",
                    ex);
            }
        }

        return census;
    }
}
