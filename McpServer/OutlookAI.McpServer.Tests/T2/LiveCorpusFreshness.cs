using System.Globalization;
using OutlookAI.RemediationTools;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Coordinates of the synthetic measurement corpus this machine holds, so the live tier can
/// prove it is still usable before it measures anything against it. Every field is needed to
/// make the check at all: without the seed and the anchor there is no plan to compare the
/// store against, and without the manifest there is nothing recording what the store was
/// last set to.
/// </summary>
public sealed class CorpusSettings
{
    /// <summary>Display name of the store the corpus lives in.</summary>
    public string StoreDisplayName { get; set; } = string.Empty;

    /// <summary>Path to the corpus manifest, as handed to <c>corpus-build --manifest</c>.</summary>
    public string ManifestPath { get; set; } = string.Empty;

    /// <summary>The corpus id embedded in every subject.</summary>
    public string CorpusId { get; set; } = string.Empty;

    /// <summary>The generator seed.</summary>
    public long Seed { get; set; }

    /// <summary>The anchor the corpus was generated against, ISO-8601 UTC. Never the clock.</summary>
    public string AnchorUtc { get; set; } = string.Empty;

    /// <summary>How many ordinals the corpus holds.</summary>
    public int ItemCount { get; set; }

    /// <summary>
    /// The measurement windows this machine's tests actually ask about, in days. Empty means
    /// the plan's own marks. A machine that never asks about a one-day window should not be
    /// stopped by a one-day window having emptied, and saying so here is how it declares that.
    /// </summary>
    public List<int> WindowDays { get; set; } = new();

    /// <summary>Whether every field needed to make the check is present.</summary>
    public bool IsComplete
        => !string.IsNullOrWhiteSpace(StoreDisplayName)
            && !string.IsNullOrWhiteSpace(ManifestPath)
            && !string.IsNullOrWhiteSpace(CorpusId)
            && !string.IsNullOrWhiteSpace(AnchorUtc)
            && ItemCount > 0;
}

/// <summary>
/// Refuses to run the live tier against a corpus that has aged out of its own measurement
/// windows.
/// <para>
/// <b>The failure this removes.</b> The corpus is generated against a FIXED anchor. Every
/// test asking about "the last N days" selects against the CLOCK. Six weeks after generation
/// a seven-day window selects nothing at all - and every one of those tests still PASSES,
/// because selecting nothing is a valid answer about an empty window. Nothing goes red.
/// Nothing is logged. The suite simply stops measuring and keeps saying it measured.
/// </para>
/// <para>
/// <b>Why it is a fixture-time refusal and not a per-test assertion.</b> The staleness is a
/// property of the machine, not of any one test, and by the time a test is running it is
/// already too late to tell "this window is empty" from "this window is empty on purpose".
/// So it sits beside the count tripwire, is checked once, and fails the whole tier - which is
/// what the remedy needs anyway, because re-anchoring is an operator action against the store
/// and not something a test may do to itself.
/// </para>
/// <para>
/// It reads the MANIFEST, never the mailbox: no COM, no Outlook, no store. That is deliberate
/// - the check has to be able to run before anything has started Outlook, and a check that
/// needed the thing it is guarding could not be the first thing to run.
/// </para>
/// </summary>
public static class LiveCorpusFreshness
{
    private static readonly object Gate = new();
    private static bool _checked;

    /// <summary>
    /// Proves the corpus can still answer its windows, or throws naming the repair. A no-op
    /// on a machine whose settings declare no corpus, and computed once per process.
    /// </summary>
    public static void EnsureFresh(LiveTestSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        lock (Gate)
        {
            if (_checked)
            {
                return;
            }

            CorpusSettings? corpus = settings.Corpus;
            if (corpus == null)
            {
                _checked = true;
                return;
            }

            CorpusFreshnessReport report = Evaluate(corpus, DateTime.UtcNow);
            (bool proceed, string message) = CorpusFreshness.Decide(report);
            Console.WriteLine("[corpus] " + message);
            if (!proceed)
            {
                throw new InvalidOperationException(
                    "The live tier refuses to run against this corpus. " + message
                    + " Nothing was read from the mailbox to reach this conclusion - it is derived from the "
                    + "manifest at '" + corpus.ManifestPath + "' and the plan the corpus was generated from.");
            }

            _checked = true;
        }
    }

    /// <summary>
    /// The check itself, taking its clock as an argument so it can be exercised at any
    /// instant rather than only at the one the test host happens to run at.
    /// </summary>
    public static CorpusFreshnessReport Evaluate(CorpusSettings corpus, DateTime asOfUtc)
    {
        ArgumentNullException.ThrowIfNull(corpus);
        if (!corpus.IsComplete)
        {
            throw new InvalidOperationException("The corpus settings block is incomplete.");
        }

        if (!File.Exists(corpus.ManifestPath))
        {
            throw new InvalidOperationException(
                $"The corpus manifest '{corpus.ManifestPath}' does not exist, so nothing can say whether the "
                + "corpus is still measurable - and without it the corpus could not be torn down either. It is "
                + "the one file that must survive a VM rebuild.");
        }

        DateTime anchor = CorpusManifest.ParseUtc(corpus.AnchorUtc)
            ?? throw new InvalidOperationException(
                $"The corpus anchor '{corpus.AnchorUtc}' is not an instant. Use yyyy-MM-dd or yyyy-MM-ddTHH:mm:ssZ.");

        var plan = new CorpusPlan(new CorpusPlanOptions(corpus.CorpusId, corpus.Seed, anchor));
        CorpusManifest manifest = CorpusManifest.Parse(File.ReadLines(corpus.ManifestPath));
        (TimeSpan applied, bool provable) =
            CorpusReanchor.DeriveAppliedShift(plan, manifest, out int agreeing, out int dated);
        Console.WriteLine(string.Format(
            CultureInfo.InvariantCulture,
            "[corpus] manifest '{0}': {1:N0} item(s), {2:N0} dated, {3:N0} agreeing on the applied shift.",
            corpus.ManifestPath,
            manifest.Items.Count,
            dated,
            agreeing));

        return CorpusFreshness.Evaluate(
            plan,
            corpus.ItemCount,
            applied,
            asOfUtc,
            corpus.WindowDays.Count > 0 ? corpus.WindowDays : null,
            provable);
    }
}
