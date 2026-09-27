using System.Globalization;
using OutlookAI.RemediationTools;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>What a hub population's manifest says about the population, as far as its age goes.</summary>
/// <param name="CorpusId">The population's corpus id.</param>
/// <param name="Store">The store it was built into - the hub.</param>
/// <param name="AnchorUtc">The anchor it was last built against.</param>
/// <param name="NewestDatedUtc">Its newest DATED item's received instant - one minute before the anchor, by the plan.</param>
public sealed record HubPopulationFact(string CorpusId, string Store, DateTime AnchorUtc, DateTime NewestDatedUtc);

/// <summary>
/// Whether the hub's generated population is young enough for
/// <c>LiveIndexSearchTests.Staleness_SelfReportsPlausibleFrontier</c> to mean anything - the fail-closed
/// half of the per-run hub rebuild the maintainer decided on 2026-09-24 (option (a), "not something a
/// human must remember").
/// <para>
/// <b>Why age decides it.</b> The frontier test asserts the index frontier is not in the future. A
/// product that read the index's local time as UTC would shift the frontier forward by the machine's
/// UTC offset - into the future, and so caught, ONLY while the true frontier is younger than that
/// offset less the test's own five minutes of tolerance. On a W. Europe guest that is 55 minutes in
/// winter. A hub population built weeks ago puts the frontier weeks old, and the misreading passes.
/// The rebuild script (<c>Testbed/guest/Reset-HubPopulation.ps1</c>) makes the hub fresh; this is what
/// turns forgetting it into a red test that names the script, the same way the corpus builder refuses
/// while <c>ImportPRF</c> is set rather than trusting anyone to clear it.
/// </para>
/// <para>Pure: manifest lines and a clock in, a verdict out. T1 pins every branch.</para>
/// </summary>
public static class LiveHubPopulationFreshness
{
    /// <summary>
    /// How far in the future the frontier test lets the frontier sit before it calls it wrong - its own
    /// <c>ClockUtc.AddMinutes(5)</c>, restated here because the margin below is derived from it.
    /// </summary>
    public static readonly TimeSpan FrontierFutureTolerance = TimeSpan.FromMinutes(5);

    /// <summary>
    /// Reads a hub population's manifest header and refuses anything else: another population's, the
    /// measurement corpus's, or one whose shape this generator would not reproduce - an old format, or
    /// a store name that does not match. Returns what its age is judged on.
    /// </summary>
    public static HubPopulationFact Read(IEnumerable<string> manifestLines)
    {
        ArgumentNullException.ThrowIfNull(manifestLines);
        CorpusManifest manifest = CorpusManifest.Parse(manifestLines);
        CorpusManifestHeader header = manifest.Header;
        DateTime anchor = CorpusManifest.ParseUtc(header.AnchorUtc)
            ?? throw new InvalidOperationException(
                $"The hub population manifest records an anchor that is not a UTC instant ('{header.AnchorUtc}').");

        var options = new CorpusPlanOptions(header.CorpusId, header.Seed, anchor)
        {
            Population = CorpusPopulationKind.Hub,
            Owner = CorpusMailboxOwner.ForStore(header.StoreDisplayName),
        };
        if (!string.Equals(options.ShapeKey, header.ShapeKey, StringComparison.Ordinal))
        {
            throw new InvalidOperationException(
                "The manifest named by 'hubPopulationManifestPath' is not a hub population this build of the generator "
                + $"would reproduce (its shape key is '{header.ShapeKey}', a v{CorpusPopulation.Version.ToString(CultureInfo.InvariantCulture)} "
                + "hub population's would end '|p:hub:v...'). Point the setting at the hub's own corpus-<id>.jsonl, and rebuild "
                + "the hub with Testbed/guest/Reset-HubPopulation.ps1 if it was built by an older generator.");
        }

        var plan = new CorpusPlan(options);
        DateTime newest = plan.Report(1, plan.FixedItemCount!.Value).NewestReceivedUtc;
        return new HubPopulationFact(header.CorpusId, header.StoreDisplayName, anchor, newest);
    }

    /// <summary>
    /// The oldest the hub's newest item may be for the frontier test to catch a local-time misreading
    /// on a machine at <paramref name="utcOffset"/>: the offset's size less the test's tolerance. Zero
    /// or less means that machine cannot tell local time from UTC at all.
    /// </summary>
    public static TimeSpan DiscriminatingAge(TimeSpan utcOffset) => utcOffset.Duration() - FrontierFutureTolerance;

    /// <summary>Whether the frontier test may run against this hub population now, and what to print either way.</summary>
    public static (bool Proceed, string Message) Decide(HubPopulationFact fact, DateTime nowUtc, TimeSpan utcOffset)
    {
        ArgumentNullException.ThrowIfNull(fact);
        CultureInfo invariant = CultureInfo.InvariantCulture;
        TimeSpan margin = DiscriminatingAge(utcOffset);
        TimeSpan age = DateTime.SpecifyKind(nowUtc, DateTimeKind.Utc) - fact.NewestDatedUtc;
        string what = $"hub population '{fact.CorpusId}' in '{fact.Store}', anchored {CorpusManifest.FormatUtc(fact.AnchorUtc)}; "
            + $"its newest item is {Minutes(age)} old, and this machine's UTC offset of {Minutes(utcOffset)} lets the "
            + $"frontier test catch a local-time misreading while that is under {Minutes(margin)}";

        if (margin <= TimeSpan.Zero)
        {
            return (false, "REFUSING the frontier check: " + what + ". A machine on UTC cannot tell a local-time frontier "
                + "from a UTC one at all, so this test proves nothing here whatever the hub holds. Run it on a guest whose "
                + "time zone is not UTC (the testbed guests are W. Europe).");
        }

        if (age < -FrontierFutureTolerance)
        {
            return (false, "REFUSING the frontier check: " + what + ". The population's newest item is in the FUTURE, which "
                + "means it was built against an anchor ahead of this clock - check the guest's clock and rebuild the hub.");
        }

        if (age > margin)
        {
            return (false, "STALE HUB: " + what + ". A frontier that old passes this test whether the product reads the index's "
                + "time as UTC or as local time, so the run is not measuring what the test is for. Rebuild the hub before "
                + "the run: Testbed/guest/Reset-HubPopulation.ps1 -Execute, through Register-InteractiveTask.ps1 "
                + "(Docs/live-tier-on-the-vm.md section 3b).");
        }

        return (true, "Hub population fresh: " + what + ".");
    }

    private static string Minutes(TimeSpan span)
        => (span < TimeSpan.Zero ? "-" : string.Empty)
            + ((long)Math.Round(span.Duration().TotalMinutes)).ToString(CultureInfo.InvariantCulture) + " min";
}
