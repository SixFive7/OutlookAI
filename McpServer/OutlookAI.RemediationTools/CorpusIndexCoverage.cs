using System.Globalization;

namespace OutlookAI.RemediationTools;

/// <summary>One population item as the Windows Search index returned it.</summary>
/// <param name="Ordinal">The ordinal parsed out of the row's subject with the corpus's own ordinal parse.</param>
/// <param name="DateReceivedUtc">The row's <c>System.Message.DateReceived</c>, as the product maps it; null when the index gave it none.</param>
public sealed record CorpusIndexedRow(int Ordinal, DateTime? DateReceivedUtc);

/// <summary>What the index holds of a population, measured against its plan.</summary>
/// <param name="Planned">Ordinals the population holds.</param>
/// <param name="Indexed">Distinct planned ordinals the index returned at least one message row for.</param>
/// <param name="Missing">Planned ordinals the index returned no row for, in order.</param>
/// <param name="UndatedPlanned">Planned UNDATED ordinals.</param>
/// <param name="UndatedIndexedWithADate">Undated ordinals the index nevertheless gave a received date.</param>
/// <param name="DatedCompared">Dated ordinals whose index date could be compared with what the store holds.</param>
/// <param name="DatedMismatched">Of those, how many differ by more than <see cref="CorpusIndexCoverage.DateTolerance"/>.</param>
/// <param name="ModalMismatchSeconds">The most common difference among the mismatches, index minus store, in seconds; null when none.</param>
/// <param name="NewestIndexedUtc">The newest received date the index holds for the population.</param>
/// <param name="NewestPlannedUtc">The newest received date the population is supposed to hold.</param>
public sealed record CorpusIndexCoverageReport(
    int Planned,
    int Indexed,
    IReadOnlyList<int> Missing,
    int UndatedPlanned,
    int UndatedIndexedWithADate,
    int DatedCompared,
    int DatedMismatched,
    long? ModalMismatchSeconds,
    DateTime? NewestIndexedUtc,
    DateTime? NewestPlannedUtc);

/// <summary>
/// Whether the Windows Search index has taken a freshly built population in - the question the
/// per-run hub rebuild has to answer before a live run may start, because every index test reads
/// the hub and a hub the indexer has not reached is one they would measure half of.
/// <para>
/// <b>Coverage decides; the dates are REPORTED.</b> The build is complete when every planned
/// ordinal has a message row. The received dates the index returns are compared with what the store
/// holds - the manifest's read-back where there is one, the plan otherwise - and a systematic
/// difference is printed with its size, because a difference equal to the machine's UTC offset is
/// exactly the local-time misreading <c>LiveIndexSearchTests.Staleness_SelfReportsPlausibleFrontier</c>
/// exists to catch. It is that test's to fail, not this check's: this one only says when the index
/// is ready to be measured.
/// </para>
/// <para>Pure: rows in, a report out. The index query that produces the rows is the command's.</para>
/// </summary>
public static class CorpusIndexCoverage
{
    /// <summary>How far an index date may sit from the store's before it counts as different.</summary>
    public static readonly TimeSpan DateTolerance = TimeSpan.FromSeconds(2);

    /// <summary>Compares what the index returned with the population's plan.</summary>
    /// <param name="plan">A population plan.</param>
    /// <param name="itemCount">The population's item count.</param>
    /// <param name="rows">Every message row the index returned whose subject parsed as this population's.</param>
    /// <param name="manifest">The build's manifest, whose read-back dates are what the store holds; optional.</param>
    public static CorpusIndexCoverageReport Compare(
        CorpusPlan plan, int itemCount, IEnumerable<CorpusIndexedRow> rows, CorpusManifest? manifest)
    {
        ArgumentNullException.ThrowIfNull(plan);
        ArgumentNullException.ThrowIfNull(rows);
        if (plan.Population == null)
        {
            throw new ArgumentException("Index coverage is measured for a population, whose size is fixed.", nameof(plan));
        }

        var byOrdinal = new Dictionary<int, List<CorpusIndexedRow>>();
        foreach (CorpusIndexedRow row in rows)
        {
            if (row.Ordinal < 1 || row.Ordinal > itemCount)
            {
                continue;
            }

            if (!byOrdinal.TryGetValue(row.Ordinal, out List<CorpusIndexedRow>? list))
            {
                byOrdinal[row.Ordinal] = list = new List<CorpusIndexedRow>();
            }

            list.Add(row);
        }

        var missing = new List<int>();
        int undatedPlanned = 0;
        int undatedDated = 0;
        int compared = 0;
        int mismatched = 0;
        var offsets = new Dictionary<long, int>();
        DateTime? newestIndexed = null;
        DateTime? newestPlanned = null;
        for (int ordinal = 1; ordinal <= itemCount; ordinal++)
        {
            CorpusItemSpec spec = plan.Describe(ordinal);
            if (spec.IsUndated)
            {
                undatedPlanned++;
            }
            else if (newestPlanned == null || spec.ReceivedUtc > newestPlanned)
            {
                newestPlanned = spec.ReceivedUtc;
            }

            if (!byOrdinal.TryGetValue(ordinal, out List<CorpusIndexedRow>? seen))
            {
                missing.Add(ordinal);
                continue;
            }

            if (spec.IsUndated)
            {
                if (seen.Any(r => r.DateReceivedUtc != null))
                {
                    undatedDated++;
                }

                continue;
            }

            DateTime? indexed = seen.Select(r => r.DateReceivedUtc).FirstOrDefault(d => d != null);
            if (indexed == null)
            {
                continue;
            }

            if (newestIndexed == null || indexed > newestIndexed)
            {
                newestIndexed = indexed;
            }

            DateTime truth = spec.ReceivedUtc;
            if (manifest != null && manifest.Items.TryGetValue(ordinal, out CorpusManifestItem? recorded)
                && CorpusManifest.ParseUtc(recorded.ReceivedUtc) is DateTime readBack)
            {
                truth = readBack;
            }

            compared++;
            long offset = (long)Math.Round((indexed.Value - truth).TotalSeconds);
            if (Math.Abs(offset) > DateTolerance.TotalSeconds)
            {
                mismatched++;
                offsets[offset] = offsets.TryGetValue(offset, out int n) ? n + 1 : 1;
            }
        }

        long? modal = offsets.Count == 0
            ? null
            : offsets.OrderByDescending(kv => kv.Value).ThenBy(kv => Math.Abs(kv.Key)).First().Key;
        return new CorpusIndexCoverageReport(
            itemCount, itemCount - missing.Count, missing, undatedPlanned, undatedDated,
            compared, mismatched, modal, newestIndexed, newestPlanned);
    }

    /// <summary>Whether the index holds the whole population, and what to print either way.</summary>
    public static (bool Complete, string Message) Decide(CorpusIndexCoverageReport report)
    {
        ArgumentNullException.ThrowIfNull(report);
        CultureInfo invariant = CultureInfo.InvariantCulture;
        string head = $"Index coverage: {report.Indexed.ToString(invariant)} of {report.Planned.ToString(invariant)} "
            + $"population item(s) indexed ({report.UndatedPlanned.ToString(invariant)} of them undated)";
        string newest = report.NewestPlannedUtc == null
            ? string.Empty
            : $"; newest planned {CorpusManifest.FormatUtc(report.NewestPlannedUtc.Value)}, newest indexed "
                + (report.NewestIndexedUtc == null ? "(none)" : CorpusManifest.FormatUtc(report.NewestIndexedUtc.Value));

        var notes = new List<string>();
        if (report.DatedMismatched > 0)
        {
            notes.Add($"NOTE: {report.DatedMismatched.ToString(invariant)} of {report.DatedCompared.ToString(invariant)} dated "
                + $"item(s) carry a different received date in the index than in the store, most often by "
                + $"{report.ModalMismatchSeconds!.Value.ToString(invariant)} s - if that is this machine's UTC offset, the "
                + "index reports local time as if it were UTC, which is what the staleness test is for");
        }

        if (report.UndatedIndexedWithADate > 0)
        {
            notes.Add($"NOTE: the index gave {report.UndatedIndexedWithADate.ToString(invariant)} UNDATED item(s) a received "
                + "date anyway, so LiveOrderKeyCollationTests will find fewer undated rows than the plan holds");
        }

        string tail = notes.Count == 0 ? "." : ". " + string.Join(". ", notes) + ".";
        if (report.Missing.Count == 0)
        {
            return (true, head + newest + tail);
        }

        string sample = string.Join(", ", report.Missing.Take(12).Select(o => o.ToString(invariant)))
            + (report.Missing.Count > 12 ? ", ..." : string.Empty);
        return (false, head + newest + $". NOT YET: {report.Missing.Count.ToString(invariant)} ordinal(s) have no row in "
            + $"the index (first: {sample}). The indexer only advances while Outlook runs; a run started now would "
            + "measure part of this population" + tail);
    }
}
