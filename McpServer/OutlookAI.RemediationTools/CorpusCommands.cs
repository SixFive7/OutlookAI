using System.Globalization;

namespace OutlookAI.RemediationTools;

/// <summary>
/// Parsed command line for the corpus commands. Its own parser rather than the console's
/// shared one because <c>--allow-store</c> is REPEATABLE: the allowlist is the guard that
/// decides whether tens of thousands of items may be written into a mailbox, and squeezing
/// several store names into one comma-separated value would put a delimiter inside a name
/// that users are free to put commas in.
/// </summary>
public sealed class CorpusOptions
{
    private CorpusOptions()
    {
    }

    /// <summary>Target store display name.</summary>
    public string? Store { get; private set; }

    /// <summary>Stores the caller explicitly permits writing to.</summary>
    public List<string> AllowStores { get; } = new();

    /// <summary>Corpus id embedded in every subject.</summary>
    public string? CorpusId { get; private set; }

    /// <summary>Generator seed.</summary>
    public long Seed { get; private set; }

    /// <summary>Whether a seed was given at all.</summary>
    public bool HasSeed { get; private set; }

    /// <summary>The instant ages are measured back from.</summary>
    public DateTime? AnchorUtc { get; private set; }

    /// <summary>How many items the corpus should hold in total.</summary>
    public int Count { get; private set; }

    /// <summary>Manifest path.</summary>
    public string? ManifestPath { get; private set; }

    /// <summary>Report progress every N items.</summary>
    public int ProgressEvery { get; private set; } = 250;

    /// <summary>Build even when no date-write method verified.</summary>
    public bool AllowUndated { get; private set; }

    /// <summary>Actually write; without it every command dry-runs, as the rest of this console does.</summary>
    public bool Execute { get; private set; }

    /// <summary>Parses the arguments after the command word. Throws on anything unrecognised.</summary>
    public static CorpusOptions Parse(IEnumerable<string> args)
    {
        ArgumentNullException.ThrowIfNull(args);
        var options = new CorpusOptions();
        string? pending = null;
        foreach (string arg in args)
        {
            if (arg.StartsWith("--", StringComparison.Ordinal))
            {
                if (pending != null)
                {
                    options.SetFlag(pending);
                }

                pending = arg.Substring(2);
                continue;
            }

            if (pending == null)
            {
                throw new ArgumentException($"Unexpected argument '{arg}'.");
            }

            options.SetValue(pending, arg);
            pending = null;
        }

        if (pending != null)
        {
            options.SetFlag(pending);
        }

        return options;
    }

    /// <summary>Builds the plan options these arguments describe. Throws when one is missing.</summary>
    public CorpusPlanOptions ToPlanOptions()
    {
        if (string.IsNullOrWhiteSpace(CorpusId))
        {
            throw new ArgumentException("--corpus-id <id> is required.");
        }

        if (!HasSeed)
        {
            throw new ArgumentException("--seed <integer> is required - a corpus with no seed is not reproducible.");
        }

        if (AnchorUtc == null)
        {
            throw new ArgumentException(
                "--anchor <yyyy-MM-dd or yyyy-MM-ddTHH:mm:ssZ> is required. It is deliberately not defaulted to "
                + "today: an anchor taken from the clock would move every time the corpus was rebuilt, and the "
                + "same seed would stop meaning the same corpus.");
        }

        return new CorpusPlanOptions(CorpusId!, Seed, AnchorUtc.Value);
    }

    private void SetFlag(string name)
    {
        switch (name.ToLowerInvariant())
        {
            case "allow-undated":
                AllowUndated = true;
                break;
            case "execute":
                Execute = true;
                break;
            default:
                throw new ArgumentException($"--{name} needs a value.");
        }
    }

    private void SetValue(string name, string value)
    {
        switch (name.ToLowerInvariant())
        {
            case "store":
                Store = value;
                break;
            case "allow-store":
                AllowStores.Add(value);
                break;
            case "corpus-id":
                CorpusId = value;
                break;
            case "seed":
                Seed = long.Parse(value, CultureInfo.InvariantCulture);
                HasSeed = true;
                break;
            case "anchor":
                AnchorUtc = ParseAnchor(value);
                break;
            case "count":
                Count = int.Parse(value, CultureInfo.InvariantCulture);
                break;
            case "manifest":
                ManifestPath = value;
                break;
            case "progress-every":
                ProgressEvery = Math.Max(1, int.Parse(value, CultureInfo.InvariantCulture));
                break;
            default:
                throw new ArgumentException($"Unknown option --{name}.");
        }
    }

    /// <summary>
    /// Reads the anchor as UTC under the INVARIANT culture. Both properties matter: a
    /// machine-locale parse would read 2026-05-09 as 9 May here and 5 September elsewhere -
    /// the exact defect that once blew this project's sweep budget - and a local-time anchor
    /// would move the whole corpus when the VM's time zone or DST changed.
    /// </summary>
    public static DateTime ParseAnchor(string value)
    {
        if (!DateTime.TryParse(
                value,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal,
                out DateTime parsed))
        {
            throw new ArgumentException($"--anchor '{value}' is not a date. Use yyyy-MM-dd or yyyy-MM-ddTHH:mm:ssZ.");
        }

        return DateTime.SpecifyKind(parsed, DateTimeKind.Utc);
    }
}

/// <summary>
/// The corpus commands of the operator console. Each one prints what it decided before it
/// does anything, and each writes only after <c>--execute</c>, which is the same discipline
/// the incident-remediation commands beside them keep.
/// </summary>
public static class CorpusCommands
{
    /// <summary>
    /// <c>corpus-plan</c>: prints what the corpus WOULD be. Pure - no Outlook, no store, no
    /// writes. Run this first: it is the sheet a measurement is read against, and it costs
    /// nothing to get wrong here rather than after four hours of building.
    /// </summary>
    public static int RunPlan(CorpusOptions options, TextWriter output)
    {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(output);
        if (options.Count < 1)
        {
            throw new ArgumentException("--count <n> is required.");
        }

        var plan = new CorpusPlan(options.ToPlanOptions());
        CorpusPlanReport report = plan.Report(1, options.Count);
        WriteReport(plan, report, output);
        return 0;
    }

    /// <summary>
    /// Renders a plan report. Shared by <c>corpus-plan</c> and the build's preamble.
    /// <para>
    /// Numbers are formatted under the INVARIANT culture, not the machine's. This output is
    /// meant to be saved beside the measurement results and compared across machines, and a
    /// Dutch-locale VM would otherwise write 426.407.429 where an English one writes
    /// 426,407,429 - the same figure, but not the same string, and a reader comparing two
    /// runs should not have to work out which convention each was written under.
    /// </para>
    /// </summary>
    public static void WriteReport(CorpusPlan plan, CorpusPlanReport report, TextWriter output)
    {
        ArgumentNullException.ThrowIfNull(plan);
        ArgumentNullException.ThrowIfNull(report);
        ArgumentNullException.ThrowIfNull(output);
        CultureInfo invariant = CultureInfo.InvariantCulture;
        output.WriteLine($"== corpus plan '{plan.Options.CorpusId}' seed {plan.Options.Seed} "
            + $"anchor {CorpusManifest.FormatUtc(plan.Options.AnchorUtc)} ==");
        output.WriteLine("  items                 : " + report.ItemCount.ToString("N0", invariant));
        output.WriteLine("  body bytes (total)    : " + report.TotalBodyBytes.ToString("N0", invariant)
            + "  (mean " + report.MeanBodyBytes.ToString("N0", invariant) + ")");
        output.WriteLine("  bodies >= 24 KB       : " + report.BodiesAtLeast24Kb.ToString("N0", invariant));
        output.WriteLine("  bodies >= 96 KB       : " + report.BodiesAtLeast96Kb.ToString("N0", invariant));
        output.WriteLine("  bodies over sweep cap : " + report.BodiesOverSweepBodyCap.ToString("N0", invariant)
            + "  (cap " + OutlookAI.Core.Com.OutlookComSession.SweepBodyCharsCap.ToString("N0", invariant) + " chars)");
        output.WriteLine("  received range        : " + CorpusManifest.FormatUtc(report.OldestReceivedUtc)
            + " .. " + CorpusManifest.FormatUtc(report.NewestReceivedUtc));
        output.WriteLine("  per folder            : "
            + string.Join(", ", report.ByFolderId.Select(kv => FolderName(kv.Key) + "=" + kv.Value.ToString("N0", invariant))));
        output.WriteLine("  per size class        : "
            + string.Join(", ", report.BySizeClass.Select(kv => kv.Key + "=" + kv.Value.ToString("N0", invariant))));
        output.WriteLine("  per date band         : "
            + string.Join(", ", report.ByDateBand.Select(kv => kv.Key + "=" + kv.Value.ToString("N0", invariant))));
        output.WriteLine("  selected by window    : "
            + string.Join(", ", report.WithinDays.Select(kv => kv.Key + "d=" + kv.Value.ToString("N0", invariant))));
    }

    /// <summary>
    /// <c>corpus-probe</c>: settles whether this store will accept back-dated mail, and
    /// prints every rung's result. Creates and deletes a handful of throwaway items and
    /// nothing else. Always safe to run before a build - and the build runs it anyway.
    /// </summary>
    public static int RunProbe(CorpusOptions options, TextWriter output)
    {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(output);
        CorpusPlanOptions planOptions = options.ToPlanOptions();
        if (!Vet(options, output, out _))
        {
            return 1;
        }

        DateTime probeInstant = planOptions.AnchorUtc.AddDays(-30);
        IReadOnlyList<CorpusDateProbe> probes =
            ComCorpusMailbox.ProbeDateFidelity(options.Store!, planOptions.CorpusId, probeInstant);
        CorpusDateWriteMethod chosen = ReportProbes(probes, output);
        (bool proceed, string message) = CorpusDateFidelity.Decide(chosen, options.AllowUndated);
        output.WriteLine(message);
        return proceed ? 0 : 1;
    }

    /// <summary>
    /// <c>corpus-build</c>: creates the corpus. Vets the store, probes the dates, prints the
    /// plan, and only then - and only with <c>--execute</c> - writes anything. Resumable:
    /// it builds the ordinals the manifest does not already record, so re-running it after
    /// an interruption continues, and re-running it after completion does nothing.
    /// </summary>
    public static int RunBuild(CorpusOptions options, TextWriter output)
    {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(output);
        if (options.Count < 1)
        {
            throw new ArgumentException("--count <n> is required.");
        }

        if (string.IsNullOrWhiteSpace(options.ManifestPath))
        {
            throw new ArgumentException("--manifest <path> is required - without it nothing could ever be torn down.");
        }

        CorpusPlanOptions planOptions = options.ToPlanOptions();
        var plan = new CorpusPlan(planOptions);
        if (!Vet(options, output, out CorpusStoreFacts facts))
        {
            return 1;
        }

        WriteReport(plan, plan.Report(1, options.Count), output);

        CorpusManifest? existing = LoadManifest(options.ManifestPath!, output);
        if (existing != null)
        {
            CorpusManifestMismatch mismatch = existing.CheckCompatible(planOptions, options.Store!);
            if (mismatch != CorpusManifestMismatch.None)
            {
                output.WriteLine($"REFUSING to continue: {CorpusManifest.Explain(mismatch)}.");
                return 1;
            }

            output.WriteLine($"Existing manifest: {existing.Items.Count:N0} items, {existing.Folders.Count} created folders"
                + $"{(existing.UnparseableLines.Count > 0 ? $", {existing.UnparseableLines.Count} unreadable line(s)" : string.Empty)}.");
        }

        DateTime probeInstant = planOptions.AnchorUtc.AddDays(-30);
        IReadOnlyList<CorpusDateProbe> probes = options.Execute
            ? ComCorpusMailbox.ProbeDateFidelity(options.Store!, planOptions.CorpusId, probeInstant)
            : Array.Empty<CorpusDateProbe>();
        if (!options.Execute)
        {
            int todo = existing == null ? options.Count : existing.MissingOrdinals(options.Count).Count();
            output.WriteLine($"Dry-run complete; {todo:N0} item(s) would be created. Nothing written, and the date "
                + "probe was NOT run (it creates items). Re-run with --execute.");
            return 0;
        }

        CorpusDateWriteMethod chosen = ReportProbes(probes, output);
        (bool proceed, string message) = CorpusDateFidelity.Decide(chosen, options.AllowUndated);
        output.WriteLine(message);
        if (!proceed)
        {
            return 1;
        }

        TimeSpan writeShift = ShiftFrom(probes, chosen);
        if (writeShift != TimeSpan.Zero)
        {
            output.WriteLine($"Applying a write shift of {writeShift} so the stored date lands on the requested one "
                + "(the PropertyAccessor converted this store's write from local time).");
        }

        CorpusManifest manifest = existing ?? CorpusManifest.Create(new CorpusManifestHeader(
            CorpusManifest.CurrentVersion,
            planOptions.CorpusId,
            planOptions.Seed,
            CorpusManifest.FormatUtc(planOptions.AnchorUtc),
            planOptions.ShapeKey,
            options.Store!,
            facts.FilePath,
            chosen.ToString()));

        using StreamWriter writer = OpenManifest(options.ManifestPath!, existing == null, manifest.Header);
        ComCorpusMailbox.BuildOutcome outcome = ComCorpusMailbox.Build(
            plan,
            options.Store!,
            options.Count,
            chosen,
            writeShift,
            manifest,
            item =>
            {
                // Flushed per item on purpose. A build runs for hours and will be
                // interrupted; the cost of a flush is nothing beside a COM item write, and
                // an unflushed line is an item nothing can ever delete.
                writer.WriteLine(CorpusManifest.RenderLine(item));
                writer.Flush();
            },
            folder =>
            {
                writer.WriteLine(CorpusManifest.RenderLine(folder));
                writer.Flush();
            },
            p => output.WriteLine(
                $"  progress: created {p.Created:N0}, skipped {p.Skipped:N0}, failed {p.Failed:N0}, "
                + $"remaining {p.Remaining:N0}, {p.BodyBytesWritten:N0} body bytes, {p.Elapsed:hh\\:mm\\:ss} elapsed"),
            options.ProgressEvery);

        output.WriteLine($"Build finished: created {outcome.Created:N0}, already present {outcome.Skipped:N0}, "
            + $"failed {outcome.Failed:N0}, {outcome.BodyBytesWritten:N0} body bytes in {outcome.Elapsed:hh\\:mm\\:ss}"
            + $" ({Rate(outcome.Created, outcome.Elapsed)}).");
        if (outcome.FirstError != null)
        {
            output.WriteLine($"  first failure: {outcome.FirstError}");
        }

        return outcome.Failed == 0 ? 0 : 1;
    }

    /// <summary>
    /// <c>corpus-teardown</c>: removes exactly what the manifest records, then whatever
    /// those deletions turned into. Refuses without a manifest - deleting by subject alone
    /// is the thing the mailbox-safety rules forbid, and there is no flag for it.
    /// </summary>
    public static int RunTeardown(CorpusOptions options, TextWriter output)
    {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(output);
        if (string.IsNullOrWhiteSpace(options.ManifestPath))
        {
            throw new ArgumentException(
                "--manifest <path> is required. Teardown deletes by EntryID allowlist AND subject tag, both; "
                + "without the manifest there is no allowlist and the command has nothing it is allowed to do. "
                + "Use corpus-reindex to rebuild a candidate manifest by reading the store.");
        }

        if (!Vet(options, output, out _))
        {
            return 1;
        }

        CorpusManifest manifest = LoadManifest(options.ManifestPath!, output)
            ?? throw new ArgumentException($"Manifest not found: {options.ManifestPath}");
        CorpusPlanOptions planOptions = options.ToPlanOptions();
        CorpusManifestMismatch mismatch = manifest.CheckCompatible(planOptions, options.Store!);
        if (mismatch != CorpusManifestMismatch.None)
        {
            output.WriteLine($"REFUSING to tear down: {CorpusManifest.Explain(mismatch)}.");
            return 1;
        }

        output.WriteLine($"Manifest records {manifest.Items.Count:N0} item(s) and {manifest.Folders.Count} created folder(s).");
        if (!options.Execute)
        {
            IReadOnlyList<ComCorpusMailbox.ScanRow> present =
                ComCorpusMailbox.Scan(options.Store!, planOptions.CorpusId, manifest);
            output.WriteLine($"Dry-run: a read-only scan finds {present.Count:N0} corpus item(s) in the store. "
                + "Nothing deleted. Re-run with --execute.");
            return 0;
        }

        ComCorpusMailbox.TeardownOutcome outcome =
            ComCorpusMailbox.Teardown(options.Store!, planOptions.CorpusId, manifest);
        output.WriteLine($"Teardown: considered {outcome.Considered:N0}, deleted {outcome.Deleted:N0}, "
            + $"refused by rule {outcome.RefusedByRule:N0}, already gone {outcome.AlreadyGone:N0}, "
            + $"failed {outcome.Failed:N0}, folders removed {outcome.FoldersRemoved}.");
        output.WriteLine($"Post-teardown scan finds {outcome.RemainingInStore:N0} corpus item(s) remaining (expected 0).");
        return outcome.RemainingInStore == 0 && outcome.Failed == 0 ? 0 : 1;
    }

    /// <summary>
    /// <c>corpus-reindex</c>: READ-ONLY. Walks the store, finds every item whose subject
    /// parses as this corpus's, and writes a fresh manifest of what it found. The recovery
    /// path for a lost or truncated manifest - and it is deliberately a separate command
    /// producing a file a human can look at, rather than something teardown does for itself.
    /// </summary>
    public static int RunReindex(CorpusOptions options, TextWriter output)
    {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(output);
        if (string.IsNullOrWhiteSpace(options.ManifestPath))
        {
            throw new ArgumentException("--manifest <path> is required - it is where the rebuilt manifest is written.");
        }

        CorpusPlanOptions planOptions = options.ToPlanOptions();
        if (!Vet(options, output, out CorpusStoreFacts facts))
        {
            return 1;
        }

        IReadOnlyList<ComCorpusMailbox.ScanRow> rows =
            ComCorpusMailbox.Scan(options.Store!, planOptions.CorpusId, null);
        output.WriteLine($"Read-only scan found {rows.Count:N0} corpus item(s) "
            + $"across {rows.Select(r => r.FolderId).Distinct().Count()} folder(s).");
        if (!options.Execute)
        {
            output.WriteLine("Dry-run: no manifest written. Re-run with --execute.");
            return 0;
        }

        var header = new CorpusManifestHeader(
            CorpusManifest.CurrentVersion,
            planOptions.CorpusId,
            planOptions.Seed,
            CorpusManifest.FormatUtc(planOptions.AnchorUtc),
            planOptions.ShapeKey,
            options.Store!,
            facts.FilePath,
            "reindexed");
        using (StreamWriter writer = OpenManifest(options.ManifestPath!, writeHeader: true, header))
        {
            foreach (ComCorpusMailbox.ScanRow row in rows.OrderBy(r => r.Ordinal))
            {
                writer.WriteLine(CorpusManifest.RenderLine(
                    new CorpusManifestItem(row.Ordinal, row.EntryId, row.FolderId, 0, null)));
            }
        }

        output.WriteLine($"Wrote {rows.Count:N0} entries to {options.ManifestPath}. "
            + "Inspect it before handing it to corpus-teardown - it records what is in the store now, "
            + "not what a build claimed to create.");
        return 0;
    }

    /// <summary>
    /// The gate every corpus command passes through: the caller's allowlist AND four
    /// independent COM facts must agree the target is a local .pst. Prints the verdict
    /// either way.
    /// </summary>
    private static bool Vet(CorpusOptions options, TextWriter output, out CorpusStoreFacts facts)
    {
        if (string.IsNullOrWhiteSpace(options.Store))
        {
            throw new ArgumentException("--store <display name> is required.");
        }

        facts = ComCorpusMailbox.ReadStoreFacts(options.Store!);
        CorpusStoreRefusal refusal = CorpusSafety.EvaluateStore(facts, options.AllowStores);
        output.WriteLine(CorpusSafety.Explain(refusal, facts));
        if (refusal != CorpusStoreRefusal.None)
        {
            return false;
        }

        output.WriteLine($"  store file: {facts.FilePath}");
        return true;
    }

    private static CorpusDateWriteMethod ReportProbes(IReadOnlyList<CorpusDateProbe> probes, TextWriter output)
    {
        output.WriteLine("== date fidelity probe ==");
        foreach (CorpusDateProbe probe in probes)
        {
            output.WriteLine($"  {probe.Method,-28} requested {CorpusManifest.FormatUtc(probe.RequestedUtc)}"
                + $" wrote {CorpusManifest.FormatUtc(probe.WrittenUtc)}"
                + $" readBack {(probe.ReadBackReceivedUtc == null ? "(unreadable)" : CorpusManifest.FormatUtc(probe.ReadBackReceivedUtc.Value))}"
                + $" daslIn={probe.DaslSelectedInWindow} daslOut={probe.DaslExcludedOutsideWindow}"
                + $" usable={CorpusDateFidelity.IsUsable(probe)}"
                + (probe.Error == null ? string.Empty : $" error={probe.Error}"));
        }

        return CorpusDateFidelity.Choose(probes);
    }

    /// <summary>
    /// The correction the chosen rung needed, if any: the difference between what the probe
    /// asked for and what it had to write to get it. Zero on a store that stores what it is
    /// given.
    /// </summary>
    private static TimeSpan ShiftFrom(IReadOnlyList<CorpusDateProbe> probes, CorpusDateWriteMethod chosen)
    {
        foreach (CorpusDateProbe probe in probes)
        {
            if (probe.Method == chosen && CorpusDateFidelity.IsUsable(probe))
            {
                return probe.WrittenUtc - probe.RequestedUtc;
            }
        }

        return TimeSpan.Zero;
    }

    private static CorpusManifest? LoadManifest(string path, TextWriter output)
    {
        if (!File.Exists(path))
        {
            return null;
        }

        CorpusManifest manifest = CorpusManifest.Parse(File.ReadLines(path));
        if (manifest.UnparseableLines.Count > 0)
        {
            output.WriteLine($"  NOTE: {manifest.UnparseableLines.Count} manifest line(s) could not be read "
                + "(the shape an interrupted build leaves). Those items, if any, are not in the teardown allowlist - "
                + "corpus-reindex finds them.");
        }

        return manifest;
    }

    private static StreamWriter OpenManifest(string path, bool writeHeader, CorpusManifestHeader header)
    {
        string? directory = Path.GetDirectoryName(Path.GetFullPath(path));
        if (!string.IsNullOrEmpty(directory))
        {
            Directory.CreateDirectory(directory);
        }

        var writer = new StreamWriter(path, append: !writeHeader);
        if (writeHeader)
        {
            writer.WriteLine(CorpusManifest.RenderLine(header));
            writer.Flush();
        }

        return writer;
    }

    private static string Rate(int items, TimeSpan elapsed)
        => elapsed.TotalSeconds <= 0
            ? "rate unavailable"
            : (items / elapsed.TotalSeconds).ToString("F1", CultureInfo.InvariantCulture) + " items/s";

    private static string FolderName(int folderId) => folderId switch
    {
        3 => "Deleted Items",
        5 => "Sent Items",
        6 => "Inbox",
        23 => "Junk Email",
        _ => "folder " + folderId.ToString(CultureInfo.InvariantCulture),
    };
}
