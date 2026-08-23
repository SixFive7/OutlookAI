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

    /// <summary>
    /// Where <c>corpus-reanchor</c> should move the corpus's newest edge to. <c>now</c> is
    /// accepted and is what an operator wants after restoring a checkpoint; a literal date is
    /// accepted so a re-anchor can be reproduced exactly.
    /// </summary>
    public DateTime? ToUtc { get; private set; }

    /// <summary>Whether a re-anchor may move the corpus BACKWARDS in time.</summary>
    public bool AllowBackwards { get; private set; }

    /// <summary>
    /// The windows a freshness check judges the corpus against, in days. Repeatable, and
    /// empty means <see cref="CorpusPlan.MeasurementWindowDays"/> - a caller that names its
    /// own set is saying which questions it actually asks.
    /// </summary>
    public List<int> Windows { get; } = new();

    /// <summary>Build even when no date-write method verified.</summary>
    public bool AllowUndated { get; private set; }

    /// <summary>
    /// Build even when no placement method put items in the folders the plan names, so every
    /// item is filed as a draft. There is no equivalent override for the store or profile
    /// guards, and deliberately so: this one costs a measurement, those cost real mail.
    /// </summary>
    public bool AllowDraftsPlacement { get; private set; }

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
            case "allow-drafts-placement":
                AllowDraftsPlacement = true;
                break;
            case "allow-backwards":
                AllowBackwards = true;
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
            case "to":
                ToUtc = string.Equals(value, "now", StringComparison.OrdinalIgnoreCase)
                    ? DateTime.UtcNow
                    : ParseAnchor(value);
                break;
            case "window":
                Windows.Add(int.Parse(value, CultureInfo.InvariantCulture));
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
        output.WriteLine("  unread                : " + report.UnreadItems.ToString("N0", invariant));
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

        // Placement FIRST, and the date probe inherits it. Probing dates against an item
        // that was filed somewhere other than the folder being queried cannot distinguish
        // "the date does not drive selection" from "the item is not in this folder", and the
        // first version of this tool reported the second as if it were the first.
        IReadOnlyList<CorpusPlacementProbe> placements =
            ComCorpusMailbox.ProbePlacement(options.Store!, planOptions.CorpusId);
        CorpusPlacementMethod placement = ReportPlacementProbes(placements, output);
        (bool placementOk, string placementMessage) =
            CorpusPlacement.Decide(placement, options.AllowDraftsPlacement, Math.Max(options.Count, 1), placements);
        output.WriteLine(placementMessage);

        DateTime probeInstant = planOptions.AnchorUtc.AddDays(-30);
        IReadOnlyList<CorpusDateProbe> probes =
            ComCorpusMailbox.ProbeDateFidelity(options.Store!, planOptions.CorpusId, probeInstant, placement);
        CorpusDateWriteMethod chosen = ReportProbes(probes, output);
        (bool dateOk, string message) =
            CorpusDateFidelity.Decide(chosen, options.AllowUndated, Math.Max(options.Count, 1));
        output.WriteLine(message);
        return placementOk && dateOk ? 0 : 1;
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

        if (!options.Execute)
        {
            int todo = existing == null ? options.Count : existing.MissingOrdinals(options.Count).Count();
            output.WriteLine($"Dry-run complete; {todo:N0} item(s) would be created. Nothing written, and neither "
                + "the placement probe nor the date probe was run (both create items). Re-run with --execute.");
            return 0;
        }

        IReadOnlyList<CorpusPlacementProbe> placements =
            ComCorpusMailbox.ProbePlacement(options.Store!, planOptions.CorpusId);
        CorpusPlacementMethod placement = ReportPlacementProbes(placements, output);
        (bool placementOk, string placementMessage) =
            CorpusPlacement.Decide(placement, options.AllowDraftsPlacement, options.Count, placements);
        output.WriteLine(placementMessage);
        if (!placementOk)
        {
            return 1;
        }

        DateTime probeInstant = planOptions.AnchorUtc.AddDays(-30);
        IReadOnlyList<CorpusDateProbe> probes =
            ComCorpusMailbox.ProbeDateFidelity(options.Store!, planOptions.CorpusId, probeInstant, placement);
        CorpusDateWriteMethod chosen = ReportProbes(probes, output);
        (bool proceed, string message) =
            CorpusDateFidelity.Decide(chosen, options.AllowUndated, options.Count);
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
            chosen.ToString(),
            placement.ToString()));

        using StreamWriter writer = OpenManifest(options.ManifestPath!, existing == null, manifest.Header);
        ComCorpusMailbox.BuildOutcome outcome = ComCorpusMailbox.Build(
            plan,
            options.Store!,
            options.Count,
            chosen,
            placement,
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

        // A build that reports success is not a build that produced a corpus. The first real
        // one created 40 000 items with zero failures and put every one of them in Drafts,
        // plus 5 532 copies in the Outbox, and said nothing. The census is a read-only scan
        // of what is now in the store, compared against the plan, and it decides the exit
        // code alongside the failure count.
        bool clean = RunCensusPass(options, plan, output);
        return outcome.Failed == 0 && clean ? 0 : 1;
    }

    /// <summary>
    /// <c>corpus-census</c>: READ-ONLY. Scans the store and says whether the corpus that is
    /// there is the corpus the plan describes - right count, right folders, one copy each,
    /// and nothing stranded in Drafts or the Outbox. Safe at any time; it is what a build
    /// runs on itself.
    /// </summary>
    public static int RunCensus(CorpusOptions options, TextWriter output)
    {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(output);
        if (options.Count < 1)
        {
            throw new ArgumentException("--count <n> is required - a census compares against a plan of a known size.");
        }

        var plan = new CorpusPlan(options.ToPlanOptions());
        if (!Vet(options, output, out _))
        {
            return 1;
        }

        return RunCensusPass(options, plan, output) ? 0 : 1;
    }

    /// <summary>The census itself, shared by <c>corpus-census</c> and the build's own check.</summary>
    private static bool RunCensusPass(CorpusOptions options, CorpusPlan plan, TextWriter output)
    {
        IReadOnlyList<ComCorpusMailbox.ScanRow> rows =
            ComCorpusMailbox.Scan(options.Store!, plan.Options.CorpusId, LoadManifest(options.ManifestPath, output));
        CorpusCensusReport census = CorpusCensus.Compare(
            plan, options.Count, rows.Select(r => new CorpusSighting(r.Ordinal, r.FolderId)));
        (bool clean, string message) = CorpusCensus.Decide(census);
        output.WriteLine(message);
        return clean;
    }

    /// <summary>
    /// <c>corpus-verify</c>: PURE. Reads the manifest, works out what shift the store already
    /// carries, and says whether the corpus can still answer the questions it exists for.
    /// No Outlook, no store, nothing written - so it can run anywhere, including as the first
    /// thing a test run does.
    /// <para>
    /// This is the check that converts a silent lie into a loud failure. A corpus anchored on
    /// a fixed date stops filling the narrow measurement windows a few weeks later, and every
    /// test asking about those windows keeps PASSING, because selecting nothing is a valid
    /// answer about an empty window.
    /// </para>
    /// </summary>
    public static int RunVerify(CorpusOptions options, TextWriter output)
    {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(output);
        if (options.Count < 1)
        {
            throw new ArgumentException("--count <n> is required.");
        }

        if (string.IsNullOrWhiteSpace(options.ManifestPath))
        {
            throw new ArgumentException(
                "--manifest <path> is required. The shift the store already carries is derived from what the "
                + "manifest records each item as holding; without it there is nothing to compare the plan against.");
        }

        var plan = new CorpusPlan(options.ToPlanOptions());
        CorpusManifest manifest = LoadManifest(options.ManifestPath!, output)
            ?? throw new ArgumentException($"Manifest not found: {options.ManifestPath}");

        (TimeSpan applied, bool provable) =
            CorpusReanchor.DeriveAppliedShift(plan, manifest, out int agreeing, out int dated);
        output.WriteLine($"Manifest records {manifest.Items.Count:N0} item(s), {dated:N0} of them dated; "
            + $"{agreeing:N0} agree on the shift now applied.");

        CorpusFreshnessReport report = CorpusFreshness.Evaluate(
            plan,
            options.Count,
            applied,
            DateTime.UtcNow,
            options.Windows.Count > 0 ? options.Windows : null,
            provable);
        (bool proceed, string message) = CorpusFreshness.Decide(report);
        output.WriteLine(message);
        return proceed ? 0 : 1;
    }

    /// <summary>
    /// <c>corpus-reanchor</c>: moves every item's received and submit instants forward so the
    /// corpus's newest edge lands where <c>--to</c> says, without regenerating anything.
    /// Idempotent and resumable - it writes the items whose recorded instant is not already
    /// the target one - and it never creates, moves or removes an item.
    /// </summary>
    public static int RunReanchor(CorpusOptions options, TextWriter output)
    {
        ArgumentNullException.ThrowIfNull(options);
        ArgumentNullException.ThrowIfNull(output);
        if (options.Count < 1)
        {
            throw new ArgumentException("--count <n> is required.");
        }

        if (string.IsNullOrWhiteSpace(options.ManifestPath))
        {
            throw new ArgumentException(
                "--manifest <path> is required. A re-anchor addresses items by the EntryIDs the manifest records, "
                + "and there is no other way for it to know what it is allowed to write to.");
        }

        if (options.ToUtc == null)
        {
            throw new ArgumentException("--to <now|yyyy-MM-dd|yyyy-MM-ddTHH:mm:ssZ> is required.");
        }

        CorpusPlanOptions planOptions = options.ToPlanOptions();
        var plan = new CorpusPlan(planOptions);
        if (!Vet(options, output, out _))
        {
            return 1;
        }

        CorpusManifest manifest = LoadManifest(options.ManifestPath!, output)
            ?? throw new ArgumentException($"Manifest not found: {options.ManifestPath}");
        CorpusManifestMismatch mismatch = manifest.CheckCompatible(planOptions, options.Store!);
        if (mismatch != CorpusManifestMismatch.None)
        {
            output.WriteLine($"REFUSING to re-anchor: {CorpusManifest.Explain(mismatch)}.");
            return 1;
        }

        (TimeSpan applied, bool provable) =
            CorpusReanchor.DeriveAppliedShift(plan, manifest, out int agreeing, out int dated);
        output.WriteLine($"Manifest records {manifest.Items.Count:N0} item(s), {dated:N0} dated, "
            + $"{agreeing:N0} agreeing on the shift now applied.");
        if (!provable)
        {
            output.WriteLine("REFUSING: the manifest cannot say what shift the store already carries, so this run "
                + "could not tell an item it has already moved from one it has not. Rebuild the manifest with "
                + "corpus-reindex first.");
            return 1;
        }

        CorpusReanchorPlan work = CorpusReanchor.Build(plan, manifest, options.Count, options.ToUtc.Value);
        (bool proceed, string message) = CorpusReanchor.Decide(work, applied, options.AllowBackwards);
        output.WriteLine(message);
        if (!proceed)
        {
            return 1;
        }

        if (!options.Execute)
        {
            output.WriteLine("Dry-run complete; nothing written, and the date probe was not run (it creates items). "
                + "Re-run with --execute.");
            return 0;
        }

        // The same two probes the build runs, and for the same reason: a re-anchor writes the
        // same two date properties through the same PropertyAccessor, so it needs the rung
        // this store verified AND the local-time compensation that rung required. Skipping
        // them would move the whole corpus by the machine's UTC offset.
        IReadOnlyList<CorpusPlacementProbe> placements =
            ComCorpusMailbox.ProbePlacement(options.Store!, planOptions.CorpusId);
        CorpusPlacementMethod placement = ReportPlacementProbes(placements, output);
        DateTime probeInstant = work.TargetAnchorUtc.AddDays(-30);
        IReadOnlyList<CorpusDateProbe> probes =
            ComCorpusMailbox.ProbeDateFidelity(options.Store!, planOptions.CorpusId, probeInstant, placement);
        CorpusDateWriteMethod chosen = ReportProbes(probes, output);
        (bool dateOk, string dateMessage) = CorpusDateFidelity.Decide(chosen, options.AllowUndated, options.Count);
        output.WriteLine(dateMessage);
        if (!dateOk)
        {
            return 1;
        }

        TimeSpan writeShift = ShiftFrom(probes, chosen);
        if (writeShift != TimeSpan.Zero)
        {
            output.WriteLine($"Applying a write shift of {writeShift} so the stored date lands on the requested one.");
        }

        using StreamWriter writer = OpenManifest(options.ManifestPath!, writeHeader: false, manifest.Header);
        ComCorpusMailbox.ReanchorOutcome outcome = ComCorpusMailbox.Reanchor(
            work,
            options.Store!,
            planOptions.CorpusId,
            chosen,
            writeShift,
            CorpusSafety.BuildEntryIdAllowlist(manifest.EntryIds),
            item =>
            {
                // Flushed per item, exactly as the build does: the run is long, it will be
                // interrupted, and an unflushed line is an item the next run would rewrite
                // for no reason.
                writer.WriteLine(CorpusManifest.RenderLine(item));
                writer.Flush();
            },
            p => output.WriteLine(
                $"  progress: rewritten {p.Created:N0}, remaining {p.Remaining:N0}, failed {p.Failed:N0}, "
                + $"{p.Elapsed:hh\\:mm\\:ss} elapsed"),
            options.ProgressEvery);

        output.WriteLine($"Re-anchor finished: rewritten {outcome.Rewritten:N0}, already correct "
            + $"{outcome.AlreadyCorrect:N0}, refused by rule {outcome.Refused:N0}, already gone {outcome.Gone:N0}, "
            + $"failed {outcome.Failed:N0}, in {outcome.Elapsed:hh\\:mm\\:ss}.");
        if (outcome.FirstError != null)
        {
            output.WriteLine($"  first failure: {outcome.FirstError}");
        }

        return outcome.Failed == 0 && outcome.Refused == 0 && outcome.Gone == 0 ? 0 : 1;
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
            "reindexed",
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
        CorpusProfileFacts profile = ComCorpusMailbox.ReadProfileFacts(options.Store!);
        CorpusStoreRefusal refusal = CorpusSafety.Evaluate(facts, profile, options.AllowStores);
        output.WriteLine(CorpusSafety.Explain(refusal, facts));
        output.WriteLine($"  profile accounts: {(profile.AccountCount == null ? "(unreadable)" : profile.AccountCount)}"
            + $", delivering into this store: {profile.AccountsDeliveringToTarget}"
            + $", unreadable delivery store: {profile.AccountsWithUnreadableDeliveryStore}");
        if (refusal != CorpusStoreRefusal.None)
        {
            return false;
        }

        output.WriteLine($"  store file: {facts.FilePath}");
        return true;
    }

    /// <summary>
    /// Prints every placement rung's result and returns the one to build with. The
    /// "landed in" column is the useful one when a rung fails: it says where the store put
    /// the item instead, which is the difference between a diagnosis and a shrug.
    /// </summary>
    private static CorpusPlacementMethod ReportPlacementProbes(
        IReadOnlyList<CorpusPlacementProbe> probes, TextWriter output)
    {
        output.WriteLine("== placement probe ==");
        foreach (CorpusPlacementProbe probe in probes)
        {
            output.WriteLine($"  {probe.Method,-28} target={probe.TargetFolderName}"
                + $" landedIn={probe.LandedInFolderName ?? "(unknown)"}"
                + $" parentMatches={probe.ParentIsTargetFolder}"
                + $" inFolderTable={probe.TargetFolderTableContainsIt}"
                + $" sentFlag={probe.SentFlagSet}"
                + $" usable={CorpusPlacement.IsUsable(probe)}"
                + (probe.Error == null ? string.Empty : $" error={probe.Error}"));
        }

        return CorpusPlacement.Choose(probes);
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

    private static CorpusManifest? LoadManifest(string? path, TextWriter output)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
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
