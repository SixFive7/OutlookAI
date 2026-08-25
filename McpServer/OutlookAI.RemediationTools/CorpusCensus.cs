using System.Globalization;

namespace OutlookAI.RemediationTools;

/// <summary>One item a read-only scan found, reduced to what a census judges it on.</summary>
/// <param name="Ordinal">Ordinal parsed out of the subject.</param>
/// <param name="FolderId">Outlook default-folder id it was found under.</param>
public sealed record CorpusSighting(int Ordinal, int FolderId);

/// <summary>What one folder was supposed to hold and what it actually holds.</summary>
/// <param name="FolderId">Outlook default-folder id.</param>
/// <param name="Name">Readable folder name.</param>
/// <param name="Planned">How many items the plan puts here.</param>
/// <param name="Observed">How many the scan found here.</param>
public sealed record CorpusCensusFolder(int FolderId, string Name, int Planned, int Observed);

/// <summary>
/// What a corpus actually looks like in the store, measured against what it was supposed to
/// look like.
/// </summary>
/// <param name="PlannedItems">Ordinals the plan covers.</param>
/// <param name="PlannedUnread">How many of those the plan wants left unread.</param>
/// <param name="Sightings">Total items found, copies included.</param>
/// <param name="DistinctOrdinals">Distinct ordinals found.</param>
/// <param name="DuplicatedOrdinals">Ordinals found more than once - one item existing twice.</param>
/// <param name="MissingOrdinals">Ordinals the plan covers that the scan did not find anywhere.</param>
/// <param name="Misplaced">Sightings in a folder the plan does not put that ordinal in.</param>
/// <param name="Folders">Per-folder planned/observed, including the two folders nothing is ever planned into.</param>
/// <param name="StrayOutboxPlannedUnread">Outbox sightings whose ordinal the plan wants UNREAD.</param>
/// <param name="StrayOutboxPlannedRead">Outbox sightings whose ordinal the plan wants read.</param>
/// <param name="LegacyTagged">
/// Items found carrying the OLD corpus tag (<see cref="CorpusPlan.LegacySubjectTag"/>). They
/// are NOT sightings - nothing in this build may address them - so without this count they
/// would show up only as an equal number of missing ordinals, which reads as "the build never
/// happened" rather than "this corpus predates the tag split".
/// </param>
public sealed record CorpusCensusReport(
    int PlannedItems,
    int PlannedUnread,
    int Sightings,
    int DistinctOrdinals,
    int DuplicatedOrdinals,
    int MissingOrdinals,
    int Misplaced,
    IReadOnlyList<CorpusCensusFolder> Folders,
    int StrayOutboxPlannedUnread,
    int StrayOutboxPlannedRead,
    int LegacyTagged)
{
    /// <summary>Items sitting in the Outbox, where the plan never puts anything.</summary>
    public int StrayOutbox => StrayOutboxPlannedUnread + StrayOutboxPlannedRead;

    /// <summary>Items sitting in Drafts, where the plan never puts anything.</summary>
    public int StrayDrafts
        => Folders.Where(f => f.FolderId == CorpusCensus.DraftsFolderId).Sum(f => f.Observed);
}

/// <summary>
/// Counts what a build actually produced and says so out loud.
/// <para>
/// <b>Why this exists.</b> The first real build of 40 000 items had three faults and the
/// build reported success on all three. Every item was filed in Drafts, so the measurement
/// the corpus exists for selected six items instead of forty thousand. And 5 532 items were
/// additionally QUEUED FOR DELIVERY into the Outbox - inert only because that profile could
/// not send. Nothing in the tool noticed either; both were found by a person looking at
/// Outlook. This turns both into a number the build prints and a non-zero exit code.
/// </para>
/// <para>
/// <b>The Outbox split is a diagnosis, not decoration.</b> 5 532 is exactly the number of
/// items the plan for that shape marks UNREAD (seed 4242, 40 000 items - the plan report
/// prints it), and unread items are the only ones the builder treats differently. So the
/// census reports Outbox strays SPLIT by the plan's intended read state: if the next build
/// strays and the split is all-unread, the read-state write is the cause and nothing else
/// needs eliminating; if it strays evenly, that explanation is dead. One small build now
/// settles what a 12-minute build previously only hinted at.
/// </para>
/// </summary>
public static class CorpusCensus
{
    /// <summary>Outlook default-folder id for Drafts. Nothing is ever planned here.</summary>
    public const int DraftsFolderId = 16;

    /// <summary>Outlook default-folder id for the Outbox. Nothing is ever planned here either.</summary>
    public const int OutboxFolderId = 4;

    /// <summary>
    /// Compares a read-only scan against the plan.
    /// </summary>
    /// <param name="plan">The corpus shape - the authority on where each ordinal belongs.</param>
    /// <param name="itemCount">How many ordinals the corpus is supposed to hold.</param>
    /// <param name="sightings">Every item a scan found, one entry per copy.</param>
    /// <param name="legacyTagged">
    /// How many items the same scan found carrying the OLD corpus tag. Defaulted so the pure
    /// tests that only exercise the plan-vs-sightings arithmetic stay unchanged; the COM scan
    /// always passes the real number.
    /// </param>
    public static CorpusCensusReport Compare(
        CorpusPlan plan, int itemCount, IEnumerable<CorpusSighting> sightings, int legacyTagged = 0)
    {
        ArgumentNullException.ThrowIfNull(plan);
        ArgumentNullException.ThrowIfNull(sightings);
        if (itemCount < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(itemCount), "A corpus with no items has nothing to census.");
        }

        var plannedByFolder = new SortedDictionary<int, int>();
        var plannedFolderOf = new Dictionary<int, int>(itemCount);
        var plannedUnreadOrdinals = new HashSet<int>();
        for (int ordinal = 1; ordinal <= itemCount; ordinal++)
        {
            CorpusItemSpec spec = plan.Describe(ordinal);
            plannedFolderOf[ordinal] = spec.FolderId;
            plannedByFolder[spec.FolderId] = plannedByFolder.TryGetValue(spec.FolderId, out int n) ? n + 1 : 1;
            if (!spec.IsRead)
            {
                plannedUnreadOrdinals.Add(ordinal);
            }
        }

        var observedByFolder = new SortedDictionary<int, int>();
        var seen = new Dictionary<int, int>();
        int total = 0;
        int misplaced = 0;
        int outboxUnread = 0;
        int outboxRead = 0;
        foreach (CorpusSighting sighting in sightings)
        {
            total++;
            observedByFolder[sighting.FolderId] =
                observedByFolder.TryGetValue(sighting.FolderId, out int m) ? m + 1 : 1;
            seen[sighting.Ordinal] = seen.TryGetValue(sighting.Ordinal, out int c) ? c + 1 : 1;

            if (!plannedFolderOf.TryGetValue(sighting.Ordinal, out int want) || want != sighting.FolderId)
            {
                misplaced++;
            }

            if (sighting.FolderId == OutboxFolderId)
            {
                if (plannedUnreadOrdinals.Contains(sighting.Ordinal))
                {
                    outboxUnread++;
                }
                else
                {
                    outboxRead++;
                }
            }
        }

        var folders = new List<CorpusCensusFolder>();
        foreach (int folderId in plannedByFolder.Keys.Concat(observedByFolder.Keys).Distinct().OrderBy(f => f))
        {
            folders.Add(new CorpusCensusFolder(
                folderId,
                FolderName(folderId),
                plannedByFolder.TryGetValue(folderId, out int p) ? p : 0,
                observedByFolder.TryGetValue(folderId, out int o) ? o : 0));
        }

        int missing = 0;
        for (int ordinal = 1; ordinal <= itemCount; ordinal++)
        {
            if (!seen.ContainsKey(ordinal))
            {
                missing++;
            }
        }

        return new CorpusCensusReport(
            itemCount,
            plannedUnreadOrdinals.Count,
            total,
            seen.Count,
            seen.Values.Count(c => c > 1),
            missing,
            misplaced,
            folders,
            outboxUnread,
            outboxRead,
            legacyTagged);
    }

    /// <summary>
    /// Whether the corpus is what it claims to be, and what to print either way. Every
    /// consequence is a COUNT, never a description - the same correction the placement guard
    /// carries, made for the same reason.
    /// </summary>
    public static (bool Clean, string Message) Decide(CorpusCensusReport report)
    {
        ArgumentNullException.ThrowIfNull(report);
        CultureInfo invariant = CultureInfo.InvariantCulture;
        var faults = new List<string>();

        // First, because it EXPLAINS the faults underneath it. An old-tagged corpus otherwise
        // reports as "every ordinal missing", which reads as a build that never ran.
        if (report.LegacyTagged > 0)
        {
            faults.Add($"{report.LegacyTagged.ToString("N0", invariant)} item(s) carry the OLD corpus tag "
                + $"'{CorpusPlan.LegacySubjectTag}' instead of '{CorpusPlan.SubjectTag}', so this store holds a "
                + "corpus built before 2026-08-25. Nothing in this build may delete or rewrite them - the delete "
                + "and rewrite predicates require the current tag - and corpus-teardown will refuse. REBUILD: "
                + "remove the .pst and build a fresh corpus, which is the supported way to deal with a stale one");
        }

        if (report.StrayDrafts > 0)
        {
            faults.Add($"{report.StrayDrafts.ToString("N0", invariant)} item(s) are in DRAFTS, which the freshness "
                + "sweep does not cover - those items are invisible to the measurement this corpus exists for");
        }

        if (report.StrayOutbox > 0)
        {
            faults.Add($"{report.StrayOutbox.ToString("N0", invariant)} item(s) are in the OUTBOX, i.e. queued for "
                + $"delivery ({report.StrayOutboxPlannedUnread.ToString("N0", invariant)} of them ordinals the plan "
                + $"marks unread, {report.StrayOutboxPlannedRead.ToString("N0", invariant)} marked read, against "
                + $"{report.PlannedUnread.ToString("N0", invariant)} unread in the whole plan) - inert only while "
                + "the profile has no account, and real mail the moment one exists");
        }

        if (report.DuplicatedOrdinals > 0)
        {
            faults.Add($"{report.DuplicatedOrdinals.ToString("N0", invariant)} ordinal(s) exist more than once, so "
                + "the corpus holds more items than the plan describes and every per-item number measured against "
                + "it is wrong");
        }

        if (report.MissingOrdinals > 0)
        {
            faults.Add($"{report.MissingOrdinals.ToString("N0", invariant)} ordinal(s) are not in the store at all");
        }

        int misplacedElsewhere = report.Misplaced - report.StrayDrafts - report.StrayOutbox;
        if (misplacedElsewhere > 0)
        {
            faults.Add($"{misplacedElsewhere.ToString("N0", invariant)} item(s) are in a folder the plan does not "
                + "put them in");
        }

        string perFolder = string.Join(", ", report.Folders.Select(f =>
            f.Name + "=" + f.Observed.ToString("N0", invariant) + "/" + f.Planned.ToString("N0", invariant)));
        string head = $"Census: {report.Sightings.ToString("N0", invariant)} item(s) found for "
            + $"{report.PlannedItems.ToString("N0", invariant)} planned; per folder found/planned: {perFolder}.";

        return faults.Count == 0
            ? (true, head + " Every ordinal exists exactly once, in the folder the plan names.")
            : (false, head + " FAULTS: " + string.Join("; ", faults) + ".");
    }

    private static string FolderName(int folderId) => folderId switch
    {
        3 => "Deleted Items",
        OutboxFolderId => "Outbox",
        5 => "Sent Items",
        6 => "Inbox",
        DraftsFolderId => "Drafts",
        23 => "Junk Email",
        0 => "created folder",
        _ => "folder " + folderId.ToString(CultureInfo.InvariantCulture),
    };
}
