using System.Globalization;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The live half of the artifact sweep - what <see cref="ArtifactSweepPolicy"/> is handed to count and
/// purge with - and, on the Exchange VM's test account, the two things a real mailbox adds to it
/// (2026-10-08, the Phase 2 write-safety design):
/// <list type="bullet">
/// <item><b>The purge takes only items of a RECORDED run</b>: the tag AND this run's marker, then the tag
/// AND each earlier run's marker the host recorded and handed in (<see cref="LiveRunMarker.Prior"/>) - an
/// aborted run's leftovers are taken back by the next run, by marker, never by a subject pattern. Every
/// delete is the mailer's tested one: an item is found by its subject, deleted by its EntryID, and only
/// when its subject carries both strings; Deleted Items is swept last, so a deleted test item is purged
/// from there too. Test folders left in the hub go with them, deepest first, and only when every item in
/// them carries the tag (<see cref="LiveOutlookTestMailer.DeleteTestFolders"/>).</item>
/// <item><b>Before the run, the hub is read</b> (<see cref="CheckHubBeforeTheRun"/>): leftovers of
/// recorded runs are reported and left for this run's sweep; a tagged item that carries no recorded marker
/// REFUSES the run - nothing will delete it, so the run's own zero-artifact proof could never pass.</item>
/// </list>
/// Everywhere else the sweep is what it always was: tag-matched, over the default folders.
/// </summary>
public static class LiveArtifactSweep
{
    /// <summary>
    /// What a count matches: the live-tier tag without its brackets (the DASL restriction matches a plain
    /// fragment). The same text the two sweeps always spelled out; <c>T1/CorpusTagSeparationTests</c> pins
    /// the corpus apart from it.
    /// </summary>
    public const string TagText = LiveOutlookTestMailer.TagText;

    /// <summary>
    /// The folders counted and swept on the Exchange VM: the hub's set with its archive folder, and Junk
    /// Email (23) - Exchange Online may file a self-sent test mail there, where no PST guest ever puts one -
    /// before Deleted Items, which stays last so the final pass purges what the others moved there.
    /// </summary>
    public static readonly int[] ExchangeFolders = { 16, 6, 5, 39, 4, 23, 20, 19, 21, 22, 3 };

    /// <summary>True on the machine whose hub is a real Exchange account (the Exchange VM's test account).</summary>
    public static bool SweepsByRecordedMarkers(LiveTestSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        return settings.MachineProfile == LiveMachineProfile.ExchangeTestAccount;
    }

    /// <summary>
    /// The folders counted and swept: there, the hub's set WITH its archive folder (a test account's
    /// archive is small, and archive_mail writes to it); elsewhere the mailer's default set.
    /// </summary>
    public static int[]? FoldersFor(LiveTestSettings settings)
    {
        return SweepsByRecordedMarkers(settings) ? ExchangeFolders : null;
    }

    /// <summary>
    /// The post-suite proof (S3) both sweeping suites run - LiveDraftTests and LiveSendTests, also run explicitly
    /// after the full suite: NO item tagged as a test artifact, from this or any earlier run, remains in
    /// Drafts/Inbox/Sent/Outbox/Deleted or the SYNC ISSUES subtree (Sync Issues, Conflicts, Local Failures, Server
    /// Failures - widened in soak fix 15 after a tagged item was found stranded in the hub's Local Failures folder,
    /// where nothing had ever swept) of ANY store. Self-send copies of earlier collections can MATERIALIZE with lag
    /// (Phase-2/Phase-4 documented race; worst on a store mid-resync), outliving their test's stable-zero window -
    /// such stragglers are purged here first (S3-legal: tag-matched), then stable zero is asserted. Counts only (S4).
    /// <para>
    /// WHICH stores are visited, and which of them may be purged, is <see cref="ArtifactSweepPolicy"/>'s decision and
    /// not this one's. It walks the COUNT TRIPWIRE'S OWN watched set - the primaries, the delegate/shared mailboxes
    /// and the declared bystanders - so the two guards cover one set rather than two overlapping ones. Everything the
    /// allowlist refuses a delete on is COUNTED and left alone, and a non-zero count in one of those FAILS the run
    /// rather than being tidied away: nothing else would report it, because the count tripwire fires on a decrease and
    /// an artifact arriving is an increase. On the Exchange VM's test account (2026-10-08) the purge takes the tag AND
    /// a recorded run's marker only, and a store nothing may write to may keep the tagged items it held before the run
    /// (<see cref="LiveArtifactBaseline"/>). The one copy of the two suites' test bodies (Q154).
    /// </para>
    /// </summary>
    public static void ProveNoTaggedArtifactRemains(LiveTestSettings settings, Action<string> output)
    {
        ArtifactSweepPolicy.Run(
            settings,
            store => CountTagged(settings, store),
            store => PurgeUntilStableZero(settings, store),
            output,
            LiveArtifactBaseline.Current);
    }

    /// <summary>Counts the tagged items in one store. Read-only.</summary>
    public static int CountTagged(LiveTestSettings settings, string store)
    {
        return LiveOutlookTestMailer.CountTaggedArtifacts(store, TagText, FoldersFor(settings));
    }

    /// <summary>Purges one store the plan says may be purged, until its count holds at zero.</summary>
    public static void PurgeUntilStableZero(LiveTestSettings settings, string store)
    {
        if (!SweepsByRecordedMarkers(settings))
        {
            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(store, TagText);
            return;
        }

        int[] folders = ExchangeFolders;
        LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(store, LiveRunMarker.Current, folderIds: folders);
        foreach (string prior in LiveRunMarker.Prior)
        {
            if (LiveOutlookTestMailer.CountTaggedArtifacts(store, prior, folders) > 0)
            {
                LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(store, prior, folderIds: folders);
            }
        }

        if (LiveOutlookTestMailer.CountTestFolders(store) > 0)
        {
            LiveOutlookTestMailer.DeleteTestFolders(store);
        }
    }

    /// <summary>
    /// Before any test writes, on the Exchange VM's test account only: the hub's tagged items, by recorded
    /// marker. Read-only. Throws - refusing the run - when a tagged item in the hub carries no marker of this
    /// run or of a recorded earlier one.
    /// </summary>
    public static void CheckHubBeforeTheRun(LiveTestSettings settings, Action<string> report)
    {
        ArgumentNullException.ThrowIfNull(settings);
        ArgumentNullException.ThrowIfNull(report);
        if (!SweepsByRecordedMarkers(settings))
        {
            return;
        }

        string hub = settings.TestHubStoreDisplayName;
        int[] folders = ExchangeFolders;
        int total = LiveOutlookTestMailer.CountTaggedArtifacts(hub, TagText, folders);
        int testFolders = LiveOutlookTestMailer.CountTestFolders(hub);
        Dictionary<string, int> byMarker = new(StringComparer.Ordinal);
        foreach (string marker in new[] { LiveRunMarker.Current }.Concat(LiveRunMarker.Prior))
        {
            byMarker[marker] = total == 0 ? 0 : LiveOutlookTestMailer.CountTaggedArtifacts(hub, marker, folders);
        }

        HubLeftovers leftovers = HubLeftovers.From(total, byMarker, testFolders, LiveRunMarker.Current);
        report("[artifacts] " + leftovers.Describe(hub));
        string? refusal = leftovers.Refusal(hub);
        if (refusal != null)
        {
            throw new InvalidOperationException(refusal);
        }
    }
}

/// <summary>
/// What the hub held before a run, by recorded marker - the decision <see cref="LiveArtifactSweep.CheckHubBeforeTheRun"/>
/// takes, pure so that T1 pins it.
/// </summary>
/// <param name="Total">Tagged items in the hub's sweep folders.</param>
/// <param name="FromRecordedRuns">Of those, the ones carrying an earlier recorded run's marker, by marker.</param>
/// <param name="FromThisRun">Of those, the ones carrying this run's marker (none, before the run).</param>
/// <param name="TestFolders">Test folders in the hub.</param>
public sealed record HubLeftovers(int Total, IReadOnlyDictionary<string, int> FromRecordedRuns, int FromThisRun, int TestFolders)
{
    /// <summary>The phrase the refusal carries.</summary>
    public const string Unrecorded = "TAGGED ITEM OF NO RECORDED RUN IN THE HUB";

    /// <summary>Tagged items whose marker no record names.</summary>
    public int OfNoRecordedRun => Math.Max(0, Total - FromThisRun - FromRecordedRuns.Values.Sum());

    /// <summary>Builds the reading from per-marker counts; <paramref name="current"/> is this run's marker.</summary>
    public static HubLeftovers From(int total, IReadOnlyDictionary<string, int> byMarker, int testFolders, string current)
    {
        ArgumentNullException.ThrowIfNull(byMarker);
        Dictionary<string, int> prior = byMarker
            .Where(kv => !string.Equals(kv.Key, current, StringComparison.Ordinal) && kv.Value > 0)
            .ToDictionary(kv => kv.Key, kv => kv.Value, StringComparer.Ordinal);
        int fromThisRun = byMarker.TryGetValue(current, out int mine) ? mine : 0;
        return new HubLeftovers(total, prior, fromThisRun, testFolders);
    }

    /// <summary>One line for the run's log: counts and markers, never a subject.</summary>
    public string Describe(string hub)
    {
        if (Total == 0 && TestFolders == 0)
        {
            return "hub '" + hub + "' before the run: no tagged item, no test folder.";
        }

        string recorded = FromRecordedRuns.Count == 0
            ? "none of recorded runs"
            : string.Join(", ", FromRecordedRuns.Select(kv => kv.Value.ToString(CultureInfo.InvariantCulture) + " of " + kv.Key))
                + " - left for this run's sweep, which takes them back by marker";
        return "hub '" + hub + "' before the run: " + Total.ToString(CultureInfo.InvariantCulture) + " tagged item(s) - "
            + recorded + "; " + OfNoRecordedRun.ToString(CultureInfo.InvariantCulture) + " of no recorded run; "
            + TestFolders.ToString(CultureInfo.InvariantCulture) + " test folder(s), which the sweep removes.";
    }

    /// <summary>The refusal, or null: only a tagged item of no recorded run refuses.</summary>
    public string? Refusal(string hub)
    {
        if (OfNoRecordedRun == 0)
        {
            return null;
        }

        return "REFUSING to run the live tier: " + Unrecorded + " - '" + hub + "' holds " + OfNoRecordedRun.ToString(CultureInfo.InvariantCulture)
            + " item(s) carrying the live-tier tag and neither this run's marker nor any marker the host recorded for an earlier "
            + "run (Testbed/host/Invoke-LiveTierOnGuest.ps1 hands those in). Nothing will delete them - a sweep takes back only "
            + "items of a recorded run, never by a subject pattern - so this run's zero-artifact proof could not pass. Find the "
            + "run that made them in the host's run ledger and hand its marker in, or leave them to the maintainer.";
    }
}

/// <summary>
/// The tagged items each store the sweep may NOT delete from held before the run - so that a tagged item
/// that was already in a read-only bystander (on the Exchange VM: one in the real mailbox's Sent Items, from the
/// workstation years) is reported and left alone, while one that ARRIVES during the run still fails it.
/// Captured once per process, before any test writes, on the Exchange VM's test account only; empty
/// everywhere else, where every such store must hold none, as before.
/// </summary>
public static class LiveArtifactBaseline
{
    private static readonly object Gate = new();
    private static IReadOnlyDictionary<string, int> _current = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// The counts captured, by store, less the confirmed removals made since (<see cref="ConfirmedRemovals"/>,
    /// Q138); empty until <see cref="Capture"/> ran.
    /// </summary>
    public static IReadOnlyDictionary<string, int> Current
    {
        get
        {
            lock (Gate)
            {
                return AdjustedFor(_current, ConfirmedRemovals.Current);
            }
        }
    }

    /// <summary>
    /// <paramref name="counts"/> with one taken off a store's count for each confirmed removal made in it - never
    /// below zero, and a store the counts do not name is left out. Pure, so T1 pins it.
    /// </summary>
    public static IReadOnlyDictionary<string, int> AdjustedFor(IReadOnlyDictionary<string, int> counts, IEnumerable<ConfirmedRemoval> removals)
    {
        ArgumentNullException.ThrowIfNull(counts);
        ArgumentNullException.ThrowIfNull(removals);
        Dictionary<string, int> adjusted = new(counts, StringComparer.OrdinalIgnoreCase);
        foreach (ConfirmedRemoval removal in removals)
        {
            if (removal != null && adjusted.TryGetValue(removal.Store, out int count))
            {
                adjusted[removal.Store] = Math.Max(0, count - 1);
            }
        }

        return adjusted;
    }

    /// <summary>Counts the tagged items of every store the sweep may not delete from. Read-only.</summary>
    public static void Capture(LiveTestSettings settings, Action<string> report)
    {
        ArgumentNullException.ThrowIfNull(settings);
        ArgumentNullException.ThrowIfNull(report);
        if (!LiveArtifactSweep.SweepsByRecordedMarkers(settings))
        {
            return;
        }

        Dictionary<string, int> counts = new(StringComparer.OrdinalIgnoreCase);
        foreach (ArtifactSweepStep step in ArtifactSweepPolicy.Assess(settings).Steps.Where(s => !s.MayDelete))
        {
            counts[step.Store] = LiveArtifactSweep.CountTagged(settings, step.Store);
        }

        lock (Gate)
        {
            _current = counts;
        }

        report("[artifacts] before the run, in stores no test may write to: " + (counts.Count == 0
            ? "none watched"
            : string.Join(", ", counts.Select(kv => "'" + kv.Key + "' " + kv.Value.ToString(CultureInfo.InvariantCulture) + " tagged item(s)")))
            + " - present before the run, left alone, and the sweep fails only on a change.");
    }
}
