using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Shared state for the D39 move_mail/archive_mail live tier: the MailService under
/// test, an independent verify session, and one unique run marker (S3 double-match).
/// ALL move/archive writes target the test hub only (S2); the other four stores are
/// touched READ-ONLY (archive-folder resolution checks). Disposal purges this run's
/// tagged artifacts (Drafts/Inbox/Sent/Archive/Deleted) AND any leftover test folders.
///
/// Belt-and-braces (D39, user-mandated after the 7d incident): the constructor takes a
/// per-folder item-count SNAPSHOT of the ENTIRE hub store before any test write, and
/// <see cref="VerifyHubReconciled"/> re-walks the store after cleanup (each test's
/// post-cleanup assert AND fixture disposal) proving (a) zero items carrying this
/// run's marker anywhere in the store and (b) every folder's baseline-era item count
/// back at its pre-run value - so damage to REAL mail (which tag-based sweeps cannot
/// see) fails the run loudly. Writes are expected ONLY inside the hard folder
/// allowlist (Inbox, Sent Items, Drafts, Deleted Items as cleanup transit, the tagged
/// test folder, and the hub's designated Archive folder); a count change anywhere
/// else is reported as an allowlist escape. If the baseline snapshot cannot be taken,
/// the constructor throws and the whole collection is refused (D38 guard precedent).
/// </summary>
public sealed class LiveMoveArchiveFixture : IDisposable
{
    private readonly Lazy<OutlookComSession> _verifySession;
    private readonly Dictionary<string, int> _baselineCounts;
    private readonly DateTime _baselineUtc;
    private readonly string _archiveFolderFirstSegment;

    public LiveMoveArchiveFixture()
    {
        Settings = LiveTestSettings.Load();

        // Fail-closed per-store count tripwire: no census, no live tier. Cheap after
        // the first fixture (one process-wide baseline).
        LiveStoreCountTripwire.EnsureBaseline(Settings);
        Service = MailService.CreateDefault();
        RunMarker = "d39" + Guid.NewGuid().ToString("N").Substring(0, 14);
        _verifySession = new Lazy<OutlookComSession>(
            () => OutlookComSession.Connect(allowStartingOutlook: true),
            LazyThreadSafetyMode.ExecutionAndPublication);

        // Pre-clean leftover test folders from a crashed earlier run BEFORE the
        // baseline so their removal cannot skew the reconciliation (allowlist helper
        // only - never shell patterns, the 7d standing rule).
        LiveOutlookTestMailer.DeleteTestFolders(Settings.TestHubStoreDisplayName);

        // Baseline snapshot: per-folder mail-item counts across the WHOLE hub store.
        // A failure here throws out of the ctor and refuses the collection.
        _baselineCounts = CountByFolder(VerifySession.WalkStoreMailItems(Settings.TestHubStoreDisplayName));
        _baselineUtc = DateTime.UtcNow;

        // Resolved once (read-only) for allowlist classification in failure reports.
        ComArchiveFolderInfo hubArchive = VerifySession.TryResolveArchiveFolder(Settings.TestHubStoreDisplayName, out string? archiveError)
            ?? throw new InvalidOperationException("Hub archive folder resolution failed - refusing the live tier: " + archiveError);
        _archiveFolderFirstSegment = hubArchive.StoreRelativePath.Split('\\')[0];
    }

    public LiveTestSettings Settings { get; }

    public MailService Service { get; }

    /// <summary>Per-run unique marker; every artifact subject carries tag + marker (S3).</summary>
    public string RunMarker { get; }

    /// <summary>Independent COM session for verification (and read-only resolution checks).</summary>
    public OutlookComSession VerifySession => _verifySession.Value;

    /// <summary>Builds a tagged subject: [OutlookAI-McpTest] + run marker + label.</summary>
    public string TaggedSubject(string label)
    {
        return LiveOutlookTestMailer.SubjectTag + " " + RunMarker + " " + label;
    }

    /// <summary>This run's test folder name (carries the folder name prefix, D39).</summary>
    public string TestFolderName => LiveOutlookTestMailer.TestFolderNamePrefix;

    /// <summary>StoreID of a store by display name (via the verify session).</summary>
    public string GetStoreId(string storeDisplayName)
    {
        return VerifySession.GetStores()
                .FirstOrDefault(s => string.Equals(s.DisplayName, storeDisplayName, StringComparison.OrdinalIgnoreCase))?.StoreId
            ?? throw new InvalidOperationException("Store not found by display name.");
    }

    /// <summary>
    /// Post-cleanup reconciliation belt (call AFTER a test's own stable-zero sweep):
    /// re-walks the whole hub store and throws unless (a) zero items carry this run's
    /// marker anywhere and (b) every folder's baseline-era item count equals the
    /// pre-run snapshot. Late-materializing tagged copies (the documented sent-copy
    /// lag) get ONE extra stable-zero sweep before failing. Items that ARRIVED after
    /// the snapshot and carry no marker are reported as external arrivals, not
    /// failures. Returns a one-line report for test output.
    /// </summary>
    public string VerifyHubReconciled()
    {
        string hub = Settings.TestHubStoreDisplayName;
        IReadOnlyList<ComWalkedItem> walk = VerifySession.WalkStoreMailItems(hub);

        List<ComWalkedItem> strays = FindMarkerStrays(walk);
        if (strays.Count > 0)
        {
            // Benign race candidate: a tagged copy materialized after the test's
            // stable-zero window - sweep once more via the allowlisted helper, re-walk.
            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(
                hub, RunMarker, folderIds: LiveOutlookTestMailer.HubSweepFolderIdsWithArchive);
            walk = VerifySession.WalkStoreMailItems(hub);
            strays = FindMarkerStrays(walk);
        }

        if (strays.Count > 0)
        {
            IEnumerable<string> escaped = strays
                .Select(s => s.FolderPath + (IsAllowedWriteFolder(s.FolderPath) ? string.Empty : " [OUTSIDE WRITE ALLOWLIST]"))
                .Distinct(StringComparer.OrdinalIgnoreCase);
            throw new InvalidOperationException(
                $"Hub reconciliation FAILED: {strays.Count} item(s) carrying run marker {RunMarker} remain in: "
                + string.Join(", ", escaped));
        }

        Dictionary<string, int> post = CountByFolder(walk);
        if (CountsMatch(post, out _))
        {
            return $"hub reconciled: {_baselineCounts.Count} baseline folders, all item counts back at pre-run values";
        }

        // Counts moved - separate external arrivals (received after the snapshot,
        // unmarked) from real damage by re-counting baseline-era items only.
        Dictionary<string, int> refined = CountByFolder(
            walk.Where(i => i.ReceivedTime == null || i.ReceivedTime <= _baselineUtc));
        if (CountsMatch(refined, out _))
        {
            Dictionary<string, int> arrivals = new(StringComparer.OrdinalIgnoreCase);
            foreach (ComWalkedItem item in walk.Where(i => i.ReceivedTime != null && i.ReceivedTime > _baselineUtc))
            {
                arrivals[item.FolderPath] = arrivals.TryGetValue(item.FolderPath, out int n) ? n + 1 : 1;
            }

            return "hub reconciled: baseline-era counts match; external arrivals (not test writes): "
                + string.Join(", ", arrivals.Select(a => $"{a.Key}+{a.Value}"));
        }

        CountsMatch(refined, out List<string> damage);
        throw new InvalidOperationException(
            "Hub reconciliation FAILED - baseline-era item counts moved (pre -> post-cleanup): "
            + string.Join(", ", damage));
    }

    public void Dispose()
    {
        try
        {
            // Folders first (their tagged contents purge into Deleted Items), then
            // the item sweep - the same order the tests use.
            LiveOutlookTestMailer.DeleteTestFolders(Settings.TestHubStoreDisplayName);
        }
        catch (Exception)
        {
            // Best-effort final belt for folders.
        }

        try
        {
            LiveOutlookTestMailer.DeleteTaggedArtifacts(
                Settings.TestHubStoreDisplayName, RunMarker, LiveOutlookTestMailer.HubSweepFolderIdsWithArchive);
        }
        catch (Exception)
        {
            // Best-effort - each test already cleaned up and asserted in finally.
        }

        try
        {
            // Final belt: the snapshot reconciliation must hold at collection end;
            // a violation throws and fails the run loudly (D39 task order).
            VerifyHubReconciled();
        }
        finally
        {
            if (_verifySession.IsValueCreated)
            {
                // Releases COM references only - Outlook keeps running (S7: never kill/close).
                _verifySession.Value.Dispose();
            }

            Service.Dispose();
        }
    }

    private List<ComWalkedItem> FindMarkerStrays(IReadOnlyList<ComWalkedItem> walk)
    {
        return walk
            .Where(i => i.Subject != null && i.Subject.Contains(RunMarker, StringComparison.Ordinal))
            .ToList();
    }

    private bool CountsMatch(Dictionary<string, int> post, out List<string> deltas)
    {
        deltas = new List<string>();
        foreach (string folder in _baselineCounts.Keys.Union(post.Keys, StringComparer.OrdinalIgnoreCase))
        {
            int pre = _baselineCounts.TryGetValue(folder, out int p) ? p : 0;
            int now = post.TryGetValue(folder, out int n) ? n : 0;
            if (pre != now)
            {
                deltas.Add($"{folder}: {pre} -> {now}"
                    + (IsAllowedWriteFolder(folder) ? string.Empty : " [OUTSIDE WRITE ALLOWLIST]"));
            }
        }

        return deltas.Count == 0;
    }

    /// <summary>
    /// The hard write allowlist: hub Inbox, Sent Items, Drafts, Deleted Items
    /// (cleanup transit), the tagged test folder (any nesting), and the designated
    /// Archive folder. Used to CLASSIFY reconciliation failures - the count check
    /// itself covers every folder in the store.
    /// </summary>
    private bool IsAllowedWriteFolder(string folderPath)
    {
        if (folderPath.Contains(LiveOutlookTestMailer.TestFolderNamePrefix, StringComparison.Ordinal))
        {
            return true;
        }

        string first = folderPath.Split('\\')[0];
        return string.Equals(first, "Inbox", StringComparison.OrdinalIgnoreCase)
            || string.Equals(first, "Sent Items", StringComparison.OrdinalIgnoreCase)
            || string.Equals(first, "Drafts", StringComparison.OrdinalIgnoreCase)
            || string.Equals(first, "Deleted Items", StringComparison.OrdinalIgnoreCase)
            || string.Equals(first, _archiveFolderFirstSegment, StringComparison.OrdinalIgnoreCase);
    }

    private static Dictionary<string, int> CountByFolder(IEnumerable<ComWalkedItem> items)
    {
        Dictionary<string, int> counts = new(StringComparer.OrdinalIgnoreCase);
        foreach (ComWalkedItem item in items)
        {
            counts[item.FolderPath] = counts.TryGetValue(item.FolderPath, out int n) ? n + 1 : 1;
        }

        return counts;
    }
}

[CollectionDefinition("LiveMoveArchive")]
public sealed class LiveMoveArchiveCollection : ICollectionFixture<LiveMoveArchiveFixture>
{
}
