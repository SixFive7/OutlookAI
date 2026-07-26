using OutlookAI.Core.Audit;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// D39 T2 live acceptance - HUB ONLY writes (S2): seed mails, one tagged test folder,
/// the full move chain (Inbox -> test folder -> back: EntryID change + fromFolder undo
/// verified), archive into the hub's DESIGNATED Archive folder (resolution verified
/// live), the live guard errors (missing folder, Deleted Items target, cross-store,
/// already-archived), audit lines for every move, and read-only archive-folder
/// resolution across ALL FIVE stores. Everything created is deleted via the tested
/// allowlist helpers (items AND the test folder) with zero-remaining asserts.
/// </summary>
[Collection("LiveMoveArchive")]
[Trait("Category", "Live")]
public sealed class LiveMoveArchiveTests
{
    private readonly LiveMoveArchiveFixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveMoveArchiveTests(LiveMoveArchiveFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    private MailService Service => _fixture.Service;

    [Fact]
    public void ArchiveResolution_AllFiveStores_ReadOnly()
    {
        List<string> stores = _fixture.Settings.ExpectedStoreDisplayNames
            .Concat(_fixture.Settings.ExpectedDelegateStoreDisplayNames)
            .ToList();
        Assert.True(stores.Count >= 3, "live settings must name the account stores");

        foreach (string store in stores)
        {
            ComArchiveFolderInfo? archive = _fixture.VerifySession.TryResolveArchiveFolder(store, out string? error);
            Assert.True(archive != null, $"store '{store}': archive resolution failed ({error})");
            Assert.Equal("outlookDefaultFolder", archive!.Via);
            Assert.False(string.IsNullOrEmpty(archive.Name));
            Assert.False(string.IsNullOrEmpty(archive.StoreRelativePath));
            Assert.DoesNotContain('\\', archive.StoreRelativePath);
            Assert.True(archive.EntryId.Length >= 48, "folder EntryID expected");
            Assert.Equal(store, archive.StoreDisplayName);

            // Folder names of well-known folders are not message content (S4-consistent
            // with the store audits); they prove localization-proof resolution.
            _output.WriteLine($"store='{store}' archive='{archive.StoreRelativePath}' via={archive.Via}");
        }
    }

    [Fact]
    public void MoveChain_TestFolderRoundTrip_Archive_Guards_Audit_Cleanup()
    {
        // Pre-clean: a crashed earlier run may have left a test folder behind - the
        // createdFolders assert below needs a clean slate (allowlist helper only).
        LiveOutlookTestMailer.DeleteTestFolders(Hub);

        int auditLinesBefore = CountAuditLines();
        string seedSubject = _fixture.TaggedSubject("d39 move seed");
        string? currentEntryId = null;
        try
        {
            // --- seed: hub -> itself (D20), wait for the Inbox copy, take its REAL EntryID.
            LiveOutlookTestMailer.SendSelfMail(Hub, seedSubject, "D39 move/archive seed body " + Marker, null);
            currentEntryId = WaitForInboxSeed(seedSubject);
            _output.WriteLine("seed arrived in hub Inbox");

            // --- guard live: target folder missing without create_folder.
            MoveMailOutcome missing = Service.MoveMail(new[] { currentEntryId }, _fixture.TestFolderName);
            Assert.Equal(0, missing.Moved);
            Assert.False(missing.Items[0].Ok);
            Assert.Contains("create_folder=true", missing.Items[0].Error, StringComparison.Ordinal);

            // --- move 1: Inbox -> test folder, created on demand.
            MoveMailOutcome move1 = Service.MoveMail(new[] { currentEntryId }, _fixture.TestFolderName, createFolder: true);
            Assert.Equal(1, move1.Moved);
            MoveItemView item1 = Assert.Single(move1.Items);
            Assert.True(item1.Ok, item1.Error);
            Assert.Equal(Hub, item1.Store, ignoreCase: true);
            Assert.Equal("Inbox", item1.FromFolder);
            Assert.Equal(_fixture.TestFolderName, item1.ToFolder);
            Assert.Equal(currentEntryId, item1.OldEntryId, ignoreCase: true);
            Assert.NotEqual(item1.OldEntryId, item1.NewEntryId, StringComparer.OrdinalIgnoreCase); // EntryIDs change on ANY move
            Assert.NotNull(move1.CreatedFolders);
            Assert.Contains(_fixture.TestFolderName, move1.CreatedFolders!);
            Assert.NotNull(move1.Advice);
            Assert.Contains(MailService.MoveEntryIdAdvice, move1.Advice!);

            // Independent verify: new id opens with the seed subject, OLD id is stale.
            ComOpenResult? reopened = _fixture.VerifySession.TryOpenItem(item1.NewEntryId!, _fixture.GetStoreId(Hub), out string? openError);
            Assert.True(reopened != null, $"moved item must open by newEntryId ({openError})");
            Assert.Equal(seedSubject, reopened!.Subject);

            // The OLD EntryID is stale after a move. On cached Exchange the store MAY
            // keep answering old-id lookups for a short window - both shapes are
            // documented; what matters (and is asserted hard) is that the NEW id is
            // the item's identity and the ids differ.
            ComOpenResult? staleOpen = _fixture.VerifySession.TryOpenItem(item1.OldEntryId!, _fixture.GetStoreId(Hub), out string? staleError);
            _output.WriteLine($"move 1 ok: Inbox -> test folder, EntryID changed; old id open -> {(staleOpen == null ? "stale (" + staleError + ")" : "still mapped")}");

            // --- move 2 (UNDO): back to fromFolder using ONLY the result's undo info.
            MoveMailOutcome move2 = Service.MoveMail(new[] { item1.NewEntryId! }, item1.FromFolder!);
            MoveItemView item2 = Assert.Single(move2.Items);
            Assert.True(item2.Ok, item2.Error);
            Assert.Equal(_fixture.TestFolderName, item2.FromFolder);
            Assert.Equal("Inbox", item2.ToFolder);
            Assert.NotEqual(item2.OldEntryId, item2.NewEntryId, StringComparer.OrdinalIgnoreCase);
            currentEntryId = item2.NewEntryId!;
            _output.WriteLine("move 2 ok: undo via fromFolder landed back in Inbox");

            // --- cross-store guard (READ-ONLY for the other store): requesting another
            // store as target refuses per-item, the hub item stays untouched.
            string otherStore = _fixture.Settings.ExpectedStoreDisplayNames
                .First(s => !string.Equals(s, Hub, StringComparison.OrdinalIgnoreCase));
            MoveMailOutcome cross = Service.MoveMail(new[] { currentEntryId }, "Inbox", createFolder: false, store: otherStore);
            Assert.False(cross.Items[0].Ok);
            Assert.Contains("same-store only", cross.Items[0].Error, StringComparison.Ordinal);
            Assert.NotNull(_fixture.VerifySession.TryOpenItem(currentEntryId, _fixture.GetStoreId(Hub), out _));

            // --- Deleted Items target refusal (S1 v2: moves must never become deletes).
            MoveMailOutcome trash = Service.MoveMail(new[] { currentEntryId }, "Deleted Items");
            Assert.False(trash.Items[0].Ok);
            Assert.Contains("deletion semantics", trash.Items[0].Error, StringComparison.Ordinal);

            // --- archive: lands in the hub's DESIGNATED Archive folder.
            ComArchiveFolderInfo hubArchive = _fixture.VerifySession.TryResolveArchiveFolder(Hub, out string? resolveError)
                ?? throw new InvalidOperationException("hub archive resolution failed: " + resolveError);
            ArchiveMailOutcome archived = Service.ArchiveMail(new[] { currentEntryId });
            Assert.Equal(1, archived.Archived);
            MoveItemView archivedItem = Assert.Single(archived.Items);
            Assert.True(archivedItem.Ok, archivedItem.Error);
            Assert.Equal("Inbox", archivedItem.FromFolder);
            Assert.Equal(hubArchive.StoreRelativePath, archivedItem.ToFolder);
            ArchiveFolderView hubView = Assert.Single(archived.ArchiveFolders!);
            Assert.Equal(Hub, hubView.Store, ignoreCase: true);
            Assert.Equal(hubArchive.StoreRelativePath, hubView.Folder);
            Assert.Equal("outlookDefaultFolder", hubView.Via);

            // Independent verify: the item's parent folder IS the designated archive folder.
            ComDraftInfo? inArchive = _fixture.VerifySession.TryGetMailInfo(
                archivedItem.NewEntryId!, _fixture.GetStoreId(Hub), out string? infoError);
            Assert.True(inArchive != null, $"archived item must open ({infoError})");
            Assert.Equal(hubArchive.EntryId, inArchive!.ParentFolderEntryId, ignoreCase: true);
            currentEntryId = archivedItem.NewEntryId!;
            _output.WriteLine($"archive ok: landed in designated '{hubArchive.StoreRelativePath}' (via {hubArchive.Via})");

            // --- archiving an already-archived item refuses.
            ArchiveMailOutcome again = Service.ArchiveMail(new[] { currentEntryId });
            Assert.Equal(0, again.Archived);
            Assert.Contains("already in the target folder", again.Items[0].Error, StringComparison.Ordinal);

            // --- audit: one line per completed move, from -> to, op split move/archive.
            IReadOnlyList<string> auditLines = ReadAuditLinesAfter(auditLinesBefore);
            AssertMoveAuditLine(auditLines, "move_mail", item1.OldEntryId!, item1.NewEntryId!, "Inbox", _fixture.TestFolderName);
            AssertMoveAuditLine(auditLines, "move_mail", item2.OldEntryId!, item2.NewEntryId!, _fixture.TestFolderName, "Inbox");
            AssertMoveAuditLine(auditLines, "archive_mail", archivedItem.OldEntryId!, archivedItem.NewEntryId!, "Inbox", hubArchive.StoreRelativePath);
            Assert.Equal(2, auditLines.Count(l => l.Contains(" op=move_mail ", StringComparison.Ordinal)));
            Assert.Equal(1, auditLines.Count(l => l.Contains(" op=archive_mail ", StringComparison.Ordinal)));
            _output.WriteLine("audit: 2 move_mail + 1 archive_mail lines with exact from->to");
        }
        finally
        {
            // Folders FIRST (their tagged contents are purged into Deleted Items),
            // then the stable-zero item sweep (which covers Deleted Items and the
            // Archive folder). Allowlist helpers only - NEVER shell patterns (7d rule).
            if (currentEntryId != null)
            {
                try
                {
                    LiveOutlookTestMailer.DeleteItemByEntryId(Hub, currentEntryId, Marker);
                }
                catch (Exception)
                {
                    // The stable-zero sweep below is the authority.
                }
            }

            int foldersDeleted = LiveOutlookTestMailer.DeleteTestFolders(Hub);
            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(Hub, Marker);
            _output.WriteLine($"cleanup: foldersDeleted={foldersDeleted}");
        }

        // Post-suite (S3): zero tagged artifacts anywhere in the hub sweep set - the
        // hub Archive folder included - and zero test folders anywhere in the store.
        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
        Assert.Equal(0, LiveOutlookTestMailer.CountTestFolders(Hub));
        _output.WriteLine("post-suite: 0 tagged artifacts (incl. Archive), 0 test folders");
    }

    /// <summary>
    /// Waits for the self-send's Inbox copy and returns its REAL EntryID via a hub
    /// store walk (index-independent - the hub is tiny by design).
    /// </summary>
    private string WaitForInboxSeed(string seedSubject)
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(120);
        while (DateTime.UtcNow < deadline)
        {
            IReadOnlyList<ComWalkedItem> items = _fixture.VerifySession.WalkStoreMailItems(Hub);
            ComWalkedItem? seed = items.FirstOrDefault(i =>
                i.Subject == seedSubject
                && string.Equals(i.FolderPath, "Inbox", StringComparison.OrdinalIgnoreCase));
            if (seed != null)
            {
                return seed.EntryId;
            }

            Thread.Sleep(3000);
        }

        throw new TimeoutException("Seed mail did not arrive in the hub Inbox within 120 s.");
    }

    private static int CountAuditLines()
    {
        string path = AuditLog.DefaultLogPath;
        return File.Exists(path) ? File.ReadAllLines(path).Length : 0;
    }

    private static IReadOnlyList<string> ReadAuditLinesAfter(int skipCount)
    {
        string path = AuditLog.DefaultLogPath;
        Assert.True(File.Exists(path), "audit log must exist after write operations");
        return File.ReadAllLines(path).Skip(skipCount).ToList();
    }

    private static void AssertMoveAuditLine(
        IReadOnlyList<string> lines, string operation, string oldEntryId, string newEntryId, string fromFolder, string toFolder)
    {
        Assert.Contains(lines, l =>
            l.Contains(" op=" + operation + " ", StringComparison.Ordinal)
            && l.Contains("entryId=\"" + oldEntryId + "\"", StringComparison.OrdinalIgnoreCase)
            && l.Contains("newEntryId=\"" + newEntryId + "\"", StringComparison.OrdinalIgnoreCase)
            && l.Contains("fromFolder=\"" + fromFolder + "\"", StringComparison.Ordinal)
            && l.Contains("toFolder=\"" + toFolder + "\"", StringComparison.Ordinal));
    }
}
