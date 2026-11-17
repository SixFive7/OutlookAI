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
[Collection(LiveCollections.MoveArchive)]
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
    [Trait("Requires", "MultipleStores")]
    public void ArchiveResolution_AllFiveStores_ReadOnly()
    {
        List<string> stores = _fixture.Settings.ExpectedStoreDisplayNames
            .Concat(_fixture.Settings.ExpectedDelegateStoreDisplayNames)
            .ToList();
        // Two or more - what MultipleStores promises. It said three, the floor of the PST guests, until the
        // Exchange VM's two mailboxes met it on 2026-10-08.
        Assert.True(stores.Count >= 2, "live settings must name the account stores");
        IReadOnlyList<ComStoreDetail> details = _fixture.VerifySession.GetStoreDetails();

        foreach (string store in stores)
        {
            ComStoreDetail? detail = details.FirstOrDefault(
                d => string.Equals(d.DisplayName, store, StringComparison.OrdinalIgnoreCase));
            Assert.True(detail != null, $"store '{store}' is not in the profile");
            bool exchange = SpecialFolders.IsExchangeStore(detail!.ExchangeStoreType);

            // Q84: READ-ONLY means it. On a PST the old lookup made an Archive folder (and the
            // verification a Junk Email folder) in every store that lacked one, so a non-Exchange
            // store's whole folder list - walked from the root, which creates nothing - must read
            // the same after the lookup as before it. Exchange stores are not compared: a
            // delegate store's hierarchy syncs lazily (D42), so two walks of an unchanged store
            // can differ, and on Exchange the lookup is the unchanged GetDefaultFolder(39).
            IReadOnlyList<string>? foldersBefore = exchange ? null : FolderPathsOf(store);
            ComArchiveFolderInfo? archive = _fixture.VerifySession.TryResolveArchiveFolder(store, out string? error);
            if (foldersBefore != null)
            {
                AssertFolderListUnchanged(store, foldersBefore, FolderPathsOf(store));
            }

            if (archive == null)
            {
                // Only a non-Exchange store may have no designated Archive folder (Exchange keeps
                // one as a default folder), and only "none" is an answer: a designation that
                // could not be READ is the resolver failing on this store type, not a result.
                Assert.False(exchange, $"store '{store}': archive resolution failed ({error})");
                Assert.True(
                    error == ArchiveFolderResolution.NoDesignatedArchiveFolder,
                    $"store '{store}': archive resolution failed ({error}) - expected either a designated folder "
                    + "or " + ArchiveFolderResolution.NoDesignatedArchiveFolder);
                _output.WriteLine($"store='{store}' has no designated Archive folder - reported, nothing created");
                continue;
            }

            if (exchange)
            {
                // Exactly as before - Q84 leaves Exchange alone: Outlook's own GetDefaultFolder(39).
                Assert.Equal(ArchiveFolderResolution.ViaOutlookDefaultFolder, archive.Via);
            }
            else
            {
                // Found from the store's own designation, never by asking Outlook for the folder. On a
                // PST that is the Inbox's PR_ADDITIONAL_REN_ENTRYIDS_EX block 0x800F, where Outlook
                // records the Archive folder it made (measured 2026-10-03).
                Assert.Contains(
                    archive.Via,
                    new[]
                    {
                        ArchiveFolderResolution.ViaInboxArchiveProperty,
                        ArchiveFolderResolution.ViaInboxPersistData,
                        ArchiveFolderResolution.ViaStoreArchiveProperty,
                        ArchiveFolderResolution.ViaRootArchiveProperty,
                    });
            }

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

    /// <summary>
    /// A folder of one of <paramref name="stores"/> that move_mail takes as a target - read through the test's own
    /// read-only session (<see cref="OutlookComSession.TryDescribeFolder"/>): not a store's top, not refused by the
    /// move guard (Deleted Items or under it, the Outbox, not a mail folder). Fails when none of them has one.
    /// </summary>
    private (string Store, string FolderId) AFolderAMoveCouldTarget(IEnumerable<string> stores)
    {
        foreach (string store in stores)
        {
            foreach (FolderView folder in LiveIds.Folders(Service, store))
            {
                if (folder.FolderId == null
                    || !OutlookAI.Core.Mapi.OutlookIds.TryDecodeFolderId(folder.FolderId, "folder", out string hex, out _))
                {
                    continue;
                }

                ComFolderDetail? detail = _fixture.VerifySession.TryDescribeFolder(hex, out _);
                if (detail != null && !detail.IsStoreRoot && detail.RefusedAsTarget == null)
                {
                    return (store, folder.FolderId);
                }
            }
        }

        throw new InvalidOperationException(
            "No store but the hub has a folder move_mail could target, so the cross-store refusal cannot be shown here.");
    }

    /// <summary>A store's folder paths, walked from the root - a walk that creates nothing.</summary>
    private IReadOnlyList<string> FolderPathsOf(string store)
    {
        ComFolderPathList paths = _fixture.VerifySession.ListFolderPaths(store, MailService.FolderWalkAbsoluteCap);
        Assert.False(paths.Incomplete, $"store '{store}': the folder walk was cut short, so an unchanged list would prove nothing");
        return paths.Paths.OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToList();
    }

    /// <summary>
    /// Content-free on failure: counts, plus the names of folders that APPEARED - which can only
    /// be folders a lookup made, never mail - and nothing about the rest of the store.
    /// </summary>
    private static void AssertFolderListUnchanged(string store, IReadOnlyList<string> before, IReadOnlyList<string> after)
    {
        List<string> appeared = after.Except(before, StringComparer.OrdinalIgnoreCase).ToList();
        int vanished = before.Except(after, StringComparer.OrdinalIgnoreCase).Count();
        Assert.True(
            appeared.Count == 0 && vanished == 0,
            $"store '{store}': the READ-ONLY archive lookup changed the folder list ({before.Count} -> {after.Count}); "
            + $"appeared: [{string.Join(", ", appeared)}], vanished: {vanished}");
    }

    [Fact]
    [Trait("Requires", "MultipleStores")]
    [Trait("Requires", "Transport")]
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
            DateTime sentUtc = LiveOutlookTestMailer.SendSelfMail(
                Hub, seedSubject, "D39 move/archive seed body " + Marker, null);
            currentEntryId = LiveInboxArrival.WaitFor(_fixture.VerifySession, Hub, seedSubject, sentUtc).EntryId;
            _output.WriteLine("seed arrived in hub Inbox");
            string hubStoreId = LiveIds.StoreId(Service, Hub);

            // --- the target is looked up by NAME first, the way an agent does it (Q136 (ii)): find_folder
            // offers no such folder yet, and nothing was created by asking.
            FindFolderOutcome before = Service.FindFolder(_fixture.TestFolderName, hubStoreId);
            Assert.DoesNotContain(before.Candidates, c => string.Equals(c.Name, _fixture.TestFolderName, StringComparison.OrdinalIgnoreCase));

            // --- create_folder makes it, at the top of the hub, and gives its ID (Q134 (i)).
            CreateFolderOutcome made = LiveIds.CreateTestFolder(
                Service, Hub, LiveIds.RootFolderId(Service, Hub), _fixture.TestFolderName);
            Assert.True(made.Created);
            Assert.Equal(_fixture.TestFolderName, made.Name);
            Assert.Equal(_fixture.TestFolderName, made.Folder);
            Assert.Equal(hubStoreId, made.StoreId);
            string testFolderId = made.FolderId;
            FindFolderOutcome after = Service.FindFolder(_fixture.TestFolderName, hubStoreId);
            Assert.Equal(testFolderId, Assert.Single(after.Candidates, c => c.Match == "exact").FolderId);
            OperationOutcomeException secondCreate = Assert.Throws<OperationOutcomeException>(
                () => LiveIds.CreateTestFolder(Service, Hub, LiveIds.RootFolderId(Service, Hub), _fixture.TestFolderName));
            Assert.Equal(MutationOutcome.Unchanged, secondCreate.Outcome);
            Assert.Contains(testFolderId, secondCreate.Message, StringComparison.Ordinal);
            _output.WriteLine("create_folder made the test folder, find_folder found it by name, a second create was refused naming it");

            // --- move 1: Inbox -> test folder, by its ID.
            MoveMailOutcome move1 = Service.MoveMail(new[] { currentEntryId }, testFolderId);
            Assert.Equal(1, move1.Moved);
            MoveItemView item1 = Assert.Single(move1.Items);
            Assert.True(item1.Ok, item1.Error);
            Assert.Equal(Hub, item1.Store, ignoreCase: true);
            Assert.Equal(hubStoreId, item1.StoreId);
            Assert.Equal("Inbox", item1.FromFolder);
            Assert.Equal(LiveIds.FolderId(Service, Hub, "Inbox"), item1.FromFolderId);
            Assert.Equal(_fixture.TestFolderName, item1.ToFolder);
            Assert.Equal(testFolderId, item1.ToFolderId);
            Assert.Equal(testFolderId, move1.TargetFolderId);
            Assert.Equal(currentEntryId, item1.OldEntryId, ignoreCase: true);
            Assert.NotEqual(item1.OldEntryId, item1.NewEntryId, StringComparer.OrdinalIgnoreCase); // EntryIDs change on ANY move
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

            // --- move 2 (UNDO): back to fromFolderId using ONLY the result's undo info.
            MoveMailOutcome move2 = Service.MoveMail(new[] { item1.NewEntryId! }, item1.FromFolderId!);
            MoveItemView item2 = Assert.Single(move2.Items);
            Assert.True(item2.Ok, item2.Error);
            Assert.Equal(_fixture.TestFolderName, item2.FromFolder);
            Assert.Equal(testFolderId, item2.FromFolderId);
            Assert.Equal("Inbox", item2.ToFolder);
            Assert.NotEqual(item2.OldEntryId, item2.NewEntryId, StringComparer.OrdinalIgnoreCase);
            currentEntryId = item2.NewEntryId!;
            _output.WriteLine("move 2 ok: undo via fromFolderId landed back in Inbox");

            // --- cross-store guard (READ-ONLY for the other store): a target folder in another store is
            // refused per item, and the hub item stays untouched. The other store's folder is only read:
            // any folder a move could otherwise go to - not its top, not Deleted Items, not the Outbox, a
            // mail folder - since a refused target refuses the whole call before any item is looked at.
            // A data file added to the profile need not have an Inbox (the guests' bystander has none).
            (string otherStore, string otherFolderId) = AFolderAMoveCouldTarget(
                _fixture.Settings.ExpectedStoreDisplayNames.Where(s => !string.Equals(s, Hub, StringComparison.OrdinalIgnoreCase)));
            _output.WriteLine("cross-store target: a folder of '" + otherStore + "'");
            MoveMailOutcome cross = Service.MoveMail(new[] { currentEntryId }, otherFolderId);
            Assert.False(cross.Items[0].Ok);
            Assert.Contains("same-store only", cross.Items[0].Error, StringComparison.Ordinal);
            Assert.NotNull(_fixture.VerifySession.TryOpenItem(currentEntryId, _fixture.GetStoreId(Hub), out _));

            // --- Deleted Items target refusal (S1 v2: moves must never become deletes): the folder is
            // checked once, before any item, and the whole call is refused with nothing moved.
            OperationOutcomeException trash = Assert.Throws<OperationOutcomeException>(
                () => Service.MoveMail(new[] { currentEntryId }, LiveIds.FolderId(Service, Hub, "Deleted Items")));
            Assert.Equal(MutationOutcome.Unchanged, trash.Outcome);
            Assert.Contains("deletion semantics", trash.Message, StringComparison.Ordinal);
            Assert.NotNull(_fixture.VerifySession.TryOpenItem(currentEntryId, _fixture.GetStoreId(Hub), out _));

            // --- archive: lands in the hub's DESIGNATED Archive folder. Whether the hub HAS one
            // yet decides what archive_mail must report (Q84): a PST hub - the test guests - has
            // none until archive_mail creates it, and then createdFolders must say so; a hub that
            // has one (Exchange, or any later run) reports nothing created. Asked read-only, which
            // no longer creates it, so the question does not answer itself.
            ComArchiveFolderInfo? archiveBefore = _fixture.VerifySession.TryResolveArchiveFolder(Hub, out string? beforeError);
            Assert.True(
                archiveBefore != null || beforeError == ArchiveFolderResolution.NoDesignatedArchiveFolder,
                "hub archive lookup failed before archiving: " + beforeError);
            ArchiveMailOutcome archived = Service.ArchiveMail(new[] { currentEntryId });
            MoveItemView archivedItem = Assert.Single(archived.Items);
            Assert.True(archivedItem.Ok, archivedItem.Error);
            Assert.Equal(1, archived.Archived);

            // Resolved AFTER the move: the folder archive_mail used or made, read back independently.
            ComArchiveFolderInfo hubArchive = _fixture.VerifySession.TryResolveArchiveFolder(Hub, out string? resolveError)
                ?? throw new InvalidOperationException("hub archive resolution failed after archiving: " + resolveError);
            if (archiveBefore == null)
            {
                string created = Assert.Single(archived.CreatedFolders!);
                Assert.Equal(MailService.CreatedArchiveFolderLabel(Hub, hubArchive.StoreRelativePath), created, ignoreCase: true);
                _output.WriteLine($"archive_mail created and reported the hub's Archive folder: {created}");
            }
            else
            {
                Assert.Null(archived.CreatedFolders);
                Assert.Equal(archiveBefore.EntryId, hubArchive.EntryId, ignoreCase: true);
            }

            Assert.Equal("Inbox", archivedItem.FromFolder);
            Assert.Equal(hubArchive.StoreRelativePath, archivedItem.ToFolder);
            ArchiveFolderView hubView = Assert.Single(archived.ArchiveFolders!);
            Assert.Equal(Hub, hubView.Store, ignoreCase: true);
            Assert.Equal(hubStoreId, hubView.StoreId);
            Assert.Equal(hubArchive.StoreRelativePath, hubView.Folder);
            Assert.Equal(OutlookAI.Core.Mapi.OutlookIds.FolderIdFromEntryId(hubArchive.EntryId), hubView.FolderId);
            Assert.Equal(hubView.FolderId, archivedItem.ToFolderId);

            // The MOVE path resolves exactly as archive_mail always did, on every store type.
            Assert.Equal(ArchiveFolderResolution.ViaOutlookDefaultFolder, hubView.Via);

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
            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(
                Hub, Marker, folderIds: LiveOutlookTestMailer.HubSweepFolderIdsWithArchive);
            _output.WriteLine($"cleanup: foldersDeleted={foldersDeleted}");
        }

        // Post-suite (S3): zero tagged artifacts anywhere in the hub sweep set - the
        // hub Archive folder included - and zero test folders anywhere in the store.
        int remaining = LiveOutlookTestMailer.CountTaggedArtifactsAfterPurgingStragglers(
            Hub, Marker, LiveOutlookTestMailer.HubSweepFolderIdsWithArchive, out int stragglers);
        if (stragglers > 0)
        {
            _output.WriteLine($"post-suite: {stragglers} late-materialized self-send copy/copies purged (documented lag)");
        }

        Assert.Equal(0, remaining);
        // Folders: LIVE ones only - an empty test folder can wedge in Deleted Items
        // for the rest of an Outlook session (documented in DeleteTestFolders, which
        // tolerates and reports it). Asserting the raw count pins that limitation
        // instead of the contract; see CountLiveTestFolders.
        Assert.Equal(0, LiveOutlookTestMailer.CountLiveTestFolders(Hub, out int wedgedEmptyFolders));
        if (wedgedEmptyFolders > 0)
        {
            Console.WriteLine($"cleanup[{Hub}]: {wedgedEmptyFolders} empty test folder(s) wedged in Deleted Items "
                + "until Outlook restarts (documented same-session limitation, no items involved)");
        }
        _output.WriteLine("post-suite: 0 tagged artifacts (incl. Archive), 0 test folders");

        // Belt-and-braces (D39): whole-store snapshot reconciliation - every hub
        // folder's baseline-era count back at its pre-run value, zero marker strays.
        _output.WriteLine(_fixture.VerifyHubReconciled());
    }

    // The log THIS process appends to. The service runs in-process here, so since Q86 its lines
    // go to the test run's throwaway directory (AuditIsolation), never the machine's real log.
    private static int CountAuditLines()
    {
        string path = AuditLog.EffectiveLogPath;
        return File.Exists(path) ? File.ReadAllLines(path).Length : 0;
    }

    private static IReadOnlyList<string> ReadAuditLinesAfter(int skipCount)
    {
        string path = AuditLog.EffectiveLogPath;
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
