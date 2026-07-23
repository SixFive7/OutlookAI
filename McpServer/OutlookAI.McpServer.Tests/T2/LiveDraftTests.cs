using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Phase-4 T2 live acceptance (v3.MD section 0.6 Phase 4). Full draft matrix on the
/// designated test hub (S2): new_draft lands in the hub's Drafts with the hub identity
/// and signature state recorded; reply/replyall/forward derive from a self-sent seed
/// (D20) with ConversationIndex EXTENDING the original, quoted history present and
/// agent text above the quote; ONE display case opens an Inspector and takes the S5
/// screenshot (agent-authored content only), then the Inspector is closed via the
/// Phase-3 helper. Identity-only checks create one tagged, never-displayed draft in
/// each business account (Q-it2-3a), property-asserted content-free (S4) and deleted
/// immediately. Every artifact carries tag + run marker and is deleted after assert
/// (S3); each test ends by proving 0 marker artifacts remain in the folders it touched.
/// </summary>
[Collection("LivePhase4")]
[Trait("Category", "Live")]
public sealed class LiveDraftTests
{
    private const int OlFolderDrafts = 16;

    private readonly LivePhase4Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveDraftTests(LivePhase4Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    [Fact]
    public void NewDraft_Hub_LandsInHubDrafts_IdentityAndSignatureRecorded()
    {
        ComDefaultFolderInfo hubDrafts = RequireDefaultFolder(Hub, OlFolderDrafts);
        string hubStoreId = _fixture.GetStoreId(Hub);
        string agentBody = "Agent-authored Phase-4 body " + Marker + "\r\nSecond line for break handling.";
        string subject = _fixture.TaggedSubject("new-draft");

        DraftOutcome outcome = Service.NewDraft(Hub, Hub, cc: null, subject, agentBody, display: false);
        try
        {
            Assert.Equal("new", outcome.Kind);
            Assert.True(outcome.EntryId.Length >= 48, "draft must carry a real EntryID");
            Assert.Equal(Hub, outcome.Store, ignoreCase: true);
            Assert.True(outcome.AccountResolved, "SendUsingAccount must be pinned from the Account object");
            Assert.False(outcome.Displayed);

            // Persisted state through the INDEPENDENT session (the acceptance asserts).
            ComDraftInfo reopened = RequireMailInfo(outcome.EntryId, hubStoreId);
            Assert.Equal(hubDrafts.EntryId, reopened.ParentFolderEntryId, ignoreCase: true);
            string? persistedAccount = reopened.SendUsingAccountSmtp ?? outcome.Account;
            Assert.Equal(Hub, persistedAccount, ignoreCase: true);
            _output.WriteLine(
                $"new_draft: inHubDrafts=true accountPersisted={reopened.SendUsingAccountSmtp != null} account={persistedAccount}");

            // Signature findings (empirical primary; config probe for names - S4: names
            // only, never signature content of any account).
            ReadOutcome read = Service.Read(outcome.EntryId);
            Assert.Contains(agentBody.Substring(0, 30), read.Body, StringComparison.Ordinal);
            (string? newSig, string? replySig) = SignatureConfigProbe.AssignedSignatures(Hub);
            _output.WriteLine(
                $"signature[hub new]: injected={outcome.SignatureInjected} bodyTotalChars={read.BodyTotalChars} "
                + $"assignedNew='{newSig ?? "-"}' assignedReply='{replySig ?? "-"}' files={string.Join("|", SignatureConfigProbe.SignatureFileBaseNames())}");
            if (outcome.SignatureInjected)
            {
                Assert.True(read.BodyTotalChars > agentBody.Length, "injected signature must add body text");
            }

            // Recipient echo: the single To recipient is the hub itself.
            Assert.NotNull(outcome.Recipients);
            Assert.Contains(outcome.Recipients!, r => r.Address != null
                && r.Address.IndexOf(Hub, StringComparison.OrdinalIgnoreCase) >= 0);
        }
        finally
        {
            CleanupDraft(Hub, outcome.EntryId);
        }

        AssertGone(outcome.EntryId, hubStoreId);
        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
        _output.WriteLine("new_draft: artifact deleted, 0 marker artifacts remain in hub");
    }

    [Fact]
    public void DerivedDrafts_Hub_ThreadingQuotedHistoryAndPlacement()
    {
        ComDefaultFolderInfo hubDrafts = RequireDefaultFolder(Hub, OlFolderDrafts);
        string hubStoreId = _fixture.GetStoreId(Hub);
        string quoteToken = "QT" + Marker;
        string seedSubject = _fixture.TaggedSubject("seed");
        string seedBody = "Seed body for derived drafts.\r\nUnique quote token: " + quoteToken + "\r\nEnd of seed.";

        // D20 grant: hub -> itself; the reply target is the ARRIVED Inbox copy.
        DateTime sentUtc = LiveOutlookTestMailer.SendSelfMail(Hub, seedSubject, seedBody, attachmentPath: null);
        ComMailBrief seed = WaitForInboxArrival(seedSubject, sentUtc);
        _output.WriteLine($"seed arrived: folderKind={seed.FolderKind} afterSeconds={(DateTime.UtcNow - sentUtc).TotalSeconds:F1}");

        ComDraftInfo seedInfo = RequireMailInfo(seed.EntryId, seed.StoreId ?? hubStoreId);
        Assert.False(string.IsNullOrEmpty(seedInfo.ConversationIndex), "seed must carry a ConversationIndex");

        var draftIds = new List<string>();
        try
        {
            // --- reply_draft
            DraftOutcome reply = Service.ReplyDraft(seed.EntryId, "Reply agent text " + Marker, replyAll: false, display: false);
            draftIds.Add(reply.EntryId);
            AssertDerivedDraft(reply, "reply", seedInfo, hubDrafts, hubStoreId, quoteToken, "Reply agent text " + Marker);
            Assert.NotNull(reply.Recipients);
            Assert.All(
                reply.Recipients!.Where(r => r.Kind == "to"),
                r => Assert.True(r.Address != null && r.Address.IndexOf(Hub, StringComparison.OrdinalIgnoreCase) >= 0,
                    "reply recipient must be the hub (self-sent seed)"));

            // --- replyall_draft
            DraftOutcome replyAll = Service.ReplyDraft(seed.EntryId, "ReplyAll agent text " + Marker, replyAll: true, display: false);
            draftIds.Add(replyAll.EntryId);
            AssertDerivedDraft(replyAll, "replyall", seedInfo, hubDrafts, hubStoreId, quoteToken, "ReplyAll agent text " + Marker);
            Assert.True(replyAll.Recipients!.Count >= 1, "replyall must carry recipients");
            Assert.All(
                replyAll.Recipients!,
                r => Assert.True(r.Address == null || r.Address.IndexOf(Hub, StringComparison.OrdinalIgnoreCase) >= 0,
                    "self-sent replyall recipients must all be the hub"));

            // --- forward_draft
            DraftOutcome forward = Service.ForwardDraft(seed.EntryId, "Forward agent text " + Marker, to: Hub, display: false);
            draftIds.Add(forward.EntryId);
            AssertDerivedDraft(forward, "forward", seedInfo, hubDrafts, hubStoreId, quoteToken, "Forward agent text " + Marker);
            Assert.Contains(forward.Recipients!, r => r.Kind == "to" && r.Address != null
                && r.Address.IndexOf(Hub, StringComparison.OrdinalIgnoreCase) >= 0);

            (string? newSig, string? replySig) = SignatureConfigProbe.AssignedSignatures(Hub);
            _output.WriteLine(
                $"signature[hub derived]: reply={reply.SignatureInjected} replyall={replyAll.SignatureInjected} "
                + $"forward={forward.SignatureInjected} assignedReply='{replySig ?? "-"}' assignedNew='{newSig ?? "-"}'");
        }
        finally
        {
            foreach (string draftId in draftIds)
            {
                CleanupDraft(Hub, draftId);
            }

            // Seed copies (Inbox + Sent) + purge pass over Deleted Items.
            int deleted = LiveOutlookTestMailer.DeleteTaggedArtifacts(Hub, Marker);
            _output.WriteLine($"cleanup: taggedArtifactsDeleted={deleted}");
        }

        foreach (string draftId in draftIds)
        {
            AssertGone(draftId, hubStoreId);
        }

        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
        _output.WriteLine("derived drafts: threading verified (ConversationIndex child-of), 0 marker artifacts remain");
    }

    [Fact]
    public void NewDraft_Hub_DisplayCase_InspectorShown_ScreenshotTaken_ThenClosed()
    {
        string hubStoreId = _fixture.GetStoreId(Hub);
        string subject = _fixture.TaggedSubject("display-case");
        string agentBody = "Display-case draft, agent-authored content only. " + Marker;

        // The ONE .Display() case (D4 default behavior; S5: the window shows only
        // agent-authored content plus the hub's own signature).
        DraftOutcome outcome = Service.NewDraft(Hub, Hub, cc: null, subject, agentBody, display: true);
        try
        {
            Assert.True(outcome.Displayed);

            ComInspectorInfo? inspector = PollForInspector(outcome.EntryId, present: true, TimeSpan.FromSeconds(15));
            Assert.True(inspector != null, "no Inspector appeared for the displayed draft within 15 s");
            _output.WriteLine($"display case: inspector entryIdMatches=true itemClass={inspector!.ItemClass}");

            string path = ScreenCapture.CaptureOutlookWindowByCaptionFragment(
                Marker,
                _fixture.ScreenshotsDirectory,
                $"phase4-new-draft-display-{DateTime.Now:yyyyMMdd-HHmmss}.png");
            var file = new FileInfo(path);
            Assert.True(file.Exists && file.Length > 0, "screenshot must exist and be non-empty");
            _output.WriteLine($"screenshot saved: {path} bytes={file.Length}");
        }
        finally
        {
            // Close the window the test opened (Phase-3 helper; olDiscard - the saved
            // draft itself is untouched), then delete the draft.
            bool closed = _fixture.VerifySession.TryCloseInspectorByEntryId(outcome.EntryId, out string? closeError);
            _output.WriteLine($"inspector close requested: ok={closed} err={closeError ?? "-"}");
            CleanupDraft(Hub, outcome.EntryId);
        }

        ComInspectorInfo? stillOpen = PollForInspector(outcome.EntryId, present: false, TimeSpan.FromSeconds(10));
        Assert.True(stillOpen == null, "Inspector still open after the test closed it");
        AssertGone(outcome.EntryId, hubStoreId);
        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
        _output.WriteLine("display case: inspector closed, draft deleted, 0 marker artifacts remain");
    }

    [Fact]
    public void IdentityDrafts_BusinessAccounts_RightStore_NeverDisplayed_DeletedImmediately()
    {
        // Q-it2-3a: ONE tagged identity-verification draft per business account -
        // created, property-asserted CONTENT-FREE (S4: only booleans/ids in output),
        // deleted immediately, never displayed.
        foreach (string account in _fixture.IdentityAccounts)
        {
            ComDefaultFolderInfo drafts = RequireDefaultFolder(account, OlFolderDrafts);
            string storeId = _fixture.GetStoreId(account);
            string subject = _fixture.TaggedSubject("identity");

            DraftOutcome outcome = Service.NewDraft(
                account, account, cc: null, subject, "Identity verification draft (agent-authored). " + Marker, display: false);
            bool deleted = false;
            try
            {
                Assert.False(outcome.Displayed);
                Assert.True(outcome.AccountResolved, "identity draft must pin SendUsingAccount from the Account object");
                Assert.Equal(account, outcome.Store, ignoreCase: true);

                ComDraftInfo reopened = RequireMailInfo(outcome.EntryId, storeId);
                string? persistedAccount = reopened.SendUsingAccountSmtp ?? outcome.Account;
                bool accountMatches = string.Equals(persistedAccount, account, StringComparison.OrdinalIgnoreCase);
                bool inOwnDrafts = string.Equals(reopened.ParentFolderEntryId, drafts.EntryId, StringComparison.OrdinalIgnoreCase);
                Assert.True(accountMatches, "identity must resolve to the requested account");
                Assert.True(inOwnDrafts, "identity draft must land in that account's own Drafts folder");

                // Never displayed: no Inspector for the EntryID and no visible Outlook
                // window carrying the run marker in its caption.
                bool inspectorSeen = _fixture.VerifySession.GetOpenInspectors()
                    .Any(i => i.EntryId != null && string.Equals(i.EntryId, outcome.EntryId, StringComparison.OrdinalIgnoreCase));
                Assert.False(inspectorSeen, "identity draft must never get an Inspector");
                Assert.False(
                    ScreenCapture.AnyVisibleOutlookWindowWithCaptionFragment(Marker),
                    "no visible Outlook window may show the identity draft");

                // Content-free output (S4): booleans + account identifier only.
                _output.WriteLine(
                    $"identity[{account}]: accountResolved={outcome.AccountResolved} accountMatches={accountMatches} "
                    + $"inOwnDrafts={inOwnDrafts} signatureInjected={outcome.SignatureInjected} displayed=false");

                deleted = LiveOutlookTestMailer.DeleteItemByEntryId(account, outcome.EntryId, Marker);
                Assert.True(deleted, "identity draft must delete cleanly");
            }
            finally
            {
                if (!deleted)
                {
                    CleanupDraft(account, outcome.EntryId);
                }

                // Purge pass: remove the just-deleted draft from that store's Deleted
                // Items so nothing tagged remains anywhere in the business store.
                LiveOutlookTestMailer.DeleteTaggedArtifacts(account, Marker);
            }

            AssertGone(outcome.EntryId, storeId);
            Assert.Equal(
                0,
                LiveOutlookTestMailer.CountTaggedArtifacts(account, Marker, new[] { OlFolderDrafts, 3 }));
            _output.WriteLine($"identity[{account}]: deleted, 0 marker artifacts remain (Drafts+Deleted)");
        }
    }

    [Fact]
    public void ArtifactSweep_AllThreeAccounts_ZeroTaggedRemain()
    {
        // S3 post-suite proof (also run explicitly after the full suite): NO item
        // tagged [OutlookAI-McpTest] - from this or any earlier run - remains in
        // Drafts/Inbox/Sent/Deleted of ANY of the three accounts. Counts only (S4).
        foreach (string store in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            int count = LiveOutlookTestMailer.CountTaggedArtifacts(store, "OutlookAI-McpTest");
            _output.WriteLine($"sweep[{store}]: taggedArtifacts={count}");
            Assert.Equal(0, count);
        }
    }

    // ------------------------------------------------------------------ helpers

    private void AssertDerivedDraft(
        DraftOutcome outcome,
        string expectedKind,
        ComDraftInfo seedInfo,
        ComDefaultFolderInfo hubDrafts,
        string hubStoreId,
        string quoteToken,
        string agentText)
    {
        Assert.Equal(expectedKind, outcome.Kind);
        Assert.True(outcome.AccountResolved, expectedKind + ": account must derive from the source store");
        Assert.Equal(Hub, outcome.Store, ignoreCase: true);
        Assert.False(outcome.Displayed);

        ComDraftInfo reopened = RequireMailInfo(outcome.EntryId, hubStoreId);

        // ACCEPTANCE: the draft's ConversationIndex EXTENDS the original's (child-of).
        Assert.False(string.IsNullOrEmpty(reopened.ConversationIndex), expectedKind + ": draft must carry a ConversationIndex");
        Assert.StartsWith(seedInfo.ConversationIndex!, reopened.ConversationIndex!, StringComparison.OrdinalIgnoreCase);
        Assert.True(
            reopened.ConversationIndex!.Length > seedInfo.ConversationIndex!.Length,
            expectedKind + ": ConversationIndex must be LONGER than the original (child entry appended)");

        // Placement: the hub store's Drafts folder.
        Assert.Equal(hubDrafts.EntryId, reopened.ParentFolderEntryId, ignoreCase: true);
        string? persistedAccount = reopened.SendUsingAccountSmtp ?? outcome.Account;
        Assert.Equal(Hub, persistedAccount, ignoreCase: true);

        // Quoted history present + agent text ABOVE the quote.
        ReadOutcome read = Service.Read(outcome.EntryId, maxBodyChars: 100000);
        int agentIndex = read.Body.IndexOf(agentText, StringComparison.Ordinal);
        int quoteIndex = read.Body.IndexOf(quoteToken, StringComparison.Ordinal);
        Assert.True(agentIndex >= 0, expectedKind + ": agent text must be in the draft body");
        Assert.True(quoteIndex >= 0, expectedKind + ": quoted seed content must be in the draft body");
        Assert.True(agentIndex < quoteIndex, expectedKind + ": agent text must sit ABOVE the quoted history");

        _output.WriteLine(
            $"{expectedKind}: conversationIndexExtends=true (len {seedInfo.ConversationIndex!.Length} -> {reopened.ConversationIndex!.Length}) "
            + $"inHubDrafts=true agentAboveQuote=true signatureInjected={outcome.SignatureInjected}");
    }

    private ComMailBrief WaitForInboxArrival(string seedSubject, DateTime sentUtc)
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(180);
        while (DateTime.UtcNow < deadline)
        {
            ComSweepResult sweep = _fixture.VerifySession.SweepDefaultFoldersNewerThan(
                sentUtc.AddMinutes(-2), perFolderCap: 100, includeBodies: false, onlyStoreDisplayName: Hub);
            ComMailBrief? hit = sweep.Items.FirstOrDefault(i =>
                i.FolderKind == "inbox" && string.Equals(i.Subject, seedSubject, StringComparison.Ordinal));
            if (hit != null)
            {
                return hit;
            }

            Thread.Sleep(3000);
        }

        throw new TimeoutException("Seed mail did not arrive in the hub Inbox within 180 s (D20 round trip).");
    }

    private ComDefaultFolderInfo RequireDefaultFolder(string store, int folderId)
    {
        ComDefaultFolderInfo? info = _fixture.VerifySession.TryGetDefaultFolderInfo(store, folderId, out string? error);
        Assert.True(info != null, $"default folder {folderId} of '{store}' unavailable: {error}");
        return info!;
    }

    private ComDraftInfo RequireMailInfo(string entryId, string? storeId)
    {
        ComDraftInfo? info = _fixture.VerifySession.TryGetMailInfo(entryId, storeId, out string? error);
        Assert.True(info != null, $"mail info unavailable: {error}");
        return info!;
    }

    private void AssertGone(string entryId, string? storeId)
    {
        ComDraftInfo? info = _fixture.VerifySession.TryGetMailInfo(entryId, storeId, out _);
        Assert.True(info == null, "deleted item must no longer open by EntryID");
    }

    /// <summary>Delete-with-retries in finally (Phase-2/3 cleanup discipline).</summary>
    private void CleanupDraft(string store, string entryId)
    {
        for (int attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                LiveOutlookTestMailer.DeleteItemByEntryId(store, entryId, Marker);
                LiveOutlookTestMailer.DeleteTaggedArtifacts(store, Marker);
                return;
            }
            catch (Exception) when (attempt < 2)
            {
                Thread.Sleep(1000);
            }
        }
    }

    private ComInspectorInfo? PollForInspector(string entryId, bool present, TimeSpan timeout)
    {
        DateTime deadline = DateTime.UtcNow + timeout;
        ComInspectorInfo? last = null;
        while (DateTime.UtcNow < deadline)
        {
            last = _fixture.VerifySession.GetOpenInspectors()
                .FirstOrDefault(i => i.EntryId != null && string.Equals(i.EntryId, entryId, StringComparison.OrdinalIgnoreCase));
            if ((last != null) == present)
            {
                return last;
            }

            Thread.Sleep(500);
        }

        return last;
    }
}
