using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Soak fix D37 (signature steering) live acceptance on the test hub (S2): ONE
/// temporary test signature (prefix "OutlookAI-McpTest-", with an image resource, per
/// the granted live-test scope) is created in the real Signatures folder and DELETED
/// in finally - the S3 cleanup discipline extended to the signature file set, with a
/// zero-leftover verification. Coverage: list_signatures sees it (name + excerpt),
/// the signature parameter applies it on new and reply drafts of the hub account (no
/// default signature there - the INSERT branch: above the quote for replies, document
/// end for new mail), a second application through the test-support surface proves
/// the REPLACE branch (_MailAutoSig exists -> delete + reinsert, no duplication), and
/// threading/quoted content/body-above-signature stay intact throughout. All draft
/// artifacts tagged + deleted (S3, stable zero).
/// </summary>
[Collection("LivePhase4")]
[Trait("Category", "Live")]
public sealed class LiveSignatureTests
{
    private readonly LivePhase4Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveSignatureTests(LivePhase4Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    [Fact]
    public void ListSignatures_SeesTestSignature_WithExcerpt_AndAccountRows()
    {
        using TestSignature sig = TestSignature.Create(Marker);

        SignaturesOutcome outcome = Service.ListSignatures();

        SignatureView listed = Assert.Single(outcome.Signatures, s => s.Name == sig.Name);
        Assert.NotNull(listed.Excerpt);
        Assert.Contains("testhandtekening", listed.Excerpt, StringComparison.OrdinalIgnoreCase);

        // Account rows: shape-only (content-free, S4) - every row is a mail account.
        // On this machine the registry carried assignments at implementation time
        // (2026-07-24; superseding the Phase-4 "not present" finding), but the test
        // stays tolerant: assignments may be absent (note explains) without failing.
        if (outcome.Accounts != null)
        {
            Assert.All(outcome.Accounts, a => Assert.Contains("@", a.Account, StringComparison.Ordinal));
            _output.WriteLine($"accounts={outcome.Accounts.Count} withNewAssignment={outcome.Accounts.Count(a => a.NewMessage != null)} "
                + $"noteSet={outcome.Note != null}");
        }
        else
        {
            Assert.NotNull(outcome.Note);
            _output.WriteLine("accounts=unreadable note=" + outcome.Note);
        }

        _output.WriteLine($"signatures={outcome.Signatures.Count} testSigListed=true excerptOk=true");
    }

    [Fact]
    public void ReplyDraft_WithSignatureOverride_InsertsAboveQuote_ThenReplaceBranchReapplies()
    {
        using TestSignature sig = TestSignature.Create(Marker);
        string hubStoreId = _fixture.GetStoreId(Hub);
        string quoteToken = "QSIG" + Marker;
        string seedSubject = _fixture.TaggedSubject("sig-seed");
        string seedBody = "Seed for the signature override test.\r\nQuote token: " + quoteToken + "\r\nEnd of seed.";
        string agentText = "Signature-steered reply " + Marker;

        DateTime sentUtc = LiveOutlookTestMailer.SendSelfMail(Hub, seedSubject, seedBody, attachmentPath: null);
        ComMailBrief seed = WaitForInboxArrival(seedSubject, sentUtc);
        ComDraftInfo seedInfo = RequireMailInfo(seed.EntryId, seed.StoreId ?? hubStoreId);

        string? draftId = null;
        try
        {
            DraftOutcome reply = Service.ReplyDraft(seed.EntryId, agentText, replyAll: false, display: false, signature: sig.Name);
            draftId = reply.EntryId;

            Assert.Equal(sig.Name, reply.Signature);
            Assert.True(reply.SignatureApplied, $"override must apply (error: {reply.SignatureError ?? "-"})");
            Assert.Null(reply.SignatureError);

            // Threading preserved: the reply's ConversationIndex EXTENDS the seed's.
            ComDraftInfo draftInfo = RequireMailInfo(reply.EntryId, hubStoreId);
            Assert.NotNull(draftInfo.ConversationIndex);
            Assert.StartsWith(seedInfo.ConversationIndex!, draftInfo.ConversationIndex!, StringComparison.OrdinalIgnoreCase);

            // Body contract: agent text ABOVE the signature ABOVE the quoted seed.
            ReadOutcome read = Service.Read(reply.EntryId);
            int agentAt = read.Body.IndexOf(agentText, StringComparison.Ordinal);
            int sigAt = read.Body.IndexOf(sig.BodyMarker, StringComparison.Ordinal);
            int quoteAt = read.Body.IndexOf(quoteToken, StringComparison.Ordinal);
            Assert.True(agentAt >= 0, "agent text must be present");
            Assert.True(sigAt > agentAt, $"signature must follow the agent text (agent@{agentAt} sig@{sigAt})");
            Assert.True(quoteAt > sigAt, $"quoted seed must follow the signature (sig@{sigAt} quote@{quoteAt})");
            _output.WriteLine($"insert branch: agent@{agentAt} sig@{sigAt} quote@{quoteAt} (ordered ok)");

            // REPLACE branch: the draft now carries a _MailAutoSig bookmark from the
            // first application - applying again must find it, delete the old region
            // and reinsert (no duplication), leaving agent text + quote intact.
            bool reapplied = _fixture.VerifySession.TryApplySignatureOverrideToDraft(
                reply.EntryId, hubStoreId, sig.FilePath, out string? reapplyError);
            Assert.True(reapplied, $"replace-branch application failed: {reapplyError ?? "-"}");

            ReadOutcome reread = Service.Read(reply.EntryId);
            int firstSig = reread.Body.IndexOf(sig.BodyMarker, StringComparison.Ordinal);
            int lastSig = reread.Body.LastIndexOf(sig.BodyMarker, StringComparison.Ordinal);
            Assert.True(firstSig >= 0, "signature must still be present after re-application");
            Assert.Equal(firstSig, lastSig); // exactly once - the old region was deleted
            int agentAt2 = reread.Body.IndexOf(agentText, StringComparison.Ordinal);
            int quoteAt2 = reread.Body.IndexOf(quoteToken, StringComparison.Ordinal);
            Assert.True(agentAt2 >= 0 && agentAt2 < firstSig, "agent text must stay above the signature");
            Assert.True(quoteAt2 > firstSig, "quote must stay below the signature");
            _output.WriteLine($"replace branch: reapplied once (sig@{firstSig}, no duplication), agent+quote intact");
        }
        finally
        {
            if (draftId != null)
            {
                TryDeleteArtifact(draftId);
            }

            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(Hub, Marker);
        }

        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
        _output.WriteLine("artifacts deleted, 0 marker artifacts remain in hub");
    }

    [Fact]
    public void NewDraft_WithSignatureOverride_AppliedOnAccountWithoutDefault()
    {
        using TestSignature sig = TestSignature.Create(Marker);
        string agentText = "New mail with steered signature " + Marker + "\r\nSecond line.";
        string subject = _fixture.TaggedSubject("sig-new");

        DraftOutcome outcome = Service.NewDraft(
            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft"), Hub, cc: null, subject, agentText, display: false, signature: sig.Name);
        try
        {
            Assert.Equal(sig.Name, outcome.Signature);
            Assert.True(outcome.SignatureApplied, $"override must apply (error: {outcome.SignatureError ?? "-"})");

            ReadOutcome read = Service.Read(outcome.EntryId);
            int agentAt = read.Body.IndexOf("New mail with steered signature", StringComparison.Ordinal);
            int sigAt = read.Body.IndexOf(sig.BodyMarker, StringComparison.Ordinal);
            Assert.True(agentAt >= 0, "agent text must be present");
            Assert.True(sigAt > agentAt, $"signature must follow the agent text (agent@{agentAt} sig@{sigAt})");

            // The image resource of the signature: recorded, not hard-asserted (the
            // embed shape may vary by Outlook build - content-free observability).
            _output.WriteLine($"new-draft override: agent@{agentAt} sig@{sigAt} attachments={read.Attachments.Count} "
                + $"bodyTotal={read.BodyTotalChars}");
        }
        finally
        {
            TryDeleteArtifact(outcome.EntryId);
            LiveOutlookTestMailer.DeleteTaggedArtifacts(Hub, Marker);
        }

        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
    }

    [Fact]
    public void TestSignatureLifecycle_LeavesZeroLeftovers()
    {
        string directory = SignatureCatalog.DefaultSignatureDirectory;
        using (TestSignature sig = TestSignature.Create(Marker))
        {
            Assert.True(File.Exists(sig.FilePath));
        }

        // The S3 discipline extended to signature files: nothing with the test prefix
        // may survive, files or resource directories alike.
        Assert.Empty(Directory.GetFileSystemEntries(directory, SignatureCatalog.TestSignaturePrefix + "*"));
    }

    // ------------------------------------------------------------------ helpers

    private ComMailBrief WaitForInboxArrival(string seedSubject, DateTime sentUtc)
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(180);
        while (DateTime.UtcNow < deadline)
        {
            ComSweepResult sweep = _fixture.VerifySession.SweepFoldersNewerThan(
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

    private ComDraftInfo RequireMailInfo(string entryId, string? storeId)
    {
        ComDraftInfo? info = _fixture.VerifySession.TryGetMailInfo(entryId, storeId, out string? error);
        Assert.True(info != null, $"mail info unavailable: {error}");
        return info!;
    }

    private void TryDeleteArtifact(string entryId)
    {
        for (int attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                LiveOutlookTestMailer.DeleteItemByEntryId(Hub, entryId, Marker);
                return;
            }
            catch (Exception) when (attempt < 2)
            {
                Thread.Sleep(1000);
            }
        }
    }
}
