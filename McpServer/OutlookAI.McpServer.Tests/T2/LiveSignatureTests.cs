using System.Globalization;
using System.Text.RegularExpressions;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Soak fix D37 (signature steering) and Q147/Q157/Q159 live acceptance on the test hub (S2):
/// temporary test signatures (prefix "OutlookAI-McpTest-", with an image resource, per the granted
/// live-test scope) are created in the real Signatures folder and DELETED in finally - the S3
/// cleanup discipline extended to the signature file set, with a zero-leftover verification.
/// Coverage: list_signatures sees one (name, excerpt, formats, picture) and reads it in full; the
/// signature parameter applies one on new and reply drafts of the hub account (no default there -
/// the INSERT branch: above the quote for replies, document end for new mail); update_draft then
/// SWAPS it for a second, different signature (the REPLACE branch - the old region gone, the new one
/// once) and finally REMOVES it with "none"; threading, quoted content and body-above-signature stay
/// intact throughout. All draft artifacts tagged + deleted (S3, stable zero).
/// </summary>
[Collection(LiveCollections.Phase4)]
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
    [Trait("Requires", "OutlookInstance")]
    [Trait("Requires", "LocalSignatures")]
    public void ListSignatures_SeesTestSignature_InEveryFormat_AndEveryProfileWithItsAccounts()
    {
        using TestSignature sig = TestSignature.Create(Marker);

        SignaturesOutcome outcome = Service.ListSignatures();

        SignatureView listed = Assert.Single(outcome.Signatures, s => s.Name == sig.Name);
        Assert.NotNull(listed.Excerpt);
        Assert.Contains("testhandtekening", listed.Excerpt, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(new[] { "html", "text" }, listed.Formats);
        Assert.Equal(1, listed.Pictures);
        Assert.Equal(SignatureCatalog.DefaultSignatureDirectory, outcome.Folder.Path, ignoreCase: true);

        // Every profile is named (Q155 (c)). Shape-only (content-free, S4): every account row is a mail
        // account. TOLERANT IS NOT THE SAME AS VACUOUS: an EMPTY account list satisfies Assert.All with
        // no element ever examined, so it is refused on a Production profile (this machine HAS mail
        // accounts) and announced on a Portable one.
        Assert.NotEmpty(outcome.Profiles);
        IReadOnlyList<SignatureAccountView> rows = LivePopulationCoverage.Require(
            _fixture.Settings,
            outcome.Profiles.SelectMany(p => p.Accounts).ToList(),
            "a mail account row in a profile's signature registry",
            "the per-account signature-assignment shape check",
            "To exercise it, run on a profile with at least one mail account; a machine with none should leave this "
                + "test out of the filter, which is what Requires=OutlookInstance is for.",
            _output.WriteLine);
        Assert.All(rows, a => Assert.Contains("@", a.Account, StringComparison.Ordinal));

        // Outlook runs on the guest, so the profile it has open is named, and it is one of the list.
        if (outcome.RunningProfile != null)
        {
            Assert.Single(outcome.Profiles, p => p.Running == true && p.Name == outcome.RunningProfile);
        }

        _output.WriteLine($"signatures={outcome.Signatures.Count} profiles={outcome.Profiles.Count} running={outcome.RunningProfile ?? "-"} "
            + $"default={outcome.DefaultProfile ?? "-"} accounts={rows.Count} folder={outcome.Folder.Source} roaming={outcome.Roaming.State}");

        // Q159: read in full - every rendition and its picture, by path.
        SignatureDetailView detail = Service.ListSignatures(sig.Name).Signature!;
        Assert.Contains(sig.BodyMarker, detail.Html, StringComparison.Ordinal);
        Assert.Contains(sig.BodyMarker, detail.Text, StringComparison.Ordinal);
        SignaturePictureView picture = Assert.Single(detail.Pictures!);
        Assert.Equal("sigimg.png", picture.File);
        Assert.True(picture.ReferencedByHtml);
        Assert.True(File.Exists(picture.Path));
        Assert.Null(detail.MissingPictures);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "Transport")]
    [Trait("Requires", "LocalSignatures")]
    public void ReplyDraft_WithSignature_InsertsAboveQuote_ThenUpdateSwapsItForAnother_ThenNoneRemovesIt()
    {
        using TestSignature sig = TestSignature.Create(Marker);
        using TestSignature second = TestSignature.Create(Marker, "B");
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
            Assert.Equal("requested", reply.SignatureSource);
            Assert.NotNull(reply.SignatureProfile);

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

            // REPLACE branch, through the product's own path: the draft carries a _MailAutoSig
            // region from the first signature, and update_draft swaps a DIFFERENT one in. Applying the
            // same signature again - what this test used to do, through a test-only path whose commit
            // shape keeps nothing - would pass whether or not anything happened.
            UpdateDraftOutcome swapped = Service.UpdateDraft(reply.EntryId, signature: second.Name, display: false);
            draftId = swapped.EntryId;
            Assert.Equal(second.Name, swapped.Signature);
            Assert.True(swapped.SignatureApplied, "the swap must be applied");
            Assert.Contains("signature", swapped.Changed!);

            ReadOutcome reread = Service.Read(swapped.EntryId);
            int newSig = reread.Body.IndexOf(second.BodyMarker, StringComparison.Ordinal);
            Assert.True(newSig >= 0, "the second signature must be in the draft");
            Assert.Equal(newSig, reread.Body.LastIndexOf(second.BodyMarker, StringComparison.Ordinal));
            Assert.DoesNotContain(sig.BodyMarker, reread.Body, StringComparison.Ordinal);
            int agentAt2 = reread.Body.IndexOf(agentText, StringComparison.Ordinal);
            int quoteAt2 = reread.Body.IndexOf(quoteToken, StringComparison.Ordinal);
            Assert.True(agentAt2 >= 0 && agentAt2 < newSig, "agent text must stay above the signature");
            Assert.True(quoteAt2 > newSig, "quote must stay below the signature");
            _output.WriteLine($"replace branch: the first signature gone, the second once (sig@{newSig}), agent+quote intact");

            // "none" (Q157 (a)): the signature region goes; the text and the quote stay.
            UpdateDraftOutcome removed = Service.UpdateDraft(swapped.EntryId, signature: "none", display: false);
            draftId = removed.EntryId;
            Assert.True(removed.SignatureRemoved, "signature=\"none\" must remove the draft's signature region");
            ReadOutcome after = Service.Read(removed.EntryId);
            Assert.DoesNotContain(second.BodyMarker, after.Body, StringComparison.Ordinal);
            Assert.DoesNotContain(sig.BodyMarker, after.Body, StringComparison.Ordinal);
            Assert.Contains(agentText, after.Body, StringComparison.Ordinal);
            Assert.Contains(quoteToken, after.Body, StringComparison.Ordinal);
            _output.WriteLine("none: signature region removed, agent text and quote intact");
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
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "LocalSignatures")]
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
            Assert.Equal("requested", outcome.SignatureSource);
            Assert.False(outcome.SignatureInjected, "a named signature is not a default");

            ReadOutcome read = Service.Read(outcome.EntryId);
            int agentAt = read.Body.IndexOf("New mail with steered signature", StringComparison.Ordinal);
            int sigAt = read.Body.IndexOf(sig.BodyMarker, StringComparison.Ordinal);
            Assert.True(agentAt >= 0, "agent text must be present");
            Assert.True(sigAt > agentAt, $"signature must follow the agent text (agent@{agentAt} sig@{sigAt})");

            // D47 - the signature's image resource is now a CONTRACT on the create path,
            // not an observation. Word's InsertFile leaves such a picture LINKED to the
            // file on disk (which renders on this machine only, and which no re-render
            // can carry), so the composition embeds it: the saved draft must carry a real
            // inline attachment, and its HTML must reference it by cid: rather than by a
            // file:/// path. Getting this right on CREATE is what makes update_draft
            // lossless - it starts from an embedded picture.
            Assert.True(read.Attachments.Count > 0, "the signature's image must be embedded as an inline attachment");
            string html = _fixture.VerifySession.TryGetHtmlBody(outcome.EntryId, _fixture.GetStoreId(Hub), out string? htmlError)
                ?? throw new InvalidOperationException("draft HTML unavailable: " + (htmlError ?? "empty"));
            Assert.Contains("src=\"cid:", html, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("src=\"file:///", html, StringComparison.OrdinalIgnoreCase);

            // Soak fix 21 - THE ZERO-BYTE PIN. A field report filed this as data loss:
            // the draft "picked up a zero-byte image001.png from the signature template".
            // The bytes were always there; the DRAFT TOOL'S OWN ECHO was the lie, because
            // the snapshot taken inside the composing call reads Size = 0 for an
            // attachment Outlook materialized during that composition. An agent that
            // believes that echo concludes the logo is broken, so the echo is now a
            // contract: non-zero bytes, and a cid: whose local part names the attachment.
            AssertInlineImageEchoedWithRealBytes(outcome.Attachments, outcome.AttachmentsTotalBytes, html);

            _output.WriteLine($"new-draft override: agent@{agentAt} sig@{sigAt} attachments={read.Attachments.Count} "
                + $"bodyTotal={read.BodyTotalChars}; signature image embedded (cid:, no file:/// link), "
                + $"draft echo bytes={outcome.AttachmentsTotalBytes}; profile={outcome.SignatureProfile ?? "-"}");
        }
        finally
        {
            TryDeleteArtifact(outcome.EntryId);
            LiveOutlookTestMailer.DeleteTaggedArtifacts(Hub, Marker);
        }

        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    public void NewDraft_SignatureNone_HasNoSignature_AndTheResultSaysWhy()
    {
        string subject = _fixture.TaggedSubject("sig-none");
        DraftOutcome outcome = Service.NewDraft(
            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft"), Hub, cc: null, subject,
            "No signature wanted " + Marker, display: false, signature: "none");
        try
        {
            Assert.Null(outcome.Signature);
            Assert.Equal("none", outcome.SignatureSource);
            Assert.Contains("\"none\"", outcome.SignatureReason, StringComparison.Ordinal);
            Assert.False(outcome.SignatureInjected);
            Assert.DoesNotContain("_MailAutoSig", RequireHtml(outcome.EntryId), StringComparison.Ordinal);
        }
        finally
        {
            TryDeleteArtifact(outcome.EntryId);
            LiveOutlookTestMailer.DeleteTaggedArtifacts(Hub, Marker);
        }

        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
    }

    [Fact]
    [Trait("Requires", "OutlookInstance")]
    [Trait("Requires", "LocalSignatures")]
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

    /// <summary>
    /// Soak fix 21: the draft tool's OWN attachment echo must describe the inline image
    /// the recipient will get - a real, non-zero attachment named by the cid the HTML
    /// references. Shared with <c>LiveDraftOptionsTests</c>, which runs the same contract
    /// against the real business-account signatures.
    /// </summary>
    internal static void AssertInlineImageEchoedWithRealBytes(
        IReadOnlyList<AttachmentView>? echoed,
        long? totalBytes,
        string html)
    {
        string cidName = RequireCidLocalName(html);
        Assert.True(echoed != null && echoed.Count > 0,
            "the draft result must ECHO the signature's inline attachment (a caller that sees none concludes the image is missing)");
        AttachmentView? match = echoed!.FirstOrDefault(a =>
            string.Equals(a.FileName, cidName, StringComparison.OrdinalIgnoreCase));
        Assert.True(match != null,
            $"the echoed attachments must include the one the HTML references by cid ('{cidName}'); echoed: "
            + string.Join(", ", echoed!.Select(a => a.FileName ?? "?")));
        Assert.True(match!.SizeBytes is > 0,
            $"the echoed inline image must carry REAL bytes, not zero (sizeBytes={match.SizeBytes?.ToString(CultureInfo.InvariantCulture) ?? "null"})");
        Assert.True(totalBytes is > 0, "attachmentsTotalBytes must be positive when an inline image is attached");
    }

    /// <summary>
    /// The file name a <c>src="cid:name@domain"</c> reference points at - the local part of
    /// the content id, which Outlook always sets to the attachment's file name.
    /// </summary>
    internal static string RequireCidLocalName(string html)
    {
        Match m = Regex.Match(html, "src=\"cid:([^\"@]+)", RegexOptions.IgnoreCase);
        Assert.True(m.Success, "the composed HTML must reference the signature image by cid:");
        return m.Groups[1].Value;
    }

    private string RequireHtml(string entryId)
    {
        return _fixture.VerifySession.TryGetHtmlBody(entryId, _fixture.GetStoreId(Hub), out string? error)
            ?? throw new InvalidOperationException("draft HTML unavailable: " + (error ?? "empty"));
    }

    private ComMailBrief WaitForInboxArrival(string seedSubject, DateTime sentUtc)
    {
        return LiveInboxArrival.WaitFor(_fixture.VerifySession, Hub, seedSubject, sentUtc);
    }

    private ComDraftInfo RequireMailInfo(string entryId, string? storeId) =>
        LiveDraftChecks.RequireMailInfo(_fixture.VerifySession, entryId, storeId);

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
