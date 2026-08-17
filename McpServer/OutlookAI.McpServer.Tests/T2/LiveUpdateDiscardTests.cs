using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Live acceptance for soak fix batch C (v3.MD D46) - HUB ONLY (S2), every artifact
/// tagged and helper-purged (S3).
/// <para>
/// <b>C1 update_draft</b> is proven against the RAW saved <c>HTMLBody</c>, because that
/// is the only place the contract is visible: the plain-text extraction collapses
/// exactly the markup that separates the draft region from the signature and the quoted
/// original. The load-bearing assertion is not "the new text is there" but "the OLD text
/// is GONE while the signature and quote are byte-present" - an appending bug would
/// satisfy a naive contains-check perfectly.
/// </para>
/// <para>
/// <b>C2 discard_draft</b> is the product's only mail-deleting tool (S1 v3), so all
/// three of its gates are proven INDEPENDENTLY: each test satisfies the other two gates
/// so the refusal it asserts is the one actually under test, never an earlier gate
/// firing first.
/// </para>
/// <para>
/// <b>C3 attachments</b> proves the security interlock end to end: a send token issued
/// for a draft is invalidated by attaching a file to it.
/// </para>
/// </summary>
[Collection("LivePhase4")]
[Trait("Category", "Live")]
public sealed class LiveUpdateDiscardTests
{
    private readonly LivePhase4Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveUpdateDiscardTests(LivePhase4Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    // ================================================================== C1: update_draft

    [Fact]
    public void UpdateDraft_ReplacesTheBodyRegion_LeavingTheSignatureIntact()
    {
        using TestSignature sig = TestSignature.Create(Marker);
        string original = "C1ORIGINALBODY" + Marker;
        string revised = "C1REVISEDBODY" + Marker;

        LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
        DraftOutcome draft = Service.NewDraft(
            Hub, Hub, cc: null, _fixture.TaggedSubject("c1-body"), original + " first version.",
            display: false, signature: sig.Name);
        try
        {
            string before = RequireHtmlBody(draft.EntryId);
            Assert.Contains(original, before, StringComparison.Ordinal);
            _output.WriteLine(
                $"C1 create diag: bodyPlacement={draft.BodyPlacement} signatureApplied={draft.SignatureApplied} "
                + $"signatureError={draft.SignatureError ?? "-"} htmlLen={before.Length}");

            // D47 PRECONDITION, and the whole reason the image survives below: the
            // signature's picture must be EMBEDDED (a cid: resource backed by a real
            // inline attachment), not left as the file:/// link Word's InsertFile
            // produces. A link renders on this machine and nowhere else, and no
            // re-render can carry it.
            string cidBefore = RequireCidImage(before, "created draft");
            Assert.DoesNotContain("src=\"file:///", before, StringComparison.OrdinalIgnoreCase);
            Assert.NotEmpty(Service.Read(draft.EntryId, maxBodyChars: 0).Attachments!);
            _output.WriteLine($"C1 create image: embedded as {cidBefore}, backed by an inline attachment");

            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "update_draft");
            UpdateDraftOutcome updated = Service.UpdateDraft(
                draft.EntryId, body: revised + " second version.", display: false);

            Assert.Equal("updated", updated.Status);
            Assert.Contains("body", updated.Changed!);
            Assert.Equal("text", updated.BodyFormat);
            Assert.Equal("wordEditor", updated.BodyPlacement);

            string after = RequireHtmlBody(updated.EntryId);

            // Diagnostics first, so a failure says WHICH way it broke: old text still
            // present with no new text = the Word edit never flushed; both present with
            // the new one first = the draft region was not cleared (prepend).
            _output.WriteLine(
                $"C1 diag: oldAt={after.IndexOf(original, StringComparison.Ordinal)} "
                + $"newAt={after.IndexOf(revised, StringComparison.Ordinal)} "
                + $"htmlLen {before.Length}->{after.Length} "
                + $"sameEntryId={string.Equals(updated.EntryId, draft.EntryId, StringComparison.OrdinalIgnoreCase)}");

            // The plain-text read must agree - this is what an agent sees next turn.
            ReadOutcome readBack = Service.Read(updated.EntryId, maxBodyChars: 2000);
            Assert.DoesNotContain(original, readBack.Body!, StringComparison.Ordinal);
            Assert.Contains(revised, readBack.Body!, StringComparison.Ordinal);

            // REPLACED, not appended - the whole point of the tool.
            Assert.DoesNotContain(original, after, StringComparison.Ordinal);
            int bodyAt = RequireIndexOf(after, revised, "revised body");

            // ...and the signature region survived, in place, below the new text.
            int anchorAt = RequireIndexOf(after, "_MailAutoSig", "signature region");
            int signatureAt = RequireIndexOf(after, "testhandtekening", "signature content");
            Assert.True(bodyAt < anchorAt, $"revised body must precede the signature region (body@{bodyAt} sig@{anchorAt})");
            Assert.True(bodyAt < signatureAt, "revised body must precede the signature content");
            AssertBodyInsideWordSection(after, bodyAt);

            // D47 - WAS AN OBSERVATION, IS NOW THE CONTRACT. The signature's embedded
            // image must come back through the re-rendered document intact: the same
            // cid: reference AND the inline attachment behind it. This used to fail,
            // because Word's InsertFile leaves signature pictures LINKED to the file on
            // disk and cannot re-serialize such a link - it emits a placeholder shape
            // instead. Asserting BOTH halves matters: an <img> whose attachment has gone
            // is a broken image, and an attachment no <img> points at is dead weight.
            Assert.Equal(cidBefore, RequireCidImage(after, "revised draft"));
            Assert.DoesNotContain("v:rect", after, StringComparison.OrdinalIgnoreCase);
            IReadOnlyList<AttachmentView> attachmentsAfter = Service.Read(updated.EntryId, maxBodyChars: 0).Attachments!;
            Assert.NotEmpty(attachmentsAfter);
            Assert.Null(updated.InlineImagesDropped);
            Assert.Null(updated.InlineImagesAdvice);
            _output.WriteLine(
                $"C1 signature-image after update: {cidBefore} still referenced, {attachmentsAfter.Count} attachment(s) intact");
            _output.WriteLine($"C1 body replace: bodyAt={bodyAt} anchorAt={anchorAt} contentAt={signatureAt}; old text absent");
        }
        finally
        {
            CleanupDraft(draft.EntryId);
        }

        AssertNoTaggedArtifactsRemain();
    }

    [Fact]
    public void UpdateDraft_ReportsAnInlineImageItCannotCarryOver_AndReapplyingTheSignatureRestoresIt()
    {
        // THE RESIDUAL LIMITATION (D47), proven rather than described. A draft composed
        // by an OLDER build carries its signature picture as a file:/// LINK. Word cannot
        // re-serialize such a link across a revision - it emits a placeholder shape - and
        // nothing in the update path can rescue it, because by the time the WordEditor is
        // available the picture is already a placeholder. So the contract is: the loss is
        // REPORTED, never silent, and the reported remedy actually works.
        using TestSignature sig = TestSignature.Create(Marker);

        LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
        DraftOutcome draft = Service.NewDraft(
            Hub, Hub, cc: null, _fixture.TaggedSubject("d47-legacy"), "D47LEGACY" + Marker,
            display: false, signature: sig.Name);
        try
        {
            string current = RequireHtmlBody(draft.EntryId);
            RequireCidImage(current, "created draft");

            // Rewrite the stored HTML into the PRE-D47 shape: the same picture, but
            // linked to the signature file instead of embedded. Goes through the tested,
            // allowlist-guarded helper (S3 tag AND run-marker double match).
            int cidAt = current.IndexOf("src=\"cid:", StringComparison.OrdinalIgnoreCase);
            int cidEnd = current.IndexOf('"', cidAt + 5);
            Assert.True(cidAt > 0 && cidEnd > cidAt, "the created draft must carry a cid: image to downgrade");
            string linkTarget = "file:///"
                + Path.GetDirectoryName(sig.FilePath)!.Replace('\\', '/') + "/"
                + Path.GetFileNameWithoutExtension(sig.FilePath) + "_files/sigimg.png";
            LiveOutlookTestMailer.SetDraftHtmlBody(
                Hub, draft.EntryId, Marker, current.Substring(0, cidAt) + "src=\"" + linkTarget + current.Substring(cidEnd));

            string legacy = RequireHtmlBody(draft.EntryId);
            Assert.Contains("src=\"file:///", legacy, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(1, CountImages(legacy));

            // (1) The revision loses it - and SAYS SO.
            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "update_draft");
            UpdateDraftOutcome lost = Service.UpdateDraft(
                draft.EntryId, body: "D47LOST" + Marker, display: false);

            string afterLoss = RequireHtmlBody(lost.EntryId);
            _output.WriteLine(
                $"D47 legacy: images {CountImages(legacy)}->{CountImages(afterLoss)} "
                + $"reported={lost.InlineImagesDropped?.ToString() ?? "null"}");
            Assert.Equal(0, CountImages(afterLoss));
            Assert.Equal(1, lost.InlineImagesDropped);
            Assert.NotNull(lost.InlineImagesAdvice);
            Assert.Contains("signature", lost.InlineImagesAdvice!, StringComparison.OrdinalIgnoreCase);

            // (2) The advertised remedy works: re-applying the signature brings the image
            // back, and back EMBEDDED - so the next revision keeps it.
            UpdateDraftOutcome restored = Service.UpdateDraft(
                lost.EntryId, body: "D47RESTORED" + Marker, signature: sig.Name, display: false);

            string afterRemedy = RequireHtmlBody(restored.EntryId);
            string cid = RequireCidImage(afterRemedy, "draft after the signature was re-applied");
            Assert.Null(restored.InlineImagesDropped);

            // ...and it now survives a further plain revision, which is the point.
            UpdateDraftOutcome again = Service.UpdateDraft(
                restored.EntryId, body: "D47AGAIN" + Marker, display: false);
            Assert.Equal(cid, RequireCidImage(RequireHtmlBody(again.EntryId), "draft after a further revision"));
            Assert.Null(again.InlineImagesDropped);
            _output.WriteLine($"D47 remedy: signature re-applied, image embedded as {cid}, survives a further revision");
        }
        finally
        {
            CleanupDraft(draft.EntryId);
        }

        AssertNoTaggedArtifactsRemain();
    }

    [Fact]
    public void UpdateDraft_OnAReply_ReplacesTheBody_AndKeepsTheQuotedOriginal()
    {
        string quoteToken = "C1QUOTESEED" + Marker;
        string original = "C1REPLYFIRST" + Marker;
        string revised = "C1REPLYSECOND" + Marker;

        string seedSubject = _fixture.TaggedSubject("c1-quote-seed");
        DateTime sentUtc = LiveOutlookTestMailer.SendSelfMail(Hub, seedSubject, quoteToken + " seed body.", null);
        try
        {
            ComMailBrief arrived = WaitForInboxArrival(seedSubject, sentUtc);

            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "reply_draft");
            DraftOutcome reply = Service.ReplyDraft(arrived.EntryId, original + " first.", display: false);

            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "update_draft");
            UpdateDraftOutcome updated = Service.UpdateDraft(
                reply.EntryId, bodyHtml: "<p><strong>" + revised + "</strong> second.</p>", display: false);

            Assert.Equal("html", updated.BodyFormat);
            Assert.Equal("wordEditor", updated.BodyPlacement);

            string html = RequireHtmlBody(updated.EntryId);
            _output.WriteLine(
                $"C1 reply diag: oldAt={html.IndexOf(original, StringComparison.Ordinal)} "
                + $"newAt={html.IndexOf(revised, StringComparison.Ordinal)} htmlLen={html.Length}");
            Assert.DoesNotContain(original, html, StringComparison.Ordinal);
            int bodyAt = RequireIndexOf(html, revised, "revised reply body");
            int quoteAt = RequireIndexOf(html, quoteToken, "quoted original");

            Assert.True(bodyAt < quoteAt, $"revised body must precede the quoted original (body@{bodyAt} quote@{quoteAt})");
            AssertBodyInsideWordSection(html, bodyAt);
            _output.WriteLine($"C1 reply update: bodyAt={bodyAt} quoteAt={quoteAt}; old body absent, quote intact");
        }
        finally
        {
            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(Hub, Marker);
        }

        AssertNoTaggedArtifactsRemain();
    }

    [Fact]
    public void UpdateDraft_ReplacesSubjectAndRecipients_KeepingThreadingAndTheOtherFields()
    {
        LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
        DraftOutcome draft = Service.NewDraft(
            Hub, Hub, cc: Hub, _fixture.TaggedSubject("c1-fields"), "Recipient replace probe " + Marker,
            display: false);
        try
        {
            string newSubject = _fixture.TaggedSubject("c1-fields-renamed");

            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "update_draft");
            UpdateDraftOutcome updated = Service.UpdateDraft(
                draft.EntryId,
                subject: newSubject,
                cc: string.Empty, // REPLACE with nothing = clear
                importance: "high",
                requestReadReceipt: true,
                display: false);

            Assert.Equal(newSubject, updated.Subject);
            Assert.Contains("subject", updated.Changed!);
            Assert.Contains("cc", updated.Changed!);
            Assert.Equal("high", updated.Importance);
            Assert.True(updated.ReadReceiptRequested);

            // REPLACE semantics: the Cc that was on the draft is gone, To untouched.
            Assert.DoesNotContain(updated.Recipients!, r => string.Equals(r.Kind, "cc", StringComparison.OrdinalIgnoreCase));
            Assert.Contains(updated.Recipients!, r => string.Equals(r.Kind, "to", StringComparison.OrdinalIgnoreCase));

            // The saved item agrees (round trip, not an echo).
            ComDraftInfo info = RequireMailInfo(updated.EntryId);
            Assert.Equal(newSubject, info.Subject);
            Assert.Equal(2, info.Importance);
            Assert.True(info.ReadReceiptRequested);
            Assert.DoesNotContain(info.Recipients, r => string.Equals(r.Kind, "cc", StringComparison.OrdinalIgnoreCase));
            _output.WriteLine($"C1 fields: subject replaced, cc cleared, recipients={info.Recipients.Count}, importance={info.Importance}");
        }
        finally
        {
            CleanupDraft(draft.EntryId);
        }

        AssertNoTaggedArtifactsRemain();
    }

    // ================================================================== C2: discard_draft

    [Fact]
    public void DiscardDraft_SoftDeletesItsOwnDraft_GoneFromDrafts_PresentInDeletedItems()
    {
        LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
        DraftOutcome draft = Service.NewDraft(
            Hub, Hub, cc: null, _fixture.TaggedSubject("c2-discard"), "Discard round trip " + Marker, display: false);
        string entryId = draft.EntryId;
        try
        {
            string draftsFolder = draft.Folder ?? "Drafts";

            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Delete, "discard_draft");
            DiscardDraftOutcome discarded = Service.DiscardDraft(entryId);

            Assert.True(discarded.Discarded);
            Assert.Equal(entryId, discarded.EntryId);
            Assert.Equal(draftsFolder, discarded.FromFolder);
            Assert.False(string.IsNullOrEmpty(discarded.ToFolder), "the Deleted Items folder must be named");

            // GONE from Drafts: the old EntryID no longer opens.
            Assert.Null(_fixture.VerifySession.TryGetMailInfo(entryId, _fixture.GetStoreId(Hub), out string? goneError));
            _output.WriteLine($"C2 old id no longer opens: {goneError ?? "-"}");

            // PRESENT in Deleted Items: the re-located copy opens and is the same mail.
            Assert.False(string.IsNullOrEmpty(discarded.NewEntryId), "the discarded draft must be re-located in Deleted Items");
            ComDraftInfo moved = RequireMailInfo(discarded.NewEntryId!);
            Assert.Equal(draft.Subject, moved.Subject);
            Assert.Equal(discarded.ToFolder, moved.ParentFolderName);

            // Soft, never permanent - and reversible, like a move (D39).
            Assert.Contains("move_mail", discarded.Advice!, StringComparison.Ordinal);
            _output.WriteLine($"C2 discard: {draftsFolder} -> {discarded.ToFolder}, newEntryId re-located, undo advertised");
        }
        finally
        {
            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(Hub, Marker);
        }

        AssertNoTaggedArtifactsRemain();
    }

    [Fact]
    public void DiscardDraft_RefusesADraftThisServerDidNotMake_EvenThoughItIsAnUnsentHubDraft()
    {
        // GATE 1 in isolation: the item satisfies BOTH other gates (unsent, in Drafts) -
        // it is refused purely because this server did not author it. Seeded through the
        // test mailer precisely so the product never learns its EntryID.
        string subject = _fixture.TaggedSubject("c2-foreign");
        string entryId = LiveOutlookTestMailer.SaveTaggedDraftWithAttachments(
            Hub, subject, "Not made by the server " + Marker, Array.Empty<string>());
        try
        {
            Assert.False(Service.DraftRegistry.Contains(entryId), "precondition: the server must not know this draft");

            DraftRefusedException ex = Assert.Throws<DraftRefusedException>(() => Service.DiscardDraft(entryId));

            Assert.Equal("not_created_by_this_server", ex.Reason);
            Assert.Contains("Delete it in Outlook instead", ex.Message, StringComparison.Ordinal);

            // And it is still there - a refusal is never a partial delete.
            Assert.Equal(subject, RequireMailInfo(entryId).Subject);
            _output.WriteLine("C2 refusal (a): unregistered draft refused, draft still present");
        }
        finally
        {
            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(Hub, Marker);
        }

        AssertNoTaggedArtifactsRemain();
    }

    [Fact]
    public void DiscardDraft_RefusesASentItem_EvenWhenTheRegistryGateIsSatisfied()
    {
        // GATE 2 in isolation: the registry gate is deliberately satisfied for the sent
        // item, so the refusal proves the UNSENT check itself fires rather than the
        // registry check masking it.
        string seedSubject = _fixture.TaggedSubject("c2-sent");
        DateTime sentUtc = LiveOutlookTestMailer.SendSelfMail(Hub, seedSubject, "Sent-item refusal probe " + Marker, null);
        string? registered = null;
        try
        {
            ComMailBrief arrived = WaitForInboxArrival(seedSubject, sentUtc);
            Service.DraftRegistry.Register(arrived.EntryId);
            registered = arrived.EntryId;

            DraftRefusedException ex = Assert.Throws<DraftRefusedException>(() => Service.DiscardDraft(arrived.EntryId));

            Assert.Equal("not_an_unsent_draft", ex.Reason);
            Assert.Contains("already been sent", ex.Message, StringComparison.Ordinal);

            // Still present, untouched.
            Assert.Equal(seedSubject, RequireMailInfo(arrived.EntryId).Subject);
            _output.WriteLine("C2 refusal (b): sent item refused with the registry gate satisfied, item still present");
        }
        finally
        {
            // The test widened the registry on purpose to isolate gate 2 - undo it so no
            // later test inherits a membership the product never granted.
            Service.DraftRegistry.Forget(registered);
            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(Hub, Marker);
        }

        AssertNoTaggedArtifactsRemain();
    }

    [Fact]
    public void DiscardDraft_RefusesAnUnsentItemOutsideDrafts_AndSoDoesUpdateDraft()
    {
        // GATE 3 in isolation: an UNSENT draft moved out of Drafts, with the registry
        // gate satisfied for the moved id, so only the folder check can refuse it.
        LiveOutlookTestMailer.DeleteTestFolders(Hub);
        LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
        DraftOutcome draft = Service.NewDraft(
            Hub, Hub, cc: null, _fixture.TaggedSubject("c2-elsewhere"), "Outside-drafts probe " + Marker, display: false);
        string? movedEntryId = null;
        try
        {
            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Move, "move_mail");
            MoveMailOutcome move = Service.MoveMail(
                new[] { draft.EntryId }, LiveOutlookTestMailer.TestFolderNamePrefix, createFolder: true);
            Assert.Null(move.Items[0].Error);
            string movedId = move.Items[0].NewEntryId!;
            movedEntryId = movedId;
            Service.DraftRegistry.Register(movedId);

            DraftRefusedException discardEx = Assert.Throws<DraftRefusedException>(() => Service.DiscardDraft(movedId));
            Assert.Equal("not_in_drafts_folder", discardEx.Reason);
            Assert.Contains("does not live in a Drafts folder", discardEx.Message, StringComparison.Ordinal);

            // update_draft applies the SAME precondition - the gate is shared, not
            // duplicated per tool.
            DraftRefusedException updateEx = Assert.Throws<DraftRefusedException>(
                () => Service.UpdateDraft(movedId, subject: _fixture.TaggedSubject("c2-should-not-happen"), display: false));
            Assert.Equal("not_in_drafts_folder", updateEx.Reason);

            // Untouched by both refusals.
            Assert.Equal(draft.Subject, RequireMailInfo(movedId).Subject);
            _output.WriteLine("C2 refusal (c): item outside Drafts refused by discard_draft AND update_draft, item untouched");
        }
        finally
        {
            Service.DraftRegistry.Forget(movedEntryId);
            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(Hub, Marker);
            LiveOutlookTestMailer.DeleteTestFolders(Hub);
        }

        AssertNoTaggedArtifactsRemain();
    }

    // ================================================================== C3: attachments

    [Fact]
    public void NewDraft_WithAttachments_ReportsWhatActuallyLandedOnTheSavedDraft()
    {
        string directory = Path.Combine(Path.GetTempPath(), "OutlookAI-McpTest-" + Marker + "-c3");
        Directory.CreateDirectory(directory);
        DraftOutcome? draft = null;
        try
        {
            string one = WriteTempFile(directory, "quote-" + Marker + ".txt", 120);
            string two = WriteTempFile(directory, "terms-" + Marker + ".txt", 640);
            string three = WriteTempFile(directory, "notes-" + Marker + ".txt", 2048);

            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
            draft = Service.NewDraft(
                Hub, Hub, cc: null, _fixture.TaggedSubject("c3-attach"), "Attachment probe " + Marker,
                display: false, attachments: new[] { one, two, three });

            Assert.Equal(3, draft.AttachmentsRequested);
            Assert.NotNull(draft.Attachments);
            Assert.Equal(3, draft.Attachments!.Count);

            string[] names = draft.Attachments.Select(a => a.FileName!).OrderBy(n => n, StringComparer.Ordinal).ToArray();
            Assert.Equal(
                new[] { Path.GetFileName(one), Path.GetFileName(three), Path.GetFileName(two) }
                    .OrderBy(n => n, StringComparer.Ordinal).ToArray(),
                names);

            // Sizes come from the SAVED item, so they must match what was on disk.
            AttachmentView first = draft.Attachments.Single(a => a.FileName == Path.GetFileName(one));
            Assert.Equal(120, first.SizeBytes);
            Assert.Equal(120 + 640 + 2048, draft.AttachmentsTotalBytes);

            // read agrees - the same list a later agent turn would see.
            ReadOutcome read = Service.Read(draft.EntryId, maxBodyChars: 0);
            Assert.Equal(3, read.Attachments!.Count);
            _output.WriteLine($"C3 attach: 3 files on the saved draft, {draft.AttachmentsTotalBytes} bytes total");
        }
        finally
        {
            if (draft != null)
            {
                CleanupDraft(draft.EntryId);
            }

            TryDeleteDirectory(directory);
        }

        AssertNoTaggedArtifactsRemain();
    }

    [Fact]
    public void UpdateDraft_AddsAndRemovesAttachments_AndSaysWhatMatchedNothing()
    {
        string directory = Path.Combine(Path.GetTempPath(), "OutlookAI-McpTest-" + Marker + "-c3b");
        Directory.CreateDirectory(directory);
        DraftOutcome? draft = null;
        try
        {
            string keep = WriteTempFile(directory, "keep-" + Marker + ".txt", 64);
            string drop = WriteTempFile(directory, "drop-" + Marker + ".txt", 128);
            string added = WriteTempFile(directory, "added-" + Marker + ".txt", 256);

            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
            draft = Service.NewDraft(
                Hub, Hub, cc: null, _fixture.TaggedSubject("c3-attach-update"), "Attachment update probe " + Marker,
                display: false, attachments: new[] { keep, drop });

            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "update_draft");
            UpdateDraftOutcome updated = Service.UpdateDraft(
                draft.EntryId,
                attachments: new[] { added },
                removeAttachments: new[] { Path.GetFileName(drop), "never-attached.pdf" },
                display: false);

            Assert.Equal(new[] { Path.GetFileName(added) }, updated.AttachmentsAdded);
            Assert.Equal(new[] { Path.GetFileName(drop) }, updated.AttachmentsRemoved);
            Assert.Equal(new[] { "never-attached.pdf" }, updated.AttachmentsNotFound);

            string[] final = updated.Attachments!.Select(a => a.FileName!).OrderBy(n => n, StringComparer.Ordinal).ToArray();
            Assert.Equal(
                new[] { Path.GetFileName(added), Path.GetFileName(keep) }.OrderBy(n => n, StringComparer.Ordinal).ToArray(),
                final);
            _output.WriteLine($"C3 update: added={Path.GetFileName(added)} removed={Path.GetFileName(drop)} final={string.Join(",", final)}");
        }
        finally
        {
            if (draft != null)
            {
                CleanupDraft(draft.EntryId);
            }

            TryDeleteDirectory(directory);
        }

        AssertNoTaggedArtifactsRemain();
    }

    [Fact]
    public void AttachingAFile_InvalidatesAPendingSendToken()
    {
        // THE MANDATORY INTERLOCK (D46/C3): a token the user already confirmed must not
        // survive a change to what the mail actually carries. Nothing is ever sent here -
        // the send call is expected to REFUSE.
        string directory = Path.Combine(Path.GetTempPath(), "OutlookAI-McpTest-" + Marker + "-c3c");
        Directory.CreateDirectory(directory);
        DraftOutcome? draft = null;
        try
        {
            string file = WriteTempFile(directory, "late-" + Marker + ".txt", 321);

            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
            draft = Service.NewDraft(
                Hub, Hub, cc: null, _fixture.TaggedSubject("c3-token"), "Token interlock probe " + Marker, display: false);

            SendOutcome issued = Service.Send(draft.EntryId);
            Assert.Equal("confirmation_required", issued.Status);
            Assert.False(issued.Sent);
            string token = issued.ConfirmToken!;

            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "update_draft");
            Service.UpdateDraft(draft.EntryId, attachments: new[] { file }, display: false);

            SendRefusedException refused = Assert.Throws<SendRefusedException>(
                () => Service.Send(draft.EntryId, token));

            Assert.Equal("draft_changed", refused.Reason);
            Assert.False(RequireMailInfo(draft.EntryId).Subject == null, "the draft is still an unsent draft");
            _output.WriteLine("C3 interlock: token issued, one file attached, send refused with draft_changed - nothing sent");
        }
        finally
        {
            if (draft != null)
            {
                CleanupDraft(draft.EntryId);
            }

            TryDeleteDirectory(directory);
        }

        AssertNoTaggedArtifactsRemain();
    }

    [Fact]
    public void SendContentHash_IsStableOnAnUntouchedDraft_ButSeesAMarkupOnlyEdit()
    {
        // THE HTML-COVERAGE FINDING, measured rather than assumed (D46/C3):
        //  (1) the hash must NOT flip on its own, or every send would spuriously refuse;
        //  (2) Outlook derives .Body from .HTMLBody LOSSILY, so a markup-only edit leaves
        //      the plain text identical - which is exactly why the HTML digest is in the
        //      hash. Both halves are asserted here, on one draft.
        LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
        DraftOutcome draft = Service.NewDraft(
            Hub, Hub, cc: null, _fixture.TaggedSubject("c3-hash"), body: null, display: false,
            bodyHtml: "<p>Please review the <a href=\"https://example.com/terms-a\">terms</a> before Friday.</p>");
        try
        {
            string storeId = _fixture.GetStoreId(Hub);
            string baseline = ComputeHash(draft.EntryId, storeId, out string? plainBefore, out string? digestBefore);
            string repeat = ComputeHash(draft.EntryId, storeId, out _, out string? digestRepeat);

            // (1) No spurious invalidation across two independent snapshots. If this ever
            // fails, the HTML digest must come OUT of the hash - every send would refuse.
            Assert.Equal(digestBefore, digestRepeat);
            Assert.Equal(baseline, repeat);
            Assert.False(string.IsNullOrEmpty(digestBefore), "the HTML digest must actually be computed, not silently null");

            string htmlBefore = RequireHtmlBody(draft.EntryId);

            // (2) Same visible words, different markup: only the link target changes.
            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "update_draft");
            Service.UpdateDraft(
                draft.EntryId,
                bodyHtml: "<p>Please review the <a href=\"https://example.com/terms-b\">terms</a> before Friday.</p>",
                display: false);

            string afterMarkupEdit = ComputeHash(draft.EntryId, storeId, out string? plainAfter, out string? digestAfter);
            string htmlAfter = RequireHtmlBody(draft.EntryId);

            string normalizedBefore = Normalize(plainBefore);
            string normalizedAfter = Normalize(plainAfter);
            _output.WriteLine(
                $"C3 hash: htmlLen {htmlBefore.Length}->{htmlAfter.Length}; "
                + $"htmlChanged={!string.Equals(htmlBefore, htmlAfter, StringComparison.Ordinal)}; "
                + $"termsA(before/after)={htmlBefore.Contains("terms-a", StringComparison.Ordinal)}/{htmlAfter.Contains("terms-a", StringComparison.Ordinal)}; "
                + $"termsB(before/after)={htmlBefore.Contains("terms-b", StringComparison.Ordinal)}/{htmlAfter.Contains("terms-b", StringComparison.Ordinal)}; "
                + $"digestChanged={!string.Equals(digestBefore, digestAfter, StringComparison.Ordinal)}; "
                + $"plainTextIdentical={string.Equals(normalizedBefore, normalizedAfter, StringComparison.Ordinal)}; "
                + $"plainLen {normalizedBefore.Length}->{normalizedAfter.Length}");

            // THE FINDING: the plain text is what a body-only hash would have seen. If it
            // is identical here, the HTML digest is the ONLY thing standing between an
            // agent-supplied link swap and a still-valid confirm token.
            Assert.NotEqual(baseline, afterMarkupEdit);
        }
        finally
        {
            CleanupDraft(draft.EntryId);
        }

        AssertNoTaggedArtifactsRemain();
    }

    // ================================================================== helpers

    private string ComputeHash(string entryId, string storeId, out string? plainBody, out string? htmlDigest)
    {
        ComSendableDraftState? state = _fixture.VerifySession.TryGetSendableDraftState(entryId, storeId, out string? error);
        Assert.True(state != null, $"sendable state unavailable: {error ?? "unknown"}");
        plainBody = state!.BodyText;
        htmlDigest = state.BodyHtmlDigest;
        return SendContentHash.Compute(
            state.Subject, state.Recipients, state.BodyText, null, state.Attachments, state.BodyHtmlDigest);
    }

    private static string Normalize(string? text)
    {
        return (text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n').Trim();
    }

    private static string WriteTempFile(string directory, string name, int bytes)
    {
        string path = Path.Combine(directory, name);
        File.WriteAllBytes(path, Enumerable.Repeat((byte)'x', bytes).ToArray());
        return path;
    }

    private static void TryDeleteDirectory(string directory)
    {
        try
        {
            Directory.Delete(directory, recursive: true);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    /// <summary>
    /// Asserts the body carries exactly one inline image and that it is EMBEDDED, and
    /// returns its cid: reference so before/after can be compared for identity (D47).
    /// </summary>
    private static string RequireCidImage(string html, string what)
    {
        Assert.Equal(1, CountImages(html));
        int at = html.IndexOf("src=\"cid:", StringComparison.OrdinalIgnoreCase);
        Assert.True(at > 0, $"the {what} must reference its inline image by cid:, not by a file:/// link");
        int end = html.IndexOf('"', at + 5);
        Assert.True(end > at, $"the {what}'s cid: reference must be a terminated attribute");
        return html.Substring(at + 5, end - at - 5);
    }

    private static int CountImages(string html)
    {
        return OutlookAI.Core.Text.HtmlBodyComposer.CountInlineImages(html);
    }

    private static int RequireIndexOf(string html, string needle, string what)
    {
        int at = html.IndexOf(needle, StringComparison.Ordinal);
        Assert.True(at >= 0, what + " must be present in the draft HTML");
        return at;
    }

    private static void AssertBodyInsideWordSection(string html, int bodyAt)
    {
        int section = html.IndexOf("WordSection1", StringComparison.OrdinalIgnoreCase);
        Assert.True(section >= 0, "Outlook's WordSection container must be present");
        Assert.True(bodyAt > section, $"body must sit INSIDE WordSection1 (section@{section} body@{bodyAt})");
    }

    private string RequireHtmlBody(string entryId)
    {
        string? html = _fixture.VerifySession.TryGetHtmlBody(entryId, _fixture.GetStoreId(Hub), out string? error);
        Assert.True(!string.IsNullOrEmpty(html), $"draft HTML unavailable: {error ?? "empty"}");
        return html!;
    }

    private ComDraftInfo RequireMailInfo(string entryId)
    {
        ComDraftInfo? info = _fixture.VerifySession.TryGetMailInfo(entryId, _fixture.GetStoreId(Hub), out string? error);
        Assert.True(info != null, $"item unavailable: {error ?? "unknown"}");
        return info!;
    }

    private ComMailBrief WaitForInboxArrival(string subject, DateTime sentUtc)
    {
        return LiveInboxArrival.WaitFor(_fixture.VerifySession, Hub, subject, sentUtc);
    }

    private void CleanupDraft(string entryId)
    {
        for (int attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                LiveOutlookTestMailer.DeleteItemByEntryId(Hub, entryId, Marker);
                LiveOutlookTestMailer.DeleteTaggedArtifacts(Hub, Marker);
                return;
            }
            catch (Exception) when (attempt < 2)
            {
                Thread.Sleep(1000);
            }
        }
    }

    private void AssertNoTaggedArtifactsRemain()
    {
        int remaining = LiveOutlookTestMailer.CountTaggedArtifactsAfterPurgingStragglers(
            Hub, Marker, folderIds: null, out int stragglers);
        if (stragglers > 0)
        {
            _output.WriteLine($"cleanup[{Hub}]: {stragglers} late-materialized artifact(s) purged (documented lag)");
        }

        Assert.Equal(0, remaining);
    }
}
