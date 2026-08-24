using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Soak-fix batch A live acceptance.
/// <para>
/// <b>A1 signature placement</b> - the per-account matrix the field report asked for.
/// The defect: the agent body ended up INSIDE the recreated <c>_MailAutoSig</c> bookmark
/// (the saved HTML opened with <c>&lt;a name="_MailAutoSig"&gt;</c> wrapped around the
/// agent text), because the body was written at Range(0,0) while the signature bookmark
/// started at 0 - Word absorbs text inserted at a bookmark's start. The contract pinned
/// here: the agent body comes FIRST, the account's own HTML signature follows INTACT,
/// and the signature anchor opens AFTER the body. Run on the hub (no configured
/// signature - body only, no bogus injection) and on both business accounts under the
/// standing identity-draft grant (Q-it2-3a): one tagged, never-displayed draft each,
/// deleted in-test, assertions content-free (S4 - only positions, booleans and
/// agent-authored markers).
/// </para>
/// <para>
/// <b>A2/A3/A4</b> - cc/bcc APPEND to Outlook's own recipient list with unresolvable
/// addresses reported instead of dropped, a subject override that keeps the draft
/// threaded (ConversationIndex still extends the source and the source conversation
/// topic is carried over), and the importance / read-receipt round trip.
/// </para>
/// All artifacts carry tag + run marker and are deleted through the tested helpers (S3).
/// </summary>
[Collection(LiveCollections.Phase4)]
[Trait("Category", "Live")]
public sealed class LiveDraftOptionsTests
{
    private readonly LivePhase4Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveDraftOptionsTests(LivePhase4Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "MultipleStores")]
    public void NewDraft_Hub_NoConfiguredSignature_BodyOnly_NoBogusInjection()
    {
        string bodyMarker = "A1HUB" + Marker;
        DraftOutcome outcome = CreateNewDraft(Hub, "a1-hub-default", bodyMarker, signature: null);
        try
        {
            string html = RequireHtmlBody(outcome.EntryId, Hub);
            int bodyAt = RequireIndexOf(html, bodyMarker, "agent body");

            // No configured signature on the hub: nothing may be injected and no
            // signature region may be invented.
            Assert.False(outcome.SignatureInjected, "the hub has no configured signature - nothing may be injected");
            Assert.DoesNotContain("_MailAutoSig", html, StringComparison.Ordinal);

            // The body is real content of the message, not spliced in ahead of Outlook's
            // own document container.
            AssertBodyInsideWordSection(html, bodyAt);
            _output.WriteLine($"A1[hub/default]: injected=false bodyAt={bodyAt} noSignatureRegion=true");
        }
        finally
        {
            CleanupDraft(Hub, outcome.EntryId);
        }

        AssertNoTaggedArtifactsRemain(Hub, folderIds: null);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "MultipleStores")]
    public void NewDraft_Hub_SignatureOverride_BodyAboveTheSignature_OutsideTheSignatureBookmark()
    {
        using TestSignature sig = TestSignature.Create(Marker);
        string bodyMarker = "A1SIG" + Marker;
        DraftOutcome outcome = CreateNewDraft(Hub, "a1-hub-override", bodyMarker, sig.Name);
        try
        {
            Assert.Equal(sig.Name, outcome.Signature);
            Assert.True(outcome.SignatureApplied, $"override must apply (error: {outcome.SignatureError ?? "-"})");

            string html = RequireHtmlBody(outcome.EntryId, Hub);

            // Word's spell/grammar spans split the signature's sentence in the saved
            // HTML, so the needle is a contiguous fragment of it, not the whole marker.
            AssertBodyAboveIntactSignature(html, bodyMarker, "testhandtekening", "hub/override");
        }
        finally
        {
            CleanupDraft(Hub, outcome.EntryId);
        }

        AssertNoTaggedArtifactsRemain(Hub, folderIds: null);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "MultipleStores")]
    public void NewDraft_BusinessAccounts_BodyAboveTheirOwnIntactHtmlSignature()
    {
        // Q-it2-3a identity grant: ONE tagged, never-displayed draft per business
        // account, deleted immediately; output stays content-free (S4).
        //
        // The account list announces itself: a machine whose configured primaries are all
        // declared BYSTANDERS has none, and this loop would otherwise run zero times and
        // report the A1 contract as verified (see IdentityDraftCoverage).
        foreach (string account in _fixture.IdentityAccounts(
            _output.WriteLine, "the A1 signature-placement matrix on the business accounts"))
        {
            string bodyMarker = "A1BIZ" + Marker;
            DraftOutcome outcome = CreateNewDraft(account, "a1-biz-default", bodyMarker, signature: null);
            try
            {
                Assert.False(outcome.Displayed);
                Assert.True(
                    outcome.SignatureInjected,
                    "the business accounts have a configured signature - Outlook must inject it natively");

                string html = RequireHtmlBody(outcome.EntryId, account);
                int bodyAt = RequireIndexOf(html, bodyMarker, "agent body");
                int anchorAt = RequireIndexOf(html, "_MailAutoSig", "signature region");

                // THE A1 CONTRACT: the signature region opens AFTER the agent body - the
                // body is never inside it - and the injected signature is real HTML, not
                // a flattened text blob.
                Assert.True(bodyAt < anchorAt, $"agent body must precede the signature region (body@{bodyAt} sig@{anchorAt})");
                AssertBodyInsideWordSection(html, bodyAt);
                Assert.True(
                    html.IndexOf("class=MsoNormal", StringComparison.OrdinalIgnoreCase) >= 0,
                    "the injected signature must keep Outlook's own HTML markup");

                bool inspectorSeen = _fixture.VerifySession.GetOpenInspectors()
                    .Any(i => i.EntryId != null && string.Equals(i.EntryId, outcome.EntryId, StringComparison.OrdinalIgnoreCase));
                Assert.False(inspectorSeen, "identity draft must never get an Inspector");

                // Soak fix 21 - the same zero-byte contract as LiveSignatureTests, but on
                // the REAL account signatures, which is where it was reported: a signature
                // with a company logo must be echoed by the draft tool as an attachment
                // with REAL bytes. Conditional because not every configured signature has
                // an image; when one does, the echo is a contract, never an observation.
                bool hasInlineImage = html.IndexOf("src=\"cid:", StringComparison.OrdinalIgnoreCase) >= 0;
                if (hasInlineImage)
                {
                    LiveSignatureTests.AssertInlineImageEchoedWithRealBytes(
                        outcome.Attachments, outcome.AttachmentsTotalBytes, html);
                }

                _output.WriteLine($"A1[{account}/default]: injected=true bodyBeforeSignatureRegion=true displayed=false "
                    + $"inlineImage={hasInlineImage} echoedBytes={outcome.AttachmentsTotalBytes?.ToString() ?? "-"}");
            }
            finally
            {
                CleanupDraft(account, outcome.EntryId);
            }

            AssertNoTaggedArtifactsRemain(account, new[] { 16, 3 });
        }
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "MultipleStores")]
    [Trait("Requires", "Transport")]
    public void DerivedDrafts_CcBccAppend_SubjectOverrideKeepsThreading_ImportanceAndReceiptRoundTrip()
    {
        string hubStoreId = _fixture.GetStoreId(Hub);
        string seedSubject = _fixture.TaggedSubject("optseed");
        DateTime sentUtc = LiveOutlookTestMailer.SendSelfMail(
            Hub, seedSubject, "Seed for the batch-A option matrix. " + Marker, attachmentPath: null);
        ComMailBrief seed = WaitForInboxArrival(seedSubject, sentUtc);
        ComDraftInfo seedInfo = RequireMailInfo(seed.EntryId, seed.StoreId ?? hubStoreId);
        Assert.False(string.IsNullOrEmpty(seedInfo.ConversationIndex), "seed must carry a ConversationIndex");

        string ccAddress = "cc-" + Marker + "@example.invalid";
        string bccAddress = "bcc-" + Marker + "@example.invalid";
        string unresolvable = "not a valid address " + Marker;
        string overriddenSubject = _fixture.TaggedSubject("renamed");
        var draftIds = new List<string>();
        try
        {
            // --- A2: reply-all keeps its OWN recipients; cc/bcc are appended; the
            //         unresolvable address is reported, not dropped.
            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "replyall_draft");
            DraftOutcome replyAll = Service.ReplyDraft(
                seed.EntryId, "Batch-A options " + Marker, replyAll: true, display: false, signature: null,
                cc: ccAddress + "; " + unresolvable, bcc: bccAddress,
                subject: null, importance: "high", requestReadReceipt: true);
            draftIds.Add(replyAll.EntryId);

            ComDraftInfo info = RequireMailInfo(replyAll.EntryId, hubStoreId);
            Assert.Contains(info.Recipients, r => r.Kind == "to");
            Assert.Contains(info.Recipients, r => r.Kind == "cc"
                && r.Address != null && r.Address.IndexOf(ccAddress, StringComparison.OrdinalIgnoreCase) >= 0);
            Assert.Contains(info.Recipients, r => r.Kind == "bcc"
                && r.Address != null && r.Address.IndexOf(bccAddress, StringComparison.OrdinalIgnoreCase) >= 0);

            Assert.NotNull(replyAll.UnresolvedRecipients);
            Assert.Contains(unresolvable, replyAll.UnresolvedRecipients!);
            Assert.DoesNotContain(ccAddress, replyAll.UnresolvedRecipients!);

            // --- A4: importance + read receipt round-trip through the saved item.
            Assert.Equal("high", replyAll.Importance);
            Assert.True(replyAll.ReadReceiptRequested);
            Assert.Equal(2, info.Importance);
            Assert.True(info.ReadReceiptRequested);
            _output.WriteLine(
                $"A2/A4: recipientKinds={string.Join("|", info.Recipients.Select(r => r.Kind))} "
                + $"unresolved={replyAll.UnresolvedRecipients!.Count} importance={info.Importance} receipt={info.ReadReceiptRequested}");

            // --- A3 baseline: without an override the derived subject stays RE: ... and
            //     the draft threads with the seed.
            DraftOutcome plain = Service.ReplyDraft(seed.EntryId, "Plain " + Marker, replyAll: false, display: false);
            draftIds.Add(plain.EntryId);
            ComDraftInfo plainInfo = RequireMailInfo(plain.EntryId, hubStoreId);
            Assert.Null(plain.ConversationTopicPreserved);
            Assert.StartsWith(seedInfo.ConversationIndex!, plainInfo.ConversationIndex!, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(seedInfo.ConversationTopic, plainInfo.ConversationTopic);

            // --- A3: the subject override replaces the subject AND keeps the thread.
            //     Outlook REGENERATES the conversation index when Subject is assigned;
            //     the shipped fix restores the Reply()-produced child index and the
            //     source topic through the PropertyAccessor.
            DraftOutcome renamed = Service.ReplyDraft(
                seed.EntryId, "Renamed " + Marker, replyAll: false, display: false, signature: null,
                cc: null, bcc: null, subject: overriddenSubject);
            draftIds.Add(renamed.EntryId);

            ComDraftInfo renamedInfo = RequireMailInfo(renamed.EntryId, hubStoreId);
            Assert.Equal(overriddenSubject, renamedInfo.Subject);
            Assert.True(renamed.ConversationTopicPreserved, "the subject override must preserve the conversation grouping");
            Assert.Equal(seedInfo.ConversationTopic, renamedInfo.ConversationTopic);

            // The Phase-4 threading assertion must stay green WITH the override.
            Assert.StartsWith(seedInfo.ConversationIndex!, renamedInfo.ConversationIndex!, StringComparison.OrdinalIgnoreCase);
            Assert.True(
                renamedInfo.ConversationIndex!.Length > seedInfo.ConversationIndex!.Length,
                "the renamed draft's ConversationIndex must still be a CHILD of the source's");
            Assert.Equal(seedInfo.ConversationId, renamedInfo.ConversationId);
            _output.WriteLine(
                $"A3: subjectOverridden=true topicPreserved=true indexExtends=true "
                + $"(len {seedInfo.ConversationIndex!.Length} -> {renamedInfo.ConversationIndex!.Length}) conversationIdSame=true");
        }
        finally
        {
            foreach (string id in draftIds)
            {
                CleanupDraft(Hub, id);
            }

            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(Hub, Marker);
        }

        AssertNoTaggedArtifactsRemain(Hub, folderIds: null);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "MultipleStores")]
    [Trait("Requires", "Transport")]
    public void ForwardDraft_CcBccAppend_AndSubjectOverrideKeepsTheForwardedContent()
    {
        string hubStoreId = _fixture.GetStoreId(Hub);
        string quoteToken = "FWDQ" + Marker;
        string seedSubject = _fixture.TaggedSubject("fwdseed");
        DateTime sentUtc = LiveOutlookTestMailer.SendSelfMail(
            Hub, seedSubject, "Forward seed.\r\nToken " + quoteToken, attachmentPath: null);
        ComMailBrief seed = WaitForInboxArrival(seedSubject, sentUtc);
        ComDraftInfo seedInfo = RequireMailInfo(seed.EntryId, seed.StoreId ?? hubStoreId);

        string agentText = "Forward body " + Marker;
        string overriddenSubject = _fixture.TaggedSubject("fwd-renamed");
        string? draftId = null;
        try
        {
            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "forward_draft");
            DraftOutcome forward = Service.ForwardDraft(
                seed.EntryId, agentText, Hub, display: false, signature: null,
                cc: "fwdcc-" + Marker + "@example.invalid", bcc: "fwdbcc-" + Marker + "@example.invalid",
                subject: overriddenSubject, importance: "low", requestReadReceipt: null);
            draftId = forward.EntryId;

            ComDraftInfo info = RequireMailInfo(forward.EntryId, hubStoreId);
            Assert.Equal(overriddenSubject, info.Subject);
            Assert.Contains(info.Recipients, r => r.Kind == "cc");
            Assert.Contains(info.Recipients, r => r.Kind == "bcc");
            Assert.Equal("low", forward.Importance);
            Assert.Equal(0, info.Importance);
            Assert.Null(forward.ReadReceiptRequested);

            // Threading + the forwarded content survive the rename.
            Assert.True(forward.ConversationTopicPreserved, "forward subject override must preserve grouping");
            Assert.Equal(seedInfo.ConversationTopic, info.ConversationTopic);
            Assert.StartsWith(seedInfo.ConversationIndex!, info.ConversationIndex!, StringComparison.OrdinalIgnoreCase);

            ReadOutcome read = Service.Read(forward.EntryId, maxBodyChars: 100000);
            int agentAt = read.Body.IndexOf(agentText, StringComparison.Ordinal);
            int quoteAt = read.Body.IndexOf(quoteToken, StringComparison.Ordinal);
            Assert.True(agentAt >= 0, "agent text must be present");
            Assert.True(quoteAt > agentAt, $"forwarded content must stay BELOW the agent text (agent@{agentAt} quote@{quoteAt})");
            _output.WriteLine($"A2/A3[forward]: subjectOverridden=true topicPreserved=true agent@{agentAt} quote@{quoteAt}");
        }
        finally
        {
            if (draftId != null)
            {
                CleanupDraft(Hub, draftId);
            }

            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(Hub, Marker);
        }

        AssertNoTaggedArtifactsRemain(Hub, folderIds: null);
    }

    // ------------------------------------------------------------------ helpers

    private DraftOutcome CreateNewDraft(string account, string label, string bodyMarker, string? signature)
    {
        return Service.NewDraft(
            LiveStoreWriteGuard.Writable(account, StoreWriteKind.Draft, "new_draft"),
            account,
            cc: null,
            _fixture.TaggedSubject(label),
            bodyMarker + " first line.\r\nSecond line.",
            display: false,
            signature: signature);
    }

    /// <summary>
    /// The A1 contract on the RAW HTML: agent body first, then the signature region
    /// anchor, then the signature's own content - and the body is NOT inside the
    /// signature bookmark (the exact shape of the reported defect).
    /// </summary>
    private void AssertBodyAboveIntactSignature(string html, string bodyMarker, string signatureMarker, string label)
    {
        int bodyAt = RequireIndexOf(html, bodyMarker, "agent body");
        int anchorAt = RequireIndexOf(html, "_MailAutoSig", "signature region");
        int signatureAt = RequireIndexOf(html, signatureMarker, "signature content");

        Assert.True(bodyAt < anchorAt, $"{label}: agent body must precede the signature region (body@{bodyAt} sig@{anchorAt})");
        Assert.True(bodyAt < signatureAt, $"{label}: agent body must precede the signature content");

        // The signature keeps its own HTML: the image resource of the test signature
        // survives as a real <img> element rather than being flattened to text.
        Assert.Contains("<img", html, StringComparison.OrdinalIgnoreCase);
        AssertBodyInsideWordSection(html, bodyAt);
        _output.WriteLine($"A1[{label}]: bodyAt={bodyAt} signatureAnchorAt={anchorAt} signatureContentAt={signatureAt} imgPreserved=true");
    }

    /// <summary>
    /// The body must live INSIDE Outlook's own document container, not be spliced in
    /// straight after &lt;body&gt; (which is what the retired HTMLBody string surgery did
    /// and why the agent text did not inherit the message style).
    /// </summary>
    private static void AssertBodyInsideWordSection(string html, int bodyAt)
    {
        int section = html.IndexOf("WordSection1", StringComparison.OrdinalIgnoreCase);
        Assert.True(section >= 0, "Outlook's WordSection container must be present");
        Assert.True(bodyAt > section, $"agent body must sit INSIDE WordSection1 (section@{section} body@{bodyAt})");
    }

    private static int RequireIndexOf(string html, string needle, string what)
    {
        int at = html.IndexOf(needle, StringComparison.Ordinal);
        Assert.True(at >= 0, what + " must be present in the draft HTML");
        return at;
    }

    private string RequireHtmlBody(string entryId, string store)
    {
        string? html = _fixture.VerifySession.TryGetHtmlBody(entryId, _fixture.GetStoreId(store), out string? error);
        Assert.True(!string.IsNullOrEmpty(html), $"draft HTML unavailable: {error ?? "empty"}");
        return html!;
    }

    /// <summary>
    /// Zero-artifact proof tolerant of the documented self-send lag: a straggler is
    /// purged once more (S3-legal - tag AND this run's marker) and only what survives
    /// that fails. See LiveOutlookTestMailer.CountTaggedArtifactsAfterPurgingStragglers.
    /// </summary>
    private void AssertNoTaggedArtifactsRemain(string store, int[]? folderIds)
    {
        int remaining = LiveOutlookTestMailer.CountTaggedArtifactsAfterPurgingStragglers(
            store, Marker, folderIds, out int stragglers);
        if (stragglers > 0)
        {
            _output.WriteLine($"cleanup[{store}]: {stragglers} late-materialized artifact(s) purged (documented lag)");
        }

        Assert.Equal(0, remaining);
    }

    private ComMailBrief WaitForInboxArrival(string seedSubject, DateTime sentUtc)
    {
        return LiveInboxArrival.WaitFor(_fixture.VerifySession, Hub, seedSubject, sentUtc);
    }

    private ComDraftInfo RequireMailInfo(string entryId, string? storeId)
    {
        ComDraftInfo? info = _fixture.VerifySession.TryGetMailInfo(entryId, storeId, out string? error);
        Assert.True(info != null, $"mail info unavailable: {error}");
        return info!;
    }

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
}
