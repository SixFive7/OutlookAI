using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Soak-fix batch B live acceptance (B1 + B2, v3.MD D45) - the blocker from the field
/// report: an agent could not produce a formatted formal letter.
/// <para>
/// <b>B1</b> pins that <c>body_html</c> arrives as REAL HTML inside the DRAFT REGION:
/// headings, bold, a bulleted list, a table and a link survive as markup in the SAVED
/// <c>HTMLBody</c>, they sit inside Outlook's own WordSection1 container, and the
/// signature region and (for replies) the quoted original are intact and BELOW the body.
/// The plain-text path must be byte-for-byte the behavior batch A shipped.
/// </para>
/// <para>
/// <b>B2</b> pins that <c>read include_html=true</c> hands that same HTML back for a
/// freshly created draft addressed by its EntryID - drafts are not in the search index,
/// so the direct-EntryID path is the only way in - with true totals and truncation flags.
/// </para>
/// Everything runs in the hub store only, tagged with the run marker, deleted through the
/// tested helpers (S3); the one screenshot shows agent-authored content only (S5).
/// </summary>
[Collection(LiveCollections.Phase4)]
[Trait("Category", "Live")]
[Trait("LiveTier", "ProfileBound")]
[Trait("Requires", "MailAccount")]
[Trait("Requires", "Transport")]
[Trait("Requires", "InteractiveDesktop")]
public sealed class LiveHtmlDraftTests
{
    private readonly LivePhase4Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveHtmlDraftTests(LivePhase4Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    /// <summary>
    /// The formal-letter shape the field report could not produce. It deliberately
    /// carries non-ASCII text and a currency symbol: the HTML reaches Word through a
    /// temporary file, so the wrong encoding there would silently mojibake every accented
    /// character in a Dutch or German letter.
    /// </summary>
    private string RichLetterHtml(string token)
    {
        return "<h1>Formal notice " + token + "</h1>"
            + "<h2>Reference " + token + "</h2>"
            + "<p>Geachte mevrouw Jansen,</p>"
            + "<p>Wij bevestigen de <strong>afspraken</strong> hieronder, geldig vóór 1 september:</p>"
            + "<ul><li>First point " + token + "</li><li>Tweede punt – coördinatie</li></ul>"
            + "<table border=\"1\"><thead><tr><th>Item</th><th>Amount</th></tr></thead>"
            + "<tbody><tr><td>Licence</td><td>€ 120</td></tr></tbody></table>"
            + "<p>Details are on <a href=\"https://example.com/terms\">our terms page</a>.</p>";
    }

    /// <summary>
    /// Non-ASCII may be stored literally or as a numeric/named entity depending on the
    /// charset Outlook settles on - any of those means it survived; mojibake does not.
    /// </summary>
    private void AssertNonAsciiSurvived(string html, string label)
    {
        // Escapes, not literals: the assertion must not depend on how this source file
        // happens to be decoded by the compiler.
        string[] accentForms = ["vóór", "v&oacute;&oacute;r", "v&#243;&#243;r"];
        string[] euroForms = ["€", "&euro;", "&#8364;"];

        Assert.True(
            accentForms.Any(f => html.IndexOf(f, StringComparison.Ordinal) >= 0),
            label + ": accented text must survive the temp-file round trip");
        Assert.True(
            euroForms.Any(f => html.IndexOf(f, StringComparison.Ordinal) >= 0),
            label + ": the euro sign must survive the temp-file round trip");

        // The classic UTF-8-read-as-Windows-1252 signature.
        Assert.DoesNotContain("Ã³", html, StringComparison.Ordinal);
        _output.WriteLine($"[{label}]: non-ASCII preserved (no mojibake)");
    }

    [Fact]
    public void NewDraft_BodyHtml_RichStructuresSurviveInTheDraftRegion_AboveAnIntactSignature()
    {
        using TestSignature sig = TestSignature.Create(Marker);
        string token = "B1RICH" + Marker;

        LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
        DraftOutcome outcome = Service.NewDraft(
            Hub, Hub, cc: null, _fixture.TaggedSubject("b1-rich"), body: null, display: false,
            signature: sig.Name, bcc: null, importance: null, requestReadReceipt: null,
            bodyHtml: RichLetterHtml(token));
        try
        {
            Assert.Equal("html", outcome.BodyFormat);
            Assert.Equal("wordEditor", outcome.BodyPlacement);
            Assert.True(outcome.SignatureApplied, $"override must apply (error: {outcome.SignatureError ?? "-"})");
            Assert.Null(outcome.HtmlAdjustments);

            string html = RequireHtmlBody(outcome.EntryId);

            // Every construct the field report asked for, as REAL markup - not escaped
            // text and not wrapped in <pre>.
            Assert.DoesNotContain("&lt;h1", html, StringComparison.OrdinalIgnoreCase);
            AssertRendered(html, "<h1", "heading 1");
            AssertRendered(html, "<h2", "heading 2");
            AssertRendered(html, "<table", "table");
            AssertRendered(html, "<td", "table cell");
            AssertRendered(html, "href=\"https://example.com/terms\"", "link");

            // Word renders <strong>/<ul><li> through its own HTML converter, so the tag
            // NAMES may differ in the saved markup - what must hold is that the structure
            // is real: bold formatting and a genuine list, never literal tag text.
            Assert.True(
                html.IndexOf("<strong", StringComparison.OrdinalIgnoreCase) >= 0
                || html.IndexOf("<b>", StringComparison.OrdinalIgnoreCase) >= 0
                || html.IndexOf("bold", StringComparison.OrdinalIgnoreCase) >= 0,
                "bold must survive as formatting");
            Assert.True(
                html.IndexOf("<li", StringComparison.OrdinalIgnoreCase) >= 0
                || html.IndexOf("MsoListParagraph", StringComparison.OrdinalIgnoreCase) >= 0
                || html.IndexOf("l0 level1", StringComparison.OrdinalIgnoreCase) >= 0,
                "the bulleted list must survive as a list");

            AssertNonAsciiSurvived(html, "b1-rich");
            AssertBodyAboveIntactSignature(html, token, "testhandtekening", "b1-rich");
            _output.WriteLine("B1: h1/h2/table/td/link present as markup; body above the signature region");
        }
        finally
        {
            CleanupDraft(outcome.EntryId);
        }

        AssertNoTaggedArtifactsRemain();
    }

    [Fact]
    public void ReplyDraft_BodyHtml_LeavesTheQuotedOriginalIntactAndBelowTheBody()
    {
        string quoteToken = "B1Q" + Marker;
        string seedSubject = _fixture.TaggedSubject("b1seed");
        DateTime sentUtc = LiveOutlookTestMailer.SendSelfMail(
            Hub, seedSubject, "Reply seed for the HTML body suite.\r\nToken " + quoteToken, attachmentPath: null);
        ComMailBrief seed = WaitForInboxArrival(seedSubject, sentUtc);

        string token = "B1REPLY" + Marker;
        string? draftId = null;
        try
        {
            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "reply_draft");
            DraftOutcome reply = Service.ReplyDraft(
                seed.EntryId, body: null, replyAll: false, display: false, signature: null,
                cc: null, bcc: null, subject: null, importance: null, requestReadReceipt: null,
                bodyHtml: RichLetterHtml(token));
            draftId = reply.EntryId;

            Assert.Equal("html", reply.BodyFormat);
            Assert.Equal("wordEditor", reply.BodyPlacement);

            string html = RequireHtmlBody(reply.EntryId);
            int bodyAt = RequireIndexOf(html, token, "agent body");
            int quoteAt = RequireIndexOf(html, quoteToken, "quoted original");

            Assert.True(bodyAt < quoteAt, $"agent body must precede the quoted original (body@{bodyAt} quote@{quoteAt})");
            AssertRendered(html, "<table", "table");
            AssertRendered(html, "<h1", "heading 1");
            AssertBodyInsideWordSection(html, bodyAt);
            _output.WriteLine($"B1 reply: bodyAt={bodyAt} quoteAt={quoteAt} structuresPreserved=true");
        }
        finally
        {
            if (draftId != null)
            {
                CleanupDraft(draftId);
            }

            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(Hub, Marker);
        }

        AssertNoTaggedArtifactsRemain();
    }

    [Fact]
    public void PlainTextBody_IsUnchanged_ByTheHtmlOption()
    {
        using TestSignature sig = TestSignature.Create(Marker);
        string token = "B1TEXT" + Marker;

        LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
        DraftOutcome outcome = Service.NewDraft(
            Hub, Hub, cc: null, _fixture.TaggedSubject("b1-text"),
            body: token + " first line.\r\nSecond line.", display: false, signature: sig.Name);
        try
        {
            Assert.Equal("text", outcome.BodyFormat);
            Assert.Equal("wordEditor", outcome.BodyPlacement);
            Assert.Null(outcome.HtmlAdjustments);

            string html = RequireHtmlBody(outcome.EntryId);
            AssertBodyAboveIntactSignature(html, token, "testhandtekening", "b1-text");

            // The batch-A contract: plain text stays plain - no agent-supplied structure
            // appears just because the HTML path now exists.
            Assert.DoesNotContain("<h1", html, StringComparison.OrdinalIgnoreCase);
            _output.WriteLine("B1: plain-text path unchanged (body above the intact signature, no injected structure)");
        }
        finally
        {
            CleanupDraft(outcome.EntryId);
        }

        AssertNoTaggedArtifactsRemain();
    }

    [Fact]
    public void MalformedBodyHtml_IsRepaired_AndTheAdjustmentsAreReported()
    {
        using TestSignature sig = TestSignature.Create(Marker);
        string token = "B1FIX" + Marker;

        LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
        DraftOutcome outcome = Service.NewDraft(
            Hub, Hub, cc: null, _fixture.TaggedSubject("b1-malformed"), body: null, display: false,
            signature: sig.Name, bcc: null, importance: null, requestReadReceipt: null,
            bodyHtml: "<html><body><script>alert(1)</script><p id=\"_MailAutoSig\">" + token
                + "<b>never closed<ul><li>stray");
        try
        {
            Assert.NotNull(outcome.HtmlAdjustments);
            Assert.Contains(outcome.HtmlAdjustments!, a => a.Contains("<script>", StringComparison.Ordinal));
            Assert.Contains(outcome.HtmlAdjustments!, a => a.Contains("id attribute", StringComparison.Ordinal));
            Assert.Contains(outcome.HtmlAdjustments!, a => a.Contains("unclosed", StringComparison.Ordinal));

            string html = RequireHtmlBody(outcome.EntryId);
            Assert.DoesNotContain("alert(1)", html, StringComparison.Ordinal);

            // The forged id could have redrawn the signature boundary - the signature
            // region must still be exactly one real region, below the body.
            AssertBodyAboveIntactSignature(html, token, "testhandtekening", "b1-malformed");
            _output.WriteLine($"B1: malformed input repaired; adjustments={outcome.HtmlAdjustments!.Count}");
        }
        finally
        {
            CleanupDraft(outcome.EntryId);
        }

        AssertNoTaggedArtifactsRemain();
    }

    [Fact]
    public void Read_IncludeHtml_ReturnsTheStoredHtmlOfAFreshDraftByEntryId_WithTruncationFlags()
    {
        string token = "B2READ" + Marker;

        LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
        DraftOutcome outcome = Service.NewDraft(
            Hub, Hub, cc: null, _fixture.TaggedSubject("b2-read"), body: null, display: false,
            signature: null, bcc: null, importance: null, requestReadReceipt: null,
            bodyHtml: RichLetterHtml(token));
        try
        {
            // Drafts are not in the search index: this exercises the direct-EntryID path.
            // The DEFAULT html budget must be big enough to carry a real draft past
            // Outlook's ~40 KB leading stylesheet - that is the whole point of the option.
            ReadOutcome full = Service.Read(outcome.EntryId, includeHtml: true);

            Assert.Equal("directEntryId", full.LocatedVia);
            Assert.NotNull(full.BodyHtml);
            Assert.NotNull(full.BodyHtmlTotalChars);
            Assert.False(full.BodyHtmlTruncated);
            Assert.Equal(full.BodyHtmlTotalChars, full.BodyHtml!.Length);
            AssertRendered(full.BodyHtml!, "<table", "table");
            AssertRendered(full.BodyHtml!, "<h1", "heading 1");
            Assert.Contains(token, full.BodyHtml!, StringComparison.Ordinal);

            // The plain-text rendering hides exactly this - the reason include_html exists.
            Assert.DoesNotContain("<table", full.Body, StringComparison.OrdinalIgnoreCase);

            // Payload discipline: the HTML budget is its own knob, and the true total is
            // reported whether or not the window cut it.
            ReadOutcome capped = Service.Read(outcome.EntryId, includeHtml: true, maxHtmlChars: 200);
            Assert.True(capped.BodyHtmlTruncated);
            Assert.Equal(200, capped.BodyHtml!.Length);
            Assert.Equal(full.BodyHtmlTotalChars, capped.BodyHtmlTotalChars);

            // The text budget must NOT govern the HTML (the defect this knob fixes).
            ReadOutcome tinyText = Service.Read(outcome.EntryId, maxBodyChars: 100, includeHtml: true);
            Assert.False(tinyText.BodyHtmlTruncated);
            Assert.Contains(token, tinyText.BodyHtml!, StringComparison.Ordinal);

            // A continuation read serves the TEXT body from the cache without touching
            // COM for it - the HTML must still come back rather than fall through the
            // cache shortcut.
            ReadOutcome paged = Service.Read(outcome.EntryId, maxBodyChars: 50, bodyOffset: 20, includeHtml: true);
            Assert.NotNull(paged.BodyHtml);
            Assert.Equal(full.BodyHtmlTotalChars, paged.BodyHtmlTotalChars);

            // Default stays off: no HTML unless it was asked for.
            ReadOutcome plain = Service.Read(outcome.EntryId);
            Assert.Null(plain.BodyHtml);
            Assert.Null(plain.BodyHtmlTotalChars);
            Assert.Null(plain.BodyHtmlTruncated);

            _output.WriteLine(
                $"B2: bodyHtmlTotalChars={full.BodyHtmlTotalChars} truncatedAt200=true locatedVia={full.LocatedVia}");
        }
        finally
        {
            CleanupDraft(outcome.EntryId);
        }

        AssertNoTaggedArtifactsRemain();
    }

    [Fact]
    public void NewDraft_BodyHtml_DisplayCase_ScreenshotOfTheRenderedLetter_ThenClosed()
    {
        string token = "B1SHOT" + Marker;

        // Make sure Outlook owns a normal window BEFORE this test opens one. Displaying
        // and then closing an Inspector on an Outlook that has NO other window makes the
        // instance EXIT (the D43 idle-self-exit hazard, live-proven again here: the next
        // test in the collection then died on RPC_S_SERVER_UNAVAILABLE 0x800706BA). The
        // Explorer is scoped to the hub store, which S5 allows.
        _fixture.Service.GotoFolder(Hub);

        // S5: hub store, and every visible character is agent-authored.
        LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
        DraftOutcome outcome = Service.NewDraft(
            Hub, Hub, cc: null, _fixture.TaggedSubject("b1-render"), body: null, display: true,
            signature: null, bcc: null, importance: null, requestReadReceipt: null,
            bodyHtml: RichLetterHtml(token));
        try
        {
            Assert.True(outcome.Displayed);
            Assert.True(
                WaitForInspector(outcome.EntryId, TimeSpan.FromSeconds(20)),
                "no Inspector appeared for the displayed HTML draft within 20 s");

            // S5 evidence is best-effort (soak fix 19): a lost foreground race means no
            // file, never a file of someone else's window. The rendering itself is
            // asserted against the raw HTMLBody elsewhere in this class.
            try
            {
                string path = ScreenCapture.CaptureOutlookWindowByCaptionFragment(
                    Marker,
                    _fixture.ScreenshotsDirectory,
                    $"soakfixB-html-draft-{DateTime.Now:yyyyMMdd-HHmmss}.png");
                FileInfo file = new(path);
                Assert.True(file.Exists && file.Length > 0, "screenshot must exist and be non-empty");
                _output.WriteLine($"B1 screenshot saved: {path} bytes={file.Length}");
            }
            catch (ScreenCaptureSkippedException ex)
            {
                _output.WriteLine($"S5 evidence skipped (no polluted capture written): {ex.Message}");
            }
        }
        finally
        {
            bool closed = _fixture.VerifySession.TryCloseInspectorByEntryId(outcome.EntryId, out string? closeError);
            _output.WriteLine($"inspector close requested: ok={closed} err={closeError ?? "-"}");
            CleanupDraft(outcome.EntryId);
        }

        // Outlook must have SURVIVED the close - the rest of the tier depends on it. A
        // dead instance surfaces here as RPC_S_SERVER_UNAVAILABLE, naming the real cause
        // instead of failing whichever test happens to run next.
        Assert.NotNull(_fixture.VerifySession.GetOpenInspectors());
        AssertNoTaggedArtifactsRemain();
    }

    // ------------------------------------------------------------------ helpers

    private static void AssertRendered(string html, string needle, string what)
    {
        Assert.True(
            html.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0,
            what + " must appear as real markup in the saved HTML");
    }

    private static int RequireIndexOf(string html, string needle, string what)
    {
        int at = html.IndexOf(needle, StringComparison.Ordinal);
        Assert.True(at >= 0, what + " must be present in the draft HTML");
        return at;
    }

    private void AssertBodyAboveIntactSignature(string html, string bodyMarker, string signatureMarker, string label)
    {
        int bodyAt = RequireIndexOf(html, bodyMarker, "agent body");
        int anchorAt = RequireIndexOf(html, "_MailAutoSig", "signature region");
        int signatureAt = RequireIndexOf(html, signatureMarker, "signature content");

        Assert.True(bodyAt < anchorAt, $"{label}: agent body must precede the signature region (body@{bodyAt} sig@{anchorAt})");
        Assert.True(bodyAt < signatureAt, $"{label}: agent body must precede the signature content");
        Assert.Contains("<img", html, StringComparison.OrdinalIgnoreCase);
        AssertBodyInsideWordSection(html, bodyAt);
        _output.WriteLine($"[{label}]: bodyAt={bodyAt} signatureAnchorAt={anchorAt} signatureContentAt={signatureAt}");
    }

    private static void AssertBodyInsideWordSection(string html, int bodyAt)
    {
        int section = html.IndexOf("WordSection1", StringComparison.OrdinalIgnoreCase);
        Assert.True(section >= 0, "Outlook's WordSection container must be present");
        Assert.True(bodyAt > section, $"agent body must sit INSIDE WordSection1 (section@{section} body@{bodyAt})");
    }

    private string RequireHtmlBody(string entryId)
    {
        string? html = _fixture.VerifySession.TryGetHtmlBody(entryId, _fixture.GetStoreId(Hub), out string? error);
        Assert.True(!string.IsNullOrEmpty(html), $"draft HTML unavailable: {error ?? "empty"}");
        return html!;
    }

    private bool WaitForInspector(string entryId, TimeSpan timeout)
    {
        LiveWaitBudget wait = LiveWaitBudget.Of(timeout);
        while (wait.HasTimeLeft)
        {
            if (_fixture.VerifySession.GetOpenInspectors().Any(i =>
                string.Equals(i.EntryId, entryId, StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }

            Thread.Sleep(500);
        }

        return false;
    }

    private ComMailBrief WaitForInboxArrival(string seedSubject, DateTime sentUtc)
    {
        return LiveInboxArrival.WaitFor(_fixture.VerifySession, Hub, seedSubject, sentUtc);
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
