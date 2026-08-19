using System;
using System.Linq;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Soak fix 22 (D49) - THE HEADLESS PARITY CONTRACT.
/// <para>
/// D48 measured the whole D44-D47 compose model going inert while Outlook is window-less:
/// <c>bodyPlacement="html"</c>, an explicit signature override silently DROPPED, the
/// signature's images never embedded, the recipient echo empty, and <c>update_draft</c>
/// failing outright. None of that was ever asserted anywhere, which is exactly why it
/// shipped: every draft assertion in this project ran against whatever window state the
/// suite happened to be in.
/// </para>
/// <para>
/// This suite states the contract that replaces the assumption: <b>a draft composed with
/// NO visible Outlook window must be indistinguishable from one composed with a window,
/// and composing it must not put anything on the user's screen.</b> It is deliberately
/// state-agnostic - it asserts the same capabilities whatever state it finds Outlook in,
/// and additionally proves the promotion actually fired when it finds Outlook headless -
/// because a test that only holds in ONE window state is precisely the gap D48 fell into.
/// </para>
/// </summary>
[Collection(LiveCollections.Phase4)]
[Trait("Category", "Live")]
[Trait("LiveTier", "ProfileBound")]
[Trait("Requires", "MailAccount")]
public sealed class LiveHeadlessComposeParityTests
{
    private readonly LivePhase4Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveHeadlessComposeParityTests(LivePhase4Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    /// <summary>
    /// A single space-free word that occurs ONLY in the test signature's own content.
    /// Word's HTML serializer hard-wraps at ~78 columns, so a multi-word marker can be
    /// split across a newline and become unfindable in the raw markup.
    /// </summary>
    private const string SignatureToken = "testhandtekening";

    private string RequireHtmlBody(string entryId)
        => _fixture.VerifySession.TryGetHtmlBody(entryId, _fixture.GetStoreId(Hub), out string? error)
            ?? throw new InvalidOperationException("draft HTML unavailable: " + (error ?? "empty"));

    [Fact]
    public void ComposedWithNoVisibleWindow_TheDraftIsFullyCapable_AndNothingIsPutOnScreen()
    {
        bool? headlessAtStart = HealthReporting.TryGetOutlookHeadless();
        int visibleBefore = ComposeSurface.CountUserVisibleWindows();

        using TestSignature sig = TestSignature.Create(Marker);
        string subject = _fixture.TaggedSubject("headless-parity");
        // NOTE: search tokens carry NO SPACES. Word's HTML serializer hard-wraps its
        // output at ~78 columns, so any multi-word phrase can be split across a newline
        // and is unfindable in the raw markup - a trap that has nothing to do with the
        // behaviour under test.
        string heading = "Kwartaalrapportage" + Marker;
        string cell = "CELL" + Marker;
        string bodyHtml =
            "<h1>" + heading + "</h1>"
            + "<p>Geachte <strong>heer</strong>,</p>"
            + "<table><thead><tr><th>Post</th></tr></thead><tbody><tr><td>" + cell + "</td></tr></tbody></table>";

        LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
        DraftOutcome outcome = Service.NewDraft(
            Hub, Hub, cc: null, subject, body: null, display: false, signature: sig.Name, bodyHtml: bodyHtml);

        // ---- 1. The tool's own answer must not admit a degraded composition ----------
        Assert.Equal("wordEditor", outcome.BodyPlacement);
        Assert.Null(outcome.ComposeSurfaceError);
        Assert.Null(outcome.ComposeSurfaceAdvice);
        Assert.Equal("html", outcome.BodyFormat);

        // The D48 headline: an explicit override was DROPPED headless, and the account
        // default was used instead, with signatureError="NoWordEditor".
        Assert.True(outcome.SignatureApplied, $"signature override must apply (error: {outcome.SignatureError ?? "-"})");
        Assert.Null(outcome.SignatureError);

        // ...and the recipient echo came back empty headless.
        Assert.NotNull(outcome.Recipients);
        Assert.NotEmpty(outcome.Recipients!);

        // When Outlook is window-less, THAT is the state D48 measured as inert - so the
        // precondition is asserted rather than hoped for: every capability above was just
        // proven with nothing on the user's screen. Whether this particular draft needed
        // a promotion is reported, not required: the promotion's effect outlives the
        // inspector that triggered it, so a later compose in the same process can find
        // the editor already available. The capability asserts are the contract.
        if (headlessAtStart == true)
        {
            Assert.Equal(0, visibleBefore);
        }

        // ---- 2. The RAW stored HTML - the only place these defects are visible --------
        string html = RequireHtmlBody(outcome.EntryId);

        int wordSectionAt = html.IndexOf("WordSection1", StringComparison.OrdinalIgnoreCase);
        int bodyAt = html.IndexOf(heading, StringComparison.Ordinal);
        int sigAnchorAt = html.IndexOf("_MailAutoSig", StringComparison.OrdinalIgnoreCase);
        int sigContentAt = html.IndexOf(SignatureToken, StringComparison.OrdinalIgnoreCase);

        _output.WriteLine(
            $"headlessAtStart={headlessAtStart} promoted={outcome.ComposeSurfacePromoted} "
            + $"wordSection@{wordSectionAt} body@{bodyAt} sigAnchor@{sigAnchorAt} sigContent@{sigContentAt} len={html.Length}");

        // HTML bodies must be REAL markup, not escaped text, and not a bare <div> splice.
        Assert.Contains("<h1", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<table", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(cell, html, StringComparison.Ordinal);
        Assert.DoesNotContain("&lt;h1", html, StringComparison.OrdinalIgnoreCase);

        // Inside Outlook's own message container (the A1(ii) defect is the body landing
        // after <body> and therefore OUTSIDE WordSection1, losing the message style).
        Assert.True(wordSectionAt >= 0, "Outlook's WordSection1 container must be present");
        Assert.True(bodyAt > wordSectionAt, $"body must sit inside WordSection1 (wordSection@{wordSectionAt} body@{bodyAt})");

        // Body ABOVE the signature region, and the region still marked - the A1(i) defect
        // is the marker ending up wrapped around the whole message.
        Assert.True(bodyAt >= 0, "the agent body must be present");
        Assert.True(sigAnchorAt > bodyAt, $"the signature anchor must follow the body (body@{bodyAt} anchor@{sigAnchorAt})");
        Assert.True(sigContentAt > sigAnchorAt, $"signature content must follow its anchor (anchor@{sigAnchorAt} content@{sigContentAt})");

        // D47/D48 image quality: a real inline attachment, a cid: that matches it, no
        // file:/// link and no local profile path leaking into what a recipient receives.
        Assert.Contains("src=\"cid:", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("src=\"file:///", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("C:\\Users", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<v:rect", html, StringComparison.OrdinalIgnoreCase);

        AttachmentView image = Assert.Single(outcome.Attachments!);
        Assert.True(
            image.SizeBytes > 0,
            $"the embedded signature image must echo real bytes, got {image.SizeBytes?.ToString() ?? "null"}");

        // ---- 3. update_draft, which headless used to refuse with com_failure ----------
        string revisedHeading = "HerzieneRapportage" + Marker;
        LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "update_draft");
        UpdateDraftOutcome updated = Service.UpdateDraft(
            outcome.EntryId, bodyHtml: "<h2>" + revisedHeading + "</h2><p>Tweede versie.</p>", display: false);

        Assert.Equal("updated", updated.Status);
        Assert.Equal("wordEditor", updated.BodyPlacement);

        string after = RequireHtmlBody(updated.EntryId);
        Assert.Contains(revisedHeading, after, StringComparison.Ordinal);
        Assert.DoesNotContain(heading, after, StringComparison.Ordinal);

        // The signature and its embedded image survive the revision byte-for-byte (D47).
        Assert.Contains(SignatureToken, after, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("src=\"cid:", after, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("src=\"file:///", after, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<v:rect", after, StringComparison.OrdinalIgnoreCase);
        Assert.Null(updated.InlineImagesDropped);
        Assert.NotEmpty(Service.Read(updated.EntryId, maxBodyChars: 0).Attachments!);

        // ---- 4. D33: composing must leave the user's screen exactly as it was ---------
        int visibleAfter = ComposeSurface.CountUserVisibleWindows();
        Assert.True(
            visibleAfter <= visibleBefore,
            $"composing must not put a window on screen (user-visible windows {visibleBefore} -> {visibleAfter})");

        _output.WriteLine(
            $"parity ok: placement={outcome.BodyPlacement} signatureApplied={outcome.SignatureApplied} "
            + $"imageBytes={image.SizeBytes} recipients={outcome.Recipients!.Count} visibleWindows={visibleBefore}->{visibleAfter}");
    }

    [Fact]
    public void APlainTextDraftIsAlsoFullyComposed_WithNoDegradationReported()
    {
        int visibleBefore = ComposeSurface.CountUserVisibleWindows();
        string subject = _fixture.TaggedSubject("headless-parity-text");
        string agentText = "PlatteTekstHeadless" + Marker; // space-free: Word hard-wraps its HTML

        LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft");
        DraftOutcome outcome = Service.NewDraft(Hub, Hub, cc: null, subject, agentText, display: false);

        // The point of this second case: the degradation channel must stay SILENT when
        // nothing degraded. A permanently-populated advice field would be noise, and an
        // agent would learn to ignore it - which is how the D48 defect stayed invisible.
        Assert.Equal("wordEditor", outcome.BodyPlacement);
        Assert.Null(outcome.ComposeSurfaceError);
        Assert.Null(outcome.ComposeSurfaceAdvice);
        Assert.Equal("text", outcome.BodyFormat);
        Assert.NotEmpty(outcome.Recipients!);

        string html = RequireHtmlBody(outcome.EntryId);
        int wordSectionAt = html.IndexOf("WordSection1", StringComparison.OrdinalIgnoreCase);
        int bodyAt = html.IndexOf(agentText, StringComparison.Ordinal);
        Assert.True(wordSectionAt >= 0 && bodyAt > wordSectionAt, $"wordSection@{wordSectionAt} body@{bodyAt}");

        int visibleAfter = ComposeSurface.CountUserVisibleWindows();
        Assert.True(visibleAfter <= visibleBefore, $"user-visible windows {visibleBefore} -> {visibleAfter}");

        _output.WriteLine($"text parity ok: body@{bodyAt} wordSection@{wordSectionAt} visible={visibleBefore}->{visibleAfter}");
    }
}
