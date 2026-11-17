using System.Diagnostics;
using System.Globalization;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using OutlookAI.Core.Text;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Q149 (d) stage 3 - the HTML-only draft (c) against today's Word-composed draft (b). (c) replaces Word only
/// if its drafts are indistinguishable from today's, opened in Outlook and as received, so every pair here
/// is the SAME body and the SAME signature (the test signature, with its picture) composed both ways, and
/// compared on:
/// <list type="bullet">
/// <item><b>the stylesheet</b> - the HTML-only draft borrows today's whole <c>&lt;head&gt;</c>, its best
/// case; what an item made with no inspector holds of its own is recorded beside it (C1);</item>
/// <item><b>the signature with its pictures</b> - the <c>_MailAutoSig</c> marker below the body, the picture
/// as a hidden <c>cid:</c> attachment with real bytes, no <c>file:///</c> link;</item>
/// <item><b>the quoted original</b> - what <c>Reply()</c> and <c>Forward()</c> hand over with no inspector
/// (C4), and the body placed above it;</item>
/// <item><b>update_draft</b> - Word's update against the HTML-only one, on both kinds of draft (C3);</item>
/// <item><b>how it looks opened in Outlook</b> - Word's layout of each, paragraph by paragraph: text, style,
/// font, size, spacing, alignment, pictures;</item>
/// <item><b>how it looks as received</b> - both sent through the guest's loopback mail sink with the product's
/// own two-step send, and the received copies compared the same two ways (C5).</item>
/// </list>
/// Everything runs in a windowed Outlook - the HTML-only draft needs no window, and the comparison should not
/// depend on one. Nothing is sent to anyone but the hub itself (Q110).
/// </summary>
[Collection(LiveCollections.ComposeSurface)]
[Trait("Category", "Live")]
public sealed class LiveQ149HtmlOnlyParityTests
{
    private const int NewMailPairs = 5;
    private const int DerivedPairs = 3;
    private const int UpdateSets = 2;
    private const int ReceivedPairs = 2;

    /// <summary>A space-free word that occurs only in the test signature (Word hard-wraps its HTML).</summary>
    private const string SignatureWord = "testhandtekening";

    private readonly LiveQ149Fixture _fixture;
    private readonly ITestOutputHelper _output;
    private readonly List<(string Aspect, bool Same)> _aspects = new();

    public LiveQ149HtmlOnlyParityTests(LiveQ149Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private Q149Run Run => _fixture.Run;

    private MailService Service => _fixture.Service;

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "InteractiveDesktop")]
    [Trait("Requires", "Transport")]
    public void C_TheHtmlOnlyDraft_AgainstTodaysDraft()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        Parity(NewMailPairs, DerivedPairs, UpdateSets, ReceivedPairs);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "InteractiveDesktop")]
    [Trait("Requires", "Transport")]
    public void Smoke_EveryParityCellOnce()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        Parity(1, 1, 1, 1);
    }

    private void Parity(int newPairs, int derivedPairs, int updateSets, int receivedPairs)
    {
        LiveComposeProbe.EnsureWindowedOutlook(Run.Settings, Run.Log);
        _aspects.Clear();
        using TestSignature signature = TestSignature.Create(Run.Marker);
        using Q149OutlookClient client = Q149OutlookClient.Connect(new Stopwatch(), allowStart: false);
        string storeId = client.Run((app, ns) =>
        {
            dynamic? store = Q149Drafts.FindStore(ns, Run.Hub);
            try
            {
                return (string)store!.StoreID;
            }
            finally
            {
                Q149Com.Release(store);
            }
        });

        // C2: a new mail, both ways.
        List<(string Product, string Message, string AsImported, string Token)> newMail = new();
        for (int i = 1; i <= newPairs; i++)
        {
            newMail.Add(NewMailPair(client, storeId, signature, i, html: i % 2 == 1));
        }

        // C4: a reply and a forward, both ways, of a mail this run sent itself.
        string sourceEntryId = SeedSource(storeId, out string quoteToken);
        for (int i = 1; i <= derivedPairs; i++)
        {
            DerivedPair(client, storeId, signature, sourceEntryId, quoteToken, forward: false, i);
            DerivedPair(client, storeId, signature, sourceEntryId, quoteToken, forward: true, i);
        }

        // C3: update_draft across the two kinds of draft.
        for (int i = 1; i <= updateSets; i++)
        {
            UpdateSet(client, storeId, signature, i);
        }

        // C5: as received through the loopback sink.
        for (int i = 1; i <= Math.Min(receivedPairs, newMail.Count); i++)
        {
            ReceivedPair(client, storeId, newMail[i - 1], i);
        }

        string verdict = Q149Decisions.Parity(_aspects
            .GroupBy(a => a.Aspect, StringComparer.Ordinal)
            .Select(g => (g.Key, g.All(a => a.Same)))
            .ToList());
        string line = Q149Cells.SummaryPrefix + "S3: " + string.Join("; ", _aspects
            .GroupBy(a => a.Aspect, StringComparer.Ordinal)
            .Select(g => g.Key + " same in " + g.Count(a => a.Same).ToString(CultureInfo.InvariantCulture) + "/" + g.Count().ToString(CultureInfo.InvariantCulture)))
            + "; " + verdict;
        _output.WriteLine(line);
        Console.WriteLine(line);
        Assert.NotEmpty(_aspects);
    }

    // ------------------------------------------------------------------ C2: a new mail

    /// <summary>The two ways (c) may style what it places, each compared with today's draft on its own line.</summary>
    private static readonly HtmlOnlyStyle[] Styles = { HtmlOnlyStyle.Message, HtmlOnlyStyle.AsImported };

    private (string Product, string Message, string AsImported, string Token) NewMailPair(Q149OutlookClient client, string storeId, TestSignature signature, int i, bool html)
    {
        string token = "Q149BODY" + i.ToString(CultureInfo.InvariantCulture) + Run.Marker;
        (string? text, string? fragment) = Body(token, html);
        Q149Attempt productAttempt = new(Q149Cells.C2 + "-product", i);
        DraftOutcome product = ProductDraftWithRetry(
            productAttempt,
            label => Service.NewDraft(Run.Hub, Run.Hub, cc: null, Run.TaggedSubject(label), text, display: false, signature: signature.Name, bodyHtml: fragment),
            "c2-product-" + i.ToString(CultureInfo.InvariantCulture));
        Run.Write(productAttempt);
        string productHtml = Q149Drafts.Read(Run, product.EntryId, storeId).Html;
        (string, string, string)? frame = HtmlOnlyDraftComposer.ReadFrame(productHtml);
        Assert.True(frame != null, "Q149: today's draft has no <html>, <head> and <body> to take the stylesheet from");

        List<string> made = new();
        foreach (HtmlOnlyStyle style in Styles)
        {
            Q149Attempt attempt = new(Q149Cells.C2, (i * 10) + (int)style);
            attempt.Set("bodyFormat", html ? "html" : "text").Set("variant", style.ToString())
                .Set("productPlacement", product.BodyPlacement).Set("productComposeError", product.ComposeSurfaceError);
            string cid = "image001.png@" + Run.Marker + style.ToString();
            string signatureBlock = HtmlOnlyDraftComposer.SignatureBlock(File.ReadAllText(signature.FilePath), src => SignaturePicture(src, cid), style);
            string paragraphs = html ? HtmlOnlyDraftComposer.ParagraphsFromHtml(fragment!, style) : HtmlOnlyDraftComposer.ParagraphsFromText(text!);
            (string htmlOnly, Q149HtmlFacts stub) = Q149Drafts.CreateHtmlOnlyDraft(
                Run, client, "c2-htmlonly-" + style.ToString() + "-" + i.ToString(CultureInfo.InvariantCulture), Pictures(signature, cid),
                _ => HtmlOnlyDraftComposer.ComposeNew(frame!.Value, paragraphs, signatureBlock));

            // C1: what a new item holds with no inspector ever - the stylesheet (c) would have to bring itself.
            attempt.Set("noInspectorTemplate", stub.ToFields());
            if (style == HtmlOnlyStyle.Message)
            {
                Record(Q149Cells.C1, "a stylesheet for a new mail without Word", stub.StyleBlocks > 0 && stub.WordSectionAt >= 0);
            }

            Compare(attempt, client, storeId, product.EntryId, htmlOnly, token, "new mail, " + style.ToString());
            Run.Write(attempt);
            made.Add(htmlOnly);
        }

        return (product.EntryId, made[0], made[1], token);
    }

    /// <summary>
    /// Today's draft, asked of the product. A draft it composed by the HTML fallback is no "today's draft" to
    /// compare with, so that is recorded - Outlook's own error code with it - and the draft asked for once more.
    /// </summary>
    private DraftOutcome ProductDraftWithRetry(Q149Attempt attempt, Func<string, DraftOutcome> compose, string label)
    {
        LiveStoreWriteGuard.Assert(Run.Hub, StoreWriteKind.Draft, "q149 product draft");
        DraftOutcome draft = compose(label);
        Run.Allow(draft.EntryId);
        attempt.Set("placement1", draft.BodyPlacement).Set("composeError1", draft.ComposeSurfaceError).Set("signatureApplied1", draft.SignatureApplied);
        if (string.Equals(draft.BodyPlacement, "wordEditor", StringComparison.Ordinal))
        {
            return draft;
        }

        Run.Log(label + ": the product composed by the HTML fallback (" + (draft.ComposeSurfaceError ?? "no error") + "); asking once more");
        draft = compose(label + "-again");
        Run.Allow(draft.EntryId);
        attempt.Set("placement2", draft.BodyPlacement).Set("composeError2", draft.ComposeSurfaceError).Set("signatureApplied2", draft.SignatureApplied);
        return draft;
    }

    // ------------------------------------------------------------------ C4: reply and forward

    private string SeedSource(string storeId, out string quoteToken)
    {
        quoteToken = "Q149QUOTE" + Run.Marker;
        string subject = Run.TaggedSubject("c4-source");
        DateTime sentUtc = LiveOutlookTestMailer.SendSelfMail(Run.Hub, subject, "Origineel bericht " + quoteToken + "\r\nTweede regel.", attachmentPath: null);
        using OutlookComSession session = OutlookComSession.Connect(allowStartingOutlook: false, onOutlookGone: null, lifetimePin: false);
        ComMailBrief arrived = LiveInboxArrival.WaitFor(session, Run.Hub, subject, sentUtc);
        Run.Allow(arrived.EntryId);
        return arrived.EntryId;
    }

    private void DerivedPair(Q149OutlookClient client, string storeId, TestSignature signature, string sourceEntryId, string quoteToken, bool forward, int i)
    {
        string kind = forward ? "forward" : "reply";
        string token = "Q149BODY" + kind + i.ToString(CultureInfo.InvariantCulture) + Run.Marker;
        (string? text, string? fragment) = Body(token, html: true);
        Q149Attempt productAttempt = new(Q149Cells.C4 + "-product", (forward ? 100 : 0) + i);
        productAttempt.Set("kind", kind);
        DraftOutcome product = ProductDraftWithRetry(
            productAttempt,
            label => forward
                ? Service.ForwardDraft(sourceEntryId, text, to: Run.Hub, display: false, signature: signature.Name, bodyHtml: fragment)
                : Service.ReplyDraft(sourceEntryId, text, replyAll: false, display: false, signature: signature.Name, bodyHtml: fragment),
            "c4-product-" + kind + "-" + i.ToString(CultureInfo.InvariantCulture));
        Run.Write(productAttempt);
        string productHtml = Q149Drafts.Read(Run, product.EntryId, storeId).Html;
        Dictionary<string, string> quote = new(StringComparer.Ordinal) { ["quote"] = quoteToken };
        Q149HtmlFacts today = Q149Html.Analyze(productHtml, quote);

        foreach (HtmlOnlyStyle style in Styles)
        {
            Q149Attempt attempt = new(Q149Cells.C4, (forward ? 100 : 0) + (i * 10) + (int)style);
            attempt.Set("kind", kind).Set("variant", style.ToString())
                .Set("productPlacement", product.BodyPlacement).Set("productComposeError", product.ComposeSurfaceError);

            // What the object model hands over with no inspector at all - the C4 question itself.
            (string derived, string derivedHtml) = Q149Drafts.CreateDerivedWithoutInspector(
                Run, client, sourceEntryId, storeId, forward, "c4-" + kind + "-" + style.ToString() + "-" + i.ToString(CultureInfo.InvariantCulture));
            Q149HtmlFacts bare = Q149Html.Analyze(derivedHtml, quote);
            attempt.Set("noInspector", bare.ToFields()).Set("today", today.ToFields())
                .Set("noInspectorOutline", Q149Html.Outline(derivedHtml, quote));
            if (style == HtmlOnlyStyle.Message)
            {
                bool quoteHandedOver = bare.TokenAt["quote"] >= 0 && bare.WordSectionAt >= 0 && (bare.HeaderBlockAt >= 0 || bare.MailOriginalAt >= 0);
                Record(Q149Cells.C4, "the quoted original without an inspector", quoteHandedOver);
                Record(Q149Cells.C4, "the stylesheet of a " + kind + " without an inspector", bare.StylesheetHash != null && bare.StylesheetHash == today.StylesheetHash);
            }

            string cid = "image001.png@" + Run.Marker + kind + style.ToString() + i.ToString(CultureInfo.InvariantCulture);
            string signatureBlock = HtmlOnlyDraftComposer.SignatureBlock(File.ReadAllText(signature.FilePath), src => SignaturePicture(src, cid), style);
            Q149HtmlFacts placed = Q149Drafts.RewriteHtml(
                Run, client, derived, storeId,
                html => HtmlOnlyDraftComposer.InsertAboveQuote(html, HtmlOnlyDraftComposer.ParagraphsFromHtml(fragment!, style), signatureBlock),
                Pictures(signature, cid));
            attempt.Set("htmlOnlyPlaced", placed.ToFields());
            Compare(attempt, client, storeId, product.EntryId, derived, token, kind + ", " + style.ToString());
            Run.Write(attempt);
        }
    }

    // ------------------------------------------------------------------ C3: update_draft

    private void UpdateSet(Q149OutlookClient client, string storeId, TestSignature signature, int i)
    {
        Q149Attempt attempt = new(Q149Cells.C3, i);
        string token = "Q149BODY" + "upd" + i.ToString(CultureInfo.InvariantCulture) + Run.Marker;
        string revised = "Q149REVISED" + i.ToString(CultureInfo.InvariantCulture) + Run.Marker;
        (string? text, _) = Body(token, html: false);
        string revisedText = "Herzien " + revised + "\nTweede regel.";

        string wordDraft = ProductDraft(signature, text!, "c3-word-" + i.ToString(CultureInfo.InvariantCulture));
        string wordTwin = ProductDraft(signature, text!, "c3-wordtwin-" + i.ToString(CultureInfo.InvariantCulture));
        string htmlOnly = HtmlOnlyTwin(client, storeId, signature, wordDraft, text!, "c3-htmlonly-" + i.ToString(CultureInfo.InvariantCulture), "a" + i.ToString(CultureInfo.InvariantCulture));
        string htmlOnlyTwin = HtmlOnlyTwin(client, storeId, signature, wordDraft, text!, "c3-htmlonlytwin-" + i.ToString(CultureInfo.InvariantCulture), "b" + i.ToString(CultureInfo.InvariantCulture));

        // Word's update (today) on a Word draft: the reference.
        LiveStoreWriteGuard.Assert(Run.Hub, StoreWriteKind.Draft, "q149 C3 update_draft");
        UpdateDraftOutcome viaWord = Service.UpdateDraft(wordDraft, body: revisedText, display: false);
        Run.Allow(viaWord.EntryId);

        // (c)'s update on a Word draft, on an HTML-only draft; and today's update on an HTML-only draft.
        Q149HtmlFacts htmlOnWord = Q149Drafts.RewriteHtml(Run, client, wordTwin, storeId, h => HtmlOnlyDraftComposer.ReplaceDraftRegion(h, HtmlOnlyDraftComposer.ParagraphsFromText(revisedText)));
        Q149HtmlFacts htmlOnHtml = Q149Drafts.RewriteHtml(Run, client, htmlOnly, storeId, h => HtmlOnlyDraftComposer.ReplaceDraftRegion(h, HtmlOnlyDraftComposer.ParagraphsFromText(revisedText)));
        UpdateDraftOutcome wordOnHtml = Service.UpdateDraft(htmlOnlyTwin, body: revisedText, display: false);
        Run.Allow(wordOnHtml.EntryId);
        attempt.Set("wordUpdateOnHtmlOnly", wordOnHtml.Status + "/" + wordOnHtml.BodyPlacement)
            .Set("htmlUpdateOnWord", htmlOnWord.ToFields())
            .Set("htmlUpdateOnHtmlOnly", htmlOnHtml.ToFields());

        Dictionary<string, string> tokens = new(StringComparer.Ordinal) { ["old"] = token, ["new"] = revised, ["sig"] = SignatureWord };
        Q149HtmlFacts reference = Q149Html.Analyze(Q149Drafts.Read(Run, viaWord.EntryId, storeId).Html, tokens);
        attempt.Set("wordUpdateOnWord", reference.ToFields());
        foreach ((string label, string entryId) in new[] { ("html update on a Word draft", wordTwin), ("html update on an HTML-only draft", htmlOnly), ("Word update on an HTML-only draft", wordOnHtml.EntryId) })
        {
            (string html, _, _, _, int attachments, long bytes, int hidden) = Q149Drafts.Read(Run, entryId, storeId);
            Q149HtmlFacts facts = Q149Html.Analyze(html, tokens);
            bool replaced = facts.TokenAt["new"] >= 0 && facts.TokenAt["old"] < 0 && facts.TokenAt["sig"] > facts.TokenAt["new"] && facts.CidImages >= 1 && facts.FileLinks == 0;
            (int differences, IReadOnlyList<string> listed) = LookDifference(client, storeId, viaWord.EntryId, entryId);
            attempt.Set(Key(label), new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["replaced"] = replaced,
                ["attachments"] = attachments,
                ["attachmentBytes"] = bytes,
                ["hiddenAttachments"] = hidden,
                ["lookDifferences"] = differences,
                ["firstDifferences"] = listed.ToArray(),
                ["html"] = facts.ToFields(),
            });
            Record(Q149Cells.C3, "update_draft: " + label, replaced && differences == 0);
        }

        Run.Write(attempt);
    }

    // ------------------------------------------------------------------ C5: as received

    private void ReceivedPair(Q149OutlookClient client, string storeId, (string Product, string Message, string AsImported, string Token) pair, int i)
    {
        Q149Attempt attempt = new(Q149Cells.C5, i);
        Dictionary<string, string> received = new(StringComparer.Ordinal);
        foreach ((string label, string entryId) in new[] { ("product", pair.Product), ("Message", pair.Message), ("AsImported", pair.AsImported) })
        {
            (_, IReadOnlyList<string> recipients, _, _, _, _, _) = Q149Drafts.Read(Run, entryId, storeId);
            Q149Drafts.Refuse(Q149ProbeGuards.RefuseRecipients(recipients, Run.HubAddress));
            string subject = SubjectOf(client, entryId, storeId);
            LiveStoreWriteGuard.Assert(Run.Hub, StoreWriteKind.Send, "q149 C5 send to self");
            SendOutcome issued = Service.Send(entryId);
            Assert.False(string.IsNullOrEmpty(issued.ConfirmToken), "Q149: the product issued no send token for the " + label + " draft");
            DateTime sentUtc = DateTime.UtcNow;
            SendOutcome sent = Service.Send(entryId, issued.ConfirmToken);
            attempt.Set(label + "Sent", sent.Sent);
            using OutlookComSession session = OutlookComSession.Connect(allowStartingOutlook: false, onOutlookGone: null, lifetimePin: false);
            ComMailBrief arrived = LiveInboxArrival.WaitFor(session, Run.Hub, subject, sentUtc.AddSeconds(-5));
            Run.Allow(arrived.EntryId);
            received[label] = arrived.EntryId;
            (string html, _, _, _, int attachments, long bytes, int hidden) = Q149Drafts.Read(Run, arrived.EntryId, storeId);
            Dictionary<string, string> tokens = new(StringComparer.Ordinal) { ["body"] = pair.Token, ["sig"] = SignatureWord };
            attempt.Set(label + "Received", new Dictionary<string, object>(StringComparer.Ordinal)
            {
                ["attachments"] = attachments,
                ["attachmentBytes"] = bytes,
                ["hiddenAttachments"] = hidden,
                ["html"] = Q149Html.Analyze(html, tokens).ToFields(),
                ["outline"] = Q149Html.Outline(html, tokens),
            });
        }

        Q149HtmlFacts productFacts = Q149Html.Analyze(Q149Drafts.Read(Run, received["product"], storeId).Html);
        foreach (string variant in new[] { "Message", "AsImported" })
        {
            Q149HtmlFacts variantFacts = Q149Html.Analyze(Q149Drafts.Read(Run, received[variant], storeId).Html);
            (int differences, IReadOnlyList<string> listed) = LookDifference(client, storeId, received["product"], received[variant]);
            attempt.Set("receivedLookDifferences" + variant, differences).Set("receivedFirstDifferences" + variant, listed.ToArray());
            Record(Q149Cells.C5, "looks the same as received (" + variant + ")", differences == 0);
            Record(Q149Cells.C5, "pictures as received (" + variant + ")", productFacts.CidImages >= 1 && productFacts.CidImages == variantFacts.CidImages && variantFacts.FileLinks == 0);
            Record(Q149Cells.C5, "stylesheet as received (" + variant + ")", productFacts.StylesheetHash != null && productFacts.StylesheetHash == variantFacts.StylesheetHash);
        }

        Run.Write(attempt);
    }

    // ------------------------------------------------------------------ shared

    /// <summary>The two drafts of a pair, compared on their HTML, their attachments and Word's layout of each.</summary>
    private void Compare(Q149Attempt attempt, Q149OutlookClient client, string storeId, string product, string htmlOnly, string token, string what)
    {
        Dictionary<string, string> tokens = new(StringComparer.Ordinal) { ["body"] = token, ["sig"] = SignatureWord };
        var p = Q149Drafts.Read(Run, product, storeId);
        var h = Q149Drafts.Read(Run, htmlOnly, storeId);
        Q149HtmlFacts pf = Q149Html.Analyze(p.Html, tokens);
        Q149HtmlFacts hf = Q149Html.Analyze(h.Html, tokens);
        attempt.Set("product", pf.ToFields())
            .Set("htmlOnly", hf.ToFields())
            .Set("productOutline", Q149Html.Outline(p.Html, tokens))
            .Set("htmlOnlyOutline", Q149Html.Outline(h.Html, tokens))
            .Set("productAttachments", new[] { p.Attachments, (int)Math.Min(int.MaxValue, p.AttachmentBytes), p.HiddenAttachments })
            .Set("htmlOnlyAttachments", new[] { h.Attachments, (int)Math.Min(int.MaxValue, h.AttachmentBytes), h.HiddenAttachments });

        // Each aspect counts as the same only when today's draft has it AND the HTML-only one does: two drafts
        // that both lack a signature are not evidence that (c) places one as Word does.
        bool productSignature = SignatureWhole(pf, p.HiddenAttachments, p.AttachmentBytes);
        bool htmlOnlySignature = SignatureWhole(hf, h.HiddenAttachments, h.AttachmentBytes);
        attempt.Set("productBodyInsideWordSection1", Inside(pf)).Set("htmlOnlyBodyInsideWordSection1", Inside(hf))
            .Set("productSignatureWhole", productSignature).Set("htmlOnlySignatureWhole", htmlOnlySignature);
        Record(Q149Cells.C2, "the stylesheet (" + what + ")", pf.StylesheetHash != null && pf.StylesheetHash == hf.StylesheetHash);
        Record(Q149Cells.C2, "the body inside WordSection1 (" + what + ")", Inside(pf) && Inside(hf));
        Record(Q149Cells.C2, "the signature with its pictures (" + what + ")", productSignature && htmlOnlySignature);
        if (what is "reply" or "forward")
        {
            bool productQuote = pf.HeaderBlockAt > pf.TokenAt["body"] || pf.MailOriginalAt > pf.TokenAt["body"];
            bool htmlOnlyQuote = hf.HeaderBlockAt > hf.TokenAt["body"] || hf.MailOriginalAt > hf.TokenAt["body"];
            Record(Q149Cells.C4, "the body above the quoted original (" + what + ")", productQuote && htmlOnlyQuote);
        }

        (int differences, IReadOnlyList<string> listed) = LookDifference(client, storeId, product, htmlOnly);
        attempt.Set("lookDifferences", differences).Set("firstDifferences", listed.ToArray());
        Record(Q149Cells.C2, "looks the same opened in Outlook (" + what + ")", differences == 0);
    }

    private (int Count, IReadOnlyList<string> Listed) LookDifference(Q149OutlookClient client, string storeId, string expected, string actual)
    {
        (List<WordParagraphLook>? e, string? eFailure) = Q149Editor.ReadLookOf(Run, client, expected, storeId);
        (List<WordParagraphLook>? a, string? aFailure) = Q149Editor.ReadLookOf(Run, client, actual, storeId);
        if (e == null || a == null)
        {
            return (-1, new[] { "layout not read: " + (eFailure ?? aFailure ?? "unknown") });
        }

        return Q149WordLook.Compare(e, a);
    }

    private static bool Inside(Q149HtmlFacts facts)
    {
        return facts.WordSectionAt >= 0 && facts.TokenAt["body"] > facts.WordSectionAt;
    }

    private static bool SignatureWhole(Q149HtmlFacts facts, int hiddenAttachments, long attachmentBytes)
    {
        int body = facts.TokenAt["body"];
        return facts.MailAutoSigAt > body
            && facts.TokenAt["sig"] > facts.MailAutoSigAt
            && facts.CidImages >= 1
            && facts.FileLinks == 0
            && hiddenAttachments >= 1
            && attachmentBytes > 0;
    }

    private void Record(string cell, string aspect, bool same)
    {
        _aspects.Add((aspect, same));
        Run.Log(cell + " " + aspect + ": " + (same ? "same" : "DIFFERENT"));
    }

    private int _productDrafts;

    private string ProductDraft(TestSignature signature, string text, string label)
    {
        Q149Attempt attempt = new(Q149Cells.C3 + "-product", ++_productDrafts);
        attempt.Set("label", label);
        DraftOutcome draft = ProductDraftWithRetry(
            attempt,
            l => Service.NewDraft(Run.Hub, Run.Hub, cc: null, Run.TaggedSubject(l), text, display: false, signature: signature.Name),
            label);
        Run.Write(attempt);
        return draft.EntryId;
    }

    private string HtmlOnlyTwin(Q149OutlookClient client, string storeId, TestSignature signature, string frameFrom, string text, string label, string cidSuffix)
    {
        string reference = Q149Drafts.Read(Run, frameFrom, storeId).Html;
        (string, string, string)? frame = HtmlOnlyDraftComposer.ReadFrame(reference);
        Assert.True(frame != null, "Q149: today's draft has no frame to take the stylesheet from");
        string cid = "image001.png@" + Run.Marker + cidSuffix;
        string signatureBlock = HtmlOnlyDraftComposer.SignatureBlock(File.ReadAllText(signature.FilePath), src => SignaturePicture(src, cid));
        return Q149Drafts.CreateHtmlOnlyDraft(Run, client, label, Pictures(signature, cid),
            _ => HtmlOnlyDraftComposer.ComposeNew(frame!.Value, HtmlOnlyDraftComposer.ParagraphsFromText(text), signatureBlock)).EntryId;
    }

    private static (string? Text, string? Html) Body(string token, bool html)
    {
        if (!html)
        {
            return ("Regel een " + token + "\nRegel twee.\n\nAlinea twee.", null);
        }

        string raw = "<h1>Kwartaal" + token + "</h1><p>Geachte <strong>heer</strong>,</p><p>Tekst " + token
            + " hier.</p><table><tr><th>Post</th></tr><tr><td>Cel" + token + "</td></tr></table>";
        return (null, HtmlFragmentNormalizer.Normalize(raw).Html);
    }

    private static string? SignaturePicture(string src, string cid)
    {
        return src.EndsWith("sigimg.png", StringComparison.OrdinalIgnoreCase) ? cid : null;
    }

    private static IReadOnlyList<(string Path, string Cid, string Mime)> Pictures(TestSignature signature, string cid)
    {
        string picture = Path.Combine(Path.GetDirectoryName(signature.FilePath)!, signature.Name + "_files", "sigimg.png");
        return new[] { (picture, cid, "image/png") };
    }

    private string SubjectOf(Q149OutlookClient client, string entryId, string storeId)
    {
        return client.Run((app, ns) =>
        {
            dynamic? item = null;
            try
            {
                item = ns.GetItemFromID(entryId, storeId);
                string subject = (string?)item.Subject ?? string.Empty;
                Q149Drafts.Refuse(Q149ProbeGuards.RefuseSubject(subject, Run.Marker));
                return subject;
            }
            finally
            {
                Q149Com.Release(item);
            }
        });
    }

    private static string Key(string label)
    {
        return label.Replace(" ", "_", StringComparison.Ordinal).Replace("-", "_", StringComparison.Ordinal);
    }
}
