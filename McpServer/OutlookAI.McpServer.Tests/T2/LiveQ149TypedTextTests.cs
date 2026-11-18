using System.Diagnostics;
using System.Globalization;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Q211 (b), decided 2026-11-18: <b>what Outlook's own compose window makes of typed text</b> - the paragraph style
/// and spacing - in a new mail, a reply and a forward, beside the product's drafts of the same text. Q149 found
/// (2026-11-17) that the product's body paragraphs are Word's "Normal (Web)" style with 5 pt before and after, where
/// a paragraph of the message's own <c>MsoNormal</c> class is "Normal" with none; what a person typing in Outlook gets
/// was not measured. This measures it, and reports; the maintainer decides what to change after.
/// <para>
/// <b>How "typed" is had.</b> A draft is saved in the hub with no inspector (Items.Add, or Reply() and Forward() of a
/// mail this run sent itself), re-opened by its EntryID and DISPLAYED - the compose window a person types into - and
/// the text goes in at the insertion point Outlook gave the window, through Word's own typing
/// (<c>Selection.TypeText</c>, <c>TypeParagraph</c>): the formatting a keystroke there would get, with no input
/// synthesized (the suite's user-idle checks rely on there being none). Word's layout is read paragraph by paragraph
/// and the window closed DISCARDING, so the typed text is never saved; the draft is then deleted by EntryID. The
/// product's drafts of the same text - with no signature, and with the test signature - are read the same way.
/// </para>
/// <para>
/// <b>Q180 (b).</b> When no Outlook with a window runs, this class starts one, and that cold start is logged: the time
/// until Outlook answered, its liveness verdicts and any hang (<see cref="Q180ColdStartWatch"/>). One that runs is used
/// as it is. The first version stopped it and started another, and both runs of it were held to their time limit with
/// nothing to say where; so every step of a trial is bounded now (<see cref="Step{T}"/>), and a blocked one names
/// the windows Outlook showed.
/// </para>
/// </summary>
[Collection(LiveCollections.ComposeSurface)]
[Trait("Category", "Live")]
public sealed class LiveQ149TypedTextTests
{
    private const int Rounds = 2;

    private static readonly string[] Kinds = { "new", "reply", "forward" };

    /// <summary>How long one step of a trial may take before Outlook is taken for blocked by a window that waits for a person.</summary>
    private static readonly TimeSpan StepLimit = TimeSpan.FromSeconds(90);

    /// <summary>Set once a step stayed blocked: nothing more is asked of Outlook in this class, not even a delete.</summary>
    private bool _blocked;

    private readonly LiveQ149Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveQ149TypedTextTests(LiveQ149Fixture fixture, ITestOutputHelper output)
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
    public void Q211_WhatOutlooksOwnComposeWindowMakesOfTypedText()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        TypedText(Rounds);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "InteractiveDesktop")]
    [Trait("Requires", "Transport")]
    public void Smoke_TypedTextOnce()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        TypedText(1);
    }

    private void TypedText(int rounds)
    {
        // Q180 (b): a windowed start this class makes is a cold start too, and is logged; an Outlook with a window that
        // runs already - the runner's - is used as it is.
        _fixture.ReleaseService();
        _blocked = false;
        Q149Attempt setup = new(Q149Cells.Q211 + "-process", 1);
        Q180ColdStartWatch start = new("windowed", Q149Cells.Q211, 1);
        try
        {
            LiveComposeProbe.EnsureWindowedOutlook(Run.Settings, Run.Log, start);
        }
        finally
        {
            start.Finish(Run);
        }

        using TestSignature signature = TestSignature.Create(Run.Marker);
        Q149OutlookClient? client = null;
        List<string> rows = new();
        string? blocked = null;
        try
        {
            client = Step("connect a client", setup, () => Q149OutlookClient.Connect(new Stopwatch(), allowStart: false));
            string storeId = Step("read the hub's StoreID", setup, () => Q149Drafts.HubStoreId(Run, client));
            string sourceEntryId = Step("seed the mail to reply to", setup, () => Q149Drafts.SeedSelfSentSource(Run, "q211-source", out _), TimeSpan.FromMinutes(4));
            for (int round = 1; round <= rounds; round++)
            {
                for (int k = 0; k < Kinds.Length; k++)
                {
                    string kind = Kinds[k];
                    rows.Add(Typed(client, storeId, sourceEntryId, kind, (round * 10) + k));
                    rows.Add(Product(client, storeId, sourceEntryId, kind, (round * 10) + k, signatureName: null));
                    rows.Add(Product(client, storeId, sourceEntryId, kind, (round * 10) + k, signature.Name));
                }
            }
        }
        catch (TimeoutException ex)
        {
            blocked = ex.Message;
        }
        finally
        {
            // A client whose thread is still blocked in Outlook is left as it is: its release would wait on it too.
            if (!_blocked)
            {
                client?.Dispose();
            }

            Run.Write(setup);
        }

        string line = Q149Cells.SummaryPrefix + Q149Cells.Q211 + " n=" + rounds.ToString(CultureInfo.InvariantCulture) + ": " + string.Join("; ", rows.Distinct(StringComparer.Ordinal))
            + (blocked == null ? string.Empty : "; BLOCKED: " + blocked);
        _output.WriteLine(line);
        Console.WriteLine(line);
        Assert.True(blocked == null, blocked);
        Assert.Equal(rounds * Kinds.Length * 3, rows.Count);
    }

    /// <summary>
    /// Runs one step of a trial, bounded. Past <see cref="StepLimit"/> - or <paramref name="limit"/> - Outlook is taken for
    /// blocked by a window that waits for a person: every window it shows goes on <paramref name="attempt"/> and the
    /// log, each dialog among them is asked to close (WM_CLOSE: its Cancel, never its OK), and the step gets as long
    /// again. One that still has not returned ends the class's trials, and nothing more is asked of Outlook.
    /// </summary>
    private T Step<T>(string name, Q149Attempt attempt, Func<T> work, TimeSpan? limit = null)
    {
        TimeSpan budget = limit ?? StepLimit;
        Task<T> task = Task.Run(work);
        if (task.Wait(budget))
        {
            return task.GetAwaiter().GetResult();
        }

        string[] windows = Q149Native.DescribeVisibleOutlookWindows();
        int asked = Q149Native.CloseVisibleOutlookDialogs();
        attempt.Set("blockedStep", name).Set("blockedBy", windows).Set("dialogsAskedToClose", asked);
        Run.Log("Q211: '" + name + "' had not returned after " + budget.TotalSeconds.ToString("0", CultureInfo.InvariantCulture) + " s; Outlook shows "
            + (windows.Length == 0 ? "no window" : string.Join("; ", windows)) + "; dialogs asked to close: " + asked.ToString(CultureInfo.InvariantCulture));
        if (task.Wait(budget))
        {
            attempt.Set("returnedAfterTheDialogsClosed", true);
            return task.GetAwaiter().GetResult();
        }

        _blocked = true;
        throw new TimeoutException("Q211: '" + name + "' did not return, even after Outlook's dialogs were asked to close; Outlook showed "
            + (windows.Length == 0 ? "no window" : string.Join("; ", windows)));
    }

    /// <summary>Outlook's own compose window, typed into: the draft saved with no inspector, re-opened, displayed, typed, read, discarded, deleted.</summary>
    private string Typed(Q149OutlookClient client, string storeId, string sourceEntryId, string kind, int iteration)
    {
        Q149Attempt attempt = new(Q149Cells.Q211, iteration);
        attempt.Set("kind", kind).Set("source", "outlook-typed");
        string token = "Q211TYPED" + kind.ToUpperInvariant() + iteration.ToString(CultureInfo.InvariantCulture) + Run.Marker;
        string label = "q211-typed-" + kind + "-" + iteration.ToString(CultureInfo.InvariantCulture);
        string entryId;
        try
        {
            entryId = Step("save the " + kind + " draft", attempt, () => kind == "new"
                ? Q149Drafts.CreateTaggedDraft(Run, client, label, htmlBody: null, addressToHub: true).EntryId
                : Q149Drafts.CreateDerivedWithoutInspector(Run, client, sourceEntryId, storeId, kind == "forward", label).EntryId);
        }
        catch (TimeoutException)
        {
            Run.Write(attempt);
            throw;
        }

        try
        {
            (List<WordParagraphLook>? look, int caret, string? failure) = Step(
                "type into the " + kind + " compose window",
                attempt,
                () => client.Run<(List<WordParagraphLook>?, int, string?)>((_, ns) => TypeIntoTheComposeWindow((object)ns, entryId, storeId, token)));
            attempt.Set("caretAt", caret).Set("failure", failure);
            return Record(attempt, look, token, "outlook typed " + kind);
        }
        finally
        {
            if (!_blocked)
            {
                _ = LiveOutlookTestMailer.DeleteItemByEntryId(Run.Hub, entryId, Run.Marker);
            }

            Run.Write(attempt);
        }
    }

    /// <summary>The product's draft of the same text - a text body, as an agent writes one - read the same way, then deleted.</summary>
    private string Product(Q149OutlookClient client, string storeId, string sourceEntryId, string kind, int iteration, string? signatureName)
    {
        string source = signatureName == null ? "product-no-signature" : "product-signature";
        Q149Attempt attempt = new(Q149Cells.Q211 + "-product", iteration * 10 + (signatureName == null ? 1 : 2));
        attempt.Set("kind", kind).Set("source", source);
        string token = "Q211PRODUCT" + kind.ToUpperInvariant() + iteration.ToString(CultureInfo.InvariantCulture) + (signatureName == null ? "N" : "S") + Run.Marker;
        string text = "Typed line one " + token + "\nTyped line two\n\nTyped paragraph two";
        string requested = signatureName ?? "none";
        LiveStoreWriteGuard.Assert(Run.Hub, StoreWriteKind.Draft, "q211 product draft");
        DraftOutcome draft;
        try
        {
            draft = Step("the product's " + kind + " draft (" + requested + ")", attempt, () => kind switch
            {
                "new" => Service.NewDraft(Run.Hub, Run.Hub, cc: null, Run.TaggedSubject("q211-product-" + kind + "-" + iteration.ToString(CultureInfo.InvariantCulture)), text, display: false, signature: requested),
                "reply" => Service.ReplyDraft(sourceEntryId, text, replyAll: false, display: false, signature: requested),
                _ => Service.ForwardDraft(sourceEntryId, text, to: Run.Hub, display: false, signature: requested),
            }, LiveComposeProbe.StepBudget);
        }
        catch (TimeoutException)
        {
            Run.Write(attempt);
            throw;
        }

        Run.Allow(draft.EntryId);
        try
        {
            attempt.Set("placement", draft.BodyPlacement).Set("signatureApplied", draft.SignatureApplied).Set("composeError", draft.ComposeSurfaceError);
            (List<WordParagraphLook>? look, string? failure) = Step(
                "read the product's " + kind + " draft",
                attempt,
                () => Q149Editor.ReadLookOf(Run, client, draft.EntryId, storeId));
            attempt.Set("failure", failure);
            return Record(attempt, look, token, source.Replace('-', ' ') + " " + kind);
        }
        finally
        {
            if (!_blocked)
            {
                _ = LiveOutlookTestMailer.DeleteItemByEntryId(Run.Hub, draft.EntryId, Run.Marker);
            }

            Run.Write(attempt);
        }
    }

    /// <summary>
    /// On the client's STA: opens the draft in a compose window a person would see, types at the insertion point Outlook
    /// gave it, reads Word's layout, and closes the window discarding. Returns the layout, where the insertion point was,
    /// and the failure when the editor could not be had.
    /// </summary>
    private (List<WordParagraphLook>? Look, int Caret, string? Failure) TypeIntoTheComposeWindow(object nameSpace, string entryId, string storeId, string token)
    {
        dynamic ns = nameSpace;
        dynamic? item = null;
        object? inspector = null;
        object? document = null;
        dynamic? windows = null;
        dynamic? window = null;
        dynamic? selection = null;
        try
        {
            item = ns.GetItemFromID(entryId, storeId);
            Q149Drafts.RequireOwnUnsentDraft(Run, item);
            item.Display(false);
            inspector = (object)item.GetInspector;
            Q149EditorRead read = Q149Editor.ReadEditor(inspector, 10_000, 250);
            if (read.Document == null)
            {
                return (null, -1, "no editor in the displayed compose window");
            }

            document = read.Document;
            windows = ((dynamic)document).Windows;
            window = windows.Item(1);
            selection = window.Selection;
            int caret = (int)selection.Start;
            selection.TypeText("Typed line one " + token);
            selection.TypeParagraph();
            selection.TypeText("Typed line two");
            selection.TypeParagraph();
            selection.TypeParagraph();
            selection.TypeText("Typed paragraph two");
            return (Q149Editor.ReadLook(document), caret, null);
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
        {
            return (null, -1, OutlookComSession.DescribeComFailure(ex));
        }
        finally
        {
            Q149Com.Release(selection);
            Q149Com.Release(window);
            Q149Com.Release(windows);
            if (inspector != null)
            {
                try
                {
                    ((dynamic)inspector).Close(LiveComposeProbe.OlDiscard);
                }
                catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                {
                }
            }

            Q149Com.Release(document);
            Q149Com.Release(inspector);
            Q149Com.Release(item);
        }
    }

    /// <summary>
    /// Records the layout on <paramref name="attempt"/> - every paragraph down to the quoted header block, the paragraph
    /// holding <paramref name="token"/> on its own - and returns the row the summary line carries.
    /// </summary>
    private static string Record(Q149Attempt attempt, List<WordParagraphLook>? look, string token, string what)
    {
        if (look == null)
        {
            return what + ": not read (" + (attempt.Get("failure") as string ?? "unknown") + ")";
        }

        int at = look.FindIndex(p => p.Text.Contains(token, StringComparison.Ordinal));
        int header = look.FindIndex(p => p.TopBorder > 0);
        int shown = Math.Min(look.Count, header > 0 ? header + 1 : 8);
        attempt.Set("paragraphs", look.Count)
            .Set("tokenParagraph", at + 1)
            .Set("headerBlockParagraph", header + 1)
            .Set("layout", look.Take(shown).Select((p, i) => "p" + (i + 1).ToString(CultureInfo.InvariantCulture) + " " + p.Describe()).ToArray());
        if (at < 0)
        {
            return what + ": the text is in no paragraph";
        }

        WordParagraphLook typed = look[at];
        attempt.Set("style", typed.Style)
            .Set("spaceBefore", typed.SpaceBefore)
            .Set("spaceAfter", typed.SpaceAfter)
            .Set("lineSpacing", typed.LineSpacing)
            .Set("font", typed.Font)
            .Set("size", typed.Size)
            .Set("topBorder", typed.TopBorder);
        return string.Format(
            CultureInfo.InvariantCulture,
            "{0}: p{1} '{2}' {3}/{4} pt, line {5}, {6} {7} pt",
            what,
            at + 1,
            typed.Style,
            typed.SpaceBefore,
            typed.SpaceAfter,
            typed.LineSpacing,
            typed.Font,
            typed.Size);
    }
}
