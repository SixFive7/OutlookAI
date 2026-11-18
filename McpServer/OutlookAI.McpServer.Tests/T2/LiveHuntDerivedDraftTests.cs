using System.Diagnostics;
using System.Globalization;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// THE HUNT'S VARIANTS (runbook section 4.15; the coordinator's direction of 2026-11-18, option (b)): which part of
/// making a reply or a forward the crash of OUTLOOK.EXE in mso40uiwin32client.dll needs. Seven crashes that day came
/// while the product made a reply or a forward in an Outlook with its main window up - at the creator's call, or at
/// the "Message" form page of the re-opened draft - and the new drafts the same tests made never crashed. Each variant
/// makes derived drafts in a loop for the leg's minutes, or until Outlook ends, a reply and a forward by turns from one
/// mail the leg sent the hub itself, and writes one attempt line a draft, so a run says how many came before a crash:
/// <list type="bullet">
/// <item>ProductLoop, the control: the product's reply_draft and forward_draft in the runner's windowed Outlook, its
/// main window as Outlook opened it, on the Inbox the source mail is in.</item>
/// <item>OtherFolder: the same, the main window moved to the Calendar first, so no mail of the Inbox is on screen.</item>
/// <item>NoWindow: the same in an Outlook that COM started, with no window - the windowed one ended first, after the
/// source was sent. The product promotes that Outlook's first draft and reads the rest directly.</item>
/// <item>NoFormPage: a reply or forward made by the object model and saved, re-opened by its EntryID, its editor read
/// and edited and the inspector closed with olSave - the product's direct route - without the "Message" form page,
/// which is what keeps such an edit (B3; without it, B2, the edit is not kept - nothing here needs it kept).</item>
/// <item>BareCreators: Reply() or Forward() on the source, and what it returned released unsaved.</item>
/// </list>
/// With <see cref="HuntPlan.ShowsEach"/> every draft comes after a compose window shown and closed - a reply or forward
/// the object model made, promoted the product's own way and closed discarding - as the typed-text tests do before each
/// of the product's drafts; the loops without it make no window at all.
/// Gated twice: it runs only in a run that selects the compose collection alone (<see cref="Q149RunGate"/>), and only
/// when the hunt's suite script named a variant (<see cref="HuntPlan.VariantVariable"/>); anywhere else it writes a
/// NOT RUN line and makes nothing. Every draft is the hub's, carries the tag and this run's marker, and is deleted by
/// its EntryID through the tested helper as soon as it is made; the fixture's teardown sweeps whatever is left.
/// </summary>
[Collection(LiveCollections.ComposeSurface)]
[Trait("Category", "Live")]
public sealed class LiveHuntDerivedDraftTests
{
    /// <summary>olFolderCalendar.</summary>
    private const int CalendarFolder = 9;

    /// <summary>How long a draft that failed the way a dying Outlook fails waits for the process to end - a debugger
    /// writing its full dump holds Outlook still first, for tens of seconds.</summary>
    private static readonly TimeSpan GoingWait = TimeSpan.FromSeconds(90);

    /// <summary>How long any other failed draft looks for the process to end before the loop goes on.</summary>
    private static readonly TimeSpan FailureWait = TimeSpan.FromSeconds(3);

    private readonly LiveQ149Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveHuntDerivedDraftTests(LiveQ149Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private Q149Run Run => _fixture.Run;

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "InteractiveDesktop")]
    [Trait("Requires", "Transport")]
    public void Hunt_DerivedDraftsUntilTheLegEndsOrOutlookDoes()
    {
        HuntPlan? plan = HuntPlan.Parse(
            Environment.GetEnvironmentVariable(HuntPlan.VariantVariable),
            Environment.GetEnvironmentVariable(HuntPlan.LegMinutesVariable),
            Environment.GetEnvironmentVariable(HuntPlan.ShowEachVariable),
            out string? why);
        if (plan == null)
        {
            _output.WriteLine("NOT RUN: " + why + ".");
            return;
        }

        if (!_fixture.Begin(_output))
        {
            return;
        }

        // The source is sent and arrives in an Outlook with its window up - the runner's, or one started the way a
        // user starts it - in every variant: a window-less Outlook is had only after it.
        LiveComposeProbe.EnsureWindowedOutlook(Run.Settings, Run.Log);
        if (plan.Maker == HuntMaker.None)
        {
            // The CalendarOnly step: the main window moved, said, and nothing made.
            using Q149OutlookClient mover = Q149OutlookClient.Connect(new Stopwatch(), allowStart: false);
            Q149Attempt step = new(plan.Cell + "-outlook", 0);
            step.Set("moved", MoveMainWindowToTheCalendar(mover)).Set("mainWindow", DescribeMainWindow(mover)).Set("visibleWindows", WindowProbe.VisibleOutlookWindows().Count);
            Run.Write(step);
            return;
        }

        string sourceEntryId = Q149Drafts.SeedSelfSentSource(Run, "hunt-source", out _);
        if (plan.Windowless)
        {
            _fixture.ReleaseService();
            LiveComposeProbe.StopOutlook(Run.Settings, Run.Log);
        }

        // Window-less, this client's Outlook.Application starts OUTLOOK.EXE with -Embedding, and holds it up for the leg.
        // Released after the asserts, quietly: a release that met the crashed Outlook must not hide what the leg found.
        Q149OutlookClient client = Q149OutlookClient.Connect(new Stopwatch(), allowStart: plan.Windowless);
        try
        {
            Hunt(plan, client, sourceEntryId);
        }
        finally
        {
            try
            {
                client.Dispose();
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex) || LiveOutlookRestart.IsSessionGone(ex))
            {
                Run.Log("hunt: the probe client's release failed (" + ex.GetType().Name + ") - its Outlook had ended");
            }
        }
    }

    /// <summary>One leg: the variant's main window, then drafts until the leg's minutes are up or Outlook ends.</summary>
    private void Hunt(HuntPlan plan, Q149OutlookClient client, string sourceEntryId)
    {
        IReadOnlyList<int> pids = LiveOutlookRestart.OutlookProcessIds();
        Assert.True(pids.Count == 1, "hunt: not exactly one OUTLOOK.EXE before the first draft (" + pids.Count.ToString(CultureInfo.InvariantCulture) + ")");
        int pid = pids[0];
        string storeId = Q149Drafts.HubStoreId(Run, client);

        Q149Attempt start = new(plan.Cell + "-outlook", 0);
        string? commandLine = Q149Native.TryReadCommandLine(pid);
        start.Set("pid", pid)
            .Set("embedding", commandLine == null ? null : commandLine.Contains("Embedding", StringComparison.OrdinalIgnoreCase))
            .Set("legMinutes", plan.LegMinutes)
            .Set("maker", plan.Maker.ToString())
            .Set("showsEach", plan.ShowsEach)
            .Set("typesEach", plan.TypesEach);
        if (plan.ShowsCalendar)
        {
            start.Set("moved", MoveMainWindowToTheCalendar(client));
        }

        start.Set("mainWindow", DescribeMainWindow(client)).Set("visibleWindows", WindowProbe.VisibleOutlookWindows().Count);
        Run.Write(start);

        // The drafts' own allowlist, under the run's marker: the fixture's holds the source alone, so its teardown does
        // not look up every draft this leg already deleted - the tag sweep after it finds any that were left.
        Q149Run drafts = new(Run.Settings, Run.Marker, _output.WriteLine);
        drafts.Allow(sourceEntryId);
        Stopwatch leg = Stopwatch.StartNew();
        TimeSpan budget = TimeSpan.FromMinutes(plan.LegMinutes);
        int made = 0;
        int failed = 0;
        int failedInARow = 0;
        bool alive = true;
        for (int i = 1; leg.Elapsed < budget; i++)
        {
            bool forward = i % 2 == 0;
            Q149Attempt attempt = new(plan.Cell, i);
            attempt.Set("kind", forward ? "forward" : "reply").Set("pid", pid);
            Stopwatch one = Stopwatch.StartNew();
            string? entryId = null;
            string? failure = null;
            bool going = false;
            try
            {
                if (plan.ShowsEach)
                {
                    ShowAndClose(drafts, client, sourceEntryId, storeId, forward, i, attempt, plan.TypesEach);
                }

                entryId = plan.Maker switch
                {
                    HuntMaker.Product => ProductDraft(sourceEntryId, forward, i),
                    HuntMaker.NoFormPage => DirectWithoutTheFormPage(drafts, client, sourceEntryId, storeId, forward, i, attempt),
                    _ => BareCreator(client, sourceEntryId, storeId, forward),
                };
                made++;
                failedInARow = 0;
            }
            catch (Exception ex) when (ex is InvalidOperationException or TimeoutException or AggregateException or ArgumentException or OutlookUnavailableException
                || OutlookComSession.IsComCallFailure(ex) || LiveOutlookRestart.IsSessionGone(ex))
            {
                failed++;
                failedInARow++;
                failure = Describe(ex);
                going = HuntPlan.SaysOutlookIsGoing(failure) || LiveOutlookRestart.IsSessionGone(ex) || ex is OutlookUnavailableException;
                attempt.Set("failure", Q149Com.Failure(ex)).Set("failureSaysOutlookIsGoing", going);
            }

            attempt.Set("ms", one.ElapsedMilliseconds).Set("made", made).Set("atSecond", (long)leg.Elapsed.TotalSeconds);
            if (entryId != null)
            {
                Delete(entryId, attempt);
            }

            alive = failure == null ? Alive(pid) : StaysAlive(pid, going ? GoingWait : FailureWait);
            attempt.Set("outlookAlive", alive).Set("visibleWindows", WindowProbe.VisibleOutlookWindows().Count);
            Run.Write(attempt);
            if (HuntPlan.StopsAfter(alive, failedInARow))
            {
                break;
            }
        }

        Q149Attempt end = new(plan.Cell + "-summary", 0);
        end.Set("made", made).Set("failed", failed).Set("outlookEnded", !alive).Set("seconds", (long)leg.Elapsed.TotalSeconds).Set("pid", pid);
        if (alive)
        {
            end.Set("mainWindow", DescribeMainWindow(client));
        }

        Run.Write(end);
        Assert.True(alive, "hunt: OUTLOOK.EXE " + pid.ToString(CultureInfo.InvariantCulture) + " ended after " + made.ToString(CultureInfo.InvariantCulture)
            + " derived draft(s) of variant " + plan.Variant + " in " + ((long)leg.Elapsed.TotalSeconds).ToString(CultureInfo.InvariantCulture) + " s - the crash the hunt is for; its dumps are the run's");
        Assert.True(failedInARow < HuntPlan.MostFailuresInARow, "hunt: " + failedInARow.ToString(CultureInfo.InvariantCulture) + " drafts in a row failed with Outlook still running - the attempt lines say how");
    }

    /// <summary>The product's own reply_draft or forward_draft of the source, display off.</summary>
    private string ProductDraft(string sourceEntryId, bool forward, int i)
    {
        LiveStoreWriteGuard.Assert(Run.Hub, StoreWriteKind.Draft, forward ? "hunt forward_draft" : "hunt reply_draft");
        string text = "Hunt " + (forward ? "forward " : "reply ") + i.ToString(CultureInfo.InvariantCulture) + " " + Run.Marker;
        DraftOutcome made = forward
            ? _fixture.Service.ForwardDraft(sourceEntryId, text, to: Run.Hub, display: false, signature: null)
            : _fixture.Service.ReplyDraft(sourceEntryId, text, replyAll: false, display: false, signature: null);
        return made.EntryId;
    }

    /// <summary>
    /// NoFormPage: Reply() or Forward() saved with no inspector (<see cref="Q149Drafts.CreateDerivedWithoutInspector"/>),
    /// then re-opened by its EntryID, its inspector never activated, its editor read directly, a line put above the
    /// quote and the inspector closed with olSave - the product's direct route, without <c>ModifiedFormPages.Add</c>.
    /// </summary>
    private string DirectWithoutTheFormPage(Q149Run drafts, Q149OutlookClient client, string sourceEntryId, string storeId, bool forward, int i, Q149Attempt attempt)
    {
        (string entryId, _) = Q149Drafts.CreateDerivedWithoutInspector(drafts, client, sourceEntryId, storeId, forward, "hunt-" + i.ToString(CultureInfo.InvariantCulture));
        string text = "Hunt " + (forward ? "forward " : "reply ") + i.ToString(CultureInfo.InvariantCulture) + " " + Run.Marker;
        LiveStoreWriteGuard.Assert(Run.Hub, StoreWriteKind.Draft, "hunt edit without the form page");
        client.Run<object?>((app, ns) =>
        {
            dynamic? item = null;
            dynamic? inspector = null;
            object? document = null;
            bool closed = false;
            try
            {
                item = ns.GetItemFromID(entryId, storeId);
                Q149Drafts.RequireOwnUnsentDraft(drafts, item);
                inspector = item.GetInspector;
                Q149EditorRead read = Q149Editor.ReadEditor((object)inspector, 2_000, 250);
                document = read.Document;
                attempt.Set("editorOk", read.Ok).Set("editorPolls", read.Polls).Set("editorFailure", read.Failure);
                if (document != null)
                {
                    InsertAtTheStart(document, text);
                }

                inspector.Close(LiveComposeProbe.OlSave);
                closed = true;
                return null;
            }
            finally
            {
                Q149Com.Release(document);
                if (!closed && inspector is not null)
                {
                    try
                    {
                        inspector!.Close(LiveComposeProbe.OlDiscard);
                    }
                    catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                    {
                    }
                }

                Q149Com.Release(inspector);
                Q149Com.Release(item);
            }
        });
        return entryId;
    }

    /// <summary>BareCreators: Reply() or Forward() on the source, and what it returned released unsaved - nothing is written.</summary>
    private string? BareCreator(Q149OutlookClient client, string sourceEntryId, string storeId, bool forward)
    {
        Q149Drafts.Refuse(Q149ProbeGuards.RefuseEntryId(sourceEntryId, Run.Created));
        client.Run<object?>((app, ns) =>
        {
            dynamic? source = null;
            dynamic? derived = null;
            try
            {
                source = ns.GetItemFromID(sourceEntryId, storeId);
                Q149Drafts.Refuse(Q149ProbeGuards.RefuseSubject((string?)source.Subject, Run.Marker));
                derived = forward ? source.Forward() : source.Reply();
                return null;
            }
            finally
            {
                Q149Com.Release(derived);
                Q149Com.Release(source);
            }
        });
        return null;
    }

    /// <summary>
    /// ShowsEach: a reply or forward the object model makes and saves (<see cref="Q149Drafts.CreateDerivedWithoutInspector"/>),
    /// opened in an inspector and promoted the product's own way - parked off screen, activated, hidden
    /// (<see cref="ComposeSurface.PromoteForWordEditor(object, out string?)"/>), as the typed-text tests read every draft -
    /// then closed discarding and deleted. Activating it is what OutlookAI's own add-in puts its task pane on; closing it
    /// is what the add-in takes the pane off at.
    /// </summary>
    private void ShowAndClose(Q149Run drafts, Q149OutlookClient client, string sourceEntryId, string storeId, bool forward, int i, Q149Attempt attempt, bool typed)
    {
        Stopwatch shown = Stopwatch.StartNew();
        (string entryId, _) = Q149Drafts.CreateDerivedWithoutInspector(drafts, client, sourceEntryId, storeId, forward, "hunt-shown-" + i.ToString(CultureInfo.InvariantCulture));
        try
        {
            string? error = client.Run<string?>((app, ns) =>
            {
                dynamic? item = null;
                object? inspector = null;
                object? document = null;
                try
                {
                    item = ns.GetItemFromID(entryId, storeId);
                    Q149Drafts.RequireOwnUnsentDraft(drafts, item);
                    if (typed)
                    {
                        // As the typed-text tests do: the compose window displayed, typed into at Outlook's insertion point.
                        item.Display(false);
                        inspector = (object)item.GetInspector;
                        Q149EditorRead read = Q149Editor.ReadEditor(inspector, 10_000, 250);
                        document = read.Document;
                        if (document == null)
                        {
                            return "no editor in the displayed compose window";
                        }

                        TypeAtTheInsertionPoint(document, "Hunt typed " + i.ToString(CultureInfo.InvariantCulture) + " " + Run.Marker);
                        return null;
                    }

                    inspector = (object)item.GetInspector;
                    document = ComposeSurface.PromoteForWordEditor(inspector, out string? promoteError);
                    return document == null ? (promoteError ?? "NoWordEditor") : null;
                }
                finally
                {
                    Q149Com.Release(document);
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

                    Q149Com.Release(inspector);
                    Q149Com.Release(item);
                }
            });
            attempt.Set("shownMs", shown.ElapsedMilliseconds).Set("shownError", error);
        }
        finally
        {
            try
            {
                attempt.Set("shownDeleted", LiveOutlookTestMailer.DeleteItemByEntryId(Run.Hub, entryId, Run.Marker));
            }
            catch (Exception ex) when (ex is InvalidOperationException || OutlookComSession.IsComCallFailure(ex))
            {
                attempt.Set("shownDeleteFailure", Q149Com.Failure(ex));
            }
        }
    }

    /// <summary>Two lines typed through Word's own typing at the insertion point Outlook gave the displayed window.</summary>
    private static void TypeAtTheInsertionPoint(object document, string text)
    {
        dynamic? windows = null;
        dynamic? window = null;
        dynamic? selection = null;
        try
        {
            windows = ((dynamic)document).Windows;
            window = windows.Item(1);
            selection = window.Selection;
            selection.TypeText(text);
            selection.TypeParagraph();
            selection.TypeText("Typed line two");
        }
        finally
        {
            Q149Com.Release(selection);
            Q149Com.Release(window);
            Q149Com.Release(windows);
        }
    }

    /// <summary>A line above everything in the draft - above the quote, where the product puts its text.</summary>
    private static void InsertAtTheStart(object document, string text)
    {
        dynamic doc = document;
        dynamic? range = null;
        try
        {
            range = doc.Range(0, 0);
            range.InsertBefore(text + "\r");
        }
        finally
        {
            Q149Com.Release(range);
        }
    }

    /// <summary>Deletes one draft of this leg by its EntryID, the tag and the marker (rule 2), through the tested helper.</summary>
    private void Delete(string entryId, Q149Attempt attempt)
    {
        try
        {
            attempt.Set("deleted", LiveOutlookTestMailer.DeleteItemByEntryId(Run.Hub, entryId, Run.Marker));
        }
        catch (Exception ex) when (ex is InvalidOperationException || OutlookComSession.IsComCallFailure(ex))
        {
            attempt.Set("deleteFailure", Q149Com.Failure(ex));
        }
    }

    /// <summary>OtherFolder: the main window moved to the default Calendar. Says what it shows after.</summary>
    private static string MoveMainWindowToTheCalendar(Q149OutlookClient client)
    {
        return client.Run((app, ns) =>
        {
            dynamic? explorer = null;
            dynamic? calendar = null;
            try
            {
                explorer = app.ActiveExplorer();
                if (explorer == null)
                {
                    throw new InvalidOperationException("hunt: Outlook has no main window to move to the Calendar.");
                }

                calendar = ns.GetDefaultFolder(CalendarFolder);
                explorer.CurrentFolder = calendar;
                return "moved";
            }
            finally
            {
                Q149Com.Release(calendar);
                Q149Com.Release(explorer);
            }
        });
    }

    /// <summary>What the main window shows: its folder's name and item type, and how many items are selected - or that there is none.</summary>
    private static string DescribeMainWindow(Q149OutlookClient client)
    {
        try
        {
            return client.Run((app, ns) =>
            {
                dynamic? explorer = null;
                dynamic? folder = null;
                dynamic? selection = null;
                try
                {
                    explorer = app.ActiveExplorer();
                    if (explorer == null)
                    {
                        return "no main window";
                    }

                    folder = explorer.CurrentFolder;
                    string name = folder == null ? "(no folder)" : (string)folder.Name;
                    int itemType = folder == null ? -1 : (int)folder.DefaultItemType;
                    int selected = -1;
                    try
                    {
                        selection = explorer.Selection;
                        selected = (int)selection.Count;
                    }
                    catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                    {
                    }

                    return string.Format(CultureInfo.InvariantCulture, "folder '{0}', item type {1}, {2} selected", name, itemType, selected);
                }
                finally
                {
                    Q149Com.Release(selection);
                    Q149Com.Release(folder);
                    Q149Com.Release(explorer);
                }
            }, TimeSpan.FromSeconds(30));
        }
        catch (Exception ex) when (ex is TimeoutException or InvalidOperationException || OutlookComSession.IsComCallFailure(ex))
        {
            return "not read: " + ex.GetType().Name;
        }
    }

    /// <summary>A failure and every exception inside it, each as its type, HRESULT and message - what the going test reads.</summary>
    private static string Describe(Exception ex)
    {
        List<string> parts = new();
        Stack<Exception> pending = new();
        pending.Push(ex);
        while (pending.Count > 0 && parts.Count < 8)
        {
            Exception e = pending.Pop();
            parts.Add(e.GetType().Name + " 0x" + e.HResult.ToString("X8", CultureInfo.InvariantCulture) + " " + e.Message);
            if (e is AggregateException aggregate)
            {
                foreach (Exception inner in aggregate.InnerExceptions)
                {
                    pending.Push(inner);
                }
            }
            else if (e.InnerException != null)
            {
                pending.Push(e.InnerException);
            }
        }

        return string.Join(" | ", parts);
    }

    private static bool Alive(int pid)
    {
        try
        {
            using Process process = Process.GetProcessById(pid);
            return !process.HasExited;
        }
        catch (Exception ex) when (ex is ArgumentException or InvalidOperationException)
        {
            return false;
        }
    }

    /// <summary>Whether the process is still there after <paramref name="wait"/>, looked at once a second - false as soon as it ends.</summary>
    private static bool StaysAlive(int pid, TimeSpan wait)
    {
        Stopwatch waited = Stopwatch.StartNew();
        while (waited.Elapsed < wait)
        {
            if (!Alive(pid))
            {
                return false;
            }

            Thread.Sleep(1000);
        }

        return Alive(pid);
    }
}
