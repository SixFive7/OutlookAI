using System.Diagnostics;
using System.Globalization;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Q149 (d) stage 1 - an Outlook with no window, on Office LTSC 2024. Every trial ends Outlook the S7 way and
/// cold-starts it over COM (OUTLOOK.EXE -Embedding, no Explorer, no Inspector, nothing on screen), so each
/// trial sees exactly the state an agent session meets when it is the one that starts Outlook.
/// <list type="bullet">
/// <item><b>L1 (decisive)</b>: the shipped compose - <c>ReviseHeldDocument</c>, reached through
/// <c>TryUpdateDraft</c> - in a session with NO lifetime pin. Does closing the never-shown compose window end
/// the window-less Outlook? 20 trials; L1c is the same with the product's pin, as the control.</item>
/// <item><b>L2, W2, W3</b>: an inspector fetched on a re-opened draft and never activated: the editor read
/// once, re-read every 250 ms for 30 s, and re-read from 180 s after the start - then closed, and Outlook's
/// lifetime watched.</item>
/// <item><b>W4, W5</b>: a Windows-level show without Outlook's Activate - of the inspector's own window, and of
/// the pin Explorer's - then the editor read.</item>
/// <item><b>W6</b>: the product's promotion once in a process, then later drafts read and committed with no
/// Activate; then a displayed Explorer opened and closed, and read again.</item>
/// <item><b>W7</b>: W6 with B3's commit on every later draft - <c>ModifiedFormPages.Add("Message")</c> before
/// the read, <c>Close(olSave)</c> after the edit - the one promotion per process tried end to end.</item>
/// <item><b>P1-P3</b> (added 2026-11-18, after Q208 B was decided on the cells above): the PRODUCT as it drafts
/// since - P1 one promotion per Outlook start and every later draft direct, the last session out closing the pin;
/// P2 two sessions sharing one pin, the last out closing the one the other made; P3 a COM-started Outlook with a
/// displayed Explorer, where no draft may activate anything. Every cold start of these cells is also logged for
/// Q180 (b) (<see cref="Q180ColdStartWatch"/>).</item>
/// </list>
/// Why L1 goes through update_draft and not new_draft: new_draft sets the subject only after the compose, so
/// an Outlook that ends at the compose's close would leave an UNTAGGED draft behind (rule 2). The inspector
/// work is the same method either way; update_draft's draft is tagged before it starts.
/// <para>
/// Every attempt is one <see cref="Q149Cells.LinePrefix"/> JSON line. The measurement asserts only what makes
/// it a measurement - a cold start really was one, the trials ran - and records everything else as data.
/// </para>
/// </summary>
[Collection(LiveCollections.ComposeSurface)]
[Trait("Category", "Live")]
public sealed class LiveQ149WindowlessComposeTests
{
    private const int TrialsL1 = 20;
    private const int TrialsL1Control = 5;
    private const int TrialsFetchOnly = 10;
    private const int TrialsLate = 3;
    private const int TrialsShow = 10;
    private const int ProcessesW6 = 3;
    private const int DraftsPerProcessW6 = 10;
    private const int ProcessesP1 = 3;
    private const int DraftsPerProcessP1 = 10;
    private const int DraftsP3 = 5;

    private readonly LiveQ149Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveQ149WindowlessComposeTests(LiveQ149Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private Q149Run Run => _fixture.Run;

    // One test per cell, so a run can select any group of cells by name and stay short: a run on a frozen
    // guest has to finish its suite well inside the time its checkpoint allows.

    [Fact]
    [Trait("Requires", "MailAccount")]
    public void L1_TheShippedCompose_WithoutItsPin()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        ShippedCompose(Q149Cells.L1, TrialsL1, pin: false);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    public void L1c_TheShippedCompose_WithItsPin()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        ShippedCompose(Q149Cells.L1Control, TrialsL1Control, pin: true);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    public void L2_AnInspectorNeverActivated_ReadOnce()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        FetchOnly(Q149Cells.L2, TrialsFetchOnly, readBudgetMs: 0, firstReadAfterStartMs: 0, show: ShowKind.None);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    public void W2_AnInspectorNeverActivated_ReReadFor30Seconds()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        FetchOnly(Q149Cells.W2, TrialsFetchOnly, readBudgetMs: 30_000, firstReadAfterStartMs: 0, show: ShowKind.None);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    public void W3_AnInspectorNeverActivated_ReadFrom180Seconds()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        FetchOnly(Q149Cells.W3, TrialsLate, readBudgetMs: 30_000, firstReadAfterStartMs: 180_000, show: ShowKind.None);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "InteractiveDesktop")]
    public void W4_TheInspectorsWindowShownByWindows_WithoutActivate()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        FetchOnly(Q149Cells.W4, TrialsShow, readBudgetMs: 5_000, firstReadAfterStartMs: 0, show: ShowKind.Inspector);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "InteractiveDesktop")]
    public void W5_ThePinsWindowShownByWindows_WithoutActivate()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        FetchOnly(Q149Cells.W5, TrialsShow, readBudgetMs: 5_000, firstReadAfterStartMs: 0, show: ShowKind.Pin);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "InteractiveDesktop")]
    public void W6_OnePromotionPerProcess()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        PromoteOnce(Q149Cells.W6, ProcessesW6, DraftsPerProcessW6, formPageCommit: false);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "InteractiveDesktop")]
    public void W7_OnePromotionPerProcess_ThenTheFormPageCommit()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        PromoteOnce(Q149Cells.W7, ProcessesW6, DraftsPerProcessW6, formPageCommit: true);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "InteractiveDesktop")]
    public void P1_TheProduct_PromotesOncePerOutlookStart_AndTheLastSessionOutClosesThePin()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        ProductOncePerStart(Q149Cells.P1, ProcessesP1, DraftsPerProcessP1);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "InteractiveDesktop")]
    public void P2_TwoSessionsShareOnePin_AndTheLastOutClosesIt()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        SharedPin(Q149Cells.P2);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "InteractiveDesktop")]
    public void P3_AComStartedOutlookWithAWindow_ActivatesNothing()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        ComStartWithAWindow(Q149Cells.P3, DraftsP3);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "InteractiveDesktop")]
    public void Smoke_EveryWindowlessCellOnce()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        ShippedCompose(Q149Cells.L1, 1, pin: false);
        ShippedCompose(Q149Cells.L1Control, 1, pin: true);
        FetchOnly(Q149Cells.L2, 1, readBudgetMs: 0, firstReadAfterStartMs: 0, show: ShowKind.None);
        FetchOnly(Q149Cells.W2, 1, readBudgetMs: 3_000, firstReadAfterStartMs: 0, show: ShowKind.None);
        FetchOnly(Q149Cells.W4, 1, readBudgetMs: 2_000, firstReadAfterStartMs: 0, show: ShowKind.Inspector);
        FetchOnly(Q149Cells.W5, 1, readBudgetMs: 2_000, firstReadAfterStartMs: 0, show: ShowKind.Pin);
        PromoteOnce(Q149Cells.W6, 1, 3, formPageCommit: false);
        PromoteOnce(Q149Cells.W7, 1, 3, formPageCommit: true);
        ProductOncePerStart(Q149Cells.P1, 1, 3);
        SharedPin(Q149Cells.P2);
        ComStartWithAWindow(Q149Cells.P3, 2);
    }

    // ------------------------------------------------------------------ L1: the shipped compose

    private void ShippedCompose(string cell, int trials, bool pin)
    {
        _fixture.ReleaseService();
        List<(string EntryId, string StoreId, string Token, int Iteration)> written = new();
        int exits = 0;
        int observed = 0;
        for (int i = 1; i <= trials; i++)
        {
            Q149Attempt attempt = new(cell, i);
            Run.Log(cell + " #" + i.ToString(CultureInfo.InvariantCulture) + ": stopping Outlook");
            LiveComposeProbe.StopOutlook(Run.Settings, Run.Log);
            Run.Log(cell + " #" + i.ToString(CultureInfo.InvariantCulture) + ": cold start");
            Stopwatch clock = Stopwatch.StartNew();
            long goneAt = -1;
            OutlookComSession session = OutlookComSession.Connect(
                allowStartingOutlook: true,
                onOutlookGone: _ => Interlocked.CompareExchange(ref goneAt, clock.ElapsedMilliseconds, -1),
                lifetimePin: pin);
            Q149ProcessWatch? watch = null;
            Q149WindowWatch? windows = null;
            try
            {
                attempt.Set("connectMs", clock.ElapsedMilliseconds)
                    .Set("startedOutlook", session.StartedOutlook)
                    .Set("pinRequested", pin)
                    .Set("pinned", session.ComposeSurfacePinned)
                    .Set("pinError", session.ComposeSurfacePinError);
                watch = RequireColdStart(attempt, clock);

                // The draft, tagged before anything else touches it, through a second client that lets go again.
                string entryId;
                string storeId;
                using (Q149OutlookClient seeder = Q149OutlookClient.Connect(clock, allowStart: false))
                {
                    RecordOutlookState(attempt, seeder, "before");
                    (entryId, storeId, _) = Q149Drafts.CreateTaggedDraft(Run, seeder, cell + "-" + i.ToString(CultureInfo.InvariantCulture), OldBody(i), addressToHub: false);
                }

                string token = Token("L1", i, pin);
                windows = new Q149WindowWatch(watch.Pid);
                Run.Log(cell + " #" + i.ToString(CultureInfo.InvariantCulture) + ": composing (" + clock.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture) + " ms into the trial)");
                long started = clock.ElapsedMilliseconds;
                ComDraftUpdateResult? result = null;
                string? error = null;
                try
                {
                    (ComDraftUpdateResult? Result, string? Error) outcome = LiveComposeProbe.Bounded("the shipped compose", LiveComposeProbe.StepBudget, () =>
                    {
                        ComDraftUpdateResult? r = session.TryUpdateDraft(
                            entryId, storeId, ComDraftBody.FromText(token), subject: null, toRecipients: null, ccRecipients: null, bccRecipients: null,
                            importance: null, requestReadReceipt: null, signatureOverride: null, attachmentsToAdd: Array.Empty<string>(),
                            attachmentsToRemove: Array.Empty<string>(), resume: null, display: false, out string? e);
                        return (r, e);
                    });
                    result = outcome.Result;
                    error = outcome.Error;
                }
                catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex) || ex is InvalidOperationException)
                {
                    error = OutlookComSession.DescribeComFailure(ex);
                }

                long returned = clock.ElapsedMilliseconds;
                attempt.Set("composeOk", result != null)
                    .Set("composeError", error ?? (result == null ? "see composeOk" : null))
                    .Set("composeMs", returned - started)
                    .Set("bodyPlacement", result == null ? null : (result.BodyPlacedViaWordEditor ? "wordEditor" : "html"));

                Run.Log(cell + " #" + i.ToString(CultureInfo.InvariantCulture) + ": composed in " + (returned - started).ToString(CultureInfo.InvariantCulture) + " ms; watching Outlook");
                Q149Lifetime lifetime = WatchLifetime(attempt, watch, windows, clock, started, returned, pin ? 10_000 : 30_000,
                    quitEventSeen: () => session.GoneSignal == OutlookComSession.GoneByQuitEvent);
                attempt.Set("goneSignal", session.GoneSignal)
                    .Set("goneAtMsFromCompose", Interlocked.Read(ref goneAt) < 0 ? null : Interlocked.Read(ref goneAt) - started);
                observed++;
                if (Q149LifetimeRules.Exited(lifetime))
                {
                    exits++;
                }

                written.Add((entryId, storeId, token, i));
            }
            finally
            {
                windows?.Dispose();
                session.Dispose();
                bool alive = watch != null && !watch.HasExited;
                watch?.Dispose();
                Run.Write(attempt);
                if (alive)
                {
                    LiveComposeProbe.StopOutlook(Run.Settings, Run.Log);
                }
            }
        }

        VerifyCommits(cell, written);
        string verdict = cell == Q149Cells.L1 ? Q149Decisions.Lifetime(exits, observed) : "control";
        Summary(cell, observed, "exits=" + exits.ToString(CultureInfo.InvariantCulture) + "; " + verdict);
        Assert.Equal(trials, observed);
    }

    // ------------------------------------------------------------------ L2, W2, W3, W4, W5: an inspector never activated

    private enum ShowKind
    {
        None,
        Inspector,
        Pin,
    }

    private void FetchOnly(string cell, int trials, int readBudgetMs, int firstReadAfterStartMs, ShowKind show)
    {
        _fixture.ReleaseService();
        List<(string EntryId, string StoreId, string Token, int Iteration)> written = new();
        int ok = 0;
        int shownTrials = 0;
        int exits = 0;
        int observed = 0;
        for (int i = 1; i <= trials; i++)
        {
            Q149Attempt attempt = new(cell, i);
            Run.Log(cell + " #" + i.ToString(CultureInfo.InvariantCulture) + ": stopping Outlook");
            LiveComposeProbe.StopOutlook(Run.Settings, Run.Log);
            Run.Log(cell + " #" + i.ToString(CultureInfo.InvariantCulture) + ": cold start");
            Stopwatch clock = Stopwatch.StartNew();
            Q149OutlookClient client = Q149OutlookClient.Connect(clock, allowStart: true);
            Q149ProcessWatch? watch = null;
            Q149WindowWatch? windows = null;
            object? pin = null;
            try
            {
                attempt.Set("connectMs", clock.ElapsedMilliseconds).Set("quitSink", client.QuitSinkActive);
                watch = RequireColdStart(attempt, clock);
                RecordOutlookState(attempt, client, "before");
                (string entryId, string storeId, Q149HtmlFacts stub) = Q149Drafts.CreateTaggedDraft(
                    Run, client, cell + "-" + i.ToString(CultureInfo.InvariantCulture), htmlBody: null, addressToHub: false);
                attempt.Set("itemsAddStub", stub.ToFields());
                string token = Token(cell.Substring(3, 2), i, show == ShowKind.Pin);
                windows = new Q149WindowWatch(watch.Pid);

                if (firstReadAfterStartMs > 0)
                {
                    int wait = firstReadAfterStartMs - (int)watch.MsSinceStart;
                    if (wait > 0)
                    {
                        Thread.Sleep(wait);
                    }
                }

                long started = clock.ElapsedMilliseconds;
                Q149EditorRead read = client.Run((app, ns) =>
                {
                    dynamic? item = null;
                    dynamic? inspector = null;
                    try
                    {
                        if (show == ShowKind.Pin)
                        {
                            HashSet<IntPtr> beforePin = LiveComposeProbe.FrameSnapshot();
                            pin = AddPin(app, ns, attempt);
                            (IntPtr pinFrame, string pinFoundBy) = LiveComposeProbe.FrameOf(pin, beforePin);
                            attempt.Set("windowFoundBy", pinFoundBy);
                            ShowWithoutActivate(pinFrame, attempt, "pin");
                        }

                        item = ns.GetItemFromID(entryId, storeId);
                        Q149Drafts.RequireOwnUnsentDraft(Run, item);
                        attempt.Set("htmlBeforeInspector", Q149Html.Analyze((string?)item.HTMLBody).ToFields());
                        HashSet<IntPtr> beforeInspector = LiveComposeProbe.FrameSnapshot();
                        inspector = item.GetInspector;
                        attempt.Set("htmlAfterTouch", Q149Html.Analyze((string?)item.HTMLBody).ToFields());
                        (bool? isWordMail, int? editorType) = Q149Editor.DescribeInspector((object)inspector);
                        attempt.Set("isWordMail", isWordMail).Set("editorType", editorType)
                            .Set("msSinceOutlookStartAtRead", (long)watch.MsSinceStart)
                            .Set("explorersAtRead", Count(app.Explorers))
                            .Set("inspectorsAtRead", Count(app.Inspectors))
                            .Set("opusAppAtRead", LiveComposeProbe.CountOpusAppWindows());
                        if (show == ShowKind.Inspector)
                        {
                            (IntPtr frame, string foundBy) = LiveComposeProbe.FrameOf((object)inspector, beforeInspector);
                            attempt.Set("windowFoundBy", foundBy);
                            ShowWithoutActivate(frame, attempt, "inspector");
                        }

                        Q149EditorRead r = Q149Editor.ReadEditor((object)inspector, readBudgetMs, 250);
                        if (r.Ok)
                        {
                            Q149Editor.ReplaceText(r.Document!, token);
                            dynamic current = inspector.CurrentItem;
                            try
                            {
                                current.Save();
                                attempt.Set("savedThroughCurrentItem", true);
                            }
                            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                            {
                                attempt.Set("saveFailure", Q149Com.Failure(ex));
                            }
                            finally
                            {
                                Q149Com.Release(current);
                                Q149Com.Release(r.Document);
                            }
                        }

                        inspector.Close(LiveComposeProbe.OlDiscard);
                        return r;
                    }
                    finally
                    {
                        Q149Com.Release(inspector);
                        Q149Com.Release(item);
                    }
                });

                long closed = clock.ElapsedMilliseconds;
                Run.Log(cell + " #" + i.ToString(CultureInfo.InvariantCulture) + ": editor " + (read.Ok ? "obtained" : "refused") + ", inspector closed; watching Outlook");
                attempt.Set("editorOk", read.Ok)
                    .Set("editorMs", read.Ms)
                    .Set("editorPolls", read.Polls)
                    .Set("editorFailure", read.Failure)
                    .Set("treatmentMs", closed - started);
                if (read.Ok)
                {
                    ok++;
                }

                Q149Lifetime lifetime = WatchLifetime(attempt, watch, windows, clock, started, closed, show == ShowKind.Pin ? 5_000 : 10_000,
                    quitEventSeen: () => client.QuitEventAtMs != null);
                attempt.Set("quitEventAtMsFromTreatment", client.QuitEventAtMs - started);
                if (windows.Shows > 0)
                {
                    shownTrials++;
                }

                observed++;
                if (Q149LifetimeRules.Exited(lifetime))
                {
                    exits++;
                }

                written.Add((entryId, storeId, token, i));
            }
            finally
            {
                windows?.Dispose();
                if (pin != null)
                {
                    ClosePin(client, pin, attempt);
                }

                client.Dispose();
                bool alive = watch != null && !watch.HasExited;
                watch?.Dispose();
                Run.Write(attempt);
                if (alive)
                {
                    LiveComposeProbe.StopOutlook(Run.Settings, Run.Log);
                }
            }
        }

        VerifyCommits(cell, written);
        Summary(cell, observed, string.Format(
            CultureInfo.InvariantCulture,
            "editor {0} of {1}; trials with a window shown {2}; Outlook exited after the close in {3}",
            ok,
            observed,
            shownTrials,
            exits));
        Assert.Equal(trials, observed);
    }

    // ------------------------------------------------------------------ W6, W7: one promotion per process

    /// <summary>
    /// The product's promotion on a process's first draft, then the editor read directly on every later one.
    /// W6 commits those as B1 does (<c>CurrentItem.Save</c>, <c>Close(olDiscard)</c>); W7
    /// (<paramref name="formPageCommit"/>) as B3 does (<c>ModifiedFormPages.Add("Message")</c> before the
    /// read, <c>Close(olSave)</c> after the edit).
    /// </summary>
    private void PromoteOnce(string cell, int processes, int draftsPerProcess, bool formPageCommit)
    {
        string label = cell == Q149Cells.W7 ? "W7" : "W6";
        _fixture.ReleaseService();
        List<(string EntryId, string StoreId, string Token, int Iteration)> written = new();
        int directReads = 0;
        int directOk = 0;
        int afterExplorerReads = 0;
        int afterExplorerOk = 0;
        for (int p = 1; p <= processes; p++)
        {
            Run.Log(cell + " process " + p.ToString(CultureInfo.InvariantCulture) + ": stopping Outlook");
            LiveComposeProbe.StopOutlook(Run.Settings, Run.Log);
            Stopwatch clock = Stopwatch.StartNew();
            Q149OutlookClient client = Q149OutlookClient.Connect(clock, allowStart: true);
            Q149ProcessWatch? watch = null;
            Q149WindowWatch? windows = null;
            object? pin = null;
            Q149Attempt process = new(cell + "-process", p);
            try
            {
                watch = RequireColdStart(process, clock);
                windows = new Q149WindowWatch(watch.Pid);
                pin = client.Run((app, ns) => AddPin(app, ns, process));
                for (int d = 1; d <= draftsPerProcess + 1; d++)
                {
                    bool afterExplorer = d == draftsPerProcess + 1;
                    if (afterExplorer)
                    {
                        // A6b: a displayed Explorer opened and closed while the pin stays - does the editor survive it?
                        OpenAndCloseADisplayedExplorer(client, process);
                    }

                    Q149Attempt attempt = new(cell, (p * 100) + d);
                    bool formPage = formPageCommit && d > 1;
                    attempt.Set("process", p).Set("draft", d).Set("pid", watch.Pid)
                        .Set("mode", d == 1 ? "product-promotion" : afterExplorer ? "direct-after-displayed-explorer" : "direct")
                        .Set("commit", formPage ? "formpage-close-olsave" : "save-close-oldiscard");
                    (string entryId, string storeId, _) = Q149Drafts.CreateTaggedDraft(
                        Run, client, label.ToLowerInvariant() + "-" + p.ToString(CultureInfo.InvariantCulture) + "-" + d.ToString(CultureInfo.InvariantCulture), OldBody((p * 100) + d), addressToHub: false);
                    string token = Token(label, (p * 100) + d, pin: true);
                    int showsBefore = windows.Shows;
                    client.Run<object?>((app, ns) =>
                    {
                        dynamic? item = null;
                        dynamic? inspector = null;
                        object? document = null;
                        try
                        {
                            item = ns.GetItemFromID(entryId, storeId);
                            Q149Drafts.RequireOwnUnsentDraft(Run, item);
                            inspector = item.GetInspector;
                            if (d == 1)
                            {
                                document = ComposeSurface.PromoteForWordEditor((object)inspector, out string? error, out bool windowless);
                                attempt.Set("promotionError", error).Set("windowlessAtPromotion", windowless);
                            }
                            else
                            {
                                if (formPage)
                                {
                                    Q149Editor.AddMessageFormPage((object)inspector, attempt);
                                }

                                Q149EditorRead read = Q149Editor.ReadEditor((object)inspector, 0, 0);
                                document = read.Document;
                                attempt.Set("editorFailure", read.Failure);
                            }

                            attempt.Set("editorOk", document != null);
                            if (document != null && formPage)
                            {
                                Q149Editor.ReplaceText(document, token);
                                try
                                {
                                    inspector.Close(LiveComposeProbe.OlSave);
                                    return null;
                                }
                                catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                                {
                                    attempt.Set("commitFailure", Q149Com.Failure(ex));
                                }
                            }
                            else if (document != null)
                            {
                                Q149Editor.ReplaceText(document, token);
                                dynamic current = inspector.CurrentItem;
                                try
                                {
                                    current.Save();
                                }
                                catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                                {
                                    attempt.Set("saveFailure", Q149Com.Failure(ex));
                                }
                                finally
                                {
                                    Q149Com.Release(current);
                                }
                            }

                            try
                            {
                                inspector.Close(LiveComposeProbe.OlDiscard);
                            }
                            catch (Exception ex) when (formPage && OutlookComSession.IsComCallFailure(ex))
                            {
                                // Only after a failed Close(olSave): recorded, and the cell goes on.
                                attempt.Set("closeFailure", Q149Com.Failure(ex));
                            }

                            return null;
                        }
                        finally
                        {
                            Q149Com.Release(document);
                            Q149Com.Release(inspector);
                            Q149Com.Release(item);
                        }
                    });

                    attempt.Set("windowShows", windows.Shows - showsBefore).Set("outlookAlive", !watch.HasExited);
                    if (d > 1)
                    {
                        bool ok = attempt.Get("editorOk") is true;
                        if (afterExplorer)
                        {
                            afterExplorerReads++;
                            afterExplorerOk += ok ? 1 : 0;
                        }
                        else
                        {
                            directReads++;
                            directOk += ok ? 1 : 0;
                        }
                    }

                    written.Add((entryId, storeId, token, (p * 100) + d));
                    Run.Write(attempt);
                }
            }
            finally
            {
                windows?.Dispose();
                if (pin != null)
                {
                    ClosePin(client, pin, process);
                }

                client.Dispose();
                bool alive = watch != null && !watch.HasExited;
                watch?.Dispose();
                Run.Write(process);
                if (alive)
                {
                    LiveComposeProbe.StopOutlook(Run.Settings, Run.Log);
                }
            }
        }

        (int committed, int changedForm) = VerifyCommits(cell, written);
        Summary(cell, processes, string.Format(
            CultureInfo.InvariantCulture,
            "direct reads after one promotion: {0} of {1}; after a displayed Explorer was opened and closed: {2} of {3}; edits in the stored draft {4} of {5}, {6} with a message class or form properties changed",
            directOk,
            directReads,
            afterExplorerOk,
            afterExplorerReads,
            committed,
            written.Count,
            changedForm));
    }

    // ------------------------------------------------------------------ P1, P2, P3: the product as it drafts since Q208 B

    /// <summary>
    /// P1: per cold start, one product session - the default registry, the product's own pin rules - drafts
    /// <paramref name="draftsPerProcess"/> drafts through update_draft. Each draft's route, the windows it put on screen and
    /// any change of foreground are recorded; then the session ends - the last session out - and Outlook's exit is
    /// watched. The commits are re-read through a separate client at the end.
    /// </summary>
    private void ProductOncePerStart(string cell, int processes, int draftsPerProcess)
    {
        _fixture.ReleaseService();
        List<(string EntryId, string StoreId, string Token, int Iteration)> written = new();
        int firsts = 0;
        int firstsPromoted = 0;
        int firstsShown = 0;
        int laters = 0;
        int latersDirect = 0;
        int latersShown = 0;
        int latersForeground = 0;
        int composed = 0;
        int exitsAfterLastOut = 0;
        int pinsLeft = 0;
        for (int p = 1; p <= processes; p++)
        {
            Q149Attempt process = new(cell + "-process", p);
            Stopwatch clock = Stopwatch.StartNew();
            OutlookComSession? session = null;
            Q149ProcessWatch? watch = null;
            Q149WindowWatch? windows = null;
            Q180ColdStartWatch? start = null;
            try
            {
                Run.Log(cell + " process " + p.ToString(CultureInfo.InvariantCulture) + ": stopping Outlook");
                LiveComposeProbe.StopOutlook(Run.Settings, Run.Log);
                clock.Restart();
                start = new Q180ColdStartWatch("com", cell, p);
                session = ProductColdStart(start);
                watch = RequireColdStart(process, clock);
                process.Set("pinMadeAtConnect", session.ComposeSurfacePinned).Set("registryAccessAtConnect", session.LastRegistryAccess);
                windows = new Q149WindowWatch(watch.Pid);
                for (int d = 1; d <= draftsPerProcess; d++)
                {
                    int iteration = (p * 100) + d;
                    Q149Attempt attempt = new(cell, iteration);
                    attempt.Set("process", p).Set("draft", d).Set("pid", watch.Pid);
                    (string entryId, string storeId, string token, bool ok) = ProductDraft(session, windows, attempt, cell, iteration, clock);
                    written.Add((entryId, storeId, token, iteration));
                    composed += ok ? 1 : 0;
                    bool shown = attempt.Get("windowShows") is int shows && shows > 0;
                    if (d == 1)
                    {
                        firsts++;
                        firstsPromoted += Equals(attempt.Get("route"), "Promote") ? 1 : 0;
                        firstsShown += shown ? 1 : 0;
                    }
                    else
                    {
                        laters++;
                        latersDirect += Equals(attempt.Get("route"), "Direct") ? 1 : 0;
                        latersShown += shown ? 1 : 0;
                        latersForeground += attempt.Get("foregroundToOutlook") is int fg && fg > 0 ? 1 : 0;
                    }

                    Run.Write(attempt);
                }

                OutlookComSession leaving = session;
                session = null;
                (bool exited, bool pinLeft) = LastOut(leaving, watch, process, clock);
                exitsAfterLastOut += exited ? 1 : 0;
                pinsLeft += pinLeft ? 1 : 0;
            }
            finally
            {
                windows?.Dispose();
                session?.Dispose();
                start?.Finish(Run);
                bool alive = watch != null && !watch.HasExited;
                watch?.Dispose();
                Run.Write(process);
                if (alive)
                {
                    LiveComposeProbe.StopOutlook(Run.Settings, Run.Log);
                }
            }
        }

        (int committed, int changedForm) = VerifyCommits(cell, written);
        Summary(cell, processes, string.Format(
            CultureInfo.InvariantCulture,
            "first draft of each start promoted {0} of {1}, a window shown in {2}; later drafts direct {3} of {4}, a window shown in {5}, "
            + "the foreground taken in {6}; composed {7} of {8}; edits in the stored draft {9}, {10} with a message class or form "
            + "properties changed; Outlook ended after the last session out {11} of {12}; pins left in the registry {13}",
            firstsPromoted,
            firsts,
            firstsShown,
            latersDirect,
            laters,
            latersShown,
            latersForeground,
            composed,
            written.Count,
            committed,
            changedForm,
            exitsAfterLastOut,
            processes,
            pinsLeft));
        Assert.Equal(processes * draftsPerProcess, written.Count);
    }

    /// <summary>
    /// P2: two product sessions in one window-less Outlook. The first drafts (the promotion, then direct) and so makes
    /// the pin; the second drafts on it; the first leaves - Outlook must stay up - the second drafts once more and
    /// leaves last, closing by its window a pin it did not make.
    /// </summary>
    private void SharedPin(string cell)
    {
        _fixture.ReleaseService();
        List<(string EntryId, string StoreId, string Token, int Iteration)> written = new();
        Q149Attempt process = new(cell + "-process", 1);
        Stopwatch clock = Stopwatch.StartNew();
        OutlookComSession? first = null;
        OutlookComSession? second = null;
        Q149ProcessWatch? watch = null;
        Q149WindowWatch? windows = null;
        Q180ColdStartWatch? start = null;
        int iteration = 0;
        List<string> routes = new();
        try
        {
            Run.Log(cell + ": stopping Outlook");
            LiveComposeProbe.StopOutlook(Run.Settings, Run.Log);
            clock.Restart();
            start = new Q180ColdStartWatch("com", cell, 1);
            first = ProductColdStart(start);
            watch = RequireColdStart(process, clock);
            second = LiveComposeProbe.Bounded("the second session's connect", LiveComposeProbe.StepBudget, () => OutlookComSession.Connect(allowStartingOutlook: false));
            windows = new Q149WindowWatch(watch.Pid);

            void Draft(OutlookComSession session, string who)
            {
                iteration++;
                Q149Attempt attempt = new(cell, iteration);
                attempt.Set("session", who).Set("pid", watch!.Pid);
                (string entryId, string storeId, string token, _) = ProductDraft(session, windows!, attempt, cell, iteration, clock);
                written.Add((entryId, storeId, token, iteration));
                routes.Add(who + ":" + (attempt.Get("route") as string ?? "none"));
                Run.Write(attempt);
            }

            Draft(first, "first");
            Draft(first, "first");
            Draft(second, "second");
            Draft(second, "second");
            ProcessStamp? outlook = first.RegisteredOutlook;
            process.Set("pinMadeByFirst", first.ComposeSurfacePinned)
                .Set("pinMadeBySecond", second.ComposeSurfacePinned)
                .Set("pinRecorded", RecordedPin(outlook) != 0);

            OutlookComSession firstOut = first;
            first = null;
            firstOut.Dispose();
            bool endedWithTheFirst = watch.WaitForExit(5_000);
            process.Set("firstLeave", firstOut.LeaveOutcome)
                .Set("outlookRunningAfterFirstOut", !endedWithTheFirst)
                .Set("pinRecordedAfterFirstOut", RecordedPin(outlook) != 0);
            if (!endedWithTheFirst)
            {
                Draft(second, "second, after the first left");
            }

            OutlookComSession lastOut = second;
            second = null;
            _ = LastOut(lastOut, watch, process, clock);
        }
        finally
        {
            windows?.Dispose();
            first?.Dispose();
            second?.Dispose();
            start?.Finish(Run);
            bool alive = watch != null && !watch.HasExited;
            watch?.Dispose();
            process.Set("routes", routes.ToArray());
            Run.Write(process);
            if (alive)
            {
                LiveComposeProbe.StopOutlook(Run.Settings, Run.Log);
            }
        }

        (int committed, _) = VerifyCommits(cell, written);
        Summary(cell, 1, string.Format(
            CultureInfo.InvariantCulture,
            "routes {0}; Outlook kept running after the first session out: {1}; the last session out: {2}; Outlook ended: {3}; "
            + "edits in the stored draft {4} of {5}",
            string.Join(", ", routes),
            process.Get("outlookRunningAfterFirstOut"),
            process.Get("leaveOutcome"),
            process.Get("outlookExitedAfterLastOut"),
            committed,
            written.Count));
    }

    /// <summary>
    /// P3: a COM-started Outlook in which an Explorer is then displayed (parked off-screen) - the state an agent's
    /// cold start is in once the user opens Outlook. Every product draft must read its editor directly, with no window
    /// activated and no pin made; closing the Explorer, its only window, then ends it.
    /// </summary>
    private void ComStartWithAWindow(string cell, int drafts)
    {
        _fixture.ReleaseService();
        List<(string EntryId, string StoreId, string Token, int Iteration)> written = new();
        Q149Attempt process = new(cell + "-process", 1);
        Stopwatch clock = Stopwatch.StartNew();
        OutlookComSession? session = null;
        Q149OutlookClient? client = null;
        object? explorer = null;
        Q149ProcessWatch? watch = null;
        Q149WindowWatch? windows = null;
        Q180ColdStartWatch? start = null;
        int direct = 0;
        int shown = 0;
        try
        {
            Run.Log(cell + ": stopping Outlook");
            LiveComposeProbe.StopOutlook(Run.Settings, Run.Log);
            clock.Restart();
            start = new Q180ColdStartWatch("com", cell, 1);
            session = ProductColdStart(start);
            watch = RequireColdStart(process, clock);
            client = Q149OutlookClient.Connect(clock, allowStart: false);
            explorer = DisplayAParkedExplorer(client, process);
            process.Set("outlookShowsAWindow", ComposeSurface.OutlookShowsAWindow());
            windows = new Q149WindowWatch(watch.Pid);
            for (int d = 1; d <= drafts; d++)
            {
                Q149Attempt attempt = new(cell, d);
                attempt.Set("pid", watch.Pid);
                (string entryId, string storeId, string token, _) = ProductDraft(session, windows, attempt, cell, d, clock);
                written.Add((entryId, storeId, token, d));
                direct += Equals(attempt.Get("route"), "Direct") ? 1 : 0;
                shown += attempt.Get("windowShows") is int shows && shows > 0 ? 1 : 0;
                Run.Write(attempt);
            }

            process.Set("pinMade", session.ComposeSurfacePinned).Set("pinRecorded", RecordedPin(session.RegisteredOutlook) != 0);

            // The user's close of their window: with no pin, the window-less Outlook ends.
            if (explorer != null)
            {
                object closing = explorer;
                explorer = null;
                CloseTheExplorer(client, closing, process);
                bool exited = watch.WaitForExit(30_000);
                process.Set("outlookEndedAtTheExplorersClose", exited)
                    .Set("goneSignal", session.GoneSignal);
            }
        }
        finally
        {
            windows?.Dispose();
            if (explorer != null && client != null)
            {
                CloseTheExplorer(client, explorer, process);
            }

            client?.Dispose();
            session?.Dispose();
            start?.Finish(Run);
            bool alive = watch != null && !watch.HasExited;
            watch?.Dispose();
            Run.Write(process);
            if (alive)
            {
                LiveComposeProbe.StopOutlook(Run.Settings, Run.Log);
            }
        }

        (int committed, int changedForm) = VerifyCommits(cell, written);
        Summary(cell, drafts, string.Format(
            CultureInfo.InvariantCulture,
            "direct {0} of {1}, a window shown in {2}; a pin made: {3}; edits in the stored draft {4} of {5}, {6} with a message class "
            + "or form properties changed; Outlook ended at the Explorer's close: {7}",
            direct,
            drafts,
            shown,
            process.Get("pinMade"),
            committed,
            written.Count,
            changedForm,
            process.Get("outlookEndedAtTheExplorersClose")));
    }

    /// <summary>
    /// The product's own cold start: <c>OutlookComSession.Connect</c> as a COM host makes it, then the first store read -
    /// the moment Outlook counts as ready for Q180 (b). A step that does not return within its budget is logged as a hang
    /// and fails the cell; nothing is killed.
    /// </summary>
    private static OutlookComSession ProductColdStart(Q180ColdStartWatch start)
    {
        start.Milestone("startRequested");
        OutlookComSession session;
        try
        {
            session = LiveComposeProbe.Bounded("the product's connect, a cold start", LiveComposeProbe.StepBudget, () => OutlookComSession.Connect(allowStartingOutlook: true));
        }
        catch (TimeoutException)
        {
            start.Hang("connect");
            throw;
        }

        start.Milestone("connected");
        try
        {
            _ = LiveComposeProbe.Bounded("the first store read", LiveComposeProbe.StepBudget, () => session.GetStores().Count);
        }
        catch (TimeoutException)
        {
            start.Hang("the first store read");
            throw;
        }

        start.Ready();
        return session;
    }

    /// <summary>
    /// One product draft: a tagged draft seeded through a second client that lets go again, then the product's compose
    /// through <paramref name="session"/>'s update_draft. Records the route, the time, the windows shown and any change of
    /// foreground on <paramref name="attempt"/>.
    /// </summary>
    private (string EntryId, string StoreId, string Token, bool Composed) ProductDraft(
        OutlookComSession session, Q149WindowWatch windows, Q149Attempt attempt, string cell, int iteration, Stopwatch clock)
    {
        string entryId;
        string storeId;
        using (Q149OutlookClient seeder = Q149OutlookClient.Connect(clock, allowStart: false))
        {
            (entryId, storeId, _) = Q149Drafts.CreateTaggedDraft(Run, seeder, cell.Substring(3, 2).ToLowerInvariant() + "-" + iteration.ToString(CultureInfo.InvariantCulture), OldBody(iteration), addressToHub: false);
        }

        string token = Token(cell.Substring(3, 2), iteration, pin: true);
        int shows = windows.Shows;
        int onScreen = windows.OnScreenShows;
        int foreground = windows.ForegroundToOutlook;
        int classes = windows.ShownClasses.Count;
        long started = clock.ElapsedMilliseconds;
        ComDraftUpdateResult? result = null;
        string? error = null;
        try
        {
            (result, error) = LiveComposeProbe.Bounded("the product's compose", LiveComposeProbe.StepBudget, () =>
            {
                ComDraftUpdateResult? r = session.TryUpdateDraft(
                    entryId, storeId, ComDraftBody.FromText(token), subject: null, toRecipients: null, ccRecipients: null, bccRecipients: null,
                    importance: null, requestReadReceipt: null, signatureOverride: null, attachmentsToAdd: Array.Empty<string>(),
                    attachmentsToRemove: Array.Empty<string>(), resume: null, display: false, out string? e);
                return (r, e);
            });
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex) || ex is InvalidOperationException)
        {
            error = OutlookComSession.DescribeComFailure(ex);
        }

        long returned = clock.ElapsedMilliseconds;

        // A window shown at the very end of the compose reaches the watch's own message loop a moment later.
        Thread.Sleep(300);
        attempt.Set("composeOk", result != null)
            .Set("composeError", error ?? (result == null ? "see composeOk" : null))
            .Set("composeMs", returned - started)
            .Set("bodyPlacement", result == null ? null : (result.BodyPlacedViaWordEditor ? "wordEditor" : "html"))
            .Set("route", session.LastEditorRoute)
            .Set("windowShows", windows.Shows - shows)
            .Set("onScreenShows", windows.OnScreenShows - onScreen)
            .Set("foregroundToOutlook", windows.ForegroundToOutlook - foreground)
            .Set("shownClasses", windows.ShownClasses.Skip(classes).Distinct(StringComparer.Ordinal).ToArray())
            .Set("pinMadeByThisSession", session.ComposeSurfacePinned)
            .Set("pinRecorded", RecordedPin(session.RegisteredOutlook) != 0)
            .Set("registryAccess", session.LastRegistryAccess)
            .Set("userVisibleWindowsAfter", ComposeSurface.CountUserVisibleWindows());
        return (entryId, storeId, token, result != null);
    }

    /// <summary>
    /// Ends <paramref name="session"/> - the last OutlookAI session out - and watches what follows: what its leave
    /// decided, whether Outlook ended within 30 s and how, and whether a pin was left in the registry.
    /// </summary>
    private static (bool Exited, bool PinLeft) LastOut(OutlookComSession session, Q149ProcessWatch watch, Q149Attempt process, Stopwatch clock)
    {
        ProcessStamp? outlook = session.RegisteredOutlook;
        bool pinBefore = RecordedPin(outlook) != 0;
        long leftAt = clock.ElapsedMilliseconds;
        session.Dispose();
        bool exited = watch.WaitForExit(30_000);
        bool pinLeft = RecordedPin(outlook) != 0;
        process.Set("pinRecordedBeforeLastOut", pinBefore)
            .Set("leaveOutcome", session.LeaveOutcome)
            .Set("outlookExitedAfterLastOut", exited)
            .Set("exitMsAfterLastOut", watch.ExitedAtMs - leftAt)
            .Set("exitCode", watch.ExitCode == null ? null : "0x" + watch.ExitCode.Value.ToString("X8", CultureInfo.InvariantCulture))
            .Set("pinLeftInRegistry", pinLeft);
        return (exited, pinLeft);
    }

    /// <summary>The window of the pin the session registry records for <paramref name="outlook"/>, 0 when none.</summary>
    private static long RecordedPin(ProcessStamp? outlook)
    {
        return outlook == null ? 0 : OutlookSessionRegistry.Default.TryRead()?.Find(outlook.Value)?.PinWindow ?? 0;
    }

    /// <summary>An Explorer on the hub's Drafts - never the Inbox, where an Add can hand back a pin - displayed and parked off-screen.</summary>
    private object? DisplayAParkedExplorer(Q149OutlookClient client, Q149Attempt process)
    {
        return client.Run<object?>((app, ns) =>
        {
            dynamic? store = null;
            dynamic? drafts = null;
            dynamic? explorers = null;
            object? explorer = null;
            try
            {
                store = Q149Drafts.FindStore(ns, Run.Hub);
                drafts = store!.GetDefaultFolder(LiveComposeProbe.Drafts);
                explorers = app.Explorers;
                explorer = (object)explorers.Add(drafts, 0);
                ((dynamic)explorer).Display();
                IntPtr hwnd = LiveComposeProbe.WindowOf(explorer);
                if (hwnd != IntPtr.Zero)
                {
                    Q149Native.Park(hwnd);
                }

                PumpedStaRunner.PumpedWait(500);
                process.Set("explorerDisplayed", true);
                return explorer;
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                process.Set("explorerFailure", Q149Com.Failure(ex));
                Q149Com.Release(explorer);
                return null;
            }
            finally
            {
                Q149Com.Release(explorers);
                Q149Com.Release(drafts);
                Q149Com.Release(store);
            }
        });
    }

    private static void CloseTheExplorer(Q149OutlookClient client, object explorer, Q149Attempt process)
    {
        try
        {
            client.Run<object?>((_, _) =>
            {
                try
                {
                    ((dynamic)explorer).Close();
                    process.Set("explorerClosed", true);
                }
                catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                {
                    process.Set("explorerCloseFailure", Q149Com.Failure(ex));
                }
                finally
                {
                    Q149Com.Release(explorer);
                }

                return null;
            });
        }
        catch (Exception ex) when (ex is TimeoutException or InvalidOperationException or ObjectDisposedException)
        {
            process.Set("explorerCloseFailure", ex.GetType().Name);
        }
    }

    // ------------------------------------------------------------------ shared steps

    /// <summary>Asserts the trial really cold-started a window-less Outlook, and returns the watch on it.</summary>
    private Q149ProcessWatch RequireColdStart(Q149Attempt attempt, Stopwatch clock)
    {
        Q149ProcessWatch? watch = Q149ProcessWatch.AttachToTheOnlyOutlook(clock);
        Assert.True(watch != null, "Q149: not exactly one OUTLOOK.EXE after the cold start");
        attempt.Set("pid", watch!.Pid)
            .Set("embedding", watch.Embedding)
            .Set("msSinceOutlookStart", (long)watch.MsSinceStart)
            .Set("userVisibleWindows", ComposeSurface.CountUserVisibleWindows())
            .Set("headless", HealthReporting.TryGetOutlookHeadless());
        Assert.True(watch.Embedding, "Q149: OUTLOOK.EXE was not started by COM (-Embedding) - this trial would not be a window-less cold start");
        Assert.Equal(0, ComposeSurface.CountUserVisibleWindows());
        return watch;
    }

    private static void RecordOutlookState(Q149Attempt attempt, Q149OutlookClient client, string when)
    {
        client.Run<object?>((app, _) =>
        {
            attempt.Set("version", (string?)app.Version)
                .Set("explorers_" + when, Count(app.Explorers))
                .Set("inspectors_" + when, Count(app.Inspectors));
            return null;
        });
    }

    /// <summary>
    /// Watches OUTLOOK.EXE from the end of the treatment for up to <paramref name="watchMs"/>, records the
    /// presence samples, the exit time and code, the windows shown and the classification, and returns it.
    /// </summary>
    private static Q149Lifetime WatchLifetime(
        Q149Attempt attempt,
        Q149ProcessWatch watch,
        Q149WindowWatch windows,
        Stopwatch clock,
        long treatmentStarted,
        long treatmentEnded,
        int watchMs,
        Func<bool> quitEventSeen)
    {
        _ = watch.WaitForExit(watchMs);
        long? exitedAt = watch.ExitedAtMs;
        Q149Lifetime lifetime = Q149LifetimeRules.Classify(!watch.HasExited, watch.ExitCode, quitEventSeen());
        attempt.Set("lifetime", lifetime.ToString())
            .Set("watchMs", watchMs)
            .Set("exitedAtMsFromTreatmentStart", exitedAt - treatmentStarted)
            .Set("exitedDuringTreatment", exitedAt != null && exitedAt.Value <= treatmentEnded)
            .Set("exitCode", watch.ExitCode == null ? null : "0x" + watch.ExitCode.Value.ToString("X8", CultureInfo.InvariantCulture))
            .Set("samples", Q149LifetimeRules.Samples(Q149LifetimeRules.SampleOffsets(watchMs), exitedAt == null ? null : exitedAt.Value - treatmentEnded))
            .Set("windowShows", windows.Shows)
            .Set("onScreenShows", windows.OnScreenShows)
            .Set("shownClasses", windows.ShownClasses.Distinct(StringComparer.Ordinal).ToArray())
            .Set("windowEvents", windows.Events.ToArray())
            .Set("foregroundToOutlook", windows.ForegroundToOutlook)
            .Set("foregroundOutlookAfter", Q149Native.ForegroundOutlookClass())
            .Set("userVisibleAfter", watch.HasExited ? 0 : ComposeSurface.CountUserVisibleWindows());
        return lifetime;
    }

    /// <summary>The product's pin: a never-displayed Explorer on the default Inbox, made the way TryPinProcess makes it.</summary>
    private static object AddPin(dynamic app, dynamic ns, Q149Attempt attempt)
    {
        dynamic folder = ns.GetDefaultFolder(LiveComposeProbe.Inbox);
        try
        {
            dynamic explorers = app.Explorers;
            try
            {
                object pin = explorers.Add(folder, 0);
                attempt.Set("pinWindowViaIOleWindow", LiveComposeProbe.WindowOf(pin) != IntPtr.Zero);
                return pin;
            }
            finally
            {
                Q149Com.Release(explorers);
            }
        }
        finally
        {
            Q149Com.Release(folder);
        }
    }

    private static void ClosePin(Q149OutlookClient client, object pin, Q149Attempt attempt)
    {
        try
        {
            client.Run<object?>((_, _) =>
            {
                try
                {
                    ((dynamic)pin).Close();
                }
                catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                {
                    attempt.Set("pinCloseFailure", Q149Com.Failure(ex));
                }
                finally
                {
                    Q149Com.Release(pin);
                }

                return null;
            });
        }
        catch (Exception ex) when (ex is TimeoutException or InvalidOperationException or ObjectDisposedException)
        {
            attempt.Set("pinCloseFailure", ex.GetType().Name);
        }
    }

    /// <summary>
    /// W4/W5's treatment: every invisible Outlook window parked off-screen first (the product's own order),
    /// then the one window shown by Windows WITHOUT activation, a pumped 60 ms, and hidden again.
    /// </summary>
    private static void ShowWithoutActivate(IntPtr hwnd, Q149Attempt attempt, string whose)
    {
        attempt.Set("shownWindow", whose).Set("shownWindowFound", hwnd != IntPtr.Zero);
        if (hwnd == IntPtr.Zero)
        {
            return;
        }

        foreach (ComposeSurface.WindowState w in ComposeSurface.SnapshotOutlookWindows())
        {
            if (!w.Visible)
            {
                Q149Native.Park(w.Handle);
            }
        }

        IntPtr foreground = Q149Native.GetForegroundWindow();
        _ = Q149Native.ShowWindow(hwnd, Q149Native.SwShowNoActivate);
        PumpedStaRunner.PumpedWait(60);
        _ = Q149Native.ShowWindow(hwnd, Q149Native.SwHide);
        attempt.Set("foregroundChangedByShow", Q149Native.GetForegroundWindow() != foreground);
    }

    /// <summary>A6b: an Explorer on the hub's Drafts - never the pin's folder, which would hand back the pin - displayed, parked, closed.</summary>
    private void OpenAndCloseADisplayedExplorer(Q149OutlookClient client, Q149Attempt process)
    {
        client.Run<object?>((app, ns) =>
        {
            dynamic? store = null;
            dynamic? drafts = null;
            dynamic? explorers = null;
            dynamic? explorer = null;
            try
            {
                store = Q149Drafts.FindStore(ns, Run.Hub);
                drafts = store!.GetDefaultFolder(LiveComposeProbe.Drafts);
                explorers = app.Explorers;
                explorer = explorers.Add(drafts, 0);
                explorer.Display();
                IntPtr hwnd = LiveComposeProbe.WindowOf((object)explorer);
                if (hwnd != IntPtr.Zero)
                {
                    Q149Native.Park(hwnd);
                }

                PumpedStaRunner.PumpedWait(500);
                explorer.Close();
                process.Set("displayedExplorerOpenedAndClosed", true);
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                process.Set("displayedExplorerFailure", Q149Com.Failure(ex));
            }
            finally
            {
                Q149Com.Release(explorer);
                Q149Com.Release(explorers);
                Q149Com.Release(drafts);
                Q149Com.Release(store);
            }

            return null;
        });
        Thread.Sleep(1000);
    }

    /// <summary>
    /// Whether each trial's edit reached the store - read in a fresh window-less Outlook through a separate
    /// client - then every draft deleted by EntryID, tag and marker.
    /// </summary>
    /// <summary>
    /// Re-reads every written draft through a separate client: the new text in, the old out, and - what B3's
    /// commit could leave behind - the message class and the form-definition properties. Returns how many
    /// carried the edit and how many came back changed.
    /// </summary>
    private (int Committed, int ChangedForm) VerifyCommits(string cell, IReadOnlyList<(string EntryId, string StoreId, string Token, int Iteration)> written)
    {
        if (written.Count == 0)
        {
            return (0, 0);
        }

        Run.Log(cell + ": verifying " + written.Count.ToString(CultureInfo.InvariantCulture) + " stored draft(s)");
        Stopwatch clock = Stopwatch.StartNew();
        using Q149OutlookClient holder = Q149OutlookClient.Connect(clock, allowStart: true);
        int committed = 0;
        int changedForm = 0;
        foreach ((string entryId, string storeId, string token, int iteration) in written)
        {
            Q149Attempt attempt = new(cell + "-commit", iteration);
            try
            {
                (string html, _, string messageClass, IReadOnlyList<string> formProperties, _, _, _) = Q149Drafts.Read(Run, entryId, storeId);
                bool present = html.Contains(token, StringComparison.Ordinal);
                bool oldGone = !html.Contains(OldToken(iteration), StringComparison.Ordinal);
                bool clean = string.Equals(messageClass, "IPM.Note", StringComparison.OrdinalIgnoreCase) && formProperties.Count == 0;
                attempt.Set("tokenPresent", present).Set("oldAbsent", oldGone).Set("html", Q149Html.Analyze(html).ToFields())
                    .Set("messageClass", messageClass)
                    .Set("formPropertiesPresent", formProperties.Count == 0 ? null : formProperties.ToArray());
                committed += present ? 1 : 0;
                changedForm += clean ? 0 : 1;
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex) || ex is InvalidOperationException)
            {
                attempt.Set("readFailure", Q149Com.Failure(ex));
            }

            Run.Write(attempt);
            _ = LiveOutlookTestMailer.DeleteItemByEntryId(Run.Hub, entryId, Run.Marker);
        }

        Run.Log(cell + ": " + committed.ToString(CultureInfo.InvariantCulture) + " of " + written.Count.ToString(CultureInfo.InvariantCulture)
            + " edits are in the stored draft");
        return (committed, changedForm);
    }

    private void Summary(string cell, int trials, string text)
    {
        string line = Q149Cells.SummaryPrefix + cell + " n=" + trials.ToString(CultureInfo.InvariantCulture) + ": " + text;
        _output.WriteLine(line);
        Console.WriteLine(line);
    }

    private static int Count(dynamic collection)
    {
        try
        {
            return (int)collection.Count;
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
        {
            return -1;
        }
        finally
        {
            Q149Com.Release((object)collection);
        }
    }

    private string Token(string cell, int iteration, bool pin)
    {
        return "Q149NEW" + cell + (pin ? "P" : "N") + iteration.ToString(CultureInfo.InvariantCulture) + Run.Marker;
    }

    private string OldToken(int iteration)
    {
        return "Q149OLD" + iteration.ToString(CultureInfo.InvariantCulture) + Run.Marker;
    }

    private string OldBody(int iteration)
    {
        return "<html><body><p>" + OldToken(iteration) + "</p></body></html>";
    }
}
