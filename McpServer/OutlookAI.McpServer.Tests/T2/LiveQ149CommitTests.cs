using System.Diagnostics;
using System.Globalization;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Q149 (d) stage 2 - can an edit on a RE-OPENED draft be committed without <c>Inspector.Activate()</c>?
/// Since 2026-10-04 every compose works on the draft re-opened by its EntryID, and on such a draft an edit made
/// through a never-activated inspector did not commit (D46: 3 of 3; runbook 4.6 variant m30c: 400 of 400 on
/// Office LTSC 2024), which is why every compose now activates a hidden window even in a windowed Outlook.
/// <list type="bullet">
/// <item>B0: the product's update_draft (with its Activate) - the control.</item>
/// <item>B1: the editor read without Activate, the edit, <c>Inspector.CurrentItem.Save()</c>, <c>Close(olDiscard)</c> (D46).</item>
/// <item>B2: the same edit, <c>Close(olSave)</c> (m30c).</item>
/// <item>B3: <c>ModifiedFormPages.Add("Message")</c> before the edit, <c>Close(olSave)</c>; then the message
/// class and the form-definition properties checked (it can turn an item into a one-off form).</item>
/// <item>B4: Word's <c>Document.Save()</c>, then <c>Close(olDiscard)</c>.</item>
/// </list>
/// In a windowed Outlook - started the way a user starts it - with 20 drafts per variant, each verified by a
/// separate client's re-read: the new text present and the old text gone. The same question in a window-less
/// Outlook after one promotion is W6's (<see cref="LiveQ149WindowlessComposeTests"/>).
/// </summary>
[Collection(LiveCollections.ComposeSurface)]
[Trait("Category", "Live")]
public sealed class LiveQ149CommitTests
{
    private const int Drafts = 20;

    private readonly LiveQ149Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveQ149CommitTests(LiveQ149Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private Q149Run Run => _fixture.Run;

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "InteractiveDesktop")]
    public void B_ACommitWithoutActivate_OnAReopenedDraft()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        CommitCells(Drafts);
    }

    [Fact]
    [Trait("Requires", "MailAccount")]
    [Trait("Requires", "InteractiveDesktop")]
    public void Smoke_EveryCommitCellOnce()
    {
        if (!_fixture.Begin(_output))
        {
            return;
        }

        CommitCells(1);
    }

    private void CommitCells(int drafts)
    {
        LiveComposeProbe.EnsureWindowedOutlook(Run.Settings, Run.Log);
        Stopwatch clock = Stopwatch.StartNew();
        using Q149ProcessWatch? watch = Q149ProcessWatch.AttachToTheOnlyOutlook(clock);
        Assert.True(watch != null, "Q149: not exactly one OUTLOOK.EXE in the windowed stage");
        using Q149WindowWatch windows = new(watch!.Pid);
        List<(string Variant, int Committed, int N, int SideEffects)> tallies = new();

        tallies.Add(Variant(Q149Cells.B0, drafts, windows, ProductUpdate));
        using (Q149OutlookClient client = Q149OutlookClient.Connect(clock, allowStart: false))
        {
            // B4 (Word's Document.Save) is not repeated: measured once on 2026-11-17 (OutlookAI-Indexed, run
            // 20261117-181203-indexed-ee5fc276b528), it put one window on screen, did not return within the step's
            // three-minute budget, and committed nothing. Word documents Save() on a document never saved as a file as
            // prompting with its Save As dialog, a modal wait for a person; a second observation would only hold every
            // cell after it.
            foreach (string variant in new[] { Q149Cells.B1, Q149Cells.B2, Q149Cells.B3 })
            {
                tallies.Add(Variant(variant, drafts, windows, (entryId, storeId, token, attempt) => EditWithoutActivate(client, variant, entryId, storeId, token, attempt)));
            }
        }

        string verdict = Q149Decisions.Commit(tallies.Where(t => t.Variant != Q149Cells.B0).ToList());
        string line = Q149Cells.SummaryPrefix + "S2 n=" + drafts.ToString(CultureInfo.InvariantCulture) + ": "
            + string.Join("; ", tallies.Select(t => string.Format(CultureInfo.InvariantCulture, "{0} committed {1}/{2}, side effects {3}", t.Variant, t.Committed, t.N, t.SideEffects)))
            + "; " + verdict;
        _output.WriteLine(line);
        Console.WriteLine(line);
        Assert.All(tallies, t => Assert.Equal(drafts, t.N));
    }

    private (string Variant, int Committed, int N, int SideEffects) Variant(
        string cell, int drafts, Q149WindowWatch windows, Action<string, string, string, Q149Attempt> edit)
    {
        int committed = 0;
        int sideEffects = 0;
        int n = 0;
        for (int i = 1; i <= drafts; i++)
        {
            Q149Attempt attempt = new(cell, i);
            string oldToken = "Q149OLD" + i.ToString(CultureInfo.InvariantCulture) + Run.Marker;
            string newToken = "Q149NEW" + cell.Substring(3, 2) + i.ToString(CultureInfo.InvariantCulture) + Run.Marker;
            string entryId;
            string storeId;
            using (Q149OutlookClient seeder = Q149OutlookClient.Connect(new Stopwatch(), allowStart: false))
            {
                (entryId, storeId, _) = Q149Drafts.CreateTaggedDraft(
                    Run, seeder, cell + "-" + i.ToString(CultureInfo.InvariantCulture), "<html><body><p>" + oldToken + "</p></body></html>", addressToHub: false);
            }

            int shows = windows.Shows;
            int onScreen = windows.OnScreenShows;
            Stopwatch elapsed = Stopwatch.StartNew();
            try
            {
                edit(entryId, storeId, newToken, attempt);
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex) || ex is InvalidOperationException or TimeoutException)
            {
                attempt.Set("editFailure", Q149Com.Failure(ex));
            }

            attempt.Set("editMs", elapsed.ElapsedMilliseconds)
                .Set("windowShows", windows.Shows - shows)
                .Set("onScreenShows", windows.OnScreenShows - onScreen);

            (string html, _, string messageClass, IReadOnlyList<string> formProperties, int attachments, _, _) = Q149Drafts.Read(Run, entryId, storeId);
            bool present = html.Contains(newToken, StringComparison.Ordinal);
            bool oldGone = !html.Contains(oldToken, StringComparison.Ordinal);
            bool clean = string.Equals(messageClass, "IPM.Note", StringComparison.OrdinalIgnoreCase) && formProperties.Count == 0;
            attempt.Set("committed", present && oldGone)
                .Set("tokenPresent", present)
                .Set("oldAbsent", oldGone)
                .Set("messageClass", messageClass)
                .Set("formPropertiesPresent", formProperties.Count == 0 ? null : formProperties.ToArray())
                .Set("attachments", attachments)
                .Set("html", Q149Html.Analyze(html).ToFields());
            committed += present && oldGone ? 1 : 0;
            sideEffects += clean ? 0 : 1;
            n++;
            Run.Write(attempt);
            _ = LiveOutlookTestMailer.DeleteItemByEntryId(Run.Hub, entryId, Run.Marker);
        }

        return (cell, committed, n, sideEffects);
    }

    /// <summary>B0: the product's update_draft - its activation, its commit.</summary>
    private void ProductUpdate(string entryId, string storeId, string token, Q149Attempt attempt)
    {
        LiveStoreWriteGuard.Assert(Run.Hub, StoreWriteKind.Draft, "q149 B0 update_draft");
        UpdateDraftOutcome updated = _fixture.Service.UpdateDraft(entryId, body: token, display: false);
        attempt.Set("status", updated.Status).Set("bodyPlacement", updated.BodyPlacement);
        if (!string.Equals(updated.EntryId, entryId, StringComparison.OrdinalIgnoreCase))
        {
            Run.Allow(updated.EntryId);
        }
    }

    /// <summary>B1-B4: an inspector never activated, its editor read directly, the edit, and the variant's commit.</summary>
    private void EditWithoutActivate(Q149OutlookClient client, string variant, string entryId, string storeId, string token, Q149Attempt attempt)
    {
        client.Run<object?>((app, ns) =>
        {
            dynamic? item = null;
            dynamic? inspector = null;
            object? document = null;
            bool closed = false;
            try
            {
                item = ns.GetItemFromID(entryId, storeId);
                Q149Drafts.RequireOwnUnsentDraft(Run, item);
                inspector = item.GetInspector;
                if (variant == Q149Cells.B3)
                {
                    Q149Editor.AddMessageFormPage((object)inspector, attempt);
                }

                Q149EditorRead read = Q149Editor.ReadEditor((object)inspector, 2_000, 250);
                document = read.Document;
                attempt.Set("editorOk", read.Ok).Set("editorFailure", read.Failure).Set("editorPolls", read.Polls);
                if (document == null)
                {
                    return null;
                }

                Q149Editor.ReplaceText(document, token);
                switch (variant)
                {
                    case Q149Cells.B1:
                        dynamic current = inspector.CurrentItem;
                        try
                        {
                            current.Save();
                        }
                        finally
                        {
                            Q149Com.Release(current);
                        }

                        inspector.Close(LiveComposeProbe.OlDiscard);
                        break;
                    case Q149Cells.B2:
                    case Q149Cells.B3:
                        inspector.Close(LiveComposeProbe.OlSave);
                        break;
                    case Q149Cells.B4:
                        ((dynamic)document).Save();
                        inspector.Close(LiveComposeProbe.OlDiscard);
                        break;
                    default:
                        throw new InvalidOperationException("unknown variant " + variant);
                }

                closed = true;
                return null;
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                attempt.Set("commitFailure", Q149Com.Failure(ex));
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
    }
}
