using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Phase-3 T2 live acceptance for the show-me layer (v3.MD section 0.6 Phase 3):
/// goto_folder must land ActiveExplorer().CurrentFolder.FolderPath on the target,
/// open_in_outlook must yield an Inspector for the requested EntryID (closed again by
/// the test - closing a window the test itself opened is allowed and is not an Outlook
/// restart, S7), and show_search_results must drive the real search UI without error,
/// with the olSearchScope enum values feature-tested (v3.MD risk register) and ONE S5
/// screenshot of a test-hub-scoped result list as visual evidence (navigation pane and
/// to-do bar hidden during capture so no other store's folder names are in frame; the
/// capture is window-rect only). All UI work targets the test-hub store (S2/S5);
/// logging stays content-free for business stores (S4).
/// </summary>
[Collection(LiveCollections.Phase3)]
[Trait("Category", "Live")]
public sealed class LiveShowMeTests
{
    /// <summary>
    /// How long a second COM client may take to see the explorer the first one just moved.
    /// Short because this is UI settling, not a mail round trip.
    /// </summary>
    private const int ExplorerSettleSeconds = 10;

    private const int OlNavigationPane = 4;
    private const int OlToDoBar = 5;

    private readonly LivePhase3Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveShowMeTests(LivePhase3Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    [Fact]
    [Trait("LiveTier", "Portable")]
    [Trait("Requires", "InteractiveDesktop")]
    public void GotoFolder_HubFolder_ActiveExplorerCurrentFolderMatches()
    {
        // Locale-proof target: a real TOP-LEVEL folder path reported by list_folders
        // (the listing now always returns the full tree - filter to depth 1).
        FoldersOutcome folders = Service.ListFolders(Hub);
        StoreFoldersView store = Assert.Single(folders.Stores);
        FolderView target = store.Folders.FirstOrDefault(f => !f.Path.Contains('/') && (f.Items ?? 0) > 0)
            ?? store.Folders.First(f => !f.Path.Contains('/'));
        string expectedPath = "\\\\" + Hub + "\\" + target.Path.Replace('/', '\\');

        GotoFolderOutcome outcome = Service.GotoFolder(Hub, target.Path);
        Assert.True(outcome.Displayed);
        Assert.Equal(expectedPath, outcome.ExplorerFolderPath, ignoreCase: true);

        // Independent verification through a second COM client: the ACCEPTANCE assert.
        // Polled: right after a headless cold start creates the first window, a second
        // client can momentarily see no active explorer yet.
        ComExplorerState? state = null;
        string? error = null;
        LiveWaitBudget wait = LiveWaitBudget.OfSeconds(ExplorerSettleSeconds);
        while (wait.HasTimeLeft)
        {
            state = _fixture.VerifySession.TryGetActiveExplorerState(out error);
            if (state?.CurrentFolderPath != null
                && string.Equals(state.CurrentFolderPath, expectedPath, StringComparison.OrdinalIgnoreCase))
            {
                break;
            }

            Thread.Sleep(500);
        }

        Assert.True(state != null, $"active explorer state unavailable: {error}");
        Assert.Equal(expectedPath, state!.CurrentFolderPath, ignoreCase: true);
        _output.WriteLine($"goto_folder ok: folderPathMatches=true windowState={state.WindowState}");

        // Default navigation (no folder): lands on the hub's Inbox (or root) without error.
        GotoFolderOutcome defaultOutcome = Service.GotoFolder(Hub);
        Assert.True(defaultOutcome.Displayed);
        Assert.False(string.IsNullOrEmpty(defaultOutcome.ExplorerFolderPath));
        Assert.StartsWith("\\\\" + Hub, defaultOutcome.ExplorerFolderPath!, StringComparison.OrdinalIgnoreCase);
        _output.WriteLine("goto_folder default (Inbox) ok");
    }

    [Fact]
    [Trait("LiveTier", "ProfileBound")]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "InteractiveDesktop")]
    public void OpenInOutlook_HubMail_InspectorForRightEntryIdThenClosedByTest()
    {
        // A hub mail to display - prefer an already-read one (Display can mark unread
        // mail read; the hub grant S2 covers it, but why touch state needlessly).
        List<HitSummary> hits = Service.Search(new SearchRequest
        {
            IndexOnly = true,
            Store = Hub,
            IncludeAttachmentHits = false,
            Top = 20,
        }).Hits.Where(h => !string.IsNullOrEmpty(h.Subject) && !HubCorpus.IsTestArtifact(h.Subject)).ToList();
        Assert.True(hits.Count > 0, "no hub hits to display");
        HitSummary hit = hits.FirstOrDefault(h => h.IsRead == true) ?? hits[0];

        HashSet<string> baseline = new(
            _fixture.VerifySession.GetOpenInspectors().Where(i => i.EntryId != null).Select(i => i.EntryId!),
            StringComparer.OrdinalIgnoreCase);
        _output.WriteLine($"baseline inspectors={baseline.Count}");

        OpenInOutlookOutcome outcome = Service.OpenInOutlook(hit.Id);
        Assert.True(outcome.Displayed);
        Assert.True(outcome.EntryId.Length >= 48, "open_in_outlook must report the real EntryID");

        try
        {
            // ACCEPTANCE: an Inspector whose EntryID matches the requested item exists.
            ComInspectorInfo? inspector = PollForInspector(outcome.EntryId, present: true, TimeSpan.FromSeconds(15));
            Assert.True(inspector != null, "no Inspector appeared for the displayed EntryID within 15 s");
            Assert.False(baseline.Contains(outcome.EntryId), "test item was already open before the test");
            _output.WriteLine($"inspector ok: entryIdMatches=true itemClass={inspector!.ItemClass}");
        }
        finally
        {
            // The test closes the window it opened (olDiscard - nothing saved/prompted).
            bool closed = _fixture.VerifySession.TryCloseInspectorByEntryId(outcome.EntryId, out string? closeError);
            _output.WriteLine($"inspector close requested: ok={closed} err={closeError ?? "-"}");
        }

        ComInspectorInfo? stillOpen = PollForInspector(outcome.EntryId, present: false, TimeSpan.FromSeconds(10));
        Assert.True(stillOpen == null, "Inspector still open after the test closed it");
        _output.WriteLine("inspector closed and gone");
    }

    [Fact]
    [Trait("LiveTier", "Portable")]
    [Trait("Requires", "InteractiveDesktop")]
    public void ShowSearchResults_ScopeFeatureTest_AndHubScopedScreenshot()
    {
        // Park the window on the hub store first: current_folder/subfolders scopes then
        // apply to hub content only, and the later screenshot is S5-clean.
        Service.GotoFolder(Hub);

        // --- Feature test all four olSearchScope values (v3.MD risk register) with a
        // query that matches nothing, so no business-store content is put on screen.
        const string nonsense = "OutlookAiMcpNoSuchTerm7391";
        var findings = new List<string>();
        bool currentFolderWorks = false;
        foreach (string scope in new[] { "current_folder", "subfolders", "all_folders", "all_outlook" })
        {
            try
            {
                ShowSearchResultsOutcome probe = Service.ShowSearchResults(nonsense, scope);
                findings.Add($"{scope}=ok(displayed={probe.Displayed})");
                currentFolderWorks |= scope == "current_folder";
            }
            catch (Exception ex) when (ex is not Xunit.Sdk.XunitException)
            {
                findings.Add($"{scope}=FAIL({ex.GetType().Name}:{ex.Message})");
            }

            Thread.Sleep(400);
            _fixture.VerifySession.TryClearSearch(out _);
            Thread.Sleep(200);
        }

        _output.WriteLine("olSearchScope feature test: " + string.Join(" ", findings));
        Assert.True(currentFolderWorks, "olSearchScopeCurrentFolder must work: " + string.Join(" ", findings));

        // --- S5 screenshot case: a REAL hub-corpus term so the visible result list is
        // non-empty telefonie content. Hide the navigation pane + to-do bar first so
        // nothing outside the hub store is in frame; restore afterwards.
        IReadOnlyList<string> ranked = HubCorpus.RankedCleanTerms(_fixture.TestHubCorpus);
        Assert.True(ranked.Count > 0, "no corpus term derivable from the hub store");
        string term = ranked[0];
        _output.WriteLine($"screenshot query term='{term}' (hub corpus, matches>=1)");

        bool? navWasVisible = _fixture.VerifySession.TryGetExplorerPaneVisible(OlNavigationPane, out _);
        bool? todoWasVisible = _fixture.VerifySession.TryGetExplorerPaneVisible(OlToDoBar, out _);
        if (navWasVisible == true)
        {
            _fixture.VerifySession.TrySetExplorerPaneVisible(OlNavigationPane, false, out _);
        }

        if (todoWasVisible == true)
        {
            _fixture.VerifySession.TrySetExplorerPaneVisible(OlToDoBar, false, out _);
        }

        try
        {
            ShowSearchResultsOutcome shown = Service.ShowSearchResults(term, "current_folder", Hub);
            Assert.True(shown.Displayed);
            Assert.Equal("current_folder", shown.Scope);
            _output.WriteLine($"show_search_results ok: captionPresent={shown.ExplorerCaption != null} folderPresent={shown.ExplorerFolderPath != null}");

            // The UI search populates asynchronously - give it a moment, then record
            // what COM can see of the search state (enum findings for v3.MD 0.8).
            Thread.Sleep(4000);
            ComExplorerState? searchState = _fixture.VerifySession.TryGetActiveExplorerState(out _);
            _output.WriteLine($"post-search explorer: folderName='{searchState?.CurrentFolderName}' (search-results view swap observable={searchState?.CurrentFolderName != null})");

            // S5 evidence is best-effort (soak fix 19): the helper refuses to write a
            // capture it cannot prove is Outlook's own window.
            try
            {
                string path = ScreenCapture.CaptureOutlookWindow(
                    searchState?.Caption,
                    _fixture.ScreenshotsDirectory,
                    $"phase3-show-search-results-{DateTime.Now:yyyyMMdd-HHmmss}.png");

                var file = new FileInfo(path);
                Assert.True(file.Exists, "screenshot file must exist");
                Assert.True(file.Length > 0, "screenshot file must be non-empty");
                _output.WriteLine($"screenshot saved: {path} bytes={file.Length}");
            }
            catch (ScreenCaptureSkippedException ex)
            {
                _output.WriteLine($"S5 evidence skipped (no polluted capture written): {ex.Message}");
            }
        }
        finally
        {
            _fixture.VerifySession.TryClearSearch(out _);
            if (navWasVisible == true)
            {
                _fixture.VerifySession.TrySetExplorerPaneVisible(OlNavigationPane, true, out _);
            }

            if (todoWasVisible == true)
            {
                _fixture.VerifySession.TrySetExplorerPaneVisible(OlToDoBar, true, out _);
            }
        }
    }

    private ComInspectorInfo? PollForInspector(string entryId, bool present, TimeSpan timeout)
    {
        LiveWaitBudget wait = LiveWaitBudget.Of(timeout);
        ComInspectorInfo? last = null;
        while (wait.HasTimeLeft)
        {
            last = _fixture.VerifySession.GetOpenInspectors()
                .FirstOrDefault(i => i.EntryId != null && string.Equals(i.EntryId, entryId, StringComparison.OrdinalIgnoreCase));
            if ((last != null) == present)
            {
                return last;
            }

            Thread.Sleep(500);
        }

        return last;
    }
}
