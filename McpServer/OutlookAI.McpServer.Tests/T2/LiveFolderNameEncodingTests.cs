using System.Diagnostics;
using System.Globalization;
using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// T2 live proof and permanent pin for FOLDER names holding the five characters Microsoft documents
/// as percent-encoded in a Windows Search MAPI URL - <c>% / \ * ?</c> (<see cref="MapiUrlSegment"/>;
/// the Q99 finding in <c>TODO.md</c>, 2026-10-03).
/// <para>
/// <b>What it measures, every run, before it asserts anything.</b> For each character: whether
/// Outlook accepts it in a folder name at all (the folder is made through the product's own
/// <c>move_mail</c> with <c>create_folder</c>, and the one name the product cannot express - one
/// holding <c>/</c> - through <see cref="LiveOutlookTestMailer.FileTaggedItemInNewTestFolder"/>);
/// how the index spells the folder in an item's URL and in <c>System.ItemFolderPathDisplay</c> and
/// <c>System.ItemPathDisplay</c>; the folder's own index row; which statement shapes find the item
/// - the raw-name scope the product used to build, the encoded one, and both spellings of the
/// display path; and what the product does end to end - an index-only folder search, recursive and
/// not, the folder the hit reports, and how <c>read</c> located it. One nested folder puts two
/// encoded segments in one path, and one name merely LOOKS like an escape (<c>%2A</c>), which only
/// a one-pass encoder and decoder get right.
/// </para>
/// <para>
/// <b>What it asserts.</b> For every name Outlook accepted: the URL spells it exactly as
/// <see cref="MapiUrlSegment.EncodePath"/> does; a folder-scoped search through the product finds
/// the item in the index tier, recursive and not, says nothing about an unresolved folder, and
/// reports the folder by its real name; and <c>read</c> opens it by the URL route. A name Outlook
/// refuses is reported, not failed - that is Outlook's property - but a run in which no name with
/// one of the characters got as far as the index proves nothing and fails.
/// </para>
/// <para>
/// SAFETY: every write targets the hub (S2), every item carries the tag and this run's marker (S3),
/// every folder carries <see cref="LiveOutlookTestMailer.TestFolderNamePrefix"/> and is made inside
/// this test's own parent folder, and all of it is removed through the tested helpers; the index is
/// only read.
/// </para>
/// </summary>
[Collection(LiveCollections.MoveArchive)]
[Trait("Category", "Live")]
public sealed class LiveFolderNameEncodingTests
{
    /// <summary>
    /// How long the run waits for the index to take every seeded item in its new folder. A
    /// CEILING: on the indexed guest the hub's own population is in the index within seconds
    /// (<c>Docs/live-tier-on-the-vm.md</c> section 4.2d, 16 s).
    /// </summary>
    private const int IndexWaitSeconds = 300;

    /// <summary>Gap between polls of that wait.</summary>
    private const int IndexPollSeconds = 3;

    /// <summary>
    /// Per-statement bound for the wait's polls, far below the client default: a lost poll costs
    /// one poll (<see cref="SeededCrawlPoll"/> says why that is the right trade for a poll loop).
    /// </summary>
    private const int StatementTimeoutSeconds = 15;

    private readonly LiveMoveArchiveFixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveFolderNameEncodingTests(LiveMoveArchiveFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    private MailService Service => _fixture.Service;

    /// <summary>This test's own parent folder - no special character, so it is addressable plainly.</summary>
    private static string Parent => LiveOutlookTestMailer.TestFolderNamePrefix + "-Enc";

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "MailAccount")]
    public void FolderNamesWithPercentEncodedCharacters_AreSearchedByFolder_AndReportedByTheirOwnNames()
    {
        LiveOutlookTestMailer.DeleteTestFolders(Hub);

        string runTerm = "fnenc" + Marker;
        List<EncodingCase> cases = BuildCases(runTerm);
        List<string> entryIds = new();
        try
        {
            foreach (EncodingCase c in cases)
            {
                Seed(c, runTerm, entryIds);
            }

            ReadBackFolderNames(cases);

            IIndexClient client = IndexClientFactory.CreateAuto(out string providerReport);
            _output.WriteLine("index client: " + providerReport);
            string? hubPrefix = WaitForIndex(client, runTerm, cases);
            if (hubPrefix != null)
            {
                ReportFolderRows(client, hubPrefix, cases);
                IndexSearchService index = new IndexSearchService(client);
                foreach (EncodingCase c in cases.Where(c => c.ItemUrl != null))
                {
                    MeasureStatements(index, hubPrefix, c);
                }

                foreach (EncodingCase c in cases.Where(c => c.ItemUrl != null))
                {
                    SearchThroughTheProduct(c);
                }
            }

            foreach (EncodingCase c in cases)
            {
                _output.WriteLine(c.Describe());
            }

            AssertMeasuredBehaviour(cases);
        }
        finally
        {
            CleanUp(entryIds);
        }

        AssertHubClean();
    }

    // ------------------------------------------------------------------ the cases

    /// <summary>
    /// One folder per character, one name mixing three of them, one name that only LOOKS like an
    /// escape, a nested folder - two encoded segments in one path - and the name only the helper
    /// can make. Every folder name carries the test-folder prefix, so the cleanup finds it at any
    /// depth; the percent folder comes before the nested one, which is made inside it.
    /// </summary>
    private static List<EncodingCase> BuildCases(string runTerm)
    {
        string prefix = LiveOutlookTestMailer.TestFolderNamePrefix;
        string percent = prefix + " 50% off";
        return new List<EncodingCase>
        {
            new("%", new[] { Parent, percent }, runTerm + "pct"),
            new("*", new[] { Parent, prefix + " star*" }, runTerm + "star"),
            new("?", new[] { Parent, prefix + " why?" }, runTerm + "qmark"),
            new("\\", new[] { Parent, prefix + " back\\slash" }, runTerm + "bslash"),
            new("mixed % * ?", new[] { Parent, prefix + " 100%*? mix" }, runTerm + "mixed"),
            new("escape look-alike", new[] { Parent, prefix + " %2A not a star" }, runTerm + "lookalike"),
            new("nested", new[] { Parent, percent, prefix + " inner*?" }, runTerm + "nested"),
            new("/", new[] { Parent, prefix + " a/b" }, runTerm + "slash"),
        };
    }

    /// <summary>
    /// A tagged draft, filed into the case's folder - by <c>move_mail</c> with
    /// <c>create_folder</c>, the product's own folder-creating write, or for a name holding
    /// <c>/</c> by the helper. Outlook refusing the name is recorded and the draft stays in
    /// Drafts, where the tagged sweep removes it.
    /// </summary>
    private void Seed(EncodingCase c, string runTerm, List<string> entryIds)
    {
        // The run's term finds every seed in one statement; the case's own term, last in the
        // subject, tells them apart and is what each search asks for.
        c.Subject = _fixture.TaggedSubject("folder-name encoding " + runTerm + " " + c.Term);
        DraftOutcome draft = Service.NewDraft(
            LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft"), to: Hub, cc: null, subject: c.Subject,
            body: "Folder-name encoding probe " + c.Term + " (Q99 finding).", display: false);
        Assert.False(string.IsNullOrEmpty(draft.EntryId));
        entryIds.Add(draft.EntryId);

        if (!c.NeedsHelper)
        {
            MoveMailOutcome moved = Service.MoveMail(new[] { draft.EntryId }, c.ProductPath, createFolder: true);
            MoveItemView item = Assert.Single(moved.Items);
            if (item.Ok && item.NewEntryId != null)
            {
                c.Placed = true;
                c.EntryId = item.NewEntryId;
                entryIds.Add(item.NewEntryId);
                c.Creation = "move_mail created [" + string.Join(", ", moved.CreatedFolders ?? Array.Empty<string>()) + "]";
            }
            else
            {
                c.Refusal = "move_mail: " + item.Error + " (outcome " + (item.Outcome ?? "-") + ")";
            }

            return;
        }

        TestFolderFiling filing = LiveOutlookTestMailer.FileTaggedItemInNewTestFolder(
            Hub, Marker, draft.EntryId, c.Segments.Take(c.Segments.Count - 1).ToList(), c.Segments[^1]);
        if (filing.FolderCreated && filing.NewItemEntryId != null)
        {
            c.Placed = true;
            c.EntryId = filing.NewItemEntryId;
            entryIds.Add(filing.NewItemEntryId);
            c.Creation = "Folders.Add made '" + filing.FolderName + "'";
            c.ActualLeafName = filing.FolderName;
        }
        else
        {
            c.Refusal = "Folders.Add refused: " + filing.Error;
        }
    }

    /// <summary>
    /// The name Outlook gave each folder, read back through <c>list_folders</c>: a name Outlook
    /// silently changed is a different measurement from one it kept.
    /// </summary>
    private void ReadBackFolderNames(List<EncodingCase> cases)
    {
        IReadOnlyList<FolderView> tree = LiveFolderProbe.FolderTree(Service, Hub);
        foreach (EncodingCase c in cases.Where(c => c.Placed))
        {
            string joined = string.Join("/", c.Segments);
            c.ListedByListFolders = tree.Any(f => string.Equals(f.Path, joined, StringComparison.Ordinal));
        }

        _output.WriteLine("list_folders under " + Parent + ": "
            + string.Join(" | ", tree.Where(f => f.Path.StartsWith(Parent + "/", StringComparison.Ordinal)).Select(f => f.Path)));
    }

    // ------------------------------------------------------------------ the index

    /// <summary>
    /// Polls the user's MAPI root for this run's items until each placed one has a row in its new
    /// folder (not the Drafts row it had first), or the budget is spent. Returns the hub's store
    /// prefix as the index spells it, read off the first such row, or null when none arrived.
    /// </summary>
    private string? WaitForIndex(IIndexClient client, string runTerm, List<EncodingCase> cases)
    {
        string userRoot = IndexSearchService.CurrentUserMapiRoot();
        string sql = "SELECT TOP 200 System.ItemUrl, System.Subject, System.ItemFolderPathDisplay, System.ItemPathDisplay, "
            + "System.ItemNameDisplay FROM SystemIndex WHERE SCOPE='" + userRoot + "' AND CONTAINS(System.Subject, '\""
            + runTerm + "\"')";
        string marker = "/0/" + Parent + "/";
        string? hubPrefix = null;
        int completed = 0;
        int lost = 0;
        Stopwatch waited = Stopwatch.StartNew();
        while (true)
        {
            Stopwatch statement = Stopwatch.StartNew();
            try
            {
                foreach (IReadOnlyDictionary<string, object?> row in client.ExecuteRows(sql, 200, StatementTimeoutSeconds))
                {
                    string? url = Text(row, "System.ItemUrl");
                    string? subject = Text(row, "System.Subject");
                    if (url == null || subject == null || url.IndexOf(marker, StringComparison.Ordinal) < 0)
                    {
                        continue;
                    }

                    EncodingCase? c = cases.FirstOrDefault(k => subject.EndsWith(" " + k.Term, StringComparison.Ordinal));
                    if (c == null || c.ItemUrl != null)
                    {
                        continue;
                    }

                    c.ItemUrl = url;
                    c.ItemFolderPathDisplay = Text(row, "System.ItemFolderPathDisplay");
                    c.ItemPathDisplay = Text(row, "System.ItemPathDisplay");
                    c.IndexedAfter = waited.Elapsed;
                    if (hubPrefix == null && MapiItemUrl.TryParse(url, out MapiItemUrl? parsed) && parsed != null)
                    {
                        hubPrefix = parsed.StorePrefix;
                    }
                }

                completed++;
            }
            catch (Exception ex) when (ex is System.Data.OleDb.OleDbException && statement.Elapsed.TotalSeconds >= StatementTimeoutSeconds - 1)
            {
                lost++;
            }

            if (cases.Where(c => c.Placed).All(c => c.ItemUrl != null) || waited.Elapsed.TotalSeconds >= IndexWaitSeconds)
            {
                break;
            }

            Thread.Sleep(TimeSpan.FromSeconds(IndexPollSeconds));
        }

        _output.WriteLine(
            $"index wait: {waited.Elapsed.TotalSeconds:F0} s, {completed} completed poll(s), {lost} lost to the "
            + $"{StatementTimeoutSeconds}s bound; {cases.Count(c => c.ItemUrl != null)} of {cases.Count(c => c.Placed)} placed item(s) "
            + "found in their folders; hub prefix " + (hubPrefix ?? "(none)"));
        return hubPrefix;
    }

    /// <summary>
    /// The folders' OWN rows: a shallow listing of this test's parent folder, and of the percent
    /// folder for the nested one - at the URL the parent listing measured, not at a spelling
    /// assumed here. Reported only; the item rows carry the assertions.
    /// </summary>
    private void ReportFolderRows(IIndexClient client, string hubPrefix, List<EncodingCase> cases)
    {
        string parentUrl = hubPrefix + "/0/" + Parent;
        List<string> childUrls = ListDirectory(client, parentUrl);
        EncodingCase percent = cases[0];
        string? percentRow = childUrls.FirstOrDefault(u =>
            u.EndsWith("/" + MapiUrlSegment.Encode(percent.Segments[^1]), StringComparison.Ordinal)
            || u.EndsWith("/" + percent.Segments[^1], StringComparison.Ordinal));
        if (percentRow != null)
        {
            ListDirectory(client, percentRow);
        }
    }

    private List<string> ListDirectory(IIndexClient client, string directoryUrl)
    {
        List<string> urls = new();
        string sql = "SELECT TOP 100 System.ItemUrl, System.ItemNameDisplay, System.ItemFolderPathDisplay, System.ItemPathDisplay "
            + "FROM SystemIndex WHERE DIRECTORY='" + directoryUrl.Replace("'", "''", StringComparison.Ordinal) + "'";
        try
        {
            foreach (IReadOnlyDictionary<string, object?> row in client.ExecuteRows(sql, 100, StatementTimeoutSeconds))
            {
                string? url = Text(row, "System.ItemUrl");
                if (url == null)
                {
                    continue;
                }

                urls.Add(url);
                _output.WriteLine(
                    "row under '" + directoryUrl + "': url=" + url + " | name=" + Text(row, "System.ItemNameDisplay")
                    + " | folderPathDisplay=" + Text(row, "System.ItemFolderPathDisplay")
                    + " | pathDisplay=" + Text(row, "System.ItemPathDisplay"));
            }
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            _output.WriteLine("listing '" + directoryUrl + "' failed: " + ex.GetType().Name + ": " + ex.Message);
        }

        _output.WriteLine("listing '" + directoryUrl + "': " + urls.Count + " row(s)");
        return urls;
    }

    /// <summary>
    /// The statement shapes, each run straight against the index with the case's own term: the
    /// raw-name scope the product used to build (recursive, and with its folder-path equality),
    /// the encoded scope, the encoded scope with the display path spelled as NAMES and as the
    /// URL spells it, and the product's own resolution as it now stands.
    /// </summary>
    private void MeasureStatements(IndexSearchService index, string hubPrefix, EncodingCase c)
    {
        Assert.True(MapiItemUrl.TryParse(hubPrefix, out MapiItemUrl? hub) && hub != null);
        string rawPath = string.Join("/", c.Segments);
        string encodedPath = string.Join("/", c.Segments.Select(MapiUrlSegment.Encode));
        string rawScope = hubPrefix + "/0/" + rawPath;
        string encodedScope = hubPrefix + "/0/" + encodedPath;
        string displayNames = "/" + hub!.StoreDisplayName + "/" + rawPath;
        int hash = hub.StoreSegment.LastIndexOf("($", StringComparison.Ordinal);
        string hubAsSpelled = hash > 0 ? hub.StoreSegment[..hash].TrimEnd() : hub.StoreSegment;
        string displayUrlSpelling = "/" + hubAsSpelled + "/" + encodedPath;

        c.RawRecursive = Count(index, rawScope, null, c.Term);
        c.RawNonRecursive = Count(index, rawScope, new[] { displayNames }, c.Term);
        c.EncodedRecursive = Count(index, encodedScope, null, c.Term);
        c.EncodedWithDisplayNames = Count(index, encodedScope, new[] { displayNames }, c.Term);
        c.EncodedWithDisplayUrlSpelling = displayUrlSpelling == displayNames
            ? "(same as names)"
            : Count(index, encodedScope, new[] { displayUrlSpelling }, c.Term);

        if (!c.NeedsHelper)
        {
            FolderScopeResolution deep = FolderScopeResolver.ForPrimaryStore(hubPrefix, c.ProductPath, true);
            FolderScopeResolution own = FolderScopeResolver.ForPrimaryStore(hubPrefix, c.ProductPath, false);
            c.ResolverRecursive = Count(index, deep.Scope!, deep.FolderPaths, c.Term) + " via " + deep.Scope;
            c.ResolverNonRecursive = Count(index, own.Scope!, own.FolderPaths, c.Term)
                + " via " + own.Scope + " + " + string.Join(",", own.FolderPaths ?? Array.Empty<string>());
        }
    }

    private static string Count(IndexSearchService index, string scope, IReadOnlyList<string>? folderPaths, string term)
    {
        try
        {
            IndexSearchResult result = index.Search(new IndexQuery
            {
                Scope = scope,
                FolderPathsAnyOf = folderPaths,
                Terms = new[] { term },
                Kinds = KindFilter.MessagesAndAttachments,
                Top = 10,
            });
            return result.Hits.Count.ToString(CultureInfo.InvariantCulture);
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            return "THREW " + ex.GetType().Name + ": " + ex.Message;
        }
    }

    // ------------------------------------------------------------------ the product

    /// <summary>
    /// The product end to end, index tier only - the sweep would find a just-filed item whatever
    /// the scope, which is exactly how a wrong scope stayed invisible: a folder search, recursive
    /// and not, the folder the hit reports, the zero-row guard, and how <c>read</c> located it.
    /// A name holding <c>/</c> cannot be ASKED for - every product folder path splits on it - so
    /// its item is searched for through its parent, and the direct request is reported.
    /// </summary>
    private void SearchThroughTheProduct(EncodingCase c)
    {
        if (c.NeedsHelper)
        {
            c.Recursive = Probe(c, Parent, includeSubfolders: true);
            c.DirectRequestOfASlashName = Probe(c, c.ProductPath, includeSubfolders: false);
            return;
        }

        c.Recursive = Probe(c, c.ProductPath, includeSubfolders: true);
        c.NonRecursive = Probe(c, c.ProductPath, includeSubfolders: false);

        if (c.Segments.Count > 2)
        {
            // The nested item from its PARENT's folder: in scope with subfolders, out without.
            string parentPath = string.Join("/", c.Segments.Take(c.Segments.Count - 1));
            c.FromParentWithSubfolders = Probe(c, parentPath, includeSubfolders: true);
            c.FromParentWithoutSubfolders = Probe(c, parentPath, includeSubfolders: false);
        }
    }

    private ProductProbe Probe(EncodingCase c, string folder, bool includeSubfolders)
    {
        SearchOutcome outcome;
        try
        {
            outcome = Service.Search(new SearchRequest
            {
                Query = c.Term,
                Store = Hub,
                Folder = folder,
                IncludeSubfolders = includeSubfolders,
                IndexOnly = true,
                Top = 10,
                SnippetChars = 0,
            });
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            return new ProductProbe(folder, includeSubfolders, false, null, null, null, false, "THREW " + ex.GetType().Name + ": " + ex.Message);
        }

        bool unresolved = (outcome.Advice ?? Array.Empty<string>())
            .Any(a => a.Contains("matched NOTHING in the index", StringComparison.Ordinal));
        HitSummary? hit = outcome.Hits.FirstOrDefault(h => h.Subject == c.Subject);
        if (hit == null)
        {
            return new ProductProbe(folder, includeSubfolders, false, null, null, null, unresolved, null);
        }

        string? locatedVia;
        try
        {
            locatedVia = Service.Read(hit.Id, maxBodyChars: 0).LocatedVia;
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            locatedVia = "THREW " + ex.GetType().Name + ": " + ex.Message;
        }

        return new ProductProbe(folder, includeSubfolders, true, hit.Folder, hit.Source, locatedVia, unresolved, null);
    }

    // ------------------------------------------------------------------ the verdict

    private void AssertMeasuredBehaviour(List<EncodingCase> cases)
    {
        List<EncodingCase> exercised = cases.Where(c => c.ItemUrl != null).ToList();
        Assert.True(
            exercised.Count > 0,
            "PROVED NOTHING: no folder name holding one of % / \\ * ? got as far as the index - "
            + string.Join("; ", cases.Select(c => c.Label + ": " + (c.Refusal ?? (c.Placed ? "placed, never indexed" : "not placed")))));

        List<string> problems = new();
        foreach (EncodingCase c in cases.Where(c => c.Placed && c.ItemUrl == null))
        {
            problems.Add(c.Label + ": filed in its folder but not in the index after " + IndexWaitSeconds + " s");
        }

        foreach (EncodingCase c in exercised)
        {
            string expectedSpelling = string.Join("/", c.Segments.Select(MapiUrlSegment.Encode));
            if (!string.Equals(c.UrlFolderSpelling, expectedSpelling, StringComparison.Ordinal))
            {
                problems.Add(c.Label + ": the index spells the folder '" + c.UrlFolderSpelling + "', the encoder '" + expectedSpelling + "'");
            }

            string reportedFolder = string.Join("/", c.Segments);
            Check(problems, c, "recursive", c.Recursive, reportedFolder);
            if (!c.NeedsHelper)
            {
                Check(problems, c, "non-recursive", c.NonRecursive, reportedFolder);
            }

            if (c.FromParentWithSubfolders != null)
            {
                Check(problems, c, "from its parent with subfolders", c.FromParentWithSubfolders, reportedFolder);
                if (c.FromParentWithoutSubfolders?.Found == true)
                {
                    problems.Add(c.Label + ": found from its parent WITHOUT subfolders - the non-recursive bound did not hold");
                }
            }
        }

        Assert.True(problems.Count == 0, string.Join(Environment.NewLine, problems));
    }

    private static void Check(List<string> problems, EncodingCase c, string shape, ProductProbe? probe, string reportedFolder)
    {
        if (probe == null || !probe.Found)
        {
            problems.Add(c.Label + ": the product's " + shape + " folder search did not find the item (" + probe + ")");
            return;
        }

        if (!string.Equals(probe.HitFolder, reportedFolder, StringComparison.Ordinal))
        {
            problems.Add(c.Label + ": " + shape + " hit reports folder '" + probe.HitFolder + "', expected '" + reportedFolder + "'");
        }

        if (probe.Unresolved)
        {
            problems.Add(c.Label + ": " + shape + " search called the folder unresolved although it found the item");
        }

        if (!string.Equals(probe.LocatedVia, "urlSegments", StringComparison.Ordinal))
        {
            problems.Add(c.Label + ": " + shape + " hit was located via '" + probe.LocatedVia + "', not the URL's own folder path");
        }
    }

    // ------------------------------------------------------------------ cleanup (as LiveFolderScopeTests)

    private void CleanUp(List<string> entryIds)
    {
        foreach (string entryId in entryIds.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            try
            {
                LiveOutlookTestMailer.DeleteItemByEntryId(Hub, entryId, Marker);
            }
            catch (Exception)
            {
                // The stable-zero sweep below is the authority.
            }
        }

        LiveOutlookTestMailer.DeleteTestFolders(Hub);
        LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(
            Hub, Marker, folderIds: LiveOutlookTestMailer.HubSweepFolderIdsWithArchive);
    }

    private void AssertHubClean()
    {
        LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(
            Hub, Marker, folderIds: LiveOutlookTestMailer.HubSweepFolderIdsWithArchive);

        int remaining = LiveOutlookTestMailer.CountTaggedArtifactsAfterPurgingStragglers(
            Hub, Marker, LiveOutlookTestMailer.HubSweepFolderIdsWithArchive, out int stragglersPurged);
        if (stragglersPurged > 0)
        {
            _output.WriteLine($"cleanup[{Hub}]: {stragglersPurged} late-materialized artifact(s) purged (documented lag)");
        }

        Assert.Equal(0, remaining);

        int liveTestFolders = LiveOutlookTestMailer.CountLiveTestFolders(Hub, out int wedgedEmpty);
        if (wedgedEmpty > 0)
        {
            _output.WriteLine(
                $"cleanup[{Hub}]: {wedgedEmpty} empty test folder(s) wedged in Deleted Items until Outlook restarts "
                + "(documented same-session limitation, no items involved)");
        }

        Assert.Equal(0, liveTestFolders);
        _output.WriteLine(_fixture.VerifyHubReconciled());
    }

    private static string? Text(IReadOnlyDictionary<string, object?> row, string column)
    {
        return row.TryGetValue(column, out object? value) && value is string s ? s : null;
    }

    /// <summary>One product search: whether the item came back, as what, and how read located it.</summary>
    private sealed record ProductProbe(
        string Folder, bool IncludeSubfolders, bool Found, string? HitFolder, string? Source, string? LocatedVia,
        bool Unresolved, string? Error)
    {
        public override string ToString()
        {
            return "folder='" + Folder + "' subfolders=" + IncludeSubfolders + ": "
                + (Error ?? (Found
                    ? "FOUND, hit folder='" + HitFolder + "' source=" + Source + " locatedVia=" + LocatedVia
                    : "not found"))
                + (Unresolved ? " [advice: folder matched NOTHING in the index]" : string.Empty);
        }
    }

    /// <summary>One folder name under test, and everything measured about it.</summary>
    private sealed class EncodingCase
    {
        public EncodingCase(string label, IReadOnlyList<string> segments, string term)
        {
            Label = label;
            Segments = segments;
            Term = term;
        }

        public string Label { get; }

        /// <summary>Folder names from the store root down, as Outlook is asked to name them.</summary>
        public IReadOnlyList<string> Segments { get; }

        public string Term { get; }

        /// <summary>A name holding '/' cannot be a product folder path; the helper makes it.</summary>
        public bool NeedsHelper => Segments.Any(s => s.IndexOf('/') >= 0);

        /// <summary>The path an agent would pass as <c>folder</c>.</summary>
        public string ProductPath => string.Join("/", Segments);

        public string Subject { get; set; } = string.Empty;

        public string? EntryId { get; set; }

        public bool Placed { get; set; }

        public string? Creation { get; set; }

        public string? Refusal { get; set; }

        public string? ActualLeafName { get; set; }

        public bool? ListedByListFolders { get; set; }

        public string? ItemUrl { get; set; }

        public string? ItemFolderPathDisplay { get; set; }

        public string? ItemPathDisplay { get; set; }

        public TimeSpan? IndexedAfter { get; set; }

        public string? RawRecursive { get; set; }

        public string? RawNonRecursive { get; set; }

        public string? EncodedRecursive { get; set; }

        public string? EncodedWithDisplayNames { get; set; }

        public string? EncodedWithDisplayUrlSpelling { get; set; }

        public string? ResolverRecursive { get; set; }

        public string? ResolverNonRecursive { get; set; }

        public ProductProbe? Recursive { get; set; }

        public ProductProbe? NonRecursive { get; set; }

        public ProductProbe? FromParentWithSubfolders { get; set; }

        public ProductProbe? FromParentWithoutSubfolders { get; set; }

        public ProductProbe? DirectRequestOfASlashName { get; set; }

        /// <summary>
        /// The folder path exactly as the index spells it in the item's URL: everything between
        /// <c>/0/</c> and the encoded EntryID.
        /// </summary>
        public string? UrlFolderSpelling
        {
            get
            {
                if (ItemUrl == null)
                {
                    return null;
                }

                int start = ItemUrl.IndexOf("/0/", StringComparison.Ordinal);
                int end = ItemUrl.LastIndexOf('/');
                return start < 0 || end <= start + 3 ? null : ItemUrl.Substring(start + 3, end - start - 3);
            }
        }

        public string Describe()
        {
            return "case '" + Label + "': requested [" + string.Join(" > ", Segments) + "]"
                + (Placed
                    ? " -> " + Creation + (ActualLeafName != null ? " (actual name '" + ActualLeafName + "')" : string.Empty)
                        + "; list_folders lists it: " + (ListedByListFolders?.ToString() ?? "-")
                    : " -> REFUSED: " + Refusal)
                + Environment.NewLine + "    url folder spelling: " + (UrlFolderSpelling ?? "(no row)")
                + (IndexedAfter.HasValue ? " (in the index after " + IndexedAfter.Value.TotalSeconds.ToString("F0", CultureInfo.InvariantCulture) + " s)" : string.Empty)
                + Environment.NewLine + "    item url: " + (ItemUrl ?? "-")
                + Environment.NewLine + "    ItemFolderPathDisplay: " + (ItemFolderPathDisplay ?? "-")
                + Environment.NewLine + "    ItemPathDisplay: " + (ItemPathDisplay ?? "-")
                + Environment.NewLine + "    statements: raw scope " + (RawRecursive ?? "-") + ", raw scope + names " + (RawNonRecursive ?? "-")
                + ", encoded scope " + (EncodedRecursive ?? "-") + ", encoded + display as names " + (EncodedWithDisplayNames ?? "-")
                + ", encoded + display as url spelling " + (EncodedWithDisplayUrlSpelling ?? "-")
                + Environment.NewLine + "    resolver: recursive " + (ResolverRecursive ?? "-") + "; non-recursive " + (ResolverNonRecursive ?? "-")
                + Environment.NewLine + "    product: recursive " + (Recursive?.ToString() ?? "-")
                + "; non-recursive " + (NonRecursive?.ToString() ?? "-")
                + (FromParentWithSubfolders != null ? "; from parent with subfolders " + FromParentWithSubfolders : string.Empty)
                + (FromParentWithoutSubfolders != null ? "; from parent without subfolders " + FromParentWithoutSubfolders : string.Empty)
                + (DirectRequestOfASlashName != null ? "; asked for directly " + DirectRequestOfASlashName : string.Empty);
        }
    }
}
