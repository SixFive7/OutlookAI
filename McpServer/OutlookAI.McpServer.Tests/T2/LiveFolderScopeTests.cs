using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// T2 live acceptance for soak fix 15 - folder scope across all three tiers.
/// <list type="number">
/// <item><b>THE DELEGATE DEFECT (read-only).</b> Delegate mailboxes are indexed FLAT, so
/// the nested delegate URL the product used to build addressed a folder that does not
/// exist and every delegate SUBFOLDER search returned zero rows, silently. Proven here
/// against COM ground truth: the old shape still returns 0, the shipped shape returns
/// the folder's real population, for BOTH delegate mailboxes.</item>
/// <item><b>include_subfolders (hub writes).</b> Recursion and narrowing hold in the
/// freshness sweep and the exhaustive scan, and the sweep cache keeps the two apart.</item>
/// <item><b>Escaping + the non-silent zero guard.</b> An apostrophe in a folder name is
/// searchable instead of throwing; a folder path that resolves to nothing says so.</item>
/// </list>
/// SAFETY: the two delegate mailboxes are READ-ONLY here - counts, booleans and folder
/// paths only, never a subject, sender or body (S4). Every write targets the hub (S2),
/// carries the run tag + marker (S3) and is removed through the TESTED allowlist helpers.
/// </summary>
[Collection("LiveMoveArchive")]
[Trait("Category", "Live")]
public sealed class LiveFolderScopeTests
{
    private readonly LiveMoveArchiveFixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveFolderScopeTests(LiveMoveArchiveFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    private MailService Service => _fixture.Service;

    // ==================================================== 1. the delegate defect (read-only)

    [Fact]
    public void DelegateSubfolders_AreReachableAgain_AndTheOldNestedShapeStillReturnsZero()
    {
        IReadOnlyList<string> delegates = _fixture.Settings.ExpectedDelegateStoreDisplayNames;
        Assert.True(delegates.Count > 0, "the live settings must name at least one delegate store");

        IndexSearchService index = IndexSearchService.CreateDefault(out _);
        int provenSubfolders = 0;

        foreach (string delegateStore in delegates)
        {
            // COM ground truth: a NESTED folder (depth >= 2) of this delegate mailbox
            // that actually holds mail. Discovered at runtime - no mailbox identifier
            // is ever committed to this public repo (S6).
            string delegateRoot = ResolveDelegateRootScope(index, delegateStore);
            Assert.True(
                MapiItemUrl.TryBuildFolderPathDisplay(delegateRoot, out string? rootPath) && rootPath != null,
                "the delegate root scope must yield a folder display path");

            // Candidates are NESTED folders holding mail. Item counts alone are not
            // enough: a delegate mailbox's biggest subfolder can be a CALENDAR subtree,
            // whose items are not email rows at all - so a candidate is accepted only
            // once the shipped shape actually returns email for it.
            FolderView? nested = null;
            long comCount = 0;
            int newRows = 0;
            foreach (FolderView candidate in FolderTree(delegateStore)
                .Where(f => f.Path.Contains('/') && (f.Items ?? 0) >= 5)
                .OrderByDescending(f => f.Items ?? 0)
                .Take(12))
            {
                string candidateLeaf = candidate.Path[(candidate.Path.LastIndexOf('/') + 1)..];
                int rows = DrainCount(index, new IndexQuery
                {
                    Scope = delegateRoot,
                    FolderPathsAnyOf = new[] { rootPath + "/" + candidateLeaf },
                    Kinds = KindFilter.EmailOnly,
                    Top = 5000,
                });

                if (rows > 0)
                {
                    nested = candidate;
                    comCount = candidate.Items ?? 0;
                    newRows = rows;
                    break;
                }
            }

            Assert.True(nested != null, $"no reachable nested MAIL folder found in delegate store '{delegateStore}'");
            _output.WriteLine($"[{delegateStore}] nested folder '{nested!.Path}' COM items={comCount}");

            // --- (a) THE DEFECT, still reproducible: the pre-fix nested delegate URL.
            string oldShape = delegateRoot + "/" + nested.Path;
            int oldRows = DrainCount(index, new IndexQuery
            {
                Scope = oldShape,
                Kinds = KindFilter.EmailAndDocuments,
                Top = 5000,
            });
            Assert.Equal(0, oldRows);
            _output.WriteLine($"[{delegateStore}] pre-fix nested URL -> {oldRows} rows (the silent zero)");

            // --- (b) THE FIX: delegate store root + flat folder-name equality.
            _output.WriteLine($"[{delegateStore}] shipped shape -> {newRows} email rows (COM {comCount})");
            Assert.True(newRows > 0, $"the delegate subfolder is still unreachable ({delegateStore}/{nested.Path})");

            // The index census and COM disagree only by index lag (and, on a colliding
            // leaf name, by the merged folder's extra rows); a predicate error would be
            // an order-of-magnitude miss, not a few percent.
            double ratio = comCount == 0 ? 1 : newRows / (double)comCount;
            Assert.InRange(ratio, 0.70, 2.00);

            // --- (c) end to end through the product, with the flag both ways.
            foreach (bool includeSubfolders in new[] { false, true })
            {
                SearchOutcome outcome = Service.Search(new SearchRequest
                {
                    Store = delegateStore,
                    Folder = nested.Path,
                    IncludeSubfolders = includeSubfolders,
                    IndexOnly = true,
                    Top = 5,
                    SnippetChars = 0,
                });

                Assert.NotEmpty(outcome.Hits);
                Assert.All(outcome.Hits, h => Assert.Equal(delegateStore, h.Store));
                Assert.NotNull(outcome.Scope);
                Assert.Equal(nested.Path, outcome.Scope!.Folder);
                Assert.Equal(includeSubfolders, outcome.Scope.IncludeSubfolders);
                Assert.StartsWith("delegate_", outcome.Scope.Shape, StringComparison.Ordinal);

                _output.WriteLine(
                    $"[{delegateStore}] search include_subfolders={includeSubfolders}: hits={outcome.Hits.Count} "
                    + $"shape={outcome.Scope.Shape} widened={outcome.Scope.Widened} "
                    + $"folderNames={outcome.Scope.FolderNamesMatched}");
            }

            provenSubfolders++;
        }

        Assert.Equal(delegates.Count, provenSubfolders);
    }

    [Fact]
    public void DelegateFirstLevelFolders_StillResolve_AndTheWholeMailboxIsUnfiltered()
    {
        foreach (string delegateStore in _fixture.Settings.ExpectedDelegateStoreDisplayNames)
        {
            IReadOnlyList<FolderView> tree = FolderTree(delegateStore);
            FolderView? topLevel = tree
                .Where(f => !f.Path.Contains('/') && (f.Items ?? 0) >= 1)
                .OrderByDescending(f => f.Items ?? 0)
                .FirstOrDefault();
            Assert.True(topLevel != null, $"no populated first-level folder in '{delegateStore}'");

            SearchOutcome folderScoped = Service.Search(new SearchRequest
            {
                Store = delegateStore,
                Folder = topLevel!.Path,
                IncludeSubfolders = false,
                IndexOnly = true,
                Top = 3,
                SnippetChars = 0,
            });
            Assert.NotEmpty(folderScoped.Hits);

            // The whole delegate mailbox needs no folder filter at all - its root scope
            // already covers every flat folder.
            SearchOutcome wholeStore = Service.Search(new SearchRequest
            {
                Store = delegateStore,
                IndexOnly = true,
                Top = 3,
                SnippetChars = 0,
            });
            Assert.NotEmpty(wholeStore.Hits);
            Assert.Null(wholeStore.Scope);

            _output.WriteLine(
                $"[{delegateStore}] first-level '{topLevel.Path}' hits={folderScoped.Hits.Count}; "
                + $"whole-mailbox hits={wholeStore.Hits.Count}");
        }
    }

    // ============================================ 2. primary-store narrowing (read-only)

    [Fact]
    public void PrimaryStore_ExcludeSubfolders_NarrowsExactly_AndCostsNothing()
    {
        // A primary-store folder WITH children: recursive must return strictly more than
        // non-recursive, and the difference must be the children's own populations.
        string store = _fixture.Settings.ExpectedStoreDisplayNames
            .First(s => !string.Equals(s, Hub, StringComparison.OrdinalIgnoreCase));

        IReadOnlyList<FolderView> tree = FolderTree(store);
        IndexSearchService index = IndexSearchService.CreateDefault(out _);
        StoreScopeInfo scope = index.DiscoverStoreScopes(2000)
            .FirstOrDefault(s => string.Equals(s.StoreDisplayName, store, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"store scope for '{store}' not discovered");

        // A MAIL folder with populated children. Item counts alone would happily select
        // Contacts or Calendar, whose items are not email rows - so a candidate counts
        // only once its recursive scope actually returns rows AND the narrowed one
        // returns fewer. Smallest first: the assertion needs no big drain.
        FolderView? parent = null;
        FolderScopeResolution? recursive = null;
        FolderScopeResolution? shallow = null;
        int recursiveRows = 0;
        int shallowRows = 0;
        long recursiveMs = 0;
        long shallowMs = 0;

        foreach (FolderView candidate in tree
            .Where(f => (f.Items ?? 0) >= 5 && (f.Items ?? 0) <= 20000
                && tree.Any(c => c.Path.StartsWith(f.Path + "/", StringComparison.OrdinalIgnoreCase)
                    && (c.Items ?? 0) >= 1))
            .OrderBy(f => f.Items ?? 0)
            .Take(12))
        {
            FolderScopeResolution deep = FolderScopeResolver.ForPrimaryStore(scope.StorePrefix, candidate.Path, true);
            FolderScopeResolution own = FolderScopeResolver.ForPrimaryStore(scope.StorePrefix, candidate.Path, false);

            System.Diagnostics.Stopwatch probe = System.Diagnostics.Stopwatch.StartNew();
            int deepRows = DrainCount(index, new IndexQuery
            {
                Scope = deep.Scope, Kinds = KindFilter.EmailAndDocuments, Top = 5000,
            });
            long deepMs = probe.ElapsedMilliseconds;

            probe.Restart();
            int ownRows = DrainCount(index, new IndexQuery
            {
                Scope = own.Scope, FolderPathsAnyOf = own.FolderPaths, Kinds = KindFilter.EmailAndDocuments, Top = 5000,
            });
            long ownMs = probe.ElapsedMilliseconds;

            if (ownRows > 0 && ownRows < deepRows)
            {
                parent = candidate;
                recursive = deep;
                shallow = own;
                recursiveRows = deepRows;
                shallowRows = ownRows;
                recursiveMs = deepMs;
                shallowMs = ownMs;
                break;
            }
        }

        Assert.True(parent != null, $"no mail folder with populated children found in '{store}'");
        Assert.Null(recursive!.FolderPaths);
        Assert.NotNull(shallow!.FolderPaths);

        _output.WriteLine(
            $"[{store}] '{parent!.Path}' recursive={recursiveRows} ({recursiveMs} ms) "
            + $"nonRecursive={shallowRows} ({shallowMs} ms)");

        Assert.True(shallowRows < recursiveRows,
            $"non-recursive scope did not narrow ({shallowRows} vs {recursiveRows})");
        Assert.True(shallowRows > 0, "non-recursive scope returned nothing at all");

        // Attachment rows survive the narrowing - the reason DIRECTORY= is not used.
        int shallowDocs = DrainCount(index, new IndexQuery
        {
            Scope = shallow.Scope, FolderPathsAnyOf = shallow.FolderPaths, Kinds = KindFilter.DocumentsOnly, Top = 5000,
        });
        _output.WriteLine($"[{store}] '{parent.Path}' attachment-content rows inside the narrowed scope: {shallowDocs}");

        // Free, not merely affordable (measured -26 to +3 ms across all shapes).
        Assert.True(shallowMs <= recursiveMs + 250,
            $"the non-recursive predicate cost {shallowMs - recursiveMs} ms more than the recursive scope");
    }

    // ================================== 3. include_subfolders end to end (hub writes only)

    [Fact]
    public void HubSubtree_SweepAndExhaustive_HonorIncludeSubfolders_AndTheCacheKeepsThemApart()
    {
        LiveOutlookTestMailer.DeleteTestFolders(Hub);

        // The CHILD carries the test-folder prefix too: cleanup matches folders by name,
        // so an unprefixed child would survive its parent's removal and strand the tree.
        string parentFolder = _fixture.TestFolderName;
        string childFolder = parentFolder + "/" + _fixture.TestFolderName + "-Nested";
        string probeTerm = "sfxnest" + Marker;
        string seedSubject = _fixture.TaggedSubject("subtree seed " + probeTerm);
        string? entryId = null;
        try
        {
            // A DRAFT seed keeps this test independent of mail delivery AND of index
            // latency: both tiers exercised below read Outlook directly.
            DraftOutcome draft = Service.NewDraft(
                LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft"), to: Hub, cc: null, subject: seedSubject,
                body: "Folder-scope subtree probe " + probeTerm + " (soak fix 15).", display: false);
            entryId = draft.EntryId;
            Assert.False(string.IsNullOrEmpty(entryId));

            MoveMailOutcome moved = Service.MoveMail(new[] { entryId }, childFolder, createFolder: true);
            MoveItemView movedItem = Assert.Single(moved.Items);
            Assert.True(movedItem.Ok, movedItem.Error);
            entryId = movedItem.NewEntryId!;
            Assert.NotNull(moved.CreatedFolders);
            _output.WriteLine($"seed filed into '{childFolder}' (created: {string.Join(", ", moved.CreatedFolders!)})");

            // --- EXHAUSTIVE tier: recursion is the default, narrowing is explicit.
            SearchOutcome deepScan = Exhaustive(parentFolder, includeSubfolders: true);
            Assert.Contains(deepScan.Hits, h => h.Subject == seedSubject);
            Assert.NotNull(deepScan.Exhaustive);
            Assert.True(deepScan.Exhaustive!.FoldersScanned >= 2,
                $"a recursive exhaustive scan must reach the child folder (scanned {deepScan.Exhaustive.FoldersScanned})");
            Assert.NotNull(deepScan.Scope);
            Assert.Equal("folder", deepScan.Scope!.Shape);

            SearchOutcome shallowScan = Exhaustive(parentFolder, includeSubfolders: false);
            Assert.DoesNotContain(shallowScan.Hits, h => h.Subject == seedSubject);
            Assert.Equal(1, shallowScan.Exhaustive!.FoldersScanned);
            Assert.Equal("folder_only", shallowScan.Scope!.Shape);

            SearchOutcome childScan = Exhaustive(childFolder, includeSubfolders: false);
            Assert.Contains(childScan.Hits, h => h.Subject == seedSubject);
            _output.WriteLine(
                $"exhaustive: parent+subtree={deepScan.Hits.Count} hit(s) / {deepScan.Exhaustive.FoldersScanned} folders; "
                + $"parent only={shallowScan.Hits.Count} / {shallowScan.Exhaustive.FoldersScanned}; "
                + $"child only={childScan.Hits.Count}");

            // --- FRESHNESS SWEEP tier, and the cache key (constraint C6). Order matters:
            // the SHALLOW sweep runs FIRST and populates the cache, so if the flag were
            // missing from the key the recursive search below would be served the shallow
            // entry - and would not find the seed.
            Service.ClearSweepCache();
            SearchOutcome shallowSweep = FolderSearch(parentFolder, includeSubfolders: false, probeTerm);
            Assert.Equal("folder (no subfolders)", shallowSweep.Sweep!.Scope);
            Assert.Equal(1, shallowSweep.Sweep.FoldersSwept);
            Assert.DoesNotContain(shallowSweep.Hits, h => h.Subject == seedSubject);

            SearchOutcome recursiveSweep = FolderSearch(parentFolder, includeSubfolders: true, probeTerm);
            Assert.Equal("folder", recursiveSweep.Sweep!.Scope);
            Assert.True(recursiveSweep.Sweep.FoldersSwept >= 2,
                $"the recursive sweep did not walk the subtree (swept {recursiveSweep.Sweep.FoldersSwept}) - "
                + "a shallow cache entry may have answered a recursive query (C6)");
            Assert.False(recursiveSweep.Sweep.Cached == true,
                "a shallow sweep was served to a recursive query - include_subfolders is missing from the cache key");
            HitSummary live = Assert.Single(recursiveSweep.Hits, h => h.Subject == seedSubject);

            // The contract under test is the SCOPE (C6): the recursive query walks the
            // subtree and returns the seed, the shallow one does not. WHICH tier supplies
            // the row is a race - Windows Search sometimes indexes the seed before this
            // line runs, and then the merged hit is legitimately sourced "index". Pinning
            // "live" pinned the race, and lost it nondeterministically under full-suite
            // load. Either source proves the scope; the shallow asserts above and below
            // are what prove the subfolder flag is honored and cache-keyed.
            Assert.True(
                live.Source is "live" or "index",
                $"unexpected hit source '{live.Source}' for the seeded subfolder item");

            // And the reverse direction: the recursive entry must not answer the shallow
            // query it now precedes in the cache.
            SearchOutcome shallowAgain = FolderSearch(parentFolder, includeSubfolders: false, probeTerm);
            Assert.Equal(1, shallowAgain.Sweep!.FoldersSwept);
            Assert.DoesNotContain(shallowAgain.Hits, h => h.Subject == seedSubject);

            _output.WriteLine(
                $"sweep: shallow folders={shallowSweep.Sweep.FoldersSwept} hits={shallowSweep.Hits.Count}; "
                + $"recursive folders={recursiveSweep.Sweep.FoldersSwept} hits={recursiveSweep.Hits.Count} "
                + $"({recursiveSweep.Sweep.ElapsedMs} ms); cache keeps them apart in both directions");
        }
        finally
        {
            CleanUp(entryId);
        }

        AssertHubClean();
    }

    // ==================================== 4. escaping + the non-silent zero-row guard

    [Fact]
    public void ApostropheInAFolderName_IsSearchable_InsteadOfThrowing()
    {
        LiveOutlookTestMailer.DeleteTestFolders(Hub);

        // ValidateScope used to THROW on any scope containing an apostrophe, so this
        // whole shape was un-searchable by hard exception.
        string folder = _fixture.TestFolderName + "-O'Brien";
        string probeTerm = "sfxquote" + Marker;
        string seedSubject = _fixture.TaggedSubject("apostrophe seed " + probeTerm);
        string? entryId = null;
        try
        {
            DraftOutcome draft = Service.NewDraft(
                LiveStoreWriteGuard.Writable(Hub, StoreWriteKind.Draft, "new_draft"), to: Hub, cc: null, subject: seedSubject,
                body: "Apostrophe folder probe " + probeTerm + " (soak fix 15).", display: false);
            entryId = draft.EntryId;

            MoveMailOutcome moved = Service.MoveMail(new[] { entryId }, folder, createFolder: true);
            MoveItemView movedItem = Assert.Single(moved.Items);
            Assert.True(movedItem.Ok, movedItem.Error);
            entryId = movedItem.NewEntryId!;

            Service.ClearSweepCache();
            SearchOutcome outcome = Service.Search(new SearchRequest
            {
                Query = probeTerm,
                Store = Hub,
                Folder = folder,
                Top = 10,
                SnippetChars = 0,
            });

            // The index tier emitted a valid statement (no exception) and the sweep found
            // the item live, which is the freshness contract for a just-created item.
            Assert.NotNull(outcome.Sweep);
            Assert.True(outcome.Sweep!.Performed);
            Assert.Contains(outcome.Hits, h => h.Subject == seedSubject);

            // And the zero-row guard stays quiet: this folder IS new to the index, but
            // the answer is not empty - the guard judges the merged result, not the index
            // rows alone, or every just-created folder would be called unresolvable.
            Assert.DoesNotContain(
                outcome.Advice ?? Array.Empty<string>(),
                a => a.Contains("matched NOTHING in the index", StringComparison.Ordinal));
            _output.WriteLine($"apostrophe folder searched without throwing: hits={outcome.Hits.Count}, no false resolution advice");

            // Exhaustive too - a different escaping path (DASL, not WS-SQL).
            Assert.Contains(Exhaustive(folder, includeSubfolders: false).Hits, h => h.Subject == seedSubject);
        }
        finally
        {
            CleanUp(entryId);
        }

        AssertHubClean();
    }

    [Fact]
    public void UnresolvedFolderPath_IsReportedAsAResolutionProblem_NotAnEmptyResult()
    {
        Service.ClearSweepCache();
        SearchOutcome outcome = Service.Search(new SearchRequest
        {
            Query = "sfxabsent" + Marker,
            Store = Hub,
            Folder = "OutlookAI-McpTest-NoSuchFolder/Deeper",
            Top = 5,
            SnippetChars = 0,
            IndexOnly = true,
        });

        Assert.Empty(outcome.Hits);
        Assert.NotNull(outcome.Advice);
        Assert.Contains(outcome.Advice!, a => a.Contains("matched NOTHING in the index", StringComparison.Ordinal));
        _output.WriteLine("zero-row guard fired: an unresolvable folder path is named as such");

        // A REAL folder that simply holds no match must stay quiet - the guard must not
        // cry wolf on every empty result.
        SearchOutcome realFolder = Service.Search(new SearchRequest
        {
            Query = "sfxabsent" + Marker,
            Store = Hub,
            Folder = "Inbox",
            Top = 5,
            SnippetChars = 0,
            IndexOnly = true,
        });
        Assert.Empty(realFolder.Hits);
        Assert.DoesNotContain(
            realFolder.Advice ?? Array.Empty<string>(),
            a => a.Contains("matched NOTHING in the index", StringComparison.Ordinal));
        _output.WriteLine("an empty-but-valid folder produces no resolution advice");
    }

    [Fact]
    public void TopAboveTheCap_IsReportedAsAClamp()
    {
        SearchOutcome outcome = Service.Search(new SearchRequest
        {
            Store = Hub,
            Top = 500,
            SnippetChars = 0,
            IndexOnly = true,
        });

        Assert.NotNull(outcome.Advice);
        Assert.Contains(outcome.Advice!, a => a.Contains("was reduced to 100", StringComparison.Ordinal));
        _output.WriteLine("top clamp reported instead of silently applied");
    }

    // ------------------------------------------------------------------ helpers

    private SearchOutcome FolderSearch(string folder, bool includeSubfolders, string probeTerm)
    {
        return Service.Search(new SearchRequest
        {
            Query = probeTerm,
            Store = Hub,
            Folder = folder,
            IncludeSubfolders = includeSubfolders,
            Top = MailService.SearchTopCap,
            SnippetChars = 0,
        });
    }

    /// <summary>
    /// A TERM-LESS exhaustive folder scan. Deliberately term-less: the exhaustive tier
    /// matches terms with ci_phrasematch, which is index-backed, so a termed scan of a
    /// just-created seed would race the indexer. Without a term the scan enumerates the
    /// folder through Outlook directly, which is precisely the property under test -
    /// does the scan REACH this folder.
    /// </summary>
    private SearchOutcome Exhaustive(string folder, bool includeSubfolders)
    {
        return Service.Search(new SearchRequest
        {
            Store = Hub,
            Folder = folder,
            IncludeSubfolders = includeSubfolders,
            Exhaustive = true,
            Top = MailService.SearchTopCap,
            SnippetChars = 0,
        });
    }

    /// <summary>Flattened folder tree of one store (list_folders, all pages).</summary>
    private IReadOnlyList<FolderView> FolderTree(string store)
    {
        List<FolderView> all = new();
        int offset = 0;
        while (true)
        {
            FoldersOutcome page = Service.ListFolders(store, offset);
            foreach (StoreFoldersView view in page.Stores)
            {
                all.AddRange(view.Folders);
            }

            if (!page.Truncated || page.NextOffset is not int next || next <= offset)
            {
                return all;
            }

            offset = next;
        }
    }

    private static int DrainCount(IndexSearchService index, IndexQuery query)
    {
        return index.Search(query).Hits.Count;
    }

    /// <summary>
    /// The delegate mailbox's index root (<c>&lt;host&gt;/1/&lt;delegate&gt;</c>),
    /// discovered by probing every host store - delegate mailboxes hang off a HOST
    /// account, which is exactly the fact the naive display-name construction gets wrong.
    /// </summary>
    private static string ResolveDelegateRootScope(IndexSearchService index, string delegateStore)
    {
        foreach (StoreScopeInfo host in index.DiscoverStoreScopes(2000))
        {
            string candidate = host.StorePrefix + "/1/" + delegateStore;
            if (index.ScopeHasAnyItem(candidate))
            {
                return candidate;
            }
        }

        throw new InvalidOperationException($"no index root found for delegate store '{delegateStore}'");
    }

    private void CleanUp(string? entryId)
    {
        if (entryId != null)
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
        // Includes the Sync Issues subtree since soak fix 15 - a tagged artifact was once
        // found stranded in Local Failures, where nothing had ever swept.
        //
        // Straggler-tolerant like the move/archive suites (batch A): the stable-zero sweep
        // returns after its 10 s window, but under a full-suite load a seeded item can
        // still materialize after it - which failed this class nondeterministically, a
        // different test each run. Purge once more (S3-legal: tag AND this run's marker)
        // and assert only what SURVIVES that.
        // ONE purge pass is still not enough under full-suite load: a straggler can
        // materialize between the purge and the count that follows it (measured - this
        // class failed on exactly that, one surviving artifact, in an otherwise green
        // run). Drive to a STABLE zero first, the way the self-send suites do, and only
        // then apply the straggler-tolerant count.
        LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(
            Hub, Marker, folderIds: LiveOutlookTestMailer.HubSweepFolderIdsWithArchive);

        int remaining = LiveOutlookTestMailer.CountTaggedArtifactsAfterPurgingStragglers(
            Hub, Marker, LiveOutlookTestMailer.HubSweepFolderIdsWithArchive, out int stragglersPurged);
        if (stragglersPurged > 0)
        {
            _output.WriteLine($"cleanup[{Hub}]: {stragglersPurged} late-materialized artifact(s) purged (documented lag)");
        }

        Assert.Equal(0, remaining);
        Assert.Equal(0, LiveOutlookTestMailer.CountTestFolders(Hub));
        _output.WriteLine(_fixture.VerifyHubReconciled());
    }
}
