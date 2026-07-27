using System.Text.RegularExpressions;
using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// T2 live acceptance for soak fix 13 - the two recall fixes:
/// <list type="number">
/// <item><b>Cross-column AND.</b> A mail whose first query term appears only in the
/// SUBJECT and whose second appears only in the BODY must be found by
/// <c>query="A B"</c>, and must NOT be found once <c>search_in</c> narrows to one of
/// them. Proven on the index tier from the hub's own ground truth (read-only, terms
/// derived per item so a shifting corpus moves the expectation) and end-to-end on the
/// tool tier with a controlled corpus.</item>
/// <item><b>Sweep folder coverage.</b> The freshness sweep follows the SEARCH scope: a
/// folder-scoped search sweeps that folder (and its subfolders), so mail placed in a
/// NON-default folder - the rule-filed-on-arrival case - is found before the index has
/// it. The store-wide default set deliberately does NOT cover custom folders; that
/// residual gap is asserted here so it stays documented rather than assumed.</item>
/// </list>
/// All writes target the hub (S2), carry the run tag + marker (S3), are removed via the
/// TESTED allowlist helpers only, and the fixture's whole-store reconciliation runs
/// afterwards. Business stores are not touched. Logging is counts, timings and
/// self-authored markers only (S4).
/// </summary>
[Collection("LiveMoveArchive")]
[Trait("Category", "Live")]
public sealed class LiveSweepScopeTests
{
    private readonly LiveMoveArchiveFixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveSweepScopeTests(LiveMoveArchiveFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    private MailService Service => _fixture.Service;

    // ------------------------------------------------ fix 1: cross-column AND (index tier)

    [Fact]
    public void IndexTier_TwoTerms_OneInSubjectOneInBody_AreFoundTogether()
    {
        // Ground truth: a hub mail carrying a word in its subject that its own body does
        // not contain, plus a word in its body that its own subject does not contain.
        // Before this fix the index tier ANDed the terms inside each column, so such a
        // mail matched NOTHING - the terms never co-occur in one column.
        IndexSearchService index = IndexSearchService.CreateDefault(out _);
        StoreScopeInfo hub = ResolveStoreScope(index, Hub);
        (ComWalkedItem Item, string SubjectTerm, string BodyTerm) probe = SelectCrossColumnProbe(index, hub);

        _output.WriteLine(
            $"derived probe: subjectTerm.len={probe.SubjectTerm.Length} bodyTerm.len={probe.BodyTerm.Length} "
            + $"folder={probe.Item.FolderPath}");

        string[] bothTerms = { probe.SubjectTerm, probe.BodyTerm };

        // THE regression: both terms together, default scope.
        Assert.True(
            ContainsSubject(Rows(index, hub, bothTerms, SearchIn.SubjectAndBody), probe.Item.Subject),
            "the default scope did not find a mail with one term in the subject and the other in the body "
            + "(cross-column AND regression)");

        // Narrowing to one part must NOT find it - neither part holds both terms.
        Assert.False(ContainsSubject(Rows(index, hub, bothTerms, SearchIn.SubjectOnly), probe.Item.Subject));
        Assert.False(ContainsSubject(Rows(index, hub, bothTerms, SearchIn.BodyOnly), probe.Item.Subject));

        // Prefix stems keep working in both positions of a multi-term query.
        string[] stemmed = { Stem(probe.SubjectTerm), Stem(probe.BodyTerm) };
        Assert.True(
            ContainsSubject(Rows(index, hub, stemmed, SearchIn.SubjectAndBody), probe.Item.Subject),
            "prefix stems lost the cross-column match");
    }

    // -------------------------- fix 1 (tool tier) + fix 2, on ONE controlled corpus

    [Fact]
    public void ControlledCorpus_CrossColumnTermsMatch_AndTheSweepFollowsTheSearchScope()
    {
        LiveOutlookTestMailer.DeleteTestFolders(Hub);

        // One seed carries both regressions: its subject and body hold DISJOINT probe
        // terms (fix 1), and it then gets filed into a non-default folder the way a
        // server-side rule would (fix 2 - the live case is info@'s HAProxy alerts
        // landing in Deleted Items on arrival). One delivery, both contracts.
        string subjectTerm = "sfxsubj" + Marker;
        string bodyTerm = "sfxbody" + Marker;
        string seedSubject = _fixture.TaggedSubject("crosscolumn " + subjectTerm);
        string seedBody = "Cross-column and rule-filed probe " + bodyTerm + " (soak fix 13).";
        string? entryId = null;
        try
        {
            LiveOutlookTestMailer.SendSelfMail(Hub, seedSubject, seedBody, null);
            entryId = WaitForInboxSeed(seedSubject);
            _output.WriteLine("seed arrived in hub Inbox (subject term and body term are disjoint)");

            // --- fix 1, end to end: both terms together match, either part alone does not.
            string bothTerms = subjectTerm + " " + bodyTerm;
            IReadOnlyList<HitSummary> both = SearchHits(bothTerms, SearchInValues.Default);
            IReadOnlyList<HitSummary> subjectScope = SearchHits(bothTerms, SearchIn.SubjectOnly);
            IReadOnlyList<HitSummary> bodyScope = SearchHits(bothTerms, SearchIn.BodyOnly);

            _output.WriteLine($"tool tier: default={both.Count} subject={subjectScope.Count} body={bodyScope.Count}");
            Assert.Contains(both, h => h.Subject == seedSubject);
            Assert.DoesNotContain(subjectScope, h => h.Subject == seedSubject);
            Assert.DoesNotContain(bodyScope, h => h.Subject == seedSubject);

            // Each term alone still resolves in its own part.
            Assert.Contains(SearchHits(subjectTerm, SearchIn.SubjectOnly), h => h.Subject == seedSubject);
            Assert.Contains(SearchHits(bodyTerm, SearchIn.BodyOnly), h => h.Subject == seedSubject);

            // --- fix 2: file the seed into a folder no default sweep covers.
            string probeTerm = subjectTerm;
            MoveMailOutcome moved = Service.MoveMail(new[] { entryId }, _fixture.TestFolderName, createFolder: true);
            MoveItemView movedItem = Assert.Single(moved.Items);
            Assert.True(movedItem.Ok, movedItem.Error);
            entryId = movedItem.NewEntryId!;
            _output.WriteLine($"seed filed into non-default folder '{_fixture.TestFolderName}'");

            // A store-scoped search sweeps the arrival-path default set - NOT custom
            // folders. That residual gap is pinned here so it cannot change silently.
            Service.ClearSweepCache();
            SearchOutcome storeWide = Service.Search(new SearchRequest
            {
                Query = probeTerm,
                Store = Hub,
                Top = MailService.SearchTopCap,
                SnippetChars = 0,
            });
            Assert.NotNull(storeWide.Sweep);
            Assert.True(storeWide.Sweep!.Performed, "the store-wide sweep must run");
            Assert.Equal(MailService.DefaultSweepScopeDescription, storeWide.Sweep.Scope);
            Assert.True(storeWide.Sweep.FoldersSwept >= 3,
                $"the default set should cover the arrival-path folders, swept {storeWide.Sweep.FoldersSwept}");
            Assert.DoesNotContain(storeWide.Sweep.Folders ?? Array.Empty<string>(),
                f => f.EndsWith("/" + _fixture.TestFolderName, StringComparison.OrdinalIgnoreCase));
            _output.WriteLine(
                $"store-wide sweep: scope='{storeWide.Sweep.Scope}' folders={storeWide.Sweep.FoldersSwept} "
                + $"({storeWide.Sweep.ElapsedMs} ms) - custom folder not covered, as designed");

            // --- folder-scoped search: the sweep follows the scope and finds it live.
            Service.ClearSweepCache();
            SearchOutcome folderScoped = Service.Search(new SearchRequest
            {
                Query = probeTerm,
                Store = Hub,
                Folder = _fixture.TestFolderName,
                Top = MailService.SearchTopCap,
                SnippetChars = 0,
            });

            Assert.NotNull(folderScoped.Sweep);
            Assert.True(folderScoped.Sweep!.Performed, "the folder-scoped sweep must run");
            Assert.Equal("folder", folderScoped.Sweep.Scope);
            Assert.True(folderScoped.Sweep.FoldersSwept >= 1);
            Assert.NotNull(folderScoped.Sweep.Folders);
            Assert.Contains(folderScoped.Sweep.Folders!,
                f => f.Equals(Hub + "/" + _fixture.TestFolderName, StringComparison.OrdinalIgnoreCase));
            _output.WriteLine(
                $"folder-scoped sweep: scope='{folderScoped.Sweep.Scope}' folders={folderScoped.Sweep.FoldersSwept} "
                + $"({folderScoped.Sweep.ElapsedMs} ms) folders=[{string.Join(", ", folderScoped.Sweep.Folders!)}]");

            HitSummary hit = Assert.Single(folderScoped.Hits, h => h.Subject == seedSubject);
            Assert.Equal("live", hit.Source);
            _output.WriteLine("THE regression: rule-filed mail found by the folder-scoped freshness sweep");

            // --- cache-key correctness, live: the folder-scoped sweep just populated the
            // cache; a store-wide search within the TTL must NOT be served that narrow
            // result (it would report one folder of coverage as the whole gap).
            SearchOutcome afterScoped = Service.Search(new SearchRequest
            {
                Query = probeTerm,
                Store = Hub,
                Top = MailService.SearchTopCap,
                SnippetChars = 0,
            });
            Assert.Equal(MailService.DefaultSweepScopeDescription, afterScoped.Sweep!.Scope);
            Assert.True(afterScoped.Sweep.FoldersSwept >= 3,
                $"a narrow folder sweep leaked into a store-wide search (foldersSwept={afterScoped.Sweep.FoldersSwept})");
            Assert.DoesNotContain(afterScoped.Sweep.Folders ?? Array.Empty<string>(),
                f => f.EndsWith("/" + _fixture.TestFolderName, StringComparison.OrdinalIgnoreCase));
            _output.WriteLine("cache key separates the folder scope from the store-wide scope");

            // --- and the folder scope is itself cacheable (rapid iteration stays cheap).
            SearchOutcome repeat = Service.Search(new SearchRequest
            {
                Query = probeTerm,
                Store = Hub,
                Folder = _fixture.TestFolderName,
                Top = MailService.SearchTopCap,
                SnippetChars = 0,
            });
            Assert.Equal("folder", repeat.Sweep!.Scope);
            Assert.Contains(repeat.Hits, h => h.Subject == seedSubject);
            _output.WriteLine($"repeat folder-scoped search: cached={repeat.Sweep.Cached} ({repeat.Sweep.ElapsedMs} ms)");
        }
        finally
        {
            CleanUp(entryId);
        }

        int remaining = LiveOutlookTestMailer.CountTaggedArtifactsAfterPurgingStragglers(
            Hub, Marker, LiveOutlookTestMailer.HubSweepFolderIdsWithArchive, out int stragglers);
        if (stragglers > 0)
        {
            _output.WriteLine($"post-test: {stragglers} late-materialized self-send copy/copies purged (documented lag)");
        }

        Assert.Equal(0, remaining);
        Assert.Equal(0, LiveOutlookTestMailer.CountTestFolders(Hub));
        _output.WriteLine(_fixture.VerifyHubReconciled());
    }

    [Fact]
    public void FolderScopedSweep_UnknownFolder_DegradesWithAdvice_NeverThrows()
    {
        Service.ClearSweepCache();
        SearchOutcome outcome = Service.Search(new SearchRequest
        {
            Query = "sfxnothing" + Marker,
            Store = Hub,
            Folder = "OutlookAI-McpTest-NoSuchFolder",
            Top = 10,
            SnippetChars = 0,
        });

        Assert.NotNull(outcome.Sweep);
        Assert.True(outcome.Sweep!.Performed);
        Assert.Equal(0, outcome.Sweep.FoldersSwept);
        Assert.True(outcome.Sweep.FoldersSkipped >= 1);
        Assert.NotNull(outcome.Advice);
        Assert.Contains(outcome.Advice!, a => a.Contains("could not be opened in Outlook", StringComparison.Ordinal));
        _output.WriteLine($"unknown folder: swept={outcome.Sweep.FoldersSwept} skipped={outcome.Sweep.FoldersSkipped}, advice present");
    }

    [Fact]
    public void DefaultSweep_CoversTheArrivalPathFolders_WithinBudget()
    {
        // The default set is the freshness contract for every non-folder-scoped search:
        // Inbox, Sent Items, Deleted Items and Junk Email of each store in scope.
        ComSweepResult hubSweep = _fixture.VerifySession.SweepFoldersNewerThan(
            DateTime.UtcNow.AddMinutes(-15), perFolderCap: 50, includeBodies: false, onlyStoreDisplayName: Hub);

        Assert.Equal(OutlookComSession.DefaultSweepFolderKinds.Count, hubSweep.FoldersSwept + hubSweep.FoldersSkipped);
        Assert.Equal(hubSweep.FoldersSwept, hubSweep.SweptFolders.Count);
        Assert.All(hubSweep.SweptFolders, f => Assert.StartsWith(Hub + "/", f, StringComparison.OrdinalIgnoreCase));
        _output.WriteLine($"hub default sweep folders: [{string.Join(", ", hubSweep.SweptFolders)}] skipped={hubSweep.FoldersSkipped}");

        // All stores, timed: this runs on EVERY unscoped search, so it has to stay cheap.
        System.Diagnostics.Stopwatch clock = System.Diagnostics.Stopwatch.StartNew();
        ComSweepResult allStores = _fixture.VerifySession.SweepFoldersNewerThan(
            DateTime.UtcNow.AddMinutes(-15), perFolderCap: 50, includeBodies: false, onlyStoreDisplayName: null);
        clock.Stop();

        Assert.True(allStores.FoldersSwept >= hubSweep.FoldersSwept);
        _output.WriteLine(
            $"all-stores default sweep: {allStores.FoldersSwept} folders swept, {allStores.FoldersSkipped} skipped, "
            + $"{clock.ElapsedMilliseconds} ms");
        Assert.True(clock.ElapsedMilliseconds < 5000,
            $"the always-on default sweep took {clock.ElapsedMilliseconds} ms - far outside its measured budget");
    }

    // ------------------------------------------------------------------ helpers

    private IReadOnlyList<HitSummary> SearchHits(string query, SearchIn searchIn)
    {
        Service.ClearSweepCache();
        return Service.Search(new SearchRequest
        {
            Query = query,
            SearchIn = searchIn,
            Store = Hub,
            IncludeAttachmentHits = false,
            Top = MailService.SearchTopCap,
            SnippetChars = 0,
        }).Hits;
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

    /// <summary>
    /// Waits for the self-send's Inbox copy (index-independent hub walk). The bound is
    /// generous on purpose: delivery is a real round trip through the mail server and a
    /// slow one must not fail an unattended suite (measured typical: 6-40 s).
    /// </summary>
    private string WaitForInboxSeed(string seedSubject)
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(240);
        while (DateTime.UtcNow < deadline)
        {
            ComWalkedItem? seed = _fixture.VerifySession.WalkStoreMailItems(Hub).FirstOrDefault(i =>
                i.Subject == seedSubject
                && string.Equals(i.FolderPath, "Inbox", StringComparison.OrdinalIgnoreCase));
            if (seed != null)
            {
                return seed.EntryId;
            }

            Thread.Sleep(3000);
        }

        throw new TimeoutException("Seed mail did not arrive in the hub Inbox within 120 s.");
    }

    private static string Stem(string term)
    {
        return term.Length >= 5 ? term.Substring(0, term.Length - 2) + "*" : term;
    }

    private static bool ContainsSubject(IReadOnlyList<IndexHit> hits, string? subject)
    {
        return hits.Any(h => string.Equals(h.Subject, subject, StringComparison.Ordinal));
    }

    private static IReadOnlyList<IndexHit> Rows(
        IndexSearchService index, StoreScopeInfo scope, IReadOnlyList<string> terms, SearchIn searchIn)
    {
        return index.Search(new IndexQuery
        {
            Scope = scope.StorePrefix,
            Terms = terms,
            SearchIn = searchIn,
            Kinds = KindFilter.EmailOnly,
            Top = 5000,
        }).Hits;
    }

    private static StoreScopeInfo ResolveStoreScope(IndexSearchService index, string storeDisplayName)
    {
        StoreScopeInfo? scope = index.DiscoverStoreScopes(2000)
                .FirstOrDefault(s => string.Equals(s.StoreDisplayName, storeDisplayName, StringComparison.OrdinalIgnoreCase))
            ?? index.TryDiscoverStoreScopeByAddress(storeDisplayName);
        Assert.True(scope != null, "store scope not discoverable in the index");
        return scope!;
    }

    /// <summary>
    /// Picks a hub mail plus a (subjectTerm, bodyTerm) pair that the INDEX confirms is
    /// split across the two columns for that item: the subject term is found by a
    /// subject-scoped query and not by a body-scoped one, and vice versa. Ground truth
    /// (the COM walk) proposes candidates; the catalog confirms them, because index rows
    /// outlive deleted items (Phase-2 fact 9) and tokenization is not the substring rule.
    /// </summary>
    private (ComWalkedItem Item, string SubjectTerm, string BodyTerm) SelectCrossColumnProbe(
        IndexSearchService index, StoreScopeInfo hub)
    {
        IReadOnlyList<ComWalkedItem> corpus = _fixture.VerifySession.WalkStoreMailItems(Hub);
        Assert.True(corpus.Count > 0, "hub corpus is empty");

        var tokens = new Regex("[A-Za-z]{5,20}", RegexOptions.CultureInvariant);
        var subjectCounts = corpus
            .Where(i => i.Subject != null)
            .GroupBy(i => i.Subject!, StringComparer.Ordinal)
            .ToDictionary(g => g.Key, g => g.Count(), StringComparer.Ordinal);

        foreach (ComWalkedItem item in corpus)
        {
            string subject = item.Subject ?? string.Empty;
            string body = item.Body ?? string.Empty;
            if (subject.Length == 0 || body.Length == 0 || subjectCounts[subject] != 1)
            {
                continue; // identity below is by subject - it has to be unique
            }

            foreach (string subjectTerm in Distinct(tokens, subject, body))
            {
                foreach (string bodyTerm in Distinct(tokens, body, subject))
                {
                    if (!IsSplitConfirmedByIndex(index, hub, item.Subject, subjectTerm, bodyTerm))
                    {
                        continue;
                    }

                    return (item, subjectTerm, bodyTerm);
                }
            }
        }

        Assert.Fail("no hub mail offers an index-confirmed subject-term/body-term split - the corpus moved");
        return default;
    }

    private static IEnumerable<string> Distinct(Regex tokens, string ownField, string otherField)
    {
        return tokens.Matches(ownField)
            .Select(m => m.Value.ToLowerInvariant())
            .Distinct(StringComparer.Ordinal)
            .Where(t => otherField.IndexOf(t, StringComparison.OrdinalIgnoreCase) < 0)
            .Take(6);
    }

    private static bool IsSplitConfirmedByIndex(
        IndexSearchService index, StoreScopeInfo hub, string? subject, string subjectTerm, string bodyTerm)
    {
        return ContainsSubject(Rows(index, hub, new[] { subjectTerm }, SearchIn.SubjectOnly), subject)
            && !ContainsSubject(Rows(index, hub, new[] { subjectTerm }, SearchIn.BodyOnly), subject)
            && ContainsSubject(Rows(index, hub, new[] { bodyTerm }, SearchIn.BodyOnly), subject)
            && !ContainsSubject(Rows(index, hub, new[] { bodyTerm }, SearchIn.SubjectOnly), subject);
    }
}
