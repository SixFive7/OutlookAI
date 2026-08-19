using System.Diagnostics;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Phase-3 T2 live acceptance for exhaustive search (v3.MD section 0.6, D34 boolean): for control
/// queries on the test-hub store, the COM scan (index path bypassed by construction:
/// exhaustive never queries the SystemIndex for hits and its results carry real
/// EntryIDs from birth) must return the same known-answer set as the index. Identity
/// comparison is EntryID-level: walk ground truth vs index hits located via HitLocator
/// vs exhaustive hits' EntryIDs surfaced through read (cache-only, no locate). Only the
/// tiny hub store is scanned (S2); logging is counts/timings/terms only (S4).
/// </summary>
[Collection(LiveCollections.Phase3)]
[Trait("Category", "Live")]
public sealed class LiveExhaustiveSearchTests
{
    private const int TimeToleranceSeconds = 5;

    private readonly LivePhase3Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveExhaustiveSearchTests(LivePhase3Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    [Fact]
    [Trait("LiveTier", "ProfileBound")]
    [Trait("Requires", "SearchIndex")]
    public void Exhaustive_KnownAnswer_MatchesIndexAndGroundTruth_OnHubStore()
    {
        IReadOnlyList<OutlookAI.Core.Com.ComWalkedItem> corpus = _fixture.TestHubCorpus;
        Assert.True(corpus.Count > 0, "hub corpus is empty");

        IReadOnlyList<string> ranked = HubCorpus.RankedCleanTerms(corpus);
        Assert.True(ranked.Count >= 2, "not enough clean corpus terms");

        // Control terms from both ends of the frequency ranking (oracle pattern).
        var controlTerms = new List<string> { ranked[0], ranked[ranked.Count - 1] };

        // Index side runs subject+body parity mode directly against the SystemIndex.
        IndexSearchService index = IndexSearchService.CreateDefault(out _);
        StoreScopeInfo? hubScope = index.DiscoverStoreScopes(2000)
            .FirstOrDefault(s => string.Equals(s.StoreDisplayName, Hub, StringComparison.OrdinalIgnoreCase))
            ?? index.TryDiscoverStoreScopeByAddress(Hub);
        Assert.True(hubScope != null, "hub store scope not discoverable in the index");

        foreach (string term in controlTerms)
        {
            VerifyControlTerm(term, corpus, index, hubScope!);
        }
    }

    private void VerifyControlTerm(
        string term,
        IReadOnlyList<OutlookAI.Core.Com.ComWalkedItem> corpus,
        IndexSearchService index,
        StoreScopeInfo hubScope)
    {
        System.Text.RegularExpressions.Regex word = HubCorpus.WordRegex(term);
        var expected = new HashSet<string>(
            corpus.Where(i => word.IsMatch(HubCorpus.TextOf(i))).Select(i => i.EntryId),
            StringComparer.OrdinalIgnoreCase);
        Assert.True(expected.Count > 0, $"term '{term}' has no ground-truth matches");

        // --- Index known-answer set (real EntryIDs via HitLocator).
        Stopwatch indexClock = Stopwatch.StartNew();
        IndexSearchResult indexResult = index.Search(new IndexQuery
        {
            Scope = hubScope.StorePrefix,
            Kinds = KindFilter.MailKindOnly,
            Terms = new[] { term },
            SearchIn = SearchIn.SubjectAndBody,
            Top = 500,
        });
        indexClock.Stop();

        var indexIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (IndexHit hit in indexResult.Hits.Where(h => h.DateReceivedUtc.HasValue && !HubCorpus.IsTestArtifact(h.Subject)))
        {
            OutlookAI.Core.Com.HitLocationResult location =
                OutlookAI.Core.Com.HitLocator.Locate(_fixture.VerifySession, hit, TimeToleranceSeconds);
            if (location.Located != null)
            {
                indexIds.Add(location.Located.EntryId);
            }
        }

        // --- Exhaustive set: date-bounded store scan (the required bound), index bypassed.
        Stopwatch exhaustiveClock = Stopwatch.StartNew();
        SearchOutcome outcome = Service.Search(new SearchRequest
        {
            Exhaustive = true,
            Store = Hub,
            Query = term,
            AfterUtc = new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Utc),
            Top = 100,
        });
        exhaustiveClock.Stop();

        Assert.Equal(0, outcome.IndexElapsedMs); // structural: the index path is bypassed
        Assert.Null(outcome.Sweep);
        Assert.NotNull(outcome.Exhaustive); // D34: the exhaustive block IS the mode marker
        Assert.False(outcome.Exhaustive!.Truncated, "hub scan must not hit the result cap");
        Assert.False(outcome.Exhaustive.TimedOut, "hub scan must not hit the time budget");
        Assert.All(outcome.Hits, h => Assert.Equal("exhaustive", h.Source));

        // Exhaustive hits carry real EntryIDs from birth - read serves them from cache.
        var exhaustiveIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (HitSummary hit in outcome.Hits.Where(h => !HubCorpus.IsTestArtifact(h.Subject)))
        {
            ReadOutcome read = Service.Read(hit.Id, maxBodyChars: 0);
            Assert.Equal("cached", read.LocatedVia);
            exhaustiveIds.Add(read.EntryId);
        }

        _output.WriteLine(
            $"term '{term}': groundTruth={expected.Count} index={indexIds.Count} (queryMs={indexResult.ElapsedMilliseconds} "
            + $"totalMs={indexClock.ElapsedMilliseconds}) exhaustive={exhaustiveIds.Count} "
            + $"(scanMs={outcome.Exhaustive.ElapsedMs} totalMs={exhaustiveClock.ElapsedMilliseconds} "
            + $"engine={outcome.Exhaustive.Engine} instantSearch={outcome.Exhaustive.InstantSearchEnabled} "
            + $"folders={outcome.Exhaustive.FoldersScanned} skipped={outcome.Exhaustive.FoldersSkipped})");

        // ACCEPTANCE: exhaustive == index known-answer set (EntryID-level).
        AssertSetsEqual(indexIds, exhaustiveIds, term, "index", "exhaustive");

        // And both match the walk ground truth (any index shortfall would have to be
        // staleness-explained; the idle hub store has been exact since Phase 1).
        AssertSetsEqual(expected, exhaustiveIds, term, "groundTruth", "exhaustive");
    }

    [Fact]
    [Trait("LiveTier", "Portable")]
    public void Exhaustive_FolderBounded_ReturnsExactlyThatFoldersMatches()
    {
        IReadOnlyList<OutlookAI.Core.Com.ComWalkedItem> corpus = _fixture.TestHubCorpus;
        IReadOnlyList<string> ranked = HubCorpus.RankedCleanTerms(corpus);
        Assert.True(ranked.Count > 0, "no corpus terms");
        string term = ranked[0];
        System.Text.RegularExpressions.Regex word = HubCorpus.WordRegex(term);

        // The folder with the most matches becomes the bound.
        var byFolder = corpus
            .Where(i => word.IsMatch(HubCorpus.TextOf(i)))
            .GroupBy(i => i.FolderPath, StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(g => g.Count())
            .ToList();
        Assert.True(byFolder.Count > 0, "no folder contains term matches");
        string folderPath = byFolder[0].Key;
        var expected = new HashSet<string>(byFolder[0].Select(i => i.EntryId), StringComparer.OrdinalIgnoreCase);

        SearchOutcome outcome = Service.Search(new SearchRequest
        {
            Exhaustive = true,
            Store = Hub,
            Folder = folderPath,
            Query = term,
            Top = 100,
        });

        Assert.NotNull(outcome.Exhaustive);
        var actual = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (HitSummary hit in outcome.Hits.Where(h => !HubCorpus.IsTestArtifact(h.Subject)))
        {
            actual.Add(Service.Read(hit.Id, maxBodyChars: 0).EntryId);
        }

        _output.WriteLine($"folder-bounded term '{term}': folderDepth={folderPath.Split('/').Length} "
            + $"expected={expected.Count} actual={actual.Count} scanMs={outcome.Exhaustive!.ElapsedMs} "
            + $"folders={outcome.Exhaustive.FoldersScanned} engine={outcome.Exhaustive.Engine}");
        AssertSetsEqual(expected, actual, term, "folderGroundTruth", "exhaustive");
    }

    private static void AssertSetsEqual(HashSet<string> left, HashSet<string> right, string term, string leftName, string rightName)
    {
        List<string> leftOnly = left.Where(id => !right.Contains(id)).ToList();
        List<string> rightOnly = right.Where(id => !left.Contains(id)).ToList();
        Assert.True(leftOnly.Count == 0 && rightOnly.Count == 0,
            $"term '{term}': {leftName} vs {rightName} sets differ - {leftName}Only={leftOnly.Count} {rightName}Only={rightOnly.Count}");
    }
}
