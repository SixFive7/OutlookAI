using System.Diagnostics;
using System.Text.RegularExpressions;
using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// T2 live regression for the SF-6 recall bug and its D40 fix (user order 2026-07-26:
/// search subject+body by default, optionally narrow to one of them).
/// <para>
/// The headline test is built from the discovery case itself: a real, stable population
/// whose term sits in the SUBJECT only (coordinates in the gitignored live settings,
/// S6/D13). Before the fix, <c>query=</c> could not match a single one of them because
/// the unqualified <c>CONTAINS('term')</c> predicate searches
/// <c>System.Search.Contents</c> alone; the expected count is derived LIVE in the test
/// from a <c>from:</c>-scoped query over the same folder, so a changed corpus moves the
/// expectation instead of breaking the test.
/// </para>
/// <para>
/// The scope semantics (subject-only / body-only) are then pinned across all three
/// tiers - index, freshness sweep and exhaustive COM scan - on the tiny hub store (S2),
/// using terms derived live from its COM ground truth. Everything here is READ-ONLY: no
/// mailbox writes, no test artifacts, nothing to clean up. Logging is counts and
/// timings only (S4).
/// </para>
/// </summary>
[Collection(LiveCollections.Phase3)]
[Trait("Category", "Live")]
public sealed class LiveSearchInTests
{
    private const int MaxQueryMs = 2000;

    private readonly LivePhase3Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveSearchInTests(LivePhase3Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private SubjectOnlyProbeSettings Probe => _fixture.Settings.SubjectOnlyProbe!;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    // ------------------------------------------------ the SF-6 discovery case (index tier)

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "ProbePopulation")]
    public void Sf6DiscoveryCase_IndexTier_SubjectOnlyPopulationIsFoundByDefaultScope()
    {
        IndexSearchService index = IndexSearchService.CreateDefault(out _);
        string scope = ResolveProbeFolderScope(index);

        // Independent expectation: the same population selected by sender address only -
        // a per-column CONTAINS that was never affected by the SF-6 bug (Phase-1 fact 4).
        int expected = CountRows(index, scope, terms: null, SearchInValues.Default, Probe.SenderFragment);
        Assert.True(expected > 0, "the from:-derived expectation population is empty - the corpus moved");

        int byDefaultScope = CountRows(index, scope, new[] { Probe.SubjectTerm }, SearchIn.SubjectAndBody, from: null);
        int bySubjectScope = CountRows(index, scope, new[] { Probe.SubjectTerm }, SearchIn.SubjectOnly, from: null);
        int byBodyScope = CountRows(index, scope, new[] { Probe.SubjectTerm }, SearchIn.BodyOnly, from: null);

        _output.WriteLine($"expected(from:)={expected} default={byDefaultScope} subject={bySubjectScope} body={byBodyScope}");

        // THE regression: before D40 this was 0 while the mail sat right there.
        Assert.True(byDefaultScope > 0, "default search_in scope found none of the subject-only population (SF-6 regression)");
        Assert.Equal(expected, byDefaultScope);
        Assert.Equal(byDefaultScope, bySubjectScope);

        // ... and the reason it used to fail: the body stream carries no subject text.
        Assert.Equal(0, byBodyScope);
    }

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "ProbePopulation")]
    public void Sf6DiscoveryCase_IndexTier_PrefixStemsWorkInTheSubjectColumnToo()
    {
        Assert.True(Probe.SubjectTerm.Length >= 5, "probe term too short to stem");

        IndexSearchService index = IndexSearchService.CreateDefault(out _);
        string scope = ResolveProbeFolderScope(index);
        string stem = Probe.SubjectTerm.Substring(0, Probe.SubjectTerm.Length - 2) + "*";

        int whole = CountRows(index, scope, new[] { Probe.SubjectTerm }, SearchIn.SubjectAndBody, from: null);
        int prefixed = CountRows(index, scope, new[] { stem }, SearchIn.SubjectAndBody, from: null);
        int prefixedSubject = CountRows(index, scope, new[] { stem }, SearchIn.SubjectOnly, from: null);

        _output.WriteLine($"whole={whole} prefix(default)={prefixed} prefix(subject)={prefixedSubject}");
        Assert.True(prefixed >= whole, "prefix matching lost hits the whole term found");
        Assert.Equal(prefixed, prefixedSubject);
    }

    // ------------------------------------------------ the SF-6 discovery case (tool tier)

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "ProbePopulation")]
    public void Sf6DiscoveryCase_ToolTier_DefaultQueryReturnsHits_BodyScopeReturnsNone()
    {
        SearchOutcome byDefault = Service.Search(NewProbeRequest(SearchInValues.Default));
        SearchOutcome bySubject = Service.Search(NewProbeRequest(SearchIn.SubjectOnly));
        SearchOutcome byBody = Service.Search(NewProbeRequest(SearchIn.BodyOnly));

        _output.WriteLine(
            $"tool tier: default={byDefault.Hits.Count}(truncated={byDefault.Truncated}, {byDefault.IndexElapsedMs} ms) "
            + $"subject={bySubject.Hits.Count} body={byBody.Hits.Count}");

        Assert.True(byDefault.Hits.Count > 0, "the search tool found none of the subject-only population (SF-6 regression)");
        Assert.Equal(byDefault.Hits.Count, bySubject.Hits.Count);
        Assert.Empty(byBody.Hits);
    }

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "ProbePopulation")]
    public void Sf6DiscoveryCase_ExhaustiveTier_HonorsSearchIn()
    {
        SearchRequest subjectScoped = NewProbeRequest(SearchIn.SubjectOnly);
        subjectScoped.Exhaustive = true;
        SearchRequest bodyScoped = NewProbeRequest(SearchIn.BodyOnly);
        bodyScoped.Exhaustive = true;

        Stopwatch clock = Stopwatch.StartNew();
        SearchOutcome bySubject = Service.Search(subjectScoped);
        clock.Stop();
        SearchOutcome byBody = Service.Search(bodyScoped);

        _output.WriteLine(
            $"exhaustive tier: subject={bySubject.Hits.Count} ({clock.ElapsedMilliseconds} ms, "
            + $"engine={bySubject.Exhaustive?.Engine}) body={byBody.Hits.Count}");

        Assert.True(bySubject.Hits.Count > 0, "exhaustive subject scope found none of the subject-only population");
        Assert.Empty(byBody.Hits);
    }

    // ------------------------------------------------ latency delta of the OR-pair (measured)

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "ProbePopulation")]
    public void IndexTier_OrPairLatency_StaysAcceptableVersusSingleColumn()
    {
        IndexSearchService index = IndexSearchService.CreateDefault(out _);
        string probeScope = ResolveProbeFolderScope(index);
        StoreScopeInfo probeStore = ResolveStoreScope(index, Probe.StoreDisplayName);

        // Agent-sized shape: what MailService actually emits for a default search
        // (top 25 over-fetched by one, ORDER BY DateReceived DESC).
        (long subjectMs, long bodyMs, long pairMs) folder = MeasureShapes(index, probeScope, Probe.SubjectTerm);
        (long subjectMs, long bodyMs, long pairMs) store = MeasureShapes(index, probeStore.StorePrefix, _fixture.Settings.ProbeTerm);
        (long subjectMs, long bodyMs, long pairMs) allStores = MeasureShapes(index, null, _fixture.Settings.ProbeTerm);

        _output.WriteLine($"folder-scoped  subject={folder.subjectMs} body={folder.bodyMs} orPair={folder.pairMs} ms");
        _output.WriteLine($"store-scoped   subject={store.subjectMs} body={store.bodyMs} orPair={store.pairMs} ms");
        _output.WriteLine($"all stores     subject={allStores.subjectMs} body={allStores.bodyMs} orPair={allStores.pairMs} ms");

        foreach ((long subjectMs, long bodyMs, long pairMs) measured in new[] { folder, store, allStores })
        {
            Assert.InRange(measured.pairMs, 0, MaxQueryMs);

            // The OR-pair queries two columns instead of one; it must stay in the same
            // league as a single-column CONTAINS, not degrade into a property scan.
            long singleColumnBudget = Math.Max(measured.subjectMs, measured.bodyMs);
            Assert.True(
                measured.pairMs <= (singleColumnBudget * 3) + 250,
                $"OR-pair {measured.pairMs} ms is far above the single-column cost {singleColumnBudget} ms");
        }
    }

    // ------------------------------------------------ scope semantics across all three tiers

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "ProbePopulation")]
    public void AllTiers_SubjectOnlyAndBodyOnlyTerms_AreSeparatedConsistently()
    {
        IReadOnlyList<ComWalkedItem> corpus = _fixture.TestHubCorpus;
        Assert.True(corpus.Count > 0, "hub corpus is empty");

        IndexSearchService index = IndexSearchService.CreateDefault(out _);
        StoreScopeInfo hub = ResolveStoreScope(index, Hub);

        // Candidates come from the COM ground truth (a word in one field, absent as a
        // substring from the other across the whole store). The index then confirms the
        // separation before a term is used: the catalog outlives deleted items
        // (IncludeDeletedItems=1, Phase-2 fact 9), so a word the live walk sees in no
        // subject can still sit in a subject row of a long-gone item. Confirming here
        // keeps the cross-tier assertions below honest instead of flaky.
        string subjectOnlyTerm = SelectIndexConfirmedTerm(index, hub, corpus, subjectSide: true);
        string bodyOnlyTerm = SelectIndexConfirmedTerm(index, hub, corpus, subjectSide: false);

        _output.WriteLine($"derived hub terms: subjectOnly.len={subjectOnlyTerm.Length} bodyOnly.len={bodyOnlyTerm.Length}");

        // Default scope is the union - both terms must be found by a plain query.
        Assert.True(CountRows(index, hub.StorePrefix, new[] { subjectOnlyTerm }, SearchIn.SubjectAndBody, null, KindFilter.MailKindOnly) > 0);
        Assert.True(CountRows(index, hub.StorePrefix, new[] { bodyOnlyTerm }, SearchIn.SubjectAndBody, null, KindFilter.MailKindOnly) > 0);

        // --- tool tier (index + freshness sweep merged, D34)
        AssertToolTierSeparation(subjectOnlyTerm, expectedInSubjectScope: true);
        AssertToolTierSeparation(bodyOnlyTerm, expectedInSubjectScope: false);

        // --- exhaustive tier (index bypassed; bounded whole-hub date scan)
        AssertExhaustiveSeparation(subjectOnlyTerm, expectedInSubjectScope: true);
        AssertExhaustiveSeparation(bodyOnlyTerm, expectedInSubjectScope: false);
    }

    private void AssertToolTierSeparation(string term, bool expectedInSubjectScope)
    {
        int byDefault = Service.Search(NewHubRequest(term, SearchInValues.Default)).Hits.Count;
        int bySubject = Service.Search(NewHubRequest(term, SearchIn.SubjectOnly)).Hits.Count;
        int byBody = Service.Search(NewHubRequest(term, SearchIn.BodyOnly)).Hits.Count;

        _output.WriteLine($"tool tier: default={byDefault} subject={bySubject} body={byBody} (subjectSide={expectedInSubjectScope})");

        Assert.True(byDefault > 0, "the default scope must find a term that exists in the hub corpus");
        if (expectedInSubjectScope)
        {
            Assert.True(bySubject > 0);
            Assert.Equal(0, byBody);
        }
        else
        {
            Assert.True(byBody > 0);
            Assert.Equal(0, bySubject);
        }
    }

    private void AssertExhaustiveSeparation(string term, bool expectedInSubjectScope)
    {
        SearchRequest MakeRequest(SearchIn scope)
        {
            SearchRequest request = NewHubRequest(term, scope);
            request.Exhaustive = true;
            request.AfterUtc = new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            return request;
        }

        SearchOutcome byDefault = Service.Search(MakeRequest(SearchInValues.Default));
        int bySubject = Service.Search(MakeRequest(SearchIn.SubjectOnly)).Hits.Count;
        int byBody = Service.Search(MakeRequest(SearchIn.BodyOnly)).Hits.Count;

        _output.WriteLine(
            $"exhaustive tier: default={byDefault.Hits.Count} subject={bySubject} body={byBody} "
            + $"engine={byDefault.Exhaustive?.Engine} (subjectSide={expectedInSubjectScope})");

        Assert.True(byDefault.Hits.Count > 0, "the exhaustive default scope must find a term that exists in the hub corpus");
        if (expectedInSubjectScope)
        {
            Assert.True(bySubject > 0);
            Assert.Equal(0, byBody);
        }
        else
        {
            Assert.True(byBody > 0);
            Assert.Equal(0, bySubject);
        }
    }

    // ------------------------------------------------------------------ helpers

    private SearchRequest NewProbeRequest(SearchIn searchIn) => new()
    {
        Query = Probe.SubjectTerm,
        SearchIn = searchIn,
        Store = Probe.StoreDisplayName,
        Folder = Probe.FolderPath,
        Top = MailService.SearchTopCap,
        SnippetChars = 0,
    };

    private SearchRequest NewHubRequest(string term, SearchIn searchIn) => new()
    {
        Query = term,
        SearchIn = searchIn,
        Store = Hub,
        IncludeAttachmentHits = false,
        Top = MailService.SearchTopCap,
        SnippetChars = 0,
    };

    private string ResolveProbeFolderScope(IndexSearchService index)
    {
        // Same shape MailService.ResolveScope builds for store + folder.
        return ResolveStoreScope(index, Probe.StoreDisplayName).StorePrefix + "/0/" + Probe.FolderPath.Trim('/');
    }

    private static StoreScopeInfo ResolveStoreScope(IndexSearchService index, string storeDisplayName)
    {
        StoreScopeInfo? scope = index.DiscoverStoreScopes(2000)
                .FirstOrDefault(s => string.Equals(s.StoreDisplayName, storeDisplayName, StringComparison.OrdinalIgnoreCase))
            ?? index.TryDiscoverStoreScopeByAddress(storeDisplayName);
        Assert.True(scope != null, "store scope not discoverable in the index");
        return scope!;
    }

    private static int CountRows(
        IndexSearchService index,
        string? scope,
        IReadOnlyList<string>? terms,
        SearchIn searchIn,
        string? from,
        KindFilter kinds = KindFilter.MessagesAndAttachments)
    {
        return index.Search(new IndexQuery
        {
            Scope = scope,
            Terms = terms,
            SearchIn = searchIn,
            SenderContains = from,
            Kinds = kinds,
            Top = 5000,
        }).Hits.Count;
    }

    private static (long SubjectMs, long BodyMs, long PairMs) MeasureShapes(
        IndexSearchService index, string? scope, string term)
    {
        long Measure(SearchIn searchIn)
        {
            // Warm the shape once, then take the best of three (the index caches nothing
            // between shapes, so a single cold sample would measure the OS, not the SQL).
            long best = long.MaxValue;
            for (int i = 0; i < 4; i++)
            {
                Stopwatch clock = Stopwatch.StartNew();
                index.Search(new IndexQuery
                {
                    Scope = scope,
                    Terms = new[] { term },
                    SearchIn = searchIn,
                    Top = MailService.SearchTopDefault + 1,
                });
                clock.Stop();
                if (i > 0)
                {
                    best = Math.Min(best, clock.ElapsedMilliseconds);
                }
            }

            return best;
        }

        return (Measure(SearchIn.SubjectOnly), Measure(SearchIn.BodyOnly), Measure(SearchIn.SubjectAndBody));
    }

    /// <summary>
    /// Picks the first ground-truth candidate the index also separates cleanly, and
    /// asserts along the way that the index DOES separate the two columns (the whole
    /// point of D40 - before the fix the body column was all the term predicate saw).
    /// </summary>
    private string SelectIndexConfirmedTerm(
        IndexSearchService index, StoreScopeInfo hub, IReadOnlyList<ComWalkedItem> corpus, bool subjectSide)
    {
        IReadOnlyList<string> candidates = RankFieldExclusiveTerms(corpus, subjectSide);
        Assert.True(candidates.Count > 0,
            subjectSide ? "no hub term occurs in a subject and in no body" : "no hub term occurs in a body and in no subject");

        SearchIn ownScope = subjectSide ? SearchIn.SubjectOnly : SearchIn.BodyOnly;
        SearchIn otherScope = subjectSide ? SearchIn.BodyOnly : SearchIn.SubjectOnly;

        foreach (string candidate in candidates)
        {
            int inOwnField = CountRows(index, hub.StorePrefix, new[] { candidate }, ownScope, null, KindFilter.MailKindOnly);
            int inOtherField = CountRows(index, hub.StorePrefix, new[] { candidate }, otherScope, null, KindFilter.MailKindOnly);
            if (inOwnField > 0 && inOtherField == 0)
            {
                _output.WriteLine($"index-confirmed {(subjectSide ? "subject" : "body")}-only term: own={inOwnField} other={inOtherField}");
                return candidate;
            }
        }

        Assert.Fail(
            $"the index separated none of {candidates.Count} ground-truth "
            + $"{(subjectSide ? "subject" : "body")}-exclusive hub terms into its own column");
        return string.Empty;
    }

    /// <summary>
    /// Terms that occur as a clean word in at least one subject (or body) of the hub
    /// corpus and NOWHERE AT ALL - not even as a substring - in the other field, ranked
    /// by occurrence count. That is the ground-truth definition the index/sweep/scan
    /// predicates must agree with; substring exclusion is deliberately stricter than the
    /// word breaker so the negative assertions cannot be undone by tokenization
    /// differences.
    /// </summary>
    private static IReadOnlyList<string> RankFieldExclusiveTerms(
        IReadOnlyList<ComWalkedItem> corpus, bool subjectSide)
    {
        var tokenRegex = new Regex("[A-Za-z]{4,}", RegexOptions.CultureInvariant);
        List<string> ownField = new();
        List<string> otherField = new();
        foreach (ComWalkedItem item in corpus)
        {
            ownField.Add((subjectSide ? item.Subject : item.Body) ?? string.Empty);
            otherField.Add((subjectSide ? item.Body : item.Subject) ?? string.Empty);
        }

        var candidates = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (string text in ownField)
        {
            foreach (Match match in tokenRegex.Matches(text))
            {
                string token = match.Value.ToLowerInvariant();
                if (candidates.ContainsKey(token))
                {
                    continue;
                }

                if (otherField.Any(o => o.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0))
                {
                    continue; // present in the other field - cannot separate the scopes
                }

                Regex word = HubCorpus.WordRegex(token);
                int matches = ownField.Count(t => word.IsMatch(t));
                if (matches > 0 && AllOccurrencesAreCleanWords(ownField, token))
                {
                    candidates[token] = matches;
                }
            }
        }

        return candidates
            .OrderByDescending(kv => kv.Value)
            .ThenBy(kv => kv.Key, StringComparer.Ordinal)
            .Select(kv => kv.Key)
            .ToList();
    }

    private static bool AllOccurrencesAreCleanWords(IReadOnlyList<string> texts, string token)
    {
        foreach (string text in texts)
        {
            int start = 0;
            while (true)
            {
                int index = text.IndexOf(token, start, StringComparison.OrdinalIgnoreCase);
                if (index < 0)
                {
                    break;
                }

                bool beforeOk = index == 0 || !char.IsLetterOrDigit(text[index - 1]);
                int end = index + token.Length;
                bool afterOk = end >= text.Length || !char.IsLetterOrDigit(text[end]);
                if (!beforeOk || !afterOk)
                {
                    return false;
                }

                start = index + 1;
            }
        }

        return true;
    }
}
