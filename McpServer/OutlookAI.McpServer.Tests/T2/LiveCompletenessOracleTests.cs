using System.Text.RegularExpressions;
using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// T2 completeness oracle (v3.MD section 0.6 Phase 1, S2/D19): a full COM walk of the
/// tiny designated test-hub store builds ground truth; for at least three probe terms the
/// set of expected matches (direct Subject/Body word inspection) must equal the
/// index-search results, with any difference explained by the reported staleness
/// frontier. Identity: index hits are mapped to REAL COM EntryIDs through
/// <see cref="HitLocator"/> (the 24-byte decoded id is not openable on cached Exchange
/// stores - Phase-1 finding) and compared against the walked items' EntryIDs.
///
/// Term-matching parity: the oracle queries subject+body columns only
/// (SearchIn.SubjectAndBody) and picks terms whose corpus occurrences are all clean
/// word-boundary occurrences, mirroring the index word breaker. Drafts and other items
/// without ReceivedTime are excluded on both sides. Only the test-hub store is ever
/// walked (S2); logging stays within the test-hub grant: generic single-word terms,
/// counts and ids.
///
/// Residual rows: a small minority of index rows may point at items that no longer
/// exist at the indexed location (deletes and cache rebuilds leave rows until the
/// indexer garbage-collects - seen live after the Phase-6 full-caching resync). Those
/// are logged and bounded, not forbidden; recall/precision asserts stay strict.
///
/// Deep-content extras (live-bitten 2026-07-26): System.Search.Contents indexes MORE
/// than the COM plain-text subject+body - HTML-only tokens (link URLs, alt text),
/// attachment text, address fields. A corpus-derived term can therefore match an
/// index row whose walked item's plain text does NOT contain the term (verified live
/// on a real mail: term absent from subject+body, present for the index). Such an
/// extra is precision-POSITIVE for the product; the oracle tolerates it when the
/// located item IS a walked real item whose corpus text provably lacks the term,
/// bounded together with residual rows. Located mail items OUTSIDE the walk remain a
/// hard failure.
/// </summary>
[Collection("LivePhase1")]
[Trait("Category", "Live")]
public sealed class LiveCompletenessOracleTests
{
    private const int MinTerms = 3;
    private const int TimeToleranceSeconds = 5;

    private readonly LivePhase1Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveCompletenessOracleTests(LivePhase1Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    public void CompletenessOracle_IndexMatchesGroundTruth_ForProbeTerms()
    {
        string hubStore = _fixture.Settings.TestHubStoreDisplayName;
        StoreScopeInfo hubScope = _fixture.GetScope(hubStore);
        string hubStoreId = _fixture.GetComStoreId(hubStore);

        // Ground truth: every mail item of the tiny test-hub store (recursive COM walk).
        IReadOnlyList<ComWalkedItem> walk = _fixture.TestHubWalk;
        List<ComWalkedItem> corpus = walk.Where(i => i.ReceivedTime.HasValue).ToList();
        _output.WriteLine($"walked mail items={walk.Count} with ReceivedTime={corpus.Count}");
        Assert.True(corpus.Count > 0, "test-hub store walk found no mail items with ReceivedTime");

        // Staleness frontier for the hub store, measured BEFORE the oracle queries.
        IndexStalenessReport staleness = _fixture.Service.GetStaleness(hubScope.StorePrefix);
        _output.WriteLine($"hub staleness frontierUtc={staleness.NewestIndexedReceivedUtc:O} ageMin={staleness.Age?.TotalMinutes:F1}");

        List<string> terms = SelectProbeTerms(corpus);
        _output.WriteLine("probe terms: " + string.Join(", ", terms));
        Assert.True(terms.Count >= MinTerms, $"only {terms.Count} usable oracle terms derived from the hub corpus");

        foreach (string term in terms)
        {
            VerifyTerm(term, corpus, hubScope, hubStoreId, staleness);
        }
    }

    private void VerifyTerm(
        string term,
        List<ComWalkedItem> corpus,
        StoreScopeInfo hubScope,
        string hubStoreId,
        IndexStalenessReport staleness)
    {
        Regex word = WordRegex(term);

        var expected = new HashSet<string>(
            corpus.Where(i => word.IsMatch(TextOf(i))).Select(i => i.EntryId.ToUpperInvariant()),
            StringComparer.OrdinalIgnoreCase);

        IndexSearchResult result = _fixture.Service.Search(new IndexQuery
        {
            Scope = hubScope.StorePrefix,
            Kinds = KindFilter.EmailOnly,
            Terms = new[] { term },
            SearchIn = SearchIn.SubjectAndBody,
            Top = 500,
        });

        // Map every index hit to its real COM EntryID.
        var located = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var locateFailures = new List<IndexHit>();
        foreach (IndexHit hit in result.Hits.Where(h => h.DateReceivedUtc.HasValue))
        {
            HitLocationResult location = HitLocator.Locate(_fixture.Session, hit, TimeToleranceSeconds);
            if (location.Located != null)
            {
                located.Add(location.Located.EntryId.ToUpperInvariant());
            }
            else
            {
                locateFailures.Add(hit);
            }
        }

        List<string> expectedOnly = expected.Where(id => !located.Contains(id)).ToList();
        List<string> locatedOnly = located.Where(id => !expected.Contains(id)).ToList();
        _output.WriteLine($"term '{term}': expected={expected.Count} indexHits={result.Hits.Count} located={located.Count} "
            + $"expectedOnly={expectedOnly.Count} locatedOnly={locatedOnly.Count} locateFailures={locateFailures.Count} "
            + $"queryMs={result.ElapsedMilliseconds}");

        // Ground-truth matches the index misses must lie beyond the staleness frontier.
        foreach (string id in expectedOnly)
        {
            ComWalkedItem item = corpus.First(i => string.Equals(i.EntryId, id, StringComparison.OrdinalIgnoreCase));
            bool explained = staleness.NewestIndexedReceivedUtc.HasValue
                && ReceivedBeyondFrontier(item.ReceivedTime!.Value, staleness.NewestIndexedReceivedUtc.Value);
            Assert.True(explained,
                $"term '{term}': item {id} expected from ground truth but missing in the index, "
                + "and not explained by the staleness frontier");
            _output.WriteLine($"  expectedOnly {id} explained by staleness (received beyond frontier)");
        }

        // Located index hits absent from the term's ground-truth SET. Two tolerable
        // shapes: (a) the item IS in the walked corpus but its plain subject+body do
        // not contain the term - the index matched deeper content
        // (System.Search.Contents: HTML-only tokens, attachment text, address
        // fields; class doc remarks) - logged and bounded below; (b) a non-mail
        // item (outside oracle scope). A located MAIL item that is not in the walk
        // at all remains a hard precision failure.
        int deepContentExtras = 0;
        foreach (string id in locatedOnly)
        {
            ComWalkedItem? walked = corpus.FirstOrDefault(i => string.Equals(i.EntryId, id, StringComparison.OrdinalIgnoreCase));
            if (walked != null)
            {
                Assert.False(word.IsMatch(TextOf(walked)),
                    $"term '{term}': located item {id} IS a plain-text corpus match but was not in the expected set - oracle bug");
                deepContentExtras++;
                _output.WriteLine($"  locatedOnly {id} is a walked item whose plain subject+body lack the term - "
                    + "index matched deeper content (HTML/attachment/address fields); tolerated + bounded");
                continue;
            }

            ComOpenResult? opened = _fixture.Session.TryOpenItem(id, hubStoreId, out string? error);
            if (opened == null)
            {
                _output.WriteLine($"  locatedOnly {id} no longer opens ({error})");
                Assert.Fail($"term '{term}': located index hit {id} could not be re-opened");
            }

            Assert.True(opened!.ItemClass != 43,
                $"term '{term}': index returned mail item {id} that ground truth does not contain");
            _output.WriteLine($"  locatedOnly {id} is a non-mail item (class={opened.ItemClass}) - outside oracle scope");
        }

        // Location failures = residual rows whose item no longer exists at the indexed
        // location. Index rows CAN outlive their item (v3.MD Phase-2 fact 9 for deleted
        // artifacts; discovered again in Phase 7 after the tuning service's slider=All
        // full-caching resync rebuilt this store's OST and left orphan rows pending
        // indexer garbage collection). The product surfaces these as re-run-search
        // errors, so the oracle tolerates a SMALL residue - correctness is carried by
        // the strict asserts above: recall (expectedOnly=0 unless staleness-explained)
        // and precision (no located mail row outside ground truth). A residue burst
        // beyond 10% of rows still fails: that would indicate a locator regression or
        // an index integrity problem, not leftovers.
        foreach (IndexHit residue in locateFailures)
        {
            bool tagged = residue.Subject != null
                && residue.Subject.IndexOf("[OutlookAI-McpTest]", StringComparison.OrdinalIgnoreCase) >= 0;
            _output.WriteLine($"  residual row: received={residue.DateReceivedUtc:O} taggedArtifact={tagged} "
                + $"error={HitLocator.Locate(_fixture.Session, residue, TimeToleranceSeconds).Error}");
        }

        // Tolerance: 10% of the term's hits, with a floor of ONE - a single
        // residual/deep-content row on a low-frequency term (e.g. 1 of 3 hits) is
        // normal store history, not an integrity signal; a real locator or index
        // regression shows up as MANY tolerated rows and still fails here.
        int tolerated = locateFailures.Count + deepContentExtras;
        Assert.True(tolerated <= Math.Max(1, result.Hits.Count / 10),
            $"term '{term}': {locateFailures.Count} residual + {deepContentExtras} deep-content rows of {result.Hits.Count} "
            + "index hits - beyond the tolerated minority (locator regression or index integrity problem?)");

        Assert.True(expected.Count > 0, $"term '{term}' unexpectedly has no ground-truth matches");
    }

    /// <summary>
    /// Chooses at least three ASCII probe terms from the corpus. A term qualifies when
    /// every corpus occurrence is a clean word occurrence (no letter/digit neighbors), so
    /// regex ground truth and the index word breaker agree. Deterministic: candidates
    /// ordered by match count, then picked from the extremes and middle for diversity.
    /// </summary>
    private static List<string> SelectProbeTerms(List<ComWalkedItem> corpus)
    {
        var matchCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var tokenRegex = new Regex("[A-Za-z]{4,}", RegexOptions.CultureInvariant);

        var texts = corpus.Select(TextOf).ToList();
        var tokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (string text in texts)
        {
            foreach (Match m in tokenRegex.Matches(text))
            {
                tokens.Add(m.Value.ToLowerInvariant());
            }
        }

        foreach (string token in tokens)
        {
            if (!AllOccurrencesAreCleanWords(texts, token))
            {
                continue;
            }

            Regex word = WordRegex(token);
            int count = texts.Count(t => word.IsMatch(t));
            if (count >= 1)
            {
                matchCounts[token] = count;
            }
        }

        List<string> ordered = matchCounts
            .OrderByDescending(kv => kv.Value)
            .ThenBy(kv => kv.Key, StringComparer.Ordinal)
            .Select(kv => kv.Key)
            .ToList();

        var picked = new List<string>();
        if (ordered.Count > 0)
        {
            picked.Add(ordered[0]);
        }

        if (ordered.Count > 2)
        {
            picked.Add(ordered[ordered.Count / 2]);
        }

        if (ordered.Count > 1)
        {
            picked.Add(ordered[ordered.Count - 1]);
        }

        foreach (string candidate in ordered)
        {
            if (picked.Count >= MinTerms)
            {
                break;
            }

            if (!picked.Contains(candidate, StringComparer.OrdinalIgnoreCase))
            {
                picked.Add(candidate);
            }
        }

        return picked;
    }

    private static bool AllOccurrencesAreCleanWords(List<string> texts, string token)
    {
        foreach (string text in texts)
        {
            int start = 0;
            while (true)
            {
                int idx = text.IndexOf(token, start, StringComparison.OrdinalIgnoreCase);
                if (idx < 0)
                {
                    break;
                }

                bool beforeOk = idx == 0 || !char.IsLetterOrDigit(text[idx - 1]);
                int end = idx + token.Length;
                bool afterOk = end >= text.Length || !char.IsLetterOrDigit(text[end]);
                if (!beforeOk || !afterOk)
                {
                    return false;
                }

                start = idx + 1;
            }
        }

        return true;
    }

    private static Regex WordRegex(string term)
    {
        return new Regex(
            "(?<![A-Za-z0-9])" + Regex.Escape(term) + "(?![A-Za-z0-9])",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    }

    private static string TextOf(ComWalkedItem item)
    {
        return (item.Subject ?? string.Empty) + "\n" + (item.Body ?? string.Empty);
    }

    /// <summary>
    /// True when the item's ReceivedTime lies beyond the staleness frontier under either
    /// clock interpretation (COM reports local wall time; the frontier is UTC).
    /// </summary>
    private static bool ReceivedBeyondFrontier(DateTime receivedLocal, DateTime frontierUtc)
    {
        DateTime frontierWithSlack = frontierUtc.AddSeconds(-60);
        DateTime asUtc = DateTime.SpecifyKind(receivedLocal, DateTimeKind.Local).ToUniversalTime();
        DateTime asRaw = DateTime.SpecifyKind(receivedLocal, DateTimeKind.Utc);
        return asUtc > frontierWithSlack || asRaw > frontierWithSlack;
    }
}
