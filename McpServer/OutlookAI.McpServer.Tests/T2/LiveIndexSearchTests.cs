using OutlookAI.Core.IndexSearch;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// T2 live tier (v3.MD section 0.6 Phase 1): reproduces the section-5 probe queries
/// through the IndexSearch module against the real SystemIndex. Read-only. Logging is
/// content-free for business stores (S4): counts, ids, timings, booleans - never
/// subjects/bodies.
/// </summary>
[Collection("LivePhase1")]
[Trait("Category", "Live")]
public sealed class LiveIndexSearchTests
{
    private const int MaxQueryMs = 2000;

    private readonly LivePhase1Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveIndexSearchTests(LivePhase1Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    public void ProviderSelection_OleDbPrimaryPath_IsRecorded()
    {
        _output.WriteLine(_fixture.ProviderReport);

        // The fallback would still be a pass functionally, but the chosen path must be
        // recorded either way (v3.MD Phase 1 row). OleDb is the expected primary.
        Assert.False(string.IsNullOrWhiteSpace(_fixture.ProviderReport));
        Assert.Equal(IndexProviderKind.OleDb, _fixture.Service.Provider);
    }

    [Fact]
    public void ProbeParity_Top5Email_HitsUnder2s()
    {
        IndexSearchResult result = _fixture.Service.Search(new IndexQuery
        {
            Kinds = KindFilter.MailKindOnly,
            Top = 5,
        });

        _output.WriteLine($"rows={result.Hits.Count} ms={result.ElapsedMilliseconds}");
        Assert.Equal(5, result.Hits.Count);
        Assert.InRange(result.ElapsedMilliseconds, 0, MaxQueryMs);
    }

    [Fact]
    public void ProbeParity_AllThreeStores_ReturnRowsUnder2s()
    {
        foreach (string storeName in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            StoreScopeInfo scope = _fixture.GetScope(storeName);
            IndexSearchResult result = _fixture.Service.Search(new IndexQuery
            {
                Scope = scope.StorePrefix,
                Kinds = KindFilter.MailKindOnly,
                Top = 5,
            });

            _output.WriteLine($"store={storeName} rows={result.Hits.Count} ms={result.ElapsedMilliseconds}");
            Assert.True(result.Hits.Count > 0, $"store {storeName}: no rows");
            Assert.InRange(result.ElapsedMilliseconds, 0, MaxQueryMs);
        }
    }

    [Fact]
    public void ProbeParity_McpShapedQuery_ScopeKindContainsOrderBy()
    {
        // Section-5 R3 shape: store scope + kind + CONTAINS + ORDER BY DESC, TOP 25.
        int totalHits = 0;
        foreach (string storeName in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            StoreScopeInfo scope = _fixture.GetScope(storeName);
            IndexSearchResult result = _fixture.Service.Search(new IndexQuery
            {
                Scope = scope.StorePrefix,
                Terms = new[] { _fixture.Settings.ProbeTerm },
                Top = 25,
            });

            _output.WriteLine($"store={storeName} rows={result.Hits.Count} ms={result.ElapsedMilliseconds}");
            Assert.InRange(result.ElapsedMilliseconds, 0, MaxQueryMs);
            totalHits += result.Hits.Count;

            // Newest-first ordering among email hits with dates.
            List<DateTime> dates = result.Hits
                .Where(h => h.DateReceivedUtc.HasValue)
                .Select(h => h.DateReceivedUtc!.Value)
                .ToList();
            for (int i = 1; i < dates.Count; i++)
            {
                Assert.True(dates[i - 1] >= dates[i], "hits not ordered newest-first");
            }
        }

        Assert.True(totalHits > 0, "probe term produced no hits in any store");
    }

    [Fact]
    public void ProbeParity_DateRangeQuery_HitsUnder2s()
    {
        IndexSearchResult result = _fixture.Service.Search(new IndexQuery
        {
            Kinds = KindFilter.MailKindOnly,
            ReceivedOnOrAfterUtc = DateTime.UtcNow.AddDays(-30),
            Top = 10,
        });

        _output.WriteLine($"rows={result.Hits.Count} ms={result.ElapsedMilliseconds}");
        Assert.True(result.Hits.Count > 0, "no mail indexed in the last 30 days");
        Assert.InRange(result.ElapsedMilliseconds, 0, MaxQueryMs);
    }

    [Fact]
    public void FilterShapes_ReadAndAttachmentFlags_WorkUnder2s()
    {
        StoreScopeInfo scope = _fixture.GetScope(_fixture.Settings.ExpectedStoreDisplayNames[0]);

        IndexSearchResult unread = _fixture.Service.Search(new IndexQuery
        {
            Scope = scope.StorePrefix,
            Kinds = KindFilter.MailKindOnly,
            IsRead = false,
            Top = 5,
        });
        IndexSearchResult withAttachments = _fixture.Service.Search(new IndexQuery
        {
            Scope = scope.StorePrefix,
            Kinds = KindFilter.MailKindOnly,
            HasAttachments = true,
            Top = 5,
        });

        _output.WriteLine($"unread rows={unread.Hits.Count} ms={unread.ElapsedMilliseconds}; "
            + $"withAttachments rows={withAttachments.Hits.Count} ms={withAttachments.ElapsedMilliseconds}");
        Assert.InRange(unread.ElapsedMilliseconds, 0, MaxQueryMs);
        Assert.InRange(withAttachments.ElapsedMilliseconds, 0, MaxQueryMs);
        Assert.All(withAttachments.Hits, h => Assert.NotEqual(false, h.HasAttachments));
    }

    [Fact]
    public void SenderFilter_PerColumnContains_IndexBackedUnder2s()
    {
        // Any sender address seen in recent mail of the first store; asserted content-free.
        StoreScopeInfo scope = _fixture.GetScope(_fixture.Settings.ExpectedStoreDisplayNames[0]);
        IndexSearchResult recent = _fixture.Service.Search(new IndexQuery
        {
            Scope = scope.StorePrefix,
            Kinds = KindFilter.MailKindOnly,
            Top = 25,
        });
        List<string> candidates = recent.Hits
            .Select(h => h.FromAddress)
            .Where(a => !string.IsNullOrWhiteSpace(a))
            .Select(a => a!)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        Assert.True(candidates.Count > 0, "no sender addresses in recent hits");

        IndexSearchResult? filtered = null;
        foreach (string candidate in candidates)
        {
            try
            {
                filtered = _fixture.Service.Search(new IndexQuery
                {
                    Scope = scope.StorePrefix,
                    Kinds = KindFilter.MailKindOnly,
                    SenderContains = candidate,
                    Top = 5,
                });
                break;
            }
            catch (ArgumentException)
            {
                // Address contains characters outside the term allowlist - try the next.
            }
        }

        Assert.NotNull(filtered);

        _output.WriteLine($"senderFiltered rows={filtered!.Hits.Count} ms={filtered.ElapsedMilliseconds}");
        Assert.True(filtered.Hits.Count > 0, "sender-filtered query returned no rows");
        Assert.InRange(filtered.ElapsedMilliseconds, 0, MaxQueryMs);
    }

    [Fact]
    public void DelegateStoreSubtree_ReturnsRowsUnder2s()
    {
        List<StoreScopeInfo> withDelegates = _fixture.StoreScopes.Where(s => s.HasDelegateSubtree).ToList();
        _output.WriteLine("delegate subtrees under: "
            + string.Join(", ", withDelegates.Select(s => s.StoreDisplayName)));
        Assert.True(withDelegates.Count >= 1, "no delegate-store subtree (store-type /1/) found in the index");

        IndexSearchResult result = _fixture.Service.Search(new IndexQuery
        {
            Scope = withDelegates[0].StorePrefix + "/1",
            Top = 5,
        });

        _output.WriteLine($"delegate rows={result.Hits.Count} ms={result.ElapsedMilliseconds}");
        Assert.True(result.Hits.Count > 0, "delegate-scoped query returned no rows");
        Assert.InRange(result.ElapsedMilliseconds, 0, MaxQueryMs);
    }

    [Fact]
    public void Staleness_SelfReportsPlausibleFrontier()
    {
        IndexStalenessReport report = _fixture.Service.GetStaleness();

        Assert.NotNull(report.NewestIndexedReceivedUtc);
        Assert.NotNull(report.Age);
        _output.WriteLine($"newestIndexedUtc={report.NewestIndexedReceivedUtc:O} ageMinutes={report.Age!.Value.TotalMinutes:F1}");

        Assert.True(report.NewestIndexedReceivedUtc!.Value.Year >= 2020, "frontier implausibly old");
        // Allow small clock skew but the frontier must not sit in the future.
        Assert.True(report.NewestIndexedReceivedUtc.Value <= report.ClockUtc.AddMinutes(5), "frontier lies in the future");
    }

    [Fact]
    public void StoreDiscovery_FindsAllExpectedStores()
    {
        _output.WriteLine("discovered scopes: "
            + string.Join(", ", _fixture.StoreScopes.Select(s => $"{s.StoreDisplayName}({s.SampleCount})")));

        foreach (string expected in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            Assert.NotNull(_fixture.GetScope(expected));
        }
    }
}
