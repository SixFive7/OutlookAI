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
[Collection(LiveCollections.Phase1)]
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

    /// <summary>
    /// The stores the index tier measures - the settings' INDEXED list, never the watched one, and
    /// refused rather than empty (see <see cref="LiveTestSettings.RequireIndexedStores"/>).
    /// </summary>
    private List<string> Indexed => _fixture.Settings.RequireIndexedStores().ToList();

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "MultipleStores")]
    public void ProviderSelection_OleDbPrimaryPath_IsRecorded()
    {
        _output.WriteLine(_fixture.ProviderReport);

        // The fallback would still be a pass functionally, but the chosen path must be
        // recorded either way (v3.MD Phase 1 row). OleDb is the expected primary.
        Assert.False(string.IsNullOrWhiteSpace(_fixture.ProviderReport));
        Assert.Equal(IndexProviderKind.OleDb, _fixture.Service.Provider);
    }

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "MultipleStores")]
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
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "MultipleStores")]
    public void ProbeParity_AllThreeStores_ReturnRowsUnder2s()
    {
        foreach (string storeName in Indexed)
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
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "MultipleStores")]
    public void ProbeParity_McpShapedQuery_ScopeKindContainsOrderBy()
    {
        // Section-5 R3 shape: store scope + kind + CONTAINS + ORDER BY DESC, TOP 25.
        int totalHits = 0;
        foreach (string storeName in Indexed)
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
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "MultipleStores")]
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
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "MultipleStores")]
    public void FilterShapes_ReadAndAttachmentFlags_WorkUnder2s()
    {
        // The CONTENT half reads the first indexed store - the hub on a guest, whose population
        // carries attachments and unread mail. The LATENCY half is timed on the LARGEST indexed
        // store (decided 2026-09-24): on a few-dozen-item hub a two-second bound is met by
        // construction, and T1/LatencyTargetTests fails if this ever goes back to timing it there.
        StoreScopeInfo scope = _fixture.GetScope(Indexed[0]);

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

        StoreScopeInfo timed = _fixture.LargestIndexedScope;
        IndexSearchResult unreadTimed = _fixture.Service.Search(new IndexQuery
        {
            Scope = timed.StorePrefix,
            Kinds = KindFilter.MailKindOnly,
            IsRead = false,
            Top = 5,
        });
        IndexSearchResult withAttachmentsTimed = _fixture.Service.Search(new IndexQuery
        {
            Scope = timed.StorePrefix,
            Kinds = KindFilter.MailKindOnly,
            HasAttachments = true,
            Top = 5,
        });

        AssertTimedOnTheLargestStore(unreadTimed, timed, "unread filter");
        AssertTimedOnTheLargestStore(withAttachmentsTimed, timed, "has-attachments filter");

        // The only assertion in this test that says anything about the FILTER rather than about how
        // long it took - and an empty hit list satisfies Assert.All without examining a single row,
        // so on a store with no attachment-bearing indexed mail the has-attachments shape was never
        // checked at all and the test still reported green. Found by the 2026-08-23 VM coverage
        // analysis (section 8 item 4) and fixed with the same idiom as the other three.
        Assert.All(
            LivePopulationCoverage.Require(
                _fixture.Settings,
                withAttachments.Hits,
                "an indexed mail item carrying an attachment in the first indexed store",
                "the has-attachments index filter shape check",
                "To exercise it, point the first entry of 'indexedStoreDisplayNames' (or, with that list absent, "
                    + "of 'expectedStoreDisplayNames') at a store "
                    + "whose indexed mail includes at least one message with an attachment.",
                _output.WriteLine),
            h => Assert.NotEqual(false, h.HasAttachments));
    }

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "MultipleStores")]
    public void SenderFilter_PerColumnContains_IndexBackedUnder2s()
    {
        // Any sender address seen in recent mail of the first store; asserted content-free.
        StoreScopeInfo scope = _fixture.GetScope(Indexed[0]);
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
        string? filteredOn = null;
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
                filteredOn = candidate;
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

        // The LATENCY half: the same per-column CONTAINS, timed on the LARGEST indexed store
        // (decided 2026-09-24; T1/LatencyTargetTests holds it there). It may match nothing there -
        // the measurement corpus has no senders - and a shape that matches nothing over a large
        // store is exactly the one that pays for a scan if the predicate is not index-backed.
        StoreScopeInfo largestScope = _fixture.LargestIndexedScope;
        IndexSearchResult timed = _fixture.Service.Search(new IndexQuery
        {
            Scope = largestScope.StorePrefix,
            Kinds = KindFilter.MailKindOnly,
            SenderContains = filteredOn!,
            Top = 5,
        });
        AssertTimedOnTheLargestStore(timed, largestScope, "sender filter");
    }

    /// <summary>
    /// The ONE way the filter-shape tests assert a latency bound, and it refuses to be handed anything
    /// but the largest indexed store's scope - decided 2026-09-24 (<see cref="LiveLatencyTarget"/>).
    /// T1/LatencyTargetTests pins, from the IL, that those tests time through here and never call
    /// <c>Assert.InRange</c> themselves, so timing one on the first store again cannot happen quietly.
    /// </summary>
    private void AssertTimedOnTheLargestStore(IndexSearchResult result, StoreScopeInfo timedOn, string what)
    {
        (IReadOnlyList<LiveStoreSize> sizes, string largest) = LiveLatencyTarget.Measure(_fixture.Settings);
        _output.WriteLine(LiveLatencyTarget.Describe(sizes, largest));
        Assert.True(
            string.Equals(timedOn.StorePrefix, _fixture.GetScope(largest).StorePrefix, StringComparison.OrdinalIgnoreCase),
            $"the {what} latency was timed on '{timedOn.StoreDisplayName}', not on the largest indexed store '{largest}'");
        _output.WriteLine($"{what}: timed on '{largest}' rows={result.Hits.Count} ms={result.ElapsedMilliseconds}");
        Assert.InRange(result.ElapsedMilliseconds, 0, MaxQueryMs);
    }

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "MultipleStores")]
    [Trait("Requires", "DelegateStore")]
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
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "MultipleStores")]
    public void Staleness_SelfReportsPlausibleFrontier()
    {
        IndexStalenessReport report = _fixture.Service.GetStaleness();

        Assert.NotNull(report.NewestIndexedReceivedUtc);
        Assert.NotNull(report.Age);
        _output.WriteLine($"newestIndexedUtc={report.NewestIndexedReceivedUtc:O} ageMinutes={report.Age!.Value.TotalMinutes:F1}");

        Assert.True(report.NewestIndexedReceivedUtc!.Value.Year >= 2020, "frontier implausibly old");
        // Allow small clock skew but the frontier must not sit in the future.
        Assert.True(report.NewestIndexedReceivedUtc.Value <= report.ClockUtc.AddMinutes(5), "frontier lies in the future");

        // The "not in the future" half catches a product that reads the index's local time as UTC ONLY
        // while the real frontier is younger than this machine's UTC offset. On a test guest the hub's
        // generated population is what makes it that young - rebuilt against the moment the run starts,
        // by Testbed/guest/Reset-HubPopulation.ps1 (decided 2026-09-24) - so a guest that declares one
        // must have run that step: a stale hub FAILS here with the remedy rather than passing having
        // measured nothing. And a frontier OLDER than the population's own newest item is the other
        // direction of the same misreading - or an index that has not taken the rebuilt hub in yet.
        string? manifest = _fixture.Settings.HubPopulationManifestPath;
        TimeSpan offset = TimeZoneInfo.Local.GetUtcOffset(report.ClockUtc);
        if (manifest != null)
        {
            Assert.True(File.Exists(manifest), $"the hub population manifest the settings name is not there: {manifest}");
            HubPopulationFact hub = LiveHubPopulationFreshness.Read(File.ReadLines(manifest));
            (bool fresh, string why) = LiveHubPopulationFreshness.Decide(hub, report.ClockUtc, offset);
            _output.WriteLine(why);
            Assert.True(fresh, why);
            Assert.True(
                report.NewestIndexedReceivedUtc.Value >= hub.NewestDatedUtc - TimeSpan.FromSeconds(2),
                $"the index frontier {report.NewestIndexedReceivedUtc.Value:O} is OLDER than the hub population's own newest "
                + $"item {hub.NewestDatedUtc:O}: either the product reads the index's time shifted back by the UTC offset, or "
                + "the index has not taken the rebuilt hub in yet - Reset-HubPopulation.ps1 waits for it with corpus-indexed");
            return;
        }

        // No generated hub declared - the maintainer's machine, whose hub is real mail. Nothing to
        // rebuild, but the same limit applies and is said out loud rather than assumed away.
        TimeSpan margin = LiveHubPopulationFreshness.DiscriminatingAge(offset);
        LivePopulationCoverage.Require(
            _fixture.Settings,
            report.Age!.Value <= margin ? new[] { report.NewestIndexedReceivedUtc.Value } : Array.Empty<DateTime>(),
            $"an index frontier younger than this machine's UTC offset less {LiveHubPopulationFreshness.FrontierFutureTolerance.TotalMinutes:F0} min ({margin.TotalMinutes:F0} min)",
            "the local-time half of the frontier check",
            "On a test guest declare the hub population (hubPopulationManifestPath) and rebuild it before the run with "
                + "Testbed/guest/Reset-HubPopulation.ps1; elsewhere, run where mail has arrived within the last hour.",
            _output.WriteLine);
    }

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "MultipleStores")]
    public void StoreDiscovery_FindsAllExpectedStores()
    {
        _output.WriteLine("discovered scopes: "
            + string.Join(", ", _fixture.StoreScopes.Select(s => $"{s.StoreDisplayName}({s.SampleCount})")));

        foreach (string expected in Indexed)
        {
            Assert.NotNull(_fixture.GetScope(expected));
        }
    }
}
