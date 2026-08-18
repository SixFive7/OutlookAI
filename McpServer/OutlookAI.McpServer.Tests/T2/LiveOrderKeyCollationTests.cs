using System;
using System.Collections.Generic;

using OutlookAI.Core.IndexSearch;

using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// T2 live tier, READ-ONLY: measures the one thing <c>IndexOrderGuard</c> is built to be
/// safe without, and proves its recovery query works against the real provider.
/// <para>
/// Since gap B3 the search statement carries no Kind predicate, so rows with no
/// <c>System.Message.DateReceived</c> (appointments, contacts, unsent items) are candidates
/// for the same <c>ORDER BY System.Message.DateReceived DESC</c> cut as mail. Where the
/// Windows Search provider sorts a NULL under DESC therefore decides whether they can fill
/// the TOP and push real mail out of the answer entirely. Nothing in this repo has ever
/// measured that, and the guard is written so it does not need to be known - but knowing it
/// tells the maintainer whether the conditional refetch fires on every truncated search or
/// on none of them, which is the difference between one index query per search and two.
/// </para>
/// <para>
/// Nothing here writes: index statements only, no COM, no mailbox item is opened or touched.
/// Logging is content-free (counts, positions, timings) per the S4 rule.
/// </para>
/// </summary>
[Collection("LivePhase1")]
[Trait("Category", "Live")]
public sealed class LiveOrderKeyCollationTests
{
    private const int ProbeTop = 500;

    private readonly LivePhase1Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveOrderKeyCollationTests(LivePhase1Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    /// <summary>
    /// THE MEASUREMENT. Runs the shipped widened statement over each store and records where
    /// the undated rows landed. Asserts nothing about the answer, because either answer is
    /// legitimate provider behaviour: what it produces is the number that belongs in
    /// Docs/magic-numbers.md beside the guard.
    /// </summary>
    [Fact]
    public void NullCollation_UnderDateReceivedDescending_IsMeasured()
    {
        IIndexClient client = IndexClientFactory.CreateAuto(out string report);
        _output.WriteLine(report);

        foreach (string storeName in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            StoreScopeInfo scope = _fixture.GetScope(storeName);
            IndexQuery query = new()
            {
                Scope = scope.StorePrefix,
                Kinds = KindFilter.MessagesAndAttachments,
                Top = ProbeTop,
            };

            IReadOnlyList<IReadOnlyDictionary<string, object?>> rows =
                client.ExecuteRows(WsSqlBuilder.Build(query, ProbeTop), ProbeTop);

            int undated = 0;
            int firstUndated = -1;
            int lastDated = -1;
            for (int i = 0; i < rows.Count; i++)
            {
                bool hasDate = IndexRowMapper.Map(rows[i]).DateReceivedUtc.HasValue;
                if (hasDate)
                {
                    lastDated = i;
                }
                else
                {
                    undated++;
                    if (firstUndated < 0)
                    {
                        firstUndated = i;
                    }
                }
            }

            string verdict = undated == 0
                ? "no-undated-rows-in-sample"
                : firstUndated > lastDated ? "NULLS LAST (guard rarely fires)"
                : firstUndated == 0 && lastDated < 0 ? "NULLS FIRST (guard fires on every truncated search)"
                : "INTERLEAVED or NULLS FIRST (guard fires)";

            _output.WriteLine(
                $"store={storeName} rows={rows.Count} undated={undated} firstUndated={firstUndated} "
                + $"lastDated={lastDated} verdict={verdict}");
        }
    }

    /// <summary>
    /// THE RECOVERY QUERY, which is the one statement shape the T1 suite can only assert the
    /// TEXT of. If the provider rejects the 1601 literal, or treats the comparison as
    /// anything other than "has a value", the guard degrades to a flagged short answer on
    /// exactly the searches it exists to protect - so this must be run before the guarantee
    /// is called measured rather than constructed.
    /// </summary>
    [Fact]
    public void OrderKeyFloorPredicate_IsAccepted_AndAdmitsOnlyDatedRows()
    {
        IIndexClient client = IndexClientFactory.CreateAuto(out _);

        foreach (string storeName in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            StoreScopeInfo scope = _fixture.GetScope(storeName);
            IndexQuery query = new()
            {
                Scope = scope.StorePrefix,
                Kinds = KindFilter.MessagesAndAttachments,
                Top = ProbeTop,
            };

            IReadOnlyList<IReadOnlyDictionary<string, object?>> rows =
                client.ExecuteRows(WsSqlBuilder.Build(query, ProbeTop, true), ProbeTop);

            int undated = 0;
            foreach (IReadOnlyDictionary<string, object?> row in rows)
            {
                if (!IndexRowMapper.Map(row).DateReceivedUtc.HasValue)
                {
                    undated++;
                }
            }

            _output.WriteLine($"store={storeName} rows={rows.Count} undated={undated}");
            Assert.True(rows.Count > 0, $"store {storeName}: the floor predicate returned no rows at all");
            Assert.Equal(0, undated);
        }
    }

    /// <summary>
    /// THE GUARANTEE, on real data: the widened shape (every item class, gap B3) must never
    /// hand back fewer rows than the narrow pre-B3 shape it replaced. This is the assertion
    /// that would have failed on a NULLS-FIRST provider before the guard existed.
    /// </summary>
    [Fact]
    public void WidenedSearch_NeverReturnsFewerRowsThanTheOldMailKindShape()
    {
        foreach (string storeName in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            StoreScopeInfo scope = _fixture.GetScope(storeName);

            IndexSearchResult widened = _fixture.Service.Search(new IndexQuery
            {
                Scope = scope.StorePrefix,
                Kinds = KindFilter.MessagesOnly,
                Top = 25,
            });

            IndexSearchResult mailKindOnly = _fixture.Service.Search(new IndexQuery
            {
                Scope = scope.StorePrefix,
                Kinds = KindFilter.MailKindOnly,
                Top = 25,
            });

            _output.WriteLine(
                $"store={storeName} widened={widened.Hits.Count} mailKindOnly={mailKindOnly.Hits.Count} "
                + $"widenedScanned={widened.RowsScanned} widenedMs={widened.ElapsedMilliseconds} "
                + $"mailKindMs={mailKindOnly.ElapsedMilliseconds}");

            Assert.True(
                widened.Hits.Count >= mailKindOnly.Hits.Count,
                $"store {storeName}: widening COST rows ({widened.Hits.Count} < {mailKindOnly.Hits.Count})");
        }
    }
}
