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
[Collection(LiveCollections.Phase1)]
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
    /// The stores the index tier measures - the settings' INDEXED list, never the watched one, and
    /// refused rather than empty (see <see cref="LiveTestSettings.RequireIndexedStores"/>).
    /// </summary>
    private List<string> Indexed => _fixture.Settings.RequireIndexedStores().ToList();

    /// <summary>
    /// THE MEASUREMENT. Runs the shipped widened statement over each store and records where
    /// the undated rows landed. Asserts nothing about the answer, because either answer is
    /// legitimate provider behaviour: what it produces is the number that belongs in
    /// Docs/magic-numbers.md beside the guard.
    /// </summary>
    [Fact]
    [Trait("Requires", "SearchIndex")]
    public void NullCollation_UnderDateReceivedDescending_IsMeasured()
    {
        IIndexClient client = IndexClientFactory.CreateAuto(out string report);
        _output.WriteLine(report);
        var measured = new List<string>();

        foreach (string storeName in Indexed)
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
            if (undated > 0 && lastDated >= 0)
            {
                measured.Add(storeName);
            }
        }

        // A sample with no undated row - or with nothing BUT undated rows - has no NULL collation to
        // report, and "no-undated-rows-in-sample" is not a measurement. That was every store on a test
        // guest until the fixture populations carried undated items (2026-09-24); saying so out loud
        // keeps it from quietly becoming true again.
        LivePopulationCoverage.Require(
            _fixture.Settings,
            measured,
            "an indexed store whose sampled rows include both dated rows and rows with no System.Message.DateReceived",
            "the NULL-collation measurement",
            UndatedRemedy,
            _output.WriteLine);
    }

    /// <summary>What every order-key test prints when there was nothing undated to measure.</summary>
    private const string UndatedRemedy =
        "On a test guest the hub and bystander populations carry undated appointments, contacts and tasks "
        + "(corpus-build --population hub|bystander, generator v2; Docs/live-tier-on-the-vm.md section 3b) - rebuild "
        + "them, and wait for corpus-indexed to report them indexed, before the run.";

    /// <summary>How many rows of <paramref name="scope"/>'s widest sample carry no received date - the order-key tests' population.</summary>
    private static int CountUndatedRows(IIndexClient client, string scope)
    {
        IndexQuery query = new()
        {
            Scope = scope,
            Kinds = KindFilter.MessagesAndAttachments,
            Top = ProbeTop,
        };

        int undated = 0;
        foreach (IReadOnlyDictionary<string, object?> row in client.ExecuteRows(WsSqlBuilder.Build(query, ProbeTop), ProbeTop))
        {
            if (!IndexRowMapper.Map(row).DateReceivedUtc.HasValue)
            {
                undated++;
            }
        }

        return undated;
    }

    /// <summary>
    /// THE RECOVERY QUERY, which is the one statement shape the T1 suite can only assert the
    /// TEXT of. If the provider rejects the 1601 literal, or treats the comparison as
    /// anything other than "has a value", the guard degrades to a flagged short answer on
    /// exactly the searches it exists to protect - so this must be run before the guarantee
    /// is called measured rather than constructed.
    /// </summary>
    [Fact]
    [Trait("Requires", "SearchIndex")]
    public void OrderKeyFloorPredicate_IsAccepted_AndAdmitsOnlyDatedRows()
    {
        IIndexClient client = IndexClientFactory.CreateAuto(out _);
        var excludedSomething = new List<string>();

        foreach (string storeName in Indexed)
        {
            StoreScopeInfo scope = _fixture.GetScope(storeName);

            // What the floor has to exclude: the undated rows the SAME statement returns without it.
            // With none, "undated = 0" below is true of the store rather than of the predicate.
            int withoutFloor = CountUndatedRows(client, scope.StorePrefix);
            if (withoutFloor > 0)
            {
                excludedSomething.Add(storeName);
            }
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

            _output.WriteLine($"store={storeName} rows={rows.Count} undated={undated} undatedWithoutTheFloor={withoutFloor}");
            Assert.True(rows.Count > 0, $"store {storeName}: the floor predicate returned no rows at all");
            Assert.Equal(0, undated);
        }

        LivePopulationCoverage.Require(
            _fixture.Settings,
            excludedSomething,
            "an indexed store whose unfloored statement returns rows with no System.Message.DateReceived",
            "the proof that the order-key floor predicate excludes undated rows",
            UndatedRemedy,
            _output.WriteLine);
    }

    /// <summary>
    /// THE GUARANTEE, on real data: the widened shape (every item class, gap B3) must never
    /// hand back fewer rows than the narrow pre-B3 shape it replaced. This is the assertion
    /// that would have failed on a NULLS-FIRST provider before the guard existed.
    /// </summary>
    [Fact]
    [Trait("Requires", "SearchIndex")]
    public void WidenedSearch_NeverReturnsFewerRowsThanTheOldMailKindShape()
    {
        IIndexClient client = IndexClientFactory.CreateAuto(out _);
        var contested = new List<string>();

        foreach (string storeName in Indexed)
        {
            StoreScopeInfo scope = _fixture.GetScope(storeName);
            int undatedRows = CountUndatedRows(client, scope.StorePrefix);
            if (undatedRows > 0)
            {
                contested.Add(storeName);
            }

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

            int widenedDated = widened.Hits.Count(h => h.DateReceivedUtc.HasValue);
            int mailKindDated = mailKindOnly.Hits.Count(h => h.DateReceivedUtc.HasValue);
            _output.WriteLine(
                $"store={storeName} widened={widened.Hits.Count} (dated {widenedDated}) "
                + $"mailKindOnly={mailKindOnly.Hits.Count} (dated {mailKindDated}) undatedInScope={undatedRows} "
                + $"widenedScanned={widened.RowsScanned} widenedMs={widened.ElapsedMilliseconds} "
                + $"mailKindMs={mailKindOnly.ElapsedMilliseconds}");

            Assert.True(
                widened.Hits.Count >= mailKindOnly.Hits.Count,
                $"store {storeName}: widening COST rows ({widened.Hits.Count} < {mailKindOnly.Hits.Count})");

            // The guarantee in the words IndexOrderGuard states it: an undated row can never reduce
            // the number of DATED rows a search returns. The count above cannot see that - since gap B3
            // an appointment IS a hit, so a page of appointments counts the same as a page of mail - and
            // this can. Added 2026-09-24, with the undated items that make it testable on a guest.
            Assert.True(
                widenedDated >= mailKindDated,
                $"store {storeName}: widening COST DATED rows ({widenedDated} < {mailKindDated}) - rows with no "
                + "received date took slots dated mail should have had");
        }

        LivePopulationCoverage.Require(
            _fixture.Settings,
            contested,
            "an indexed store holding rows with no System.Message.DateReceived for the widened shape to rank",
            "the proof that undated rows cannot displace dated mail from a widened search",
            UndatedRemedy,
            _output.WriteLine);
    }
}
