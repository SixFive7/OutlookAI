using System;
using System.Collections.Generic;
using System.Linq;

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
    [Trait("Writes", "Nothing")]
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
        // report, and "no-undated-rows-in-sample" is not a measurement. On the indexed guest the hub and
        // the bystander carry undated CONTACTS since 2026-10-03 (Q98 (f), and with all three kinds since
        // D62 (b)) - the one kind the index leaves without a received date - so this measures there;
        // anywhere they are missing, saying so out loud keeps it from passing as a measurement.
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
        "On the INDEXED test guest the hub and bystander populations carry appointments, contacts and tasks - four and "
        + "fourteen of each - since 2026-10-03 (D62 (b), built with --all-kinds; before that, twelve and forty-two undated "
        + "contacts with --undated-contacts): a contact saved into a PST is the one kind the index gives no "
        + "System.Message.DateReceived, so the contacts are the undated rows here, and the appointments and tasks are "
        + "dated by the plan, older than every mail item. None found means the populations were built without either "
        + "option, or the index dated the contacts after all - rebuild them with --all-kinds, and read corpus-indexed's "
        + "'with no received date' count (Docs/live-tier-on-the-vm.md section 3b). A store needs at least twelve such rows "
        + "before a widened search can be sized so they out-number its over-fetch (T2/OrderKeyContest).";

    /// <summary>
    /// <paramref name="scope"/>'s widest sample - the shipped statement's own shape at <see cref="ProbeTop"/>,
    /// mapped - and how many of its rows carry no received date: the order-key tests' population.
    /// </summary>
    private static (IReadOnlyList<IndexHit> Sample, int Undated) SampleUndatedRows(IIndexClient client, string scope)
    {
        IndexQuery query = new()
        {
            Scope = scope,
            Kinds = KindFilter.MessagesAndAttachments,
            Top = ProbeTop,
        };

        var sample = new List<IndexHit>();
        int undated = 0;
        foreach (IReadOnlyDictionary<string, object?> row in client.ExecuteRows(WsSqlBuilder.Build(query, ProbeTop), ProbeTop))
        {
            IndexHit hit = IndexRowMapper.Map(row);
            sample.Add(hit);
            if (!hit.DateReceivedUtc.HasValue)
            {
                undated++;
            }
        }

        return (sample, undated);
    }

    /// <summary>How many rows of <paramref name="scope"/>'s widest sample carry no received date - the order-key tests' population.</summary>
    private static int CountUndatedRows(IIndexClient client, string scope) => SampleUndatedRows(client, scope).Undated;

    /// <summary>
    /// THE RECOVERY QUERY, which is the one statement shape the T1 suite can only assert the
    /// TEXT of. If the provider rejects the 1601 literal, or treats the comparison as
    /// anything other than "has a value", the guard degrades to a flagged short answer on
    /// exactly the searches it exists to protect - so this must be run before the guarantee
    /// is called measured rather than constructed.
    /// </summary>
    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Writes", "Nothing")]
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
    /// THE GUARANTEE, on real data: the widened shape (every item class, gap B3) must never hand back
    /// fewer rows - or fewer DATED rows - than the narrow pre-B3 shape it replaced. This is the assertion
    /// that would have failed on a NULLS-FIRST provider before the guard existed.
    /// <para>
    /// <b>Sized, not margined - decided 2026-10-03 (D74, "ensure there is no luck involved").</b> A store
    /// is CONTESTED when it holds enough undated rows for a search to be sized so they out-number its
    /// over-fetch (<see cref="OrderKeyContest.Size"/>): there, if the provider sorts them ahead, the
    /// statement alone cannot hold <c>Top</c> dated rows, and only the guard's refetch makes the answer
    /// whole. The test runs that unguarded statement itself, requires what it held to be exactly what the
    /// store's own ordering predicts, and says in so many words whether the guard was load-bearing. Every
    /// other store is still compared, at <c>Top 25</c>, as the plain regression check it always was.
    /// </para>
    /// </summary>
    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Writes", "Nothing")]
    public void WidenedSearch_NeverReturnsFewerRowsThanTheOldMailKindShape()
    {
        const int UncontestedTop = 25;
        IIndexClient client = IndexClientFactory.CreateAuto(out _);
        var contested = new List<string>();

        foreach (string storeName in Indexed)
        {
            StoreScopeInfo scope = _fixture.GetScope(storeName);
            (IReadOnlyList<IndexHit> sample, int undatedRows) = SampleUndatedRows(client, scope.StorePrefix);
            OrderKeyContest? contest = OrderKeyContest.Size(undatedRows);
            int top = contest?.Top ?? UncontestedTop;
            int sqlTop = IndexRowFilter.ComputeSqlTop(top, true, WsSqlBuilder.MaxTop);
            _output.WriteLine(contest == null
                ? $"store={storeName} undatedInSample={undatedRows}: fewer than {OrderKeyContest.MinimumUndatedRows} undated rows, "
                    + $"so no search here can be sized to make them out-number its over-fetch - compared at Top {UncontestedTop}, not contested"
                : $"store={storeName} {contest.Describe()}");

            IndexQuery widenedQuery = new()
            {
                Scope = scope.StorePrefix,
                Kinds = KindFilter.MessagesOnly,
                Top = top,
            };
            IndexSearchResult widened = _fixture.Service.Search(widenedQuery);
            IndexSearchResult mailKindOnly = _fixture.Service.Search(new IndexQuery
            {
                Scope = scope.StorePrefix,
                Kinds = KindFilter.MailKindOnly,
                Top = top,
            });

            // The statement the widened search starts from, run WITHOUT the guard: how many dated rows it
            // could have answered with on its own. The prediction is the same count over the head of the
            // wider sample above - the same statement, the same ordering, cut at the same row - so the two
            // differ only if the provider placed the undated rows differently in two statements, and then
            // whether the guard decided anything would be chance. That fails here, out loud. The one other
            // way they can differ is the index moving between the two reads - the indexer still taking in
            // an earlier test's writes - so a disagreement is read again, sample and statement together, up
            // to twice more before it counts.
            IReadOnlyList<IndexHit> statement = Array.Empty<IndexHit>();
            int statementDated = 0;
            int predictedDated = 0;
            for (int attempt = 1; attempt <= 3; attempt++)
            {
                if (attempt > 1)
                {
                    _output.WriteLine($"store={storeName} attempt {attempt}: the statement and the sample disagreed - reading both again");
                    sample = SampleUndatedRows(client, scope.StorePrefix).Sample;
                }

                statement = client.ExecuteRows(WsSqlBuilder.Build(widenedQuery, sqlTop), sqlTop).Select(IndexRowMapper.Map).ToList();
                statementDated = OrderKeyContest.AdmittedDatedRows(statement, KindFilter.MessagesOnly);
                predictedDated = OrderKeyContest.AdmittedDatedRows(sample, KindFilter.MessagesOnly, sqlTop);
                if (statementDated == predictedDated)
                {
                    break;
                }
            }

            bool guardDecided = OrderKeyContest.StatementAloneFallsShort(statement.Count, sqlTop, statementDated, top);

            int widenedDated = widened.Hits.Count(h => h.DateReceivedUtc.HasValue);
            int mailKindDated = mailKindOnly.Hits.Count(h => h.DateReceivedUtc.HasValue);
            _output.WriteLine(
                $"store={storeName} top={top} sqlTop={sqlTop} statementRows={statement.Count} statementDated={statementDated} "
                + $"predictedDated={predictedDated} widened={widened.Hits.Count} (dated {widenedDated}) "
                + $"mailKindOnly={mailKindOnly.Hits.Count} (dated {mailKindDated}) widenedScanned={widened.RowsScanned} "
                + $"widenedMs={widened.ElapsedMilliseconds} mailKindMs={mailKindOnly.ElapsedMilliseconds} guard="
                + (guardDecided
                    ? "LOAD-BEARING (the statement alone held fewer dated rows than Top; the refetch made the answer whole)"
                    : "not needed (the undated rows sorted after the cut, or the store holds too few to reach it)"));

            // The prediction holds only where the sample reaches as far as the statement does - a sample
            // the provider did not cut off holds every row there is.
            if (sample.Count < ProbeTop || sample.Count >= sqlTop)
            {
                Assert.True(
                    statementDated == predictedDated,
                    $"store {storeName}: the unguarded TOP {sqlTop} statement held {statementDated} dated row(s), but the head of "
                    + $"the TOP {ProbeTop} sample of the same statement holds {predictedDated} - the provider placed the rows with "
                    + "no received date differently in two statements of one shape, so whether the guard decided this answer "
                    + "would be luck");
            }

            if (guardDecided)
            {
                Assert.True(
                    widened.RowsScanned > statement.Count,
                    $"store {storeName}: the statement alone fell short of Top {top}, yet the search scanned only "
                    + $"{widened.RowsScanned} row(s) - the order-key refetch did not run");
            }

            Assert.True(
                widened.Hits.Count >= mailKindOnly.Hits.Count,
                $"store {storeName}: widening COST rows ({widened.Hits.Count} < {mailKindOnly.Hits.Count})");

            // The guarantee in the words IndexOrderGuard states it: an undated row can never reduce the
            // number of DATED rows a search returns. The count above cannot see that - since gap B3 an
            // appointment IS a hit, so a page of appointments counts the same as a page of mail - and
            // this can. Added 2026-09-24, with the undated items that make it testable on a guest.
            Assert.True(
                widenedDated >= mailKindDated,
                $"store {storeName}: widening COST DATED rows ({widenedDated} < {mailKindDated}) - rows with no "
                + "received date took slots dated mail should have had");

            if (contest != null)
            {
                contested.Add(storeName);
            }
        }

        LivePopulationCoverage.Require(
            _fixture.Settings,
            contested,
            $"an indexed store holding at least {OrderKeyContest.MinimumUndatedRows} rows with no System.Message.DateReceived - "
                + "enough that a widened search can be sized so they out-number its over-fetch",
            "the proof that undated rows cannot displace dated mail from a widened search",
            UndatedRemedy,
            _output.WriteLine);
    }
}
