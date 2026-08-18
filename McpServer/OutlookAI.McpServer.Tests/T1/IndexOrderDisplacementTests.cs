using System;
using System.Collections.Generic;
using System.Linq;

using OutlookAI.Core.IndexSearch;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1: the widening of gap B3 must not be able to COST mail. Since 2026-08-18 a
/// store-scoped statement carries no Kind predicate, so appointments and contacts of
/// folders the COM tiers never open are candidates for the same <c>SELECT TOP n ... ORDER
/// BY System.Message.DateReceived DESC</c> that real mail competes in - and they carry no
/// received date, so the provider's NULL collation, which nobody here has measured, decides
/// whether they fill the n.
/// <para>
/// If NULLs sort first, the pre-B3 answer and the post-B3 answer differ by every mail in the
/// result: a full page of appointments, nothing flagged short, nothing dropped. These tests
/// are the guarantee that this cannot happen WHATEVER that collation turns out to be, which
/// is why the fake provider below is driven in both directions rather than the likely one.
/// </para>
/// <para>
/// The guarantee, stated as the tests state it: a row the index cannot date never reduces the
/// number of dated rows a search returns. Dated non-mail (a meeting request, a bounce report)
/// still competes with mail on date, because under the admission rule those ARE mail.
/// </para>
/// </summary>
public sealed class IndexOrderDisplacementTests
{
    private const string StoreScope = "mapi16://{SID}/alice@example.com($ab12)";

    private const string MessageUrlPrefix = StoreScope + "/0/Inbox/item-";

    /// <summary>Top 5 under a scope over-fetches to 5 * 2 + 10 (IndexRowFilter.ComputeSqlTop).</summary>
    private const int SqlTopForTop5 = 20;

    private static readonly DateTime Noon = new(2026, 08, 18, 12, 00, 00, DateTimeKind.Utc);

    // ------------------------------------------------------------------ the pure decisions

    [Fact]
    public void OrderKeyPresence_IsAComparison_BecauseWsSqlHasNoNullTest()
    {
        // A comparison plus three-valued logic is the only null test WS-SQL offers: a row
        // with no value fails it. The floor is the FILETIME epoch, so nothing that HAS a
        // date is excluded by it - the predicate must not become a date filter.
        Assert.Equal(
            "System.Message.DateReceived >= '1601-01-01 00:00:00'",
            WsSqlBuilder.BuildOrderKeyPresence(IndexOrder.DateReceivedDescending));
        Assert.Equal("System.Size >= 0", WsSqlBuilder.BuildOrderKeyPresence(IndexOrder.SizeDescending));
        Assert.Equal(new DateTime(1601, 01, 01, 00, 00, 00, DateTimeKind.Utc), WsSqlBuilder.OrderKeyFloorUtc);
    }

    [Fact]
    public void ThePresencePredicate_IsNeverOnTheStatementThatAnswersTheSearch()
    {
        IndexQuery query = new() { Scope = StoreScope, Kinds = KindFilter.MessagesOnly, Top = 5 };

        string search = WsSqlBuilder.Build(query, SqlTopForTop5);
        string refetch = WsSqlBuilder.Build(query, SqlTopForTop5, true);

        // It NARROWS, so it belongs only to the recovery query. Emitting it on the search
        // itself would drop the undated rows B3 decided to admit - the widening undone by
        // the guard that exists to protect it.
        Assert.DoesNotContain("1601-01-01", search, StringComparison.Ordinal);
        Assert.Contains(WsSqlBuilder.BuildOrderKeyPresence(IndexOrder.DateReceivedDescending), refetch, StringComparison.Ordinal);
        Assert.Equal(search.Replace(
            " ORDER BY",
            " AND " + WsSqlBuilder.BuildOrderKeyPresence(IndexOrder.DateReceivedDescending) + " ORDER BY",
            StringComparison.Ordinal),
            refetch);
    }

    [Fact]
    public void HasOrderKey_ReadsTheColumnTheOrderingActuallyRanksBy()
    {
        IndexHit dated = Row("a", Noon, 100);
        IndexHit undated = Row("b", null, 100);
        IndexHit sizeless = Row("c", Noon, null);

        Assert.True(IndexOrderGuard.HasOrderKey(dated, IndexOrder.DateReceivedDescending));
        Assert.False(IndexOrderGuard.HasOrderKey(undated, IndexOrder.DateReceivedDescending));

        // Size ordering asks a different column, so the same row can be rankable under one
        // ordering and not the other. A guard that only knew about dates would leave the
        // size-ordered shape exposed to the identical defect.
        Assert.True(IndexOrderGuard.HasOrderKey(undated, IndexOrder.SizeDescending));
        Assert.False(IndexOrderGuard.HasOrderKey(sizeless, IndexOrder.SizeDescending));
    }

    [Theory]
    // Truncated AND an unrankable row present: the only combination where the provider can
    // have hidden rows the client can no longer see.
    [InlineData(20, 20, true, true)]
    // Not truncated: every matching row is in hand, so their order cannot have cost anything.
    [InlineData(19, 20, true, false)]
    // Truncated but no unrankable row came back: none sorted above the cut, so none took a
    // slot. This is the NULLs-last provider, and it pays for no second query.
    [InlineData(20, 20, false, false)]
    [InlineData(0, 20, false, false)]
    public void NeedsOrderKeyRefetch_IsTruncationAndDisplacementTogether(
        int rowsReturned, int sqlTop, bool anyMissing, bool expected)
    {
        Assert.Equal(expected, IndexOrderGuard.NeedsOrderKeyRefetch(rowsReturned, sqlTop, anyMissing));
    }

    [Fact]
    public void RankableFirst_PutsUnrankableRowsLast_AndKeepsProviderOrderOnTies()
    {
        IndexHit newest = Row("newest", Noon, 10);
        IndexHit tieA = Row("tie-a", Noon.AddHours(-1), 10);
        IndexHit tieB = Row("tie-b", Noon.AddHours(-1), 10);
        IndexHit undatedFirst = Row("undated-1", null, 10);
        IndexHit undatedSecond = Row("undated-2", null, 10);

        IReadOnlyList<IndexHit> ordered = IndexOrderGuard.RankableFirst(
            new[] { undatedFirst, tieA, undatedSecond, newest, tieB }, IndexOrder.DateReceivedDescending);

        Assert.Equal(
            new[] { "newest", "tie-a", "tie-b", "undated-1", "undated-2" },
            ordered.Select(Name).ToArray());
    }

    [Fact]
    public void RankableFirst_RanksBySizeWhenThatIsTheOrdering()
    {
        IndexHit big = Row("big", null, 900);
        IndexHit small = Row("small", Noon, 10);
        IndexHit sizeless = Row("sizeless", Noon, null);

        Assert.Equal(
            new[] { "big", "small", "sizeless" },
            IndexOrderGuard.RankableFirst(new[] { sizeless, small, big }, IndexOrder.SizeDescending).Select(Name).ToArray());
    }

    [Fact]
    public void Merge_IsAUnion_SoARefetchCanOnlyEverAdd()
    {
        IndexHit shared = Row("shared", Noon, 10);
        IndexHit onlyPrimary = Row("primary", Noon, 10);
        IndexHit onlyRefetch = Row("refetch", Noon, 10);

        IReadOnlyList<IndexHit> merged = IndexOrderGuard.Merge(
            new[] { onlyPrimary, shared }, new[] { shared, onlyRefetch });

        Assert.Equal(new[] { "primary", "shared", "refetch" }, merged.Select(Name).ToArray());
    }

    // ------------------------------------------------------- the guarantee, end to end

    [Fact]
    public void NullsFirstProvider_CannotHideMailBehindUndatedRows()
    {
        // THE REGRESSION THIS EXISTS FOR. The statement fills its whole TOP with undated
        // rows, so every dated mail in the folder is behind the cut and invisible. Before
        // the guard this returned five appointments and reported nothing unusual.
        List<IReadOnlyDictionary<string, object?>> appointments = Rows(SqlTopForTop5, "appt", i => null);
        List<IReadOnlyDictionary<string, object?>> mail = Rows(6, "mail", i => Noon.AddMinutes(-i));

        ScriptedIndexClient client = new(sql => IsRefetch(sql) ? mail : appointments);
        IndexSearchResult result = new IndexSearchService(client).Search(Top5Query());

        Assert.Equal(2, client.Statements.Count);
        Assert.Equal(new[] { "mail-0", "mail-1", "mail-2", "mail-3", "mail-4" }, result.Hits.Select(Name).ToArray());
        Assert.False(result.CandidatesExhausted);
    }

    [Fact]
    public void NullsLastProvider_PaysForNoSecondStatement()
    {
        // The likely collation, and the reason the guard is conditional rather than a second
        // query on every search: no undated row came back, so none can have taken a slot.
        List<IReadOnlyDictionary<string, object?>> mail = Rows(SqlTopForTop5, "mail", i => Noon.AddMinutes(-i));

        ScriptedIndexClient client = new(sql => mail);
        IndexSearchResult result = new IndexSearchService(client).Search(Top5Query());

        Assert.Single(client.Statements);
        Assert.Equal(new[] { "mail-0", "mail-1", "mail-2", "mail-3", "mail-4" }, result.Hits.Select(Name).ToArray());
    }

    [Fact]
    public void RowsAlreadyInHand_AreNotThrownAwayByTheTrim()
    {
        // The second displacement point, and the one no refetch can fix: the statement was
        // NOT cut off, so the mail is already in the client's hands, and taking the first
        // Top rows in provider order would still discard it. Nothing here queries twice.
        List<IReadOnlyDictionary<string, object?>> mixed = new();
        mixed.AddRange(Rows(5, "appt", i => null));
        mixed.AddRange(Rows(3, "mail", i => Noon.AddMinutes(-i)));

        ScriptedIndexClient client = new(sql => mixed);
        IndexSearchResult result = new IndexSearchService(client).Search(Top5Query());

        Assert.Single(client.Statements);
        Assert.Equal(
            new[] { "mail-0", "mail-1", "mail-2", "appt-0", "appt-1" },
            result.Hits.Select(Name).ToArray());
    }

    [Fact]
    public void AFailedRefetch_StillAnswers_ButSaysTheListMayBeShort()
    {
        // The recovery query is the one statement whose shape this repo has not run against
        // the real provider. If it is ever rejected, the search must still answer with what
        // it has and must not claim that answer is complete.
        List<IReadOnlyDictionary<string, object?>> appointments = Rows(SqlTopForTop5, "appt", i => null);

        ScriptedIndexClient client = new(sql => IsRefetch(sql)
            ? throw new InvalidOperationException("provider rejected the predicate")
            : appointments);
        IndexSearchResult result = new IndexSearchService(client).Search(Top5Query());

        Assert.Equal(2, client.Statements.Count);
        Assert.Equal(5, result.Hits.Count);
        Assert.True(result.CandidatesExhausted);
    }

    [Fact]
    public void CountersCoverEveryRowBothStatementsReturned()
    {
        // rowsScanned/rowsDropped reach the payload, so they have to mean what their names
        // say across a two-statement search. They used to stop counting at the first Top
        // admitted rows, which under-reported by the whole unexamined tail.
        List<IReadOnlyDictionary<string, object?>> appointments = Rows(SqlTopForTop5, "appt", i => null);
        List<IReadOnlyDictionary<string, object?>> mail = Rows(6, "mail", i => Noon.AddMinutes(-i));

        ScriptedIndexClient client = new(sql => IsRefetch(sql) ? mail : appointments);
        IndexSearchResult result = new IndexSearchService(client).Search(Top5Query());

        Assert.Equal(SqlTopForTop5 + 6, result.RowsScanned);
        Assert.Equal(0, result.RowsDropped);
    }

    // ------------------------------------------------------------------------- fixtures

    private static IndexQuery Top5Query()
    {
        return new IndexQuery { Scope = StoreScope, Kinds = KindFilter.MessagesOnly, Top = 5 };
    }

    private static bool IsRefetch(string sql)
    {
        return sql.Contains(
            WsSqlBuilder.BuildOrderKeyPresence(IndexOrder.DateReceivedDescending), StringComparison.Ordinal);
    }

    private static string Name(IndexHit hit)
    {
        return hit.ItemUrl.Substring(MessageUrlPrefix.Length);
    }

    private static IndexHit Row(string name, DateTime? received, long? size)
    {
        return IndexRowMapper.Map(RowDictionary(name, received, size));
    }

    private static List<IReadOnlyDictionary<string, object?>> Rows(
        int count, string namePrefix, Func<int, DateTime?> receivedAt)
    {
        List<IReadOnlyDictionary<string, object?>> rows = new(count);
        for (int i = 0; i < count; i++)
        {
            rows.Add(RowDictionary(namePrefix + "-" + i.ToString(System.Globalization.CultureInfo.InvariantCulture), receivedAt(i), 1000));
        }

        return rows;
    }

    private static IReadOnlyDictionary<string, object?> RowDictionary(string name, DateTime? received, long? size)
    {
        Dictionary<string, object?> row = new(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = MessageUrlPrefix + name,
            ["System.Kind"] = received.HasValue ? new[] { "email" } : new[] { "calendar" },
        };

        if (received.HasValue)
        {
            row["System.Message.DateReceived"] = received.Value;
        }

        if (size.HasValue)
        {
            row["System.Size"] = size.Value;
        }

        return row;
    }

    /// <summary>
    /// A provider whose row ORDER is dictated by the test rather than by Windows Search, so
    /// both NULL collations can be driven through the real service. It records every
    /// statement, because "did a second query run" is half of what these tests assert.
    /// </summary>
    private sealed class ScriptedIndexClient : IIndexClient
    {
        private readonly Func<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>> _answer;

        public ScriptedIndexClient(Func<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>> answer)
        {
            _answer = answer;
        }

        public List<string> Statements { get; } = new();

        public IndexProviderKind Provider => IndexProviderKind.OleDb;

        public IReadOnlyList<IReadOnlyDictionary<string, object?>> ExecuteRows(
            string sql, int maxRows, int? commandTimeoutSeconds = null)
        {
            Statements.Add(sql);
            IReadOnlyList<IReadOnlyDictionary<string, object?>> rows = _answer(sql);

            // The real clients stop draining at maxRows, and SELECT TOP caps the same
            // number, so a row list longer than the cap could never reach the service.
            return rows.Count <= maxRows ? rows : rows.Take(maxRows).ToList();
        }
    }
}
