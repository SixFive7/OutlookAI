using System;
using System.Collections.Generic;
using System.Linq;

using OutlookAI.Core.IndexSearch;
using OutlookAI.McpServer.Tests.T2;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 for <see cref="OrderKeyContest"/>: how the live widened-search test sizes its search so that the
/// order-key guard deciding the answer is arithmetic, not luck (D74, decided 2026-10-03: "ensure there is
/// no luck involved"). Before it, the test searched at a fixed <c>Top 25</c> and leaned on the bystander
/// holding 42 undated rows - seven more than the 35 rows of over-fetch room that search has.
/// </summary>
public sealed class OrderKeyContestTests
{
    private const string MessageUrlPrefix = "mapi16://{SID}/bystander@vm.invalid($ab12)/0/Inbox/item-";

    private static readonly DateTime Noon = new(2026, 10, 03, 12, 00, 00, DateTimeKind.Utc);

    [Theory]
    [InlineData(12)]
    [InlineData(13)]
    [InlineData(14)]
    [InlineData(42)]
    [InlineData(100)]
    [InlineData(1000)]
    public void Size_IsTheLargestTopWhoseOverFetchTheUndatedRowsOutnumber(int undated)
    {
        OrderKeyContest contest = OrderKeyContest.Size(undated)!;

        Assert.NotNull(contest);
        Assert.Equal(undated, contest.UndatedRows);
        Assert.Equal(IndexRowFilter.ComputeSqlTop(contest.Top, true, WsSqlBuilder.MaxTop), contest.SqlTop);

        // The contest: every statement row the undated rows do not take is at most one dated row, so if
        // they all sort ahead the statement holds fewer than Top dated rows - no margin, just this.
        Assert.True(contest.SqlTop - undated < contest.Top);

        // And it is the LARGEST such Top: one more and the over-fetch could hold the answer again.
        int next = contest.Top + 1;
        Assert.True(IndexRowFilter.ComputeSqlTop(next, true, WsSqlBuilder.MaxTop) - undated >= next);
    }

    [Fact]
    public void Size_FollowsTheProductsOverFetch_TwelveUndatedRowsAreTheFloorToday()
    {
        // Derived, never restated: Top 1 fetches TOP 12 (2 * 1 + 10), so twelve undated rows are the
        // fewest that can out-number any scoped over-fetch. If the formula moves, this number moves with
        // it - and so does every Size - which is the point.
        Assert.Equal(12, IndexRowFilter.ComputeSqlTop(1, true, WsSqlBuilder.MaxTop));
        Assert.Equal(12, OrderKeyContest.MinimumUndatedRows);
        for (int undated = 0; undated < OrderKeyContest.MinimumUndatedRows; undated++)
        {
            Assert.Null(OrderKeyContest.Size(undated));
        }

        // The sizes the populations produce: the bystander's fourteen contacts give Top 3, fetched as
        // TOP 16; the old forty-two gave Top 31 - never the fixed Top 25 the old test searched at.
        Assert.Equal((3, 16), (OrderKeyContest.Size(14)!.Top, OrderKeyContest.Size(14)!.SqlTop));
        Assert.Equal((1, 12), (OrderKeyContest.Size(12)!.Top, OrderKeyContest.Size(12)!.SqlTop));
        Assert.Equal((31, 72), (OrderKeyContest.Size(42)!.Top, OrderKeyContest.Size(42)!.SqlTop));
    }

    [Theory]
    [InlineData(12)]
    [InlineData(14)]
    [InlineData(42)]
    public void UnderNullsFirst_TheStatementAloneAlwaysFallsShort(int undated)
    {
        // Driven both ways, like IndexOrderDisplacementTests: nobody had measured the provider's NULL
        // collation when this was written, and the sizing must hold under either.
        OrderKeyContest contest = OrderKeyContest.Size(undated)!;
        List<IndexHit> nullsFirst = Undated(undated).Concat(Dated(300)).ToList();

        IReadOnlyList<IndexHit> statement = nullsFirst.Take(contest.SqlTop).ToList();
        int dated = OrderKeyContest.AdmittedDatedRows(statement, KindFilter.MessagesOnly);

        Assert.True(dated < contest.Top);
        Assert.True(OrderKeyContest.StatementAloneFallsShort(statement.Count, contest.SqlTop, dated, contest.Top));

        // ...and the head of a wider sample predicts it exactly.
        Assert.Equal(dated, OrderKeyContest.AdmittedDatedRows(nullsFirst.Take(500).ToList(), KindFilter.MessagesOnly, contest.SqlTop));
    }

    [Theory]
    [InlineData(12)]
    [InlineData(14)]
    [InlineData(42)]
    public void UnderNullsLast_TheStatementHoldsTheAnswer_AndTheGuardIsNotLoadBearing(int undated)
    {
        OrderKeyContest contest = OrderKeyContest.Size(undated)!;
        List<IndexHit> nullsLast = Dated(300).Concat(Undated(undated)).ToList();

        IReadOnlyList<IndexHit> statement = nullsLast.Take(contest.SqlTop).ToList();
        int dated = OrderKeyContest.AdmittedDatedRows(statement, KindFilter.MessagesOnly);

        Assert.True(dated >= contest.Top);
        Assert.False(OrderKeyContest.StatementAloneFallsShort(statement.Count, contest.SqlTop, dated, contest.Top));
    }

    [Fact]
    public void AStatementThatWasNotCutOff_NeverFallsShort_WhateverItHeld()
    {
        // Fewer rows than TOP means every matching row is in hand: nothing can have been displaced.
        Assert.False(OrderKeyContest.StatementAloneFallsShort(rowsReturned: 15, sqlTop: 16, admittedDated: 0, top: 3));
        Assert.True(OrderKeyContest.StatementAloneFallsShort(rowsReturned: 16, sqlTop: 16, admittedDated: 2, top: 3));
        Assert.False(OrderKeyContest.StatementAloneFallsShort(rowsReturned: 16, sqlTop: 16, admittedDated: 3, top: 3));
    }

    [Fact]
    public void AdmittedDatedRows_CountsOnlyRowsTheSearchWouldAdmit_AndCanRank()
    {
        IndexHit attachment = IndexRowMapper.Map(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = MessageUrlPrefix + "with-attachment/at=1:report.txt",
            ["System.Kind"] = new[] { "document" },
            ["System.Message.DateReceived"] = Noon,
        });
        List<IndexHit> rows = Dated(3).Concat(Undated(2)).Append(attachment).ToList();

        Assert.Equal(3, OrderKeyContest.AdmittedDatedRows(rows, KindFilter.MessagesOnly));
        Assert.Equal(2, OrderKeyContest.AdmittedDatedRows(rows, KindFilter.MessagesOnly, take: 2));
    }

    [Fact]
    public void Describe_IsContentFree_AndSaysWhyTheGuardMustDecide()
    {
        string line = OrderKeyContest.Size(14)!.Describe();

        Assert.Contains("14 undated row(s) against Top 3", line, StringComparison.Ordinal);
        Assert.Contains("TOP 16", line, StringComparison.Ordinal);
        Assert.Contains("at most 2 dated row(s), fewer than 3", line, StringComparison.Ordinal);
    }

    private static IEnumerable<IndexHit> Dated(int count)
        => Enumerable.Range(0, count).Select(i => IndexRowMapper.Map(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = MessageUrlPrefix + "dated-" + i.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["System.Kind"] = new[] { "email" },
            ["System.Message.DateReceived"] = Noon.AddMinutes(-i),
        }));

    private static IEnumerable<IndexHit> Undated(int count)
        => Enumerable.Range(0, count).Select(i => IndexRowMapper.Map(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = MessageUrlPrefix + "undated-" + i.ToString(System.Globalization.CultureInfo.InvariantCulture),
            ["System.Kind"] = new[] { "contact", "communication" },
        }));
}
