using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using OutlookAI.Core.IndexSearch;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// How <c>LiveOrderKeyCollationTests.WidenedSearch_NeverReturnsFewerRowsThanTheOldMailKindShape</c>
/// sizes its search, so that whether the order-key guard decided the answer is ARITHMETIC rather than
/// luck. Decided on the maintainer's behalf 2026-10-03 (his answer to D74: "ensure there is no luck
/// involved").
/// <para>
/// <b>The luck it removes.</b> A store-scoped search of <c>Top n</c> over-fetches to
/// <see cref="IndexRowFilter.ComputeSqlTop"/> rows - <c>2n + 10</c>, so <c>TOP 60</c> for a top 25.
/// Rows with no <c>System.Message.DateReceived</c> can only take dated mail's place when more of them
/// sort ahead of that cut than the over-fetch leaves room for (<c>60 - 25 = 35</c>). With fewer, the
/// widened search returns as many dated rows WITHOUT the guard as with it, and the test passes having
/// told nothing apart. Until now the remedy was a margin: the bystander carried 42 undated rows, seven
/// more than 35 - true only while the over-fetch formula, the test's fixed <c>Top 25</c> and the number
/// of rows the index really leaves undated all stayed where they were, and silent the day any of them
/// moved.
/// </para>
/// <para>
/// <b>What replaces it.</b> The test COUNTS the undated rows the store holds and asks for the largest
/// <c>Top</c> whose over-fetch those rows out-number - <see cref="Size"/>, computed with the product's own
/// <see cref="IndexRowFilter.ComputeSqlTop"/>, so a change to the over-fetch moves the size with it.
/// Then, wherever the provider sorts a NULL, the arithmetic holds with no margin: if the undated rows
/// sort ahead of the dated ones the statement cannot hold <c>Top</c> dated rows, and the guard's
/// refetch is what makes the answer whole. And the test does not infer which happened: it runs the
/// unguarded statement itself and reads how many dated rows it held (<see cref="StatementAloneFallsShort"/>),
/// and it requires that count to be the one the store's own ordering predicts
/// (<see cref="AdmittedDatedRows"/> over the head of a wider sample of the same statement) - so a
/// provider whose NULL placement changed between two statements fails loudly instead of deciding the
/// verdict by chance.
/// </para>
/// <para>Pure: counts and rows in, sizes and counts out. <c>T1/OrderKeyContestTests</c> pins every rule.</para>
/// </summary>
/// <param name="Top">The search's <c>Top</c>: the largest whose over-fetch the undated rows out-number.</param>
/// <param name="SqlTop">The statement's <c>TOP</c> for that search - <see cref="IndexRowFilter.ComputeSqlTop"/>, scoped.</param>
/// <param name="UndatedRows">How many undated rows the store was counted to hold.</param>
public sealed record OrderKeyContest(int Top, int SqlTop, int UndatedRows)
{
    /// <summary>
    /// The fewest undated rows a store must hold for ANY store-scoped search to be sized so they
    /// out-number its over-fetch - the smallest count <see cref="Size"/> answers for. Twelve under
    /// today's over-fetch (<c>Top 1</c> fetches <c>TOP 12</c>); derived, never restated.
    /// </summary>
    public static int MinimumUndatedRows
    {
        get
        {
            for (int undated = 1; undated <= WsSqlBuilder.MaxTop; undated++)
            {
                if (Size(undated) != null)
                {
                    return undated;
                }
            }

            throw new System.InvalidOperationException("No undated count up to the statement ceiling can out-number the over-fetch.");
        }
    }

    /// <summary>The over-fetch room: how many rows the statement fetches beyond <see cref="Top"/>.</summary>
    public int OverFetchRoom => SqlTop - Top;

    /// <summary>
    /// The contest for a store holding <paramref name="undatedRows"/> rows with no received date: the
    /// LARGEST <c>Top</c> for which <c>ComputeSqlTop(Top) - undatedRows &lt; Top</c> - every statement row
    /// the undated rows do not take is one dated row at most, so if they all sort ahead of the dated
    /// ones the statement holds fewer than <c>Top</c> dated rows. Null when no <c>Top</c> of 1 or more
    /// qualifies (fewer than <see cref="MinimumUndatedRows"/>): that store can be searched, but it cannot
    /// tell the guard apart from its absence.
    /// </summary>
    public static OrderKeyContest? Size(int undatedRows)
    {
        if (undatedRows < 1)
        {
            return null;
        }

        for (int top = System.Math.Min(undatedRows, WsSqlBuilder.MaxTop); top >= 1; top--)
        {
            int sqlTop = IndexRowFilter.ComputeSqlTop(top, true, WsSqlBuilder.MaxTop);
            if (sqlTop - undatedRows < top)
            {
                return new OrderKeyContest(top, sqlTop, undatedRows);
            }
        }

        return null;
    }

    /// <summary>
    /// How many of <paramref name="rows"/> a search for <paramref name="kinds"/> admits AND can rank by
    /// date - the dated rows that statement could put in an answer. <paramref name="take"/> limits the
    /// count to the first rows, in the provider's order: the head of a wider sample predicts what a
    /// narrower statement of the same shape holds.
    /// </summary>
    public static int AdmittedDatedRows(IReadOnlyList<IndexHit> rows, KindFilter kinds, int take = int.MaxValue)
        => rows.Take(take).Count(r => IndexRowFilter.Keep(r, kinds) && r.DateReceivedUtc.HasValue);

    /// <summary>
    /// Whether the unguarded statement ALONE could not have answered <paramref name="top"/> dated rows:
    /// it was cut off (<paramref name="rowsReturned"/> reached <paramref name="sqlTop"/>, so more rows
    /// existed) and held fewer admitted dated rows than the answer needs. Exactly when the guard's
    /// refetch is load-bearing - <see cref="IndexOrderGuard.NeedsOrderKeyRefetch"/> fires on a superset
    /// of these, generously, and this is the subset where firing CHANGED the answer.
    /// </summary>
    public static bool StatementAloneFallsShort(int rowsReturned, int sqlTop, int admittedDated, int top)
        => rowsReturned >= sqlTop && admittedDated < top;

    /// <summary>The line the test prints for one store's contest, content-free.</summary>
    public string Describe()
        => string.Format(
            CultureInfo.InvariantCulture,
            "contest: {0} undated row(s) against Top {1}, fetched as TOP {2} ({3} row(s) of over-fetch room) - "
                + "if the undated rows sort ahead, the statement can hold at most {4} dated row(s), fewer than {1}",
            UndatedRows,
            Top,
            SqlTop,
            OverFetchRoom,
            SqlTop - UndatedRows);
}
