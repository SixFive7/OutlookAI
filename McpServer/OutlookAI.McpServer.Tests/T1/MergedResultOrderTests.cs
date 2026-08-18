using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The order the caller asked for has to survive the MERGE, not just the query.
/// <para>
/// <c>search</c> asks the index provider for an <c>ORDER BY</c>, but that only decides which
/// rows come back. What the caller receives is index hits plus freshness-sweep hits (or, on
/// the exhaustive tier, a COM walk), and that combined list is re-sorted afterwards. Until
/// 2026-08-18 the re-sort was unconditionally by received date, so a size-ordered search was
/// silently reordered one layer above the query that had honoured it: the SQL asked for the
/// biggest mail, the answer came back newest-first, and nothing anywhere said so.
/// </para>
/// <para>
/// Not exploitable at the time - size ordering is not on the MCP tool surface - which is
/// precisely why it is worth a test. It was a correct query undone by a later line, and the
/// only thing that would have caught it is an assertion about the list the caller gets.
/// </para>
/// </summary>
public sealed class MergedResultOrderTests
{
    private static readonly DateTime Noon = new DateTime(2026, 08, 18, 12, 00, 00, DateTimeKind.Utc);

    [Fact]
    public void ASizeOrderedSearch_IsNotReSortedByDate()
    {
        // The defect, stated directly: the biggest mail is the oldest one here, so a list
        // that comes back newest-first is a list that ignored the request.
        List<HitSummary> summaries = new List<HitSummary>
        {
            Hit("newest-and-smallest", Noon, 1_000),
            Hit("middle", Noon.AddHours(-1), 50_000),
            Hit("oldest-and-biggest", Noon.AddDays(-30), 9_000_000),
        };

        MailService.SortForOrder(summaries, bySizeDescending: true);

        Assert.Equal(
            new[] { "oldest-and-biggest", "middle", "newest-and-smallest" },
            summaries.Select(h => h.Id).ToArray());
    }

    [Fact]
    public void ADateOrderedSearch_IsUnchanged()
    {
        // The default path, pinned so the fix above cannot quietly alter it.
        List<HitSummary> summaries = new List<HitSummary>
        {
            Hit("older", Noon.AddHours(-2), 9_000_000),
            Hit("newest", Noon, 10),
            Hit("middle", Noon.AddHours(-1), 500),
        };

        MailService.SortForOrder(summaries, bySizeDescending: false);

        Assert.Equal(new[] { "newest", "middle", "older" }, summaries.Select(h => h.Id).ToArray());
    }

    [Fact]
    public void AHitWithNoSize_SortsLastRatherThanFirst()
    {
        // A freshness-sweep hit can reach the merged list without a size. It is not a
        // zero-byte mail: treating an unknown as a zero would rank it against real mail on a
        // value nobody measured. Unknown sorts last, in both orders.
        List<HitSummary> summaries = new List<HitSummary>
        {
            Hit("unmeasured", Noon, null),
            Hit("small", Noon.AddHours(-1), 10),
            Hit("big", Noon.AddHours(-2), 90_000),
        };

        MailService.SortForOrder(summaries, bySizeDescending: true);

        Assert.Equal(new[] { "big", "small", "unmeasured" }, summaries.Select(h => h.Id).ToArray());
    }

    [Fact]
    public void EquallyLargeHits_FallBackToDate()
    {
        // List.Sort is not stable, so without a tiebreak two equally large mails could swap
        // places between runs of the same query - the kind of instability that makes a
        // result look wrong without ever being wrong.
        List<HitSummary> summaries = new List<HitSummary>
        {
            Hit("older-tie", Noon.AddDays(-1), 42_000),
            Hit("newer-tie", Noon, 42_000),
        };

        MailService.SortForOrder(summaries, bySizeDescending: true);

        Assert.Equal(new[] { "newer-tie", "older-tie" }, summaries.Select(h => h.Id).ToArray());
    }

    [Fact]
    public void AHitWithNoDate_SortsLastOnTheDatePath()
    {
        // The pre-existing rule, restated where it is now decided, so moving the comparison
        // into one method did not move this with it.
        List<HitSummary> summaries = new List<HitSummary>
        {
            Hit("undated", null, 10),
            Hit("dated", Noon.AddYears(-5), 10),
        };

        MailService.SortForOrder(summaries, bySizeDescending: false);

        Assert.Equal(new[] { "dated", "undated" }, summaries.Select(h => h.Id).ToArray());
    }

    private static HitSummary Hit(string id, DateTime? receivedUtc, long? sizeBytes)
    {
        return new HitSummary
        {
            Id = id,
            Subject = id,
            ReceivedUtc = receivedUtc,
            SizeBytes = sizeBytes,
        };
    }
}
