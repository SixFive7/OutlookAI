using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The freshness sweep runs out of time GRACEFULLY, and says so.
/// <para>
/// THE BEHAVIOUR THIS REPLACES, observed on the maintainer's real profile. The whole-profile
/// sweep had no budget of its own - only the gateway deadline above it - so a sweep that ran
/// long produced a <c>TimeoutException</c>, the supervisor concluded the COM host was wedged,
/// killed and replaced it, and every folder the sweep HAD already covered was thrown away.
/// A big mailbox therefore lost its freshness tier entirely rather than getting a partial
/// one, and it cost a host restart and two strikes toward the breaker each time.
/// </para>
/// <para>
/// It now stops at the next store or folder boundary, returns what it covered, and reports
/// the shortfall as a coverage gap - the same discipline the exhaustive scan has had since
/// 2026-08-18. Two halves are pinned here: the BOUNDARY the COM walk stops on (pure, because
/// the walk itself only exists over live folders), and the REPORTING chain from the COM
/// result through the payload to the sentence an agent reads.
/// </para>
/// </summary>
public sealed class SweepBudgetExpiryTests
{
    [Theory]
    // Inside the budget, and exactly AT it: not spent. The comparison matches
    // ExhaustiveScanState.CheckDeadline and DecideSweepWalk, so "budget" means one thing.
    [InlineData(165_000, 0L, false)]
    [InlineData(165_000, 165_000L, false)]
    [InlineData(165_000, 165_001L, true)]
    // No budget means unbounded. "No budget" must never become "no folders", which is what
    // an in-process caller carrying its own bound relies on.
    [InlineData(0, 999_999L, false)]
    [InlineData(-1, 999_999L, false)]
    public void TheSweepBudget_IsSpentOnlyOnceElapsedHasPassedIt(int budgetMs, long elapsedMs, bool expected)
    {
        Assert.Equal(expected, OutlookComSession.SweepBudgetSpent(budgetMs, elapsedMs));
    }

    [Fact]
    public void TheInnerSweepBudget_IsDerivedFromTheOuterOne_WithTheReturnTripReserved()
    {
        // Equality is the defect the exhaustive scan already paid for: an inner budget equal
        // to its outer one can never degrade, because the watchdog fires while the walk is
        // still serializing its answer.
        Assert.Equal(MailService.SweepBudgetMs - ComOperationBudgets.ResultReturnHeadroomMs, MailService.SweepWorkBudgetMs);
        Assert.True(
            MailService.SweepWorkBudgetMs < MailService.SweepBudgetMs,
            $"the sweep's inner budget ({MailService.SweepWorkBudgetMs} ms) must be strictly inside the gateway budget it "
            + $"runs under ({MailService.SweepBudgetMs} ms)");
        Assert.True(MailService.SweepWorkBudgetMs > 0);
    }

    /// <summary>
    /// One store's sweep, MEASURED 2026-08-19 with the per-folder cap engaged: 13.6 s,
    /// 11.8 s, 10.7 s and 11.9 s over four runs against a purpose-built corpus (one PST
    /// outside the local index, 20 000 items across the four arrival-path folders, 1 612 of
    /// them inside the seven-day fallback window). Rounded up to 12 s.
    /// </summary>
    private const int MeasuredSweepPerStoreMs = 12_000;

    /// <summary>
    /// Stores on the profile this budget has to serve. The maintainer's mounts five, and the
    /// sweep covers four arrival-path folders in each - so the measured extrapolation is
    /// ~60 s, against a budget that used to be 30 s. That is the direct explanation for the
    /// sweep timeout observed there, which then killed and replaced the COM host.
    /// </summary>
    private const int MeasuredProfileStores = 5;

    /// <summary>
    /// Headroom demanded over that extrapolation. Three times, and it is headroom rather
    /// than luxury: the corpus is a fast LOCAL PST and the same per-item work against
    /// Exchange is slower.
    /// </summary>
    private const int MeasuredSweepHeadroomFactor = 3;

    /// <summary>
    /// A FLOOR, and since 2026-08-24 only a floor - read the three constants above with that
    /// in mind. The ~12 s-per-store figure they encode was measured while the sweep's sort
    /// was silently failing (fixed in <c>bea7fc9</c>), so it describes broken behaviour doing
    /// different work and no longer SIZES anything: the budget is now a maintainer-set
    /// ceiling of 600 s awaiting a re-measurement with the sort working. What this still
    /// buys is the direction that has actually gone wrong twice - a budget quietly narrowed
    /// back under the only whole-profile number anybody has ever taken fails here, and says
    /// which number it fell under.
    /// </summary>
    [Fact]
    public void TheSweepBudget_HoldsTheMeasuredWholeProfileSweep_WithHeadroom()
    {
        int extrapolated = MeasuredSweepPerStoreMs * MeasuredProfileStores;

        Assert.True(
            MailService.SweepBudgetMs >= MeasuredSweepHeadroomFactor * extrapolated,
            $"the sweep budget ({MailService.SweepBudgetMs} ms) must hold at least {MeasuredSweepHeadroomFactor}x the "
            + $"measured whole-profile sweep ({extrapolated} ms = {MeasuredProfileStores} stores x "
            + $"{MeasuredSweepPerStoreMs} ms). Below that the budget expires during ordinary use, which is not a bound - "
            + "it is a coverage hole on every search, and before 2026-08-19 it was a killed COM host as well.");

        // And the INNER budget - the one the walk actually stops on - still clears the
        // measurement itself with room, so an ordinary whole-profile sweep finishes rather
        // than reporting partial coverage every time.
        Assert.True(
            MailService.SweepWorkBudgetMs > extrapolated,
            $"the sweep's own soft budget ({MailService.SweepWorkBudgetMs} ms) must exceed the measured whole-profile "
            + $"sweep ({extrapolated} ms), or every search reports a coverage gap");
    }

    [Fact]
    public void AnExpiredSweepBudget_ReachesThePayloadAsAPartialAnswerRatherThanAFailure()
    {
        SweepInfo info = new SweepInfo { Performed = true };
        MailService.ApplySweepCounters(info, ExpiredSweep(), store: null);

        // The fact itself, on the block an agent reads.
        Assert.True(info.SweepBudgetExpired);

        // And the folders it DID cover are still reported - that is the whole point of
        // stopping rather than failing.
        Assert.Equal(6, info.FoldersSwept);

        IReadOnlyList<string> gaps = FreshMerge.DescribeCoverageGaps(info)!;
        Assert.Contains(FreshMerge.GapSweepBudget, gaps);

        // Machine-readable pair, not just prose: an agent reading fields must be told the
        // answer is partial.
        Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyFreshness(info));

        // A different hole from the subtree walk's own 2 s bound, because the remedies point
        // in different directions.
        Assert.DoesNotContain(FreshMerge.GapTimeBudget, gaps);
    }

    [Fact]
    public void AStoreScopedRequest_IsStillToldTheSweepRanOutOfTime()
    {
        // Deliberately NOT attributed per store. The budget is spent across every store the
        // sweep visited, and the stores it never reached are exactly the ones with no entry
        // to attribute it to - so a store-scoped request served by a cached all-stores sweep
        // would otherwise read its own zero coverage as "nothing was there" rather than as
        // "we ran out of time before reaching you".
        SweepInfo info = new SweepInfo { Performed = true };
        MailService.ApplySweepCounters(info, ExpiredSweep(), store: "unvisited-store");

        Assert.True(info.SweepBudgetExpired);
        Assert.Equal(0, info.FoldersSwept);
        Assert.Contains(FreshMerge.GapSweepBudget, FreshMerge.DescribeCoverageGaps(info)!);
    }

    [Fact]
    public void TheAdviceQuotesTheBudgetThatActuallyStoppedIt_AndDoesNotSayTheHostFailed()
    {
        SweepInfo info = new SweepInfo
        {
            Performed = true,
            FoldersSwept = 6,
            SweepBudgetExpired = true,
            CoverageGaps = new[] { FreshMerge.GapSweepBudget },
        };

        string line = Assert.Single(MailService.DescribeSweepCoverage(info, "12 minutes", folderScoped: false));

        // Derived from the constant, so the sentence cannot quote a budget the sweep no
        // longer runs under.
        Assert.Contains(
            (MailService.SweepWorkBudgetMs / 1000).ToString(CultureInfo.InvariantCulture) + " s",
            line,
            StringComparison.Ordinal);

        // The remedy is about SCOPE, not about retrying a broken host - and it must not
        // borrow the subtree walk's advice, which cannot be acted on over a default-folder
        // sweep that is shallow by construction.
        Assert.Contains("store", line, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("include_subfolders", line, StringComparison.Ordinal);
    }

    /// <summary>A sweep that covered six folders and then ran out of its own budget.</summary>
    private static ComSweepResult ExpiredSweep()
    {
        return new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 6,
            foldersSkipped: 0,
            sweptFolders: Enumerable.Range(1, 6).Select(i => "store-a/Folder" + i).ToArray(),
            perStore: new[]
            {
                new ComStoreSweepCounters("store-a", foldersSwept: 6, foldersSkipped: 0, foldersFailed: 0, foldersAbsent: 0),
            },
            sweepBudgetExpired: true);
    }
}
