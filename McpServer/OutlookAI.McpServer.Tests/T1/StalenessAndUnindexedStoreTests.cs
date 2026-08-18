using System;
using System.Collections.Generic;
using System.Linq;

using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The two Q7 follow-ups the freshness work deliberately left undecided, answered
/// 2026-08-18.
/// <para>
/// (a) <c>staleness.newestIndexedUtc</c> is a profile-wide MAXIMUM on an unscoped search and
/// stays one, because narrowing it would make <c>search</c> and <c>outlook_health</c> report
/// different numbers for the same profile. A maximum cannot bound anyone else's lag, so it
/// gains a companion - <c>staleness.oldestStoreFrontierUtc</c> - which answers "how far
/// behind is the WORST store" without making the two tools disagree.
/// </para>
/// <para>
/// (b) The unindexed-store list was the one list in this server with no cap, in the payload
/// and in the advice sentence that joins the names into prose. Every other list here is
/// capped and every cap is reported; this was an omission, not a design choice.
/// </para>
/// </summary>
public sealed class StalenessAndUnindexedStoreTests
{
    private static readonly DateTime ProfileNewest = new DateTime(2026, 8, 18, 9, 30, 0, DateTimeKind.Utc);
    private static readonly DateTime QuietStoreFrontier = new DateTime(2026, 8, 17, 11, 15, 0, DateTimeKind.Utc);

    // ================================================================ (a) the worst store

    [Fact]
    public void UnscopedSearch_ReportsTheOldestPerStoreFrontier_NotTheProfileMaximum()
    {
        // The measured shape this exists for: two stores 11 min 19 s apart on this machine,
        // three stores 45.4 h apart the day before. The maximum says nothing about the one
        // that is behind, which is exactly the store whose recent mail is at risk.
        DateTime? oldest = MailService.OldestStoreFrontier(
            storeScoped: false, scopeFrontierUtc: ProfileNewest, oldestPerStoreFrontierUtc: QuietStoreFrontier);

        Assert.Equal(QuietStoreFrontier, oldest);
        Assert.NotEqual(ProfileNewest, oldest);
    }

    [Fact]
    public void StoreScopedSearch_ReportsThatStoresOwnFrontier_SoBothFieldsAgree()
    {
        // One store in scope means its frontier is both the newest and the oldest. Emitting
        // it rather than nothing is deliberate: a caller reading only the new field gets a
        // true answer on every search shape instead of having to know which shape it asked.
        DateTime? oldest = MailService.OldestStoreFrontier(
            storeScoped: true, scopeFrontierUtc: QuietStoreFrontier, oldestPerStoreFrontierUtc: null);

        Assert.Equal(QuietStoreFrontier, oldest);
    }

    [Fact]
    public void NoPerStoreFrontierMeasured_ReportsNothing_RatherThanTheProfileMaximum()
    {
        // An exhaustive search (no index scope by design) or an unscoped search whose store
        // catalog could not be read. Substituting the profile maximum would put a number in
        // the field that no store's index actually stands at - absence means "not measured",
        // which is a different and honest answer.
        Assert.Null(MailService.OldestStoreFrontier(
            storeScoped: false, scopeFrontierUtc: ProfileNewest, oldestPerStoreFrontierUtc: null));
        Assert.Null(MailService.OldestStoreFrontier(
            storeScoped: true, scopeFrontierUtc: null, oldestPerStoreFrontierUtc: null));
    }

    // ============================================================ (b) the uncapped list

    [Fact]
    public void UnindexedStoreCap_IsTheSameNumberAsTheSweptFolderListCap()
    {
        // Derived rather than a second 12: both bound a name list inside the same sweep
        // block for the same reason, and two independent numbers doing one job is how a
        // pair starts drifting.
        Assert.Equal(12, MailService.UnindexedStoreListCap);
        Assert.Equal(MailService.SweptFolderListCap, MailService.UnindexedStoreListCap);
    }

    [Fact]
    public void AListInsideTheCap_IsLeftAloneAndFlagsNothing()
    {
        SweepInfo sweep = SweepWithUnindexedStores(MailService.UnindexedStoreListCap);
        MailService.ApplyUnindexedStoreCap(sweep);

        Assert.Equal(MailService.UnindexedStoreListCap, sweep.StoresWithoutIndex!.Count);
        Assert.Null(sweep.StoresWithoutIndexTruncated);
        Assert.Null(sweep.StoresWithoutIndexTotal);
    }

    [Fact]
    public void AListOverTheCap_IsTruncated_AndSaysSoWithItsTrueCount()
    {
        SweepInfo sweep = SweepWithUnindexedStores(30);
        MailService.ApplyUnindexedStoreCap(sweep);

        Assert.Equal(MailService.UnindexedStoreListCap, sweep.StoresWithoutIndex!.Count);
        Assert.True(sweep.StoresWithoutIndexTruncated);
        Assert.Equal(30, sweep.StoresWithoutIndexTotal);

        // The kept names are the first ones found, in the order the sweep found them - a
        // stable prefix rather than an arbitrary sample.
        Assert.Equal("Archive 01.pst", sweep.StoresWithoutIndex[0]);
        Assert.Equal("Archive 12.pst", sweep.StoresWithoutIndex[MailService.UnindexedStoreListCap - 1]);
    }

    [Fact]
    public void NoUnindexedStores_IsUntouched()
    {
        SweepInfo sweep = new SweepInfo { Performed = true };
        MailService.ApplyUnindexedStoreCap(sweep);

        Assert.Null(sweep.StoresWithoutIndex);
        Assert.Null(sweep.StoresWithoutIndexTruncated);
        Assert.Null(sweep.StoresWithoutIndexTotal);
        Assert.Throws<ArgumentNullException>(() => MailService.ApplyUnindexedStoreCap(null!));
    }

    /// <summary>
    /// The prose half. The advice sentence NAMES the stores, so a cap applied only to the
    /// payload would leave the whole list in the text an agent relays to the user - and a
    /// sentence that names twelve of thirty and stops is a quieter lie than the unbounded
    /// list it replaced.
    /// </summary>
    [Fact]
    public void TheAdviceSentence_ReportsTheCapItIsQuoting()
    {
        SweepInfo sweep = SweepWithUnindexedStores(30);
        sweep.IndexFrontierMissing = true;
        MailService.ApplyUnindexedStoreCap(sweep);
        sweep.CoverageGaps = new[] { FreshMerge.GapNoIndexFrontier };

        string advice = Assert.Single(MailService.DescribeSweepCoverage(sweep, "12 minutes", folderScoped: false));

        Assert.Contains("Archive 01.pst", advice, StringComparison.Ordinal);
        Assert.Contains("18 more", advice, StringComparison.Ordinal);
        Assert.Contains("list capped at 12", advice, StringComparison.Ordinal);
        Assert.Contains("storesWithoutIndexTotal", advice, StringComparison.Ordinal);
        Assert.DoesNotContain("Archive 13.pst", advice, StringComparison.Ordinal);
    }

    [Fact]
    public void TheAdviceSentence_SaysNothingAboutACapItDidNotHit()
    {
        SweepInfo sweep = SweepWithUnindexedStores(2);
        sweep.IndexFrontierMissing = true;
        MailService.ApplyUnindexedStoreCap(sweep);
        sweep.CoverageGaps = new[] { FreshMerge.GapNoIndexFrontier };

        string advice = Assert.Single(MailService.DescribeSweepCoverage(sweep, "12 minutes", folderScoped: false));

        Assert.Contains("Archive 01.pst, Archive 02.pst", advice, StringComparison.Ordinal);
        Assert.DoesNotContain("more (list capped", advice, StringComparison.Ordinal);
    }

    private static SweepInfo SweepWithUnindexedStores(int count)
    {
        return new SweepInfo
        {
            Performed = true,
            StoresWithoutIndex = Enumerable
                .Range(1, count)
                .Select(i => "Archive " + i.ToString("00", System.Globalization.CultureInfo.InvariantCulture) + ".pst")
                .ToList(),
        };
    }
}
