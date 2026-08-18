using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins what the tripwire census is allowed to cost, and - more importantly - pins that the
/// two ends of one comparison ask the same question. A folder walked before the run and
/// merely counted after it cannot be compared item by item, so the post-run pass repeats the
/// baseline's choices instead of re-deciding them against a budget that has since moved.
/// </summary>
public sealed class CensusIdentityPlanTests
{
    private static FolderCensus Walked(int items)
    {
        List<CensusItem> list = new(items);
        for (int i = 0; i < items; i++)
        {
            list.Add(new CensusItem("id-" + i, "fp-" + i, false));
        }

        return FolderCensus.WithItems(list);
    }

    [Fact]
    public void ABaselinePlan_WalksSmallOrdinaryFoldersAndCountsTheRest()
    {
        CensusIdentityPlan plan = CensusIdentityPlan.Baseline();

        Assert.True(plan.ShouldIdentify("Inbox", isVolatile: false, itemCount: 168));
        Assert.False(
            plan.ShouldIdentify("Archive", isVolatile: false, itemCount: CensusIdentityPlan.DefaultPerFolderLimit + 1));
    }

    [Fact]
    public void ABaselinePlan_NeverWalksASelfPruningFolder()
    {
        // Deleted Items is both the largest folder in most stores and the one place a
        // shrink proves nothing, so identity there would be the most expensive reading the
        // guard could take and the least useful.
        CensusIdentityPlan plan = CensusIdentityPlan.Baseline();

        Assert.False(plan.ShouldIdentify(StoreCountTripwire.VolatilePrefix + "Deleted Items", true, 12));
    }

    [Fact]
    public void ABaselinePlan_StopsWhenTheStoreBudgetIsSpent()
    {
        CensusIdentityPlan plan = CensusIdentityPlan.Baseline(perFolderLimit: 100, perStoreItemBudget: 150);

        Assert.True(plan.ShouldIdentify("A", false, 100));
        plan.Spend(100);
        Assert.True(plan.ShouldIdentify("B", false, 50));
        plan.Spend(50);
        Assert.False(plan.ShouldIdentify("C", false, 1));

        Assert.Equal(2, plan.IdentifiedFolders);
        Assert.Equal(150, plan.IdentifiedItems);
        Assert.Contains("150 item(s)", plan.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void ARepeatPlan_WalksExactlyWhatTheBaselineWalked()
    {
        Dictionary<string, FolderCensus> baseline = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Inbox"] = Walked(3),
            ["Archive"] = FolderCensus.CountOnly(6153),
        };

        CensusIdentityPlan plan = CensusIdentityPlan.Repeating(baseline);

        Assert.True(plan.ShouldIdentify("Inbox", false, 3));
        Assert.False(plan.ShouldIdentify("Archive", false, 6153));
        Assert.False(plan.ShouldIdentify("A folder that did not exist before", false, 1));
    }

    [Fact]
    public void ARepeatPlan_IgnoresTheBudgetButNotUnboundedGrowth()
    {
        // Comparability outranks cost: a folder that gained mail during the run must still
        // be walked. Past the growth headroom it degrades to a count, which is a weaker
        // reading rather than a wrong one.
        Dictionary<string, FolderCensus> baseline = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Inbox"] = Walked(2),
        };

        CensusIdentityPlan plan = CensusIdentityPlan.Repeating(baseline, perFolderLimit: 100);

        Assert.True(plan.ShouldIdentify("Inbox", false, 400));
        Assert.False(plan.ShouldIdentify("Inbox", false, 401));
    }

    [Fact]
    public void ACountOnlyPlan_WalksNothing()
    {
        CensusIdentityPlan plan = CensusIdentityPlan.CountOnly();

        Assert.False(plan.ShouldIdentify("Inbox", false, 1));

        // Including the empty ones. Walking an empty folder costs nothing, but claiming it
        // marks it as compared item by item, and this plan was told to take no such reading.
        // (A BASELINE plan does still claim empty folders once its budget is spent, which is
        // deliberate: the walk is free and an empty folder cannot lose anything.)
        Assert.False(plan.ShouldIdentify("Outbox", false, 0));
        Assert.Equal(0, plan.IdentifiedFolders);
        Assert.Equal(0, plan.IdentifiedItems);
    }
}
