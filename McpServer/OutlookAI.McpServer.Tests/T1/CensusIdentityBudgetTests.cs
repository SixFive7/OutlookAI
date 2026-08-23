using System.Reflection;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the census identity TIME budget and the rung above it.
/// <para>
/// The identity walk used to have no clock of its own: it shared the mailer's 3-minute STA
/// join with the folder-tree walk and everything else, so the only thing that could stop a
/// slow walk was the join killing the whole store's census and refusing the live tier. That
/// happened, on 2026-08-20. The walk now stops itself, degrades those folders to counts - the
/// same degradation the item budget and an unusable table already produce - and says so.
/// </para>
/// <para>
/// The value is a CEILING, not a target: 120 s against the only measurement that exists
/// (16.9 s for the whole 5-store pass), because one run is not a distribution and the risk is
/// asymmetric - too low kills a working operation, too high costs nothing when the work
/// finishes early. It is to be narrowed later from VM measurements.
/// </para>
/// </summary>
public sealed class CensusIdentityBudgetTests
{
    [Fact]
    public void TheIdentityTimeBudgetIsAHundredAndTwentySeconds()
    {
        // A literal, not the constant read back through itself: the number IS the decision.
        Assert.Equal(120_000, CensusIdentityPlan.DefaultIdentityTimeBudgetMs);
    }

    [Fact]
    public void TheSizeBudgetsDidNotMoveWhenTheClockArrived()
    {
        // 500/3,000/4x decide what the guard PROVES (which folders can be compared item by
        // item at all), and moving them is the maintainer's call rather than a side effect of
        // making the census affordable. The time budget was added beside them, not instead.
        Assert.Equal(500, CensusIdentityPlan.DefaultPerFolderLimit);
        Assert.Equal(3_000, CensusIdentityPlan.DefaultPerStoreItemBudget);
        Assert.Equal(4, CensusIdentityPlan.RepeatGrowthHeadroom);
    }

    [Fact]
    public void WhileTheBudgetHasTimeLeft_FoldersAreStillWalked()
    {
        CensusIdentityPlan plan = CensusIdentityPlan.WithClock(() => TimeSpan.FromMilliseconds(119_999));

        Assert.True(plan.ShouldIdentify("Inbox", isVolatile: false, itemCount: 168));
        Assert.Equal(0, plan.FoldersDeniedByClock);
        Assert.False(plan.IdentityClockExpired);
    }

    [Fact]
    public void OnceTheBudgetIsSpent_TheFolderIsCountedRatherThanWalked()
    {
        // Counted, never skipped: the count rule still guards every folder in the store, so
        // this is a weaker reading rather than a hole. An unmeasured mailbox would be a hole,
        // and that still refuses the tier.
        CensusIdentityPlan plan = CensusIdentityPlan.WithClock(() => TimeSpan.FromMilliseconds(120_000));

        Assert.False(plan.ShouldIdentify("Inbox", isVolatile: false, itemCount: 168));
        Assert.False(plan.ShouldIdentify("Postvak IN", isVolatile: false, itemCount: 52));
        Assert.Equal(2, plan.FoldersDeniedByClock);
        Assert.True(plan.IdentityClockExpired);
    }

    [Fact]
    public void ACensusThatRanOutOfTime_SaysSoInItsOwnLine()
    {
        CensusIdentityPlan plan = CensusIdentityPlan.WithClock(() => TimeSpan.FromMilliseconds(120_000));
        plan.NoteFolderMeasured();
        plan.ShouldIdentify("Inbox", false, 168);

        Assert.Contains("IDENTITY TIME BUDGET EXPIRED (120 s)", plan.Describe(), StringComparison.Ordinal);
        Assert.Contains("1 folder(s) counted instead of walked", plan.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void ACensusThatFinishedInsideItsBudget_SaysNothingAboutTheClock()
    {
        // The expiry clause is the loud half, so it must not be a line a reader has to notice
        // is ABSENT - the same rule the fell-back-to-counting clause follows.
        CensusIdentityPlan plan = CensusIdentityPlan.Baseline();
        plan.NoteFolderMeasured();
        plan.ShouldIdentify("Inbox", false, 168);
        plan.Spend(168);

        Assert.DoesNotContain("TIME BUDGET", plan.Describe(), StringComparison.Ordinal);
        Assert.Equal(0, plan.FoldersDeniedByClock);
    }

    [Fact]
    public void TheClockNeverDeniesAFolderThePlanWasNotGoingToWalkAnyway()
    {
        // The counter is a diagnostic and has to mean one thing: folders this plan WANTED and
        // could not afford. A folder above the per-folder limit, a self-pruning folder and a
        // count-only store were never candidates, and counting them here would bury the real
        // signal under folders nothing was ever going to walk.
        CensusIdentityPlan expired = CensusIdentityPlan.WithClock(() => TimeSpan.FromMilliseconds(600_000));

        Assert.False(expired.ShouldIdentify("Archive", isVolatile: false, itemCount: 6153));
        Assert.False(expired.ShouldIdentify(StoreCountTripwire.VolatilePrefix + "Deleted Items", true, 12));
        Assert.Equal(0, expired.FoldersDeniedByClock);
    }

    [Fact]
    public void TheRepeatPassIsBoundedByTheSameClock()
    {
        // The repeat pass runs at the END of a 27-minute tier run, so it is the one most
        // likely to meet a profile that has gone slow - and it is the one that would otherwise
        // sit in a fixture's Dispose with nothing to stop it but the STA join.
        Dictionary<string, FolderCensus> baseline = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Inbox"] = FolderCensus.WithItems(new[] { new CensusItem("id-1", "fp-1", false) }),
        };

        CensusIdentityPlan inTime = CensusIdentityPlan.RepeatingWithClock(baseline, () => TimeSpan.Zero);
        CensusIdentityPlan expired = CensusIdentityPlan.RepeatingWithClock(
            baseline, () => TimeSpan.FromMilliseconds(120_000));

        Assert.True(inTime.ShouldIdentify("Inbox", false, 1));
        Assert.False(expired.ShouldIdentify("Inbox", false, 1));
        Assert.Equal(1, expired.FoldersDeniedByClock);
    }

    [Fact]
    public void TheStaJoinAboveTheBudgetIsStrictlyLargerThanTheBudget()
    {
        // The rung this project has already got wrong once: an inner budget equal to the outer
        // deadline means the outer timer fires first and kills an operation that was working
        // fine inside its own budget. The census's join therefore carries the identity ceiling
        // PLUS the three minutes every other mailer operation gets, so the budget can expire
        // inside the call and leave the counting to finish underneath it.
        TimeSpan identity = TimeSpan.FromMilliseconds(CensusIdentityPlan.DefaultIdentityTimeBudgetMs);

        Assert.Equal(TimeSpan.FromMinutes(3), LiveOutlookTestMailer.DefaultStaBudget);
        Assert.True(
            LiveOutlookTestMailer.CensusStaBudget > identity,
            "the census STA join must outlive the identity budget it contains");
        Assert.Equal(identity + LiveOutlookTestMailer.DefaultStaBudget, LiveOutlookTestMailer.CensusStaBudget);
        Assert.Equal(300d, LiveOutlookTestMailer.CensusStaBudget.TotalSeconds);

        // And the headroom is the whole of the ordinary budget, not a token margin: what is
        // left after the walk gives up still has to enumerate the tree and count every folder.
        Assert.True(
            LiveOutlookTestMailer.CensusStaBudget - identity >= TimeSpan.FromMinutes(3),
            "the counting half of the census keeps the budget it always had");
    }

    [Fact]
    public void TheCensusReallyRunsUnderTheCensusJoin_ReadOutOfTheCompiledIl()
    {
        // The one line CI cannot execute: CaptureMailFolderCensus is COM from top to bottom.
        // Without this pin, dropping the argument and leaving the census on the ordinary
        // 3-minute join would pass every test in the suite while quietly making the 120 s
        // identity budget unreachable - the budget would never expire, because the join would
        // always fire first. Same technique LiveTierInventoryTests uses for the stdio tier:
        // read the compiled method rather than trusting the source to still say it.
        MethodInfo census = typeof(LiveOutlookTestMailer).GetMethod(
            nameof(LiveOutlookTestMailer.CaptureMailFolderCensus),
            BindingFlags.Public | BindingFlags.Static)!;

        Assert.Contains(
            StaticFieldsReadBy(census),
            field => field.Name == nameof(LiveOutlookTestMailer.CensusStaBudget)
                && field.DeclaringType == typeof(LiveOutlookTestMailer));
    }

    /// <summary>
    /// Every static field one method LOADS, resolved from its IL. <c>ldsfld</c> is 0x7E
    /// followed by a 4-byte metadata token whose high byte is 0x04 (the FieldDef table);
    /// requiring that byte and then resolving the token is what keeps a stray operand from
    /// being read as an instruction.
    /// </summary>
    private static List<FieldInfo> StaticFieldsReadBy(MethodInfo method)
    {
        byte[] il = method.GetMethodBody()!.GetILAsByteArray()!;
        List<FieldInfo> loaded = new();
        for (int i = 0; i + 4 < il.Length; i++)
        {
            if (il[i] != 0x7E || il[i + 4] != 0x04)
            {
                continue;
            }

            int token = il[i + 1] | (il[i + 2] << 8) | (il[i + 3] << 16) | (il[i + 4] << 24);
            try
            {
                FieldInfo? field = method.Module.ResolveField(token);
                if (field != null)
                {
                    loaded.Add(field);
                }
            }
            catch (ArgumentException)
            {
                // Not a real ldsfld - the bytes happened to look like one.
            }
        }

        return loaded;
    }
}
