using System.Text.RegularExpressions;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The hunt's variants (runbook section 4.15; the coordinator's direction of 2026-11-18, option (b)): what each makes
/// its derived drafts with, whether its Outlook has a window and where its main window looks - read from the two
/// variables the hunt's suite script sets, and nothing at all without them.
/// </summary>
public sealed class HuntPlanTests
{
    [Theory]
    [InlineData("ProductLoop", HuntMaker.Product, false, false)]
    [InlineData("OtherFolder", HuntMaker.Product, false, true)]
    [InlineData("NoWindow", HuntMaker.Product, true, false)]
    [InlineData("NoFormPage", HuntMaker.NoFormPage, false, false)]
    [InlineData("BareCreators", HuntMaker.Bare, false, false)]
    [InlineData("ColdStarts", HuntMaker.Product, true, false)]
    public void EveryVariant_VariesOneThingFromTheControl(string variant, HuntMaker maker, bool windowless, bool showsCalendar)
    {
        HuntPlan? plan = HuntPlan.Parse(variant, "10", out string? why);

        Assert.NotNull(plan);
        Assert.Null(why);
        Assert.Equal(variant, plan!.Variant);
        Assert.Equal(maker, plan.Maker);
        Assert.Equal(windowless, plan.Windowless);
        Assert.Equal(showsCalendar, plan.ShowsCalendar);
        Assert.Equal(10, plan.LegMinutes);
        Assert.Equal("S5-hunt-" + variant, plan.Cell);
        Assert.False(plan.ShowsEach);
    }

    [Theory]
    [InlineData(null, HuntAblation.None)]
    [InlineData("", HuntAblation.None)]
    [InlineData("new", HuntAblation.NewMail)]
    [InlineData("layout", HuntAblation.Layout)]
    [InlineData("signature", HuntAblation.Signature)]
    [InlineData("source", HuntAblation.Source)]
    public void TheTypedTextTests_LeaveOutTheStepTheHuntNames(string? value, HuntAblation expected)
    {
        Assert.Equal(expected, HuntPlan.ParseAblation(value, out string? why));
        Assert.Null(why);
    }

    [Theory]
    [InlineData("Layout")]
    [InlineData("newmail")]
    [InlineData("all")]
    public void AnUnknownAblation_IsRefused(string value)
    {
        Assert.Null(HuntPlan.ParseAblation(value, out string? why));
        Assert.Contains("must be new, layout, signature or source", why, StringComparison.Ordinal);
    }

    [Fact]
    public void PersonTyping_KeepsAWindowOpenWhileTheProductDrafts()
    {
        HuntPlan? plan = HuntPlan.Parse("PersonTyping", "10", null, out string? why);

        Assert.NotNull(plan);
        Assert.Null(why);
        Assert.Equal(HuntMaker.Product, plan!.Maker);
        Assert.True(plan.KeepsAWindowOpen);
        Assert.False(plan.ShowsEach);
        Assert.False(plan.Windowless);
        Assert.Equal("S5-hunt-PersonTyping", plan.Cell);
        Assert.True(HuntPlan.DraftsPerOpenWindow > 1);
        Assert.False(HuntPlan.Parse("ProductLoop", "10", null, out _)!.KeepsAWindowOpen);
    }

    [Theory]
    [InlineData(null, false)]
    [InlineData("", false)]
    [InlineData("0", false)]
    [InlineData("1", true)]
    public void TheProductsDrafts_AreRePromotedAfterwards_WhenAskedByOne(string? value, bool promotes)
    {
        HuntPlan? plan = HuntPlan.Parse("ProductLoop", "10", "typed", out _)!.WithPromoteAfter(value, out string? why);

        Assert.NotNull(plan);
        Assert.Null(why);
        Assert.Equal(promotes, plan!.PromotesAfter);
    }

    [Fact]
    public void PromoteAfter_IsRefusedWhereTheLoopSavesNothing_OrForAnythingButOneOrZero()
    {
        Assert.True(HuntPlan.Parse("NoFormPage", "10", null, out _)!.WithPromoteAfter("1", out _)!.PromotesAfter);
        Assert.Null(HuntPlan.Parse("BareCreators", "10", null, out _)!.WithPromoteAfter("1", out string? bare));
        Assert.Contains("makes none", bare, StringComparison.Ordinal);
        Assert.Null(HuntPlan.Parse("ProductLoop", "10", null, out _)!.WithPromoteAfter("yes", out string? odd));
        Assert.Contains("must be 1 or 0", odd, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null, false)]
    [InlineData("", false)]
    [InlineData("0", false)]
    [InlineData("1", true)]
    public void TheTypedWindowsDocument_IsReleasedAfterItsClose_WhenAskedByOne(string? value, bool late)
    {
        HuntPlan? plan = HuntPlan.Parse("ProductLoop", "10", "typed", out _)!.WithReleaseAfterClose(value, out string? why);

        Assert.NotNull(plan);
        Assert.Null(why);
        Assert.Equal(late, plan!.ReleasesAfterClose);
    }

    [Fact]
    public void ReleaseAfterClose_IsRefusedWhereNoWindowIsTypedInto_OrForAnythingButOneOrZero()
    {
        Assert.Null(HuntPlan.Parse("ProductLoop", "10", "1", out _)!.WithReleaseAfterClose("1", out string? shown));
        Assert.Contains("types into none", shown, StringComparison.Ordinal);
        Assert.Null(HuntPlan.Parse("ProductLoop", "10", "typed", out _)!.WithReleaseAfterClose("yes", out string? odd));
        Assert.Contains("must be 1 or 0", odd, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null, false)]
    [InlineData("", false)]
    [InlineData("0", false)]
    [InlineData("1", true)]
    public void TheProductsDrafts_AreDisplayedAndClosed_WhenAskedByOne(string? value, bool displays)
    {
        HuntPlan? plan = HuntPlan.Parse("ProductLoop", "10", null, out _)!.WithDisplay(value, out string? why);

        Assert.NotNull(plan);
        Assert.Null(why);
        Assert.Equal(displays, plan!.DisplaysEach);
    }

    [Fact]
    public void Display_IsRefusedWhereTheProductMakesNothing_OrForAnythingButOneOrZero()
    {
        Assert.Null(HuntPlan.Parse("NoFormPage", "10", null, out _)!.WithDisplay("1", out string? direct));
        Assert.Contains("makes none", direct, StringComparison.Ordinal);
        Assert.Null(HuntPlan.Parse("ProductLoop", "10", null, out _)!.WithDisplay("on", out string? odd));
        Assert.Contains("must be 1 or 0", odd, StringComparison.Ordinal);
    }

    [Fact]
    public void ColdStarts_MakesTheProductsDraftsInOutlooksItStartsWithNoWindow_AndNothingElse()
    {
        HuntPlan plan = HuntPlan.Parse("ColdStarts", "90", null, out string? why)!;

        Assert.Null(why);
        Assert.True(plan.ColdStarts);
        Assert.True(plan.Windowless);
        Assert.Equal(HuntMaker.Product, plan.Maker);
        Assert.Equal(2, HuntPlan.DraftsPerColdStart);
        Assert.Equal(100, HuntPlan.MostColdStarts);
        Assert.False(HuntPlan.Parse("ProductLoop", "90", null, out _)!.ColdStarts);
        Assert.Null(HuntPlan.Parse("ColdStarts", "90", "1", out string? shown));
        Assert.Contains("no window", shown, StringComparison.Ordinal);
        Assert.Null(HuntPlan.Parse("ColdStarts", "90", null, out _)!.WithPromoteAfter("1", out string? promote));
        Assert.Contains("own promotion", promote, StringComparison.Ordinal);
        Assert.Null(HuntPlan.Parse("ColdStarts", "90", null, out _)!.WithDisplay("1", out string? display));
        Assert.Contains("no window", display, StringComparison.Ordinal);
    }

    [Fact]
    public void TheCalendarStep_MovesTheMainWindowAndMakesNothing()
    {
        HuntPlan? plan = HuntPlan.Parse(HuntPlan.CalendarStep, "1", "0", out string? why);

        Assert.NotNull(plan);
        Assert.Null(why);
        Assert.Equal(HuntMaker.None, plan!.Maker);
        Assert.True(plan.ShowsCalendar);
        Assert.False(plan.Windowless);
        Assert.False(plan.ShowsEach);
        Assert.DoesNotContain(HuntPlan.CalendarStep, HuntPlan.Variants);
    }

    [Theory]
    [InlineData(null, false)]
    [InlineData("", false)]
    [InlineData("0", false)]
    [InlineData("1", true)]
    public void AWindowShownBeforeEachDraft_IsAskedForByOne(string? showEach, bool shows)
    {
        HuntPlan? plan = HuntPlan.Parse("NoFormPage", "10", showEach, out string? why);

        Assert.NotNull(plan);
        Assert.Null(why);
        Assert.Equal(shows, plan!.ShowsEach);
        Assert.Equal(HuntMaker.NoFormPage, plan.Maker);
        Assert.Equal(shows ? "S5-hunt-NoFormPage-shown" : "S5-hunt-NoFormPage", plan.Cell);
    }

    [Fact]
    public void ATypedWindowBeforeEachDraft_IsAskedForByTyped()
    {
        HuntPlan? plan = HuntPlan.Parse("ProductLoop", "10", "typed", out string? why);

        Assert.NotNull(plan);
        Assert.Null(why);
        Assert.True(plan!.ShowsEach);
        Assert.True(plan.TypesEach);
        Assert.Equal("S5-hunt-ProductLoop-typed", plan.Cell);
        Assert.False(HuntPlan.Parse("ProductLoop", "10", "1", out _)!.TypesEach);
    }

    [Theory]
    [InlineData("2")]
    [InlineData("yes")]
    [InlineData("true")]
    [InlineData("Typed")]
    public void AnythingElse_IsNoPlan(string showEach)
    {
        Assert.Null(HuntPlan.Parse("ProductLoop", "10", showEach, out string? why));
        Assert.Contains("must be 1, typed or 0", why, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null, "10", "is not set")]
    [InlineData("", "10", "is not set")]
    [InlineData("Legs", "10", "no hunting variant")]
    [InlineData("productloop", "10", "no hunting variant")]
    [InlineData("ProductLoop", null, "1 to 120")]
    [InlineData("ProductLoop", "0", "1 to 120")]
    [InlineData("ProductLoop", "121", "1 to 120")]
    [InlineData("ProductLoop", "-5", "1 to 120")]
    [InlineData("ProductLoop", "ten", "1 to 120")]
    public void WithoutAVariantAndItsMinutes_ThereIsNoPlan(string? variant, string? minutes, string reason)
    {
        Assert.Null(HuntPlan.Parse(variant, minutes, out string? why));
        Assert.Contains(reason, why, StringComparison.Ordinal);
    }

    [Fact]
    public void TheRunnersVariants_AreTheTestsAndItsOwnDefault()
    {
        // The runner's -HuntVariant: Legs, its default, runs the tests the crashes struck in; every other value is one
        // this test knows. A value the runner offered and the test did not know would be a hunting run that made nothing.
        string runner = File.ReadAllText(Path.Combine(RepoPaths.RepoRoot(), "Testbed", "host", "Invoke-LiveTierOnGuest.ps1"));
        Match set = Regex.Match(runner, @"\[ValidateSet\(([^\)]*)\)\]\s*\[string\]\s*\$HuntVariant = 'Legs'");
        Assert.True(set.Success, "the runner's -HuntVariant, with its ValidateSet and its default Legs, was not found");
        string[] offered = Regex.Matches(set.Groups[1].Value, "'([^']+)'").Select(m => m.Groups[1].Value).ToArray();

        Assert.Equal(new[] { "Legs", "LegsOtherFolder" }.Concat(HuntPlan.Variants).ToArray(), offered);
        Assert.Contains("'LegsOtherFolder' { return '" + HuntPlan.CalendarStep + "' }", runner, StringComparison.Ordinal);
        Assert.Contains("$HuntVariantVariable = '" + HuntPlan.VariantVariable + "'", runner, StringComparison.Ordinal);
        Assert.Contains("$HuntLegMinutesVariable = '" + HuntPlan.LegMinutesVariable + "'", runner, StringComparison.Ordinal);
        Assert.Contains("$HuntShowEachVariable = '" + HuntPlan.ShowEachVariable + "'", runner, StringComparison.Ordinal);
        Assert.Contains("$HuntAblateVariable = '" + HuntPlan.AblateVariable + "'", runner, StringComparison.Ordinal);
        Assert.Contains("$HuntPromoteAfterVariable = '" + HuntPlan.PromoteAfterVariable + "'", runner, StringComparison.Ordinal);
        Assert.Contains("$HuntReleaseAfterCloseVariable = '" + HuntPlan.ReleaseAfterCloseVariable + "'", runner, StringComparison.Ordinal);
        Assert.Contains("$HuntDisplayVariable = '" + HuntPlan.DisplayVariable + "'", runner, StringComparison.Ordinal);
        Assert.Contains("FullyQualifiedName~" + typeof(LiveHuntDerivedDraftTests).FullName, runner, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("OperationOutcomeException 0x80131509 The draft could not be created (COMException 0x800706BE), and NO DRAFT WAS SAVED", true)]
    [InlineData("COMException 0x800706BA The RPC server is unavailable.", true)]
    [InlineData("COMException 0x80010108 The object invoked has disconnected from its clients.", true)]
    [InlineData("COMException 0x80004005 The operation failed.", false)]
    [InlineData("TimeoutException 0x80131505 Q149: 'probe step' did not finish within 180 s", false)]
    [InlineData("", false)]
    [InlineData(null, false)]
    public void AFailureWaitsForOutlookToEnd_OnlyWhenItSaysOutlookIsGoing(string? failure, bool going)
    {
        Assert.Equal(going, HuntPlan.SaysOutlookIsGoing(failure));
    }

    [Theory]
    [InlineData(true, 0, false)]
    [InlineData(true, 4, false)]
    [InlineData(true, 5, true)]
    [InlineData(false, 0, true)]
    public void ALegStops_WhenOutlookEndsOrFiveDraftsInARowFail(bool alive, int failuresInARow, bool stops)
    {
        Assert.Equal(stops, HuntPlan.StopsAfter(alive, failuresInARow));
    }
}
