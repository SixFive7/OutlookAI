using System.Reflection;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the census's account of WHY a folder was counted rather than walked, and the verdict it
/// may honestly reach over a pair of readings of unequal strength.
/// <para>
/// <b>The defect, found by reading during the 2026-08-24 tripwire investigation.</b> Whether a
/// folder is compared by identity or by count is decided independently on each pass, and the
/// post-run decision is timing-dependent - the identity clock, a folder that grew past the repeat
/// headroom, or one transient COM failure silently switches it from <c>EvaluateByIdentity</c> to
/// <c>EvaluateByCount</c>. The count rule cannot exonerate a filing and cannot see a departure
/// masked by an arrival, so the SAME mailbox state yields <c>note: filed (not loss)</c> on one
/// reading and <c>ITEMS LOST</c> on the next. That is a census reporting differently twice with
/// nothing having changed, and it is exactly what a re-census 30 s later clears - which is how it
/// has been hiding behind the retry ladder rather than being fixed.
/// </para>
/// <para>
/// <b>It also lied about itself.</b> <c>EvaluateByCount</c>'s failure text said
/// <c>(folder above the identity budget)</c> unconditionally. That is one of six possible causes
/// and the wrong one whenever the cause was the clock, an unusable table, a self-pruning folder, a
/// count-only plan or a repeat pass with no baseline reading to match. These messages are read
/// exactly once, in an emergency, by somebody who believes mail has just been deleted; naming the
/// wrong cause sends them to the wrong remedy.
/// </para>
/// <para>
/// <b>What is fixed and what is not.</b> This is option (2) of the three the TODO records - carry
/// the reason, and say which reading was the weak one. Option (1), re-walking a degraded folder
/// before concluding anything, is the one that removes the false failure at its source and needs a
/// COM call, so it cannot land without a live run. Nothing here changes which runs FAIL: the same
/// deltas fail as before, with a verdict that now says what it is made of.
/// </para>
/// <para>
/// Synthetic store, folder and item names. No COM, no Outlook, no mailbox, no settings file.
/// </para>
/// </summary>
public sealed class CensusReadingStrengthTests
{
    private const string Hub = "hub@example.test";
    private const string Watched = "watched@example.test";

    // ------------------------------------------------------------------ the reason is carried

    [Fact]
    public void AWalkedFolderSaysItWasWalked()
    {
        FolderCensus walked = FolderCensus.WithItems(new[] { Item("a"), Item("b") });

        Assert.True(walked.HasIdentities);
        Assert.Equal(CensusCountReason.Walked, walked.CountReason);
    }

    [Fact]
    public void ACountedFolderCannotClaimItWasWalked()
    {
        // Left constructible, "reason not recorded" becomes the commonest answer within a month
        // and the verdict is back to guessing. Refused at the factory instead.
        Assert.Throws<ArgumentOutOfRangeException>(
            () => FolderCensus.CountOnly(9, CensusCountReason.Walked));
    }

    [Theory]
    [InlineData(CensusCountReason.PlanIsCountOnly)]
    [InlineData(CensusCountReason.SelfPruningFolder)]
    [InlineData(CensusCountReason.AbovePerFolderLimit)]
    [InlineData(CensusCountReason.StoreItemBudgetSpent)]
    [InlineData(CensusCountReason.NotIdentifiedAtBaseline)]
    [InlineData(CensusCountReason.IdentityClockExpired)]
    [InlineData(CensusCountReason.TableUnusable)]
    public void EveryReasonSurvivesOntoTheFolderAndHasWordsOfItsOwn(CensusCountReason reason)
    {
        FolderCensus counted = FolderCensus.CountOnly(9, reason);

        Assert.False(counted.HasIdentities);
        Assert.Equal(reason, counted.CountReason);

        // Words of its own, not a shared placeholder: a message that described two causes
        // identically would be the defect this closes, one indirection further along.
        string described = CensusReadingStrength.Describe(reason);
        Assert.False(string.IsNullOrWhiteSpace(described));
        Assert.DoesNotContain("not recorded", described, StringComparison.Ordinal);
        Assert.Equal(
            1,
            Enum.GetValues<CensusCountReason>()
                .Count(r => string.Equals(CensusReadingStrength.Describe(r), described, StringComparison.Ordinal)));
    }

    // ------------------------------------------------------------------ the plan's own answer

    [Fact]
    public void ACountOnlyPlanSaysSoRatherThanBlamingTheFolderSize()
    {
        CensusIdentityPlan plan = CensusIdentityPlan.CountOnly();

        Assert.False(plan.TryIdentify("Inbox", isVolatile: false, itemCount: 1, out CensusCountReason reason));
        Assert.Equal(CensusCountReason.PlanIsCountOnly, reason);
    }

    [Fact]
    public void ASelfPruningFolderIsNotBlamedOnTheBudgetEither()
    {
        CensusIdentityPlan plan = CensusIdentityPlan.Baseline();

        Assert.False(plan.TryIdentify(
            StoreCountTripwire.VolatilePrefix + "Deleted Items", isVolatile: true, itemCount: 3,
            out CensusCountReason reason));
        Assert.Equal(CensusCountReason.SelfPruningFolder, reason);
    }

    [Fact]
    public void AFolderAboveThePerFolderLimitIsTheOneCaseTheOldTextWasRIGHTAbout()
    {
        CensusIdentityPlan plan = CensusIdentityPlan.Baseline();

        Assert.False(plan.TryIdentify(
            "Archive", isVolatile: false, itemCount: CensusIdentityPlan.DefaultPerFolderLimit + 1,
            out CensusCountReason reason));
        Assert.Equal(CensusCountReason.AbovePerFolderLimit, reason);
    }

    [Fact]
    public void TheSTOREBudgetRunningOutIsADifferentAnswerFromTheFolderBeingTooBig()
    {
        // They point at different numbers - DefaultPerStoreItemBudget against
        // DefaultPerFolderLimit - so reporting one for the other sends a reader to the wrong
        // constant. Per-folder is asked FIRST, deliberately: a folder too big on its own is too
        // big whatever the store has left.
        CensusIdentityPlan plan = CensusIdentityPlan.Baseline(perFolderLimit: 500, perStoreItemBudget: 600);
        Assert.True(plan.TryIdentify("Inbox", false, 400, out _));
        plan.Spend(400);

        Assert.False(plan.TryIdentify("Sent Items", false, 300, out CensusCountReason reason));
        Assert.Equal(CensusCountReason.StoreItemBudgetSpent, reason);

        Assert.False(plan.TryIdentify("Archive", false, 900, out CensusCountReason tooBig));
        Assert.Equal(CensusCountReason.AbovePerFolderLimit, tooBig);
    }

    [Fact]
    public void ARepeatPassNamesTheBaselineRatherThanABudget()
    {
        Dictionary<string, FolderCensus> baseline = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Inbox"] = FolderCensus.WithItems(new[] { Item("a") }),
            ["Archive"] = FolderCensus.CountOnly(6153, CensusCountReason.AbovePerFolderLimit),
        };
        CensusIdentityPlan plan = CensusIdentityPlan.Repeating(baseline);

        Assert.True(plan.TryIdentify("Inbox", false, 1, out _));

        Assert.False(plan.TryIdentify("Archive", false, 6153, out CensusCountReason reason));
        Assert.Equal(CensusCountReason.NotIdentifiedAtBaseline, reason);

        // And a folder the baseline DID identify, grown past the repeat headroom, is a size
        // answer again rather than a baseline one.
        Assert.False(plan.TryIdentify(
            "Inbox", false, (CensusIdentityPlan.DefaultPerFolderLimit * CensusIdentityPlan.RepeatGrowthHeadroom) + 1,
            out CensusCountReason grown));
        Assert.Equal(CensusCountReason.AbovePerFolderLimit, grown);
    }

    [Fact]
    public void TheCLOCKIsNamedAsTheClockAndCountedSeparately()
    {
        // The case the old text got wrong most expensively: a folder well inside every size
        // budget, counted because the census had gone on too long, reported as "above the
        // identity budget" - which is a sentence about a number nobody needs to change.
        TimeSpan elapsed = TimeSpan.Zero;
        CensusIdentityPlan plan = CensusIdentityPlan.WithClock(() => elapsed, identityTimeBudgetMs: 1_000);

        Assert.True(plan.TryIdentify("Inbox", false, 3, out CensusCountReason walked));
        Assert.Equal(CensusCountReason.Walked, walked);

        elapsed = TimeSpan.FromMilliseconds(1_000);
        Assert.False(plan.TryIdentify("Sent Items", false, 3, out CensusCountReason reason));
        Assert.Equal(CensusCountReason.IdentityClockExpired, reason);
        Assert.Equal(1, plan.FoldersDeniedByClock);
        Assert.True(plan.IdentityClockExpired);
    }

    [Fact]
    public void ShouldIdentifyStillAnswersExactlyWhatTryIdentifyDoes()
    {
        // The old entry point is now the new one with the reason dropped, so the two cannot
        // disagree - which they would within one edit if both held their own copy of the rules.
        foreach ((int limit, int budget, bool isVolatile, int count) in new[]
                 {
                     (500, 3_000, false, 10),
                     (500, 3_000, true, 10),
                     (500, 3_000, false, 501),
                     (0, 0, false, 0),
                 })
        {
            CensusIdentityPlan a = CensusIdentityPlan.Baseline(limit, budget);
            CensusIdentityPlan b = CensusIdentityPlan.Baseline(limit, budget);
            Assert.Equal(
                a.ShouldIdentify("Inbox", isVolatile, count),
                b.TryIdentify("Inbox", isVolatile, count, out CensusCountReason reason));
            Assert.Equal(b.ShouldIdentify("Inbox", isVolatile, count), reason == CensusCountReason.Walked);
        }
    }

    // ------------------------------------------------------------------ strength of a PAIR

    [Fact]
    public void DegradedIsExactlyIdentifiedThenCounted()
    {
        FolderCensus identified = FolderCensus.WithItems(new[] { Item("a") });
        FolderCensus counted = FolderCensus.CountOnly(1, CensusCountReason.IdentityClockExpired);

        Assert.True(CensusReadingStrength.Degraded(identified, counted));
        Assert.False(CensusReadingStrength.Degraded(counted, identified));
        Assert.False(CensusReadingStrength.Degraded(identified, identified));
        Assert.False(CensusReadingStrength.Degraded(counted, counted));
    }

    [Fact]
    public void ADegradedPairSaysTheAFTERReadingWasTheWeakOne_AndWhy()
    {
        string explained = CensusReadingStrength.Explain(
            FolderCensus.WithItems(new[] { Item("a") }),
            FolderCensus.CountOnly(1, CensusCountReason.TableUnusable));

        Assert.Contains("POST-RUN reading", explained, StringComparison.Ordinal);
        Assert.Contains("WEAKER", explained, StringComparison.Ordinal);
        Assert.Contains(
            CensusReadingStrength.Describe(CensusCountReason.TableUnusable), explained, StringComparison.Ordinal);

        // The sentence that makes this actionable rather than merely honest.
        Assert.Contains("re-read this folder", explained, StringComparison.Ordinal);
        Assert.DoesNotContain("BASELINE reading was the weaker", explained, StringComparison.Ordinal);
    }

    [Fact]
    public void AnUPGRADEDPairSaysTheBaselineWasTheWeakOneAndDoesNotAskForARereRead()
    {
        // Nothing can be done about a weak baseline now, so the advice would be noise. It is
        // still SAID, because "both readings were counts" would be a lie about this pair.
        string explained = CensusReadingStrength.Explain(
            FolderCensus.CountOnly(1, CensusCountReason.StoreItemBudgetSpent),
            FolderCensus.WithItems(new[] { Item("a") }));

        Assert.Contains("BASELINE reading was the weaker", explained, StringComparison.Ordinal);
        Assert.DoesNotContain("re-read this folder", explained, StringComparison.Ordinal);
    }

    [Fact]
    public void TwoCOUNTSWithTheSameReasonSayItOnce_AndWithDifferentReasonsSayBoth()
    {
        Assert.Contains(
            "Both readings were counts (" + CensusReadingStrength.Describe(CensusCountReason.AbovePerFolderLimit) + ")",
            CensusReadingStrength.Explain(
                FolderCensus.CountOnly(9, CensusCountReason.AbovePerFolderLimit),
                FolderCensus.CountOnly(8, CensusCountReason.AbovePerFolderLimit)),
            StringComparison.Ordinal);

        string mixed = CensusReadingStrength.Explain(
            FolderCensus.CountOnly(9, CensusCountReason.AbovePerFolderLimit),
            FolderCensus.CountOnly(8, CensusCountReason.IdentityClockExpired));
        Assert.Contains("baseline: ", mixed, StringComparison.Ordinal);
        Assert.Contains("after the run: ", mixed, StringComparison.Ordinal);
        Assert.Contains(
            CensusReadingStrength.Describe(CensusCountReason.IdentityClockExpired), mixed, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ what a run actually prints

    [Fact]
    public void TheLostItemsVerdictNoLongerBlamesABudgetItCannotKnowWasTheCause()
    {
        // End to end through the real Evaluate, which is where the wrong sentence was printed.
        TripwireVerdict verdict = StoreCountTripwire.Evaluate(
            Census((Watched, "Inbox", FolderCensus.WithItems(new[] { Item("a"), Item("b"), Item("c") }))),
            Census((Watched, "Inbox", FolderCensus.CountOnly(1, CensusCountReason.IdentityClockExpired))),
            Hub);

        string failure = Assert.Single(verdict.Failures);
        Assert.Contains("ITEMS LOST", failure, StringComparison.Ordinal);
        Assert.Contains("WEAKER", failure, StringComparison.Ordinal);
        Assert.Contains(
            CensusReadingStrength.Describe(CensusCountReason.IdentityClockExpired), failure, StringComparison.Ordinal);

        // The exact sentence that used to be printed for every cause.
        Assert.DoesNotContain("folder above the identity budget", failure, StringComparison.Ordinal);
    }

    [Fact]
    public void ADegradedPairWithNOChangeInCountIsNOTED_BecauseTheGuardGotWeakerAndNothingSaidSo()
    {
        // The quietest way this guard loses its teeth: the baseline could tell a filing from a
        // deletion in this folder and the post-run pass cannot, so one item removed while another
        // arrived is now invisible here. Never a failure - nothing was observed to leave.
        TripwireVerdict verdict = StoreCountTripwire.Evaluate(
            Census((Watched, "Inbox", FolderCensus.WithItems(new[] { Item("a"), Item("b") }))),
            Census((Watched, "Inbox", FolderCensus.CountOnly(2, CensusCountReason.TableUnusable))),
            Hub);

        Assert.False(verdict.Failed);
        Assert.Contains(verdict.Notes, n => n.Contains("weaker reading", StringComparison.Ordinal));
    }

    [Fact]
    public void TheHubIsStillExemptFromAllOfIt()
    {
        // The hub is where the suite writes; its churn is tagged and the zero-artifact sweep
        // polices it. A weaker-reading note there would be noise on every single run.
        TripwireVerdict verdict = StoreCountTripwire.Evaluate(
            Census((Hub, "Inbox", FolderCensus.WithItems(new[] { Item("a"), Item("b") }))),
            Census((Hub, "Inbox", FolderCensus.CountOnly(2, CensusCountReason.TableUnusable))),
            Hub);

        Assert.False(verdict.Failed);
        Assert.DoesNotContain(verdict.Notes, n => n.Contains("weaker reading", StringComparison.Ordinal));
    }

    [Fact]
    public void AnEqualPairOfCountsStillSaysNothingAtAll()
    {
        // Regression guard on the note above: the census walks every mail folder of every store,
        // and most of them do not move. A line per unchanged folder would bury the ones that did.
        TripwireVerdict verdict = StoreCountTripwire.Evaluate(
            Census((Watched, "Inbox", FolderCensus.CountOnly(2, CensusCountReason.AbovePerFolderLimit))),
            Census((Watched, "Inbox", FolderCensus.CountOnly(2, CensusCountReason.AbovePerFolderLimit))),
            Hub);

        Assert.False(verdict.Failed);
        Assert.Empty(verdict.Notes);
    }

    // ------------------------------------------------------------------ the census call site

    [Fact]
    public void TheLiveCensusStillRecordsTheReasonItWasGiven()
    {
        // The one thing these pure tests cannot pin: that the COM census still asks. CaptureFolder
        // is driven by a live Outlook folder, so no test on a runner can enter it - and dropping
        // the reason there would restore the guess while leaving every assertion above green.
        string source = LiveSource("LiveOutlookTestMailer.cs");

        Assert.Contains(
            "plan.TryIdentify(key, isVolatile, count, out CensusCountReason refused)",
            source, StringComparison.Ordinal);
        Assert.Contains("FolderCensus.CountOnly(count, refused)", source, StringComparison.Ordinal);

        // The walk-failure path is the one the plan cannot know about: the plan SAID walk, and the
        // table then would not answer. Reporting the plan's reason there would say "walked".
        Assert.Contains(
            "FolderCensus.CountOnly(count, CensusCountReason.TableUnusable)", source, StringComparison.Ordinal);
        Assert.DoesNotContain("plan.ShouldIdentify(", source, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ helpers

    private static CensusItem Item(string id)
    {
        return new CensusItem(id, "fp-" + id, tagged: false);
    }

    private static Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> Census(
        params (string Store, string Folder, FolderCensus Census)[] entries)
    {
        Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> census =
            new(StringComparer.OrdinalIgnoreCase);
        foreach ((string store, string folder, FolderCensus folderCensus) in entries)
        {
            if (!census.TryGetValue(store, out IReadOnlyDictionary<string, FolderCensus>? byFolder))
            {
                byFolder = new Dictionary<string, FolderCensus>(StringComparer.OrdinalIgnoreCase);
                census[store] = byFolder;
            }

            ((Dictionary<string, FolderCensus>)byFolder)[folder] = folderCensus;
        }

        return census;
    }

    private static string LiveSource(string fileName)
    {
        string dir = typeof(LiveTestSettings).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
        string path = Path.Combine(dir, "T2", fileName);
        Assert.True(File.Exists(path), "live census source is missing: " + path);
        return File.ReadAllText(path);
    }
}
