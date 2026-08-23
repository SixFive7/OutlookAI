using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Pins the three pure decisions that keep a synthetic measurement corpus honest over time:
/// whether it can still answer the windows it was cut for (<see cref="CorpusFreshness"/>),
/// how it is moved forward without being regenerated (<see cref="CorpusReanchor"/>), and
/// whether what is in the store is what the plan describes (<see cref="CorpusCensus"/>).
/// <para>
/// <b>These are NOT live tests</b> - no Outlook, no COM, no mailbox, no settings file - and
/// they carry no <c>Category=Live</c> trait, so they run in the ordinary CI tier. They sit in
/// T2 rather than beside <c>CorpusGeneratorTests</c> in T1 only because T1 was owned by
/// another worktree while they were written; the natural home is T1 and moving them there is
/// a rename with no other consequence.
/// </para>
/// </summary>
public class CorpusFreshnessTests
{
    private static readonly DateTime Anchor = new(2026, 8, 1, 0, 0, 0, DateTimeKind.Utc);

    private static CorpusPlan Plan(long seed = 4242, string id = "vm1")
        => new(new CorpusPlanOptions(id, seed, Anchor));

    /// <summary>A manifest holding every ordinal, each dated as if built with <paramref name="shift"/> applied.</summary>
    private static CorpusManifest ManifestFor(CorpusPlan plan, int count, TimeSpan shift, int undatedEvery = 0)
    {
        CorpusManifest manifest = CorpusManifest.Create(new CorpusManifestHeader(
            CorpusManifest.CurrentVersion,
            plan.Options.CorpusId,
            plan.Options.Seed,
            CorpusManifest.FormatUtc(plan.Options.AnchorUtc),
            plan.Options.ShapeKey,
            "Synthetic Data File",
            null,
            "PropertyAccessor",
            "InPlaceWithSentFlag"));
        for (int ordinal = 1; ordinal <= count; ordinal++)
        {
            CorpusItemSpec spec = plan.Describe(ordinal);
            bool undated = undatedEvery > 0 && ordinal % undatedEvery == 0;
            manifest.Add(new CorpusManifestItem(
                ordinal,
                "SYNTH" + ordinal.ToString("D8", System.Globalization.CultureInfo.InvariantCulture),
                spec.FolderId,
                spec.BodyBytes,
                undated ? null : CorpusManifest.FormatUtc(spec.ReceivedUtc + shift)));
        }

        return manifest;
    }

    // ------------------------------------------------------------------ freshness

    [Fact]
    public void Freshness_IsFreshWhileTheCorpusSitsOnItsAnchor()
    {
        CorpusFreshnessReport report =
            CorpusFreshness.Evaluate(Plan(), 4_000, TimeSpan.Zero, Anchor.AddMinutes(30));
        Assert.Equal(CorpusFreshnessVerdict.Fresh, report.Verdict);
        Assert.True(CorpusFreshness.Decide(report).Proceed);
        Assert.All(report.Windows, w => Assert.True(w.StillInWindow > 0));
    }

    [Fact]
    public void Freshness_SixWeeksOnTheSevenDayWindowIsEmptyAndTheCheckFails()
    {
        // The whole reason this exists. At six weeks the corpus still holds thousands of
        // items and a test asking about the last seven days gets a perfectly valid answer of
        // "none" - so it passes, and the measurement it was performing stopped happening
        // weeks ago without anything going red.
        CorpusFreshnessReport report =
            CorpusFreshness.Evaluate(Plan(), 4_000, TimeSpan.Zero, Anchor.AddDays(42));

        Assert.Equal(CorpusFreshnessVerdict.WindowsEmptied, report.Verdict);
        Assert.Contains(7, report.EmptiedWindowDays);
        Assert.Contains(30, report.EmptiedWindowDays);
        Assert.DoesNotContain(365, report.EmptiedWindowDays);

        (bool proceed, string message) = CorpusFreshness.Decide(report);
        Assert.False(proceed);
        Assert.Contains("STALE", message, StringComparison.Ordinal);
        Assert.Contains("corpus-reanchor", message, StringComparison.Ordinal);
        // The refusal must state the consequence as a count, never as prose: the date guard's
        // prose refusal is what once invited an operator to override it and lose a build.
        Assert.Contains("select 0 items", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Freshness_PastTheWidestWindowIsDeadRatherThanMerelyStale()
    {
        CorpusFreshnessReport report =
            CorpusFreshness.Evaluate(Plan(), 4_000, TimeSpan.Zero, Anchor.AddDays(400));
        Assert.Equal(CorpusFreshnessVerdict.Dead, report.Verdict);
        Assert.Equal(0, report.Windows[^1].StillInWindow);

        (bool proceed, string message) = CorpusFreshness.Decide(report);
        Assert.False(proceed);
        Assert.Contains("DEAD", message, StringComparison.Ordinal);
        Assert.Contains("EVERY window", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Freshness_AnAppliedShiftIsWhatKeepsItFresh()
    {
        // The re-anchor's whole promise: the same corpus, six weeks later, still answering.
        DateTime now = Anchor.AddDays(42);
        CorpusFreshnessReport report =
            CorpusFreshness.Evaluate(Plan(), 4_000, TimeSpan.FromDays(42), now);
        Assert.Equal(CorpusFreshnessVerdict.Fresh, report.Verdict);
        Assert.Equal(now, report.EffectiveAnchorUtc);
    }

    [Fact]
    public void Freshness_HonoursTheWindowsTheCallerActuallyTests()
    {
        // A machine that never asks about a seven-day window is not broken by one having
        // emptied. It says so by naming its own windows.
        CorpusFreshnessReport all =
            CorpusFreshness.Evaluate(Plan(), 4_000, TimeSpan.Zero, Anchor.AddDays(42));
        CorpusFreshnessReport narrow =
            CorpusFreshness.Evaluate(Plan(), 4_000, TimeSpan.Zero, Anchor.AddDays(42), new[] { 60, 365 });

        Assert.Equal(CorpusFreshnessVerdict.WindowsEmptied, all.Verdict);
        Assert.Equal(CorpusFreshnessVerdict.Fresh, narrow.Verdict);
        Assert.Equal(new[] { 60, 365 }, narrow.Windows.Select(w => w.Days).ToArray());
    }

    [Fact]
    public void Freshness_AnUnprovableShiftRefusesRatherThanReadingAsFresh()
    {
        CorpusFreshnessReport report = CorpusFreshness.Evaluate(
            Plan(), 4_000, TimeSpan.Zero, Anchor.AddMinutes(30), windowDays: null, shiftProvable: false);
        Assert.Equal(CorpusFreshnessVerdict.Unprovable, report.Verdict);

        (bool proceed, string message) = CorpusFreshness.Decide(report);
        Assert.False(proceed);
        Assert.Contains("UNPROVABLE", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Freshness_WindowsAreReportedNarrowestFirstAndCountsNeverShrinkWithWidth()
    {
        CorpusFreshnessReport report =
            CorpusFreshness.Evaluate(Plan(), 4_000, TimeSpan.Zero, Anchor.AddDays(10));
        for (int i = 1; i < report.Windows.Count; i++)
        {
            Assert.True(report.Windows[i].Days > report.Windows[i - 1].Days);
            Assert.True(report.Windows[i].StillInWindow >= report.Windows[i - 1].StillInWindow);
            Assert.True(report.Windows[i].PlannedCount >= report.Windows[i - 1].PlannedCount);
        }
    }

    [Fact]
    public void Freshness_PlannedCountsAgreeWithThePlanReportItselfl()
    {
        // The same quantity is computed in two places - the plan report's WithinDays, which
        // is the sheet a measurement is read against, and the freshness check's own
        // at-anchor column. Two places means they can drift, and a drift here would be
        // invisible: the check would still refuse at the right moment and would tell the
        // reader the corpus once held a number of items it never held.
        CorpusPlan plan = Plan();
        CorpusPlanReport sheet = plan.Report(1, 4_000);
        CorpusFreshnessReport freshness =
            CorpusFreshness.Evaluate(plan, 4_000, TimeSpan.Zero, Anchor.AddDays(42));

        Assert.Equal(CorpusPlan.MeasurementWindowDays.Count, freshness.Windows.Count);
        foreach (CorpusWindowFreshness window in freshness.Windows)
        {
            Assert.Equal(sheet.WithinDays[window.Days], window.PlannedCount);
        }
    }

    [Fact]
    public void Freshness_LiveCountsAreThePlannedOnesShiftedByTheCorpusAge()
    {
        // And the live column must be the same function of a moved clock, so a corpus that
        // has aged by exactly one window's width reports the next window's population.
        CorpusPlan plan = Plan();
        CorpusPlanReport sheet = plan.Report(1, 4_000);
        CorpusFreshnessReport onAnchor =
            CorpusFreshness.Evaluate(plan, 4_000, TimeSpan.Zero, Anchor);

        foreach (CorpusWindowFreshness window in onAnchor.Windows)
        {
            Assert.Equal(sheet.WithinDays[window.Days], window.StillInWindow);
        }
    }

    // ------------------------------------------------------------------ shift derivation

    [Fact]
    public void Shift_IsDerivedFromTheManifestRatherThanRecordedAnywhere()
    {
        CorpusPlan plan = Plan();
        TimeSpan applied = TimeSpan.FromDays(42);
        (TimeSpan shift, bool provable) = CorpusReanchor.DeriveAppliedShift(
            plan, ManifestFor(plan, 500, applied), out int agreeing, out int dated);

        Assert.True(provable);
        Assert.Equal(applied, shift);
        Assert.Equal(500, agreeing);
        Assert.Equal(500, dated);
    }

    [Fact]
    public void Shift_SurvivesAMinorityOfItemsThatDisagree()
    {
        CorpusPlan plan = Plan();
        CorpusManifest manifest = ManifestFor(plan, 500, TimeSpan.FromDays(42));

        // Twenty items whose date write went somewhere else entirely. A mean would move; a
        // mode does not, which is the point.
        for (int ordinal = 1; ordinal <= 20; ordinal++)
        {
            manifest.Add(new CorpusManifestItem(
                ordinal, "SYNTH", 6, 100, CorpusManifest.FormatUtc(new DateTime(2001, 1, 1, 0, 0, 0, DateTimeKind.Utc))));
        }

        (TimeSpan shift, bool provable) =
            CorpusReanchor.DeriveAppliedShift(plan, manifest, out int agreeing, out _);
        Assert.True(provable);
        Assert.Equal(TimeSpan.FromDays(42), shift);
        Assert.Equal(480, agreeing);
    }

    [Fact]
    public void Shift_IsNotProvableWhenTooFewItemsAgree()
    {
        CorpusPlan plan = Plan();
        CorpusManifest manifest = ManifestFor(plan, 500, TimeSpan.FromDays(42));
        for (int ordinal = 1; ordinal <= 200; ordinal++)
        {
            manifest.Add(new CorpusManifestItem(
                ordinal,
                "SYNTH",
                6,
                100,
                CorpusManifest.FormatUtc(plan.Describe(ordinal).ReceivedUtc.AddDays(ordinal))));
        }

        (_, bool provable) = CorpusReanchor.DeriveAppliedShift(plan, manifest, out _, out _);
        Assert.False(provable);
    }

    [Fact]
    public void Shift_IsNotProvableWhenNothingIsDated()
    {
        CorpusPlan plan = Plan();
        CorpusManifest manifest = ManifestFor(plan, 100, TimeSpan.Zero, undatedEvery: 1);
        (TimeSpan shift, bool provable) = CorpusReanchor.DeriveAppliedShift(plan, manifest, out _, out int dated);
        Assert.False(provable);
        Assert.Equal(0, dated);
        Assert.Equal(TimeSpan.Zero, shift);
    }

    // ------------------------------------------------------------------ re-anchor

    [Fact]
    public void Reanchor_TargetsAnAbsoluteShiftSoRunningItTwiceIsANoOp()
    {
        CorpusPlan plan = Plan();
        DateTime target = Anchor.AddDays(42);

        CorpusReanchorPlan first = CorpusReanchor.Build(plan, ManifestFor(plan, 300, TimeSpan.Zero), 300, target);
        Assert.Equal(300, first.Todo.Count);
        Assert.Equal(0, first.AlreadyCorrect);

        // The manifest a completed run would leave.
        CorpusReanchorPlan second =
            CorpusReanchor.Build(plan, ManifestFor(plan, 300, TimeSpan.FromDays(42)), 300, target);
        Assert.Empty(second.Todo);
        Assert.Equal(300, second.AlreadyCorrect);
    }

    [Fact]
    public void Reanchor_ResumesAnInterruptedRunFromTheManifestRatherThanACursor()
    {
        CorpusPlan plan = Plan();
        DateTime target = Anchor.AddDays(42);
        CorpusManifest manifest = ManifestFor(plan, 300, TimeSpan.Zero);
        for (int ordinal = 1; ordinal <= 120; ordinal++)
        {
            manifest.Add(new CorpusManifestItem(
                ordinal,
                "SYNTH",
                6,
                100,
                CorpusManifest.FormatUtc(plan.Describe(ordinal).ReceivedUtc.AddDays(42))));
        }

        CorpusReanchorPlan work = CorpusReanchor.Build(plan, manifest, 300, target);
        Assert.Equal(180, work.Todo.Count);
        Assert.Equal(120, work.AlreadyCorrect);
        Assert.DoesNotContain(work.Todo, i => i.Ordinal <= 120);
    }

    [Fact]
    public void Reanchor_KeepsEachItemsSubmitToDeliveryGap()
    {
        CorpusPlan plan = Plan();
        CorpusReanchorPlan work =
            CorpusReanchor.Build(plan, ManifestFor(plan, 200, TimeSpan.Zero), 200, Anchor.AddDays(42));
        foreach (CorpusReanchorItem item in work.Todo)
        {
            CorpusItemSpec spec = plan.Describe(item.Ordinal);
            Assert.Equal(spec.ReceivedUtc - spec.SentUtc, item.ReceivedUtc - item.SentUtc);
        }
    }

    [Fact]
    public void Reanchor_RefusesToMoveACorpusBackwardsUnlessToldTo()
    {
        CorpusPlan plan = Plan();
        CorpusManifest manifest = ManifestFor(plan, 100, TimeSpan.FromDays(42));
        CorpusReanchorPlan work = CorpusReanchor.Build(plan, manifest, 100, Anchor.AddDays(10));

        (bool proceed, string message) = CorpusReanchor.Decide(work, TimeSpan.FromDays(42), allowBackwards: false);
        Assert.False(proceed);
        Assert.Contains("BACKWARDS", message, StringComparison.Ordinal);
        Assert.True(CorpusReanchor.Decide(work, TimeSpan.FromDays(42), allowBackwards: true).Proceed);
    }

    [Fact]
    public void Reanchor_RefusesWhenTheManifestDoesNotCoverTheWholeCorpus()
    {
        CorpusPlan plan = Plan();
        CorpusReanchorPlan work =
            CorpusReanchor.Build(plan, ManifestFor(plan, 90, TimeSpan.Zero), 100, Anchor.AddDays(42));
        Assert.Equal(10, work.Unrecorded);

        (bool proceed, string message) = CorpusReanchor.Decide(work, TimeSpan.Zero, allowBackwards: false);
        Assert.False(proceed);
        Assert.Contains("corpus-reindex", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Reanchor_WritesAnItemWhoseCurrentInstantIsUnknown()
    {
        CorpusPlan plan = Plan();
        CorpusReanchorPlan work = CorpusReanchor.Build(
            plan, ManifestFor(plan, 100, TimeSpan.Zero, undatedEvery: 10), 100, Anchor.AddDays(42));
        Assert.Equal(10, work.Undated);
        Assert.Equal(100, work.Todo.Count);
    }

    [Fact]
    public void Reanchor_AndFreshnessComposeIntoARepair()
    {
        // The pair, end to end: stale at six weeks, re-anchored to now, fresh again - with
        // the plan, the seed and the shape untouched throughout.
        CorpusPlan plan = Plan();
        DateTime now = Anchor.AddDays(42);
        CorpusManifest before = ManifestFor(plan, 2_000, TimeSpan.Zero);

        (TimeSpan applied, bool provable) = CorpusReanchor.DeriveAppliedShift(plan, before, out _, out _);
        Assert.False(CorpusFreshness.Decide(
            CorpusFreshness.Evaluate(plan, 2_000, applied, now, null, provable)).Proceed);

        CorpusReanchorPlan work = CorpusReanchor.Build(plan, before, 2_000, now);
        Assert.True(CorpusReanchor.Decide(work, applied, allowBackwards: false).Proceed);

        CorpusManifest after = ManifestFor(plan, 2_000, work.TargetShift);
        (TimeSpan appliedAfter, bool provableAfter) =
            CorpusReanchor.DeriveAppliedShift(plan, after, out _, out _);
        Assert.True(CorpusFreshness.Decide(
            CorpusFreshness.Evaluate(plan, 2_000, appliedAfter, now, null, provableAfter)).Proceed);
    }

    [Fact]
    public void Manifest_TakesTheLastLineForAnOrdinal_WhichIsWhatMakesAReanchorAppendOnly()
    {
        // Load-bearing: a re-anchor records its work by APPENDING a replacement line per
        // item. If the parser preferred the first line instead, an interrupted run would be
        // unresumable and a completed one would still read as un-re-anchored.
        string header = CorpusManifest.RenderLine(new CorpusManifestHeader(
            CorpusManifest.CurrentVersion, "vm1", 4242, CorpusManifest.FormatUtc(Anchor),
            "shape", "Synthetic Data File", null, "PropertyAccessor", "InPlaceWithSentFlag"));
        string first = CorpusManifest.RenderLine(
            new CorpusManifestItem(1, "ID1", 6, 100, CorpusManifest.FormatUtc(Anchor.AddDays(-3))));
        string second = CorpusManifest.RenderLine(
            new CorpusManifestItem(1, "ID1", 0, 0, CorpusManifest.FormatUtc(Anchor.AddDays(39))));

        CorpusManifest manifest = CorpusManifest.Parse(new[] { header, first, second });
        Assert.Single(manifest.Items);
        Assert.Equal(CorpusManifest.FormatUtc(Anchor.AddDays(39)), manifest.Items[1].ReceivedUtc);
    }

    // ------------------------------------------------------------------ census

    private static IEnumerable<CorpusSighting> AsPlanned(CorpusPlan plan, int count)
    {
        for (int ordinal = 1; ordinal <= count; ordinal++)
        {
            yield return new CorpusSighting(ordinal, plan.Describe(ordinal).FolderId);
        }
    }

    [Fact]
    public void Census_IsCleanWhenEveryOrdinalExistsOnceWhereThePlanPutsIt()
    {
        CorpusPlan plan = Plan();
        CorpusCensusReport report = CorpusCensus.Compare(plan, 1_000, AsPlanned(plan, 1_000));
        Assert.Equal(0, report.Misplaced);
        Assert.Equal(0, report.DuplicatedOrdinals);
        Assert.Equal(0, report.MissingOrdinals);
        Assert.True(CorpusCensus.Decide(report).Clean);
    }

    [Fact]
    public void Census_CatchesTheWholeCorpusFiledAsDrafts()
    {
        // The first real build's second fault: 40 000 items, zero failures reported, every
        // one of them in Drafts, and the sweep therefore selecting six.
        CorpusPlan plan = Plan();
        CorpusCensusReport report = CorpusCensus.Compare(
            plan,
            1_000,
            Enumerable.Range(1, 1_000).Select(o => new CorpusSighting(o, CorpusCensus.DraftsFolderId)));

        Assert.Equal(1_000, report.StrayDrafts);
        (bool clean, string message) = CorpusCensus.Decide(report);
        Assert.False(clean);
        Assert.Contains("DRAFTS", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Census_SplitsOutboxStraysByThePlannedReadState()
    {
        // The diagnosis the first build never got. 5 532 of 40 000 items were queued for
        // delivery, and 5 532 is exactly the plan's unread count for that shape - so the
        // split is what turns a coincidence into a cause, or kills the theory outright.
        CorpusPlan plan = Plan();
        var sightings = new List<CorpusSighting>();
        for (int ordinal = 1; ordinal <= 2_000; ordinal++)
        {
            CorpusItemSpec spec = plan.Describe(ordinal);
            sightings.Add(new CorpusSighting(ordinal, spec.FolderId));
            if (!spec.IsRead)
            {
                sightings.Add(new CorpusSighting(ordinal, CorpusCensus.OutboxFolderId));
            }
        }

        CorpusCensusReport report = CorpusCensus.Compare(plan, 2_000, sightings);
        Assert.True(report.StrayOutbox > 0);
        Assert.Equal(report.StrayOutbox, report.StrayOutboxPlannedUnread);
        Assert.Equal(0, report.StrayOutboxPlannedRead);
        Assert.Equal(report.PlannedUnread, report.StrayOutboxPlannedUnread);

        (bool clean, string message) = CorpusCensus.Decide(report);
        Assert.False(clean);
        Assert.Contains("OUTBOX", message, StringComparison.Ordinal);
        Assert.Contains("queued for", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Census_ReportsTheUnreadPopulationTheOutboxCountHasToBeMatchedAgainst()
    {
        // 5,532 is not an estimate: it is what the plan for the corpus that was built
        // produces, and it is exactly the number of items that landed in the Outbox.
        CorpusPlanReport report = Plan().Report(1, 40_000);
        Assert.Equal(5_532, report.UnreadItems);
    }

    [Fact]
    public void Census_CountsDuplicatesAndMissingSeparately()
    {
        CorpusPlan plan = Plan();
        var sightings = AsPlanned(plan, 500).ToList();
        sightings.Add(new CorpusSighting(7, plan.Describe(7).FolderId));
        sightings.RemoveAll(s => s.Ordinal == 9);

        CorpusCensusReport report = CorpusCensus.Compare(plan, 500, sightings);
        Assert.Equal(1, report.DuplicatedOrdinals);
        Assert.Equal(1, report.MissingOrdinals);
        Assert.False(CorpusCensus.Decide(report).Clean);
    }

    // ------------------------------------------------------------------ placement probe

    [Fact]
    public void Placement_AnUnansweredTableCheckIsNotAFailedRung()
    {
        // The second generator defect. Against a folder holding tens of thousands of items
        // the probe's table walk hit its row cap, reported the item ABSENT, and the build
        // refused a placement that works. The refusal was right about what it saw.
        var inconclusive = new CorpusPlacementProbe(
            CorpusPlacementMethod.InPlaceWithSentFlag, "Inbox",
            ParentIsTargetFolder: true, TargetFolderTableContainsIt: false, SentFlagSet: true,
            LandedInFolderName: "Inbox", Error: null, TableCheckConclusive: false);

        Assert.False(CorpusPlacement.IsUsable(inconclusive));
        (bool proceed, string message) = CorpusPlacement.Decide(
            CorpusPlacementMethod.None, allowDraftsPlacement: false, itemCount: 40_000, probes: new[] { inconclusive });

        Assert.False(proceed);
        Assert.Contains("UNPROVEN", message, StringComparison.Ordinal);
        Assert.DoesNotContain("NOT ACHIEVABLE", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Placement_AConclusiveAbsenceIsStillARefusalAboutTheStore()
    {
        var absent = new CorpusPlacementProbe(
            CorpusPlacementMethod.InPlaceWithSentFlag, "Inbox",
            ParentIsTargetFolder: true, TargetFolderTableContainsIt: false, SentFlagSet: true,
            LandedInFolderName: "Drafts", Error: null, TableCheckConclusive: true);

        (bool proceed, string message) = CorpusPlacement.Decide(
            CorpusPlacementMethod.None, allowDraftsPlacement: false, itemCount: 40_000, probes: new[] { absent });
        Assert.False(proceed);
        Assert.Contains("NOT ACHIEVABLE", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Placement_DefaultsToConclusiveSoOlderCallSitesKeepTheirMeaning()
    {
        var probe = new CorpusPlacementProbe(
            CorpusPlacementMethod.InPlaceWithSentFlag, "Inbox", true, true, true, "Inbox", null);
        Assert.True(probe.TableCheckConclusive);
        Assert.True(CorpusPlacement.IsUsable(probe));
    }

    [Fact]
    public void ProbeFragment_IsBracketFreeAndSelectsOneOrdinal()
    {
        // Bracket-free is a safety rule, not a style: a '[' inside a DASL LIKE pattern opens
        // a character class, which is the mechanism that once destroyed real mail.
        string fragment = CorpusPlan.DaslSubjectFragment("vm1", CorpusPlan.ProbeOrdinal);
        Assert.DoesNotContain('[', fragment);
        Assert.DoesNotContain(']', fragment);
        Assert.DoesNotContain('%', fragment);

        // And it must actually occur in the subject it is meant to select.
        string subject = new CorpusPlan(new CorpusPlanOptions("vm1", 4242, Anchor)).BuildSubject(1234);
        Assert.Contains(CorpusPlan.DaslSubjectFragment("vm1", 1234), subject, StringComparison.Ordinal);
        Assert.DoesNotContain(CorpusPlan.DaslSubjectFragment("vm1", 1235), subject, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ rewrite guard

    [Fact]
    public void Rewrite_NeedsTheEntryIdTheTagsAndTheRightOrdinal()
    {
        HashSet<string> allowlist = CorpusSafety.BuildEntryIdAllowlist(new[] { "ID-1" });
        string subject = new CorpusPlan(new CorpusPlanOptions("vm1", 4242, Anchor)).BuildSubject(42);

        Assert.True(CorpusSafety.MayRewrite("ID-1", subject, allowlist, "vm1", 42));
        Assert.False(CorpusSafety.MayRewrite("ID-2", subject, allowlist, "vm1", 42));
        Assert.False(CorpusSafety.MayRewrite("ID-1", subject, allowlist, "vm1", 43));
        Assert.False(CorpusSafety.MayRewrite("ID-1", subject, allowlist, "other", 42));
        Assert.False(CorpusSafety.MayRewrite("ID-1", "an ordinary mail about invoices", allowlist, "vm1", 42));
        Assert.False(CorpusSafety.MayRewrite("ID-1", null, allowlist, "vm1", 42));
    }

    // ------------------------------------------------------------------ live settings block

    [Fact]
    public void CorpusSettings_AreCompleteOnlyWithEveryFieldTheCheckNeeds()
    {
        var complete = new CorpusSettings
        {
            StoreDisplayName = "Corpus A",
            ManifestPath = @"C:\corpus\vm1.jsonl",
            CorpusId = "vm1",
            Seed = 4242,
            AnchorUtc = "2026-08-01T00:00:00Z",
            ItemCount = 40_000,
        };
        Assert.True(complete.IsComplete);

        Assert.False(new CorpusSettings { StoreDisplayName = "Corpus A" }.IsComplete);
        Assert.False(new CorpusSettings
        {
            StoreDisplayName = "Corpus A",
            ManifestPath = @"C:\corpus\vm1.jsonl",
            CorpusId = "vm1",
            Seed = 4242,
            AnchorUtc = "2026-08-01T00:00:00Z",
            ItemCount = 0,
        }.IsComplete);
    }

    [Fact]
    public void LiveSettings_AcceptMachineProfileSpelledAsAString()
    {
        // The documented example spells it "Portable". Without a string enum converter that
        // threw inside Load, before any test ran, so a correctly written settings file made
        // the whole live tier refuse to start with an error about JSON.
        const string json = """
        {
          "machineProfile": "Portable",
          "testHubStoreDisplayName": "test@vm.invalid",
          "expectedStoreDisplayNames": [ "test@vm.invalid" ]
        }
        """;
        LiveTestSettings settings = LiveTestSettings.Parse(json);
        Assert.Equal(LiveMachineProfile.Portable, settings.MachineProfile);
    }

    [Fact]
    public void LiveSettings_StillAcceptMachineProfileSpelledAsANumber()
    {
        const string json = """
        {
          "machineProfile": 1,
          "testHubStoreDisplayName": "test@vm.invalid",
          "expectedStoreDisplayNames": [ "test@vm.invalid" ]
        }
        """;
        Assert.Equal(LiveMachineProfile.Portable, LiveTestSettings.Parse(json).MachineProfile);
    }

    // ------------------------------------------------------------------ mail sink

    [Fact]
    public void Sink_IsCompleteOnlyWithBothHalves()
    {
        Assert.True(new MailSinkSettings().IsComplete);
        Assert.False(new MailSinkSettings { SubmitPort = 0 }.IsComplete);
        Assert.False(new MailSinkSettings { RetrievePort = 0 }.IsComplete);
        Assert.False(new MailSinkSettings { SubmitHost = "  " }.IsComplete);
        Assert.False(new MailSinkSettings { RetrieveHost = string.Empty }.IsComplete);
    }

    [Fact]
    public void Sink_ProbePassesWhenBothListenersAnswer()
    {
        using var submit = new LoopbackListener();
        using var retrieve = new LoopbackListener();
        (bool reachable, string message) = LiveMailSink.Probe(new MailSinkSettings
        {
            SubmitPort = submit.Port,
            RetrievePort = retrieve.Port,
        });

        Assert.True(reachable, message);
        Assert.Contains("both answering", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Sink_ProbeNamesTheHalfThatIsMissingRatherThanTheFirstFailure()
    {
        // A sink that accepts mail and cannot hand it back is the exact failure this design
        // exists to avoid, so the message has to name the retrieval half even when it is the
        // second thing checked.
        using var submit = new LoopbackListener();
        int deadPort = LoopbackListener.ReserveAndRelease();
        (bool reachable, string message) = LiveMailSink.Probe(new MailSinkSettings
        {
            SubmitPort = submit.Port,
            RetrievePort = deadPort,
            ConnectTimeoutMs = 750,
        });

        Assert.False(reachable);
        Assert.Contains(deadPort.ToString(System.Globalization.CultureInfo.InvariantCulture), message, StringComparison.Ordinal);
        Assert.DoesNotContain(
            submit.Port.ToString(System.Globalization.CultureInfo.InvariantCulture), message, StringComparison.Ordinal);
        Assert.Contains("Outbox", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Sink_ProbeReportsBothHalvesWhenNeitherAnswers()
    {
        int a = LoopbackListener.ReserveAndRelease();
        int b = LoopbackListener.ReserveAndRelease();
        (bool reachable, string message) = LiveMailSink.Probe(new MailSinkSettings
        {
            SubmitPort = a,
            RetrievePort = b,
            ConnectTimeoutMs = 750,
        });

        Assert.False(reachable);
        Assert.Contains(a.ToString(System.Globalization.CultureInfo.InvariantCulture), message, StringComparison.Ordinal);
        Assert.Contains(b.ToString(System.Globalization.CultureInfo.InvariantCulture), message, StringComparison.Ordinal);
    }

    [Fact]
    public void LiveSettings_RejectAHalfWrittenMailSinkBlock()
    {
        const string json = """
        {
          "machineProfile": "Portable",
          "testHubStoreDisplayName": "test@vm.invalid",
          "expectedStoreDisplayNames": [ "test@vm.invalid" ],
          "mailSink": { "submitHost": "127.0.0.1", "submitPort": 25, "retrievePort": 0 }
        }
        """;
        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() => LiveTestSettings.Parse(json));
        Assert.Contains("partially filled 'mailSink' block", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void LiveSettings_AbsentMailSinkMeansRealTransport()
    {
        const string json = """
        {
          "machineProfile": "Portable",
          "testHubStoreDisplayName": "test@vm.invalid",
          "expectedStoreDisplayNames": [ "test@vm.invalid" ]
        }
        """;
        LiveTestSettings settings = LiveTestSettings.Parse(json);
        Assert.Null(settings.MailSink);
        Assert.Contains("mailSink=none (real transport)", settings.Describe(), StringComparison.Ordinal);
    }

    /// <summary>A loopback listener on an ephemeral port, so the probe has something real to connect to.</summary>
    private sealed class LoopbackListener : IDisposable
    {
        private readonly System.Net.Sockets.TcpListener _listener;

        internal LoopbackListener()
        {
            _listener = new System.Net.Sockets.TcpListener(System.Net.IPAddress.Loopback, 0);
            _listener.Start();
            Port = ((System.Net.IPEndPoint)_listener.LocalEndpoint).Port;
        }

        internal int Port { get; }

        /// <summary>An ephemeral port that was bound and released - so nothing is listening on it now.</summary>
        internal static int ReserveAndRelease()
        {
            var probe = new System.Net.Sockets.TcpListener(System.Net.IPAddress.Loopback, 0);
            probe.Start();
            int port = ((System.Net.IPEndPoint)probe.LocalEndpoint).Port;
            probe.Stop();
            return port;
        }

        public void Dispose() => _listener.Stop();
    }

    [Fact]
    public void LiveSettings_RejectAHalfWrittenCorpusBlock()
    {
        const string json = """
        {
          "machineProfile": "Portable",
          "testHubStoreDisplayName": "test@vm.invalid",
          "expectedStoreDisplayNames": [ "test@vm.invalid" ],
          "corpus": { "storeDisplayName": "Corpus A", "corpusId": "vm1" }
        }
        """;
        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() => LiveTestSettings.Parse(json));
        Assert.Contains("partially filled 'corpus' block", ex.Message, StringComparison.Ordinal);
    }
}
