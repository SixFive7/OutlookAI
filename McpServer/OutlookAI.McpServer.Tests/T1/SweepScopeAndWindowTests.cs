using System;
using System.Collections.Generic;

using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The freshness block must describe THIS search - its store, its window, its frontier -
/// and nothing else. Four defects, one theme, all of them visible to an agent through
/// <c>degraded</c>, which the search tool's own description tells it to relay to the user.
/// <para>
/// (1) COUNTERS CROSSED STORE BOUNDARIES. A cached all-stores sweep may serve a
/// store-scoped search - safe for the DATA, since a superset contains what was asked - and
/// the folder LISTS were narrowed to the requested store from the start. The scalar
/// counters were not, and were merely informational until they began driving the coverage
/// codes; from then on a search scoped to store A reported <c>degraded: true</c> because
/// store B had an unreadable folder.
/// </para>
/// <para>
/// (2) AN EMPTY WINDOW LOOKED LIKE A FAILED SWEEP. A search bounded by <c>before</c> to
/// mail older than the index frontier has nothing for the sweep to find; the code set
/// <c>performed = false</c>, which means "could not run", so the answer came back
/// <c>degraded</c> and <c>index-only</c> - warned that it might be missing recent mail that
/// cannot exist inside its own bounds.
/// </para>
/// <para>
/// (3) THE FRONTIER WAS PROFILE-WIDE. The staleness probe ran unscoped for every search,
/// so a busy store's frontier set a quiet store's sweep window (measured 2026-08-18: five
/// store-scoped searches, one frontier, against per-store probes spanning 45.4 hours).
/// </para>
/// <para>
/// (4) FIXING (1) BROKE (1) AGAIN, ONE SCOPE DOWN. Per-store counters gave a store with no
/// arrival-path folders <c>foldersSwept: 0, foldersAbsent: 4</c>, and "swept nothing" was
/// read as a coverage hole without asking why - so a PST, an archive-only store or a shared
/// mailbox mounted without the four defaults made every search naming it <c>degraded</c>,
/// the very alarm the absent counter had been introduced to remove. Whole-sweep, the same
/// call reported <c>live</c>, because another store's folders masked the zero.
/// </para>
/// <para>
/// WHAT IS PROVEN HERE AND WHAT IS NOT. All these decisions are pure functions over data
/// the COM layer produces, and every branch of them is pinned below. What needs a real
/// mailbox is the COM layer FILLING that data in - a multi-store profile with a folder that
/// will not open, and a store whose index lags another's - so the sweep result is modelled
/// here in the shape <c>OutlookComSession.SweepFoldersNewerThan</c> builds, and the live
/// tier exercises the production of it (T2 LiveSweepScopeTests, LiveFreshModeTests).
/// </para>
/// </summary>
public sealed class SweepScopeAndWindowTests
{
    private const string StoreA = "alice@example.com";
    private const string StoreB = "bob@example.com";

    /// <summary>An archive-only data file: a store with none of the four arrival-path folders.</summary>
    private const string StoreC = "Archive 2019.pst";

    // ============================================ (1) counters belong to ONE store

    /// <summary>
    /// The shape the default-folder sweep produces on a two-store profile where store B has
    /// one folder that will not open: four folders swept in A, three plus a failure in B.
    /// </summary>
    private static ComSweepResult TwoStoresOneFailingFolderInB()
    {
        return new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 7,
            foldersSkipped: 1,
            sweptFolders: new[]
            {
                StoreA + "/Inbox", StoreA + "/Sent Items", StoreA + "/Deleted Items", StoreA + "/Junk Email",
                StoreB + "/Inbox", StoreB + "/Sent Items", StoreB + "/Deleted Items",
            },
            foldersFailed: 1,
            foldersAbsent: 0,
            perStore: new[]
            {
                new ComStoreSweepCounters(StoreA, foldersSwept: 4, foldersSkipped: 0, foldersFailed: 0, foldersAbsent: 0),
                new ComStoreSweepCounters(StoreB, foldersSwept: 3, foldersSkipped: 1, foldersFailed: 1, foldersAbsent: 0),
            });
    }

    private static SweepInfo Applied(ComSweepResult result, string? store)
    {
        SweepInfo info = new SweepInfo { Performed = true };
        MailService.ApplySweepCounters(info, result, store);
        info.CoverageGaps = FreshMerge.DescribeCoverageGaps(info);
        return info;
    }

    [Fact]
    public void ASearchScopedToAHealthyStore_IsNotDegradedByAnotherStoresFailure()
    {
        // THE DEFECT: store B's unreadable folder made this answer - which is entirely
        // about store A - report degraded: true with a folders_failed coverage gap.
        SweepInfo info = Applied(TwoStoresOneFailingFolderInB(), StoreA);

        Assert.Equal(4, info.FoldersSwept);
        Assert.Equal(0, info.FoldersSkipped);
        Assert.Equal(0, info.FoldersFailed);
        Assert.Null(info.CoverageGaps);
        Assert.Equal(FreshMerge.FreshnessLive, FreshMerge.ClassifyFreshness(info));

        // The list was always narrowed; it and the counters must now agree.
        Assert.NotNull(info.Folders);
        Assert.Equal(info.FoldersSwept, info.Folders!.Count);
        Assert.All(info.Folders, f => Assert.StartsWith(StoreA + "/", f, StringComparison.Ordinal));
    }

    [Fact]
    public void ASearchScopedToTheFailingStore_StillSeesTheFailure()
    {
        // The other half: buying quiet by dropping a real coverage hole would be the worse
        // bug. Scoped to B, the failure is B's own and must be reported.
        SweepInfo info = Applied(TwoStoresOneFailingFolderInB(), StoreB);

        Assert.Equal(3, info.FoldersSwept);
        Assert.Equal(1, info.FoldersSkipped);
        Assert.Equal(1, info.FoldersFailed);
        Assert.Contains(FreshMerge.GapFoldersFailed, info.CoverageGaps!);
        Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyFreshness(info));
    }

    [Fact]
    public void AnUnscopedSearch_ReadsTheWholeSweep_BecauseThatIsWhatItAsked()
    {
        SweepInfo info = Applied(TwoStoresOneFailingFolderInB(), store: null);

        Assert.Equal(7, info.FoldersSwept);
        Assert.Equal(1, info.FoldersSkipped);
        Assert.Equal(1, info.FoldersFailed);
        Assert.Contains(FreshMerge.GapFoldersFailed, info.CoverageGaps!);
    }

    [Fact]
    public void AStoreTheSweepNeverReached_ReportsNoCoverage_NotSomeoneElses()
    {
        // A missing entry is ZERO coverage, not "fall back to the totals" - falling back
        // would resurrect the defect in the one case where it misleads most, and the
        // nothing_swept gap is the honest answer.
        SweepInfo info = Applied(TwoStoresOneFailingFolderInB(), "carol@example.com");

        Assert.Equal(0, info.FoldersSwept);
        Assert.Equal(0, info.FoldersFailed);
        Assert.Null(info.Folders);
        Assert.Contains(FreshMerge.GapNothingSwept, info.CoverageGaps!);
    }

    /// <summary>
    /// Three stores, so absence is exercised at BOTH strengths rather than only the easy
    /// one: A has all four arrival-path folders, B is missing one of them, and C - a PST, an
    /// archive-only store, a shared mailbox mounted without the defaults - has none of them.
    /// Nothing anywhere failed or was skipped, so absence is the only thing under test.
    /// </summary>
    private static ComSweepResult ThreeStoresOneMissingAFolder_AndOneMissingThemAll()
    {
        return new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 7,
            foldersSkipped: 0,
            sweptFolders: new[]
            {
                StoreA + "/Inbox", StoreA + "/Sent Items", StoreA + "/Deleted Items", StoreA + "/Junk Email",
                StoreB + "/Inbox", StoreB + "/Sent Items", StoreB + "/Deleted Items",
            },
            foldersFailed: 0,
            foldersAbsent: 5,
            perStore: new[]
            {
                new ComStoreSweepCounters(StoreA, foldersSwept: 4, foldersSkipped: 0, foldersFailed: 0, foldersAbsent: 0),
                new ComStoreSweepCounters(StoreB, foldersSwept: 3, foldersSkipped: 0, foldersFailed: 0, foldersAbsent: 1),
                new ComStoreSweepCounters(StoreC, foldersSwept: 0, foldersSkipped: 0, foldersFailed: 0, foldersAbsent: 4),
            });
    }

    [Fact]
    public void AbsentFoldersAreAttributedToo_AndStillDegradeNothing()
    {
        // A store without a Junk Email folder must not lend its absence to another store's
        // arithmetic, and absence remains a non-gap either way.
        //
        // "EITHER WAY" IS THE PART THAT HAD NO FIXTURE. This test used to give its absent
        // store foldersSwept: 3, so it only ever proved the easy half - absence alongside
        // real coverage - and the case it is named for never ran. Store C below is the hard
        // half: every folder the sweep set out to walk is missing, so it swept nothing at
        // all, and until 2026-08-18 that made every search naming such a store report
        // freshness: partial and degraded: true. Read whole-sweep (before the per-store
        // split of c515565) the same call had said live, because A's and B's seven folders
        // masked the zero.
        ComSweepResult result = ThreeStoresOneMissingAFolder_AndOneMissingThemAll();

        // A: nothing missing, nothing borrowed from the others.
        SweepInfo healthy = Applied(result, StoreA);
        Assert.Equal(4, healthy.FoldersSwept);
        Assert.Null(healthy.FoldersAbsent);
        Assert.Null(healthy.CoverageGaps);

        // B: one folder absent, three swept - absence attributed, and no gap.
        SweepInfo partlyAbsent = Applied(result, StoreB);
        Assert.Equal(3, partlyAbsent.FoldersSwept);
        Assert.Equal(1, partlyAbsent.FoldersAbsent);
        Assert.Null(partlyAbsent.CoverageGaps);
        Assert.Equal(FreshMerge.FreshnessLive, FreshMerge.ClassifyFreshness(partlyAbsent));

        // C: ALL four absent, so nothing was swept - because there was nothing to sweep. A
        // folder that does not exist cannot be hiding mail, which is the rule e706315 set
        // one level up; the per-store split must not smuggle it back in.
        SweepInfo whollyAbsent = Applied(result, StoreC);
        Assert.Equal(0, whollyAbsent.FoldersSwept);
        Assert.Equal(0, whollyAbsent.FoldersSkipped);
        Assert.Equal(0, whollyAbsent.FoldersFailed);
        Assert.Equal(4, whollyAbsent.FoldersAbsent);
        Assert.Null(whollyAbsent.Folders);
        Assert.Null(whollyAbsent.CoverageGaps);
        Assert.Equal(FreshMerge.FreshnessLive, FreshMerge.ClassifyFreshness(whollyAbsent));

        // And the prose an agent relays says nothing either - a flag that cries wolf is
        // worse than no flag, and so is a sentence.
        Assert.Empty(MailService.DescribeSweepCoverage(whollyAbsent, "12 minutes", folderScoped: false));
    }

    [Fact]
    public void AStoreWhoseFoldersAllFAILED_SweptNothingEither_AndIsStillDegraded()
    {
        // The other half, and the reason this is not simply "foldersSwept == 0 is fine".
        // Same zero, opposite meaning: those four folders exist and hold mail nobody could
        // read, so the answer really is missing something and must say so twice over.
        ComSweepResult result = new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 4,
            foldersSkipped: 4,
            sweptFolders: new[] { StoreA + "/Inbox" },
            foldersFailed: 4,
            foldersAbsent: 0,
            perStore: new[]
            {
                new ComStoreSweepCounters(StoreA, foldersSwept: 4, foldersSkipped: 0, foldersFailed: 0, foldersAbsent: 0),
                new ComStoreSweepCounters(StoreB, foldersSwept: 0, foldersSkipped: 4, foldersFailed: 4, foldersAbsent: 0),
            });

        SweepInfo scoped = Applied(result, StoreB);

        Assert.Equal(0, scoped.FoldersSwept);
        Assert.Null(scoped.FoldersAbsent);
        Assert.Contains(FreshMerge.GapNothingSwept, scoped.CoverageGaps!);
        Assert.Contains(FreshMerge.GapFoldersFailed, scoped.CoverageGaps!);
        Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyFreshness(scoped));
    }

    [Fact]
    public void AStoreWhereOneFolderIsMissingAndAnotherUnreadable_IsStillDegraded()
    {
        // The mixed case, which is where a naive "absent > 0 means nothing was wrong" rule
        // would fail: two folders absent and two unreadable is still two folders of mail
        // nobody checked, and absence must not launder them.
        ComSweepResult result = new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 0,
            foldersSkipped: 2,
            sweptFolders: Array.Empty<string>(),
            foldersFailed: 2,
            foldersAbsent: 2,
            perStore: new[]
            {
                new ComStoreSweepCounters(StoreB, foldersSwept: 0, foldersSkipped: 2, foldersFailed: 2, foldersAbsent: 2),
            });

        SweepInfo scoped = Applied(result, StoreB);

        Assert.Equal(2, scoped.FoldersAbsent);
        Assert.Contains(FreshMerge.GapNothingSwept, scoped.CoverageGaps!);
        Assert.Contains(FreshMerge.GapFoldersFailed, scoped.CoverageGaps!);
        Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyFreshness(scoped));
    }

    [Fact]
    public void AWholeProfileOfStoresWithoutDefaultFolders_IsNotDegradedEither()
    {
        // Unscoped, the totals say the same thing: the sweep ran, found no arrival-path
        // folder anywhere to walk, and withheld nothing. This is the reading the whole-sweep
        // counters gave BEFORE the per-store split, so the two must not disagree.
        ComSweepResult result = new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 0,
            foldersSkipped: 0,
            sweptFolders: Array.Empty<string>(),
            foldersFailed: 0,
            foldersAbsent: 4,
            perStore: new[]
            {
                new ComStoreSweepCounters(StoreB, foldersSwept: 0, foldersSkipped: 0, foldersFailed: 0, foldersAbsent: 4),
            });

        SweepInfo unscoped = Applied(result, store: null);

        Assert.Equal(0, unscoped.FoldersSwept);
        Assert.Equal(4, unscoped.FoldersAbsent);
        Assert.Null(unscoped.CoverageGaps);
        Assert.Equal(FreshMerge.FreshnessLive, FreshMerge.ClassifyFreshness(unscoped));
    }

    [Fact]
    public void AFolderScopedSweep_AttributesItsWholeTallyToItsOneStore()
    {
        // The scoped path covers one store by construction, and reports itself that way so
        // the caller has ONE rule for reading counters instead of one rule per sweep shape.
        ComSweepResult result = new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 2,
            foldersSkipped: 1,
            sweptFolders: new[] { StoreA + "/Projects", StoreA + "/Projects/2026" },
            foldersFailed: 1,
            perStore: new[]
            {
                new ComStoreSweepCounters(StoreA, foldersSwept: 2, foldersSkipped: 1, foldersFailed: 1, foldersAbsent: 0),
            });

        SweepInfo info = Applied(result, StoreA);
        Assert.Equal(2, info.FoldersSwept);
        Assert.Equal(1, info.FoldersSkipped);
        Assert.Equal(1, info.FoldersFailed);
        Assert.Contains(FreshMerge.GapFoldersFailed, info.CoverageGaps!);
    }

    [Fact]
    public void TheSweptFolderListCap_MeasuresWhatIsINSCOPE_NotTheWholeSweep()
    {
        // Same family, one layer down: folderListOmitted said "the cap dropped the list"
        // whenever the WHOLE sweep exceeded the cap, including when the requested store
        // contributed nothing at all - a cap reported over a list that was never long.
        List<string> wide = new List<string>();
        for (int i = 0; i < MailService.SweptFolderListCap + 2; i++)
        {
            wide.Add(StoreB + "/Folder " + i.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        ComSweepResult result = new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: wide.Count,
            foldersSkipped: 0,
            sweptFolders: wide,
            perStore: new[]
            {
                new ComStoreSweepCounters(StoreB, wide.Count, 0, 0, 0),
            });

        SweepInfo elsewhere = Applied(result, StoreA);
        Assert.Null(elsewhere.Folders);
        Assert.Null(elsewhere.FolderListOmitted);

        // And the real omission still reports itself.
        SweepInfo owner = Applied(result, StoreB);
        Assert.Null(owner.Folders);
        Assert.True(owner.FolderListOmitted);
    }

    // ==================================== (2) "did not need to run" is not "could not run"

    [Theory]
    // before strictly inside the gap: there IS unindexed ground to cover.
    [InlineData(30, false)]
    [InlineData(1, false)]
    // The boundary is inclusive because 'before' is exclusive: at before == gapStart the
    // window holds no instant at all.
    [InlineData(0, true)]
    [InlineData(-1, true)]
    [InlineData(-2880, true)]
    public void DecideSweepWindow_TurnsOnWhetherTheWindowHoldsAnything(int beforeOffsetMinutes, bool notNeeded)
    {
        DateTime gapStart = new DateTime(2026, 8, 18, 3, 0, 0, DateTimeKind.Utc);
        FreshMerge.SweepWindowVerdict expected = notNeeded
            ? FreshMerge.SweepWindowVerdict.NotNeeded
            : FreshMerge.SweepWindowVerdict.Needed;

        Assert.Equal(expected, FreshMerge.DecideSweepWindow(gapStart, gapStart.AddMinutes(beforeOffsetMinutes)));
    }

    [Fact]
    public void NoBeforeBound_AlwaysNeedsTheSweep()
    {
        Assert.Equal(
            FreshMerge.SweepWindowVerdict.Needed,
            FreshMerge.DecideSweepWindow(DateTime.UtcNow, beforeUtc: null));
    }

    [Fact]
    public void ASweepThatWasNotNeeded_IsACompleteAnswer_NotADegradedOne()
    {
        // THE DEFECT: this exact block used to classify as "index-only", so a search
        // deliberately aimed at older mail was told it might be missing recent mail.
        SweepInfo sweep = new SweepInfo
        {
            Performed = false,
            NotNeeded = true,
            GapStartUtc = new DateTime(2026, 8, 18, 3, 0, 0, DateTimeKind.Utc),
        };

        Assert.Null(FreshMerge.DescribeCoverageGaps(sweep));
        Assert.Equal(FreshMerge.FreshnessLive, FreshMerge.ClassifyFreshness(sweep));
    }

    [Fact]
    public void ASweepThatCouldNotRun_IsStillIndexOnly()
    {
        // The regression guard on the other side: the new state must not swallow the old
        // one. Both are performed=false; only the REASON tells them apart.
        foreach (SweepInfo failed in new[]
                 {
                     new SweepInfo { Performed = false, Error = "OutlookUnavailable" },
                     new SweepInfo { Performed = false, Error = FreshMerge.AttachmentContentNotSweepable },
                     new SweepInfo { Performed = false },
                 })
        {
            Assert.Equal(FreshMerge.FreshnessIndexOnly, FreshMerge.ClassifyFreshness(failed));
        }
    }

    // ====================================== (3) the frontier is measured over what is searched

    [Fact]
    public void AnUnscopedSearch_MeasuresTheFrontierOverTheWholeProfile()
    {
        Assert.Null(MailService.StalenessScopeFor(null));
    }

    [Fact]
    public void AStoreScopedSearch_MeasuresTheFrontierOverThatStore()
    {
        string prefix = "mapi16://{S-1-5-21-1-2-3-1001}/" + StoreA + "($abcd1234)";
        FolderScopeResolution resolution = FolderScopeResolver.ForPrimaryStore(prefix, folder: null, includeSubfolders: true);

        Assert.Equal(prefix, MailService.StalenessScopeFor(resolution));
    }

    [Fact]
    public void AFolderScopedSearch_MeasuresItOverTheSTORE_NotTheFolder()
    {
        // Deliberate, and the reason is what the number IS: an ingestion frontier. A store
        // indexes as its own subtree, so its newest item tracks how far ingestion has got;
        // a quiet Archive folder's newest item is old because nothing ARRIVES there, which
        // says nothing about the index - and scoping to it would widen that search's sweep
        // window to years of already-indexed mail.
        string prefix = "mapi16://{S-1-5-21-1-2-3-1001}/" + StoreA + "($abcd1234)";
        FolderScopeResolution resolution =
            FolderScopeResolver.ForPrimaryStore(prefix, "Archive/2019", includeSubfolders: true);

        Assert.Equal(prefix, MailService.StalenessScopeFor(resolution));
        Assert.NotEqual(resolution.Scope, MailService.StalenessScopeFor(resolution));
    }

    [Fact]
    public void ADelegateFolderScope_MeasuresItOverTheDelegateStoreRoot()
    {
        string root = "mapi16://{S-1-5-21-1-2-3-1001}/owner@example.com($abcd1234)/1/Delegate Name";
        FolderScopeResolution resolution = FolderScopeResolver.ForDelegateStore(
            root, "Postvak IN", includeSubfolders: false, comFolderPaths: new[] { "Postvak IN" });

        Assert.Equal(root, MailService.StalenessScopeFor(resolution));
    }

    [Fact]
    public void AQuietStore_IsNotAccusedOfAStaleIndex()
    {
        // Scope-aware staleness makes a large age COMMON (a low-traffic account is quiet,
        // not lagging), so the scoped wording states what the number actually proves.
        string scoped = MailService.DescribeStaleIndex(MailService.VeryStaleAdviceMinutes + 60, storeScoped: true)!;
        Assert.Contains("newest indexed mail in this store is 13 h old", scoped, StringComparison.Ordinal);
        Assert.Contains("quiet", scoped, StringComparison.Ordinal);
        Assert.DoesNotContain("very stale", scoped, StringComparison.Ordinal);

        // Unscoped, the profile-wide frontier really does describe the index.
        string profile = MailService.DescribeStaleIndex(MailService.VeryStaleAdviceMinutes + 60, storeScoped: false)!;
        Assert.Contains("The index is very stale (13 h behind)", profile, StringComparison.Ordinal);

        // Same remedy either way - it does not depend on which cause it is.
        Assert.Contains("exhaustive:true", scoped, StringComparison.Ordinal);
        Assert.Contains("exhaustive:true", profile, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null)]
    [InlineData(0d)]
    [InlineData(60d)]
    [InlineData(MailService.VeryStaleAdviceMinutes)]
    public void AnIndexInsideTheThreshold_SaysNothingAtAll(double? ageMinutes)
    {
        Assert.Null(MailService.DescribeStaleIndex(ageMinutes, storeScoped: true));
        Assert.Null(MailService.DescribeStaleIndex(ageMinutes, storeScoped: false));
    }

    // ============================================ (5) one window PER STORE, not one per search

    // THE DEFECT: (3) fixed the frontier for a STORE-SCOPED search only. An unscoped search
    // still opened ONE window from the profile-wide frontier - the newest instant ANY store
    // ingested - so a store lagging by hours was swept back only as far as the busiest
    // store's clock, and the rest of its gap sat in neither tier. Measured on this machine
    // the same day: two catalog stores, frontiers 11 minutes apart. Half a fix reads as a
    // whole one, so the half that had not landed was invisible.

    private static readonly DateTime NowUtc = new(2026, 08, 18, 12, 00, 00, DateTimeKind.Utc);

    [Fact]
    public void EachStoreIsSweptFromItsOwnFrontier_AndAnUnnamedStoreGetsTheFallback()
    {
        Dictionary<string, DateTime> windows = new(StringComparer.Ordinal)
        {
            [StoreA] = NowUtc.AddMinutes(-11),
            [StoreB] = NowUtc.AddHours(-9),
        };

        IReadOnlyDictionary<string, DateTime> resolved = OutlookComSession.NormalizeSweepWindows(windows);
        DateTime fallback = NowUtc.AddDays(-7);

        Assert.Equal(NowUtc.AddMinutes(-11), OutlookComSession.WindowFor(resolved, StoreA, fallback));
        Assert.Equal(NowUtc.AddHours(-9), OutlookComSession.WindowFor(resolved, StoreB, fallback));

        // The store nobody measured gets the WIDEST window, not the profile frontier: it is
        // the one store whose gap is unknown, so the narrow window is the wrong guess.
        Assert.Equal(fallback, OutlookComSession.WindowFor(resolved, StoreC, fallback));
    }

    [Fact]
    public void AWindowMatchesItsStore_CaseInsensitively_BecauseEverythingElseDoes()
    {
        // The map crosses a JSON boundary and comes back with an ORDINAL comparer whatever
        // the sender used. A miss here is silent: the store just gets the fallback window and
        // nothing in the payload says its own one was not applied.
        IReadOnlyDictionary<string, DateTime> resolved = OutlookComSession.NormalizeSweepWindows(
            new Dictionary<string, DateTime>(StringComparer.Ordinal) { ["Archive 2019.PST"] = NowUtc.AddHours(-3) });

        Assert.Equal(NowUtc.AddHours(-3), OutlookComSession.WindowFor(resolved, "archive 2019.pst", NowUtc));
    }

    [Fact]
    public void TwoSpellingsOfOneStore_KeepTheEarlierWindow()
    {
        // They cannot both be that store's frontier, and the wider window is the one that
        // cannot hide mail.
        IReadOnlyDictionary<string, DateTime> resolved = OutlookComSession.NormalizeSweepWindows(
            new Dictionary<string, DateTime>(StringComparer.Ordinal)
            {
                ["Archive.pst"] = NowUtc.AddHours(-1),
                ["ARCHIVE.PST"] = NowUtc.AddHours(-6),
            });

        Assert.Equal(NowUtc.AddHours(-6), OutlookComSession.WindowFor(resolved, "Archive.pst", NowUtc));
    }

    [Fact]
    public void TheReportedWindow_IsTheWidestOneOpened_NotTheNarrowest()
    {
        // One number over a per-store decision. The earliest is the honest one: the claim it
        // supports - "the merged answer covers everything from here to now" - holds for every
        // store, because a store swept from a LATER start has its index covering the span in
        // front of that. The latest would understate coverage that was actually delivered.
        Dictionary<string, DateTime> windows = new(StringComparer.OrdinalIgnoreCase)
        {
            [StoreA] = NowUtc.AddMinutes(-11),
            [StoreB] = NowUtc.AddHours(-9),
        };

        Assert.Equal(NowUtc.AddHours(-9), MailService.WidestWindow(NowUtc.AddMinutes(-30), windows));

        // The fallback counts too - it is the window an undiscovered store gets.
        Assert.Equal(NowUtc.AddDays(-7), MailService.WidestWindow(NowUtc.AddDays(-7), windows));
    }

    [Fact]
    public void TheReportedWindow_CountsOnlyTheStoresTheSweepActuallyVisited()
    {
        // Before the sweep runs, the fallback has to be assumed to apply to SOMEONE, so the
        // planned widest window is 7 days on every unscoped search. On a fully catalogued
        // profile it applies to no one, and reporting the plan would say "swept back 7 days"
        // over a sweep that swept back eleven minutes.
        Dictionary<string, DateTime> windows = new(StringComparer.OrdinalIgnoreCase)
        {
            [StoreA] = NowUtc.AddMinutes(-11),
            [StoreB] = NowUtc.AddHours(-9),
        };

        ComSweepResult visited = TwoStoresOneFailingFolderInB();
        Assert.Equal(
            NowUtc.AddHours(-9),
            MailService.WindowActuallyUsed(visited, requestedStore: null, NowUtc.AddDays(-7), windows));

        // A store-scoped request reads ITS window out of a broad sweep, not the widest.
        Assert.Equal(
            NowUtc.AddMinutes(-11),
            MailService.WindowActuallyUsed(visited, StoreA, NowUtc.AddDays(-7), windows));
    }

    [Fact]
    public void AStoreTheSweepVisitedWithNoWindowOfItsOwn_ReportsTheFallback()
    {
        // Which is the whole point of the fallback being the widest span rather than the
        // profile frontier: this is the store the index could not be asked about.
        ComSweepResult visited = new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 4,
            foldersSkipped: 0,
            perStore: new[] { new ComStoreSweepCounters(StoreC, 4, 0, 0, 0) });

        Assert.Equal(
            NowUtc.AddDays(-7),
            MailService.WindowActuallyUsed(visited, requestedStore: null, NowUtc.AddDays(-7), null));
    }
}
