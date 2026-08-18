using System;
using System.Collections.Generic;

using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The freshness block must describe THIS search - its store, its window, its frontier -
/// and nothing else. Three defects, one theme, all of them visible to an agent through
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
/// WHAT IS PROVEN HERE AND WHAT IS NOT. All three decisions are pure functions over data
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

    [Fact]
    public void AbsentFoldersAreAttributedToo_AndStillDegradeNothing()
    {
        // A store without a Junk Email folder must not lend its absence to another store's
        // arithmetic, and absence remains a non-gap either way.
        ComSweepResult result = new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 7,
            foldersSkipped: 0,
            sweptFolders: new[] { StoreA + "/Inbox", StoreB + "/Inbox" },
            foldersAbsent: 1,
            perStore: new[]
            {
                new ComStoreSweepCounters(StoreA, foldersSwept: 4, foldersSkipped: 0, foldersFailed: 0, foldersAbsent: 0),
                new ComStoreSweepCounters(StoreB, foldersSwept: 3, foldersSkipped: 0, foldersFailed: 0, foldersAbsent: 1),
            });

        Assert.Null(Applied(result, StoreA).FoldersAbsent);
        Assert.Equal(1, Applied(result, StoreB).FoldersAbsent);
        Assert.Null(Applied(result, StoreB).CoverageGaps);
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
}
