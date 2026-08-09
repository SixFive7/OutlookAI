using System;
using System.Collections.Generic;
using OutlookAI.Core.Com;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The bounds of the folder-scoped freshness sweep's subtree walk
/// (<see cref="OutlookComSession.DecideSweepWalk"/>).
/// <para>
/// The walk itself only exists over live COM folders, so the descend/stop DECISION is a
/// pure function and is pinned here - the same shape as
/// <c>ComposeSurface.SelectWindowsToPark</c>, and for the same reason: the rule is a
/// safety boundary that must not be left implicit inside the COM path.
/// </para>
/// <para>
/// What it is a boundary against: the walk used to be bounded only by
/// <c>sweptFolders.Count &gt;= MaxScopedSweepFolders</c>, which counts folders
/// SUCCESSFULLY SWEPT. Non-mail folders (Calendar/Contacts/Tasks) are never swept but
/// ARE recursed into, and a folder whose table fails to open is counted as failed rather
/// than swept - so a wide subtree of either kind was walked in full at zero cap cost,
/// with no depth guard and no clock. This runs inline on every folder-scoped search.
/// </para>
/// </summary>
public sealed class SweepWalkBoundsTests
{
    private static readonly TimeSpan Budget = TimeSpan.FromMilliseconds(OutlookComSession.ScopedSweepTimeBudgetMs);

    [Fact]
    public void Bounds_ArePinned()
    {
        // Cap creep in either direction is a decision, not a refactor: the cap is the
        // sweep's cost ceiling (~10 ms per folder) and the budget is inline latency on
        // every folder-scoped search.
        Assert.Equal(40, OutlookComSession.MaxScopedSweepFolders);
        Assert.Equal(2_000, OutlookComSession.ScopedSweepTimeBudgetMs);

        // Two orders of magnitude below the opt-in exhaustive scan's 120 s budget: this
        // walk is not opt-in, it runs as part of an ordinary search.
        Assert.True(OutlookComSession.ScopedSweepTimeBudgetMs <= 5_000);
    }

    [Fact]
    public void NormalWalk_VisitsTheFolder_WhenNoBoundApplies()
    {
        Assert.Equal(
            OutlookComSession.SweepWalkVerdict.Visit,
            OutlookComSession.DecideSweepWalk(foldersVisited: 0, depth: 0, elapsed: TimeSpan.Zero));

        // The ordinary case this must never disturb: a folder-scoped sweep a few levels
        // deep, a handful of folders in, well inside its budget.
        Assert.Equal(
            OutlookComSession.SweepWalkVerdict.Visit,
            OutlookComSession.DecideSweepWalk(foldersVisited: 12, depth: 4, elapsed: TimeSpan.FromMilliseconds(310)));

        // Right up to the last folder the cap allows.
        Assert.Equal(
            OutlookComSession.SweepWalkVerdict.Visit,
            OutlookComSession.DecideSweepWalk(
                OutlookComSession.MaxScopedSweepFolders - 1, depth: 3, elapsed: Budget - TimeSpan.FromMilliseconds(1)));
    }

    [Fact]
    public void FolderCap_StopsTheWalkAtTheCap()
    {
        Assert.Equal(
            OutlookComSession.SweepWalkVerdict.FolderCap,
            OutlookComSession.DecideSweepWalk(OutlookComSession.MaxScopedSweepFolders, depth: 1, elapsed: TimeSpan.Zero));
    }

    [Fact]
    public void DepthLimit_MatchesTheFullTreeWalkGuardExactly()
    {
        // FolderWalkDepthGuard = 64 on CollectFolders, with the same `depth > guard`
        // comparison - a walk that recurses in C# cannot survive a cyclic or pathological
        // tree, and StackOverflowException is uncatchable: it kills the server process.
        Assert.Equal(
            OutlookComSession.SweepWalkVerdict.Visit,
            OutlookComSession.DecideSweepWalk(foldersVisited: 0, depth: 64, elapsed: TimeSpan.Zero));
        Assert.Equal(
            OutlookComSession.SweepWalkVerdict.DepthLimit,
            OutlookComSession.DecideSweepWalk(foldersVisited: 0, depth: 65, elapsed: TimeSpan.Zero));

        // Note the guard is deliberately independent of the visit cap: at cap 40 a walk
        // can never legitimately REACH depth 65, so this is a structural backstop that
        // still holds if the cap is ever raised or bypassed. It must not be deleted as
        // "unreachable".
        Assert.Equal(
            OutlookComSession.SweepWalkVerdict.DepthLimit,
            OutlookComSession.DecideSweepWalk(foldersVisited: 0, depth: 4_000, elapsed: TimeSpan.Zero));
    }

    [Fact]
    public void TimeBudget_FiresOnlyAfterTheBudgetIsPassed()
    {
        // Same boundary as ExhaustiveScanState.CheckDeadline (`Clock.Elapsed <= Budget`
        // is still inside): exactly at the budget is not yet spent.
        Assert.Equal(
            OutlookComSession.SweepWalkVerdict.Visit,
            OutlookComSession.DecideSweepWalk(foldersVisited: 1, depth: 1, elapsed: Budget));
        Assert.Equal(
            OutlookComSession.SweepWalkVerdict.TimeBudget,
            OutlookComSession.DecideSweepWalk(foldersVisited: 1, depth: 1, elapsed: Budget + TimeSpan.FromMilliseconds(1)));
    }

    [Fact]
    public void Precedence_IsStable_WhenSeveralBoundsApplyAtOnce()
    {
        // Most global bound wins, so the reason reported to the agent does not depend on
        // which folder happened to trip first.
        Assert.Equal(
            OutlookComSession.SweepWalkVerdict.TimeBudget,
            OutlookComSession.DecideSweepWalk(
                OutlookComSession.MaxScopedSweepFolders, depth: 900, elapsed: Budget + TimeSpan.FromSeconds(9)));
        Assert.Equal(
            OutlookComSession.SweepWalkVerdict.DepthLimit,
            OutlookComSession.DecideSweepWalk(OutlookComSession.MaxScopedSweepFolders, depth: 900, elapsed: TimeSpan.Zero));
    }

    [Fact]
    public void VisitCap_BoundsAWideSubtreeOfNonMailFolders()
    {
        // THE defect, as a walk: 500 Calendar/Contacts/Tasks folders under the scoped
        // root. Not one of them is swept, so a cap counting SWEPT folders never fired and
        // all 501 were visited - each paying its COM round trips. Counting VISITS is what
        // makes the existing cap of 40 actually bound the walk.
        WalkModel root = new WalkModel(isMail: false, Children(500, isMail: false));

        WalkResult result = Walk(root, millisecondsPerFolder: 0);

        Assert.Equal(OutlookComSession.MaxScopedSweepFolders, result.Visited);
        Assert.Equal(0, result.Swept);
        Assert.True(result.FolderCapReached);
        Assert.False(result.TimeBudgetExceeded);
        Assert.False(result.DepthLimitReached);
    }

    [Fact]
    public void VisitCap_BoundsAWideSubtreeOfUnreadableFolders()
    {
        // The other uncounted class: a folder whose table will not open is tallied as
        // FAILED, never as swept, so it was equally free under the old cap.
        WalkModel root = new WalkModel(isMail: true, Children(500, isMail: true, readable: false));

        WalkResult result = Walk(root, millisecondsPerFolder: 0);

        Assert.Equal(OutlookComSession.MaxScopedSweepFolders, result.Visited);
        Assert.Equal(1, result.Swept); // only the root itself
        Assert.True(result.FolderCapReached);
    }

    [Fact]
    public void TimeBudget_StopsAWalkWhoseFoldersAreSlow()
    {
        // The case the folder cap alone cannot bound: 40 visits are affordable at ~10 ms
        // per folder, but not when every folder costs 100 ms (a slow delegate/online
        // store). The clock is checked PER FOLDER - a folder holding no fresh mail never
        // reaches an inner loop, so per-folder checking is the only thing that bounds it.
        WalkModel root = new WalkModel(isMail: true, Children(500, isMail: true));

        WalkResult result = Walk(root, millisecondsPerFolder: 100);

        // Refusal happens once elapsed has PASSED 2000 ms: visits 1..21 cost 0..2000 ms.
        Assert.Equal(21, result.Visited);
        Assert.True(result.TimeBudgetExceeded);
        Assert.False(result.FolderCapReached);
    }

    [Fact]
    public void SmallRealTree_IsWalkedWholeAndTripsNothing()
    {
        // The everyday shape a folder-scoped search actually has: a folder with a few
        // levels of subfolders. No bound may fire, or the sweep would report a coverage
        // hole that does not exist.
        WalkModel root = new WalkModel(
            isMail: true,
            new WalkModel(isMail: true, Children(3, isMail: true)),
            new WalkModel(isMail: false, Children(2, isMail: true)),
            new WalkModel(isMail: true));

        WalkResult result = Walk(root, millisecondsPerFolder: 10);

        Assert.Equal(9, result.Visited);
        Assert.Equal(8, result.Swept); // every visited folder except the non-mail one
        Assert.False(result.FolderCapReached);
        Assert.False(result.TimeBudgetExceeded);
        Assert.False(result.DepthLimitReached);
    }

    // --- a model of the shipped walk, driven entirely by the pure decision ------------

    private static WalkModel[] Children(int count, bool isMail, bool readable = true)
    {
        WalkModel[] children = new WalkModel[count];
        for (int i = 0; i < count; i++)
        {
            children[i] = new WalkModel(isMail, readable, Array.Empty<WalkModel>());
        }

        return children;
    }

    private sealed class WalkModel
    {
        internal WalkModel(bool isMail, params WalkModel[] children)
            : this(isMail, readable: true, children)
        {
        }

        internal WalkModel(bool isMail, bool readable, WalkModel[] children)
        {
            IsMail = isMail;
            Readable = readable;
            Children = children;
        }

        /// <summary>DefaultItemType == 0. Only these are swept; the rest are still walked.</summary>
        internal bool IsMail { get; }

        /// <summary>False models a folder whose table will not open (tallied as failed).</summary>
        internal bool Readable { get; }

        internal IReadOnlyList<WalkModel> Children { get; }
    }

    private sealed class WalkResult
    {
        internal int Visited { get; set; }

        internal int Swept { get; set; }

        internal bool FolderCapReached { get; set; }

        internal bool DepthLimitReached { get; set; }

        internal bool TimeBudgetExceeded { get; set; }

        internal bool Latch(OutlookComSession.SweepWalkVerdict verdict)
        {
            switch (verdict)
            {
                case OutlookComSession.SweepWalkVerdict.TimeBudget:
                    TimeBudgetExceeded = true;
                    return false;
                case OutlookComSession.SweepWalkVerdict.DepthLimit:
                    DepthLimitReached = true;
                    return false;
                case OutlookComSession.SweepWalkVerdict.FolderCap:
                    FolderCapReached = true;
                    return false;
                default:
                    return true;
            }
        }
    }

    /// <summary>
    /// Mirrors SweepFolderTree: check the bounds per folder, count the VISIT, sweep only
    /// mail folders, then recurse. Wall clock is modelled as a fixed cost per visited
    /// folder so the budget is deterministic.
    /// </summary>
    private static WalkResult Walk(WalkModel root, int millisecondsPerFolder)
    {
        WalkResult result = new WalkResult();
        Descend(root, depth: 0, millisecondsPerFolder, result);
        return result;
    }

    private static void Descend(WalkModel folder, int depth, int millisecondsPerFolder, WalkResult result)
    {
        TimeSpan elapsed = TimeSpan.FromMilliseconds((long)result.Visited * millisecondsPerFolder);
        if (!result.Latch(OutlookComSession.DecideSweepWalk(result.Visited, depth, elapsed)))
        {
            return;
        }

        result.Visited++;
        if (folder.IsMail && folder.Readable)
        {
            result.Swept++;
        }

        foreach (WalkModel child in folder.Children)
        {
            // The shipped walk peeks before fetching a child, because reaching one costs
            // two COM round trips even when it is about to be refused.
            TimeSpan now = TimeSpan.FromMilliseconds((long)result.Visited * millisecondsPerFolder);
            if (!result.Latch(OutlookComSession.DecideSweepWalk(result.Visited, depth + 1, now)))
            {
                return;
            }

            Descend(child, depth + 1, millisecondsPerFolder, result);
        }
    }
}
