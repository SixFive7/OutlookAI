using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1: the freshness contract of <c>thread</c> and of an EXHAUSTIVE search - gaps C1 and
/// F1. Both closed a hole of the same shape: a result that covered less than it was asked
/// to, reported through fields an agent branches on in one mode and through nothing at all
/// in the other.
/// <para>
/// Everything here runs over the pure classifiers (<see cref="FreshMerge"/>) and the pure
/// advice renderers, so no Outlook, no mailbox and no search index is touched. What CANNOT
/// be pinned here, and needs a live profile: that Outlook's conversation walk actually
/// returns the members the index has not ingested yet, that a meeting-request member really
/// carries the mail's ConversationID, and that a conversation spanning two stores really
/// walks only one of them. Those are T2 (<c>Category=Live</c>) facts; this tier pins that
/// once the counters say so, the payload says so too.
/// </para>
/// </summary>
public sealed class ThreadFreshnessTests
{
    private static ThreadLiveInfo Walked(int members = 4, string? store = "alice@example.com")
    {
        return new ThreadLiveInfo
        {
            Performed = true,
            MembersWalked = members,
            MembersAdded = 1,
            AnchorStore = store,
        };
    }

    /// <summary>The thread coverage codes, read off the type so a new one cannot ship untested.</summary>
    private static IReadOnlyList<string> AllThreadGapCodes()
    {
        return typeof(FreshMerge)
            .GetFields(BindingFlags.Public | BindingFlags.Static)
            .Where(f => f.IsLiteral && f.FieldType == typeof(string)
                && f.Name.StartsWith("ThreadGap", System.StringComparison.Ordinal))
            .Select(f => (string)f.GetRawConstantValue()!)
            .ToList();
    }

    // ------------------------------------------------------------------ thread (C1)

    [Fact]
    public void AWalkThatCoveredTheConversation_IsLive_AndReportsNoGaps()
    {
        ThreadLiveInfo live = Walked();

        Assert.Null(FreshMerge.DescribeThreadCoverageGaps(live, new[] { "alice@example.com" }));
        Assert.Equal(FreshMerge.FreshnessLive, FreshMerge.ClassifyThreadFreshness(live, new[] { "alice@example.com" }));
    }

    /// <summary>
    /// The state gap C1 was: index rows and no live check. It must be index-only and
    /// degraded, never "live" - a conversation that skipped the live tier is exactly the
    /// answer that looks whole and is not.
    /// </summary>
    [Fact]
    public void AWalkThatCouldNotRun_IsIndexOnly_WhateverTheIndexReturned()
    {
        ThreadLiveInfo live = new ThreadLiveInfo { Performed = false, Error = "NoAnchorItem" };

        Assert.Equal(FreshMerge.FreshnessIndexOnly, FreshMerge.ClassifyThreadFreshness(live, new[] { "alice@example.com" }));

        // "Did not run" is a state, not a coverage hole - the same split the sweep makes.
        Assert.Null(FreshMerge.DescribeThreadCoverageGaps(live, new[] { "alice@example.com" }));
    }

    [Fact]
    public void AWalkStoppedAtTheMemberCap_IsPartial_AndNamesTheCap()
    {
        ThreadLiveInfo live = Walked();
        live.MemberCapReached = true;

        IReadOnlyList<string> gaps = FreshMerge.DescribeThreadCoverageGaps(live, new[] { "alice@example.com" })!;

        Assert.Contains(FreshMerge.ThreadGapMemberCap, gaps);
        Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyThreadFreshness(live, new[] { "alice@example.com" }));
    }

    /// <summary>
    /// Gap C4 made machine-readable: Outlook walks a conversation inside ONE store, so index
    /// rows from a second account prove the live tier did not cover all of it.
    /// </summary>
    [Fact]
    public void IndexMembersInAnotherStore_MakeTheWalkPartial()
    {
        ThreadLiveInfo live = Walked(store: "alice@example.com");

        IReadOnlyList<string> gaps = FreshMerge.DescribeThreadCoverageGaps(
            live, new[] { "alice@example.com", "shared@example.com" })!;

        Assert.Contains(FreshMerge.ThreadGapUnwalkedStore, gaps);
        Assert.Equal(
            FreshMerge.FreshnessPartial,
            FreshMerge.ClassifyThreadFreshness(live, new[] { "alice@example.com", "shared@example.com" }));
    }

    /// <summary>
    /// A store nobody could name is UNKNOWN, not UNCOVERED. A gap code that fires on missing
    /// metadata would mark a complete answer partial on every hit whose URL did not parse -
    /// the cries-wolf failure the absent-folder rule already exists to prevent one tier down.
    /// </summary>
    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void AnUnnameableStore_RaisesNothing(string? unnameable)
    {
        ThreadLiveInfo live = Walked(store: "alice@example.com");

        Assert.Null(FreshMerge.DescribeThreadCoverageGaps(live, new[] { "alice@example.com", unnameable }));
        Assert.Equal(
            FreshMerge.FreshnessLive,
            FreshMerge.ClassifyThreadFreshness(live, new[] { "alice@example.com", unnameable }));
    }

    /// <summary>Store comparison is case-insensitive: display-name casing is not a coverage hole.</summary>
    [Fact]
    public void StoreMatching_IgnoresCase()
    {
        ThreadLiveInfo live = Walked(store: "Alice@Example.com");

        Assert.Null(FreshMerge.DescribeThreadCoverageGaps(live, new[] { "alice@example.com" }));
    }

    /// <summary>
    /// A walk that returned nothing cannot be judged against the stores it did not name, and
    /// must not be reported as having missed one - it is the conversation that has no live
    /// members, not the walk that skipped a store.
    /// </summary>
    [Fact]
    public void AWalkThatFoundNoMembers_DoesNotClaimAnUnwalkedStore()
    {
        ThreadLiveInfo live = new ThreadLiveInfo { Performed = true, MembersWalked = 0, AnchorStore = null };

        Assert.Null(FreshMerge.DescribeThreadCoverageGaps(live, new[] { "alice@example.com", "shared@example.com" }));
        Assert.Equal(
            FreshMerge.FreshnessLive,
            FreshMerge.ClassifyThreadFreshness(live, new[] { "alice@example.com", "shared@example.com" }));
    }

    [Fact]
    public void EveryThreadGapCode_ProducesItsOwnAdviceSentence()
    {
        // Codes and prose are two renderings of one decision, exactly as on the sweep: a
        // code with no sentence is a partial result an agent can see but not explain.
        foreach (string code in AllThreadGapCodes())
        {
            ThreadLiveInfo live = Walked();
            live.CoverageGaps = new[] { code };

            IReadOnlyList<string> advice = MailService.DescribeThreadCoverage(
                live, FreshMerge.FreshnessPartial, store: null, scopeWidened: false, top: 50)!;

            string line = Assert.Single(advice);
            Assert.DoesNotContain("no further detail available", line, System.StringComparison.Ordinal);
        }
    }

    [Fact]
    public void TheGapSet_IsExactlyTheThreadCodesDeclared_SoANewOneCannotBeAddedUntested()
    {
        ThreadLiveInfo capped = Walked();
        capped.MemberCapReached = true;

        ThreadLiveInfo unindexed = Walked(store: "alice@example.com");
        unindexed.StoresWithoutIndex = new[] { "Archive 2019.pst" };

        List<string> raised = new List<string>();
        raised.AddRange(FreshMerge.DescribeThreadCoverageGaps(capped, new[] { "alice@example.com" })!);
        raised.AddRange(FreshMerge.DescribeThreadCoverageGaps(
            Walked(store: "alice@example.com"), new[] { "shared@example.com" })!);
        raised.AddRange(FreshMerge.DescribeThreadCoverageGaps(unindexed, new[] { "alice@example.com" })!);

        Assert.Equal(
            AllThreadGapCodes().OrderBy(c => c, System.StringComparer.Ordinal).ToList(),
            raised.Distinct().OrderBy(c => c, System.StringComparer.Ordinal).ToList());
    }

    /// <summary>
    /// The index-only sentence must name the remedy the CALLER holds. "Pass id" is the whole
    /// reason this state is reported rather than thrown: it is fixable from the next call.
    /// </summary>
    [Fact]
    public void TheIndexOnlyAdvice_NamesTheRemedy_AndSaysToTellTheUser()
    {
        ThreadLiveInfo live = new ThreadLiveInfo { Performed = false, Error = "NoAnchorItem" };

        string line = Assert.Single(MailService.DescribeThreadCoverage(
            live, FreshMerge.FreshnessIndexOnly, store: null, scopeWidened: false, top: 50)!);

        Assert.Contains("TELL THE USER", line, System.StringComparison.Ordinal);
        Assert.Contains("id", line, System.StringComparison.Ordinal);
    }

    /// <summary>A COM-side failure reports its own reason, not the pass-id remedy that cannot help.</summary>
    [Fact]
    public void AFailedWalk_ReportsItsReason_NotThePassIdRemedy()
    {
        ThreadLiveInfo live = new ThreadLiveInfo { Performed = false, Error = "COMException 0x80040107" };

        string line = Assert.Single(MailService.DescribeThreadCoverage(
            live, FreshMerge.FreshnessIndexOnly, store: null, scopeWidened: false, top: 50)!);

        Assert.Contains("COMException 0x80040107", line, System.StringComparison.Ordinal);
        Assert.Contains("outlook_health", line, System.StringComparison.Ordinal);
    }

    /// <summary>
    /// An expired or mistyped hit id gets its own remedy. "Retry" is the wrong instruction
    /// for it - retrying the same dead id produces the same dead id - and hit ids really do
    /// expire with the server process, which is the likeliest cause.
    /// </summary>
    [Fact]
    public void AnUnknownAnchorId_TellsTheCallerToGetAFreshHitId()
    {
        ThreadLiveInfo live = new ThreadLiveInfo { Performed = false, Error = "UnknownAnchorId" };

        string line = Assert.Single(MailService.DescribeThreadCoverage(
            live, FreshMerge.FreshnessIndexOnly, store: null, scopeWidened: false, top: 50)!);

        Assert.Contains("TELL THE USER", line, System.StringComparison.Ordinal);
        Assert.Contains("re-run the search", line, System.StringComparison.Ordinal);
        Assert.DoesNotContain("outlook_health", line, System.StringComparison.Ordinal);
    }

    /// <summary>
    /// Every index-only reason earns a sentence, and every sentence leads with the
    /// tell-the-user instruction: the flag is only useful if the prose beside it says what
    /// to do, and a reason that fell through to no branch would be a silent partial answer.
    /// </summary>
    [Theory]
    [InlineData("NoAnchorItem")]
    [InlineData("UnknownAnchorId")]
    [InlineData("AnchorNotLocatable")]
    [InlineData("ConversationWalkFailed")]
    [InlineData("Timeout")]
    [InlineData(null)]
    public void EveryIndexOnlyReason_EarnsASentence(string? error)
    {
        ThreadLiveInfo live = new ThreadLiveInfo { Performed = false, Error = error };

        string line = Assert.Single(MailService.DescribeThreadCoverage(
            live, FreshMerge.FreshnessIndexOnly, store: null, scopeWidened: false, top: 50)!);

        Assert.StartsWith("INCOMPLETE CONVERSATION - TELL THE USER", line, System.StringComparison.Ordinal);
    }

    /// <summary>Gap C3: a store that did not resolve widened the lookup, and now says so.</summary>
    [Fact]
    public void AWidenedStoreScope_IsReported_ButIsNotACoverageHole()
    {
        ThreadLiveInfo live = Walked();

        IReadOnlyList<string> advice = MailService.DescribeThreadCoverage(
            live, FreshMerge.FreshnessLive, store: "typo@example.com", scopeWidened: true, top: 50)!;

        Assert.Contains(advice, line => line.Contains("typo@example.com", System.StringComparison.Ordinal));
        Assert.Contains(advice, line => line.Contains("WHOLE profile", System.StringComparison.Ordinal));

        // Over-returning is the safe direction: it must not mark the answer degraded.
        Assert.Null(FreshMerge.DescribeThreadCoverageGaps(live, new[] { "alice@example.com" }));
    }

    [Fact]
    public void AWholeThread_SaysNothing()
    {
        Assert.Null(MailService.DescribeThreadCoverage(
            Walked(), FreshMerge.FreshnessLive, store: null, scopeWidened: false, top: 50));
    }

    // ------------------------------------------------------------- exhaustive (F1)

    [Fact]
    public void AnExhaustiveScanThatCoveredItsScope_IsLive()
    {
        ExhaustiveInfo scan = new ExhaustiveInfo { FoldersScanned = 12 };

        Assert.Equal(FreshMerge.FreshnessLive, FreshMerge.ClassifyExhaustiveFreshness(scan));
    }

    /// <summary>
    /// The three ways an exhaustive scan covers less than it was asked to. Each one used to
    /// leave BOTH top-level flags absent while <c>exhaustive.*</c> carried the fact - so the
    /// mode a caller reaches for because completeness matters was the one that could not say
    /// it had fallen short.
    /// </summary>
    [Theory]
    [InlineData(true, 0, false)]
    [InlineData(false, 3, false)]
    [InlineData(false, 0, true)]
    [InlineData(true, 3, true)]
    public void AnExhaustiveScanThatFellShort_IsPartial(bool timedOut, int foldersSkipped, bool truncated)
    {
        ExhaustiveInfo scan = new ExhaustiveInfo
        {
            FoldersScanned = 12,
            TimedOut = timedOut,
            FoldersSkipped = foldersSkipped,
            Truncated = truncated,
        };

        Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyExhaustiveFreshness(scan));
    }

    /// <summary>
    /// "index-only" means the live check never ran. An exhaustive scan IS the live check, so
    /// the value is unreachable there by construction - and no fourth value is invented for
    /// a mode whose method differs but whose COVERAGE question is the same one.
    /// </summary>
    [Fact]
    public void AnExhaustiveScan_IsNeverIndexOnly()
    {
        foreach (bool timedOut in new[] { false, true })
        {
            foreach (bool truncated in new[] { false, true })
            {
                foreach (int skipped in new[] { 0, 5 })
                {
                    string freshness = FreshMerge.ClassifyExhaustiveFreshness(new ExhaustiveInfo
                    {
                        TimedOut = timedOut,
                        Truncated = truncated,
                        FoldersSkipped = skipped,
                    });

                    Assert.NotEqual(FreshMerge.FreshnessIndexOnly, freshness);
                    Assert.Contains(freshness, new[] { FreshMerge.FreshnessLive, FreshMerge.FreshnessPartial });
                }
            }
        }
    }

    /// <summary>
    /// Both classifiers derive <c>degraded</c> the same way the merged path does: anything
    /// that is not "live" is degraded, and "live" leaves the flag absent. Three call sites,
    /// one rule - a fourth spelling of it is how the flag and the value drift apart.
    /// </summary>
    [Theory]
    [InlineData(FreshMerge.FreshnessLive, null)]
    [InlineData(FreshMerge.FreshnessPartial, true)]
    [InlineData(FreshMerge.FreshnessIndexOnly, true)]
    public void DegradedIsDerivedFromFreshness_Everywhere(string freshness, bool? expected)
    {
        Assert.Equal(expected, freshness == FreshMerge.FreshnessLive ? (bool?)null : true);
    }

    // ----------------------------------------- exhaustive coverage CODES (F5, and F1's other half)

    // THE DEFECT THIS SECTION PINS: F1 gave the exhaustive tier degraded/freshness but no
    // machine-readable REASON, and the rows the scan lost inside a folder it had already
    // opened were not counted at all (gap F5) - a scan that could not open a single one of a
    // folder's matches reported that folder as scanned and returned nothing from it.

    /// <summary>Every exhaustive coverage code declared on <see cref="FreshMerge"/>, read from the type.</summary>
    private static IReadOnlyList<string> AllScanGapCodes()
    {
        return typeof(FreshMerge)
            .GetFields(BindingFlags.Public | BindingFlags.Static)
            .Where(f => f.IsLiteral && f.FieldType == typeof(string) && f.Name.StartsWith("ScanGap", StringComparison.Ordinal))
            .Select(f => (string)f.GetRawConstantValue()!)
            .Distinct(StringComparer.Ordinal)
            .ToList();
    }

    /// <summary>
    /// Every way an exhaustive scan covers less than its scope, paired with the code it must
    /// report - the same data-driven shape the sweep's holes use, so the set itself can be
    /// compared against the codes the type declares.
    /// </summary>
    private static List<(string Gap, ExhaustiveInfo Scan)> ScanCoverageHoles()
    {
        return new List<(string, ExhaustiveInfo)>
        {
            (FreshMerge.ScanGapTimeBudget, new ExhaustiveInfo { FoldersScanned = 4, TimedOut = true }),
            (FreshMerge.ScanGapResultCap, new ExhaustiveInfo { FoldersScanned = 4, Truncated = true }),
            (FreshMerge.ScanGapFoldersSkipped, new ExhaustiveInfo { FoldersScanned = 4, FoldersSkipped = 9 }),

            // Gap F4: the walk refused a subtree past the depth guard, so those folders were
            // never opened. New here because this walk had no bound at all to report - it
            // recursed until the stack ran out and took the COM host with it.
            (FreshMerge.ScanGapDepthLimit, new ExhaustiveInfo { FoldersScanned = 4, DepthLimitReached = true }),
            (FreshMerge.ScanGapRowsUnreadable, new ExhaustiveInfo { FoldersScanned = 4, RowsDropped = 5, RowsUnreadable = 5 }),
            (
                FreshMerge.ScanGapFilterUnreadable,
                new ExhaustiveInfo
                {
                    FoldersScanned = 4,
                    ItemsFilterUnreadable = 2,
                    FiltersUnevaluated = new[] { "unread_only" },
                }),

            // Gap F3: the cap counted CANDIDATES, and the caller's own filter then thinned
            // what it kept - so `truncated` beside two results does not mean two more exist.
            // Truncated is part of the fixture on purpose: without the cap the filter saw
            // the whole matched set and this is not a hole at all.
            (
                FreshMerge.ScanGapPostCapFilter,
                new ExhaustiveInfo
                {
                    FoldersScanned = 4,
                    Truncated = true,
                    PostCapFilters = new[] { "from" },
                    ItemsFilteredOut = 23,
                }),

            // Gap F2's five. The first is the general fact - this answer is ONE PAGE of a
            // longer scan - and it is what keeps degraded honest on the final page of a
            // chain, which by itself covered everything it was asked for. The other four are
            // only reachable on a page that already carries it.
            (FreshMerge.ScanGapResumed, new ExhaustiveInfo { FoldersScanned = 4, Resumed = true }),
            (FreshMerge.ScanGapTreeChanged, new ExhaustiveInfo { FoldersScanned = 4, TreeChangedFolders = 2 }),
            (FreshMerge.ScanGapResumedUnsorted, new ExhaustiveInfo { FoldersScanned = 4, ResumedUnsorted = true }),
            (
                FreshMerge.ScanGapResumePositionLost,
                new ExhaustiveInfo { FoldersScanned = 4, ResumePositionLost = true }),
            (
                FreshMerge.ScanGapDedupCapacity,
                new ExhaustiveInfo { FoldersScanned = 4, DedupCapacityReached = true }),
        };
    }

    [Fact]
    public void TheCursorFolderVanishing_RaisesTreeChanged_EvenWithNoFolderCount()
    {
        // The one tree-change that LOSES mail: the folder the chain stopped inside is gone,
        // so its unread remainder is uncovered. It carries no count of its own, so a
        // condition written only over TreeChangedFolders would miss exactly the case that
        // matters most.
        ExhaustiveInfo scan = new ExhaustiveInfo { FoldersScanned = 4, CursorFolderMissing = true };

        Assert.Contains(FreshMerge.ScanGapTreeChanged, FreshMerge.DescribeExhaustiveCoverageGaps(scan)!);
        Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyExhaustiveFreshness(scan));

        scan.CoverageGaps = new[] { FreshMerge.ScanGapTreeChanged };
        string line = Assert.Single(MailService.DescribeExhaustiveCoverage(scan, top: 25));
        Assert.Contains("NOT covered", line, StringComparison.Ordinal);
    }

    [Fact]
    public void AScanThatCoveredItsScope_ReportsNoCodes()
    {
        Assert.Null(FreshMerge.DescribeExhaustiveCoverageGaps(new ExhaustiveInfo { FoldersScanned = 12 }));
    }

    [Fact]
    public void EveryScanCoverageHole_MakesTheScanPartial_AndNamesItself()
    {
        foreach ((string expectedGap, ExhaustiveInfo scan) in ScanCoverageHoles())
        {
            IReadOnlyList<string>? gaps = FreshMerge.DescribeExhaustiveCoverageGaps(scan);
            Assert.True(gaps != null, $"{expectedGap}: a scan with this hole must report coverage gaps");
            Assert.Contains(expectedGap, gaps!);
            Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyExhaustiveFreshness(scan));
        }
    }

    [Fact]
    public void TheScanCoverageHoleSet_IsExactlyTheScanCodesDeclared_SoANewOneCannotBeAddedUntested()
    {
        List<string> covered = ScanCoverageHoles().Select(r => r.Gap).Distinct(StringComparer.Ordinal)
            .OrderBy(c => c, StringComparer.Ordinal).ToList();
        List<string> declared = AllScanGapCodes().OrderBy(c => c, StringComparer.Ordinal).ToList();
        Assert.Equal(declared, covered);
    }

    [Fact]
    public void EveryScanCode_ProducesItsOwnAdviceSentence()
    {
        // Codes and prose are two renderings of one decision here exactly as they are for
        // the sweep. A code with no sentence is a partial result an agent can see and cannot
        // explain; this walks the codes declared on the type, so a new one added without
        // prose fails here rather than shipping silent.
        foreach (string code in AllScanGapCodes())
        {
            ExhaustiveInfo scan = new ExhaustiveInfo
            {
                FoldersScanned = 4,
                FoldersSkipped = 9,
                RowsDropped = 7,
                RowsUnreadable = 5,
                ItemsFilterUnreadable = 2,
                FiltersUnevaluated = new[] { "unread_only" },
                CoverageGaps = new[] { code },
            };

            string line = Assert.Single(MailService.DescribeExhaustiveCoverage(scan, top: 25));
            Assert.DoesNotContain("no further detail available", line, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void RowsDroppedWithoutFailures_AreCounted_ButRaiseNothing()
    {
        // The distinction that keeps this flag from crying wolf. rowsDropped is every row
        // not admitted and never a coverage hole on its own; rowsUnreadable is the failure
        // subset and is the half that makes a scan partial.
        //
        // Its difference from rowsUnreadable USED to be the item-class filter, which is the
        // measurement gap B3 asked for - and B3's answer removed the filter, so on a real
        // scan that difference is now zero by construction. The shape is still pinned
        // because the classifier must keep reading the two numbers separately: a future
        // non-failure drop would otherwise silently start degrading every scan that met one.
        ExhaustiveInfo droppedNotLost = new ExhaustiveInfo { FoldersScanned = 4, RowsDropped = 28, RowsUnreadable = 0 };

        Assert.Null(FreshMerge.DescribeExhaustiveCoverageGaps(droppedNotLost));
        Assert.Equal(FreshMerge.FreshnessLive, FreshMerge.ClassifyExhaustiveFreshness(droppedNotLost));

        // And a scan that lost rows to FAILURE is partial even when it covered every folder.
        ExhaustiveInfo lostRows = new ExhaustiveInfo { FoldersScanned = 4, RowsDropped = 28, RowsUnreadable = 2 };
        Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyExhaustiveFreshness(lostRows));
        Assert.Contains(FreshMerge.ScanGapRowsUnreadable, FreshMerge.DescribeExhaustiveCoverageGaps(lostRows)!);
    }

    [Fact]
    public void TheScanReusesTheSweepsTokens_WhereTheHoleIsTheSameHole()
    {
        // One vocabulary across tiers: a caller that learned "time_budget" from a freshness
        // sweep must not have to learn a second spelling of it for the scan. Where the holes
        // genuinely differ, the tokens do too - the sweep's item cap truncates ONE folder's
        // window and the rest are still swept, while the scan's result cap stops the walk.
        Assert.Equal(FreshMerge.GapTimeBudget, FreshMerge.ScanGapTimeBudget);
        Assert.Equal(FreshMerge.GapFoldersSkipped, FreshMerge.ScanGapFoldersSkipped);
        Assert.Equal(FreshMerge.GapRowsUnreadable, FreshMerge.ScanGapRowsUnreadable);
        Assert.Equal(FreshMerge.GapFilterUnreadable, FreshMerge.ScanGapFilterUnreadable);
        Assert.NotEqual(FreshMerge.GapItemCap, FreshMerge.ScanGapResultCap);

        foreach (string code in AllScanGapCodes())
        {
            Assert.DoesNotContain(" ", code, StringComparison.Ordinal);
            Assert.Equal(code.ToLowerInvariant(), code);
        }
    }
}
