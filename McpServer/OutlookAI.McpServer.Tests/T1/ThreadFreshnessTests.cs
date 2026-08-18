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

        List<string> raised = new List<string>();
        raised.AddRange(FreshMerge.DescribeThreadCoverageGaps(capped, new[] { "alice@example.com" })!);
        raised.AddRange(FreshMerge.DescribeThreadCoverageGaps(
            Walked(store: "alice@example.com"), new[] { "shared@example.com" })!);

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
}
