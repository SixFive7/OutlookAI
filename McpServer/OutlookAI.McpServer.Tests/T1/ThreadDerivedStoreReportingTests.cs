using System.Reflection;

using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// What <c>thread</c> says about a store it DERIVED, now that deriving one no longer narrows
/// anything (C5's behavioural half, 2026-08-24).
/// <para>
/// The behaviour change left <c>scopeStoreDerived</c> structurally unreachable under its old
/// meaning - "the scope applied here was derived rather than asked for" - because a scope now
/// exists only when the caller named a store, and a named store is not derived. The field was
/// re-pointed at the fact that survived: a store WAS derived from the referenced hit and
/// deliberately not applied. That payload half is pinned end to end in
/// <see cref="ThreadDerivedStoreScopeTests"/>, which owns the stand-in profile; this file
/// holds the half that needs no fixture at all - what the ADVICE may and may not say about it.
/// </para>
/// <para>
/// The rule the two tests below defend is one rule: a field that reports a narrowing which did
/// not happen has no remedy attached to it, and the sentence that DID branch on it must lose
/// the branch rather than keep it for a state the code can no longer be in. Advice for an
/// unreachable state is worse than none - a caller who follows it watches the warning stay.
/// </para>
/// </summary>
public sealed class ThreadDerivedStoreReportingTests
{
    private const string AliceStore = "alice@example.com";

    private const string BobStore = "bob@example.com";

    /// <summary>
    /// The <c>unqueried_store</c> sentence names ONE remedy, and it is the one that works.
    /// <para>
    /// It named two until the derived scope went away: drop the <c>store</c> you passed, or
    /// pass <c>conversation_id</c> beside <c>id</c> so that no store is derived. The second is
    /// now advice about an argument that changes nothing, because only a caller-chosen scope
    /// can raise this code at all - <c>live.storesNotQueried</c> is computed from the store the
    /// index query was narrowed to, and nothing else narrows it.
    /// </para>
    /// </summary>
    [Fact]
    public void TheNarrowedLookupSentence_OffersOnlyTheRemedyThatStillWorks()
    {
        string line = Assert.Single(MailService.DescribeThreadCoverage(
            Unqueried(), FreshMerge.FreshnessPartial, AliceStore, scopeWidened: false, top: 50)!);

        Assert.Contains("without store", line, StringComparison.Ordinal);
        Assert.DoesNotContain("conversation_id", line, StringComparison.Ordinal);
        Assert.DoesNotContain("DERIVED", line, StringComparison.Ordinal);
    }

    /// <summary>
    /// And no sentence anywhere selects on a derived store, which is asserted against the
    /// SIGNATURE rather than against one call: the flag was a parameter, so a branch could only
    /// come back by the parameter coming back. A derived store costs the caller no members and
    /// leaves nothing to do about it, so the field is informational and the advice list - which
    /// exists to name partial coverage - stays out of it.
    /// </summary>
    [Fact]
    public void TheCoverageAdvice_TakesNoDerivedStoreFlagAtAll()
    {
        MethodInfo method = typeof(MailService).GetMethod(
            nameof(MailService.DescribeThreadCoverage),
            BindingFlags.Public | BindingFlags.Static)
            ?? throw new InvalidOperationException(
                "MailService.DescribeThreadCoverage is gone or renamed; it is the one place a thread lookup's "
                + "partial coverage is narrated, so move this test with it.");

        Assert.Equal(
            new[] { "live", "freshness", "store", "scopeWidened", "top" },
            method.GetParameters().Select(p => p.Name).ToArray());
    }

    /// <summary>A walk of Alice's store with Bob's named as the one nobody asked about.</summary>
    private static ThreadLiveInfo Unqueried()
    {
        return new ThreadLiveInfo
        {
            Performed = true,
            MembersWalked = 3,
            MembersAdded = 3,
            AnchorStore = AliceStore,
            StoresNotQueried = new[] { BobStore },
            CoverageGaps = new[] { FreshMerge.ThreadGapUnqueriedStore },
        };
    }
}
