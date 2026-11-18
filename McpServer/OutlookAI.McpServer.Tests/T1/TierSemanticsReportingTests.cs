using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;

using Xunit;

using static OutlookAI.McpServer.Tests.AliceMailbox;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Three things a search did and never said: WHAT text each tier matched (gap B4), HOW it
/// matched it (gap B5), and that the freshness sweep it reports may have run seconds ago
/// rather than now (gap E3). Plus the last silent clamp on this surface, <c>snippet_chars</c>.
/// <para>
/// B4 - <c>search_in: "body"</c> named three different bodies. The index matches
/// <c>System.Search.Contents</c>, which is the body stream PLUS attachment text; the sweep
/// reads <c>MailItem.Body</c> through Outlook; the exhaustive scan restricts on
/// <c>urn:schemas:httpmail:textdescription</c>, the plain-text body property alone. Only the
/// README said so, and the narrowest of the three belongs to the mode a caller chooses BECAUSE
/// it is the complete one - so an agent switching to <c>exhaustive:true</c> to find a term the
/// sweep could not see inside an attachment was switching to the tier least able to find it.
/// </para>
/// <para>
/// B5 - the index matches whole words and the sweep matches substrings. It over-matches, which
/// is the safe direction for a freshness tier, but it is also why a just-arrived mail can
/// appear in a search and drop out of the identical search once the index catches up. An agent
/// that can read both fields can explain that instead of reporting a bug.
/// </para>
/// <para>
/// E3 - a sweep served from the ~10 s cache did its live check of Outlook up to that long ago,
/// so an arrival inside the window is in neither tier. <c>sweep.cached</c> and
/// <c>sweep.cacheAgeSeconds</c> have always been in the payload; they sat beside
/// <c>freshness: "live"</c>, which is the one word an agent reads as "nothing is missing".
/// </para>
/// <para>
/// Driven through the real <see cref="MailService"/> against a stand-in session and a stand-in
/// index client. No mailbox and no Windows Search index are touched.
/// </para>
/// </summary>
public sealed class TierSemanticsReportingTests
{
    // ======================================================= B4: what each tier read

    [Fact]
    public void TheThreeTiers_NameThreeDIFFERENTBodies()
    {
        // The row's whole claim in one assertion: if any two of these ever became equal the
        // field would be noise, and if they were dropped the difference would be invisible
        // again.
        Assert.Equal(3, new HashSet<string>(StringComparer.Ordinal)
        {
            FreshMerge.BodyScopeBodyAndAttachments,
            FreshMerge.BodyScopeItemBody,
            FreshMerge.BodyScopePlainTextBody,
        }.Count);
    }

    [Theory]
    [InlineData(SearchIn.SubjectAndBody)]
    [InlineData(SearchIn.BodyOnly)]
    public void AQueryThatMatchesAgainstABody_GetsTheTiersOwnAnswer(SearchIn searchIn)
    {
        Assert.Equal(
            FreshMerge.BodyScopeItemBody,
            FreshMerge.BodyTextScope(FreshMerge.BodyScopeItemBody, searchIn, hasTerms: true));
    }

    [Fact]
    public void ASubjectOnlySearch_AndOneWithNoTermsAtAll_SayNothing()
    {
        // The question does not arise, and a field that answered it anyway would report a
        // difference that cannot have cost this search anything. Same gate as B2's
        // attachmentTextCovered, deliberately: one rule, one place.
        Assert.Null(FreshMerge.BodyTextScope(
            FreshMerge.BodyScopeItemBody, SearchIn.SubjectOnly, hasTerms: true));
        Assert.Null(FreshMerge.BodyTextScope(
            FreshMerge.BodyScopeItemBody, SearchIn.SubjectAndBody, hasTerms: false));
    }

    [Fact]
    [Trait("TimeZone", "Machine")]
    public void AnOrdinarySearch_ReportsBothTiersBodyScopes()
    {
        using MailService service = Service(Index(), Sweep);

        SearchOutcome outcome = service.Search(Request());

        Assert.Equal(FreshMerge.BodyScopeBodyAndAttachments, outcome.Index!.BodyTextScope);
        Assert.Equal(FreshMerge.BodyScopeItemBody, outcome.Sweep!.BodyTextScope);
    }

    [Fact]
    [Trait("TimeZone", "Machine")]
    public void ASubjectOnlySearch_CarriesNeither()
    {
        using MailService service = Service(Index(), Sweep);

        SearchRequest request = Request();
        request.SearchIn = SearchIn.SubjectOnly;
        SearchOutcome outcome = service.Search(request);

        Assert.Null(outcome.Index!.BodyTextScope);
        Assert.Null(outcome.Sweep!.BodyTextScope);
    }

    [Fact]
    [Trait("TimeZone", "Machine")]
    public void AnExhaustiveSearch_SaysItReadThePlainTextBodyAndNoAttachment()
    {
        using MailService service = Service(Index(), Sweep, ScanOf);

        SearchOutcome outcome = service.Search(ExhaustiveRequest());

        Assert.Equal(FreshMerge.BodyScopePlainTextBody, outcome.Exhaustive!.BodyTextScope);
        Assert.False(outcome.Exhaustive.AttachmentTextCovered);
    }

    [Fact]
    public void TheExhaustiveSentence_SaysWhichTwoThingsItCannotSee_AndWhereToGoInstead()
    {
        string line = MailService.DescribeExhaustiveBodyTextGap(
            new ExhaustiveInfo { BodyTextScope = FreshMerge.BodyScopePlainTextBody })!;

        Assert.Contains("attachment", line, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("HTML", line, StringComparison.Ordinal);

        // The remedy is the opposite of the instinct: the complete-looking mode is the one
        // that cannot read attachment text, so the answer is to search WITHOUT it.
        Assert.Contains("WITHOUT exhaustive:true", line, StringComparison.Ordinal);
    }

    [Fact]
    public void TheExhaustiveSentence_IsSilentWhenNoTermMetABody()
    {
        // Generated FROM the field, so it cannot claim a hole the payload does not carry.
        Assert.Null(MailService.DescribeExhaustiveBodyTextGap(new ExhaustiveInfo()));
    }

    [Fact]
    [Trait("TimeZone", "Machine")]
    public void AnExhaustiveSearch_CarriesTheSentenceItsFieldImplies()
    {
        using MailService service = Service(Index(), Sweep, ScanOf);

        SearchOutcome outcome = service.Search(ExhaustiveRequest());

        Assert.Contains(
            outcome.Advice!,
            a => a.Contains("WITHOUT exhaustive:true", StringComparison.Ordinal));
    }

    // ======================================================= B5: how each tier matched

    [Fact]
    [Trait("TimeZone", "Machine")]
    public void TheIndexMatchesWholeWords_AndTheSweepMatchesSubstrings()
    {
        using MailService service = Service(Index(), Sweep);

        SearchOutcome outcome = service.Search(Request());

        Assert.Equal(FreshMerge.TermMatchWholeWord, outcome.Index!.TermMatch);
        Assert.Equal(FreshMerge.TermMatchSubstring, outcome.Sweep!.TermMatch);
    }

    [Fact]
    [Trait("TimeZone", "Machine")]
    public void ATermlessSearch_ClaimsNoMatchingSemanticsAtAll()
    {
        using MailService service = Service(Index(), Sweep);

        SearchRequest request = Request();
        request.Query = null;
        request.From = "bob@example.com";
        SearchOutcome outcome = service.Search(request);

        Assert.Null(outcome.Index!.TermMatch);
        Assert.Null(outcome.Sweep!.TermMatch);
    }

    [Theory]
    [InlineData("ci_phrasematch", FreshMerge.TermMatchWholeWord)]
    [InlineData("like", FreshMerge.TermMatchSubstring)]
    [InlineData("ci_phrasematch+like", FreshMerge.TermMatchSubstring)]
    public void TheExhaustiveAnswerIsReadOffTheEngineItActuallyUsed(string engine, string expected)
    {
        // The mixed case answers "substring", which is the honest reading of a result set
        // that is broader than whole-word somewhere and cannot say where.
        Assert.Equal(expected, FreshMerge.ExhaustiveTermMatch(engine, hasTerms: true));
    }

    [Fact]
    public void TheExhaustiveAnswerIsAbsentWithoutTerms()
    {
        Assert.Null(FreshMerge.ExhaustiveTermMatch("ci_phrasematch", hasTerms: false));
        Assert.Null(FreshMerge.ExhaustiveTermMatch(null, hasTerms: false));
    }

    [Fact]
    [Trait("TimeZone", "Machine")]
    public void AnExhaustiveSearch_ReportsItsOwnMatchingInTheSameVocabulary()
    {
        using MailService service = Service(Index(), Sweep, ScanOf);

        SearchOutcome outcome = service.Search(ExhaustiveRequest());

        Assert.Equal(FreshMerge.TermMatchWholeWord, outcome.Exhaustive!.TermMatch);
    }

    // ================================================== E3: the sweep that already ran

    [Fact]
    public void ACachedSweep_RaisesItsCode_AndMakesTheSearchPartial()
    {
        SweepInfo sweep = new SweepInfo
        {
            Performed = true,
            FoldersSwept = 4,
            Cached = true,
            CacheAgeSeconds = 7.4,
        };

        Assert.Contains(FreshMerge.GapCachedSweep, FreshMerge.DescribeCoverageGaps(sweep)!);
        Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyFreshness(sweep));
    }

    [Fact]
    public void ALiveSweepThatCoveredItsScope_IsStillFullyLive()
    {
        // The cry-wolf guard: this must fire for a repeat inside the window and for nothing
        // else, or it devalues every other code on the block.
        SweepInfo sweep = new SweepInfo { Performed = true, FoldersSwept = 4 };

        Assert.Null(FreshMerge.DescribeCoverageGaps(sweep));
        Assert.Equal(FreshMerge.FreshnessLive, FreshMerge.ClassifyFreshness(sweep));
    }

    [Fact]
    public void TheCachedSentence_QuotesTheAge_TheTtl_AndSaysTheHoleClosesItself()
    {
        SweepInfo sweep = new SweepInfo
        {
            Performed = true,
            FoldersSwept = 4,
            Cached = true,
            CacheAgeSeconds = 7.4,
            CoverageGaps = new[] { FreshMerge.GapCachedSweep },
        };

        string line = Assert.Single(MailService.DescribeSweepCoverage(sweep, "12 minutes", folderScoped: false));

        Assert.Contains("7.4 s ago", line, StringComparison.Ordinal);
        Assert.Contains("10 s", line, StringComparison.Ordinal);
        Assert.Contains("closes by itself", line, StringComparison.Ordinal);
    }

    [Fact]
    [Trait("TimeZone", "Machine")]
    public void TheSecondOfTwoRapidSearches_SaysItsSweepWasNotRunAgain()
    {
        // End to end, and the state the cache exists to produce: two searches inside the TTL,
        // the second served from the first's sweep. Before this the second answered
        // freshness "live" with nothing but a cache age to contradict it.
        //
        // STORE-SCOPED deliberately, and it was not an arbitrary fixture choice: an unscoped
        // search used to key the cache on a window base recomputed from the wall clock every
        // call, so no two unscoped searches shared a key and none of them could ever hit. That
        // defect is fixed (the unscoped key is the profile frontier again) and pinned in
        // T1/SweepCacheKeyTests; this test pins the REPORTING, which is the same either way.
        using MailService service = Service(Index(), Sweep);

        SearchOutcome first = service.Search(ScopedRequest());
        SearchOutcome second = service.Search(ScopedRequest());

        Assert.Null(first.Sweep!.Cached);
        Assert.Equal(FreshMerge.FreshnessLive, first.Freshness);

        Assert.True(second.Sweep!.Cached);
        Assert.Contains(FreshMerge.GapCachedSweep, second.Sweep.CoverageGaps!);
        Assert.Equal(FreshMerge.FreshnessPartial, second.Freshness);
        Assert.True(second.Degraded);
        Assert.Contains(
            second.Advice!,
            a => a.Contains("served from cache", StringComparison.Ordinal));
    }

    // ============================================== the last silent clamp on this surface

    [Fact]
    [Trait("TimeZone", "Machine")]
    public void ASnippetLengthAboveTheCap_IsReported()
    {
        using MailService service = Service(Index(), Sweep);

        SearchRequest request = Request();
        request.SnippetChars = 5000;
        SearchOutcome outcome = service.Search(request);

        Assert.Contains(
            outcome.Advice!,
            a => a.Contains("snippet_chars=5000 was reduced to 1000", StringComparison.Ordinal));
    }

    [Fact]
    [Trait("TimeZone", "Machine")]
    public void ANegativeSnippetLength_IsReportedToo_BecauseItTurnsSnippetsOff()
    {
        using MailService service = Service(Index(), Sweep);

        SearchRequest request = Request();
        request.SnippetChars = -20;
        SearchOutcome outcome = service.Search(request);

        Assert.Contains(
            outcome.Advice!,
            a => a.Contains("snippet_chars=-20 was raised to 0", StringComparison.Ordinal));
    }

    [Fact]
    [Trait("TimeZone", "Machine")]
    public void AnExhaustiveSearch_SaysItDroppedTheSnippetLengthAltogether()
    {
        // Found by the C5 asymmetry scan: the ordinary search honours and reports this
        // argument, the exhaustive one drops it whole - and hits with no snippet look
        // identical to mail with an empty body.
        using MailService service = Service(Index(), Sweep, ScanOf);

        SearchRequest request = ExhaustiveRequest();
        request.SnippetChars = 200;
        SearchOutcome outcome = service.Search(request);

        Assert.Contains(
            outcome.Advice!,
            a => a.Contains("snippet_chars=200 was ignored", StringComparison.Ordinal));
    }

    [Fact]
    [Trait("TimeZone", "Machine")]
    public void ASnippetLengthInsideTheRange_SaysNothing()
    {
        using MailService service = Service(Index(), Sweep);

        SearchRequest request = Request();
        request.SnippetChars = 200;
        SearchOutcome outcome = service.Search(request);

        Assert.DoesNotContain(
            outcome.Advice ?? Array.Empty<string>(),
            a => a.Contains("snippet_chars", StringComparison.Ordinal));
    }

    // =============================================================== fixtures

    private static SearchRequest Request()
    {
        return new SearchRequest { Query = "test", Top = 25, SnippetChars = 0 };
    }

    private static SearchRequest ScopedRequest()
    {
        return new SearchRequest { Query = "test", StoreId = TestFolders.StoreId(Store), Top = 25, SnippetChars = 0 };
    }

    /// <summary>A scan that covered its scope with the whole-word engine.</summary>
    private static ComExhaustiveResult ScanOf(string? store)
    {
        return new ComExhaustiveResult(
            new[] { Mail("EX1") }, foldersScanned: 3, foldersSkipped: 0, engine: "ci_phrasematch",
            instantSearchEnabled: true, truncated: false, timedOut: false);
    }
}
