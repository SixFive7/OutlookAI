using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 shapes for the exhaustive-scan DASL builder (v3.MD section 0.6 Phase 3 /
/// section 12): ci_phrasematch only in Restrict/GetTable shapes, LIKE fallback,
/// prefix stems via LIKE in both engines, UTC date literals and quote escaping.
/// <para>
/// THE MAIL-ONLY CLAUSE IS GONE (gap B3, maintainer decision 2026-08-18). Every filter
/// used to open with <c>PR_MESSAGE_CLASS like 'IPM.Note%'</c>, so the one search mode a
/// caller reaches for BECAUSE completeness matters was the only one that could not find a
/// bounce report, a read receipt, a meeting request or a post. The assertions that changed
/// are marked where they changed; what replaces the clause is nothing at all, except in the
/// one call that would otherwise emit an empty restriction.
/// </para>
/// </summary>
public sealed class ExhaustiveDaslFilterTests
{
    private const string Subject = "\"urn:schemas:httpmail:subject\"";
    private const string Body = "\"urn:schemas:httpmail:textdescription\"";
    private const string DateReceived = "\"urn:schemas:httpmail:datereceived\"";
    private const string MessageClass = "\"http://schemas.microsoft.com/mapi/proptag/0x001A001E\"";

    [Fact]
    public void CiEngine_SingleTerm_SubjectOrBodyPhraseMatch()
    {
        string filter = ExhaustiveDaslFilter.Build(new[] { "factuur" }, null, null, ExhaustiveEngine.CiPhraseMatch);

        Assert.StartsWith("@SQL=", filter, StringComparison.Ordinal);
        // CHANGED BY B3: this used to assert the presence of the IPM.Note clause.
        Assert.DoesNotContain("IPM.Note", filter, StringComparison.Ordinal);
        Assert.Contains(Subject + " ci_phrasematch 'factuur' OR " + Body + " ci_phrasematch 'factuur'", filter, StringComparison.Ordinal);
        Assert.DoesNotContain(" like '%factuur%'", filter, StringComparison.Ordinal);
    }

    [Fact]
    public void LikeEngine_SingleTerm_SubjectOrBodySubstring()
    {
        string filter = ExhaustiveDaslFilter.Build(new[] { "factuur" }, null, null, ExhaustiveEngine.Like);

        Assert.Contains(Subject + " like '%factuur%' OR " + Body + " like '%factuur%'", filter, StringComparison.Ordinal);
        Assert.DoesNotContain("ci_phrasematch", filter, StringComparison.Ordinal);
    }

    [Fact]
    public void MultipleTerms_AreAnded()
    {
        string filter = ExhaustiveDaslFilter.Build(new[] { "alpha", "beta" }, null, null, ExhaustiveEngine.CiPhraseMatch);

        Assert.Contains("ci_phrasematch 'alpha'", filter, StringComparison.Ordinal);
        Assert.Contains("ci_phrasematch 'beta'", filter, StringComparison.Ordinal);
        // CHANGED BY B3: two parenthesized term clauses joined with AND, where there used
        // to be three because the message-class filter always led.
        Assert.Equal(1, CountOccurrences(filter, ") AND ("));
    }

    [Fact]
    public void MultipleTerms_AndAcrossSubjectAndBody_NotInsideOneProperty()
    {
        // Tier parity with the index builder (soak fix 13): each term gets its own
        // subject-OR-body clause, so a mail with one term only in the subject and the
        // other only in the body matches. This tier already had the right shape - the
        // pin exists so it stays that way.
        foreach (ExhaustiveEngine engine in new[] { ExhaustiveEngine.CiPhraseMatch, ExhaustiveEngine.Like })
        {
            string filter = ExhaustiveDaslFilter.Build(new[] { "balans", "energie" }, null, null, engine);
            string op = engine == ExhaustiveEngine.CiPhraseMatch ? " ci_phrasematch '" : " like '%";
            string close = engine == ExhaustiveEngine.CiPhraseMatch ? "'" : "%'";

            foreach (string term in new[] { "balans", "energie" })
            {
                Assert.Contains(
                    "(" + Subject + op + term + close + " OR " + Body + op + term + close + ")",
                    filter,
                    StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public void PrefixStem_UsesLikeInBothEngines()
    {
        string ci = ExhaustiveDaslFilter.Build(new[] { "fact*" }, null, null, ExhaustiveEngine.CiPhraseMatch);
        string like = ExhaustiveDaslFilter.Build(new[] { "fact*" }, null, null, ExhaustiveEngine.Like);

        foreach (string filter in new[] { ci, like })
        {
            Assert.Contains(Subject + " like '%fact%' OR " + Body + " like '%fact%'", filter, StringComparison.Ordinal);
            Assert.DoesNotContain("ci_phrasematch", filter, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void DateBounds_EmitUtcYearFirstDaslLiterals()
    {
        // The SINCE bound deliberately falls on day 1: Outlook parses DASL date literals
        // in the machine locale, so the month-first literal this used to emit was read
        // with day and month swapped on any day 12 or lower (measured 2026-08-18 - an
        // exhaustive search for 1-5 August returned only mail from January to June).
        // Culture-independence of these literals is pinned in DaslDateLiteralTests.
        DateTime since = new DateTime(2026, 7, 1, 8, 30, 0, DateTimeKind.Utc);
        DateTime before = new DateTime(2026, 7, 15, 0, 0, 0, DateTimeKind.Utc);
        string filter = ExhaustiveDaslFilter.Build(null, since, before, ExhaustiveEngine.Like);

        Assert.Contains("(" + DateReceived + " >= '2026-07-01 08:30:00')", filter, StringComparison.Ordinal);
        Assert.Contains("(" + DateReceived + " < '2026-07-15 00:00:00')", filter, StringComparison.Ordinal);
        Assert.DoesNotContain("07/01/2026", filter, StringComparison.Ordinal);
    }

    [Fact]
    public void InvertedDateWindow_Throws()
    {
        DateTime since = new DateTime(2026, 7, 15, 0, 0, 0, DateTimeKind.Utc);
        Assert.Throws<ArgumentException>(() =>
            ExhaustiveDaslFilter.Build(null, since, since, ExhaustiveEngine.Like));
    }

    /// <summary>
    /// The one place a message-class predicate survives, and it is not a filter: an
    /// unbounded folder scan ("show me this folder") has nothing else to restrict on, and
    /// <c>@SQL=</c> with no predicate is not a restriction Outlook accepts. PR_MESSAGE_CLASS
    /// is mandatory on every MAPI message, so <c>like '%'</c> over it excludes nothing.
    /// <para>
    /// CHANGED BY B3: this call used to emit the real <c>like 'IPM.Note%'</c> filter, which
    /// is why the shape needed replacing rather than deleting.
    /// </para>
    /// </summary>
    [Fact]
    public void NoTermsNoDates_EmitsAPredicateThatAdmitsEveryClass()
    {
        string filter = ExhaustiveDaslFilter.Build(null, null, null, ExhaustiveEngine.Like);

        Assert.Equal("@SQL=(" + MessageClass + " like '%')", filter);
        Assert.Equal("@SQL=(" + ExhaustiveDaslFilter.AdmitEveryClassClause + ")", filter);
        Assert.DoesNotContain("IPM.Note", filter, StringComparison.Ordinal);
    }

    /// <summary>
    /// And the corollary: as soon as there is anything to restrict on, no class predicate is
    /// emitted at all. A tautology that leaked into every filter would be a wasted predicate
    /// on every scan, and would read to the next person like the filter that was removed.
    /// </summary>
    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void AnyRealBound_MeansNoClassPredicateAtAll(bool withTerm, bool withDate)
    {
        string filter = ExhaustiveDaslFilter.Build(
            withTerm ? new[] { "factuur" } : null,
            withDate ? new DateTime(2026, 7, 1, 0, 0, 0, DateTimeKind.Utc) : (DateTime?)null,
            null,
            ExhaustiveEngine.Like);

        Assert.DoesNotContain(MessageClass, filter, StringComparison.Ordinal);
    }

    /// <summary>
    /// The classes the removed filter dropped, spelled out: every one of them is now
    /// admitted by every tier. Nothing in the emitted DASL mentions any of them, which is
    /// the point - admission is decided by not asking.
    /// </summary>
    [Fact]
    public void TheClassesTheOldFilterDropped_AreNoLongerNamedInAnyFilter()
    {
        foreach (ExhaustiveEngine engine in new[] { ExhaustiveEngine.CiPhraseMatch, ExhaustiveEngine.Like })
        {
            string filter = ExhaustiveDaslFilter.Build(new[] { "factuur" }, null, null, engine);
            foreach (string dropped in OutlookAI.Core.Mapi.MailItemAdmission.ClassesTheOldFiltersDropped)
            {
                Assert.DoesNotContain(dropped, filter, StringComparison.OrdinalIgnoreCase);
            }
        }
    }

    [Fact]
    public void SingleQuote_IsDoubledInBothEngines()
    {
        string ci = ExhaustiveDaslFilter.Build(new[] { "o'brien" }, null, null, ExhaustiveEngine.CiPhraseMatch);
        string like = ExhaustiveDaslFilter.Build(new[] { "o'brien" }, null, null, ExhaustiveEngine.Like);

        Assert.Contains("ci_phrasematch 'o''brien'", ci, StringComparison.Ordinal);
        Assert.Contains("like '%o''brien%'", like, StringComparison.Ordinal);
    }

    [Fact]
    public void Underscore_IsBracketEscapedInLike()
    {
        string like = ExhaustiveDaslFilter.Build(new[] { "a_b" }, null, null, ExhaustiveEngine.Like);
        Assert.Contains("like '%a[_]b%'", like, StringComparison.Ordinal);

        // In the ci engine an exact underscore term stays a phrase literal.
        string ci = ExhaustiveDaslFilter.Build(new[] { "a_b" }, null, null, ExhaustiveEngine.CiPhraseMatch);
        Assert.Contains("ci_phrasematch 'a_b'", ci, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("*")]
    [InlineData("fa*r")]
    [InlineData("a;b")]
    [InlineData("a\"b")]
    [InlineData("   ")]
    public void InvalidTerms_Throw(string term)
    {
        Assert.Throws<ArgumentException>(() =>
            ExhaustiveDaslFilter.Build(new[] { term }, null, null, ExhaustiveEngine.Like));
    }

    // ------------------------------------------- search_in scopes (D40, user 2026-07-26)

    [Fact]
    public void DefaultSearchIn_StaysSubjectOrBody()
    {
        // The exhaustive tier already covered subject+body - that predicate asymmetry is
        // why it found the SF-6 population the index tier missed. It must not regress.
        string implicitDefault = ExhaustiveDaslFilter.Build(new[] { "factuur" }, null, null, ExhaustiveEngine.CiPhraseMatch);
        string explicitDefault = ExhaustiveDaslFilter.Build(
            new[] { "factuur" }, null, null, ExhaustiveEngine.CiPhraseMatch, SearchIn.SubjectAndBody);

        Assert.Equal(implicitDefault, explicitDefault);
        Assert.Contains(Subject + " ci_phrasematch 'factuur' OR " + Body + " ci_phrasematch 'factuur'", implicitDefault, StringComparison.Ordinal);
    }

    [Fact]
    public void SubjectOnlyScope_DropsTheBodyClause_BothEngines()
    {
        string ci = ExhaustiveDaslFilter.Build(new[] { "factuur" }, null, null, ExhaustiveEngine.CiPhraseMatch, SearchIn.SubjectOnly);
        string like = ExhaustiveDaslFilter.Build(new[] { "factuur" }, null, null, ExhaustiveEngine.Like, SearchIn.SubjectOnly);

        Assert.Contains("(" + Subject + " ci_phrasematch 'factuur')", ci, StringComparison.Ordinal);
        Assert.Contains("(" + Subject + " like '%factuur%')", like, StringComparison.Ordinal);
        Assert.DoesNotContain(Body, ci, StringComparison.Ordinal);
        Assert.DoesNotContain(Body, like, StringComparison.Ordinal);
    }

    [Fact]
    public void BodyOnlyScope_DropsTheSubjectClause_BothEngines()
    {
        string ci = ExhaustiveDaslFilter.Build(new[] { "factuur" }, null, null, ExhaustiveEngine.CiPhraseMatch, SearchIn.BodyOnly);
        string like = ExhaustiveDaslFilter.Build(new[] { "factuur" }, null, null, ExhaustiveEngine.Like, SearchIn.BodyOnly);

        Assert.Contains("(" + Body + " ci_phrasematch 'factuur')", ci, StringComparison.Ordinal);
        Assert.Contains("(" + Body + " like '%factuur%')", like, StringComparison.Ordinal);
        Assert.DoesNotContain(Subject, ci, StringComparison.Ordinal);
        Assert.DoesNotContain(Subject, like, StringComparison.Ordinal);
    }

    [Fact]
    public void PrefixStem_HonorsSearchIn()
    {
        string subjectOnly = ExhaustiveDaslFilter.Build(new[] { "fact*" }, null, null, ExhaustiveEngine.CiPhraseMatch, SearchIn.SubjectOnly);
        string bodyOnly = ExhaustiveDaslFilter.Build(new[] { "fact*" }, null, null, ExhaustiveEngine.CiPhraseMatch, SearchIn.BodyOnly);

        // CHANGED BY B3: both used to be prefixed with the IPM.Note class clause.
        Assert.Equal("@SQL=(" + Subject + " like '%fact%')", subjectOnly);
        Assert.Equal("@SQL=(" + Body + " like '%fact%')", bodyOnly);
    }

    [Fact]
    public void ScopedTerms_StillAndTogether()
    {
        string filter = ExhaustiveDaslFilter.Build(
            new[] { "alpha", "beta" }, null, null, ExhaustiveEngine.CiPhraseMatch, SearchIn.SubjectOnly);

        Assert.Equal(1, CountOccurrences(filter, ") AND ("));
        Assert.Equal(2, CountOccurrences(filter, "ci_phrasematch"));
    }

    [Fact]
    public void UnknownSearchIn_Throws()
    {
        Assert.Throws<ArgumentException>(() =>
            ExhaustiveDaslFilter.Build(new[] { "factuur" }, null, null, ExhaustiveEngine.Like, (SearchIn)99));
    }

    private static int CountOccurrences(string haystack, string needle)
    {
        int count = 0;
        int index = 0;
        while ((index = haystack.IndexOf(needle, index, StringComparison.Ordinal)) >= 0)
        {
            count++;
            index += needle.Length;
        }

        return count;
    }
}
