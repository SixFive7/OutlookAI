using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 shapes for the exhaustive-scan DASL builder (v3.MD section 0.6 Phase 3 /
/// section 12): ci_phrasematch only in Restrict/GetTable shapes, LIKE fallback,
/// prefix stems via LIKE in both engines, UTC date literals, quote escaping, and the
/// always-present IPM.Note mail-only clause.
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
        Assert.Contains("(" + MessageClass + " like 'IPM.Note%')", filter, StringComparison.Ordinal);
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
        // Three parenthesized clauses (message class + 2 terms) joined with AND.
        Assert.Equal(2, CountOccurrences(filter, ") AND ("));
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
    public void DateBounds_EmitUtcDaslLiterals()
    {
        DateTime since = new DateTime(2026, 7, 1, 8, 30, 0, DateTimeKind.Utc);
        DateTime before = new DateTime(2026, 7, 15, 0, 0, 0, DateTimeKind.Utc);
        string filter = ExhaustiveDaslFilter.Build(null, since, before, ExhaustiveEngine.Like);

        Assert.Contains("(" + DateReceived + " >= '07/01/2026 08:30:00')", filter, StringComparison.Ordinal);
        Assert.Contains("(" + DateReceived + " < '07/15/2026 00:00:00')", filter, StringComparison.Ordinal);
    }

    [Fact]
    public void InvertedDateWindow_Throws()
    {
        DateTime since = new DateTime(2026, 7, 15, 0, 0, 0, DateTimeKind.Utc);
        Assert.Throws<ArgumentException>(() =>
            ExhaustiveDaslFilter.Build(null, since, since, ExhaustiveEngine.Like));
    }

    [Fact]
    public void NoTermsNoDates_StillEmitsMailOnlyClause()
    {
        string filter = ExhaustiveDaslFilter.Build(null, null, null, ExhaustiveEngine.Like);
        Assert.Equal("@SQL=(" + MessageClass + " like 'IPM.Note%')", filter);
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

    // ------------------------------------------- term scopes (D40, user 2026-07-26)

    [Fact]
    public void DefaultTermScope_StaysSubjectOrBody()
    {
        // The exhaustive tier already covered subject+body - that predicate asymmetry is
        // why it found the SF-6 population the index tier missed. It must not regress.
        string implicitDefault = ExhaustiveDaslFilter.Build(new[] { "factuur" }, null, null, ExhaustiveEngine.CiPhraseMatch);
        string explicitDefault = ExhaustiveDaslFilter.Build(
            new[] { "factuur" }, null, null, ExhaustiveEngine.CiPhraseMatch, TermScope.SubjectAndBody);

        Assert.Equal(implicitDefault, explicitDefault);
        Assert.Contains(Subject + " ci_phrasematch 'factuur' OR " + Body + " ci_phrasematch 'factuur'", implicitDefault, StringComparison.Ordinal);
    }

    [Fact]
    public void SubjectOnlyScope_DropsTheBodyClause_BothEngines()
    {
        string ci = ExhaustiveDaslFilter.Build(new[] { "factuur" }, null, null, ExhaustiveEngine.CiPhraseMatch, TermScope.SubjectOnly);
        string like = ExhaustiveDaslFilter.Build(new[] { "factuur" }, null, null, ExhaustiveEngine.Like, TermScope.SubjectOnly);

        Assert.Contains("(" + Subject + " ci_phrasematch 'factuur')", ci, StringComparison.Ordinal);
        Assert.Contains("(" + Subject + " like '%factuur%')", like, StringComparison.Ordinal);
        Assert.DoesNotContain(Body, ci, StringComparison.Ordinal);
        Assert.DoesNotContain(Body, like, StringComparison.Ordinal);
    }

    [Fact]
    public void BodyOnlyScope_DropsTheSubjectClause_BothEngines()
    {
        string ci = ExhaustiveDaslFilter.Build(new[] { "factuur" }, null, null, ExhaustiveEngine.CiPhraseMatch, TermScope.BodyOnly);
        string like = ExhaustiveDaslFilter.Build(new[] { "factuur" }, null, null, ExhaustiveEngine.Like, TermScope.BodyOnly);

        Assert.Contains("(" + Body + " ci_phrasematch 'factuur')", ci, StringComparison.Ordinal);
        Assert.Contains("(" + Body + " like '%factuur%')", like, StringComparison.Ordinal);
        Assert.DoesNotContain(Subject, ci, StringComparison.Ordinal);
        Assert.DoesNotContain(Subject, like, StringComparison.Ordinal);
    }

    [Fact]
    public void PrefixStem_HonorsTermScope()
    {
        string subjectOnly = ExhaustiveDaslFilter.Build(new[] { "fact*" }, null, null, ExhaustiveEngine.CiPhraseMatch, TermScope.SubjectOnly);
        string bodyOnly = ExhaustiveDaslFilter.Build(new[] { "fact*" }, null, null, ExhaustiveEngine.CiPhraseMatch, TermScope.BodyOnly);

        Assert.Equal("@SQL=(" + MessageClass + " like 'IPM.Note%') AND (" + Subject + " like '%fact%')", subjectOnly);
        Assert.Equal("@SQL=(" + MessageClass + " like 'IPM.Note%') AND (" + Body + " like '%fact%')", bodyOnly);
    }

    [Fact]
    public void ScopedTerms_StillAndTogether()
    {
        string filter = ExhaustiveDaslFilter.Build(
            new[] { "alpha", "beta" }, null, null, ExhaustiveEngine.CiPhraseMatch, TermScope.SubjectOnly);

        Assert.Equal(2, CountOccurrences(filter, ") AND ("));
        Assert.Equal(2, CountOccurrences(filter, "ci_phrasematch"));
    }

    [Fact]
    public void UnknownTermScope_Throws()
    {
        Assert.Throws<ArgumentException>(() =>
            ExhaustiveDaslFilter.Build(new[] { "factuur" }, null, null, ExhaustiveEngine.Like, (TermScope)99));
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
