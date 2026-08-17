using System;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The two search tiers agree on what a valid TERM is.
/// <para>
/// <c>WsSqlBuilder</c> (index) and <c>ExhaustiveDaslFilter</c> (COM) each held their own
/// <c>private const int MaxTermLength = 128</c>, undocumented on both sides. They answer the
/// same user query in two modes, so raising one alone would have made a term accepted by one
/// mode and rejected by the other, for the same input, with nothing to notice - and the
/// symptom would have been "exhaustive search says my query is invalid" long after the edit.
/// </para>
/// <para>
/// The limit is now declared once and derived by the other tier, and this test proves the
/// BEHAVIOUR rather than the equality of two fields: both builders are driven with a term at
/// the boundary and one past it.
/// </para>
/// </summary>
public sealed class SearchTermLimitParityTests
{
    private static string TermOfLength(int length) => new string('a', length);

    private static string BuildIndexStatement(string term)
    {
        return WsSqlBuilder.Build(new IndexQuery { Terms = new[] { term }, Top = 10 });
    }

    private static string BuildExhaustiveFilter(string term)
    {
        return ExhaustiveDaslFilter.Build(new[] { term }, null, null, ExhaustiveEngine.Like);
    }

    /// <summary>A term exactly at the limit is accepted by BOTH tiers.</summary>
    [Fact]
    public void ATermAtTheLimit_IsAcceptedByBothTiers()
    {
        string term = TermOfLength(WsSqlBuilder.MaxTermLength);

        string sql = BuildIndexStatement(term);
        string dasl = BuildExhaustiveFilter(term);

        Assert.Contains(term, sql, StringComparison.Ordinal);
        Assert.Contains(term, dasl, StringComparison.Ordinal);
    }

    /// <summary>
    /// A term one character past the limit is rejected by BOTH tiers. Break either constant
    /// and exactly one of these two assertions fails, naming the tier that drifted.
    /// </summary>
    [Fact]
    public void ATermPastTheLimit_IsRejectedByBothTiers()
    {
        string term = TermOfLength(WsSqlBuilder.MaxTermLength + 1);

        ArgumentException indexTier = Assert.Throws<ArgumentException>(() => BuildIndexStatement(term));
        ArgumentException comTier = Assert.Throws<ArgumentException>(() => BuildExhaustiveFilter(term));

        Assert.Contains("too long", indexTier.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("too long", comTier.Message, StringComparison.OrdinalIgnoreCase);
    }
}
