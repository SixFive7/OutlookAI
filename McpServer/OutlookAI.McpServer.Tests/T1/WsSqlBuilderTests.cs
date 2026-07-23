using OutlookAI.Core.IndexSearch;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 shape tests for the WS-SQL builder, including the v3.MD section-12 anti-patterns as
/// NEGATIVE tests: no LIKE on System.ItemPathDisplay, never SELECT System.Message.MessageId
/// / System.Search.Contents / System.Search.EntryID, no CONTAINS('*'), mapi-only SCOPE,
/// index-backed sender/recipient shapes (per-column CONTAINS - Phase-1 probes measured
/// =/LIKE on address columns at 1-10 s property-scan cost).
/// </summary>
public sealed class WsSqlBuilderTests
{
    private const string SyntheticScope = "mapi16://{S-1-5-21-1111111111-2222222222-3333333333-1001}/alice@example.com($deadbeef)";

    private static IndexQuery BaseQuery() => new()
    {
        Scope = SyntheticScope,
        Terms = new[] { "factuur" },
    };

    [Fact]
    public void Build_DefaultShape_MatchesProvenProbeQuery()
    {
        string sql = WsSqlBuilder.Build(BaseQuery());

        Assert.StartsWith("SELECT TOP 25 ", sql, StringComparison.Ordinal);
        Assert.Contains(" FROM SystemIndex WHERE ", sql, StringComparison.Ordinal);
        Assert.Contains($"SCOPE='{SyntheticScope}'", sql, StringComparison.Ordinal);
        Assert.Contains("(System.Kind='email' OR System.Kind='document')", sql, StringComparison.Ordinal);
        Assert.Contains("CONTAINS('\"factuur\"')", sql, StringComparison.Ordinal);
        Assert.EndsWith(" ORDER BY System.Message.DateReceived DESC", sql, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_EmailOnly_WhenAttachmentHitsExcluded()
    {
        var query = BaseQuery();
        query.IncludeAttachmentHits = false;

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("System.Kind='email'", sql, StringComparison.Ordinal);
        Assert.DoesNotContain("System.Kind='document'", sql, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_NeverSelectsForbiddenColumns()
    {
        string sql = WsSqlBuilder.Build(BaseQuery());

        foreach (string forbidden in WsSqlBuilder.ForbiddenSelectColumns)
        {
            Assert.DoesNotContain(forbidden, sql, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain(forbidden, WsSqlBuilder.SelectColumns, StringComparer.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void Build_NeverEmitsLike_AntiPattern()
    {
        var query = new IndexQuery
        {
            Scope = SyntheticScope,
            Terms = new[] { "factuur", "betaling" },
            FromAddressContains = "billing@example.com",
            RecipientContains = "alice@example.com",
            ReceivedOnOrAfterUtc = new DateTime(2026, 7, 1, 0, 0, 0, DateTimeKind.Utc),
            ReceivedBeforeUtc = new DateTime(2026, 7, 20, 0, 0, 0, DateTimeKind.Utc),
            IsRead = false,
            HasAttachments = true,
        };

        string sql = WsSqlBuilder.Build(query);

        // The 9-10 s property-scan anti-patterns (v3.MD sections 5/12 + Phase-1 probes).
        Assert.DoesNotContain("LIKE", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("System.ItemPathDisplay LIKE", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("FromAddress =", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("FromAddress='", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("ToAddress =", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("ToAddress='", sql, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Build_AllFilters_ComposeIntoOneStatement()
    {
        var query = new IndexQuery
        {
            Scope = SyntheticScope,
            Terms = new[] { "factuur", "betaling" },
            FromAddressContains = "billing@example.com",
            RecipientContains = "alice@example.com",
            ReceivedOnOrAfterUtc = new DateTime(2026, 7, 1, 12, 30, 45, DateTimeKind.Utc),
            ReceivedBeforeUtc = new DateTime(2026, 7, 20, 0, 0, 0, DateTimeKind.Utc),
            IsRead = false,
            HasAttachments = true,
            Top = 100,
        };

        string sql = WsSqlBuilder.Build(query);

        Assert.StartsWith("SELECT TOP 100 ", sql, StringComparison.Ordinal);
        Assert.Contains("CONTAINS('\"factuur\" AND \"betaling\"')", sql, StringComparison.Ordinal);
        Assert.Contains("CONTAINS(System.Message.FromAddress, '\"billing@example.com\"')", sql, StringComparison.Ordinal);
        Assert.Contains("CONTAINS(System.Message.ToAddress, '\"alice@example.com\"')", sql, StringComparison.Ordinal);
        Assert.Contains("CONTAINS(System.Message.CcAddress, '\"alice@example.com\"')", sql, StringComparison.Ordinal);
        Assert.Contains("System.Message.DateReceived >= '2026-07-01 12:30:45'", sql, StringComparison.Ordinal);
        Assert.Contains("System.Message.DateReceived < '2026-07-20 00:00:00'", sql, StringComparison.Ordinal);
        Assert.Contains("System.IsRead=FALSE", sql, StringComparison.Ordinal);
        Assert.Contains("System.Message.HasAttachments=TRUE", sql, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_SubjectAndBodyTermScope_UsesColumnScopedContains()
    {
        var query = BaseQuery();
        query.TermScope = TermScope.SubjectAndBody;

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("(CONTAINS(System.Subject, '\"factuur\"') OR CONTAINS(System.Search.Contents, '\"factuur\"'))", sql, StringComparison.Ordinal);
        // Contents may be queried but never selected.
        int selectEnd = sql.IndexOf(" FROM SystemIndex", StringComparison.Ordinal);
        Assert.DoesNotContain("System.Search.Contents", sql.Substring(0, selectEnd), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Build_EscapesSingleQuotesInTerms()
    {
        var query = BaseQuery();
        query.Terms = new[] { "o'brien" };

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("CONTAINS('\"o''brien\"')", sql, StringComparison.Ordinal);
    }

    [Fact]
    public void Build_AllowsTrailingStarPrefixMatch()
    {
        var query = BaseQuery();
        query.Terms = new[] { "factu*" };

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("CONTAINS('\"factu*\"')", sql, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("*")]
    [InlineData("  *  ")]
    public void Build_RejectsBareStarTerm_ContainsStarIsInvalid(string term)
    {
        var query = BaseQuery();
        query.Terms = new[] { term };

        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query));
    }

    [Theory]
    [InlineData("fac*tuur")]
    [InlineData("*factuur")]
    [InlineData("fac\"tuur")]
    [InlineData("fac;DROP TABLE x")]
    [InlineData("fac(tuur")]
    [InlineData("")]
    [InlineData("   ")]
    public void Build_RejectsUnsafeTerms(string term)
    {
        var query = BaseQuery();
        query.Terms = new[] { term };

        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query));
    }

    [Theory]
    [InlineData("file:///c:/temp")]
    [InlineData("csc://{S-1-5-21-1}/x")]
    [InlineData("/alice@example.com/Inbox")]
    [InlineData("")]
    public void Build_RejectsNonMapiScope(string scope)
    {
        var query = BaseQuery();
        query.Scope = scope;

        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query));
    }

    [Fact]
    public void Build_RejectsScopeWithSingleQuote()
    {
        var query = BaseQuery();
        query.Scope = "mapi16://{sid}/store'--injection";

        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(5001)]
    public void Build_RejectsTopOutOfRange(int top)
    {
        var query = BaseQuery();
        query.Top = top;

        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query));
    }

    [Fact]
    public void Build_RejectsInvertedDateRange()
    {
        var query = BaseQuery();
        query.ReceivedOnOrAfterUtc = new DateTime(2026, 7, 20, 0, 0, 0, DateTimeKind.Utc);
        query.ReceivedBeforeUtc = new DateTime(2026, 7, 1, 0, 0, 0, DateTimeKind.Utc);

        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query));
    }

    [Fact]
    public void Build_RejectsUnsafeSenderAndRecipientFilters()
    {
        var withBadSender = BaseQuery();
        withBadSender.FromAddressContains = "*";
        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(withBadSender));

        var withBadRecipient = BaseQuery();
        withBadRecipient.RecipientContains = "a\"b";
        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(withBadRecipient));
    }

    [Fact]
    public void Build_NoAggregatesOrJoinsEverEmitted()
    {
        string sql = WsSqlBuilder.Build(BaseQuery());

        Assert.DoesNotContain("COUNT(", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(" JOIN ", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("GROUP BY", sql, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("DISTINCT", sql, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void BuildNewestReceivedProbe_ShapesWithAndWithoutScope()
    {
        string unscoped = WsSqlBuilder.BuildNewestReceivedProbe(null);
        Assert.Equal(
            "SELECT TOP 1 System.Message.DateReceived FROM SystemIndex WHERE System.Kind='email' ORDER BY System.Message.DateReceived DESC",
            unscoped);

        string scoped = WsSqlBuilder.BuildNewestReceivedProbe(SyntheticScope);
        Assert.Contains($"SCOPE='{SyntheticScope}' AND System.Kind='email'", scoped, StringComparison.Ordinal);
    }

    [Fact]
    public void BuildScopeExistenceProbe_ValidatesScope()
    {
        Assert.Throws<ArgumentException>(() => WsSqlBuilder.BuildScopeExistenceProbe("file:///x"));
        Assert.Equal(
            $"SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE SCOPE='{SyntheticScope}/1'",
            WsSqlBuilder.BuildScopeExistenceProbe(SyntheticScope + "/1"));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(100001)]
    public void BuildStoreDiscoverySample_RejectsOutOfRange(int top)
    {
        Assert.Throws<ArgumentException>(() => WsSqlBuilder.BuildStoreDiscoverySample(top));
    }

    [Fact]
    public void Build_TreatsUnspecifiedDateKindAsUtc()
    {
        var query = BaseQuery();
        query.ReceivedOnOrAfterUtc = new DateTime(2026, 7, 1, 6, 0, 0, DateTimeKind.Unspecified);

        string sql = WsSqlBuilder.Build(query);

        Assert.Contains("System.Message.DateReceived >= '2026-07-01 06:00:00'", sql, StringComparison.Ordinal);
    }
}
