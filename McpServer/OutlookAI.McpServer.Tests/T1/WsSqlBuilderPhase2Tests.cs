using OutlookAI.Core.IndexSearch;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>T1: Phase-2 WS-SQL builder additions - ConversationID equality (thread tool) and size ordering.</summary>
public sealed class WsSqlBuilderPhase2Tests
{
    private const string SyntheticScope = "mapi16://{S-1-5-21-1111111111-2222222222-3333333333-1001}/alice@example.com($deadbeef)";

    [Fact]
    public void ConversationIdEquals_EmitsEqualityPredicate()
    {
        string sql = WsSqlBuilder.Build(new IndexQuery
        {
            Scope = SyntheticScope,
            Kinds = KindFilter.EmailOnly,
            ConversationIdEquals = "ABCDEF0123456789",
            Top = 50,
        });

        Assert.Contains("System.Message.ConversationID='ABCDEF0123456789'", sql, StringComparison.Ordinal);
        Assert.Contains($"SCOPE='{SyntheticScope}'", sql, StringComparison.Ordinal);
        Assert.EndsWith("ORDER BY System.Message.DateReceived DESC", sql, StringComparison.Ordinal);
    }

    [Fact]
    public void ConversationIdEquals_EscapesSingleQuotes()
    {
        string sql = WsSqlBuilder.Build(new IndexQuery
        {
            Kinds = KindFilter.EmailOnly,
            ConversationIdEquals = "abc'def",
        });

        Assert.Contains("System.Message.ConversationID='abc''def'", sql, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("bad\ncontrol")]
    public void ConversationIdEquals_RejectsEmptyAndControlCharacters(string value)
    {
        var query = new IndexQuery { ConversationIdEquals = value };
        Assert.Throws<ArgumentException>(() => WsSqlBuilder.Build(query));
    }

    [Fact]
    public void OrderBySizeDescending_ChangesOnlyTheOrderClause()
    {
        string sql = WsSqlBuilder.Build(new IndexQuery
        {
            Scope = SyntheticScope,
            Kinds = KindFilter.EmailOnly,
            OrderBy = IndexOrder.SizeDescending,
            Top = 25,
        });

        Assert.EndsWith("ORDER BY System.Size DESC", sql, StringComparison.Ordinal);
        Assert.DoesNotContain("ORDER BY System.Message.DateReceived", sql, StringComparison.Ordinal);
        foreach (string forbidden in WsSqlBuilder.ForbiddenSelectColumns)
        {
            Assert.DoesNotContain(forbidden, sql, StringComparison.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void DefaultOrder_RemainsDateReceivedDescending()
    {
        string sql = WsSqlBuilder.Build(new IndexQuery { Kinds = KindFilter.EmailOnly });
        Assert.EndsWith("ORDER BY System.Message.DateReceived DESC", sql, StringComparison.Ordinal);
    }
}
