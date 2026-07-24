using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// D36 (soak fix 6): the MCP initialize <c>instructions</c> field is injected passively
/// into every Claude Code session at start - even when the tool roster is deferred
/// name-only by tool search (probe-proven). These tests guard the properties that make
/// that passive load acceptable and useful; T3 pins the exact string on the wire.
/// </summary>
public sealed class ServerMetadataTests
{
    /// <summary>
    /// The text loads into EVERY session on the machine, so it must stay short. Claude
    /// Code truncates server instructions at 2 KB; our own budget is far tighter.
    /// </summary>
    [Fact]
    public void Instructions_StayWithinPassiveLoadBudget()
    {
        Assert.False(string.IsNullOrWhiteSpace(McpServer.ServerMetadata.Instructions));
        Assert.InRange(McpServer.ServerMetadata.Instructions.Length, 1, 400);
    }

    /// <summary>
    /// Discovery works two ways: the model reads the sentence AND tool search matches
    /// keywords. Pin the terms an email-shaped question would search for, plus the two
    /// policy notes (drafts open for review; send is gated - D4).
    /// </summary>
    [Fact]
    public void Instructions_CarryDiscoveryKeywordsAndPolicy()
    {
        string text = McpServer.ServerMetadata.Instructions;
        foreach (string keyword in new[]
                 {
                     "Outlook", "email", "mail", "inbox", "search", "read",
                     "accounts", "attachments", "draft", "review", "send",
                 })
        {
            Assert.Contains(keyword, text, StringComparison.Ordinal);
        }
    }

    /// <summary>A probe/canary build must never ship as the registered server.</summary>
    [Fact]
    public void Instructions_ContainNoTestCanary()
    {
        Assert.DoesNotContain("CANARY", McpServer.ServerMetadata.Instructions, StringComparison.OrdinalIgnoreCase);
    }
}
