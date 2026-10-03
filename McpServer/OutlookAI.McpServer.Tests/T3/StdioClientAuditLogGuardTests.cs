using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// The stdio tier's audit-log guard, tested rather than trusted (Q86, Q93): <see cref="McpStdioClient"/>
/// refuses to send a <c>tools/call</c> the server would answer by writing the machine's REAL audit
/// log - or by reading it, which is what <c>audit_log</c> does - unless the test declared contact
/// with the machine's own data.
/// <para>
/// Built so that a BROKEN guard still cannot leak a line. The end-to-end refusal uses an id that
/// the guard refuses (it is not a hit id) but the server would reject at id resolution if it ever
/// received it - so the failure mode of a regression is a red test, not a line in the real log.
/// The case the guard exists for, a well-formed raw EntryID, is pinned through the pure predicate
/// without starting a server at all.
/// </para>
/// </summary>
public sealed class StdioClientAuditLogGuardTests
{
    [Fact]
    public async Task AnUndeclaredClient_RefusesADiscardWhoseIdIsNotAHitId_BeforeSendingIt()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        InvalidOperationException refused = await Assert.ThrowsAsync<InvalidOperationException>(
            () => client.CallToolAsync("discard_draft", new { id = "not-a-hit-id" }));

        // The message has to name the tool, the resource and the way out, or the next person to
        // meet it deletes the guard instead of moving the test.
        Assert.Contains("discard_draft", refused.Message, StringComparison.Ordinal);
        Assert.Contains("audit log", refused.Message, StringComparison.Ordinal);
        Assert.Contains("T1", refused.Message, StringComparison.Ordinal);
        Assert.Contains(nameof(McpStdioClient.OutlookReachingToolsAllowed), refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task AnUndeclaredClient_RefusesAuditLog_BeforeSendingIt()
    {
        // audit_log READS the log, which for a server this tier starts is the machine's real one.
        // The arguments are ones the server would itself refuse at validation (a hit id), so even
        // a broken guard could not make this test read the real log.
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        InvalidOperationException refused = await Assert.ThrowsAsync<InvalidOperationException>(
            () => client.CallToolAsync("audit_log", new { entry_id = "h12" }));

        Assert.Contains("audit_log", refused.Message, StringComparison.Ordinal);
        Assert.Contains("audit log", refused.Message, StringComparison.Ordinal);
        Assert.Contains(nameof(McpStdioClient.OutlookReachingToolsAllowed), refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task AnUndeclaredClient_StillSendsADiscardOfAHitId()
    {
        // Narrow, like the mailbox guard: a hit id fails id resolution in a fresh server, before
        // the registry check and before anything is audited, and the CI-safe pins rely on it.
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("discard_draft", new { id = "h424242" });

        Assert.Equal("InvalidArgument", result.GetProperty("error").GetProperty("type").GetString());
    }

    [Theory]
    [InlineData("discard_draft", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", true)]
    [InlineData("discard_draft", "00000000ABCDEF0123456789ABCDEF0123456789ABCDEF0123456789", true)]
    [InlineData("discard_draft", "not-a-hit-id", true)]
    [InlineData("discard_draft", "h424242", false)]
    [InlineData("discard_draft", "h1", false)]
    [InlineData("discard_draft", "", false)]
    [InlineData("discard_draft", "   ", false)]
    [InlineData("update_draft", "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA", false)]
    public void ThePredicate_RefusesExactlyTheCallsTheServerAuditsWithoutOutlook(string tool, string id, bool refused)
    {
        JsonElement arguments = JsonSerializer.SerializeToElement(new { id });

        Assert.Equal(refused, McpStdioClient.DescribeAuditLogContact(tool, arguments) != null);
    }

    [Fact]
    public void ThePredicate_IgnoresACallWithNoArguments()
    {
        Assert.Null(McpStdioClient.DescribeAuditLogContact("discard_draft", default));
    }

    [Fact]
    public void ThePredicate_RefusesAuditLog_WhateverItsArguments()
    {
        // By name: every audit_log call that passes validation reads the real log, and which ones
        // pass is the server's judgement, not this client's.
        Assert.NotNull(McpStdioClient.DescribeAuditLogContact("audit_log", default));
        Assert.NotNull(McpStdioClient.DescribeAuditLogContact("audit_log", JsonSerializer.SerializeToElement(new { top = 1 })));
        Assert.NotNull(McpStdioClient.DescribeAuditLogContact("audit_log", JsonSerializer.SerializeToElement(new { entry_id = "h12" })));
    }
}
