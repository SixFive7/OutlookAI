using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// Wire pins for audit_log (Q93), from <c>tools/list</c> ONLY.
/// <para>
/// This class never calls the tool, and must not: a server this tier starts reads the machine's
/// real audit log - on the maintainer's workstation, his own - for every call that passes
/// validation, and <see cref="McpStdioClient"/> refuses to send one from an undeclared client.
/// Everything a call would show is pinned in-process by <c>T1/AuditLogToolTests</c>, where the log
/// is the test run's throwaway one. What only the wire can show is pinned here: the tool is
/// advertised, its arguments have the shape the description promises, and the read-only mark
/// actually reaches a client as MCP tool annotations.
/// </para>
/// </summary>
public sealed class AuditLogCiToolShapeTests
{
    [Fact]
    public async Task AuditLog_IsAnnotatedReadOnly_OnTheWire()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, "audit_log");
        Assert.True(
            tool.TryGetProperty("annotations", out JsonElement annotations),
            "audit_log carries no annotations on the wire: " + tool.GetRawText());

        // All four, because each says something different and a client may read any one of them.
        Assert.True(annotations.GetProperty("readOnlyHint").GetBoolean(), "readOnlyHint must be true");
        Assert.False(annotations.GetProperty("destructiveHint").GetBoolean(), "destructiveHint must be false");
        Assert.True(annotations.GetProperty("idempotentHint").GetBoolean(), "idempotentHint must be true");
        Assert.False(annotations.GetProperty("openWorldHint").GetBoolean(), "openWorldHint must be false");
    }

    [Fact]
    public async Task AuditLog_Schema_TakesTheFiveFilters_AndTheResumeToken_NoneRequired()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement schema = (await GetToolAsync(client, "audit_log")).GetProperty("inputSchema");
        string[] properties = schema.GetProperty("properties").EnumerateObject().Select(p => p.Name).OrderBy(n => n, StringComparer.Ordinal).ToArray();

        Assert.Equal(new[] { "after", "before", "entry_id", "operation", "resume_token", "top" }, properties);
        if (schema.TryGetProperty("required", out JsonElement required))
        {
            Assert.Empty(required.EnumerateArray());
        }

        Assert.Equal("integer", schema.GetProperty("properties").GetProperty("top").GetProperty("type").GetString());
    }

    [Fact]
    public async Task AuditLog_Description_SaysReadOnly_AndHowToFollowAnItem()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        string description = (await GetToolAsync(client, "audit_log")).GetProperty("description").GetString()!;

        // The read-only promise in words, for a client that ignores annotations.
        Assert.Contains("Read-only", description, StringComparison.Ordinal);
        Assert.Contains("never starts or touches Outlook", description, StringComparison.Ordinal);

        // The two uses that need the contract spelled out: an unknown outcome, and an item whose
        // EntryID changed under it.
        Assert.Contains("outcome \"unknown\"", description, StringComparison.Ordinal);
        Assert.Contains("newEntryId", description, StringComparison.Ordinal);
        Assert.Contains("truncated=true", description, StringComparison.Ordinal);

        // A read-only tool does not teach the mutation outcome field (AtomicityClaimsTests' rule).
        Assert.DoesNotContain("ON FAILURE, the error carries outcome", description, StringComparison.Ordinal);
    }

    private static async Task<JsonElement> GetToolAsync(McpStdioClient client, string name)
    {
        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        foreach (JsonElement tool in list.GetProperty("result").GetProperty("tools").EnumerateArray())
        {
            if (tool.GetProperty("name").GetString() == name)
            {
                return tool.Clone();
            }
        }

        throw new InvalidOperationException($"tool '{name}' is not advertised");
    }
}
