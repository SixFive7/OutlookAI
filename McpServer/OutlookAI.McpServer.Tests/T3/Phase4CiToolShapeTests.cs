using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// T3 CI-safe slice of the Phase-4 draft tools: the four tools are advertised over real
/// stdio MCP and argument validation reaches the caller as structured error JSON
/// without Outlook or an index. The full write-path pass over live data (draft in
/// Drafts, signature, threading, audit lines) is in
/// <see cref="Phase4LiveMcpToolShapeTests"/> (Category=Live).
/// </summary>
public sealed class Phase4CiToolShapeTests
{
    [Fact]
    public async Task ToolsList_AdvertisesAllFourDraftTools()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        IReadOnlyList<string> names = await client.ListToolNamesAsync();

        foreach (string expected in new[] { "new_draft", "reply_draft", "replyall_draft", "forward_draft" })
        {
            Assert.Contains(expected, names);
        }
    }

    [Fact]
    public async Task NewDraft_BlankAccount_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("new_draft", new
        {
            account = " ",
            to = "a@b.example",
            subject = "s",
            body = "b",
            display = false,
        });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("account", error.GetProperty("message").GetString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task NewDraft_NoRecipients_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("new_draft", new
        {
            account = "someone@example.com",
            to = " ; ",
            subject = "s",
            body = "b",
            display = false,
        });

        Assert.Equal("InvalidArgument", result.GetProperty("error").GetProperty("type").GetString());
    }

    [Fact]
    public async Task ReplyDraft_UnknownHitId_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("reply_draft", new
        {
            id = "h424242",
            body = "b",
            display = false,
        });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("Unknown id", error.GetProperty("message").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task ReplyAllDraft_BlankBody_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("replyall_draft", new
        {
            id = "h1",
            body = " ",
            display = false,
        });

        Assert.Equal("InvalidArgument", result.GetProperty("error").GetProperty("type").GetString());
    }

    [Fact]
    public async Task ForwardDraft_NoRecipients_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("forward_draft", new
        {
            id = "h1",
            body = "b",
            to = " , ",
            display = false,
        });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("to", error.GetProperty("message").GetString(), StringComparison.OrdinalIgnoreCase);
    }
}
