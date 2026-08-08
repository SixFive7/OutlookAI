using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// T3 CI-safe slice of the Phase-5 send tool (v3.MD D4/L5): the tool is advertised
/// over real stdio MCP with its STRONGLY DISCOURAGING description intact (that text is
/// a load-bearing policy control - models must read "do not use by default" on the
/// wire), and argument validation reaches the caller as structured error JSON without
/// Outlook. The full two-step token flow over live data is in
/// <see cref="Phase5LiveMcpToolShapeTests"/> (Category=Live).
/// </summary>
public sealed class Phase5CiToolShapeTests
{
    [Fact]
    public async Task ToolsList_AdvertisesSend_WithDiscouragingDescription()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        JsonElement? sendTool = null;
        foreach (JsonElement tool in list.GetProperty("result").GetProperty("tools").EnumerateArray())
        {
            if (tool.GetProperty("name").GetString() == "send")
            {
                sendTool = tool;
                break;
            }
        }

        Assert.True(sendTool != null, "the send tool must be advertised");
        string description = sendTool!.Value.GetProperty("description").GetString()!;

        // D4: the description IS the first friction layer - assert the discouraging
        // policy phrases survive to the wire verbatim enough to steer a model.
        Assert.Contains("DO NOT USE THIS BY DEFAULT", description, StringComparison.Ordinal);
        Assert.Contains("EXPLICITLY", description, StringComparison.Ordinal);
        Assert.Contains("confirm_token", description, StringComparison.Ordinal);
        Assert.Contains("NEVER sends", description, StringComparison.Ordinal);
        Assert.Contains("audit", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("draft", description, StringComparison.OrdinalIgnoreCase);

        // The two-step surface: id required, confirm_token + sent_on_behalf_of optional.
        JsonElement schema = sendTool.Value.GetProperty("inputSchema");
        JsonElement properties = schema.GetProperty("properties");
        Assert.True(properties.TryGetProperty("id", out _));
        Assert.True(properties.TryGetProperty("confirm_token", out _));
        Assert.True(properties.TryGetProperty("sent_on_behalf_of", out _));
        var required = schema.GetProperty("required").EnumerateArray().Select(e => e.GetString()).ToList();
        Assert.Contains("id", required);
        Assert.DoesNotContain("confirm_token", required);
    }

    [Fact]
    public async Task Send_BlankId_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("send", new { id = "  " });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("id", error.GetProperty("message").GetString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Send_UnknownHitId_ReturnsStructuredErrorJson_WithoutSending()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("send", new
        {
            id = "h424242",
            confirm_token = "confirm-0123456789abcdef0123456789abcdef",
        });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("Unknown id", error.GetProperty("message").GetString(), StringComparison.Ordinal);
    }
}
