using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// T3 CI-safe slice of the Phase-2 tool surface: everything here runs WITHOUT Outlook
/// and without index content (structured error paths + environment-tolerant shapes), so
/// CI exercises real tools/call round-trips for the new tools. The full golden-shape
/// pass over live data is in <see cref="LiveMcpToolShapeTests"/> (Category=Live).
/// </summary>
public sealed class Phase2CiToolShapeTests
{
    [Fact]
    public async Task Read_UnknownHitId_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("read", new { id = "h999999" });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("Unknown id", error.GetProperty("message").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task SaveAttachment_InvalidIndex_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("save_attachment", new { id = "h1", attachment_index = 0 });

        Assert.Equal("InvalidArgument", result.GetProperty("error").GetProperty("type").GetString());
    }

    [Fact]
    public async Task Search_ExhaustiveMode_ReturnsPhase3Error()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("search", new { query = "anything", mode = "exhaustive" });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("Phase 3", error.GetProperty("message").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task IndexStatus_ReturnsShape_WithOrWithoutAnIndex()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("index_status", new { });

        // Environment-tolerant: on CI the SystemIndex is unreachable and provider
        // reports 'unavailable: ...'; on a dev machine it reports OleDb/AdodbCom.
        Assert.False(string.IsNullOrWhiteSpace(result.GetProperty("provider").GetString()));
        Assert.True(result.TryGetProperty("outlookRunning", out JsonElement outlookRunning));
        Assert.True(outlookRunning.ValueKind is JsonValueKind.True or JsonValueKind.False);
        Assert.True(result.TryGetProperty("installerMutexHeld", out _));
        Assert.True(result.GetProperty("advice").GetArrayLength() >= 1);
    }

    [Fact]
    public async Task Thread_WithoutAnyReference_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("thread", new { });

        Assert.Equal("InvalidArgument", result.GetProperty("error").GetProperty("type").GetString());
    }
}
