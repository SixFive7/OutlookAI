using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// T3 CI-safe slice of the Phase-2 tool surface: every call here is refused by argument
/// validation BEFORE any COM work, so CI exercises real tools/call round-trips without
/// Outlook and without index content. The outlook_health shape moved to
/// <see cref="OutlookHealthLiveToolShapeTests"/> when it turned out that "environment
/// tolerant" and "touches no mailbox" are different claims. The full golden-shape pass
/// over live data is in <see cref="LiveMcpToolShapeTests"/> (Category=Live).
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
    public async Task Search_Exhaustive_WithoutStore_ReturnsBoundingError()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        // Exhaustive bounding rules (store + folder/after) are validated before any
        // COM/index access, so this stays CI-safe (D34: exhaustive is a boolean flag).
        JsonElement result = await client.CallToolAsync("search", new { query = "anything", exhaustive = true });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("store", error.GetProperty("message").GetString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Thread_WithoutAnyReference_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("thread", new { });

        Assert.Equal("InvalidArgument", result.GetProperty("error").GetProperty("type").GetString());
    }
}
