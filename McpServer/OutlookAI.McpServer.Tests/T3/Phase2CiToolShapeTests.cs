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
    public async Task OutlookHealth_CarriesTheFreshnessBlock_WithOrWithoutAnIndex()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("outlook_health", new { });

        // The merged index_status content (D37): environment-tolerant - on CI the
        // SystemIndex is unreachable and provider reports 'unavailable: ...' (advice is
        // then optional); on a dev machine it reports OleDb/AdodbCom plus advice.
        JsonElement index = result.GetProperty("index");
        string provider = index.GetProperty("provider").GetString()!;
        Assert.False(string.IsNullOrWhiteSpace(provider));
        JsonElement outlook = result.GetProperty("outlook");
        Assert.True(outlook.GetProperty("running").ValueKind is JsonValueKind.True or JsonValueKind.False);
        Assert.True(outlook.TryGetProperty("installerMutexHeld", out _));
        if (!provider.StartsWith("unavailable", StringComparison.Ordinal))
        {
            Assert.True(result.GetProperty("advice").GetArrayLength() >= 1);
        }
    }

    [Fact]
    public async Task Thread_WithoutAnyReference_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("thread", new { });

        Assert.Equal("InvalidArgument", result.GetProperty("error").GetProperty("type").GetString());
    }
}
