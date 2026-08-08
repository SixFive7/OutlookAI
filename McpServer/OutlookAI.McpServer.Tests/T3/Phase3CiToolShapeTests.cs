using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// T3 CI-safe slice of the Phase-3 tool surface (show-me + exhaustive): argument
/// validation reaches the caller as structured error JSON over real stdio MCP, without
/// Outlook or an index. The full golden-shape pass over live data is in
/// <see cref="Phase3LiveMcpToolShapeTests"/> (Category=Live).
/// </summary>
public sealed class Phase3CiToolShapeTests
{
    [Fact]
    public async Task OpenInOutlook_UnknownHitId_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("open_in_outlook", new { id = "h424242" });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("Unknown id", error.GetProperty("message").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task GotoFolder_BlankStore_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("goto_folder", new { store = " " });

        Assert.Equal("InvalidArgument", result.GetProperty("error").GetProperty("type").GetString());
    }

    [Fact]
    public async Task ShowSearchResults_UnknownScope_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("show_search_results", new { query = "term", scope = "everywhere" });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("current_folder", error.GetProperty("message").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task ShowSearchResults_FolderWithoutStore_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("show_search_results", new { query = "term", folder = "Inbox" });

        Assert.Equal("InvalidArgument", result.GetProperty("error").GetProperty("type").GetString());
    }

    [Fact]
    public async Task ShowSearchResults_Description_AdvertisesTheServerAssistedAdvice()
    {
        // D35: agents must know the tool can carry a divergence note so they relay it.
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        string? description = null;
        foreach (JsonElement tool in list.GetProperty("result").GetProperty("tools").EnumerateArray())
        {
            if (tool.GetProperty("name").GetString() == "show_search_results")
            {
                description = tool.GetProperty("description").GetString();
                break;
            }
        }

        Assert.NotNull(description);
        Assert.Contains("server-assisted", description!, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("advice", description!, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("diverge", description!, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// Soak-fix-14 pin (user order 2026-07-27): the scope argument must state that the
    /// default (olSearchScopeCurrentFolder) excludes subfolders while the search tool's
    /// folder scope includes them - otherwise an agent showing what it found silently
    /// displays a narrower list than its own search returned.
    /// </summary>
    [Fact]
    public async Task ShowSearchResults_ScopeDescription_StatesTheSubfolderBehavior()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        string? scopeDescription = null;
        foreach (JsonElement tool in list.GetProperty("result").GetProperty("tools").EnumerateArray())
        {
            if (tool.GetProperty("name").GetString() == "show_search_results")
            {
                scopeDescription = tool.GetProperty("inputSchema").GetProperty("properties")
                    .GetProperty("scope").GetProperty("description").GetString();
                break;
            }
        }

        Assert.NotNull(scopeDescription);
        Assert.Contains("no subfolders", scopeDescription!, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("pass subfolders", scopeDescription!, StringComparison.OrdinalIgnoreCase);

        // All four wire values stay advertised.
        foreach (string value in new[] { "current_folder", "subfolders", "all_folders", "all_outlook" })
        {
            Assert.Contains(value, scopeDescription!, StringComparison.Ordinal);
        }
    }

    [Fact]
    public async Task Search_ExhaustiveWithStoreButNoBound_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("search", new
        {
            query = "anything",
            exhaustive = true,
            store = "someone@example.com",
        });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("bound", error.GetProperty("message").GetString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Search_ExhaustiveWithRecipientFilter_ReturnsStructuredErrorJson()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("search", new
        {
            exhaustive = true,
            store = "someone@example.com",
            folder = "Inbox",
            to = "someone",
        });

        Assert.Equal("InvalidArgument", result.GetProperty("error").GetProperty("type").GetString());
    }
}
