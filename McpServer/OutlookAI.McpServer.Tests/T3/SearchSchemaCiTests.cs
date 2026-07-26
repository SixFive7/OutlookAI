using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// D34 wire-schema pin (user decision 2026-07-24: "drop the fast mode. Keep everything
/// inside of 1 tool."): the search tool advertises NO mode enum - fresh (index +
/// freshness sweep) is THE behavior - and exhaustive survives as a boolean flag on the
/// same tool. CI-safe: tools/list needs no Outlook and no index.
/// </summary>
public sealed class SearchSchemaCiTests
{
    [Fact]
    public async Task SearchInputSchema_HasNoModeParameter_AndHasExhaustiveBoolean()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        JsonElement? searchTool = null;
        foreach (JsonElement tool in list.GetProperty("result").GetProperty("tools").EnumerateArray())
        {
            if (tool.GetProperty("name").GetString() == "search")
            {
                searchTool = tool;
                break;
            }
        }

        Assert.True(searchTool != null, "the search tool must be advertised");
        JsonElement properties = searchTool!.Value.GetProperty("inputSchema").GetProperty("properties");

        Assert.False(properties.TryGetProperty("mode", out _),
            "the search schema must not expose a 'mode' parameter (removed by D34)");
        Assert.True(properties.TryGetProperty("exhaustive", out JsonElement exhaustive),
            "the search schema must expose the 'exhaustive' boolean (D34)");
        Assert.Equal("boolean", exhaustive.GetProperty("type").GetString());

        // The description must document the always-fresh contract + graceful degradation.
        string description = searchTool.Value.GetProperty("description").GetString()!;
        Assert.Contains("fresh", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("advice", description, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("mode=", description, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// D40 wire pin (user order 2026-07-26, SF-6 fix): the search_in argument exists,
    /// is a string, and the tool description states plainly what 'query' matches by
    /// default - subject AND body - instead of the old "search all ... mail" overpromise
    /// that hid a body-content-only predicate.
    /// </summary>
    [Fact]
    public async Task SearchSchema_ExposesSearchIn_AndDescribesWhatQueryMatches()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        JsonElement? searchTool = null;
        foreach (JsonElement tool in list.GetProperty("result").GetProperty("tools").EnumerateArray())
        {
            if (tool.GetProperty("name").GetString() == "search")
            {
                searchTool = tool;
                break;
            }
        }

        Assert.True(searchTool != null, "the search tool must be advertised");
        JsonElement properties = searchTool!.Value.GetProperty("inputSchema").GetProperty("properties");

        Assert.True(properties.TryGetProperty("search_in", out JsonElement searchIn),
            "the search schema must expose the 'search_in' parameter (D40, renamed 2026-07-26)");
        Assert.Contains("string", DescribeJsonType(searchIn.GetProperty("type")), StringComparison.Ordinal);

        // The pre-rename name must be gone from the wire entirely.
        Assert.False(properties.TryGetProperty("term_scope", out _),
            "the search schema must not expose the old 'term_scope' name (renamed to 'search_in')");

        // The parameter description must name all three values and explain the default.
        string paramDescription = searchIn.GetProperty("description").GetString()!;
        foreach (string wireName in new[] { "subject_and_body", "subject", "body" })
        {
            Assert.Contains(wireName, paramDescription, StringComparison.Ordinal);
        }

        Assert.Contains("default", paramDescription, StringComparison.OrdinalIgnoreCase);

        // search_in must be optional - omitting it is the default subject+body search.
        if (searchTool.Value.GetProperty("inputSchema").TryGetProperty("required", out JsonElement required))
        {
            foreach (JsonElement name in required.EnumerateArray())
            {
                Assert.NotEqual("search_in", name.GetString());
            }
        }

        // The tool description must say what 'query' matches, and must not repeat the
        // pre-D40 overpromise ("Search all locally indexed Outlook mail ...").
        string description = searchTool.Value.GetProperty("description").GetString()!;
        Assert.Contains("subject", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("body", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("search_in", description, StringComparison.Ordinal);
        Assert.DoesNotContain("term_scope", description, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Search all locally indexed Outlook mail", description, StringComparison.OrdinalIgnoreCase);
    }

    private static string DescribeJsonType(JsonElement type)
    {
        if (type.ValueKind == JsonValueKind.Array)
        {
            return string.Join(",", type.EnumerateArray().Select(t => t.GetString()));
        }

        return type.GetString() ?? string.Empty;
    }
}
