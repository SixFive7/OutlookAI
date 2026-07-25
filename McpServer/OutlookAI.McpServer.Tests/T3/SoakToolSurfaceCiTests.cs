using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// Soak fix D37 wire pins (user-ordered tool-surface refinement, 2026-07-24): the
/// tool count is EXACTLY 16 (echo/index_status/health deleted, outlook_health +
/// list_signatures added), list_folders lost its depth knob and gained stable offset
/// paging, read gained body_offset paging, thread explains its two lookup keys, and
/// the draft tools carry the optional signature parameter with the pick-the-best-
/// signature steering hint. CI-safe: schema/description pins via tools/list plus
/// pre-COM validation calls.
/// </summary>
public sealed class SoakToolSurfaceCiTests
{
    /// <summary>The 16 advertised tools after D37 (exact - a change here is a reviewed surface decision).</summary>
    private static readonly string[] ExpectedTools =
    [
        "search", "thread", "read", "save_attachment",
        "list_accounts", "list_folders", "list_signatures",
        "open_in_outlook", "goto_folder", "show_search_results",
        "new_draft", "reply_draft", "replyall_draft", "forward_draft",
        "send", "outlook_health",
    ];

    [Fact]
    public async Task ToolCount_IsExactlySixteen_WithTheExpectedRoster()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        var names = list.GetProperty("result").GetProperty("tools").EnumerateArray()
            .Select(t => t.GetProperty("name").GetString()!)
            .OrderBy(n => n, StringComparer.Ordinal)
            .ToArray();

        Assert.Equal(ExpectedTools.OrderBy(n => n, StringComparer.Ordinal).ToArray(), names);
        Assert.Equal(16, names.Length);
    }

    [Fact]
    public async Task ListFolders_Schema_HasOffset_AndNoDepth()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, "list_folders");
        JsonElement properties = tool.GetProperty("inputSchema").GetProperty("properties");

        Assert.False(properties.TryGetProperty("depth", out _), "list_folders must not expose a depth parameter (D37)");
        Assert.True(properties.TryGetProperty("offset", out JsonElement offset), "list_folders must expose offset paging");
        Assert.Equal("integer", offset.GetProperty("type").GetString());

        // The stable traversal order is part of the contract - agents must be told it.
        string description = tool.GetProperty("description").GetString()!;
        Assert.Contains("stable", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("nextOffset", description, StringComparison.Ordinal);
        Assert.Contains("FULL", description, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Read_Schema_HasBodyOffset_AndDocumentsContinuation()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, "read");
        JsonElement properties = tool.GetProperty("inputSchema").GetProperty("properties");

        Assert.True(properties.TryGetProperty("body_offset", out JsonElement bodyOffset), "read must expose body_offset paging");
        Assert.Equal("integer", bodyOffset.GetProperty("type").GetString());

        string description = tool.GetProperty("description").GetString()!;
        Assert.Contains("body_offset", description, StringComparison.Ordinal);
        Assert.Contains("bodyTruncated", description, StringComparison.Ordinal);
        Assert.Contains("CONTINUE", description, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task Thread_Description_ExplainsBothLookupKeys()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, "thread");
        string description = tool.GetProperty("description").GetString()!;

        // Why both parameters exist: conversation_id = free/fast index path; id =
        // anchor for the COM conversation-graph fallback (COM cannot look up a
        // conversation by id string). Pass both when available.
        Assert.Contains("BOTH", description, StringComparison.Ordinal);
        Assert.Contains("fast", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("COM", description, StringComparison.Ordinal);
        Assert.Contains("cannot look up a conversation", description, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("new_draft")]
    [InlineData("reply_draft")]
    [InlineData("replyall_draft")]
    [InlineData("forward_draft")]
    public async Task DraftTools_CarryOptionalSignatureParameter_WithSteeringHint(string toolName)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, toolName);
        JsonElement schema = tool.GetProperty("inputSchema");
        JsonElement properties = schema.GetProperty("properties");

        Assert.True(properties.TryGetProperty("signature", out JsonElement signature),
            $"{toolName} must expose the signature parameter (D37)");

        // Optional: never listed as required (null/omitted = account default).
        if (schema.TryGetProperty("required", out JsonElement required))
        {
            Assert.DoesNotContain("signature",
                required.EnumerateArray().Select(r => r.GetString()), StringComparer.Ordinal);
        }

        // The user's explicit steering wish: tell the model to pick the BEST
        // signature (e.g. by language), trivial when only one exists, omit = default.
        string hint = signature.GetProperty("description").GetString()!;
        Assert.Contains("BEST", hint, StringComparison.Ordinal);
        Assert.Contains("language", hint, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("list_signatures", hint, StringComparison.Ordinal);
        Assert.Contains("default", hint, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ListSignatures_IsCallableOnAnyMachine_AndAdvertisesLanguageSteering()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, "list_signatures");
        string description = tool.GetProperty("description").GetString()!;
        Assert.Contains("language", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("excerpt", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("never starts Outlook", description, StringComparison.OrdinalIgnoreCase);

        // Pure filesystem + registry: callable everywhere, empty on a bare CI runner.
        JsonElement result = await client.CallToolAsync("list_signatures", new { });
        Assert.Equal(JsonValueKind.Array, result.GetProperty("signatures").ValueKind);
    }

    [Fact]
    public async Task NewDraft_UnknownSignature_IsRejectedBeforeAnyComWork()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("new_draft", new
        {
            account = "someone@example.com",
            to = "target@example.com",
            subject = "x",
            body = "y",
            display = false,
            signature = "OutlookAI-NoSuchSignature-424242",
        });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("signature", error.GetProperty("message").GetString(), StringComparison.OrdinalIgnoreCase);
    }

    private static async Task<JsonElement> GetToolAsync(McpStdioClient client, string name)
    {
        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        foreach (JsonElement tool in list.GetProperty("result").GetProperty("tools").EnumerateArray())
        {
            if (tool.GetProperty("name").GetString() == name)
            {
                return tool;
            }
        }

        throw new InvalidOperationException($"tool '{name}' is not advertised");
    }
}
