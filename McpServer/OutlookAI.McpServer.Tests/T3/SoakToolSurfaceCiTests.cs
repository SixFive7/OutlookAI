using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// Soak fix D37/D38/D39 wire pins (user-ordered tool-surface refinements): the tool
/// count is EXACTLY 19 (D37: echo/index_status/health deleted, outlook_health +
/// list_signatures added; D38: manage_signature added; D39: move_mail + archive_mail
/// added), list_folders lost its depth knob and gained stable offset paging (page
/// 1000 since D38), read gained body_offset paging, thread explains its two lookup
/// keys, the draft tools carry the optional signature parameter with the
/// pick-the-best-signature steering hint, and manage_signature carries the
/// destructive-action warning + automatic-backup contract. CI-safe:
/// schema/description pins via tools/list plus pre-COM/pre-write validation calls.
/// </summary>
public sealed class SoakToolSurfaceCiTests
{
    /// <summary>The 19 advertised tools after D39 (exact - a change here is a reviewed surface decision).</summary>
    private static readonly string[] ExpectedTools =
    [
        "search", "thread", "read", "save_attachment",
        "move_mail", "archive_mail",
        "list_accounts", "list_folders", "list_signatures", "manage_signature",
        "open_in_outlook", "goto_folder", "show_search_results",
        "new_draft", "reply_draft", "replyall_draft", "forward_draft",
        "send", "outlook_health",
    ];

    [Fact]
    public async Task ToolCount_IsExactlyNineteen_WithTheExpectedRoster()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        var names = list.GetProperty("result").GetProperty("tools").EnumerateArray()
            .Select(t => t.GetProperty("name").GetString()!)
            .OrderBy(n => n, StringComparer.Ordinal)
            .ToArray();

        Assert.Equal(ExpectedTools.OrderBy(n => n, StringComparer.Ordinal).ToArray(), names);
        Assert.Equal(19, names.Length);
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

    [Fact]
    public async Task ManageSignature_Description_CarriesDestructiveWarning_AndBackupContract()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, "manage_signature");
        string description = tool.GetProperty("description").GetString()!;

        // D38 (user-ordered): the tool description must carry the destructive-action
        // warning for delete AND mention the automatic backup with its returned path.
        Assert.Contains("DESTRUCTIVE", description, StringComparison.Ordinal);
        Assert.Contains("delete", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("backed up", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("backupPath", description, StringComparison.Ordinal);
        Assert.Contains("audit", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("set_default_for", description, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ManageSignature_Schema_RequiresActionAndName_WithObjectShapedDefaults()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, "manage_signature");
        JsonElement schema = tool.GetProperty("inputSchema");
        JsonElement properties = schema.GetProperty("properties");

        Assert.True(properties.TryGetProperty("action", out _));
        Assert.True(properties.TryGetProperty("name", out _));
        Assert.True(properties.TryGetProperty("body_text", out _));
        Assert.True(properties.TryGetProperty("body_html", out _));
        Assert.True(properties.TryGetProperty("set_default_for", out JsonElement setDefault),
            "manage_signature must expose the set_default_for object");

        // The {account, scope} shape must be advertised (object with both keys).
        JsonElement setDefaultProps = setDefault.GetProperty("properties");
        Assert.True(setDefaultProps.TryGetProperty("account", out _));
        Assert.True(setDefaultProps.TryGetProperty("scope", out _));

        var required = schema.GetProperty("required").EnumerateArray().Select(r => r.GetString()).ToArray();
        Assert.Contains("action", required);
        Assert.Contains("name", required);
        Assert.DoesNotContain("body_text", required);
        Assert.DoesNotContain("set_default_for", required);
    }

    [Theory]
    [InlineData("destroy")]
    [InlineData("")]
    public async Task ManageSignature_UnknownAction_IsRejectedBeforeAnyWork(string action)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("manage_signature", new { action, name = "X" });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("create", error.GetProperty("message").GetString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ManageSignature_DeleteUnknownName_IsRejected_WithoutTouchingAnything()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("manage_signature", new
        {
            action = "delete",
            name = "OutlookAI-NoSuchSignature-424242",
        });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("not found", error.GetProperty("message").GetString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ManageSignature_InvalidDefaultScope_IsRejectedBeforeAnyWork()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        // Default-assignment validation (D38): a bad scope fails BEFORE any file or
        // registry work - even the (nonexistent) signature is never created.
        JsonElement result = await client.CallToolAsync("manage_signature", new
        {
            action = "create",
            name = "OutlookAI-NoSuchSignature-424242",
            body_text = "x",
            set_default_for = new { account = "someone@example.com", scope = "everything" },
        });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("scope", error.GetProperty("message").GetString(), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ManageSignature_DeleteWithDefaultAssignment_IsRejected()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("manage_signature", new
        {
            action = "delete",
            name = "OutlookAI-NoSuchSignature-424242",
            set_default_for = new { account = "someone@example.com", scope = "new" },
        });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("delete", error.GetProperty("message").GetString(), StringComparison.OrdinalIgnoreCase);
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
