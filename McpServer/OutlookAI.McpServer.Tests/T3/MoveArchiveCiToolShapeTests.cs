using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// D39 CI wire pins for move_mail/archive_mail: schemas (ids array 1-50, folder,
/// create_folder, optional store), the load-bearing description contracts (same-store
/// restriction, undo via fromFolder/newEntryId, EntryID-change semantics, Deleted
/// Items refusal, designated-Archive semantics identical to Outlook's own Archive
/// action), and the pre-COM validation / per-item error shapes - all callable on any
/// machine (unknown hit ids fail in the hit cache before any COM work).
/// </summary>
public sealed class MoveArchiveCiToolShapeTests
{
    [Fact]
    public async Task MoveMail_Schema_IdsArrayAndFolderRequired_CreateFolderAndStoreOptional()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, "move_mail");
        JsonElement schema = tool.GetProperty("inputSchema");
        JsonElement properties = schema.GetProperty("properties");

        Assert.Equal("array", properties.GetProperty("ids").GetProperty("type").GetString());
        Assert.True(properties.TryGetProperty("folder", out _));
        Assert.Equal("boolean", properties.GetProperty("create_folder").GetProperty("type").GetString());
        Assert.True(properties.TryGetProperty("store", out _));

        var required = schema.GetProperty("required").EnumerateArray().Select(r => r.GetString()).ToArray();
        Assert.Contains("ids", required);
        Assert.Contains("folder", required);
        Assert.DoesNotContain("create_folder", required);
        Assert.DoesNotContain("store", required);
    }

    [Fact]
    public async Task MoveMail_Description_CarriesSameStoreUndoAndNoDeleteContracts()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        string description = (await GetToolAsync(client, "move_mail")).GetProperty("description").GetString()!;

        Assert.Contains("SAME-STORE", description, StringComparison.Ordinal);
        Assert.Contains("REVERSIBLE", description, StringComparison.Ordinal);
        Assert.Contains("fromFolder", description, StringComparison.Ordinal);
        Assert.Contains("newEntryId", description, StringComparison.Ordinal);
        Assert.Contains("CHANGES an item's EntryID", description, StringComparison.Ordinal);
        Assert.Contains("re-run search", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Deleted Items", description, StringComparison.Ordinal);
        Assert.Contains("cannot delete", description, StringComparison.Ordinal);
        Assert.Contains("1-50", description, StringComparison.Ordinal);
        Assert.Contains("never opens windows", description, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ArchiveMail_Schema_RequiresOnlyIds()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, "archive_mail");
        JsonElement schema = tool.GetProperty("inputSchema");

        Assert.Equal("array", schema.GetProperty("properties").GetProperty("ids").GetProperty("type").GetString());
        var required = schema.GetProperty("required").EnumerateArray().Select(r => r.GetString()).ToArray();
        Assert.Equal(new[] { "ids" }, required);
    }

    [Fact]
    public async Task ArchiveMail_Description_CarriesDesignatedFolderSemantics()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        string description = (await GetToolAsync(client, "archive_mail")).GetProperty("description").GetString()!;

        // The user-ordered contract: designated-folder semantics, identical to
        // Outlook's own Archive action, localization-proof, never silently created.
        Assert.Contains("DESIGNATED Archive folder", description, StringComparison.Ordinal);
        Assert.Contains("Archive button", description, StringComparison.Ordinal);
        Assert.Contains("Backspace", description, StringComparison.Ordinal);
        Assert.Contains("localization-proof", description, StringComparison.Ordinal);
        Assert.Contains("never guessed by name", description, StringComparison.Ordinal);
        Assert.Contains("NOTHING is created", description, StringComparison.Ordinal);
        Assert.Contains("fromFolder", description, StringComparison.Ordinal);
        Assert.Contains("newEntryId", description, StringComparison.Ordinal);
    }

    [Fact]
    public async Task MoveMail_EmptyIds_IsRejectedBeforeAnyComWork()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("move_mail", new { ids = Array.Empty<string>(), folder = "X" });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("ids", error.GetProperty("message").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task MoveMail_OverFiftyIds_IsRejectedBeforeAnyComWork()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        string[] tooMany = Enumerable.Range(0, 51).Select(i => "h" + i).ToArray();
        JsonElement result = await client.CallToolAsync("move_mail", new { ids = tooMany, folder = "X" });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("50", error.GetProperty("message").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task MoveMail_MissingFolder_IsRejectedBeforeAnyComWork()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("move_mail", new { ids = new[] { "h1" }, folder = "  " });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("folder", error.GetProperty("message").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task ArchiveMail_DuplicateIds_AreRejectedBeforeAnyComWork()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("archive_mail", new { ids = new[] { "h7", "h7" } });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("Duplicate", error.GetProperty("message").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task MoveMail_UnknownHitId_FailsPerItem_WithGoldenOutcomeShape()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        // An unknown hit id fails in the per-process hit cache BEFORE any COM work, so
        // this golden shape (requested/moved/failed/items[].ok/error) runs on any machine.
        JsonElement result = await client.CallToolAsync("move_mail", new { ids = new[] { "h999999" }, folder = "X" });

        Assert.Equal(1, result.GetProperty("requested").GetInt32());
        Assert.Equal(0, result.GetProperty("moved").GetInt32());
        Assert.Equal(1, result.GetProperty("failed").GetInt32());
        Assert.Equal("X", result.GetProperty("targetFolder").GetString());
        Assert.False(result.TryGetProperty("advice", out _), "no advice when nothing moved");

        JsonElement item = result.GetProperty("items").EnumerateArray().Single();
        Assert.Equal("h999999", item.GetProperty("id").GetString());
        Assert.False(item.GetProperty("ok").GetBoolean());
        Assert.Contains("Unknown id", item.GetProperty("error").GetString(), StringComparison.Ordinal);
    }

    [Fact]
    public async Task ArchiveMail_UnknownHitId_FailsPerItem_WithGoldenOutcomeShape()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("archive_mail", new { ids = new[] { "h999999" } });

        Assert.Equal(1, result.GetProperty("requested").GetInt32());
        Assert.Equal(0, result.GetProperty("archived").GetInt32());
        Assert.Equal(1, result.GetProperty("failed").GetInt32());

        JsonElement item = result.GetProperty("items").EnumerateArray().Single();
        Assert.False(item.GetProperty("ok").GetBoolean());
        Assert.Contains("Unknown id", item.GetProperty("error").GetString(), StringComparison.Ordinal);
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
