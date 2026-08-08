using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// Wire pins for soak fix batch C (v3.MD D46): the two NEW tools - update_draft and
/// discard_draft - plus the attachments parameter on all five draft tools. CI-safe:
/// tools/list schema and description assertions, and structured-error calls whose
/// arguments are rejected BEFORE any COM work (unknown hit id, bad paths, empty
/// request), so nothing here needs Outlook.
/// <para>
/// The description pins are not decoration. discard_draft is the only mail-deleting
/// tool in the product (S1 v3), so its advertised text must state both what it can
/// reach and what it can never reach - an agent that reads "delete a draft" and nothing
/// else would form the wrong model of the guardrail.
/// </para>
/// </summary>
public sealed class SoakBatchCCiToolShapeTests
{
    private static readonly string[] AllDraftTools =
    [
        "new_draft", "reply_draft", "replyall_draft", "forward_draft", "update_draft",
    ];

    public static TheoryData<string> AllDraftToolNames => Names(AllDraftTools);

    // ------------------------------------------------------------------ C3: attachments

    [Theory]
    [MemberData(nameof(AllDraftToolNames))]
    public async Task DraftTools_ExposeOptionalAttachmentsArray_OfStrings(string toolName)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement schema = (await GetToolAsync(client, toolName)).GetProperty("inputSchema");
        JsonElement properties = schema.GetProperty("properties");

        Assert.True(properties.TryGetProperty("attachments", out JsonElement attachments), $"{toolName} must expose attachments");
        Assert.Equal("array", TypeOf(attachments));
        Assert.Equal("string", TypeOf(attachments.GetProperty("items")));
        AssertOptional(schema, toolName, "attachments");

        string hint = attachments.GetProperty("description").GetString()!;
        Assert.Contains("ABSOLUTE", hint, StringComparison.Ordinal);
        Assert.Contains("no folder restriction", hint, StringComparison.Ordinal);
        // Fail-closed whole-set semantics must be advertised, not discovered.
        Assert.Contains("NOTHING is attached", hint, StringComparison.Ordinal);
        // The send interlock must be discoverable from the parameter itself.
        Assert.Contains("confirm_token", hint, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(AllDraftToolNames))]
    public async Task Attachments_RelativePath_IsRejectedBeforeAnyComWork_NamingThePath(string toolName)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync(toolName, ArgumentsFor(toolName, new[] { "documents\\offer.pdf" }));

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        string message = error.GetProperty("message").GetString()!;
        Assert.Contains("documents\\offer.pdf", message, StringComparison.Ordinal);
        Assert.Contains("absolute path", message, StringComparison.Ordinal);
        Assert.Contains("no draft was changed", message, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(AllDraftToolNames))]
    public async Task Attachments_MissingFile_IsRejected_AndEveryBadPathIsNamed(string toolName)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        string missingA = Path.Combine(Path.GetTempPath(), "outlookai-ci-absent-a.pdf");
        string missingB = Path.Combine(Path.GetTempPath(), "outlookai-ci-absent-b.pdf");

        JsonElement result = await client.CallToolAsync(toolName, ArgumentsFor(toolName, new[] { missingA, missingB }));

        string message = result.GetProperty("error").GetProperty("message").GetString()!;
        // BOTH are named: one retry fixes everything, instead of a bad-path whack-a-mole.
        Assert.Contains(missingA, message, StringComparison.Ordinal);
        Assert.Contains(missingB, message, StringComparison.Ordinal);
        Assert.Contains("no such file", message, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(AllDraftToolNames))]
    public async Task Attachments_Directory_IsRejectedWithADirectorySpecificReason(string toolName)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync(
            toolName, ArgumentsFor(toolName, new[] { Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar) }));

        string message = result.GetProperty("error").GetProperty("message").GetString()!;
        Assert.Contains("is a directory", message, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ C1: update_draft

    [Fact]
    public async Task UpdateDraft_Schema_RequiresOnlyId_AndOffersTheRevisableParts()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, "update_draft");
        JsonElement schema = tool.GetProperty("inputSchema");
        JsonElement properties = schema.GetProperty("properties");

        string[] required = schema.GetProperty("required").EnumerateArray().Select(r => r.GetString()!).ToArray();
        Assert.Equal(new[] { "id" }, required);

        foreach (string name in new[]
                 {
                     "body", "body_html", "subject", "to", "cc", "bcc", "importance",
                     "request_read_receipt", "signature", "attachments", "remove_attachments", "display",
                 })
        {
            Assert.True(properties.TryGetProperty(name, out _), $"update_draft must expose {name}");
        }

        Assert.Equal("array", TypeOf(properties.GetProperty("remove_attachments")));
    }

    [Fact]
    public async Task UpdateDraft_Description_SpellsOutReplaceSemantics_AndTheSurvivingRegions()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        string description = (await GetToolAsync(client, "update_draft")).GetProperty("description").GetString()!;

        // The body contract: replace the draft region, keep signature + quote.
        Assert.Contains("REPLACED, NOT APPENDED", description, StringComparison.Ordinal);
        Assert.Contains("signature", description, StringComparison.Ordinal);
        Assert.Contains("quoted original", description, StringComparison.Ordinal);

        // The recipient contract - explicitly contrasted with the creators' append, since
        // an agent carrying that habit over would silently drop recipients.
        Assert.Contains("RECIPIENTS ARE REPLACED", description, StringComparison.Ordinal);
        Assert.Contains("pass the full new list", description, StringComparison.Ordinal);

        // The attachment contract.
        Assert.Contains("ATTACHMENTS ARE ADDED", description, StringComparison.Ordinal);
        Assert.Contains("remove_attachments", description, StringComparison.Ordinal);

        // Preconditions + the token interlock.
        Assert.Contains("UNSENT", description, StringComparison.Ordinal);
        Assert.Contains("confirm_token", description, StringComparison.Ordinal);
        Assert.Contains("NOTHING IS SENT", description, StringComparison.Ordinal);

        // D47: signature images survive, and the one case that cannot is named together
        // with its remedy - a limitation an agent may not discover by accident.
        Assert.Contains("SIGNATURE IMAGES survive a revision", description, StringComparison.Ordinal);
        Assert.Contains("inlineImagesDropped", description, StringComparison.Ordinal);
        Assert.Contains("older version of this server", description, StringComparison.Ordinal);
        Assert.Contains("never silent", description, StringComparison.OrdinalIgnoreCase);

        // The retired admission ("images may be lost", with no remedy) must not come back.
        Assert.DoesNotContain("images are not preserved", description, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task UpdateDraft_Cc_DescribesReplaceNotAppend()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement properties = (await GetToolAsync(client, "update_draft"))
            .GetProperty("inputSchema").GetProperty("properties");

        string cc = properties.GetProperty("cc").GetString2("description");
        Assert.Contains("REPLACEMENT", cc, StringComparison.Ordinal);
        Assert.Contains("Omit to keep", cc, StringComparison.Ordinal);

        string to = properties.GetProperty("to").GetString2("description");
        Assert.Contains("REPLACEMENT", to, StringComparison.Ordinal);
    }

    [Fact]
    public async Task UpdateDraft_WithNothingToChange_IsRejectedBeforeAnyComWork()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("update_draft", new { id = "h424242" });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("Nothing to update", error.GetProperty("message").GetString()!, StringComparison.Ordinal);
    }

    [Fact]
    public async Task UpdateDraft_EmptyToList_IsRejected_BecauseItWouldStripEveryRecipient()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("update_draft", new { id = "h424242", to = "   " });

        string message = result.GetProperty("error").GetProperty("message").GetString()!;
        Assert.Contains("REPLACES the To list", message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task UpdateDraft_UnknownSignature_IsRejectedBeforeAnyComWork()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync(
            "update_draft", new { id = "h424242", signature = "OutlookAI-NoSuchSignature-CI" });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("was not found", error.GetProperty("message").GetString()!, StringComparison.Ordinal);
    }

    [Fact]
    public async Task UpdateDraft_UnknownId_FailsAtIdResolution_ProvingEverythingElsePassedPreComValidation()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("update_draft", new { id = "h424242", subject = "New subject" });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("Unknown id", error.GetProperty("message").GetString()!, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ C2: discard_draft

    [Fact]
    public async Task DiscardDraft_Schema_TakesExactlyOneRequiredId()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement schema = (await GetToolAsync(client, "discard_draft")).GetProperty("inputSchema");

        Assert.Equal(
            new[] { "id" },
            schema.GetProperty("required").EnumerateArray().Select(r => r.GetString()!).ToArray());
        Assert.Single(schema.GetProperty("properties").EnumerateObject());
    }

    [Fact]
    public async Task DiscardDraft_Description_StatesBothWhatItCanAndCannotTouch()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        string description = (await GetToolAsync(client, "discard_draft")).GetProperty("description").GetString()!;

        // The DESTRUCTIVE warning, like manage_signature's.
        Assert.Contains("DESTRUCTIVE", description, StringComparison.Ordinal);

        // The three conditions of the S1 v3 guardrail.
        Assert.Contains("THIS server session", description, StringComparison.Ordinal);
        Assert.Contains("UNSENT", description, StringComparison.Ordinal);
        Assert.Contains("Drafts folder", description, StringComparison.Ordinal);

        // What it can NEVER reach - stated positively so the agent cannot mis-model it.
        Assert.Contains("CAN NEVER TOUCH", description, StringComparison.Ordinal);
        Assert.Contains("already sent", description, StringComparison.Ordinal);
        Assert.Contains("Deleted Items", description, StringComparison.Ordinal);
        Assert.Contains("cannot delete permanently", description, StringComparison.Ordinal);

        // Soft delete + reversibility, the D39 contract shape.
        Assert.Contains("SOFT delete", description, StringComparison.Ordinal);
        Assert.Contains("move_mail", description, StringComparison.Ordinal);

        // No silent no-op.
        Assert.Contains("never silently does nothing", description, StringComparison.Ordinal);
    }

    [Fact]
    public async Task DiscardDraft_UnknownId_IsRejectedBeforeAnyComWork()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("discard_draft", new { id = "h424242" });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("Unknown id", error.GetProperty("message").GetString()!, StringComparison.Ordinal);
    }

    [Fact]
    public async Task DiscardDraft_EntryIdThisServerNeverCreated_IsRefusedWithTheRegistryReason()
    {
        // A syntactically valid EntryID that no draft tool in this process ever returned:
        // the registry gate must refuse it WITHOUT opening anything in Outlook, which is
        // exactly why this pin is CI-safe.
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("discard_draft", new { id = new string('A', 96) });

        JsonElement error = result.GetProperty("error");
        Assert.Equal("DraftRefused", error.GetProperty("type").GetString());
        Assert.Equal("not_created_by_this_server", error.GetProperty("reason").GetString());
        string message = error.GetProperty("message").GetString()!;
        Assert.Contains("not created or last updated by this server session", message, StringComparison.Ordinal);
        Assert.Contains("Delete it in Outlook instead", message, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ helpers

    private static string TypeOf(JsonElement property)
    {
        // The SDK emits either "type":"array" or a nullable union ["array","null"].
        JsonElement type = property.GetProperty("type");
        return type.ValueKind == JsonValueKind.Array
            ? type.EnumerateArray().Select(t => t.GetString()!).First(t => t != "null")
            : type.GetString()!;
    }

    private static object ArgumentsFor(string toolName, string[] attachments)
    {
        Dictionary<string, object?> arguments = new(StringComparer.Ordinal)
        {
            ["display"] = false,
            ["attachments"] = attachments,
            ["body"] = "body",
        };

        if (toolName == "new_draft")
        {
            arguments["account"] = "nobody@example.invalid";
            arguments["to"] = "a@b.example";
            arguments["subject"] = "s";
        }
        else
        {
            arguments["id"] = "h424242";
            if (toolName == "forward_draft")
            {
                arguments["to"] = "a@b.example";
            }
        }

        return arguments;
    }

    private static TheoryData<string> Names(string[] names)
    {
        TheoryData<string> data = new();
        foreach (string name in names)
        {
            data.Add(name);
        }

        return data;
    }

    private static void AssertOptional(JsonElement schema, string toolName, params string[] parameterNames)
    {
        if (!schema.TryGetProperty("required", out JsonElement required))
        {
            return;
        }

        string[] requiredNames = required.EnumerateArray().Select(r => r.GetString()!).ToArray();
        foreach (string parameterName in parameterNames)
        {
            Assert.DoesNotContain(parameterName, requiredNames, StringComparer.Ordinal);
        }
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

internal static class JsonElementDescriptionExtensions
{
    /// <summary>Reads a string property, failing loudly when it is absent.</summary>
    internal static string GetString2(this JsonElement element, string propertyName)
    {
        return element.GetProperty(propertyName).GetString()
            ?? throw new InvalidOperationException($"'{propertyName}' is null");
    }
}
