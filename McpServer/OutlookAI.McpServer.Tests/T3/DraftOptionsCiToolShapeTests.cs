using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// Soak-fix batch A wire pins (CI-safe: tools/list schema + description pins plus
/// pre-COM validation calls). A2: every draft tool exposes optional cc and bcc whose
/// descriptions state the APPEND semantics and the unresolved-recipient reporting. A3:
/// the three derived draft tools expose an optional subject override whose description
/// states that threading survives. A4: all four expose importance and
/// request_read_receipt. None of them may become required, and new_draft must NOT gain a
/// subject OVERRIDE hint (its subject stays mandatory).
/// </summary>
public sealed class DraftOptionsCiToolShapeTests
{
    private static readonly string[] AllDraftTools =
    [
        "new_draft", "reply_draft", "replyall_draft", "forward_draft",
    ];

    private static readonly string[] DerivedDraftTools =
    [
        "reply_draft", "replyall_draft", "forward_draft",
    ];

    public static TheoryData<string> AllDraftToolNames => Names(AllDraftTools);

    public static TheoryData<string> DerivedDraftToolNames => Names(DerivedDraftTools);

    [Theory]
    [MemberData(nameof(AllDraftToolNames))]
    public async Task DraftTools_ExposeOptionalCcAndBcc_DocumentingAppendAndUnresolvedReporting(string toolName)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, toolName);
        JsonElement schema = tool.GetProperty("inputSchema");
        JsonElement properties = schema.GetProperty("properties");

        Assert.True(properties.TryGetProperty("cc", out JsonElement cc), $"{toolName} must expose cc");
        Assert.True(properties.TryGetProperty("bcc", out JsonElement bcc), $"{toolName} must expose bcc");
        AssertOptional(schema, toolName, "cc", "bcc");

        // A2 semantics that an agent cannot guess: APPEND, never replace, and every
        // address that fails to resolve comes back instead of vanishing.
        string ccHint = cc.GetProperty("description").GetString()!;
        Assert.Contains("ADDED", ccHint, StringComparison.Ordinal);
        Assert.Contains("never replaced", ccHint, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("unresolvedRecipients", ccHint, StringComparison.Ordinal);

        string bccHint = bcc.GetProperty("description").GetString()!;
        Assert.Contains("ADDED", bccHint, StringComparison.Ordinal);
        Assert.Contains("unresolvedRecipients", bccHint, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(DerivedDraftToolNames))]
    public async Task DerivedDraftTools_ExposeOptionalSubjectOverride_StatingThreadingSurvives(string toolName)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, toolName);
        JsonElement schema = tool.GetProperty("inputSchema");

        Assert.True(schema.GetProperty("properties").TryGetProperty("subject", out JsonElement subject),
            $"{toolName} must expose the subject override");
        AssertOptional(schema, toolName, "subject");

        string hint = subject.GetProperty("description").GetString()!;
        Assert.Contains("RE:/FW:", hint, StringComparison.Ordinal);
        Assert.Contains("threading", hint, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("ConversationIndex", hint, StringComparison.Ordinal);
        Assert.Contains("conversationTopicPreserved", hint, StringComparison.Ordinal);
        Assert.Contains("255", hint, StringComparison.Ordinal);
    }

    [Fact]
    public async Task NewDraft_KeepsSubjectRequired_AndDoesNotAdvertiseAnOverride()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement schema = (await GetToolAsync(client, "new_draft")).GetProperty("inputSchema");
        Assert.True(schema.TryGetProperty("required", out JsonElement required), "new_draft must declare required args");
        Assert.Contains("subject", required.EnumerateArray().Select(r => r.GetString()), StringComparer.Ordinal);

        string hint = schema.GetProperty("properties").GetProperty("subject").GetProperty("description").GetString()!;
        Assert.DoesNotContain("RE:/FW:", hint, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(AllDraftToolNames))]
    public async Task DraftTools_ExposeImportanceAndReadReceipt_WithTheAllowedValues(string toolName)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, toolName);
        JsonElement schema = tool.GetProperty("inputSchema");
        JsonElement properties = schema.GetProperty("properties");

        Assert.True(properties.TryGetProperty("importance", out JsonElement importance), $"{toolName} must expose importance");
        Assert.True(properties.TryGetProperty("request_read_receipt", out JsonElement receipt),
            $"{toolName} must expose request_read_receipt");
        AssertOptional(schema, toolName, "importance", "request_read_receipt");

        string importanceHint = importance.GetProperty("description").GetString()!;
        Assert.Contains("'low'", importanceHint, StringComparison.Ordinal);
        Assert.Contains("'normal'", importanceHint, StringComparison.Ordinal);
        Assert.Contains("'high'", importanceHint, StringComparison.Ordinal);

        string receiptHint = receipt.GetProperty("description").GetString()!;
        Assert.Contains("read receipt", receiptHint, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("recipients see", receiptHint, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [MemberData(nameof(AllDraftToolNames))]
    public async Task DraftTools_RejectUnknownImportance_AsAStructuredError_BeforeAnyComWork(string toolName)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        object arguments = toolName switch
        {
            "new_draft" => new
            {
                account = "hub@example.com",
                to = "a@b.example",
                subject = "s",
                body = "b",
                display = false,
                importance = "urgent",
            },
            "forward_draft" => new
            {
                id = "h424242",
                body = "b",
                to = "a@b.example",
                display = false,
                importance = "urgent",
            },
            _ => new { id = "h424242", body = "b", display = false, importance = "urgent" },
        };

        JsonElement result = await client.CallToolAsync(toolName, arguments);
        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("importance", error.GetProperty("message").GetString()!, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [MemberData(nameof(DerivedDraftToolNames))]
    public async Task DerivedDraftTools_RejectOverlongSubjectOverride_AsAStructuredError(string toolName)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        string tooLong = new('x', 256);
        object arguments = toolName == "forward_draft"
            ? new { id = "h424242", body = "b", to = "a@b.example", display = false, subject = tooLong }
            : (object)new { id = "h424242", body = "b", display = false, subject = tooLong };

        JsonElement result = await client.CallToolAsync(toolName, arguments);
        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("subject", error.GetProperty("message").GetString()!, StringComparison.OrdinalIgnoreCase);
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
