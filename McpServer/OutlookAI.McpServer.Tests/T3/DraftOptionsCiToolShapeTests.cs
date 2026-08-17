using System.Globalization;
using System.Text.Json;
using OutlookAI.Core.Services;
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
    private readonly Xunit.Abstractions.ITestOutputHelper _output;

    public DraftOptionsCiToolShapeTests(Xunit.Abstractions.ITestOutputHelper output)
    {
        _output = output;
    }

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
        // From the WIRE, and derived: the cap in the hint is the one the service enforces.
        Assert.Contains(
            MailService.SubjectCharsCap.ToString(CultureInfo.InvariantCulture), hint, StringComparison.Ordinal);
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
        // The first drafting call of a server process is answered with the user's writing
        // rules instead of the work (WritingRulesGate); spend it before the real call.
        await client.PrimeWritingRulesGateAsync();

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

    /// <summary>
    /// Every tool that enforces the 255-character subject cap, including update_draft -
    /// which is not a "draft tool" in the sense the lists above use, but shares the gate.
    /// </summary>
    public static TheoryData<string> SubjectCapToolNames =>
        Names(["new_draft", "reply_draft", "replyall_draft", "forward_draft", "update_draft"]);

    [Theory]
    [MemberData(nameof(SubjectCapToolNames))]
    public async Task DraftTools_RefuseAnOverlongSubject_WithAnErrorTheModelCanSelfCorrectFrom(string toolName)
    {
        // The cap is taught by the ERROR, not by the subject argument's description (user
        // decision: an over-long subject is rare, and every call would pay the description
        // budget for it). That is only a fair trade while the refusal carries everything a
        // retry needs, which is what this pins from the WIRE - the surface an agent
        // actually meets. It reaches the model as a real tool error (isError plus this
        // server's {"error": ...} shape), not as a protocol fault or a bare exception.
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();
        await client.PrimeWritingRulesGateAsync();

        int supplied = MailService.SubjectCharsCap + 57;
        string tooLong = new('x', supplied);
        object arguments = toolName switch
        {
            "new_draft" => new
            {
                account = "hub@example.com",
                to = "a@b.example",
                subject = tooLong,
                body = "b",
                display = false,
            },
            "forward_draft" => new { id = "h424242", body = "b", to = "a@b.example", display = false, subject = tooLong },
            "update_draft" => new { id = "h424242", display = false, subject = tooLong },
            _ => (object)new { id = "h424242", body = "b", display = false, subject = tooLong },
        };

        (JsonElement payload, bool isError) = await client.CallToolWithIsErrorAsync(toolName, arguments);
        _output.WriteLine($"{toolName} isError={isError} payload={payload.GetRawText()}");

        Assert.True(isError, $"{toolName} must flag an over-long subject as a tool error");
        JsonElement error = payload.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());

        string message = error.GetProperty("message").GetString()!;
        Assert.Contains("subject", message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(
            "max " + MailService.SubjectCharsCap.ToString(CultureInfo.InvariantCulture) + " characters",
            message,
            StringComparison.Ordinal);
        Assert.Contains(
            supplied.ToString(CultureInfo.InvariantCulture) + " characters supplied",
            message,
            StringComparison.Ordinal);
        Assert.Contains("Nothing was created or changed", message, StringComparison.Ordinal);
        Assert.Contains("call again", message, StringComparison.OrdinalIgnoreCase);

        // Pre-COM, and provably so: this ran on a machine that may have no Outlook at all,
        // and the answer is the validation error rather than an Outlook-state error.
        Assert.DoesNotContain("Outlook", message, StringComparison.OrdinalIgnoreCase);
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
