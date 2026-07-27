using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// Soak-fix batch B wire pins (CI-safe: tools/list schema + description pins plus pre-COM
/// validation calls over stdio). B1: all four draft tools take an optional body_html that
/// is inserted as REAL HTML, mutually exclusive with body - so body may no longer be
/// required either, and the descriptions must state the normalizer's policy an agent
/// cannot guess. B2: read exposes include_html and tells the agent to use it for
/// verifying formatting, because read's plain text hides exactly that.
/// </summary>
public sealed class HtmlBodyCiToolShapeTests
{
    private static readonly string[] AllDraftTools =
    [
        "new_draft", "reply_draft", "replyall_draft", "forward_draft",
    ];

    public static TheoryData<string> AllDraftToolNames => Names(AllDraftTools);

    [Theory]
    [MemberData(nameof(AllDraftToolNames))]
    public async Task DraftTools_ExposeOptionalBodyHtml_AndBodyIsOptionalToo(string toolName)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, toolName);
        JsonElement schema = tool.GetProperty("inputSchema");
        JsonElement properties = schema.GetProperty("properties");

        Assert.True(properties.TryGetProperty("body_html", out JsonElement bodyHtml), $"{toolName} must expose body_html");
        Assert.True(properties.TryGetProperty("body", out JsonElement body), $"{toolName} must still expose body");

        // Exactly-one-of cannot be expressed in the schema, so NEITHER may be required -
        // otherwise an agent sending only body_html would be rejected by the client.
        AssertOptional(schema, toolName, "body", "body_html");

        string hint = bodyHtml.GetProperty("description").GetString()!;
        Assert.Contains("exactly one of body or body_html", hint, StringComparison.Ordinal);
        Assert.Contains("REAL HTML", hint, StringComparison.Ordinal);

        string bodyHint = body.GetProperty("description").GetString()!;
        Assert.Contains("body_html", bodyHint, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(AllDraftToolNames))]
    public async Task BodyHtmlDescription_StatesWhereItGoes_AndWhatIsKeptDroppedOrRepaired(string toolName)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, toolName);
        string hint = tool.GetProperty("inputSchema").GetProperty("properties")
            .GetProperty("body_html").GetProperty("description").GetString()!;

        // Placement: draft region only, signature and quote untouched.
        Assert.Contains("above the signature", hint, StringComparison.Ordinal);
        Assert.Contains("quoted original", hint, StringComparison.Ordinal);

        // Do not escape / do not wrap - the exact field-reported requirement.
        Assert.Contains("do NOT escape", hint, StringComparison.Ordinal);
        Assert.Contains("<pre>", hint, StringComparison.Ordinal);

        // The policy: what survives, what is dropped with content, what is unwrapped.
        Assert.Contains("h1-h6", hint, StringComparison.Ordinal);
        Assert.Contains("table", hint, StringComparison.Ordinal);
        Assert.Contains("inline style", hint, StringComparison.Ordinal);
        Assert.Contains("script", hint, StringComparison.Ordinal);
        Assert.Contains("img", hint, StringComparison.Ordinal);
        Assert.Contains("REPAIRED", hint, StringComparison.Ordinal);

        // Reporting + the verification loop.
        Assert.Contains("htmlAdjustments", hint, StringComparison.Ordinal);
        Assert.Contains("include_html", hint, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(AllDraftToolNames))]
    public async Task SupplyingBothBodies_IsRejectedAsAStructuredError_NamingBoth(string toolName)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync(toolName, ArgumentsFor(toolName, body: "text", bodyHtml: "<p>x</p>"));

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        string message = error.GetProperty("message").GetString()!;
        Assert.Contains("body_html", message, StringComparison.Ordinal);
        Assert.Contains("mutually exclusive", message, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(AllDraftToolNames))]
    public async Task SupplyingNeitherBody_IsRejectedAsAStructuredError_NamingBoth(string toolName)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync(toolName, ArgumentsFor(toolName, body: null, bodyHtml: null));

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        string message = error.GetProperty("message").GetString()!;
        Assert.Contains("body", message, StringComparison.Ordinal);
        Assert.Contains("body_html", message, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(AllDraftToolNames))]
    public async Task BodyHtmlThatNormalizesToNothing_IsRejectedAsAStructuredError(string toolName)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync(
            toolName, ArgumentsFor(toolName, body: null, bodyHtml: "<script>alert(1)</script>"));

        JsonElement error = result.GetProperty("error");
        Assert.Equal("InvalidArgument", error.GetProperty("type").GetString());
        Assert.Contains("no usable content", error.GetProperty("message").GetString()!, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Read_ExposesOptionalIncludeHtml_DescribingTheTruncationContract()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, "read");
        JsonElement schema = tool.GetProperty("inputSchema");
        JsonElement properties = schema.GetProperty("properties");

        Assert.True(properties.TryGetProperty("include_html", out JsonElement includeHtml), "read must expose include_html");
        Assert.Equal("boolean", includeHtml.GetProperty("type").GetString());
        AssertOptional(schema, "read", "include_html");

        string hint = includeHtml.GetProperty("description").GetString()!;
        Assert.Contains("bodyHtml", hint, StringComparison.Ordinal);
        Assert.Contains("bodyHtmlTotalChars", hint, StringComparison.Ordinal);
        Assert.Contains("bodyHtmlTruncated", hint, StringComparison.Ordinal);
        Assert.Contains("max_html_chars", hint, StringComparison.Ordinal);
        Assert.Contains("body_html", hint, StringComparison.Ordinal);

        // The HTML budget is a separate, larger knob - Outlook's ~40 KB of leading
        // stylesheet would otherwise fill a text-sized window with CSS.
        Assert.True(properties.TryGetProperty("max_html_chars", out JsonElement maxHtml), "read must expose max_html_chars");
        AssertOptional(schema, "read", "max_html_chars");
        string budgetHint = maxHtml.GetProperty("description").GetString()!;
        Assert.Contains("include_html", budgetHint, StringComparison.Ordinal);
        Assert.Contains("stylesheet", budgetHint, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ReadDescription_TellsAgentsThatPlainTextHidesLayout_AndThatDraftsWorkByEntryId()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement tool = await GetToolAsync(client, "read");
        string description = tool.GetProperty("description").GetString()!;

        Assert.Contains("include_html", description, StringComparison.Ordinal);
        Assert.Contains("HIDES layout", description, StringComparison.Ordinal);
        // Drafts are not in the search index - the direct-EntryID path is the only way in.
        Assert.Contains("not in the search index", description, StringComparison.Ordinal);
    }

    private static object ArgumentsFor(string toolName, string? body, string? bodyHtml)
    {
        Dictionary<string, object?> arguments = new(StringComparer.Ordinal) { ["display"] = false };
        if (body != null)
        {
            arguments["body"] = body;
        }

        if (bodyHtml != null)
        {
            arguments["body_html"] = bodyHtml;
        }

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
