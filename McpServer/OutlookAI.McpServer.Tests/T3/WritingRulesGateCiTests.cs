using System.Text.Json;
using OutlookAI.McpServer.Tools;
using OutlookAI.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// The writing-rules gate ON THE WIRE: the rules the user typed into OutlookAI's prompt
/// settings actually reach an MCP client, in a real tools/call response from the real server
/// exe, and they stop arriving once they have been delivered.
/// <para>
/// T1 pins the gate's logic against an injected rules source. This tier pins the two things
/// only the wire can show: that the text an agent receives is the text the store holds
/// (nothing truncates or reformats it in between), and that the rejection is genuinely
/// transient - the retry meets the tool's normal behaviour instead of the gate again.
/// </para>
/// CI-safe: the gate answers from a registry read and a hash before any COM work, and the
/// arguments used here are rejected by argument validation, so nothing needs Outlook.
/// </summary>
public sealed class WritingRulesGateCiTests
{
    private readonly ITestOutputHelper _output;

    public WritingRulesGateCiTests(ITestOutputHelper output)
    {
        _output = output;
    }

    /// <summary>
    /// A drafting call whose arguments fail validation BEFORE any COM work: a blank account is
    /// refused with InvalidArgument. That is what makes the retry observable on a machine with
    /// no Outlook - and what keeps this test from starting one.
    /// </summary>
    private static object DraftArguments => new
    {
        account = " ",
        to = "someone@example.com",
        subject = "Quarterly update",
        body = "Hi, here is the quarterly update.",
        display = false,
    };

    [Fact]
    public async Task FirstDraftingCall_IsRefused_AndCarriesTheUsersWritingRulesVerbatim()
    {
        // The same store the server reads, so this asserts against THIS machine's rules -
        // the developer's own overrides if they have any, the shipped text if not.
        string expected = PromptStore.GetSection(PromptSection.Preamble);

        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();
        (JsonElement payload, bool isError) = await client.CallToolWithIsErrorAsync("new_draft", DraftArguments);
        JsonElement error = payload.GetProperty("error");
        string type = error.GetProperty("type").GetString()!;

        if (string.IsNullOrWhiteSpace(expected))
        {
            // A user who cleared the section has no rules to hand over, and the gate must not
            // spend a call refusing a draft to deliver nothing.
            Assert.NotEqual(WritingRulesGate.ErrorType, type);
            return;
        }

        // A real MCP error, not a success carrying an error-shaped payload: the specification's
        // guidance is that tool errors are what a model can see and self-correct from, and
        // self-correction is the entire mechanism here.
        Assert.True(isError, "the gate rejection must be flagged as an MCP tool error");
        Assert.Equal(WritingRulesGate.ErrorType, type);

        // Verbatim, byte for byte: what the user sees in the sidebar is what the agent gets.
        Assert.Equal(expected, error.GetProperty("writingRules").GetString());

        // It has to read as a retry rather than a dead end, or an agent reports a failure to
        // the user and stops.
        string message = error.GetProperty("message").GetString()!;
        Assert.Contains("call this tool again", message, StringComparison.Ordinal);
        Assert.Contains("Nothing failed", message, StringComparison.Ordinal);

        // And the clarification the rules themselves cannot carry: they were written for the
        // sidebar's plain-text insertion, so their no-HTML line must not read as a ban on
        // body_html, and the tool's own contract still applies.
        string advice = error.GetProperty("advice").GetString()!;
        Assert.Contains("body_html", advice, StringComparison.Ordinal);
        Assert.Contains("exactly one of body or body_html", advice, StringComparison.Ordinal);

        _output.WriteLine($"gate rejection: {payload.GetRawText()}");
    }

    [Fact]
    public async Task TheRetry_IsNoLongerRefusedByTheGate()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement first = await client.CallToolAsync("new_draft", DraftArguments);
        JsonElement second = await client.CallToolAsync("new_draft", DraftArguments);

        // What this CANNOT assert is that the retry SUCCEEDS. A successful new_draft needs a
        // resolvable account and a running Outlook, and this tier has neither by design (CI
        // runners have no Outlook, and starting one on a developer's machine would put a test
        // in front of a real mailbox). So the retry is aimed at an argument error instead, and
        // what is asserted is that the answer changed hands: the gate is done with this
        // process, and the call is now being judged on its own merits.
        string secondType = second.GetProperty("error").GetProperty("type").GetString()!;
        Assert.NotEqual(WritingRulesGate.ErrorType, secondType);
        Assert.Equal("InvalidArgument", secondType);
        Assert.Contains("account", second.GetProperty("error").GetProperty("message").GetString()!, StringComparison.OrdinalIgnoreCase);

        _output.WriteLine($"first: {first.GetProperty("error").GetProperty("type").GetString()} -> second: {secondType}");
    }

    [Fact]
    public async Task ARevisionThatWritesNoBody_IsNeverGated()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        // Fresh process, so the gate is still armed - but this call changes a subject, not a
        // body. Rules about how to write have nothing to say about it, and refusing it would
        // tax every recipient fix and every attachment removal.
        JsonElement result = await client.CallToolAsync("update_draft", new { id = "h424242", subject = "New subject" });

        Assert.NotEqual(
            WritingRulesGate.ErrorType,
            result.GetProperty("error").GetProperty("type").GetString());
    }
}
