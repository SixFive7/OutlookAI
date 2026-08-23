using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// The tier's own guard rail, tested rather than trusted: <see cref="McpStdioClient"/>
/// refuses to send a <c>tools/call</c> for a tool that reaches the machine's own Outlook
/// unless the test declared it.
/// <para>
/// Without this class the guard would be the one piece of the classification nothing
/// exercises - by construction, since a default run no longer contains a test that calls
/// such a tool undeclared. A mechanism that only runs when it is being violated is a
/// mechanism nobody would notice the loss of, and this tier has already been through one
/// version of that: sixteen files claimed CI-safety in their names for months.
/// </para>
/// <para>
/// Both tests refuse before anything is written to the server's stdin, so neither can touch
/// a mailbox even when the guard is broken.
/// </para>
/// </summary>
public sealed class StdioClientMailboxGuardTests
{
    /// <summary>
    /// All three tools, named one by one rather than as a set, because each is in the set for
    /// its own reason: outlook_health probes the store list AND the Windows Search index,
    /// list_accounts enumerates accounts and stores (and starts Outlook when it is closed),
    /// list_folders walks every folder of every store. Dropping any one of them from the set
    /// has to fail something.
    /// </summary>
    [Theory]
    [InlineData("outlook_health")]
    [InlineData("list_accounts")]
    [InlineData("list_folders")]
    public async Task AnUndeclaredClient_RefusesEveryTool_ThatAlwaysReachesOutlook(string tool)
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        InvalidOperationException refused = await Assert.ThrowsAsync<InvalidOperationException>(
            () => client.CallToolAsync(tool, new { }));

        // The message has to name the tool and the way out, or the next person to meet it
        // deletes the guard instead of classifying their test.
        Assert.Contains(tool, refused.Message, StringComparison.Ordinal);
        Assert.Contains("Category=Live", refused.Message, StringComparison.Ordinal);
        Assert.Contains(nameof(McpStdioClient.OutlookReachingToolsAllowed), refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task AnUndeclaredClient_StillCallsToolsWhoseArgumentsAreRefusedBeforeAnyComWork()
    {
        // The guard is deliberately narrow: it reads the tool NAME, and the CI-safe half of
        // this tier is built on calls that a validation error answers before Outlook is
        // reached. A guard that blocked those would have taken most of the tier with it.
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement result = await client.CallToolAsync("thread", new { });

        Assert.Equal("InvalidArgument", result.GetProperty("error").GetProperty("type").GetString());
    }
}
