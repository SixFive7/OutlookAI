using System.Text.Json;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// Phase-7 live wire proof over real stdio MCP: the search has-more contract
/// (truncated=true is DEFINITE - over-fetch-by-one) demonstrated on the real index
/// against the tiny test-hub store, both directions (D34: searches are always fresh;
/// the idle hub yields no sweep additions, so the truncation math is index-driven.
/// S4: only counts/flags asserted).
/// </summary>
[Trait("Category", "Live")]
[Collection(LiveCollections.McpToolShape)]
public sealed class Phase7LiveMcpToolShapeTests
{
    [Fact]
    [Trait("Requires", "SmallHubStore")]
    public async Task Search_TopOne_OnHubStore_SetsTruncated_AndTopHundredDoesNot()
    {
        LiveTestSettings settings = LiveTestSettings.Load();
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        // The hub store holds a small double-digit mail corpus (v3.MD Phase-1 oracle):
        // top=1 must cut the list and say so.
        JsonElement capped = await client.CallToolAsync("search", new
        {
            store = settings.TestHubStoreDisplayName,
            top = 1,
        });

        Assert.Equal(1, capped.GetProperty("hits").GetArrayLength());
        Assert.True(capped.GetProperty("truncated").GetBoolean(),
            "top=1 on the multi-item hub store must report truncated=true");
        // The has-more advice must accompany the flag (agent-facing next step).
        bool adviceMentionsTop = capped.GetProperty("advice").EnumerateArray()
            .Any(a => a.GetString()!.Contains("top", StringComparison.OrdinalIgnoreCase));
        Assert.True(adviceMentionsTop, "truncated results must carry raise-top/narrow advice");

        // top=100 fits the whole hub corpus: no truncation.
        JsonElement uncapped = await client.CallToolAsync("search", new
        {
            store = settings.TestHubStoreDisplayName,
            top = 100,
        });

        int hubCount = uncapped.GetProperty("hits").GetArrayLength();
        Assert.InRange(hubCount, 2, 99);
        Assert.False(uncapped.GetProperty("truncated").GetBoolean(),
            "the whole hub corpus fits in top=100 - truncated must be false");
    }

    [Fact]
    [Trait("Requires", "AddInRegistry")]
    public async Task Health_OverStdio_OnThisMachine_HasOutlookVersionAndTuning()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(
            timeout: null, environment: null, McpStdioClient.OutlookReachingToolsAllowed);

        JsonElement report = await client.CallToolAsync("outlook_health", new { });

        // Machine facts that hold regardless of Outlook's running state at this point
        // in the suite: classic Outlook installed (16.x), WSearch automatic, audit
        // writable, tuning managed (Phase-6 add-in state).
        Assert.StartsWith("16.", report.GetProperty("outlook").GetProperty("version").GetString(), StringComparison.Ordinal);
        Assert.Equal("automatic", report.GetProperty("index").GetProperty("wSearchStartMode").GetString());
        Assert.True(report.GetProperty("audit").GetProperty("writable").GetBoolean());
        Assert.True(report.GetProperty("tuning").GetProperty("managed").GetBoolean());
    }
}
