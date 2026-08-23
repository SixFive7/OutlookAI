using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// T3 CI-safe slice of the Phase-7 hardening surface: the outlook_health tool is ADVERTISED
/// with its read-only promise intact, and the has-more wording on the search/thread tool
/// descriptions is pinned (section 12 payload contract on the wire). Reading a description is
/// a tools/list answer and touches nothing.
/// <para>
/// CALLING outlook_health moved to <see cref="OutlookHealthLiveToolShapeTests"/>. It really
/// does degrade instead of throwing on a machine with no Outlook, which is what the old claim
/// said - but on a machine that HAS one it attaches to the profile, enumerates the stores and
/// queries the Windows Search index per store, and that is not something a default test run
/// should do.
/// </para>
/// </summary>
public sealed class Phase7CiToolShapeTests
{
    [Fact]
    public async Task ToolsList_AdvertisesOutlookHealth_AsReadOnly()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        JsonElement? healthTool = null;
        foreach (JsonElement tool in list.GetProperty("result").GetProperty("tools").EnumerateArray())
        {
            if (tool.GetProperty("name").GetString() == "outlook_health")
            {
                healthTool = tool;
                break;
            }
        }

        Assert.True(healthTool != null, "the outlook_health tool must be advertised");
        string description = healthTool!.Value.GetProperty("description").GetString()!;

        // The read-only promise is the one fact here that NO payload field states, so the
        // description is the only place it can reach an agent - pinned as the clause it
        // actually is rather than as the bare word "never", which any sentence could satisfy.
        Assert.Contains("read-only", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("NEVER starts", description, StringComparison.Ordinal);
        Assert.Contains("Outlook", description, StringComparison.Ordinal);
        Assert.Contains("index", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("audit", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("tuning", description, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SearchAndThread_Descriptions_CarryHasMoreContract()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        var tools = list.GetProperty("result").GetProperty("tools").EnumerateArray()
            .ToDictionary(t => t.GetProperty("name").GetString()!, t => t);

        // search states the has-more contract on the 'top' argument that causes it
        // (re-homed 2026-08-17: the tool description was over Claude Code's client truncation
        // cap of 2048 UTF-16 code units, so its tail was being cut silently - see
        // DescriptionBudgetCiTests). Same wire, same words, and measurably better placed: the
        // 2026-08-18 capture of client 2.1.234 showed the cut is per string with no per-tool
        // bucket, and that parameter descriptions are not cut at any length.
        string searchTop = tools["search"].GetProperty("inputSchema").GetProperty("properties")
            .GetProperty("top").GetProperty("description").GetString()!;
        Assert.Contains("truncated", searchTop, StringComparison.OrdinalIgnoreCase);

        Assert.Contains("truncated", tools["thread"].GetProperty("description").GetString()!, StringComparison.OrdinalIgnoreCase);
    }
}
