using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// T3 CI-safe slice of the Phase-7 hardening surface: the outlook_health tool is advertised
/// and CALLABLE over real stdio MCP on any machine - outlook_health never starts Outlook and
/// degrades (status=degraded + problems) instead of throwing, so a CI runner without
/// Outlook/mail stores still gets a well-formed report. Also pins the has-more wording
/// on the search/thread tool descriptions (section 12 payload contract on the wire).
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
        Assert.Contains("never", description, StringComparison.OrdinalIgnoreCase); // never starts Outlook
        Assert.Contains("Outlook", description, StringComparison.Ordinal);
        Assert.Contains("index", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("audit", description, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("tuning", description, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task OutlookHealth_OnAnyMachine_ReturnsWellFormedReport()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement report = await client.CallToolAsync("outlook_health", new { });

        string status = report.GetProperty("status").GetString()!;
        Assert.True(status is "ok" or "degraded", $"unexpected status '{status}'");

        JsonElement outlook = report.GetProperty("outlook");
        bool running = outlook.GetProperty("running").GetBoolean();
        Assert.True(outlook.GetProperty("installerMutexHeld").ValueKind is JsonValueKind.True or JsonValueKind.False);

        // SF-1/SF-3 invariants (soak fix 2026-07-23): comConnected is PROBED liveness -
        // it can never be true while Outlook is not running; headless is only reported
        // for a running Outlook (nulls are omitted from payloads).
        bool comConnected = outlook.GetProperty("comConnected").GetBoolean();
        bool hasHeadless = outlook.TryGetProperty("headless", out JsonElement headless);
        if (!running)
        {
            Assert.False(comConnected, "comConnected=true while Outlook is not running (SF-1 regression)");
            Assert.False(hasHeadless && headless.ValueKind is not JsonValueKind.Null,
                "headless must be omitted when Outlook is not running");
        }
        else if (hasHeadless)
        {
            Assert.True(headless.ValueKind is JsonValueKind.True or JsonValueKind.False);
        }

        JsonElement index = report.GetProperty("index");
        Assert.False(string.IsNullOrWhiteSpace(index.GetProperty("provider").GetString()));
        Assert.False(string.IsNullOrWhiteSpace(index.GetProperty("wSearchStartMode").GetString()));

        JsonElement audit = report.GetProperty("audit");
        Assert.False(string.IsNullOrWhiteSpace(audit.GetProperty("path").GetString()));
        Assert.True(audit.GetProperty("writable").ValueKind is JsonValueKind.True or JsonValueKind.False);

        JsonElement tuning = report.GetProperty("tuning");
        Assert.True(tuning.GetProperty("managed").ValueKind is JsonValueKind.True or JsonValueKind.False);

        // D35: the tuning block always carries the EFFECTIVE UI search backend (live
        // registry read, policy hive authoritative) - even on machines where the add-in
        // never ran (CI runners report the Outlook default, server-assisted).
        string? uiSearchBackend = tuning.GetProperty("uiSearchBackend").GetString();
        Assert.True(uiSearchBackend is "local" or "server-assisted",
            $"unexpected tuning.uiSearchBackend '{uiSearchBackend}'");

        // Phase 8: whether Claude Code's user-global registration actually points at THIS
        // executable. Always present - a machine where the add-in never reconciled still
        // gets the observed verdict, which on a CI runner (no config file) is "absent".
        JsonElement registration = report.GetProperty("registration");
        string? registrationStatus = registration.GetProperty("status").GetString();
        Assert.True(
            registrationStatus is "ok" or "drifted" or "absent" or "unreadable" or "unknown",
            $"unexpected registration.status '{registrationStatus}'");
        Assert.False(string.IsNullOrWhiteSpace(registration.GetProperty("runningFrom").GetString()));

        // Degraded reports must SAY why (compact problem lines - section 12).
        if (status == "degraded")
        {
            Assert.True(report.GetProperty("problems").GetArrayLength() > 0);
        }
    }

    [Fact]
    public async Task SearchAndThread_Descriptions_CarryHasMoreContract()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();

        JsonElement list = await client.RoundTripAsync("tools/list", new { });
        var tools = list.GetProperty("result").GetProperty("tools").EnumerateArray()
            .ToDictionary(t => t.GetProperty("name").GetString()!, t => t);

        // search states the has-more contract on the 'top' argument that causes it
        // (re-homed 2026-08-17: the tool description was over Claude Code's 2 KB client
        // truncation cap, so its tail was being cut silently - see
        // DescriptionBudgetCiTests). Same wire, same words, budgeted separately.
        string searchTop = tools["search"].GetProperty("inputSchema").GetProperty("properties")
            .GetProperty("top").GetProperty("description").GetString()!;
        Assert.Contains("truncated", searchTop, StringComparison.OrdinalIgnoreCase);

        Assert.Contains("truncated", tools["thread"].GetProperty("description").GetString()!, StringComparison.OrdinalIgnoreCase);
    }
}
