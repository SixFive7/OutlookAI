using System.Text.Json;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// Everything the tool-shape tier asserts by CALLING <c>outlook_health</c>, gathered in one
/// place because they share one property: they reach the machine's own Outlook.
/// <para>
/// <b>What is being corrected.</b> These tests were spread across
/// <c>Phase2CiToolShapeTests</c>, <c>Phase7CiToolShapeTests</c> and
/// <c>McpStdioConformanceTests</c>, all of which advertised themselves as CI-safe. They are
/// not: <c>outlook_health</c> attaches to a running Outlook to enumerate its stores, enriches
/// the store catalogue over COM, and queries the Windows Search index globally and once per
/// store. On a CI runner that all degrades to "not running / unavailable" and the tests still
/// pass - which is precisely what hid the problem, because on the maintainer's machine the
/// same code read a production mailbox on every verification run.
/// </para>
/// <para>
/// The tests themselves are unchanged, and deliberately so: what was wrong was the label,
/// not the assertions. They remain written as invariants that hold whatever the machine's
/// state, so they stay honest on the dedicated test VM, on a developer box and on a runner
/// with no Outlook at all.
/// </para>
/// </summary>
[Collection(LiveCollections.McpToolShape)]
[Trait("Category", "Live")]
public sealed class OutlookHealthLiveToolShapeTests
{
    private static Task<McpStdioClient> StartAsync()
    {
        return McpStdioClient.StartAndInitializeAsync(
            timeout: null, environment: null, McpStdioClient.OutlookReachingToolsAllowed);
    }

    /// <summary>
    /// The merged index_status content (D37), formerly Phase2CiToolShapeTests.
    /// </summary>
    [Fact]
    [Trait("Requires", "OutlookInstance")]
    public async Task OutlookHealth_CarriesTheFreshnessBlock_WithOrWithoutAnIndex()
    {
        await using McpStdioClient client = await StartAsync();

        JsonElement result = await client.CallToolAsync("outlook_health", new { });

        // Environment-tolerant: where the SystemIndex is unreachable the provider reports
        // 'unavailable: ...' and advice is then optional; on a machine with an index it
        // reports OleDb/AdodbCom plus advice.
        JsonElement index = result.GetProperty("index");
        string provider = index.GetProperty("provider").GetString()!;
        Assert.False(string.IsNullOrWhiteSpace(provider));
        JsonElement outlook = result.GetProperty("outlook");
        Assert.True(outlook.GetProperty("running").ValueKind is JsonValueKind.True or JsonValueKind.False);
        Assert.True(outlook.TryGetProperty("installerMutexHeld", out _));
        if (!provider.StartsWith("unavailable", StringComparison.Ordinal))
        {
            Assert.True(result.GetProperty("advice").GetArrayLength() >= 1);
        }
    }

    /// <summary>
    /// The whole report shape, formerly Phase7CiToolShapeTests. outlook_health never starts
    /// Outlook and degrades (status=degraded + problems) instead of throwing, so the report
    /// is well formed on any machine - that part of the old claim was always true.
    /// </summary>
    [Fact]
    [Trait("Requires", "OutlookInstance")]
    public async Task OutlookHealth_OnAnyMachine_ReturnsWellFormedReport()
    {
        await using McpStdioClient client = await StartAsync();

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

        // Which Office registry hive every registry-backed answer in this report came out of.
        // Present with one of the supported majors on a machine that has Outlook; OMITTED when
        // none is registered - and in that case the report MUST say so in problems, because the
        // symptom (empty accounts, empty signature defaults) is otherwise indistinguishable from
        // a healthy machine with nothing configured. That is the whole reason for the field.
        bool hasOfficeVersion = outlook.TryGetProperty("officeVersion", out JsonElement officeVersion)
            && officeVersion.ValueKind is not JsonValueKind.Null;
        if (hasOfficeVersion)
        {
            string? major = officeVersion.GetString();
            Assert.True(major is "16.0" or "17.0" or "15.0", $"unexpected outlook.officeVersion '{major}'");
        }
        else
        {
            Assert.Equal("degraded", status);
            bool explained = false;
            foreach (JsonElement problem in report.GetProperty("problems").EnumerateArray())
            {
                explained |= problem.GetString()?.Contains("No supported Office version", StringComparison.Ordinal) == true;
            }

            Assert.True(explained, "officeVersion is absent but problems does not explain that no supported Office version was found");
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
        // never ran (a bare runner reports the Outlook default, server-assisted).
        string? uiSearchBackend = tuning.GetProperty("uiSearchBackend").GetString();
        Assert.True(uiSearchBackend is "local" or "server-assisted",
            $"unexpected tuning.uiSearchBackend '{uiSearchBackend}'");

        // Phase 8: whether Claude Code's user-global registration actually points at THIS
        // executable. Always present - a machine where the add-in never reconciled still
        // gets the observed verdict, which with no config file is "absent".
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

    /// <summary>
    /// The smoke call that used to close <see cref="McpStdioConformanceTests"/>: a real
    /// <c>tools/call</c> for the one tool whose report is assembled in OutlookAI.Core, so the
    /// server-to-Core chain is proven over the wire rather than in process.
    /// <para>
    /// It moved here because it is the only step of that conformance run that touched a
    /// mailbox. The conformance test kept its handshake, its instructions pin, its tool
    /// roster, its stdin-close proof and a <c>tools/call</c> - it now uses
    /// <c>list_signatures</c>, which is also built in Core and reads only the signature
    /// directory. So CI still proves all four verbs end to end, and nothing proves them
    /// against a real mailbox unless someone asks for a live run.
    /// </para>
    /// </summary>
    [Fact]
    [Trait("Requires", "OutlookInstance")]
    public async Task OutlookHealth_IsCallableOverRawStdio_AndAnswersWithAStatus()
    {
        await using McpStdioClient client = await StartAsync();

        (JsonElement payload, bool isError) = await client.CallToolWithIsErrorAsync("outlook_health", new { });

        Assert.False(isError, "outlook_health tool call reported isError=true.");
        string? status = payload.GetProperty("status").GetString();
        Assert.True(status is "ok" or "degraded", $"unexpected outlook_health status '{status}'");
    }
}
