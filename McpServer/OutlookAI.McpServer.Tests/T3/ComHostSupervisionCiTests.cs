using System.Diagnostics;
using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// End-to-end proof that a wedged Outlook call becomes a bounded, structured failure and
/// that the server recovers - the whole point of the COM-host split (Docs/com-host.md).
/// <para>
/// CI-safe, and that is not an accident. Faults are injected in the COM host BEFORE the
/// call reaches Outlook, so the entire timeout / kill / respawn path runs on a machine
/// with no Outlook installed. Reproducing this against a genuinely wedged Outlook would
/// be untestable in CI and unsafe on a machine holding real mail.
/// </para>
/// <para>
/// This is the regression test for the 2026-08-15 incident, in which two search calls
/// each hung for the full 1800 s client timeout with no response of any kind.
/// </para>
/// </summary>
public sealed class ComHostSupervisionCiTests
{
    /// <summary>Short enough to keep the suite quick, long enough not to be flaky on a loaded CI box.</summary>
    private const string DeadlineMs = "4000";

    private static Dictionary<string, string> Fault(string spec) => new(StringComparer.Ordinal)
    {
        ["OUTLOOKAI_COMHOST_FAULT"] = spec,
        ["OUTLOOKAI_COMHOST_DEADLINE_MS"] = DeadlineMs,
    };

    private static async Task<JsonElement> CallRawAsync(McpStdioClient client, string tool, object arguments)
    {
        JsonElement envelope = await client.RoundTripAsync("tools/call", new { name = tool, arguments });
        return envelope.GetProperty("result");
    }

    private static JsonElement PayloadOf(JsonElement result)
    {
        string text = result.GetProperty("content")[0].GetProperty("text").GetString()!;
        return JsonDocument.Parse(text).RootElement;
    }

    [Fact]
    public async Task WedgedCall_FailsWithinItsBudgetInsteadOfHangingForever()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(
            TimeSpan.FromSeconds(120), Fault("hang:GetAccounts"));

        Stopwatch elapsed = Stopwatch.StartNew();
        JsonElement result = await CallRawAsync(client, "list_accounts", new { });
        elapsed.Stop();

        // The defining property: it ANSWERS. Before this change the same call produced no
        // response of any kind until the client gave up 1800 s later.
        Assert.True(
            elapsed.Elapsed < TimeSpan.FromSeconds(60),
            $"the call must fail on its budget, not hang; took {elapsed.Elapsed.TotalSeconds:F1}s");

        Assert.True(result.GetProperty("isError").GetBoolean());

        JsonElement error = PayloadOf(result).GetProperty("error");
        Assert.Equal("Timeout", error.GetProperty("type").GetString());

        // The message must name the operation and the budget, so a human reading it knows
        // what was slow rather than merely that something was.
        string message = error.GetProperty("message").GetString()!;
        Assert.Contains("GetAccounts", message, StringComparison.Ordinal);
        Assert.Contains(DeadlineMs, message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task TimeoutIsReportedAsATimeout_NotAsALostConnection()
    {
        // Regression guard for a real defect: killing the host tore down the connection,
        // and the teardown path failed everything outstanding with the vaguer "the host
        // stopped" cause, winning the race against the deadline watchdog and hiding both
        // that we ended it and why.
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(
            TimeSpan.FromSeconds(120), Fault("hang:GetAccounts"));

        JsonElement result = await CallRawAsync(client, "list_accounts", new { });
        JsonElement error = PayloadOf(result).GetProperty("error");

        Assert.Equal("Timeout", error.GetProperty("type").GetString());
        Assert.NotEqual("ComHostUnavailable", error.GetProperty("type").GetString());
    }

    [Fact]
    public async Task AfterAWedge_TheHostIsReplacedAndHealthSaysSo()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(
            TimeSpan.FromSeconds(120), Fault("hang:GetAccounts"));

        _ = await CallRawAsync(client, "list_accounts", new { });

        // Health must answer even though the host was just killed - it is asked precisely
        // in this situation, so joining the failure would defeat its purpose.
        JsonElement health = PayloadOf(await CallRawAsync(client, "outlook_health", new { }));
        JsonElement comHost = health.GetProperty("outlook").GetProperty("comHost");
        string reported = comHost.GetRawText();

        Assert.Equal("child-process", comHost.GetProperty("mode").GetString());
        Assert.True(
            comHost.GetProperty("restartCount").GetInt32() >= 1,
            $"a reclaimed wedge must be visible as a restart, not vanish silently. comHost={reported}");

        // The explanation must survive the recovery: a restart count with no reason tells
        // nobody what wedged.
        Assert.True(
            comHost.TryGetProperty("lastFailure", out JsonElement lastFailure),
            $"a restart must carry its explanation. comHost={reported}");
        Assert.Contains("GetAccounts", lastFailure.GetString()!, StringComparison.Ordinal);

        // An injected fault must never be mistakeable for a real one when reading health.
        Assert.Contains("hang:GetAccounts", comHost.GetProperty("injectedFault").GetString()!, StringComparison.Ordinal);
    }

    [Fact]
    public async Task AWedgedHostDoesNotTakeDownToolsThatDoNotNeedOutlook()
    {
        // Previously a single wedge disabled 19 of 21 tools for the life of the process.
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(
            TimeSpan.FromSeconds(120), Fault("hang:GetAccounts"));

        _ = await CallRawAsync(client, "list_accounts", new { });

        JsonElement result = await CallRawAsync(client, "list_signatures", new { });
        Assert.False(result.TryGetProperty("isError", out JsonElement isError) && isError.GetBoolean());
        Assert.True(PayloadOf(result).TryGetProperty("signatures", out _));
    }

    [Fact]
    public async Task TheWedgedHostProcessIsEnded_NotReused()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(
            TimeSpan.FromSeconds(120), Fault("hang:GetAccounts"));

        int restartsBefore = await RestartCountAsync(client);

        _ = await CallRawAsync(client, "list_accounts", new { });

        int restartsAfter = await RestartCountAsync(client);

        // Reclaiming a wedged COM call means ENDING the process that made it - the blocked
        // thread and its Outlook references cannot be recovered any other way. The restart
        // count is the honest signal for that.
        //
        // Deliberately not asserted by comparing pids: Windows reuses them aggressively,
        // and an immediate respawn genuinely can land on the pid just freed. That is not a
        // failure to replace the host, and a test that says otherwise is simply wrong.
        Assert.True(
            restartsAfter > restartsBefore,
            $"the wedged host must be replaced; restarts went {restartsBefore} -> {restartsAfter}");
    }

    [Fact]
    public async Task NoComHostSurvivesTheServer()
    {
        // The 2026-08-15 machine had 18 orphaned server processes, one wedged holding
        // Outlook COM. Process-tree lifetime is enforced by a job object plus the child's
        // own parent watch, because a killed parent cannot run cleanup code.
        int? childPid;
        await using (McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(
            TimeSpan.FromSeconds(120), Fault("delay:1:GetAccounts")))
        {
            childPid = await ComHostPidAsync(client);
            Assert.NotNull(childPid);

            bool exited = await client.CloseAndAwaitExitAsync(TimeSpan.FromSeconds(30));
            Assert.True(exited, "the server must exit when its stdin closes");
        }

        // Give the kernel a moment to reap the job.
        for (int attempt = 0; attempt < 50 && IsAlive(childPid!.Value); attempt++)
        {
            await Task.Delay(100);
        }

        Assert.False(IsAlive(childPid!.Value), $"COM host pid {childPid} outlived its server");
    }

    [Fact]
    public async Task AHostThatCrashes_ProducesAnErrorAndThenRecovers()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(
            TimeSpan.FromSeconds(120), Fault("crash:GetAccounts"));

        JsonElement result = await CallRawAsync(client, "list_accounts", new { });

        // A crashed host is a different failure from a wedged one, and must also answer
        // rather than leave the caller waiting on a process that no longer exists.
        Assert.True(result.GetProperty("isError").GetBoolean());
        string type = PayloadOf(result).GetProperty("error").GetProperty("type").GetString()!;
        Assert.Contains(type, new[] { "ComHostUnavailable", "Timeout" });

        // And the server itself stays usable.
        JsonElement signatures = await CallRawAsync(client, "list_signatures", new { });
        Assert.True(PayloadOf(signatures).TryGetProperty("signatures", out _));
    }

    private static async Task<int> RestartCountAsync(McpStdioClient client)
    {
        JsonElement health = PayloadOf(await CallRawAsync(client, "outlook_health", new { }));
        return health.GetProperty("outlook").GetProperty("comHost").GetProperty("restartCount").GetInt32();
    }

    private static async Task<int?> ComHostPidAsync(McpStdioClient client)
    {
        // outlook_health alone is enough, and is the only safe choice here. Its liveness
        // probe touches COM, so the host spawns; it is cheap, so it cannot exhaust the
        // deliberately tiny test budget; and it never calls the faulted operation, so it
        // does not kill the host whose pid we are reading.
        //
        // Both other candidates fail for instructive reasons: list_accounts IS the faulted
        // operation in these tests, and list_folders walks every folder of every store,
        // which legitimately exceeds a 4 s budget on a real multi-store profile.
        JsonElement health = PayloadOf(await CallRawAsync(client, "outlook_health", new { }));
        JsonElement comHost = health.GetProperty("outlook").GetProperty("comHost");
        return comHost.TryGetProperty("processId", out JsonElement pid) && pid.ValueKind == JsonValueKind.Number
            ? pid.GetInt32()
            : null;
    }

    private static bool IsAlive(int pid)
    {
        try
        {
            using Process process = Process.GetProcessById(pid);
            return !process.HasExited;
        }
        catch (ArgumentException)
        {
            return false;
        }
    }
}
