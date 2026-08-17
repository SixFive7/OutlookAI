using System.Diagnostics;
using System.Text.Json;
using OutlookAI.ComHost.Supervision;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// End-to-end proof that a wedged Outlook call becomes a bounded, structured failure and
/// that the server recovers - the whole point of the COM-host split (see "Why two
/// processes" in McpServer/README.md).
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

        // These tests are about the machinery BELOW the liveness gate - deadlines, kills,
        // respawns, the breaker. The gate sits in front of all of it and, by design,
        // refuses instantly when Outlook is absent or hung, which is exactly the state of
        // a CI box and was the state of the dev machine these were written on. Forcing the
        // observed state keeps the tests deterministic and testing what they claim to.
        ["OUTLOOKAI_COMHOST_LIVENESS"] = "Responsive",
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
        Assert.False(
            string.IsNullOrWhiteSpace(lastFailure.GetString()),
            $"the explanation must not be blank. comHost={reported}");

        // Deliberately NOT asserting that the explanation names GetAccounts. lastFailure
        // holds the LATEST failure, and health's own store probe is a COM call too - on a
        // machine where Outlook is genuinely unresponsive (which is exactly when someone
        // runs this) that probe times out first and overwrites it. Pinning the operation
        // name here made the test fail for a true and unrelated reason.
        //
        // The operation-specific attribution is asserted where it actually belongs, and
        // where nothing can overwrite it: in the error returned to the caller, by
        // WedgedCall_FailsWithinItsBudgetInsteadOfHangingForever.

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
        // Whether a host exists at all depends on the machine, not on the code under test.
        // Since 6f5a8b5 the liveness probe asks Windows whether Outlook is running BEFORE
        // spending anything on COM, so on a box with Outlook closed - a CI runner, or a
        // developer who has not started it - outlook_health answers without ever spawning
        // the host. Asserting a pid there fails on the environment rather than on a defect,
        // which is exactly the machine-dependence this tier avoids: it made the whole
        // Category!=Live suite pass or fail according to whether Outlook happened to be up.
        // So branch on the observed state and keep an assertion in both branches.
        int? childPid;
        await using (McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(
            TimeSpan.FromSeconds(120), Fault("delay:1:GetAccounts")))
        {
            childPid = await ComHostPidAsync(client);

            bool exited = await client.CloseAndAwaitExitAsync(TimeSpan.FromSeconds(30));
            Assert.True(exited, "the server must exit when its stdin closes");
        }

        if (childPid is null)
        {
            // No host was spawned, so nothing can have outlived the server and the
            // stdin-close assertion above is all this run can honestly prove. Start Outlook
            // to exercise the branch that matters.
            return;
        }

        // Give the kernel a moment to reap the job.
        for (int attempt = 0; attempt < 50 && IsAlive(childPid.Value); attempt++)
        {
            await Task.Delay(100);
        }

        Assert.False(IsAlive(childPid.Value), $"COM host pid {childPid} outlived its server");
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

    [Fact]
    public async Task RepeatedTimeouts_StopCostingTheFullBudget()
    {
        // Bounding each call is necessary but not sufficient. Measured against a genuinely
        // wedged Outlook on 2026-08-16: search, list_accounts and list_folders each burned
        // their full 120 s budget, independently, every single time - so the tenth request
        // in a row still took two minutes to rediscover what the first had established.
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(
            TimeSpan.FromSeconds(180), Fault("hang:GetAccounts"));

        // Two timeouts to open the breaker.
        for (int i = 0; i < ComHostPolicy.UnresponsiveTimeoutThreshold; i++)
        {
            JsonElement timedOut = await CallRawAsync(client, "list_accounts", new { });
            Assert.True(timedOut.GetProperty("isError").GetBoolean());
        }

        Stopwatch elapsed = Stopwatch.StartNew();
        JsonElement result = await CallRawAsync(client, "list_accounts", new { });
        elapsed.Stop();

        Assert.True(result.GetProperty("isError").GetBoolean());
        Assert.Equal("OutlookUnresponsive", PayloadOf(result).GetProperty("error").GetProperty("type").GetString());

        // The point of the whole mechanism: this must be effectively instant, not another
        // budget. Generous bound so a loaded CI box cannot make it flaky - the real
        // measurement is milliseconds against seconds.
        Assert.True(
            elapsed.Elapsed < TimeSpan.FromSeconds(2),
            $"a known-unresponsive Outlook must be reported immediately, not re-discovered; took {elapsed.Elapsed.TotalSeconds:F1}s");
    }

    [Fact]
    public async Task WithTheBreakerOpen_HealthStillReportsAndSaysWhy()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(
            TimeSpan.FromSeconds(180), Fault("hang:GetAccounts"));

        for (int i = 0; i < ComHostPolicy.UnresponsiveTimeoutThreshold; i++)
        {
            _ = await CallRawAsync(client, "list_accounts", new { });
        }

        JsonElement health = PayloadOf(await CallRawAsync(client, "outlook_health", new { }));
        JsonElement comHost = health.GetProperty("outlook").GetProperty("comHost");

        // Health must never be the tool that hides this: if requests are being refused
        // outright, the report has to say so and say why, or the state is invisible.
        Assert.True(
            comHost.GetProperty("unresponsive").GetBoolean(),
            $"health must disclose that COM requests are being refused. comHost={comHost.GetRawText()}");
        Assert.True(comHost.GetProperty("consecutiveTimeouts").GetInt32() >= ComHostPolicy.UnresponsiveTimeoutThreshold);
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
