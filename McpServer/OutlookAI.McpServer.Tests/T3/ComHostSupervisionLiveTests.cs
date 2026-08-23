using System.Diagnostics;
using OutlookAI.ComHost.Supervision;
using OutlookAI.McpServer.Tests.T2;
using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// The half of the COM-host supervision proof that has to ask <c>outlook_health</c>, and
/// therefore reaches the machine's own Outlook.
/// <para>
/// <b>Why it is a separate class.</b> Its siblings in
/// <see cref="ComHostSupervisionCiTests"/> drive the timeout / kill / respawn path with an
/// injected fault that fires in the COM host ABOVE the routing proxy, so the call never
/// reaches Outlook and they run honestly on a machine with none. These four cannot: the
/// facts they assert - the restart count, the failure explanation, the frame meter, the
/// host's pid - are only published by <c>outlook_health</c>, and health probes the store
/// list over COM and asks the Windows Search index for its freshness. No fault
/// specification can neutralise that, because the index half of the report does not go
/// through the COM host at all.
/// </para>
/// <para>
/// So they are <c>Category=Live</c>: excluded from a default run, included in a deliberate
/// one. <c>LiveTier=Portable</c> because any Outlook profile satisfies them - they read
/// nothing of the maintainer's own mail, and the dedicated test VM will run them unchanged.
/// </para>
/// </summary>
[Collection(LiveCollections.McpToolShape)]
[Trait("Category", "Live")]
[Trait("LiveTier", "Portable")]
[Trait("Requires", "OutlookInstance")]
public sealed class ComHostSupervisionLiveTests
{
    /// <summary>Short enough to keep the suite quick, long enough not to be flaky on a loaded box.</summary>
    private const string DeadlineMs = "4000";

    private static Dictionary<string, string> Fault(string spec) => new(StringComparer.Ordinal)
    {
        ["OUTLOOKAI_COMHOST_FAULT"] = spec,
        ["OUTLOOKAI_COMHOST_DEADLINE_MS"] = DeadlineMs,

        // These tests are about the machinery BELOW the liveness gate - deadlines, kills,
        // respawns, the breaker. The gate sits in front of all of it and, by design,
        // refuses instantly when Outlook is absent or hung. Forcing the observed state
        // keeps the tests deterministic and testing what they claim to.
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

    private static Task<McpStdioClient> StartAsync(TimeSpan timeout, string faultSpec)
    {
        return McpStdioClient.StartAndInitializeAsync(
            timeout, Fault(faultSpec), McpStdioClient.OutlookReachingToolsAllowed);
    }

    [Fact]
    public async Task AfterAWedge_TheHostIsReplacedAndHealthSaysSo()
    {
        await using McpStdioClient client = await StartAsync(TimeSpan.FromSeconds(120), "hang:GetAccounts");

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
        // ComHostSupervisionCiTests.WedgedCall_FailsWithinItsBudgetInsteadOfHangingForever.

        // An injected fault must never be mistakeable for a real one when reading health.
        Assert.Contains("hang:GetAccounts", comHost.GetProperty("injectedFault").GetString()!, StringComparison.Ordinal);

        // The measurement half of the same block, asserted here because this is the only
        // test that reads outlook_health out of a REAL server process talking to a REAL COM
        // host - the one path on which the numbers can be wrong without any unit test
        // noticing. It answers "is 64 MB the right limit?", which nobody could answer before
        // because the largest frame the product actually produces had never been measured.
        long limit = comHost.GetProperty("frameLimitBytes").GetInt64();
        long largest = comHost.GetProperty("largestFrameBytes").GetInt64();

        Assert.True(limit > 0, $"health must publish the ceiling a single message must fit under. comHost={reported}");
        Assert.True(largest > 0, $"frames have crossed by now; the high-water mark must show it. comHost={reported}");
        Assert.True(largest < limit, $"a frame at or over the limit could not have crossed. comHost={reported}");

        // And nothing was refused: this test wedges Outlook, it does not overflow a frame.
        // A non-zero count here would mean the meter is counting something else.
        Assert.Equal(0, comHost.GetProperty("framesRefusedTooLarge").GetInt32());
    }

    [Fact]
    public async Task TheWedgedHostProcessIsEnded_NotReused()
    {
        await using McpStdioClient client = await StartAsync(TimeSpan.FromSeconds(120), "hang:GetAccounts");

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
        // spending anything on COM, so on a box with Outlook closed outlook_health answers
        // without ever spawning the host. Asserting a pid there fails on the environment
        // rather than on a defect, so branch on the observed state and keep an assertion in
        // both branches.
        int? childPid;
        await using (McpStdioClient client = await StartAsync(TimeSpan.FromSeconds(120), "delay:1:GetAccounts"))
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
    public async Task WithTheBreakerOpen_HealthStillReportsAndSaysWhy()
    {
        await using McpStdioClient client = await StartAsync(TimeSpan.FromSeconds(180), "hang:GetAccounts");

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
