using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using OutlookAI.ComHost.Host;
using OutlookAI.ComHost.Supervision;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// End-to-end proof that a wedged Outlook call becomes a bounded, structured failure and
/// that the server recovers - the whole point of the COM-host split (Docs/com-host.md).
/// <para>
/// CI-safe, and that is not an accident. Faults are injected in the COM host BEFORE the
/// call reaches Outlook - <c>ComHostServer</c> applies them above the routing proxy, so no
/// session is ever asked for - and the entire timeout / kill / respawn path therefore runs
/// on a machine with no Outlook installed. Reproducing this against a genuinely wedged
/// Outlook would be untestable in CI and unsafe on a machine holding real mail.
/// </para>
/// <para>
/// The four tests that could NOT keep that promise now live in
/// <see cref="ComHostSupervisionLiveTests"/>: everything they assert is published only by
/// <c>outlook_health</c>, which probes the store list over COM and queries the Windows
/// Search index, neither of which any fault specification can neutralise. They were running
/// in every <c>Category!=Live</c> pass against the maintainer's own mailbox.
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

    /// <summary>
    /// Starts a server whose faulted operation is neutralised in the COM host, and says so.
    /// <para>
    /// <see cref="McpStdioClient.OutlookReachingToolsAllowed"/> is passed because every test
    /// below calls <c>list_accounts</c>, which reaches Outlook for every argument shape. Here
    /// it does not: the fault is applied in <c>ComHostServer</c> ABOVE the routing proxy, so
    /// the call is answered - or hung, or crashed - before it can reach a session. That is
    /// the one legitimate way to name an Outlook-reaching tool outside the live tier, and
    /// T1 LiveTierInventoryTests recognises it by the injected fault rather than taking the
    /// declaration on trust.
    /// </para>
    /// </summary>
    private static Task<McpStdioClient> StartAsync(TimeSpan timeout, string faultSpec)
    {
        return McpStdioClient.StartAndInitializeAsync(
            timeout, Fault(faultSpec), McpStdioClient.OutlookReachingToolsAllowed);
    }

    private static JsonElement PayloadOf(JsonElement result)
    {
        string text = result.GetProperty("content")[0].GetProperty("text").GetString()!;
        return JsonDocument.Parse(text).RootElement;
    }

    [Fact]
    public async Task WedgedCall_FailsWithinItsBudgetInsteadOfHangingForever()
    {
        await using McpStdioClient client = await StartAsync(TimeSpan.FromSeconds(120), "hang:GetAccounts");

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
        await using McpStdioClient client = await StartAsync(TimeSpan.FromSeconds(120), "hang:GetAccounts");

        JsonElement result = await CallRawAsync(client, "list_accounts", new { });
        JsonElement error = PayloadOf(result).GetProperty("error");

        Assert.Equal("Timeout", error.GetProperty("type").GetString());
        Assert.NotEqual("ComHostUnavailable", error.GetProperty("type").GetString());
    }

    [Fact]
    public async Task AWedgedHostDoesNotTakeDownToolsThatDoNotNeedOutlook()
    {
        // Previously a single wedge disabled 19 of 21 tools for the life of the process.
        await using McpStdioClient client = await StartAsync(TimeSpan.FromSeconds(120), "hang:GetAccounts");

        _ = await CallRawAsync(client, "list_accounts", new { });

        JsonElement result = await CallRawAsync(client, "list_signatures", new { });
        Assert.False(result.TryGetProperty("isError", out JsonElement isError) && isError.GetBoolean());
        Assert.True(PayloadOf(result).TryGetProperty("signatures", out _));
    }


    [Fact]
    public async Task AHostThatCrashes_ProducesAnErrorAndThenRecovers()
    {
        await using McpStdioClient client = await StartAsync(TimeSpan.FromSeconds(120), "crash:GetAccounts");

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
        await using McpStdioClient client = await StartAsync(TimeSpan.FromSeconds(180), "hang:GetAccounts");

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
    public async Task AnErrorRaisedInsideTheHost_ReachesTheCallerWithItsOwnMessage()
    {
        // The 2026-08-18 defect, end to end. An exhaustive search naming a folder that does
        // not exist came back as ComHostRemoteException: "Exception has been thrown by the
        // target of an invocation." - the message of a reflection wrapper, not of the
        // failure. The good message the session builds ("Folder 'X' was not found in store
        // 'Y' ... Use list_folders for paths") was constructed, thrown, and then discarded
        // at the process boundary, on the primary store and on a delegate store alike.
        //
        // The wrapper was earned twice: the routing proxy calls the session by reflection,
        // and the host dispatches by reflection, so the host peeled one layer and shipped
        // the other. That means it was never one message that was lost - EVERY deliberate
        // error raised inside the session read the same, and the parent's whole
        // exception-type mapping was unreachable.
        await using McpStdioClient client = await StartAsync(
            TimeSpan.FromSeconds(120), $"{ComHostFaultInjection.SessionThrowKind}:folder:GetAccounts");

        JsonElement result = await CallRawAsync(client, "list_accounts", new { });
        Assert.True(result.GetProperty("isError").GetBoolean());

        JsonElement error = PayloadOf(result).GetProperty("error");
        string type = error.GetProperty("type").GetString()!;
        string message = error.GetProperty("message").GetString()!;

        // The message is the caller's whole diagnosis, so it is asserted first and in full.
        Assert.Equal(ComHostFaultInjection.SessionFolderMessage, message);

        // And the failure this replaces, named explicitly - a future reflective layer
        // anywhere on this path would silently reintroduce it.
        Assert.DoesNotContain("target of an invocation", message, StringComparison.OrdinalIgnoreCase);

        // The TYPE has to cross too, not only the text: the tool layer branches on it to
        // choose the error payload, so a flattened type downgrades every mapped failure to
        // an anonymous one even when the message happens to survive.
        Assert.Equal(nameof(InvalidOperationException), type);
        Assert.NotEqual("ComHostRemoteException", type);
        Assert.NotEqual("TargetInvocationException", type);
    }

    [Fact]
    public async Task AComFailureInsideTheHost_KeepsItsTypeAndHresult()
    {
        // The same crossing, for the payload the machinery reads rather than the human. A
        // COMException's HRESULT is what ComGateway keys its disconnect rebuild on and what
        // the tool layer reports as ComFailure; a TargetInvocationException carries neither,
        // so a genuine RPC_E_DISCONNECTED stopped being recognisable as one.
        await using McpStdioClient client = await StartAsync(
            TimeSpan.FromSeconds(120), $"{ComHostFaultInjection.SessionThrowKind}:com:GetAccounts");

        JsonElement result = await CallRawAsync(client, "list_accounts", new { });
        Assert.True(result.GetProperty("isError").GetBoolean());

        JsonElement error = PayloadOf(result).GetProperty("error");
        Assert.Equal("ComFailure", error.GetProperty("type").GetString());

        // Guard renders a COM failure as "<TypeName> 0x<HRESULT>", so both halves of the
        // reconstruction are visible in one string.
        string message = error.GetProperty("message").GetString()!;
        Assert.Contains(nameof(System.Runtime.InteropServices.COMException), message, StringComparison.Ordinal);
        Assert.Contains(
            string.Format(CultureInfo.InvariantCulture, "0x{0:X8}", ComHostFaultInjection.SessionComHResult),
            message,
            StringComparison.Ordinal);
    }
}
