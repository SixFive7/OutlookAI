using System.Diagnostics;
using System.Text.Json;
using Xunit;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// End-to-end guarantees about what happens when Outlook is not simply available.
/// <para>
/// Written as invariants rather than as assertions about one machine state, on purpose.
/// These run in CI (no Outlook at all), on a developer box (Outlook healthy), and were
/// developed against a genuinely wedged Outlook. A test that only holds in one of those
/// would be worse than no test: it would go red for reasons that are not defects, which
/// is how a suite stops being believed.
/// </para>
/// </summary>
public sealed class OutlookAvailabilityCiTests
{
    /// <summary>Error types that mean "not now, try again" rather than "this went wrong".</summary>
    private static readonly string[] TransientTypes =
    {
        "OutlookStarting", "OutlookUnresponsive", "Timeout", "ComHostUnavailable", "OutlookUnavailable",
    };

    private static async Task<(JsonElement Result, TimeSpan Elapsed)> CallAsync(
        McpStdioClient client, string tool, object arguments)
    {
        Stopwatch clock = Stopwatch.StartNew();
        JsonElement envelope = await client.RoundTripAsync("tools/call", new { name = tool, arguments });
        clock.Stop();
        return (envelope.GetProperty("result"), clock.Elapsed);
    }

    private static JsonElement PayloadOf(JsonElement result)
    {
        string text = result.GetProperty("content")[0].GetProperty("text").GetString()!;
        return JsonDocument.Parse(text).RootElement;
    }

    private static bool IsError(JsonElement result) =>
        result.TryGetProperty("isError", out JsonElement flag) && flag.GetBoolean();

    [Fact]
    public async Task ATransientOutlookState_AnswersFastAndCarriesRetryGuidance()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(TimeSpan.FromSeconds(180));

        (JsonElement result, TimeSpan elapsed) = await CallAsync(client, "list_accounts", new { });

        // Whatever the machine's state, a COM-needing tool must not sit on the caller.
        Assert.True(elapsed < TimeSpan.FromSeconds(100), $"took {elapsed.TotalSeconds:F1}s");

        if (!IsError(result))
        {
            return; // Outlook was healthy here; nothing transient to assert.
        }

        JsonElement error = PayloadOf(result).GetProperty("error");
        string type = error.GetProperty("type").GetString()!;
        Assert.Contains(type, TransientTypes);

        // The states we can do something about must say WHEN to come back. Guidance
        // without a number is not guidance - an agent cannot act on "later".
        if (type is "OutlookStarting" or "OutlookUnresponsive")
        {
            Assert.True(
                error.TryGetProperty("retryAfterSeconds", out JsonElement retry),
                $"a retryable state must carry retryAfterSeconds; got {error.GetRawText()}");
            Assert.InRange(retry.GetInt32(), 1, 300);
        }
    }

    [Fact]
    public async Task SearchAlwaysAnswers_AndSaysWhetherItIsComplete()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(TimeSpan.FromSeconds(180));

        (JsonElement result, TimeSpan elapsed) = await CallAsync(
            client, "search", new { query = "invoice", top = 3 });

        Assert.True(elapsed < TimeSpan.FromSeconds(100), $"search took {elapsed.TotalSeconds:F1}s");

        JsonElement payload = PayloadOf(result);
        if (payload.TryGetProperty("error", out _))
        {
            return; // No index on this machine; the freshness contract does not apply.
        }

        // search is a SUCCESS even when it could not reach Outlook - losing the indexed
        // answer we already hold would be the worse failure.
        Assert.False(IsError(result), "search must degrade, never fail, when Outlook is unavailable");

        bool degraded = payload.TryGetProperty("degraded", out JsonElement d) && d.GetBoolean();
        string freshness = payload.TryGetProperty("freshness", out JsonElement f) ? f.GetString()! : "live";

        // The two markers must agree with each other and with the sweep block. A result
        // that looks complete but silently lags recent mail is the one failure mode here
        // that misleads rather than merely inconveniences.
        Assert.Equal(degraded, freshness == "index-only");

        if (payload.TryGetProperty("sweep", out JsonElement sweep) &&
            sweep.TryGetProperty("performed", out JsonElement performed))
        {
            Assert.Equal(!performed.GetBoolean(), degraded);
        }

        if (degraded)
        {
            // And it must say so in words the model will relay, not only in a field.
            Assert.True(payload.TryGetProperty("advice", out JsonElement advice), "degraded results must carry advice");
            string joined = advice.GetRawText();
            Assert.Contains("TELL THE USER", joined, StringComparison.Ordinal);
        }
    }

    [Fact]
    public async Task RepeatedCalls_NeverEachPayAFullBudget()
    {
        // The regression this guards: before the liveness gate and the breaker, every
        // request independently rediscovered an unavailable Outlook - measured at 120 s
        // EACH against a wedged one. Whatever the machine state, the fifth call must not
        // cost what the first did.
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(TimeSpan.FromSeconds(240));

        TimeSpan worstAfterFirst = TimeSpan.Zero;
        for (int i = 0; i < 5; i++)
        {
            (_, TimeSpan elapsed) = await CallAsync(client, "list_accounts", new { });
            if (i > 0 && elapsed > worstAfterFirst)
            {
                worstAfterFirst = elapsed;
            }
        }

        Assert.True(
            worstAfterFirst < TimeSpan.FromSeconds(30),
            $"repeat calls must be cheap once the state is known; worst was {worstAfterFirst.TotalSeconds:F1}s");
    }

    [Fact]
    public async Task HealthAlwaysAnswersQuickly_AndStatesOutlooksCondition()
    {
        // outlook_health is asked precisely when things are wrong, so it is the one tool
        // that must never join the failure it is reporting on.
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(TimeSpan.FromSeconds(180));

        (JsonElement result, TimeSpan elapsed) = await CallAsync(client, "outlook_health", new { });

        Assert.False(IsError(result), "health must always produce a report");
        Assert.True(elapsed < TimeSpan.FromSeconds(60), $"health took {elapsed.TotalSeconds:F1}s");

        JsonElement outlook = PayloadOf(result).GetProperty("outlook");

        // It must state Outlook's condition in words, from Windows' own view rather than
        // inferred from our own failures.
        Assert.True(outlook.TryGetProperty("state", out JsonElement state), "health must report outlook.state");
        Assert.Contains(
            state.GetString(),
            new[] { "not running", "starting", "responsive", "not responding" });

        if (outlook.GetProperty("running").GetBoolean())
        {
            Assert.True(outlook.TryGetProperty("responding", out _), "a running Outlook must report whether it responds");
        }
    }
}
