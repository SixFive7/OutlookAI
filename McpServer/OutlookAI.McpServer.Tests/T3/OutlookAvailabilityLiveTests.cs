using System.Diagnostics;
using System.Text.Json;

using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tests.T2;

using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// End-to-end guarantees about what happens when Outlook is not simply available.
/// <para>
/// Written as invariants rather than as assertions about one machine state, on purpose.
/// These hold on a machine with no Outlook at all, on a developer box (Outlook healthy),
/// and were developed against a genuinely wedged Outlook. A test that only holds in one of
/// those would be worse than no test: it would go red for reasons that are not defects,
/// which is how a suite stops being believed.
/// </para>
/// <para>
/// <b>Why this is Category=Live, despite the file having been called ...CiTests.</b> Every
/// test here calls a tool that reaches Outlook for every argument shape - <c>list_accounts</c>,
/// <c>search</c>, <c>outlook_health</c> - and there is no way to write the invariant without
/// one. On a machine that HAS Outlook, that is an attach to a real profile: the store list is
/// enumerated over COM, the freshness sweep walks folders, and health queries the Windows
/// Search index per store. Worse, <c>list_accounts</c> has no liveness escape: a machine with
/// Outlook installed but closed gets it STARTED, because the supervisor's verdict for
/// NotRunning is MayStart. None of that was declared while the tier ran under
/// <c>Category!=Live</c>.
/// </para>
/// <para>
/// Any Outlook profile satisfies these, the dedicated test VM included. They read no mail and
/// write nothing - what they need is an Outlook, which is what <c>Requires=OutlookInstance</c>
/// says, and it is the whole of what they say.
/// </para>
/// <para>
/// <b>Neither may pass having checked nothing (Q101, 2026-10-03).</b> Each used to return GREEN
/// on the machine state it could not check - a healthy Outlook, with no line at all; and ANY
/// error from <c>search</c>, which is exactly what "search must degrade, never fail" exists to
/// catch. The search test now fails on every error except the one where there was nothing to
/// search - this machine's Windows Search index unreachable, by the product's own verdict - and
/// both remaining returns go through <see cref="LivePopulationCoverage"/>: a refusal on a
/// Production profile, a <c>PROVED NOTHING:</c> line on a Portable one. Pinned by
/// <c>T1/LiveEarlyReturnGuardTests</c>.
/// </para>
/// </summary>
[Collection(LiveCollections.McpToolShape)]
[Trait("Category", "Live")]
public sealed class OutlookAvailabilityLiveTests
{
    /// <summary>
    /// What the retry-guidance check reads, named as the Production refusal wraps it. A healthy
    /// Outlook answers <c>list_accounts</c> without an error, so there is no state to check.
    /// </summary>
    internal const string TransientStatePopulation =
        "a transient Outlook state (starting, not responding or unavailable) for list_accounts to report";

    /// <summary>What a reader of a PROVED NOTHING line about that check is to do about it.</summary>
    internal const string TransientStateRemedy =
        "The check reads the error list_accounts returns while Outlook is starting, hung or unavailable, and a "
        + "healthy Outlook returns none - so it runs only on a machine whose Outlook is in that state when the test "
        + "calls: straight after Outlook was closed, while it is still starting, or while it is not responding.";

    /// <summary>What the freshness-contract assertions need, named as the Production refusal wraps it.</summary>
    internal const string AnsweredSearchPopulation =
        "a search answered from a reachable Windows Search index";

    /// <summary>What a reader of a PROVED NOTHING line about the freshness contract is to do about it.</summary>
    internal const string AnsweredSearchRemedy =
        "search failed the way it does when this machine's Windows Search index cannot be reached at all - "
        + "outlook_health reported index.provider as unavailable - so there was no answer to hold to the freshness "
        + "contract. Start the Windows Search service (WSearch) and re-run; a guest with no catalog at all is a "
        + "guest shape of its own (Docs/live-tier-on-the-vm.md section 8, item 21).";

    /// <summary>Error types that mean "not now, try again" rather than "this went wrong".</summary>
    private static readonly string[] TransientTypes =
    {
        "OutlookStarting", "OutlookUnresponsive", "Timeout", "ComHostUnavailable", "OutlookUnavailable",
    };

    private readonly LiveMcpToolShapeFixture _fixture;
    private readonly ITestOutputHelper _output;

    public OutlookAvailabilityLiveTests(LiveMcpToolShapeFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    /// <summary>
    /// True when <paramref name="provider"/> - <c>outlook_health</c>'s <c>index.provider</c> - is
    /// the product's own "the SystemIndex is unreachable" marker, classified by the product's own
    /// rule (<see cref="MailService.ClassifyIndexCurrency"/>) rather than by a second copy of it.
    /// Absent is NOT unreachable: the one error <see cref="SearchAlwaysAnswers_AndSaysWhetherItIsComplete"/>
    /// lets through needs the product to have said so.
    /// </summary>
    internal static bool IndexIsUnreachable(string? provider)
    {
        return provider != null
            && MailService.ClassifyIndexCurrency(provider, null) == MailService.IndexCurrency.Unavailable;
    }

    /// <summary>
    /// The index provider <c>outlook_health</c> reports on the SAME server - whose index tier is the
    /// one <c>search</c> just used, created once and shared - or null when health did not say.
    /// </summary>
    private static async Task<string?> IndexProviderAsync(McpStdioClient client)
    {
        (JsonElement health, TimeSpan _) = await CallAsync(client, "outlook_health", new { });
        return !IsError(health)
            && PayloadOf(health).TryGetProperty("index", out JsonElement index)
            && index.TryGetProperty("provider", out JsonElement provider)
            && provider.ValueKind == JsonValueKind.String
                ? provider.GetString()
                : null;
    }

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
    [Trait("Requires", "OutlookInstance")]
    public async Task ATransientOutlookState_AnswersFastAndCarriesRetryGuidance()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(
            TimeSpan.FromSeconds(180), environment: null, McpStdioClient.OutlookReachingToolsAllowed);

        (JsonElement result, TimeSpan elapsed) = await CallAsync(client, "list_accounts", new { });

        // Whatever the machine's state, a COM-needing tool must not sit on the caller.
        Assert.True(elapsed < TimeSpan.FromSeconds(100), $"took {elapsed.TotalSeconds:F1}s");

        // A healthy Outlook answers without an error, which leaves the retry-guidance check below
        // nothing to read. That used to return GREEN with no line at all; it is the Q57 pattern now
        // (Q101): a refusal on a Production profile, a PROVED NOTHING line on a Portable one.
        IReadOnlyList<JsonElement> transient = LivePopulationCoverage.Require(
            _fixture.Settings,
            IsError(result) ? new[] { result } : Array.Empty<JsonElement>(),
            TransientStatePopulation,
            "the retry-guidance check",
            TransientStateRemedy,
            _output.WriteLine);
        if (transient.Count == 0)
        {
            return;
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

    /// <summary>
    /// What this test is FOR, and what it deliberately stopped asserting.
    /// <para>
    /// It is a correctness test about one contract: a <c>search</c> always answers, and the
    /// answer always says whether it is complete. It used to carry a wall-clock assertion as
    /// well - the search must return inside 100 s - and that half was measuring the
    /// developer's Outlook rather than the product. It passed for months and then failed 4 of
    /// 4 at 139 s against an Outlook that had been up for 40 hours, including against a
    /// mutation of a constant the server process cannot even see. A test that goes red for
    /// reasons that are not defects is how a suite stops being believed, which is the rule
    /// this file's own header states.
    /// </para>
    /// <para>
    /// The timing contract has a better home and already lives there:
    /// <see cref="ATransientOutlookState_AnswersFastAndCarriesRetryGuidance"/> asserts that a
    /// COM-needing tool does not sit on the caller, and the budget composition itself is
    /// pinned arithmetically in T1 <c>BudgetCompositionTests</c> where no mailbox can move it.
    /// </para>
    /// <para>
    /// The client budget is DERIVED rather than left at the old flat 180 s: that number was
    /// below the budget a slow search is entitled to spend, so dropping the assertion without
    /// raising it would have replaced a failed assertion with a cancelled round trip and
    /// proved nothing at all.
    /// </para>
    /// </summary>
    [Fact]
    [Trait("Requires", "OutlookInstance")]
    public async Task SearchAlwaysAnswers_AndSaysWhetherItIsComplete()
    {
        // What a search may legitimately cost end to end: its own composed budget plus the
        // handshake that precedes it on this client.
        TimeSpan clientBudget = TimeSpan.FromMilliseconds(
            MailService.SearchBudgetMs + ComOperationBudgets.HandshakeBudgetMs);

        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(
            clientBudget, environment: null, McpStdioClient.OutlookReachingToolsAllowed);

        (JsonElement result, TimeSpan _) = await CallAsync(
            client, "search", new { query = "invoice", top = 3 });

        JsonElement payload = PayloadOf(result);
        string? searchError = payload.TryGetProperty("error", out JsonElement error) ? error.GetRawText() : null;
        if (searchError != null)
        {
            // ONE error may end this test (Q101): the one where there was nothing to search, because
            // this machine's Windows Search index cannot be reached at all. That is the product's own
            // verdict, read from outlook_health on the same server, whose index tier is the one this
            // search used. Every other error is precisely what "search must degrade, never fail"
            // below exists to catch - and it used to return green just the same.
            string? provider = await IndexProviderAsync(client);
            Assert.True(
                IndexIsUnreachable(provider),
                "search must degrade, never fail - it returned an error, and outlook_health does not report this "
                + $"machine's index as unreachable (index.provider={provider ?? "(not reported)"}): {searchError}");
        }

        // The freshness contract needs a search that ANSWERED; the no-index error above is the one
        // way left not to, and it says so rather than passing (the Q57 pattern, Q101).
        IReadOnlyList<JsonElement> answered = LivePopulationCoverage.Require(
            _fixture.Settings,
            searchError == null ? new[] { payload } : Array.Empty<JsonElement>(),
            AnsweredSearchPopulation,
            "the freshness-contract assertions",
            AnsweredSearchRemedy,
            _output.WriteLine);
        if (answered.Count == 0)
        {
            return;
        }

        // search is a SUCCESS even when it could not reach Outlook - losing the indexed
        // answer we already hold would be the worse failure.
        Assert.False(IsError(result), "search must degrade, never fail, when Outlook is unavailable");

        bool degraded = payload.TryGetProperty("degraded", out JsonElement d) && d.GetBoolean();
        string freshness = payload.TryGetProperty("freshness", out JsonElement f) ? f.GetString()! : "live";

        // THREE states, not two: the sweep ran and covered everything ("live"), it ran and
        // covered part of its scope ("partial"), or it never ran ("index-only").
        Assert.Contains(freshness, new[] { "live", "partial", "index-only" });

        // The two markers must agree with each other and with the sweep block. A result
        // that looks complete but silently lags recent mail is the one failure mode here
        // that misleads rather than merely inconveniences.
        //
        // Re-baselined from Assert.Equal(degraded, freshness == "index-only"): that read
        // "degraded means the sweep did not run", which made a partially-covered sweep a
        // NON-degraded result by definition and pinned exactly the lie this contract exists
        // to prevent. What survives unchanged is the useful half - degraded is true iff the
        // answer is not fully fresh - now stated against "live" rather than against one of
        // the two ways of failing it.
        Assert.Equal(degraded, freshness != "live");

        if (payload.TryGetProperty("sweep", out JsonElement sweep) &&
            sweep.TryGetProperty("performed", out JsonElement performed))
        {
            bool ran = performed.GetBoolean();
            bool notNeeded = sweep.TryGetProperty("notNeeded", out JsonElement skipped) && skipped.GetBoolean();
            bool gapsReported = sweep.TryGetProperty("coverageGaps", out JsonElement gaps)
                && gaps.ValueKind == JsonValueKind.Array
                && gaps.GetArrayLength() > 0;

            // A sweep that COULD NOT run is index-only, whatever else is in the block; a
            // sweep that ran is degraded exactly when it reports coverage gaps, and the
            // gap list is the machine-readable reason - so an agent never has to read prose
            // to find out that an answer is partial.
            //
            // A sweep that did not NEED to run is the third state and is not index-only:
            // its search's window ends before the index frontier, so there was nothing for
            // it to find and the answer is complete. (This query sets no 'before' bound, so
            // it never reaches that state here - the clause keeps the invariant honest for
            // the searches that do.)
            Assert.Equal(!ran && !notNeeded, freshness == "index-only");
            Assert.False(!ran && gapsReported, "a sweep that never ran cannot report coverage gaps");
            if (ran)
            {
                Assert.Equal(gapsReported, freshness == "partial");
                Assert.Equal(gapsReported, degraded);
            }
        }

        if (degraded)
        {
            // And it must say so in words the model will relay, not only in a field.
            Assert.True(payload.TryGetProperty("advice", out JsonElement advice), "degraded results must carry advice");
            string joined = advice.GetRawText();

            // "TELL THE USER" is the alarm the not-run case raises, and it stays pinned
            // there. A partial sweep does not shout: its advice names the specific hole,
            // what is missing because of it and how to close it, which is more actionable
            // than an alarm - so what is required of it is that it says something at all.
            if (freshness == "index-only")
            {
                Assert.Contains("TELL THE USER", joined, StringComparison.Ordinal);
            }
            else
            {
                Assert.Contains("Freshness sweep", joined, StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    [Trait("Requires", "OutlookInstance")]
    public async Task RepeatedCalls_NeverEachPayAFullBudget()
    {
        // The regression this guards: before the liveness gate and the breaker, every
        // request independently rediscovered an unavailable Outlook - measured at 120 s
        // EACH against a wedged one. Whatever the machine state, the fifth call must not
        // cost what the first did.
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(
            TimeSpan.FromSeconds(240), environment: null, McpStdioClient.OutlookReachingToolsAllowed);

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
    [Trait("Requires", "OutlookInstance")]
    public async Task HealthAlwaysAnswersQuickly_AndStatesOutlooksCondition()
    {
        // outlook_health is asked precisely when things are wrong, so it is the one tool
        // that must never join the failure it is reporting on.
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(
            TimeSpan.FromSeconds(180), environment: null, McpStdioClient.OutlookReachingToolsAllowed);

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
