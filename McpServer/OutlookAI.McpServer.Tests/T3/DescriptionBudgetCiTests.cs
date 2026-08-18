using System.Globalization;
using System.Text;
using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// Wire-length guardrail for every string this server sends a client as prose.
/// <para>
/// MCP clients cap the strings a server may inject. Claude Code cuts a tool
/// <c>description</c> and the server <c>instructions</c> at 2,048 UTF-16 code units - not a
/// documentation claim but a measurement, taken 2026-08-18 against client version 2.1.234 by
/// intercepting the client's own outbound <c>POST /v1/messages</c> and reading the
/// <c>tools</c> array the model actually receives (see <see cref="ClientTruncationBudget"/>
/// for the numbers and the caveats). The truncation is POSITIONAL and SILENT: the client cuts
/// wherever the limit falls - mid-sentence, mid-word - and appends
/// <see cref="ClientTruncationMarker"/>, which the model sees and this server never can. No
/// error, no notification, no re-request. A description that grew past the cap keeps looking
/// correct in source and in the JSON payload while the half that matters (typically the
/// trailing paragraphs: degradation handling, scope rules, cost warnings) is invisible to the
/// agent that has to act on it.
/// </para>
/// <para>
/// That same measurement showed <c>inputSchema.properties[*].description</c> is NOT capped by
/// the client at any length. This file still budgets those, at
/// <see cref="HouseParameterBudget"/> - a HOUSE limit with its own reasoning, kept as a
/// separate constant from the measured client cap precisely so nobody later cites it as
/// client behaviour.
/// </para>
/// <para>
/// So the measurement is taken FROM THE WIRE, not from source constants: these
/// descriptions are assembled from concatenated literals and shared constants
/// (<c>OutlookTools</c> reuses hint constants across the draft tools), so a test that
/// measured source literals would miss exactly the cases that break. The enumeration is
/// dynamic - a tool or parameter added later is covered with no edit here.
/// </para>
/// <para>
/// Maintenance note for anyone writing another stdio probe: do NOT feed the server from a
/// redirected FILE. stdin then hits EOF the instant the file is consumed and the server
/// shuts down before it answers. The pipe has to stay open, which is what
/// <see cref="McpStdioClient"/> does.
/// </para>
/// CI-safe: initialize + tools/list need neither Outlook nor the search index.
/// </summary>
public sealed class DescriptionBudgetCiTests
{
    /// <summary>
    /// The measured client cap, in UTF-16 code units, for the two surfaces Claude Code
    /// actually cuts: the initialize result's <c>instructions</c> and each tool
    /// <c>description</c>. The unit is <c>string.Length</c> - see
    /// <see cref="WireString.Measured"/>.
    /// </summary>
    /// <remarks>
    /// <para>
    /// MEASURED, NOT INFERRED. 2026-08-18, Claude Code 2.1.234 on Windows 11: the client's
    /// outbound <c>POST /v1/messages</c> was captured at a local HTTP endpoint and the
    /// <c>tools</c> array the model receives was read byte for byte. Reproduced against two
    /// models, byte-identical, because the cut is client-side. The documentation's phrase -
    /// "truncates tool descriptions and server instructions at 2KB each" - is ambiguous about
    /// the unit and about what "each" counts. The capture settles both:
    /// </para>
    /// <list type="bullet">
    /// <item><description>
    /// PER STRING, not per serialized tool. There is no per-tool bucket at all: an entry of
    /// 17,411 bytes and one of 20,172 bytes both arrived intact. So trimming a description
    /// really does buy room, rather than moving text from one capped bucket into the same one.
    /// </description></item>
    /// <item><description>
    /// UTF-16 CODE UNITS. Bytes are never counted: a 2,048-character description weighing
    /// 6,004 UTF-8 bytes arrives whole, and two strings of very different byte lengths were
    /// cut at the same CHARACTER offset. Units rather than code points - 1,539 code points
    /// spread over 3,000 units was cut - and the cut is surrogate-aware, taking 2,047 rather
    /// than splitting a pair.
    /// </description></item>
    /// <item><description>
    /// The predicate is <c>length &gt; 2048</c>. Measured as a triple in one run: 2,047
    /// intact, 2,048 intact, 2,049 cut.
    /// </description></item>
    /// <item><description>
    /// No total budget across <c>tools/list</c>: 202 tools totalling 348,314 bytes of
    /// serialized entries passed in one request, nothing dropped and nothing cut. That
    /// establishes no cap at 348 KB, not that none exists above it.
    /// </description></item>
    /// </list>
    /// <para>
    /// ONE CLIENT AT ONE VERSION, and nothing watches it for us. This is not part of the
    /// Model Context Protocol specification, which puts no length limit on <c>description</c>
    /// or <c>instructions</c> at all; other hosts cap differently or not at all, and none of
    /// them were measured. There is no version header, no notification and no server-side
    /// signal when the behaviour changes, so this number is only as current as its date -
    /// re-measure at client-bump time. The change worth re-measuring FOR is a release that
    /// introduces a per-tool bucket: servers with large schemas would fail it on day one,
    /// silently, with the description intact and the schema cut.
    /// </para>
    /// <para>
    /// <c>internal</c>, not private: <c>SearchSchemaCiTests</c> asserts that the search
    /// tool's degraded-results instruction lands inside this budget, and it used to do so
    /// against its own bare <c>2048</c> in another file of the same assembly. Two copies of
    /// one client's cap, and only one of them carried the explanation.
    /// </para>
    /// </remarks>
    internal const int ClientTruncationBudget = 2048;

    /// <summary>
    /// The budget applied to every <c>inputSchema.properties[*].description</c>. A HOUSE
    /// LIMIT - the client does not cap these at all.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The same 2026-08-18 capture put 2,600 characters and then 20,000 characters through a
    /// parameter description intact. The documentation's silence about <c>inputSchema</c> is
    /// accurate rather than an omission, and this guardrail's earlier application of 2048 here
    /// was an extrapolation - now known to have been unnecessary.
    /// </para>
    /// <para>
    /// Kept anyway, at the same value, for two reasons. It floats with a client version this
    /// project neither controls nor gets a signal about, so a release that starts cutting
    /// schemas should find us already inside the limit rather than find us in production. And
    /// this server's schema has exactly the shape that would suffer worst: <c>BodyHtmlHint</c>
    /// is ONE constant reused across five drafting tools, so a single over-long shared
    /// parameter description would not be one silent truncation but five.
    /// </para>
    /// <para>
    /// It is a separate named constant, not a second use of
    /// <see cref="ClientTruncationBudget"/>, so the label survives the next reader: a failure
    /// on this budget is a house-style failure, and citing it as documented client behaviour
    /// would be wrong.
    /// </para>
    /// </remarks>
    private const int HouseParameterBudget = 2048;

    /// <summary>
    /// What a cut looks like when it reaches the model: the published string's exact prefix,
    /// then this literal - U+2026 HORIZONTAL ELLIPSIS, space, <c>[truncated]</c> - 13 UTF-16
    /// code units, so a cut string arrives at <see cref="TruncatedStringLength"/>.
    /// </summary>
    /// <remarks>
    /// Recorded because NO TEST CAN OBSERVE IT. The client appends the marker after this
    /// server's JSON-RPC response has already left the process; there is no error, no
    /// notification and no re-request, so a server cannot detect its own truncation - only the
    /// model can see the marker, which makes "did that arrive whole?" answerable by asking and
    /// unanswerable by logging. The constant exists so the exact string a human would grep for
    /// in a transcript, when a description looks cut, is written down in the file that owns the
    /// budget rather than rediscovered. Escaped rather than pasted, so this file stays ASCII.
    /// </remarks>
    internal const string ClientTruncationMarker = "\u2026 [truncated]";

    /// <summary>
    /// The length a truncated string reaches the model at: a 2,048-unit prefix plus the
    /// 13-unit marker. Derived rather than restated, so the two cannot drift apart.
    /// </summary>
    internal static readonly int TruncatedStringLength = ClientTruncationBudget + ClientTruncationMarker.Length;

    /// <summary>
    /// Warn-only tier. Something sitting at 1600 of 2048 is one added paragraph from being
    /// cut, and the cut lands without a diagnostic - so the run says so out loud while it
    /// is still cheap to fix, without failing a build over a description that currently
    /// arrives intact.
    /// </summary>
    private const double WarnFraction = 0.75;

    /// <summary>
    /// The budget a surface is judged against: the measured client cap for the two surfaces
    /// the client actually cuts, the house limit for the one it does not. The two are equal
    /// today; the split exists so either can move without implying the other, and so a
    /// failure message can say which kind of limit was crossed.
    /// </summary>
    private static int BudgetFor(string surface) =>
        surface == "parameter" ? HouseParameterBudget : ClientTruncationBudget;

    private static int WarnThresholdFor(int budget) => (int)(budget * WarnFraction);

    private readonly ITestOutputHelper _output;

    public DescriptionBudgetCiTests(ITestOutputHelper output)
    {
        _output = output;
    }

    /// <summary>
    /// Every description on the wire - the initialize result's <c>instructions</c>, every
    /// tool <c>description</c>, and every parameter <c>description</c> nested anywhere inside
    /// an <c>inputSchema</c> - fits its budget: the measured client cap for the first two,
    /// the house limit for the third.
    /// </summary>
    [Fact]
    public async Task EveryDescriptionOnTheWire_FitsTheClientTruncationBudget()
    {
        await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();
        JsonElement toolsResult = await client.RoundTripAsync("tools/list", new { });

        IReadOnlyList<WireString> measured = CollectWireStrings(client.InitializeResult, toolsResult);

        // Guard against the enumeration going vacuous: a walk that silently stopped
        // finding descriptions would "pass" forever. These floors are deliberately far
        // below the real counts - they catch a broken walk, not a surface change.
        Assert.Contains(measured, s => s.Surface == "instructions");
        Assert.True(
            measured.Count(s => s.Surface == "tool") >= 20,
            $"expected every advertised tool to carry a description; found {measured.Count(s => s.Surface == "tool")}");
        Assert.True(
            measured.Count(s => s.Surface == "parameter") >= 50,
            $"expected the parameter walk to reach the whole schema surface; found {measured.Count(s => s.Surface == "parameter")}");

        ReportTable(measured);

        // Canary, and an honest statement of what a canary can be here. A cut string reaches
        // the model as its prefix plus ClientTruncationMarker, appended AFTER this server's
        // response has left the process - so nothing in this file can ever observe our own
        // truncation. What this DOES catch is the marker travelling the other way: text
        // copied out of a transcript that was already cut and pasted back into a description
        // or a schema, which would then ship the marker to every client as if it were prose.
        List<WireString> carryingMarker = measured
            .Where(s => s.Text.Contains(ClientTruncationMarker, StringComparison.Ordinal))
            .ToList();

        Assert.True(
            carryingMarker.Count == 0,
            $"a shipped string contains the client's truncation marker (\"{ClientTruncationMarker}\"), "
            + "which means already-truncated text was copied back into source: "
            + string.Join(", ", carryingMarker.Select(s => $"{s.Surface} '{s.Label}'")));

        // The predicate is the measured one: cut when length > 2048, so exactly 2048 passes.
        List<WireString> overBudget = measured
            .Where(s => s.Measured > s.Budget)
            .OrderByDescending(s => s.Measured - s.Budget)
            .ToList();

        Assert.True(overBudget.Count == 0, DescribeOverBudget(overBudget));
    }

    /// <summary>
    /// Writes the whole measured surface, largest first, plus the warn-tier callouts.
    /// Emitted on every run - a passing run is exactly when this table is worth having,
    /// because it shows how much head-room is left before the next paragraph gets cut.
    /// <para>
    /// UTF-8 bytes are still in the table but are no longer budgeted against anything. The
    /// client counts UTF-16 code units and never bytes (measured 2026-08-18, see
    /// <see cref="ClientTruncationBudget"/>), so a byte column is information - useful if a
    /// future client is ever found to count differently - and not a limit.
    /// </para>
    /// </summary>
    private void ReportTable(IReadOnlyList<WireString> measured)
    {
        _output.WriteLine(
            $"MCP description budget, in UTF-16 code units (string.Length): {ClientTruncationBudget} for "
            + $"instructions and tool descriptions (MEASURED Claude Code 2.1.234 truncation, 2026-08-18); "
            + $"{HouseParameterBudget} for parameter descriptions (HOUSE limit - the client does not cap "
            + $"these at all). Warn at {WarnFraction:P0} of whichever applies. UTF-8 bytes shown for "
            + "information only: the client never counts them.");
        _output.WriteLine(
            "If a cut ever happens it is invisible here and visible to the model: the string arrives as a "
            + $"{ClientTruncationBudget}-unit prefix plus the marker \"{ClientTruncationMarker}\", "
            + $"{TruncatedStringLength} units in total. That marker is what to grep a transcript for.");
        _output.WriteLine(string.Empty);
        _output.WriteLine($"{"units",6}  {"bytes",6}  {"budget",6}  {"%",5}  surface     name");

        foreach (WireString entry in measured.OrderByDescending(s => s.Measured).ThenBy(s => s.Label, StringComparer.Ordinal))
        {
            _output.WriteLine(
                $"{entry.Measured,6}  {entry.Utf8Bytes,6}  {entry.Budget,6}  "
                + $"{entry.PercentOfBudget,4:F0}%  {entry.Surface,-10}  {entry.Label}");
        }

        List<WireString> warnings = measured
            .Where(s => s.Measured >= WarnThresholdFor(s.Budget) && s.Measured <= s.Budget)
            .OrderByDescending(s => s.PercentOfBudget)
            .ToList();

        _output.WriteLine(string.Empty);
        if (warnings.Count == 0)
        {
            _output.WriteLine($"WARN TIER: nothing is above {WarnFraction:P0} of its budget.");
            return;
        }

        var warningLines = new List<string>();
        foreach (WireString warning in warnings)
        {
            // A parameter over the warn line is not the same event as a tool description over
            // it: one is approaching a limit this project chose, the other a limit the client
            // enforces. Saying which keeps the house limit from being read as client behaviour
            // in the one place a maintainer is most likely to meet it.
            string consequence = warning.Surface == "parameter"
                ? "The client does NOT cut parameter descriptions (measured 2026-08-18, 2.1.234) - this is "
                  + "the house limit, kept so a client release that starts cutting schemas finds us already "
                  + "inside it."
                : "It still arrives intact, but the client truncates silently and mid-sentence - move detail "
                  + "into the runtime payload before it crosses.";

            string line =
                $"WARNING: {warning.Surface} '{warning.Label}' is {warning.Measured} units of {warning.Budget} "
                + $"({warning.PercentOfBudget:F0}% of budget, {warning.Budget - warning.Measured} left). "
                + consequence;
            warningLines.Add(line);
            _output.WriteLine(line);

            // Also on the test host's stderr: ITestOutputHelper reaches an IDE runner, the
            // TRX log and `--logger "console;verbosity=detailed"`, but VSTest's default
            // console reporter prints a passing test's output nowhere.
            Console.Error.WriteLine(line);
        }

        PublishToCiJobSummary(warningLines);
    }

    /// <summary>
    /// The warn tier only earns its keep if a PASSING run shows it, and `dotnet test` at
    /// its default verbosity shows a passing test's output nowhere at all. On GitHub
    /// Actions the job summary is the channel that does not need the whole suite switched
    /// to a verbose logger, so the warnings are written there when CI provides it. Purely
    /// additive and best-effort: no CI, or an unwritable summary file, changes nothing.
    /// </summary>
    private static void PublishToCiJobSummary(IReadOnlyList<string> warningLines)
    {
        if (warningLines.Count == 0)
        {
            return;
        }

        string? summaryPath = Environment.GetEnvironmentVariable("GITHUB_STEP_SUMMARY");
        if (string.IsNullOrWhiteSpace(summaryPath))
        {
            return;
        }

        try
        {
            var summary = new StringBuilder();
            summary.AppendLine("### MCP description budget");
            summary.AppendLine();
            summary.AppendLine(
                $"Inside budget ({ClientTruncationBudget} UTF-16 code units for instructions and tool "
                + $"descriptions - measured Claude Code 2.1.234 truncation; {HouseParameterBudget} for "
                + $"parameter descriptions - house limit), but past the {WarnFraction:P0} warn line:");
            summary.AppendLine();
            foreach (string line in warningLines)
            {
                summary.Append("- ").AppendLine(line);
            }

            File.AppendAllText(summaryPath, summary.ToString());
        }
        catch (Exception)
        {
            // A diagnostic channel must never be able to fail a test run.
        }
    }

    private static string DescribeOverBudget(IReadOnlyList<WireString> overBudget)
    {
        var message = new StringBuilder();
        message.Append(overBudget.Count.ToString(CultureInfo.InvariantCulture));
        message.Append(overBudget.Count == 1 ? " description exceeds " : " descriptions exceed ");
        message.AppendLine("its budget, measured in UTF-16 code units:");

        foreach (WireString entry in overBudget)
        {
            string consequence = entry.Surface == "parameter"
                ? $"HOUSE limit of {HouseParameterBudget} - the client does not cut parameter descriptions "
                  + "(measured 2026-08-18, Claude Code 2.1.234), so nothing is being lost today; the limit is "
                  + "kept because it floats with a client version we do not control"
                : $"MEASURED client cap of {ClientTruncationBudget} - Claude Code cuts this silently and "
                  + "mid-sentence, and everything past the cut never reaches the model";

            message.AppendLine(
                $"  {entry.Surface} '{entry.Label}': {entry.Measured} units "
                + $"({entry.Utf8Bytes} UTF-8 bytes, which the client never counts) - "
                + $"{entry.Measured - entry.Budget} OVER the {consequence}. "
                + $"({entry.PercentOfBudget:F0}% of budget.) "
                + $"The cut would land at: \"...{Excerpt(entry.Text, entry.Budget)}\"");
        }

        message.AppendLine("Shorten the description, or move the detail into the tool's runtime payload");
        message.Append("(advice/scope/sweep blocks) or onto per-parameter descriptions, which carry their own budget.");
        return message.ToString();
    }

    /// <summary>
    /// Shows where the cut lands: the 40 code units either side of the budget mark. Only
    /// called for strings already known to be longer than the budget, which is now always
    /// true of an over-budget entry - the measure and the slice index are the same unit.
    /// </summary>
    private static string Excerpt(string text, int budget)
    {
        int start = Math.Max(0, budget - 40);
        int end = Math.Min(text.Length, budget + 40);
        return (text[start..budget] + " | " + text[budget..end]).ReplaceLineEndings(" ");
    }

    /// <summary>
    /// Walks the three surfaces a client reads as prose. Dynamic on purpose: tools and
    /// parameters are discovered from the live <c>tools/list</c> answer, so nothing needs
    /// editing here when the surface grows.
    /// </summary>
    private static IReadOnlyList<WireString> CollectWireStrings(JsonElement initializeResult, JsonElement toolsResult)
    {
        var collected = new List<WireString>();

        // Surface 1: the initialize result's instructions - injected passively into every
        // session at start (D36), so it is capped by the same client rule.
        if (initializeResult.TryGetProperty("instructions", out JsonElement instructions)
            && instructions.ValueKind == JsonValueKind.String)
        {
            collected.Add(new WireString("instructions", "initialize.instructions", instructions.GetString()!));
        }

        foreach (JsonElement tool in toolsResult.GetProperty("result").GetProperty("tools").EnumerateArray())
        {
            string name = tool.GetProperty("name").GetString()!;

            // Surface 2: the tool description.
            if (tool.TryGetProperty("description", out JsonElement toolDescription)
                && toolDescription.ValueKind == JsonValueKind.String)
            {
                collected.Add(new WireString("tool", name, toolDescription.GetString()!));
            }

            // Surface 3: every parameter description in the schema, at any nesting depth.
            if (tool.TryGetProperty("inputSchema", out JsonElement inputSchema))
            {
                CollectSchemaDescriptions(inputSchema, name, collected);
            }
        }

        return collected;
    }

    /// <summary>
    /// Recursive so nested object parameters are covered too (manage_signature's
    /// <c>set_default_for</c> is an object with its own described members, and a future
    /// array/anyOf parameter would nest the same way).
    /// </summary>
    private static void CollectSchemaDescriptions(JsonElement schema, string path, List<WireString> into)
    {
        if (schema.ValueKind != JsonValueKind.Object)
        {
            return;
        }

        if (schema.TryGetProperty("description", out JsonElement description)
            && description.ValueKind == JsonValueKind.String)
        {
            into.Add(new WireString("parameter", path, description.GetString()!));
        }

        if (schema.TryGetProperty("properties", out JsonElement properties)
            && properties.ValueKind == JsonValueKind.Object)
        {
            foreach (JsonProperty property in properties.EnumerateObject())
            {
                CollectSchemaDescriptions(property.Value, path + "." + property.Name, into);
            }
        }

        if (schema.TryGetProperty("items", out JsonElement items))
        {
            CollectSchemaDescriptions(items, path + "[]", into);
        }

        if (schema.TryGetProperty("additionalProperties", out JsonElement additional))
        {
            CollectSchemaDescriptions(additional, path + ".*", into);
        }

        foreach (string keyword in new[] { "anyOf", "oneOf", "allOf" })
        {
            if (!schema.TryGetProperty(keyword, out JsonElement branch) || branch.ValueKind != JsonValueKind.Array)
            {
                continue;
            }

            int index = 0;
            foreach (JsonElement alternative in branch.EnumerateArray())
            {
                CollectSchemaDescriptions(alternative, $"{path}.{keyword}[{index++}]", into);
            }
        }

        foreach (string keyword in new[] { "$defs", "definitions" })
        {
            if (!schema.TryGetProperty(keyword, out JsonElement defs) || defs.ValueKind != JsonValueKind.Object)
            {
                continue;
            }

            foreach (JsonProperty definition in defs.EnumerateObject())
            {
                CollectSchemaDescriptions(definition.Value, $"{path}.{keyword}.{definition.Name}", into);
            }
        }
    }

    /// <summary>One measured string from the wire.</summary>
    private sealed record WireString(string Surface, string Label, string Text)
    {
        /// <summary>
        /// UTF-8 bytes. Reported, never budgeted: the client does not count them.
        /// </summary>
        public int Utf8Bytes => Encoding.UTF8.GetByteCount(Text);

        /// <summary>
        /// The size the budget is judged on: UTF-16 code units, which in C# is
        /// <c>string.Length</c> and in the client's own JavaScript is <c>String.length</c> -
        /// the same unit on both sides.
        /// <para>
        /// This used to be <c>max(chars, UTF-8 bytes)</c>, because the documentation's "2KB"
        /// did not say which unit it meant and the two diverge on the first non-ASCII
        /// character. The 2026-08-18 capture settled it: bytes are never counted, and a
        /// 2,048-character description weighing 6,004 bytes arrives whole. Failing on bytes
        /// only ever produced FALSE failures - it would reject text the client delivers
        /// intact - so the change is a correction, not a relaxation. This project's house
        /// style avoids em dashes and curly quotes, so on today's surface the two units are
        /// equal for every string; the guard is now right rather than accidentally harmless.
        /// </para>
        /// </summary>
        public int Measured => Text.Length;

        /// <summary>Which budget applies to this surface - see <see cref="BudgetFor"/>.</summary>
        public int Budget => BudgetFor(Surface);

        public double PercentOfBudget => Measured * 100.0 / Budget;
    }
}
