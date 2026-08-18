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
/// the client at any length. They are therefore MEASURED AND REPORTED HERE, NEVER FAILED:
/// failing a build over text the client delivers whole would reject a description that
/// arrives intact, which is a false failure however tidy the number looks. The sizes stay in
/// the table because they are the evidence a future client bump would be re-read against.
/// </para>
/// <para>
/// There is no warn tier either, and its absence is a decision rather than an omission. A
/// threshold below the cap fires on strings that arrive whole, every run, forever - three of
/// them here, none of which will ever change - which trains a reader to skip the channel on
/// the day it finally carries a real cut. The rule this file implements is the narrow one:
/// fail the instant something would be truncated, allow everything that fits, warn about
/// nothing.
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
    /// The budget a surface is judged against, or <c>null</c> for a surface the client was
    /// measured not to cut. Only two surfaces have a budget at all, because only two are
    /// truncated; a parameter description is measured, printed and never judged.
    /// <para>
    /// A nullable return rather than a large sentinel on purpose: a sentinel would still be a
    /// number, and every reader downstream would have to know which number means "no limit".
    /// </para>
    /// </summary>
    private static int? BudgetFor(string surface) =>
        surface == "parameter" ? null : ClientTruncationBudget;

    private readonly ITestOutputHelper _output;

    public DescriptionBudgetCiTests(ITestOutputHelper output)
    {
        _output = output;
    }

    /// <summary>
    /// Every description on the wire - the initialize result's <c>instructions</c>, every
    /// tool <c>description</c>, and every parameter <c>description</c> nested anywhere inside
    /// an <c>inputSchema</c> - is measured and reported. The first two are failed the instant
    /// they cross the measured client cap; the third has no cap to cross, so it is only ever
    /// reported.
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
        // A surface with no budget (parameter descriptions, which the client does not cut) is
        // not a pass here - it is not a candidate at all.
        List<WireString> overBudget = measured
            .Where(s => s.Budget is int budget && s.Measured > budget)
            .OrderByDescending(s => s.Measured - s.Budget!.Value)
            .ToList();

        Assert.True(overBudget.Count == 0, DescribeOverBudget(overBudget));
    }

    /// <summary>
    /// Writes the whole measured surface, largest first. Emitted on every run, and a passing
    /// run is exactly when it is worth having: it is the only record of how large each string
    /// actually is, which is the number a future client bump gets re-read against.
    /// <para>
    /// It reports and it does not judge. Nothing here says a string is close to anything -
    /// see the type doc for why a threshold below the cap is worse than no threshold.
    /// </para>
    /// <para>
    /// UTF-8 bytes are in the table but budgeted against nothing. The client counts UTF-16
    /// code units and never bytes (measured 2026-08-18, see
    /// <see cref="ClientTruncationBudget"/>), so a byte column is information - useful if a
    /// future client is ever found to count differently - and not a limit.
    /// </para>
    /// </summary>
    private void ReportTable(IReadOnlyList<WireString> measured)
    {
        _output.WriteLine(
            $"MCP description sizes, in UTF-16 code units (string.Length). Budget: {ClientTruncationBudget} "
            + "for the initialize instructions and each tool description (MEASURED Claude Code 2.1.234 "
            + "truncation, 2026-08-18). Parameter descriptions are measured and reported but have NO budget: "
            + "the same capture put 20,000 characters through one intact. UTF-8 bytes shown for information "
            + "only: the client never counts them.");
        _output.WriteLine(
            "If a cut ever happens it is invisible here and visible to the model: the string arrives as a "
            + $"{ClientTruncationBudget}-unit prefix plus the marker \"{ClientTruncationMarker}\", "
            + $"{TruncatedStringLength} units in total. That marker is what to grep a transcript for.");
        _output.WriteLine(string.Empty);

        IReadOnlyList<string> rows = TableRows(measured);
        foreach (string row in rows)
        {
            _output.WriteLine(row);
        }

        PublishToCiJobSummary(rows);
    }

    /// <summary>
    /// The size table as text, header first. One renderer for both channels, so the CI
    /// summary cannot drift from what the test log says.
    /// </summary>
    private static IReadOnlyList<string> TableRows(IReadOnlyList<WireString> measured)
    {
        var rows = new List<string> { $"{"units",6}  {"bytes",6}  {"budget",6}  surface     name" };

        foreach (WireString entry in measured.OrderByDescending(s => s.Measured).ThenBy(s => s.Label, StringComparer.Ordinal))
        {
            // "none" rather than a blank or a dash: a surface with no budget is a measured
            // fact about this client, not a column somebody forgot to fill in.
            string budget = entry.Budget is int b ? b.ToString(CultureInfo.InvariantCulture) : "none";
            rows.Add($"{entry.Measured,6}  {entry.Utf8Bytes,6}  {budget,6}  {entry.Surface,-10}  {entry.Label}");
        }

        return rows;
    }

    /// <summary>
    /// The size table only earns its keep if a PASSING run shows it, and `dotnet test` at its
    /// default verbosity shows a passing test's output nowhere at all. On GitHub Actions the
    /// job summary is the channel that does not need the whole suite switched to a verbose
    /// logger, so the table is written there when CI provides it - collapsed, because it is a
    /// record to consult and not a thing to act on. Purely additive and best-effort: no CI, or
    /// an unwritable summary file, changes nothing.
    /// </summary>
    private static void PublishToCiJobSummary(IReadOnlyList<string> rows)
    {
        string? summaryPath = Environment.GetEnvironmentVariable("GITHUB_STEP_SUMMARY");
        if (string.IsNullOrWhiteSpace(summaryPath))
        {
            return;
        }

        try
        {
            var summary = new StringBuilder();
            summary.AppendLine("<details><summary>MCP description sizes (UTF-16 code units)</summary>");
            summary.AppendLine();
            summary.AppendLine(
                $"Budget {ClientTruncationBudget} for the initialize instructions and each tool description "
                + "(measured Claude Code 2.1.234 truncation, 2026-08-18). Parameter descriptions are not cut "
                + "by the client at any length, so they are reported without a budget. The run fails only on "
                + "a string that would actually be truncated.");
            summary.AppendLine();
            summary.AppendLine("```");
            foreach (string row in rows)
            {
                summary.AppendLine(row);
            }

            summary.AppendLine("```");
            summary.AppendLine();
            summary.AppendLine("</details>");

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
            // Only budgeted surfaces reach here, so the budget is present by construction.
            int budget = entry.Budget!.Value;

            message.AppendLine(
                $"  {entry.Surface} '{entry.Label}': {entry.Measured} units "
                + $"({entry.Utf8Bytes} UTF-8 bytes, which the client never counts) - "
                + $"{entry.Measured - budget} OVER the MEASURED client cap of {budget}. Claude Code cuts "
                + "this silently and mid-sentence, and everything past the cut never reaches the model. "
                + $"The cut would land at: \"...{Excerpt(entry.Text, budget)}\"");
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

        /// <summary>
        /// Which budget applies to this surface, or null where the client does not cut at
        /// all - see <see cref="BudgetFor"/>.
        /// </summary>
        public int? Budget => BudgetFor(Surface);
    }
}
