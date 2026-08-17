using System.Globalization;
using System.Text;
using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// Wire-length guardrail for every string this server sends a client as prose.
/// <para>
/// MCP clients cap the strings a server may inject. Claude Code's own documentation states
/// it "truncates tool descriptions and server instructions at 2KB each". That truncation is
/// POSITIONAL and SILENT: the client cuts wherever the limit falls - mid-sentence,
/// mid-word - and everything after the cut simply never reaches the model. No error, no
/// warning, no shortened-but-coherent summary. A description that grew past the cap keeps
/// looking correct in source and in the JSON payload while the half that matters
/// (typically the trailing paragraphs: degradation handling, scope rules, cost warnings)
/// is invisible to the agent that has to act on it.
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
    /// The per-string budget, in whichever unit measures larger (see <see cref="WireString.Measured"/>).
    /// <para>
    /// 2048 comes from Claude Code's MCP documentation, which states that it "truncates
    /// tool descriptions and server instructions at 2KB each". This is CLIENT behaviour of
    /// one client - it is NOT part of the Model Context Protocol specification, which puts
    /// no length limit on <c>description</c> or <c>instructions</c> at all. Other hosts cap
    /// differently or not at all. The number is pinned here because Claude Code is this
    /// server's primary client and because the failure it produces is silent.
    /// </para>
    /// <para>
    /// <c>internal</c>, not private: <c>SearchSchemaCiTests</c> asserts that the search
    /// tool's degraded-results instruction lands inside this budget, and it used to do so
    /// against its own bare <c>2048</c> in another file of the same assembly. Two copies of
    /// one client's undocumented cap, and only one of them carried the explanation.
    /// </para>
    /// </summary>
    internal const int ClientTruncationBudget = 2048;

    /// <summary>
    /// Warn-only tier. Something sitting at 1600 of 2048 is one added paragraph from being
    /// cut, and the cut lands without a diagnostic - so the run says so out loud while it
    /// is still cheap to fix, without failing a build over a description that currently
    /// arrives intact.
    /// </summary>
    private const double WarnFraction = 0.75;

    private const int WarnThreshold = (int)(ClientTruncationBudget * WarnFraction);

    private readonly ITestOutputHelper _output;

    public DescriptionBudgetCiTests(ITestOutputHelper output)
    {
        _output = output;
    }

    /// <summary>
    /// Every description on the wire - the initialize result's <c>instructions</c>, every
    /// tool <c>description</c>, and every parameter <c>description</c> nested anywhere
    /// inside an <c>inputSchema</c> - fits inside the client's truncation budget.
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

        List<WireString> overBudget = measured
            .Where(s => s.Measured > ClientTruncationBudget)
            .OrderByDescending(s => s.Measured)
            .ToList();

        Assert.True(overBudget.Count == 0, DescribeOverBudget(overBudget));
    }

    /// <summary>
    /// Writes the whole measured surface, largest first, plus the warn-tier callouts.
    /// Emitted on every run - a passing run is exactly when this table is worth having,
    /// because it shows how much head-room is left before the next paragraph gets cut.
    /// </summary>
    private void ReportTable(IReadOnlyList<WireString> measured)
    {
        _output.WriteLine(
            $"MCP description budget: {ClientTruncationBudget} (Claude Code client truncation), "
            + $"warn at {WarnThreshold} ({WarnFraction:P0}). Measured = max(chars, UTF-8 bytes).");
        _output.WriteLine(string.Empty);
        _output.WriteLine($"{"measured",8}  {"chars",6}  {"bytes",6}  {"%budget",7}  surface     name");

        foreach (WireString entry in measured.OrderByDescending(s => s.Measured).ThenBy(s => s.Label, StringComparer.Ordinal))
        {
            _output.WriteLine(
                $"{entry.Measured,8}  {entry.Chars,6}  {entry.Utf8Bytes,6}  "
                + $"{entry.PercentOfBudget(ClientTruncationBudget),6:F0}%  {entry.Surface,-10}  {entry.Label}");
        }

        List<WireString> warnings = measured
            .Where(s => s.Measured >= WarnThreshold && s.Measured <= ClientTruncationBudget)
            .OrderByDescending(s => s.Measured)
            .ToList();

        _output.WriteLine(string.Empty);
        if (warnings.Count == 0)
        {
            _output.WriteLine($"WARN TIER: nothing is above {WarnThreshold} ({WarnFraction:P0} of budget).");
            return;
        }

        var warningLines = new List<string>();
        foreach (WireString warning in warnings)
        {
            string line =
                $"WARNING: {warning.Surface} '{warning.Label}' is {warning.Measured} of {ClientTruncationBudget} "
                + $"({warning.PercentOfBudget(ClientTruncationBudget):F0}% of budget, "
                + $"{ClientTruncationBudget - warning.Measured} left). It still arrives intact, but the client "
                + "truncates silently and mid-sentence - move detail into the runtime payload before it crosses.";
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
                $"Under the {ClientTruncationBudget}-byte client truncation budget, but past the "
                + $"{WarnFraction:P0} warn line:");
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
        message.Append("the ").Append(ClientTruncationBudget.ToString(CultureInfo.InvariantCulture));
        message.AppendLine(" client truncation budget. Claude Code cuts these silently, mid-sentence, and");
        message.AppendLine("everything past the cut never reaches the model:");

        foreach (WireString entry in overBudget)
        {
            message.AppendLine(
                $"  {entry.Surface} '{entry.Label}': {entry.Measured} "
                + $"({entry.Chars} chars, {entry.Utf8Bytes} UTF-8 bytes) - "
                + $"{entry.Measured - ClientTruncationBudget} OVER the budget of {ClientTruncationBudget} "
                + $"({entry.PercentOfBudget(ClientTruncationBudget):F0}% of budget). "
                + $"The cut lands at: \"...{Excerpt(entry.Text, ClientTruncationBudget)}\"");
        }

        message.AppendLine("Shorten the description, or move the detail into the tool's runtime payload");
        message.Append("(advice/scope/sweep blocks) or into per-parameter descriptions, which are budgeted separately.");
        return message.ToString();
    }

    /// <summary>Shows where the cut lands: the 40 characters either side of the budget mark.</summary>
    private static string Excerpt(string text, int budget)
    {
        if (text.Length <= budget)
        {
            // Over budget on BYTES but not on characters: point at the tail instead.
            return text[Math.Max(0, text.Length - 80)..].ReplaceLineEndings(" ");
        }

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
        /// <summary>UTF-16 code units - what a char-counting client would see.</summary>
        public int Chars => Text.Length;

        /// <summary>UTF-8 bytes - what a byte-counting client (or "2KB") would see.</summary>
        public int Utf8Bytes => Encoding.UTF8.GetByteCount(Text);

        /// <summary>
        /// The size the budget is judged on. It is NOT documented whether the client's cap
        /// counts characters or bytes, and the two diverge the moment a description contains
        /// a non-ASCII character, so the larger of the two is the only safe reading. Taken
        /// as an explicit max rather than assuming bytes always win, so the rule still holds
        /// if the counting basis ever changes.
        /// </summary>
        public int Measured => Math.Max(Chars, Utf8Bytes);

        public double PercentOfBudget(int budget) => Measured * 100.0 / budget;
    }
}
