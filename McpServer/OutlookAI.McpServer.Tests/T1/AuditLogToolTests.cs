using System.Reflection;
using System.Text.Json;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using OutlookAI.Core.Audit;
using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 for the audit_log tool (Q93), driven in-process - through <c>MailService.ReadAuditLog</c> and
/// through the tool method itself, which builds no COM gateway - against the log THIS process
/// appends to. That is the test run's throwaway log (<see cref="AuditIsolation"/>), so every call
/// here reads only lines tests wrote, and never the machine's real log. The stdio tier only checks
/// the advertised shape, because a server it starts would read the real one.
/// </summary>
public sealed class AuditLogToolTests
{
    // ================================================================== how it is marked

    [Fact]
    public void TheTool_CarriesTheReadOnlyAnnotations_OnItsAttribute()
    {
        // The exact metadata a classifier of write-capable tools can read without calling
        // anything: these four become annotations.readOnlyHint / destructiveHint / idempotentHint /
        // openWorldHint on the wire (pinned there by T3/AuditLogCiToolShapeTests).
        McpServerToolAttribute attribute = ToolMethod().GetCustomAttribute<McpServerToolAttribute>()!;

        Assert.Equal("audit_log", attribute.Name);
        Assert.True(attribute.ReadOnly);
        Assert.False(attribute.Destructive);
        Assert.True(attribute.Idempotent);
        Assert.False(attribute.OpenWorld);
    }

    [Fact]
    public void TheTool_DoesNotTeachTheMutationOutcome_AndSaysItIsReadOnly()
    {
        string description = ToolMethod().GetCustomAttribute<System.ComponentModel.DescriptionAttribute>()!.Description;

        Assert.DoesNotContain("outcome:", description, StringComparison.Ordinal);
        Assert.Contains("Read-only", description, StringComparison.Ordinal);
    }

    [Fact]
    public void TheCaps_ArePinned_AndTheToolDefaultsToTheDefault()
    {
        Assert.Equal(100, MailService.AuditLogTopCap);
        Assert.Equal(25, MailService.AuditLogTopDefault);

        ParameterInfo top = ToolMethod().GetParameters().Single(p => p.Name == "top");
        Assert.Equal(MailService.AuditLogTopDefault, top.DefaultValue);
    }

    // ================================================================== arguments

    public static TheoryData<string?, string?, string?, string?> RefusedArguments => new()
    {
        // after must come before before.
        { "2026-10-03T10:00:00Z", "2026-10-03T10:00:00Z", null, null },
        { "2026-10-03T11:00:00Z", "2026-10-03T10:00:00Z", null, null },
        // operation: not a name, a * anywhere but the end.
        { null, null, "bad op", null },
        { null, null, "se*nd", null },
        { null, null, "send,move mail", null },
        // entry_id: a hit id, not hex, too short, odd length.
        { null, null, null, "h12" },
        { null, null, null, "not-an-entry-id" },
        { null, null, null, "00AB" },
        { null, null, null, new string('A', 97) },
    };

    [Theory]
    [MemberData(nameof(RefusedArguments))]
    public void BadArguments_AreRefusedBeforeTheLogIsOpened(string? after, string? before, string? operation, string? entryId)
    {
        // The log is locked against every reader for the duration: if the call reached the file it
        // would fail opening it (the control below), so an ArgumentException proves it never tried.
        using (LockTheThrowawayLog())
        {
            Assert.Throws<ArgumentException>(() => MailService.ReadAuditLog(Utc(after), Utc(before), operation, entryId));
        }
    }

    [Fact]
    public void ValidArguments_DoOpenTheLog_TheControlForTheTestAbove()
    {
        using (LockTheThrowawayLog())
        {
            InvalidOperationException failed = Assert.Throws<InvalidOperationException>(() => MailService.ReadAuditLog());
            Assert.Contains("could not be opened", failed.Message, StringComparison.Ordinal);
        }
    }

    [Theory]
    [InlineData("send", new[] { "send" })]
    [InlineData(" send , move_mail ", new[] { "send", "move_mail" })]
    [InlineData("send*", new[] { "send*" })]
    [InlineData("*", new[] { "*" })]
    [InlineData("send,,", new[] { "send" })]
    public void Operations_AreSplitOnCommas_AndTrimmed(string operation, string[] expected)
    {
        Assert.Equal(expected, MailService.ParseAuditOperations(operation));
    }

    [Fact]
    public void NoFilters_MeanNoFilters()
    {
        Assert.Null(MailService.ParseAuditOperations(null));
        Assert.Null(MailService.ParseAuditOperations(" , "));
        Assert.Null(MailService.ParseAuditEntryId("   "));
        Assert.Equal(new string('a', 48), MailService.ParseAuditEntryId(" " + new string('a', 48) + " "));
    }

    // ================================================================== end to end, in-process

    [Fact]
    public async Task TheTool_ReturnsWhatThisProcessAppended_NewestFirst_FollowingTheItemThroughAMove()
    {
        RequireIsolation();
        string draft = RandomEntryId();
        string moved = RandomEntryId();
        AuditLog.Append("new_draft", ("entryId", draft), ("store", "hub@example.invalid"));
        AuditLog.Append("send_token_issued", ("entryId", draft));
        AuditLog.Append("move_mail", ("entryId", draft), ("newEntryId", moved), ("toFolder", "Archive"));
        AuditLog.Append("discard_draft", ("entryId", RandomEntryId()));

        using JsonDocument payload = await CallToolAsync(entry_id: draft);
        JsonElement root = payload.RootElement;

        Assert.Equal(AuditLog.EffectiveLogPath, root.GetProperty("path").GetString());
        Assert.Equal(new[] { "move_mail", "send_token_issued", "new_draft" }, Operations(root));
        Assert.Equal(3, root.GetProperty("matched").GetInt32());
        Assert.Equal(3, root.GetProperty("returned").GetInt32());
        Assert.False(root.GetProperty("truncated").GetBoolean());

        JsonElement newest = root.GetProperty("entries")[0];
        Assert.Equal(moved, newest.GetProperty("fields").GetProperty("newEntryId").GetString());
        Assert.Equal("Archive", newest.GetProperty("fields").GetProperty("toFolder").GetString());
        Assert.True(newest.GetProperty("utc").GetDateTime() <= DateTime.UtcNow);

        // The move's NEW EntryID finds the move too - which is how an item is followed onward.
        using JsonDocument onward = await CallToolAsync(entry_id: moved);
        Assert.Equal(new[] { "move_mail" }, Operations(onward.RootElement));

        // An operation filter, ANDed with the item.
        using JsonDocument sends = await CallToolAsync(entry_id: draft, operation: "send*");
        Assert.Equal(new[] { "send_token_issued" }, Operations(sends.RootElement));

        // A compact shape: the counters that are zero or false are absent, not zero.
        Assert.False(root.TryGetProperty("malformedLines", out _));
        Assert.False(root.TryGetProperty("incompleteLastLine", out _));
        Assert.False(root.TryGetProperty("logMissing", out _));
    }

    [Fact]
    public async Task MoreMatchesThanTop_AreCut_AndTheAnswerSaysSo()
    {
        RequireIsolation();
        string id = RandomEntryId();
        for (int i = 0; i < 3; i++)
        {
            AuditLog.Append("update_draft", ("entryId", id), ("n", i.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }

        using JsonDocument payload = await CallToolAsync(entry_id: id, top: 2);
        JsonElement root = payload.RootElement;

        Assert.True(root.GetProperty("truncated").GetBoolean());
        Assert.Equal(2, root.GetProperty("returned").GetInt32());
        Assert.Equal(3, root.GetProperty("matched").GetInt32());
        Assert.Equal("2", root.GetProperty("entries")[0].GetProperty("fields").GetProperty("n").GetString());
        Assert.Contains(Advice(root), a => a.Contains("newest 2 of 3", StringComparison.Ordinal));
    }

    [Fact]
    public async Task ATopAboveTheCap_IsReduced_AndTheAnswerSaysSo()
    {
        RequireIsolation();
        string id = RandomEntryId();
        AuditLog.Append("save_attachment", ("entryId", id));

        using JsonDocument payload = await CallToolAsync(entry_id: id, top: 1000);

        Assert.Contains(Advice(payload.RootElement), a => a.Contains("reduced to 100", StringComparison.Ordinal));
    }

    [Fact]
    public async Task AnUnparsableDate_IsAnInvalidArgument_OnTheToolsOwnErrorShape()
    {
        CallToolResult result = await OutlookTools.AuditLog(after: "not a date");

        Assert.True(result.IsError);
        using JsonDocument payload = JsonDocument.Parse(Text(result));
        Assert.Equal("InvalidArgument", payload.RootElement.GetProperty("error").GetProperty("type").GetString());
    }

    [Fact]
    public async Task ALogArchivedBesideTheLiveOne_IsNeverRead()
    {
        // The maintainer's clean-up renames the old log to audit.until-<date>.log next to the new
        // one. Whatever it holds - mostly test noise - must not reach an agent.
        RequireIsolation();
        string id = RandomEntryId();
        Directory.CreateDirectory(AuditLog.EffectiveDirectory);
        string archive = Path.Combine(AuditLog.EffectiveDirectory, "audit.until-q93-test.log");
        File.WriteAllText(archive, AuditLog.FormatLine(DateTime.UtcNow, "new_draft", new (string, string?)[] { ("entryId", id) }) + "\r\n");
        try
        {
            using JsonDocument payload = await CallToolAsync(entry_id: id);
            Assert.Equal(0, payload.RootElement.GetProperty("matched").GetInt32());
        }
        finally
        {
            File.Delete(archive);
        }
    }

    // ================================================================== what the answer says

    [Fact]
    public void AMissingLog_IsSaidToBeMissing_NotEmptyOfMatches()
    {
        AuditLogOutcome outcome = MailService.DescribeAuditLogScan("x", Scan(fileFound: false), 25, operationFiltered: false);

        Assert.True(outcome.LogMissing);
        Assert.Empty(outcome.Entries);
        Assert.Contains(outcome.Advice!, a => a.Contains("no audit log", StringComparison.Ordinal));
    }

    [Fact]
    public void AnOperationThatMatchedNothing_IsAnsweredWithTheOperationsTheLogHolds()
    {
        AuditLogOutcome outcome = MailService.DescribeAuditLogScan(
            "x", Scan(lines: 4, operationsSeen: new[] { "move_mail", "new_draft" }), 25, operationFiltered: true);

        Assert.Equal(0, outcome.Matched);
        Assert.Contains(outcome.Advice!, a => a.Contains("Operations this log does contain: move_mail, new_draft.", StringComparison.Ordinal));
    }

    [Fact]
    public void MalformedAndIncompleteLines_AreReportedInFieldsAndInWords()
    {
        AuditLogOutcome outcome = MailService.DescribeAuditLogScan(
            "x", Scan(lines: 10, malformed: 3, incomplete: true), 25, operationFiltered: false);

        Assert.Equal(3, outcome.MalformedLines);
        Assert.True(outcome.IncompleteLastLine);
        Assert.Null(outcome.LogMissing);
        Assert.Contains(outcome.Advice!, a => a.StartsWith("3 line(s)", StringComparison.Ordinal));
        Assert.Contains(outcome.Advice!, a => a.Contains("still being written", StringComparison.Ordinal));
    }

    // ================================================================== helpers

    private static MethodInfo ToolMethod() =>
        typeof(OutlookTools).GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Single(m => m.GetCustomAttribute<McpServerToolAttribute>()?.Name == "audit_log");

    private static async Task<JsonDocument> CallToolAsync(string? entry_id = null, string? operation = null, int top = 25)
    {
        CallToolResult result = await OutlookTools.AuditLog(operation: operation, entry_id: entry_id, top: top);
        Assert.True(result.IsError != true, "audit_log failed: " + Text(result));
        return JsonDocument.Parse(Text(result));
    }

    private static string Text(CallToolResult result) =>
        Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text;

    private static string[] Operations(JsonElement root) =>
        root.GetProperty("entries").EnumerateArray().Select(e => e.GetProperty("operation").GetString()!).ToArray();

    private static string[] Advice(JsonElement root) =>
        root.TryGetProperty("advice", out JsonElement advice)
            ? advice.EnumerateArray().Select(a => a.GetString()!).ToArray()
            : Array.Empty<string>();

    private static DateTime? Utc(string? value) =>
        value == null ? null : DateTimeOffset.Parse(value, System.Globalization.CultureInfo.InvariantCulture).UtcDateTime;

    private static string RandomEntryId() =>
        (Guid.NewGuid().ToString("N") + Guid.NewGuid().ToString("N")).ToUpperInvariant();

    private static AuditLogScan Scan(
        bool fileFound = true, long lines = 0, long malformed = 0, bool incomplete = false, IReadOnlyList<string>? operationsSeen = null) =>
        new(fileFound, Array.Empty<AuditLogEntry>(), 0, lines, malformed, incomplete, operationsSeen ?? Array.Empty<string>());

    /// <summary>
    /// Every test here that writes, locks or reads the log asserts this FIRST: on code that does not
    /// isolate, the log is the real one, and the test must stop before touching it.
    /// </summary>
    private static void RequireIsolation()
    {
        Assert.True(AuditLog.IsRedirected, "not redirected - refusing to touch the audit log");
        Assert.Equal(AuditIsolation.ThrowawayDirectory, AuditLog.EffectiveDirectory, ignoreCase: true);
    }

    /// <summary>Holds the throwaway log open with no sharing at all, so any attempt to read it fails.</summary>
    private static FileStream LockTheThrowawayLog()
    {
        RequireIsolation();
        AuditLog.Append("q93_lock_probe");
        return new FileStream(AuditLog.EffectiveLogPath, FileMode.Open, FileAccess.Read, FileShare.None);
    }
}
