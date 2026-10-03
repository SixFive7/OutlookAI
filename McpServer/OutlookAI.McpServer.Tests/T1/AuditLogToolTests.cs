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

    // ================================================================== paging (Q119)

    [Fact]
    public async Task TheToolPages_ThroughEveryEntryOnce_UntilNextTokenIsAbsent()
    {
        RequireIsolation();
        string id = RandomEntryId();
        for (int i = 0; i < 7; i++)
        {
            // Appended in a tight loop, so several share a millisecond - which paging must not care about.
            AuditLog.Append("update_draft", ("entryId", id), ("n", i.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        }

        List<string> seen = new();
        List<bool> truncated = new();
        string? token = null;
        string[] lastAdvice = Array.Empty<string>();
        do
        {
            using JsonDocument page = await CallToolAsync(entry_id: id, top: 3, resume_token: token);
            JsonElement root = page.RootElement;
            seen.AddRange(root.GetProperty("entries").EnumerateArray()
                .Select(e => e.GetProperty("fields").GetProperty("n").GetString()!));
            truncated.Add(root.GetProperty("truncated").GetBoolean());
            Assert.Equal(7, root.GetProperty("matched").GetInt32());
            token = root.TryGetProperty("nextToken", out JsonElement next) ? next.GetString() : null;
            Assert.Equal(truncated[truncated.Count - 1], token != null);
            Assert.True(root.GetProperty("entries")[0].GetProperty("pid").GetInt32() > 0);
            lastAdvice = Advice(root);
        }
        while (token != null);

        Assert.Equal(new[] { "6", "5", "4", "3", "2", "1", "0" }, seen);
        Assert.Equal(new[] { true, true, false }, truncated);
        Assert.Contains(lastAdvice, a => a.Contains("last page", StringComparison.Ordinal));
    }

    [Fact]
    public async Task AResumeTokenUsedWithOtherFilters_IsRefused_NamingWhatChanged()
    {
        RequireIsolation();
        string id = RandomEntryId();
        for (int i = 0; i < 3; i++)
        {
            AuditLog.Append("update_draft", ("entryId", id));
        }

        using JsonDocument first = await CallToolAsync(entry_id: id, operation: "update_draft", top: 1);
        string token = first.RootElement.GetProperty("nextToken").GetString()!;

        // The same filters in another spelling are the same question.
        using JsonDocument same = await CallToolAsync(entry_id: id.ToLowerInvariant(), operation: "UPDATE_DRAFT", top: 2, resume_token: token);
        Assert.Equal(2, same.RootElement.GetProperty("returned").GetInt32());

        CallToolResult changed = await OutlookTools.AuditLog(entry_id: id, operation: "update_draft,send", top: 1, resume_token: token);
        Assert.True(changed.IsError);
        using JsonDocument error = JsonDocument.Parse(Text(changed));
        string message = error.RootElement.GetProperty("error").GetProperty("message").GetString()!;
        Assert.Equal("InvalidArgument", error.RootElement.GetProperty("error").GetProperty("type").GetString());
        Assert.Contains("operation changed", message, StringComparison.Ordinal);
        Assert.DoesNotContain("entry_id", message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("scan-0123456789abcdef0123456789abcdef", "another tool")]
    [InlineData("not a token", "not a token audit_log issued")]
    [InlineData("pg1.AAAA.00000000", "not a token audit_log issued")]
    public void AResumeTokenAuditLogDidNotIssue_IsRefusedBeforeTheLogIsOpened(string token, string expected)
    {
        using (LockTheThrowawayLog())
        {
            ArgumentException refused = Assert.Throws<ArgumentException>(() => MailService.ReadAuditLog(resumeToken: token));
            Assert.Contains(expected, refused.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void AnotherToolsPagingToken_IsRefusedAsAnotherTools()
    {
        string foreign = Paging.IssueToken("list_folders", new PagingFingerprint().ToString(), new[] { "5" });
        ArgumentException refused = Assert.Throws<ArgumentException>(() => MailService.ReadAuditLog(resumeToken: foreign));
        Assert.Contains("another tool", refused.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(AuditLogResumeStatus.FileMissing, "archived")]
    [InlineData(AuditLogResumeStatus.FileReplaced, "archived or replaced")]
    [InlineData(AuditLogResumeStatus.PositionNotFound, "truncated or edited")]
    public void ATokenTheLogNoLongerFits_IsRefusedSayingWhyAndWhatToDo(AuditLogResumeStatus status, string why)
    {
        string message = MailService.DescribeAuditResumeRefusal(status);
        Assert.Contains(why, message, StringComparison.Ordinal);
        Assert.Contains("Leave resume_token out", message, StringComparison.Ordinal);
    }

    // ================================================================== integrity in the answer (Q117)

    [Fact]
    public void DamagedMissingAndUnverifiedLines_AreReportedInFieldsAndInWords()
    {
        AuditLogScan scan = new()
        {
            FileFound = true,
            LinesScanned = 40,
            MalformedLines = 3,
            ChecksumFailures = 2,
            MissingLines = 5,
            UnverifiedLines = 4,
            LinesWithoutWriterLock = 1,
            Gaps = new[] { new AuditLogGap(4242, "0a1b2c3d", 3, 6), new AuditLogGap(77, "ffffffff", 1, 1) },
        };

        AuditLogOutcome outcome = MailService.DescribeAuditLogScan("x", scan, 25, operationFiltered: false);

        Assert.Equal(3, outcome.MalformedLines);
        Assert.Equal(2, outcome.DamagedLines);
        Assert.Equal(5, outcome.MissingLines);
        Assert.Equal(4, outcome.UnverifiedLines);
        Assert.Equal(1, outcome.LinesWithoutWriterLock);
        Assert.Equal(new[] { "pid 4242 run 0a1b2c3d: seq 3-6 (4 lines)", "pid 77 run ffffffff: seq 1 (1 line)" }, outcome.SequenceGaps);
        Assert.Contains(outcome.Advice!, a => a.StartsWith("1 line(s) of the log are not in the format", StringComparison.Ordinal));
        Assert.Contains(outcome.Advice!, a => a.StartsWith("2 line(s) failed their checksum", StringComparison.Ordinal));
        Assert.Contains(outcome.Advice!, a => a.StartsWith("5 line(s) are MISSING", StringComparison.Ordinal));
        Assert.Contains(outcome.Advice!, a => a.StartsWith("4 line(s) were written before lines carried a checksum", StringComparison.Ordinal));
        Assert.Contains(outcome.Advice!, a => a.StartsWith("1 line(s) were written without the writers' lock", StringComparison.Ordinal));
    }

    // ================================================================== helpers

    private static MethodInfo ToolMethod() =>
        typeof(OutlookTools).GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Single(m => m.GetCustomAttribute<McpServerToolAttribute>()?.Name == "audit_log");

    private static async Task<JsonDocument> CallToolAsync(
        string? entry_id = null, string? operation = null, int top = 25, string? resume_token = null)
    {
        CallToolResult result = await OutlookTools.AuditLog(operation: operation, entry_id: entry_id, top: top, resume_token: resume_token);
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
