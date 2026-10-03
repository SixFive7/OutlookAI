using System.Globalization;
using System.Text;
using Microsoft.Win32.SafeHandles;
using OutlookAI.Core.Audit;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 for reading the audit log back (Q93): the parser against the grammar
/// <see cref="AuditLog.FormatLine"/> writes - escapes, malformed lines - and the scan against real
/// files in a temp directory: a partially written last line, a file that grows while it is read,
/// and the sharing that keeps a read from ever blocking the product's appends or the maintainer's
/// rename. Every file here is under the temp directory; the real log is never named.
/// </summary>
public sealed class AuditLogReaderTests : IDisposable
{
    private static readonly DateTime Ts = new(2026, 7, 23, 10, 11, 12, 345, DateTimeKind.Utc);

    private readonly string _dir = Path.Combine(Path.GetTempPath(), "OutlookAI-AuditLogReaderTests-" + Guid.NewGuid().ToString("N"));

    public AuditLogReaderTests()
    {
        Directory.CreateDirectory(_dir);
    }

    private string LogPath => Path.Combine(_dir, AuditLog.LogFileName);

    public void Dispose()
    {
        if (Directory.Exists(_dir))
        {
            Directory.Delete(_dir, recursive: true);
        }
    }

    // ================================================================== the parser

    public static TheoryData<string, string?> RoundTripValues => new()
    {
        { "plain", "00AB" },
        { "empty", string.Empty },
        { "spaces", "a value with spaces" },
        { "quote", "say \"hi\"" },
        { "backslash", @"C:\dir\file.txt" },
        { "escape-looking text", "literally \\n and \\t and \\\" in text" },
        { "trailing backslash", @"ends with \" },
        { "CR LF TAB", "line1\r\nline2\tend" },
        { "lone CR", "a\rb" },
        { "other controls", "nul\0 bell\u0007 esc\u001b" },
        // Built from code points so this file stays ASCII: e-diaeresis, an en dash, two CJK
        // ideographs and an emoji outside the BMP (a surrogate pair).
        { "unicode", "Zo" + (char)0x00EB + " " + (char)0x2013 + " " + (char)0x65E5 + (char)0x672C + " " + char.ConvertFromUtf32(0x1F642) },
        { "equals and op lookalike", "x=\"y\" op=send ts=2026" },
        { "null", null },
    };

    [Theory]
    [MemberData(nameof(RoundTripValues))]
    public void EveryValueTheWriterCanWrite_ReadsBackExactly(string label, string? value)
    {
        string line = AuditLog.FormatLine(Ts, "new_draft", new (string, string?)[]
        {
            ("entryId", "00AB"),
            ("value", value),
            ("store", "telefonie@xxlnet.nl"),
        });

        AuditLogEntry? entry = AuditLogReader.ParseLine(line);

        Assert.True(entry != null, label + ": the reader refused a line the writer wrote: " + line);
        Assert.Equal(Ts, entry!.TimestampUtc);
        Assert.Equal(DateTimeKind.Utc, entry.TimestampUtc.Kind);
        Assert.Equal("new_draft", entry.Operation);

        // Null values are omitted by the writer, and must stay omitted.
        List<KeyValuePair<string, string>> expected = new() { new("entryId", "00AB") };
        if (value != null)
        {
            expected.Add(new("value", value));
        }

        expected.Add(new("store", "telefonie@xxlnet.nl"));
        Assert.Equal(expected, entry.Fields);
    }

    [Fact]
    public void ARandomisedRoundTrip_OfThousandsOfLines_LosesNothing()
    {
        // A fixed seed, so a failure is reproducible; a palette loaded with every character the
        // grammar treats specially, so the escapes are exercised in every combination.
        Random random = new(86_93);
        const string Palette = "ab \"\\\r\n\t=_-.:/@\u00e9\u65e5\0";
        for (int n = 0; n < 2000; n++)
        {
            int fieldCount = random.Next(0, 6);
            var fields = new (string, string?)[fieldCount];
            for (int f = 0; f < fieldCount; f++)
            {
                StringBuilder value = new();
                int length = random.Next(0, 12);
                for (int c = 0; c < length; c++)
                {
                    value.Append(Palette[random.Next(Palette.Length)]);
                }

                fields[f] = ("k" + f.ToString(CultureInfo.InvariantCulture), value.ToString());
            }

            DateTime ts = Ts.AddMilliseconds(random.Next(0, 1_000_000));
            string line = AuditLog.FormatLine(ts, "op_" + n.ToString(CultureInfo.InvariantCulture), fields);

            AuditLogEntry? entry = AuditLogReader.ParseLine(line);

            Assert.True(entry != null, "refused: " + line);
            Assert.Equal(ts, entry!.TimestampUtc);
            Assert.Equal(fields.Select(f => new KeyValuePair<string, string>(f.Item1, f.Item2!)), entry.Fields);
        }
    }

    [Fact]
    public void TheEscapes_DecodeToWhatTheyStandFor()
    {
        // Written by hand, so the reader is pinned against the grammar and not only against the
        // writer's current output.
        string line = "ts=2026-07-23T10:11:12.345Z op=save_attachment path=\"C:\\\\dir\\\\file \\\"x\\\".txt\" "
            + "note=\"line1\\r\\nline2\\tend\" empty=\"\"";

        AuditLogEntry entry = Assert.IsType<AuditLogEntry>(AuditLogReader.ParseLine(line));

        Assert.Equal("save_attachment", entry.Operation);
        Assert.Equal(@"C:\dir\file ""x"".txt", entry.Fields[0].Value);
        Assert.Equal("line1\r\nline2\tend", entry.Fields[1].Value);
        Assert.Equal(string.Empty, entry.Fields[2].Value);
    }

    [Fact]
    public void ARepeatedKey_IsKept_NotDropped()
    {
        // The writer does not forbid it, so the reader must not lose it.
        AuditLogEntry entry = Assert.IsType<AuditLogEntry>(AuditLogReader.ParseLine(
            "ts=2026-07-23T10:11:12.345Z op=x entryId=\"AA\" entryId=\"BB\""));

        Assert.Equal(new[] { "AA", "BB" }, entry.Fields.Select(f => f.Value));
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("garbage")]
    // What OutlookAI wrote before the structured log (2026-07-23): no "ts=", no quotes.
    [InlineData("2026-07-22T10:11:12Z save_attachment path=C:\\x.txt bytes=5")]
    [InlineData("2026-07-22T10:11:12Z open_in_outlook entryId=00AB")]
    // The timestamp: missing, without milliseconds, not a date, not UTC-shaped.
    [InlineData("op=new_draft entryId=\"00AB\"")]
    [InlineData("ts=2026-07-23T10:11:12Z op=new_draft")]
    [InlineData("ts=2026-13-23T10:11:12.345Z op=new_draft")]
    [InlineData("ts=2026-07-23 10:11:12.345Z op=new_draft")]
    [InlineData("ts=2026-07-23T10:11:12.345+0 op=new_draft")]
    // The operation: missing, empty, not a token.
    [InlineData("ts=2026-07-23T10:11:12.345Z")]
    [InlineData("ts=2026-07-23T10:11:12.345Z op=")]
    [InlineData("ts=2026-07-23T10:11:12.345Z op=new.draft")]
    [InlineData("ts=2026-07-23T10:11:12.345Z  op=new_draft")]
    // The fields: unquoted, unterminated, a key that is not a token, no key at all.
    [InlineData("ts=2026-07-23T10:11:12.345Z op=x k=v")]
    [InlineData("ts=2026-07-23T10:11:12.345Z op=x k=\"v")]
    [InlineData("ts=2026-07-23T10:11:12.345Z op=x bad.key=\"v\"")]
    [InlineData("ts=2026-07-23T10:11:12.345Z op=x =\"v\"")]
    [InlineData("ts=2026-07-23T10:11:12.345Z op=x k=\"v\"x")]
    // Separators the writer never emits: two spaces, a trailing space, a tab.
    [InlineData("ts=2026-07-23T10:11:12.345Z op=x  k=\"v\"")]
    [InlineData("ts=2026-07-23T10:11:12.345Z op=x k=\"v\" ")]
    [InlineData("ts=2026-07-23T10:11:12.345Z op=x\tk=\"v\"")]
    // Escapes: unknown, and a backslash with nothing after it.
    [InlineData("ts=2026-07-23T10:11:12.345Z op=x k=\"a\\xb\"")]
    [InlineData("ts=2026-07-23T10:11:12.345Z op=x k=\"a\\")]
    // A raw tab or carriage return inside a value - the writer escapes both.
    [InlineData("ts=2026-07-23T10:11:12.345Z op=x k=\"a\tb\"")]
    [InlineData("ts=2026-07-23T10:11:12.345Z op=x k=\"a\rb\"")]
    public void ALineTheWriterCouldNotHaveWritten_IsMalformed(string line)
    {
        Assert.Null(AuditLogReader.ParseLine(line));
    }

    // ================================================================== reading a file

    [Fact]
    public void AMissingLog_IsAnEmptyLog_NotAnError()
    {
        AuditLogScan scan = AuditLogReader.Read(Path.Combine(_dir, "absent", AuditLog.LogFileName), MatchAll(), 25);

        Assert.False(scan.FileFound);
        Assert.Empty(scan.Entries);
        Assert.Equal(0, scan.Matched);
        Assert.Equal(0, scan.LinesScanned);
    }

    [Fact]
    public void TheNewestComeFirst_UpToTop_AndEveryMatchIsCounted()
    {
        for (int i = 0; i < 10; i++)
        {
            WriteLine(Ts.AddSeconds(i), "op" + i.ToString(CultureInfo.InvariantCulture), ("n", i.ToString(CultureInfo.InvariantCulture)));
        }

        AuditLogScan scan = AuditLogReader.Read(LogPath, MatchAll(), 3);

        Assert.True(scan.FileFound);
        Assert.Equal(10, scan.Matched);
        Assert.Equal(10, scan.LinesScanned);
        Assert.Equal(new[] { "op9", "op8", "op7" }, scan.Entries.Select(e => e.Operation));
        Assert.Equal(0, scan.MalformedLines);
        Assert.False(scan.IncompleteLastLine);
    }

    [Fact]
    public void TheFilters_AreAnded_AndEachMeansWhatItSays()
    {
        const string Draft = "00000000AA11BB22CC33DD44EE55FF66AA11BB22CC33DD44EE55FF66";
        const string Moved = "00000000FFEEDDCCBBAA99887766554433221100FFEEDDCCBBAA9988";
        WriteLine(Ts, "new_draft", ("entryId", Draft), ("store", "s"));
        WriteLine(Ts.AddSeconds(1), "send_token_issued", ("entryId", Draft));
        WriteLine(Ts.AddSeconds(2), "send_refused", ("entryId", Draft), ("reason", "draft_changed"));
        WriteLine(Ts.AddSeconds(3), "move_mail", ("entryId", Draft), ("newEntryId", Moved));
        WriteLine(Ts.AddSeconds(4), "reply_draft", ("entryId", "00FF"), ("sourceEntryId", Moved));
        WriteLine(Ts.AddSeconds(5), "manage_signature", ("name", Draft));

        // after is inclusive, before is exclusive.
        Assert.Equal(new[] { "send_refused", "send_token_issued" }, Ops(new AuditLogFilter(Ts.AddSeconds(1), Ts.AddSeconds(3), null, null)));

        // Exact operations, case-insensitively, any of several.
        Assert.Equal(new[] { "move_mail", "new_draft" }, Ops(new AuditLogFilter(null, null, new[] { "NEW_DRAFT", "move_mail" }, null)));

        // A trailing * is a family.
        Assert.Equal(new[] { "send_refused", "send_token_issued" }, Ops(new AuditLogFilter(null, null, new[] { "send*" }, null)));

        // An EntryID is found as entryId, newEntryId or sourceEntryId - any case - but NOT in a
        // field that merely holds the same text under another name.
        Assert.Equal(
            new[] { "reply_draft", "move_mail" },
            Ops(new AuditLogFilter(null, null, null, Moved.ToLowerInvariant())));
        Assert.Equal(
            new[] { "move_mail", "send_refused", "send_token_issued", "new_draft" },
            Ops(new AuditLogFilter(null, null, null, Draft)));

        // And all of them together.
        Assert.Equal(new[] { "send_refused" }, Ops(new AuditLogFilter(Ts.AddSeconds(2), null, new[] { "send*" }, Draft)));
    }

    [Fact]
    public void MalformedLines_AreCountedAndSkipped_AndTheRestStillRead()
    {
        StringBuilder text = new();
        text.Append(AuditLog.FormatLine(Ts, "first", Array.Empty<(string, string?)>())).Append("\r\n");
        text.Append("2026-07-22T10:11:12Z save_attachment path=C:\\x.txt bytes=5\r\n"); // the pre-structured format
        text.Append("\r\n"); // a blank line
        text.Append("garbage\n");
        text.Append(AuditLog.FormatLine(Ts.AddSeconds(1), "second", new (string, string?)[] { ("k", "v") })).Append('\n');
        File.WriteAllText(LogPath, text.ToString(), new UTF8Encoding(false));

        // Invalid UTF-8, and a line far longer than any real one.
        using (FileStream append = new(LogPath, FileMode.Append, FileAccess.Write))
        {
            append.Write(new byte[] { (byte)'t', (byte)'s', 0xC3, 0x28, (byte)'\n' });
            byte[] huge = Encoding.ASCII.GetBytes(new string('x', AuditLogReader.MaxLineBytes + 10) + "\n");
            append.Write(huge);
        }

        WriteLine(Ts.AddSeconds(2), "third");

        AuditLogScan scan = AuditLogReader.Read(LogPath, MatchAll(), 25);

        Assert.Equal(new[] { "third", "second", "first" }, scan.Entries.Select(e => e.Operation));
        Assert.Equal(5, scan.MalformedLines);
        Assert.Equal(8, scan.LinesScanned);
        Assert.False(scan.IncompleteLastLine);
        Assert.Equal(new[] { "first", "second", "third" }, scan.OperationsSeen);
    }

    [Fact]
    public void AByteOrderMark_AtTheStart_IsTolerated()
    {
        // The writer never emits one; a log opened and saved in an editor may carry one.
        File.WriteAllText(
            LogPath,
            AuditLog.FormatLine(Ts, "first", Array.Empty<(string, string?)>()) + "\n",
            new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));

        AuditLogScan scan = AuditLogReader.Read(LogPath, MatchAll(), 25);

        Assert.Equal("first", Assert.Single(scan.Entries).Operation);
        Assert.Equal(0, scan.MalformedLines);
    }

    [Fact]
    public void APartiallyWrittenLastLine_IsReported_NotParsed_AndArrivesOnceItIsComplete()
    {
        WriteLine(Ts, "complete");
        string next = AuditLog.FormatLine(Ts.AddSeconds(1), "in_flight", new (string, string?)[] { ("k", "v") });

        // Half of a line, the way a writer leaves it between two writes - or a crash leaves it.
        using (FileStream writer = OpenLikeTheWriter())
        {
            byte[] half = Encoding.UTF8.GetBytes(next.Substring(0, next.Length / 2));
            writer.Write(half, 0, half.Length);
        }

        AuditLogScan during = AuditLogReader.Read(LogPath, MatchAll(), 25);
        Assert.True(during.IncompleteLastLine);
        Assert.Equal("complete", Assert.Single(during.Entries).Operation);
        Assert.Equal(1, during.LinesScanned);
        Assert.Equal(0, during.MalformedLines);

        using (FileStream writer = OpenLikeTheWriter())
        {
            byte[] rest = Encoding.UTF8.GetBytes(next.Substring(next.Length / 2) + "\r\n");
            writer.Write(rest, 0, rest.Length);
        }

        AuditLogScan after = AuditLogReader.Read(LogPath, MatchAll(), 25);
        Assert.False(after.IncompleteLastLine);
        Assert.Equal(new[] { "in_flight", "complete" }, after.Entries.Select(e => e.Operation));
    }

    [Fact]
    public void AFileThatGrowsWhileItIsRead_IsReadUpToTheSnapshot_AndTheWriterIsNeverBlocked()
    {
        for (int i = 0; i < 5; i++)
        {
            AuditLog.AppendTo(_dir, "before" + i.ToString(CultureInfo.InvariantCulture), Array.Empty<(string, string?)>());
        }

        using (FileStream reader = AuditLogReader.OpenShared(LogPath)!)
        {
            long snapshot = reader.Length;

            // The product's own writer, while the reader holds the file open: it must not be
            // blocked, because a draft or send whose audit line cannot be written fails.
            for (int i = 0; i < 5; i++)
            {
                AuditLog.AppendTo(_dir, "during" + i.ToString(CultureInfo.InvariantCulture), Array.Empty<(string, string?)>());
            }

            AuditLogScan scan = AuditLogReader.Scan(reader, snapshot, MatchAll(), 25);

            Assert.Equal(5, scan.LinesScanned);
            Assert.All(scan.Entries, e => Assert.StartsWith("before", e.Operation, StringComparison.Ordinal));
            Assert.False(scan.IncompleteLastLine);
        }

        AuditLogScan later = AuditLogReader.Read(LogPath, MatchAll(), 25);
        Assert.Equal(10, later.LinesScanned);
        Assert.Equal("during4", later.Entries[0].Operation);
    }

    [Fact]
    public void ReadsRacingAConcurrentWriter_NeverBlockItAndNeverSeeATornLine()
    {
        const int Lines = 400;
        Exception? writerFailure = null;
        Thread writer = new(() =>
        {
            try
            {
                for (int i = 0; i < Lines; i++)
                {
                    AuditLog.AppendTo(_dir, "race", new (string, string?)[]
                    {
                        ("n", i.ToString(CultureInfo.InvariantCulture)),
                        ("padding", new string('p', i % 300)),
                    });
                }
            }
            catch (Exception ex)
            {
                writerFailure = ex;
            }
        });

        writer.Start();
        long previous = 0;
        int reads = 0;
        while (writer.IsAlive || reads == 0)
        {
            AuditLogScan scan = AuditLogReader.Read(LogPath, MatchAll(), 5);
            Assert.Equal(0, scan.MalformedLines);
            Assert.True(scan.LinesScanned >= previous, "a later read saw fewer lines than an earlier one");
            previous = scan.LinesScanned;
            reads++;
        }

        writer.Join();
        Assert.Null(writerFailure);

        AuditLogScan final = AuditLogReader.Read(LogPath, MatchAll(), 1);
        Assert.Equal(Lines, final.LinesScanned);
        Assert.Equal(Lines, final.Matched);
        Assert.Equal((Lines - 1).ToString(CultureInfo.InvariantCulture), final.Entries[0].Fields[0].Value);
        Assert.False(final.IncompleteLastLine);
        Assert.True(reads > 1, "the writer finished before a second read - the race was never run");
    }

    // ================================================================== sharing and the rename

    [Fact]
    public void AnOpenReader_BlocksNeitherAnAppendNorARename()
    {
        WriteLine(Ts, "first");
        string archive = Path.Combine(_dir, "audit.until-test.log");

        using (FileStream reader = AuditLogReader.OpenShared(LogPath)!)
        {
            // FileShare.ReadWrite: the product's append goes through while the read is open.
            AuditLog.AppendTo(_dir, "second", Array.Empty<(string, string?)>());

            // FileShare.Delete: so does the maintainer's rename.
            File.Move(LogPath, archive);
            Assert.True(reader.Length > 0);
        }

        Assert.True(File.Exists(archive));
        Assert.False(File.Exists(LogPath));
    }

    [Fact]
    public void AReaderOpenedWithoutWriteSharing_WouldFailTheProductsAppend_TheControl()
    {
        // Why the reader's share flags matter: File.ReadAllLines and friends open with
        // FileShare.Read, and while such a handle is open the product's append exhausts its
        // retries and reports the write operation as failed.
        WriteLine(Ts, "first");

        using (new FileStream(LogPath, FileMode.Open, FileAccess.Read, FileShare.Read))
        {
            InvalidOperationException failed = Assert.Throws<InvalidOperationException>(
                () => AuditLog.AppendTo(_dir, "second", Array.Empty<(string, string?)>()));
            Assert.Contains("attempts", failed.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void AnAppendInProgress_BlocksARename_AndTheNextAppendAfterARenameStartsAFreshLog()
    {
        // The facts the maintainer's rename procedure rests on, measured on real files.
        WriteLine(Ts, "old");
        string archive = Path.Combine(_dir, "audit.until-test.log");

        // 1. The handle an append really holds (AuditLogFile, since Q117) - which shares read and
        //    write, NOT delete - makes the rename fail instead of interleaving with the write.
        Assert.True(AuditLogFile.TryOpenForAppend(LogPath, AuditLog.WriteThrough, out SafeFileHandle? held, out int error), "open failed: " + error);
        using (held!)
        {
            Assert.ThrowsAny<IOException>(() => File.Move(LogPath, archive));
        }

        // 2. Between appends nothing holds the file, so the rename goes through at once...
        long archivedLength = new FileInfo(LogPath).Length;
        File.Move(LogPath, archive);

        // 3. ...and the next append creates a fresh log, leaving the archive exactly as it was.
        AuditLog.AppendTo(_dir, "new", Array.Empty<(string, string?)>());

        Assert.Equal(archivedLength, new FileInfo(archive).Length);
        Assert.Equal("new", Assert.Single(AuditLogReader.Read(LogPath, MatchAll(), 25).Entries).Operation);
        Assert.Equal("old", Assert.Single(AuditLogReader.Read(archive, MatchAll(), 25).Entries).Operation);
    }

    // ================================================================== the read half of the tripwire

    [Fact]
    public void WhileRedirected_ALogOutsideTemp_IsRefusedBeforeItIsRead()
    {
        // A test process may no more READ the real log than write it (Q86): code that forgot the
        // redirect must fail here rather than scan the maintainer's own log. Aimed at a log under
        // the build output, never the real one, so a broken tripwire reads only this test's line.
        Assert.True(AuditLog.IsRedirected, "not redirected - this test proves nothing");
        string outside = Path.Combine(AppContext.BaseDirectory, "audit-read-tripwire-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(outside);
        try
        {
            string log = Path.Combine(outside, AuditLog.LogFileName);
            File.WriteAllText(log, AuditLog.FormatLine(Ts, "planted", Array.Empty<(string, string?)>()) + "\n");

            InvalidOperationException refused = Assert.Throws<InvalidOperationException>(
                () => AuditLogReader.Read(log, MatchAll(), 25));
            Assert.Contains("was not read", refused.Message, StringComparison.Ordinal);
            Assert.Contains("test process", refused.Message, StringComparison.Ordinal);
        }
        finally
        {
            Directory.Delete(outside, recursive: true);
        }
    }

    // ================================================================== helpers

    private static AuditLogFilter MatchAll() => new(null, null, null, null);

    private string[] Ops(AuditLogFilter filter) =>
        AuditLogReader.Read(LogPath, filter, 25).Entries.Select(e => e.Operation).ToArray();

    private void WriteLine(DateTime ts, string operation, params (string Key, string? Value)[] fields)
    {
        using FileStream stream = OpenLikeTheWriter();
        byte[] bytes = new UTF8Encoding(false).GetBytes(AuditLog.FormatLine(ts, operation, fields) + Environment.NewLine);
        stream.Write(bytes, 0, bytes.Length);
    }

    /// <summary>
    /// A writer-like handle for planting raw bytes: read and write sharing, no delete, like the product's.
    /// The product's own handle is <see cref="AuditLogFile.TryOpenForAppend"/> (append-only, since Q117).
    /// </summary>
    private FileStream OpenLikeTheWriter() => new(LogPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
}
