using System.Globalization;
using System.Text;
using OutlookAI.Core.Audit;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 for the audit log's integrity (Q117, 2026-10-03): the trailer every line now ends with
/// (<c>pid</c>, <c>run</c>, <c>seq</c>, an optional <c>lock</c>, <c>crc</c>), the reader that checks it,
/// the per-writer gap report, the repair of a line a crash left unfinished, and the writer lock's
/// three ways of not being held normally - each recorded on the line it concerns. Real files under
/// %TEMP% only; the many-process proof is <see cref="AuditLogStressTests"/>.
/// </summary>
public sealed class AuditLogIntegrityTests : IDisposable
{
    private static readonly DateTime Ts = new(2026, 10, 3, 10, 11, 12, 345, DateTimeKind.Utc);

    private readonly string _dir = Path.Combine(Path.GetTempPath(), "OutlookAI-AuditLogIntegrityTests-" + Guid.NewGuid().ToString("N"));

    public AuditLogIntegrityTests()
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

    // ================================================================== the checksum

    [Fact]
    public void TheChecksum_IsTheStandardCrc32()
    {
        // The published check value of CRC-32/ISO-HDLC, and a bit-by-bit reference for the table.
        Assert.Equal(0xCBF43926u, AuditCrc32.Compute(Encoding.ASCII.GetBytes("123456789"), 0, 9));
        byte[] sample = Encoding.UTF8.GetBytes("ts=2026-10-03T10:11:12.345Z op=x k=\"é日\"");
        Assert.Equal(BitwiseCrc32(sample), AuditCrc32.Compute(sample, 0, sample.Length));
        Assert.Equal("cbf43926", AuditCrc32.ToHex(0xCBF43926u));
        Assert.True(AuditCrc32.TryParseHex("cbf43926", out uint parsed));
        Assert.Equal(0xCBF43926u, parsed);
        Assert.False(AuditCrc32.TryParseHex("CBF43926", out _));
        Assert.False(AuditCrc32.TryParseHex("cbf4392", out _));
    }

    // ================================================================== the trailer

    [Fact]
    public void ASealedLine_HasItsGoldenShape()
    {
        string line = AuditLog.FormatSealedLine(
            Ts, "new_draft", new (string, string?)[] { ("entryId", "00AB"), ("store", "s") }, 4242, "0a1b2c3d", 17, null);

        string covered = "ts=2026-10-03T10:11:12.345Z op=new_draft entryId=\"00AB\" store=\"s\" pid=\"4242\" run=\"0a1b2c3d\" seq=\"17\"";
        Assert.Equal(covered + " crc=\"" + Crc(covered) + "\"", line);

        // The lock field, only when there is something to say, between seq and crc.
        string timedOut = AuditLog.FormatSealedLine(Ts, "send", Array.Empty<(string, string?)>(), 1, "ffffffff", 1, AuditLog.LockTimeout);
        Assert.Contains(" seq=\"1\" lock=\"timeout\" crc=\"", timedOut, StringComparison.Ordinal);
    }

    [Fact]
    public void EveryAppendedLine_CarriesPidRunAndSeq_AndAChecksumThatVerifies()
    {
        AuditLog.AppendTo(_dir, "first", new (string, string?)[] { ("k", "v1") });
        AuditLog.AppendTo(_dir, "second", Array.Empty<(string, string?)>());
        AuditLog.AppendTo(_dir, "third", new (string, string?)[] { ("k", "v3") });

        string[] raw = File.ReadAllLines(LogPath);
        Assert.Equal(3, raw.Length);
        for (int i = 0; i < raw.Length; i++)
        {
            Assert.Matches(
                "^ts=\\S+ op=\\w+( k=\"v\\d\")? pid=\"" + AuditLog.WriterPid.ToString(CultureInfo.InvariantCulture)
                + "\" run=\"[0-9a-f]{8}\" seq=\"" + (i + 1).ToString(CultureInfo.InvariantCulture) + "\" crc=\"[0-9a-f]{8}\"$",
                raw[i]);
        }

        AuditLogScan scan = AuditLogReader.Read(LogPath, All(), 25);
        Assert.Equal(3, scan.VerifiedLines);
        Assert.Equal(0, scan.MalformedLines);
        Assert.Equal(0, scan.MissingLines);
        Assert.Equal(1, scan.Writers);
        Assert.Equal(new long?[] { 3, 2, 1 }, scan.Entries.Select(e => e.Seq));
        Assert.All(scan.Entries, e =>
        {
            Assert.Equal(AuditLog.WriterPid, e.Pid);
            Assert.Equal(AuditLog.WriterRun, e.Run);
            Assert.Null(e.WriterLock);
            Assert.True(e.Checksummed);
        });

        // The trailer is split off: the caller's fields are exactly what the caller passed.
        Assert.Equal(new[] { new KeyValuePair<string, string>("k", "v1") }, scan.Entries[2].Fields);
        Assert.Empty(scan.Entries[1].Fields);
    }

    [Theory]
    [InlineData("pid")]
    [InlineData("run")]
    [InlineData("seq")]
    [InlineData("lock")]
    [InlineData("crc")]
    public void TheTrailersKeys_AreRefusedFromCallers_AndNothingIsWritten(string key)
    {
        (string, string?)[] fields = { (key, "1") };

        Assert.Throws<ArgumentException>(() => AuditLog.AppendTo(_dir, "op", fields));
        Assert.Throws<ArgumentException>(() => AuditLog.FormatLine(Ts, "op", fields));
        Assert.False(File.Exists(LogPath));

        // Keys are compared exactly: another spelling is a caller's own field.
        AuditLog.AppendTo(_dir, "op", new (string, string?)[] { (key.ToUpperInvariant(), "1") });
        Assert.Equal(1, AuditLogReader.Read(LogPath, All(), 25).VerifiedLines);
    }

    [Fact]
    public void AChangedByte_FailsTheChecksum_TheLineIsLeftOut_AndItsNumberIsMissing()
    {
        AuditLog.AppendTo(_dir, "one", new (string, string?)[] { ("k", "aaaa") });
        AuditLog.AppendTo(_dir, "two", new (string, string?)[] { ("k", "bbbb") });
        AuditLog.AppendTo(_dir, "three", new (string, string?)[] { ("k", "cccc") });

        // One byte of the middle line's value, the way an overwrite or an edit changes a line: it
        // still parses, but it is not what was written.
        string text = File.ReadAllText(LogPath);
        File.WriteAllText(LogPath, text.Replace("k=\"bbbb\"", "k=\"bbbc\""), new UTF8Encoding(false));
        string damaged = File.ReadAllLines(LogPath)[1];
        Assert.Null(AuditLogReader.ParseLine(damaged));
        Assert.NotNull(AuditLogReader.ParseLine(damaged.Replace("k=\"bbbc\"", "k=\"bbbb\"")));

        AuditLogScan scan = AuditLogReader.Read(LogPath, All(), 25);
        Assert.Equal(new[] { "three", "one" }, scan.Entries.Select(e => e.Operation));
        Assert.Equal(1, scan.ChecksumFailures);
        Assert.Equal(1, scan.MalformedLines);
        Assert.Equal(2, scan.VerifiedLines);
        Assert.Equal(1, scan.MissingLines);
        AuditLogGap gap = Assert.Single(scan.Gaps);
        Assert.Equal((AuditLog.WriterPid, AuditLog.WriterRun, 2L, 2L), (gap.Pid, gap.Run, gap.FirstMissing, gap.LastMissing));
    }

    public static TheoryData<string, string> LinesTheWriterCannotProduce()
    {
        string covered = "ts=2026-10-03T10:11:12.345Z op=x k=\"v\" pid=\"4242\" run=\"0a1b2c3d\" seq=\"7\"";
        return new TheoryData<string, string>
        {
            // Each checksum below is CORRECT for its line: the structure is refused on its own.
            { "no pid", Seal("ts=2026-10-03T10:11:12.345Z op=x k=\"v\" run=\"0a1b2c3d\" seq=\"7\"") },
            { "no run", Seal("ts=2026-10-03T10:11:12.345Z op=x k=\"v\" pid=\"4242\" seq=\"7\"") },
            { "out of order", Seal("ts=2026-10-03T10:11:12.345Z op=x k=\"v\" run=\"0a1b2c3d\" pid=\"4242\" seq=\"7\"") },
            { "run in capitals", Seal(covered.Replace("0a1b2c3d", "0A1B2C3D")) },
            { "run too short", Seal(covered.Replace("0a1b2c3d", "0a1b2c3")) },
            { "seq with a leading zero", Seal(covered.Replace("seq=\"7\"", "seq=\"07\"")) },
            { "seq zero", Seal(covered.Replace("seq=\"7\"", "seq=\"0\"")) },
            { "pid not a number", Seal(covered.Replace("pid=\"4242\"", "pid=\"42a\"")) },
            { "an unknown lock", Seal(covered + " lock=\"maybe\"") },
            { "lock before seq", Seal(covered.Replace(" seq=\"7\"", " lock=\"timeout\" seq=\"7\"")) },
            { "a field after crc", Seal(covered) + " k2=\"v\"" },
            { "crc in capitals", covered + " crc=\"" + Crc(covered).ToUpperInvariant() + "\"" },
            { "a trailer key among the caller's fields", Seal("ts=2026-10-03T10:11:12.345Z op=x seq=\"1\" pid=\"4242\" run=\"0a1b2c3d\" seq=\"7\"") },
            { "a trailer key on a line without one", "ts=2026-10-03T10:11:12.345Z op=x pid=\"4242\"" },
            { "a wrong checksum", covered + " crc=\"00000000\"" },
        };
    }

    [Theory]
    [MemberData(nameof(LinesTheWriterCannotProduce))]
    public void ATrailerTheWriterCannotProduce_IsMalformed(string label, string line)
    {
        Assert.True(AuditLogReader.ParseLine(line) == null, label + ": " + line);
    }

    [Fact]
    public void ALineWithoutATrailer_IsReadUnverified()
    {
        File.WriteAllText(
            LogPath,
            AuditLog.FormatLine(Ts, "old_one", Array.Empty<(string, string?)>()) + "\n"
                + AuditLog.FormatLine(Ts.AddSeconds(1), "old_two", new (string, string?)[] { ("k", "v") }) + "\n",
            new UTF8Encoding(false));

        AuditLogScan scan = AuditLogReader.Read(LogPath, All(), 25);

        Assert.Equal(2, scan.UnverifiedLines);
        Assert.Equal(0, scan.VerifiedLines);
        Assert.Equal(0, scan.Writers);
        Assert.All(scan.Entries, e =>
        {
            Assert.False(e.Checksummed);
            Assert.Null(e.Pid);
            Assert.Null(e.Seq);
        });
    }

    [Fact]
    public void Gaps_AreReportedPerWriter_AsRunsOfMissingNumbers()
    {
        // Writer A wrote 1, 2, 4 and 7 (and 2 twice); writer B's lines start at 2.
        StringBuilder text = new();
        foreach ((int pid, string run, long seq) in new[]
        {
            (4242, "0a1b2c3d", 1L), (4242, "0a1b2c3d", 2L), (100, "ffff0000", 2L), (4242, "0a1b2c3d", 4L),
            (4242, "0a1b2c3d", 2L), (100, "ffff0000", 3L), (4242, "0a1b2c3d", 7L),
        })
        {
            text.Append(AuditLog.FormatSealedLine(Ts, "x", Array.Empty<(string, string?)>(), pid, run, seq, null)).Append('\n');
        }

        File.WriteAllText(LogPath, text.ToString(), new UTF8Encoding(false));
        AuditLogScan scan = AuditLogReader.Read(LogPath, All(), 25);

        Assert.Equal(2, scan.Writers);
        Assert.Equal(1, scan.DuplicateLines);
        Assert.Equal(4, scan.MissingLines);
        Assert.Equal(
            new[] { (100, "ffff0000", 1L, 1L), (4242, "0a1b2c3d", 3L, 3L), (4242, "0a1b2c3d", 5L, 6L) },
            scan.Gaps.Select(g => (g.Pid, g.Run, g.FirstMissing, g.LastMissing)));

        // And in the tool's words.
        Assert.Equal(
            new[] { "pid 100 run ffff0000: seq 1 (1 line)", "pid 4242 run 0a1b2c3d: seq 3 (1 line)", "pid 4242 run 0a1b2c3d: seq 5-6 (2 lines)" },
            MailService.DescribeAuditGaps(scan));
    }

    [Fact]
    public void ANewFileUnderTheSameName_StartsEveryWritersNumberingAgain()
    {
        AuditLog.AppendTo(_dir, "old1", Array.Empty<(string, string?)>());
        AuditLog.AppendTo(_dir, "old2", Array.Empty<(string, string?)>());
        string archive = Path.Combine(_dir, "audit.until-test.log");
        File.Move(LogPath, archive);
        AuditLog.AppendTo(_dir, "new1", Array.Empty<(string, string?)>());

        AuditLogScan fresh = AuditLogReader.Read(LogPath, All(), 25);
        AuditLogScan old = AuditLogReader.Read(archive, All(), 25);

        // A writer's lines in one file are 1..n, so neither file reports the other's lines as lost.
        Assert.Equal(1L, Assert.Single(fresh.Entries).Seq);
        Assert.Equal(new long?[] { 2, 1 }, old.Entries.Select(e => e.Seq));
        Assert.Equal(0, fresh.MissingLines);
        Assert.Equal(0, old.MissingLines);
    }

    // ================================================================== a line a crash left unfinished

    [Fact]
    public void AnUnfinishedLastLine_IsClosedBeforeTheNextLine_SoOnlyTheFragmentIsLost()
    {
        File.WriteAllBytes(LogPath, Encoding.UTF8.GetBytes("ts=2026-10-03T10:11:12.345Z op=cut k=\"half"));

        AuditLog.AppendTo(_dir, "next", Array.Empty<(string, string?)>());

        string text = File.ReadAllText(LogPath);
        Assert.StartsWith("ts=2026-10-03T10:11:12.345Z op=cut k=\"half\nts=", text, StringComparison.Ordinal);
        AuditLogScan scan = AuditLogReader.Read(LogPath, All(), 25);
        Assert.Equal("next", Assert.Single(scan.Entries).Operation);
        Assert.Equal(1, scan.MalformedLines);
        Assert.False(scan.IncompleteLastLine);
    }

    [Fact]
    public void AFinishedLastLine_GetsNoExtraLineBreak()
    {
        AuditLog.AppendTo(_dir, "one", Array.Empty<(string, string?)>());
        AuditLog.AppendTo(_dir, "two", Array.Empty<(string, string?)>());

        string text = File.ReadAllText(LogPath);
        Assert.DoesNotContain("\n\n", text, StringComparison.Ordinal);
        Assert.EndsWith("\n", text, StringComparison.Ordinal);
        Assert.Equal(2, text.Count(c => c == '\n'));
    }

    // ================================================================== the lock, and when it is not held

    [Fact]
    public void TheTimestamp_IsTakenWhileTheLockIsHeld()
    {
        // The lock orders the file only if the time on a line is read under it: a writer that waited
        // must stamp the moment it got the lock, not the moment it started waiting.
        using Mutex blocker = new(false, AuditLog.WriterMutexName(LogPath));
        Assert.True(blocker.WaitOne(0));
        Thread append = AppendOnAnotherThread("waited", out Func<Exception?> failure);
        Thread.Sleep(600);
        DateTime released = DateTime.UtcNow;
        blocker.ReleaseMutex();
        Assert.True(append.Join(TimeSpan.FromSeconds(30)));
        Assert.Null(failure());

        AuditLogEntry entry = Assert.Single(AuditLogReader.Read(LogPath, All(), 25).Entries);
        Assert.True(entry.TimestampUtc >= released.AddMilliseconds(-1), "stamped " + entry.TimestampUtc.ToString("o") + ", released " + released.ToString("o"));
        Assert.Null(entry.WriterLock);
    }

    [Fact]
    public void ALockTimeout_IsWrittenOnTheLine_AndCountedForOutlookHealth()
    {
        long before = AuditLog.WriterLockTimeouts;
        using Mutex blocker = new(false, AuditLog.WriterMutexName(LogPath));
        Assert.True(blocker.WaitOne(0));
        try
        {
            // Another thread: the mutex is re-entrant for the thread that holds it.
            Thread append = AppendOnAnotherThread("impatient", out Func<Exception?> failure);
            Assert.True(append.Join(TimeSpan.FromSeconds(30)));
            Assert.Null(failure());
        }
        finally
        {
            blocker.ReleaseMutex();
        }

        AuditLogScan scan = AuditLogReader.Read(LogPath, All(), 25);
        Assert.Equal(AuditLog.LockTimeout, Assert.Single(scan.Entries).WriterLock);
        Assert.Equal(1, scan.LinesWithoutWriterLock);
        Assert.Equal(before + 1, AuditLog.WriterLockTimeouts);
        Assert.True(MailService.DescribeAuditLog(out _).LockTimeouts >= 1);
    }

    [Fact]
    public void ALockItsOwnerDiedHolding_IsTakenOver_AndThatIsWrittenOnTheLine()
    {
        long before = AuditLog.WriterLockAbandoned;
        string name = AuditLog.WriterMutexName(LogPath);
        Mutex? abandoned = null;
        Thread owner = new(() =>
        {
            abandoned = new Mutex(false, name);
            abandoned.WaitOne();

            // Ends without releasing it: the way a writer killed mid-append leaves it.
        });
        owner.Start();
        owner.Join();
        try
        {
            AuditLog.AppendTo(_dir, "inherited", Array.Empty<(string, string?)>());
        }
        finally
        {
            abandoned?.Dispose();
        }

        AuditLogScan scan = AuditLogReader.Read(LogPath, All(), 25);
        Assert.Equal(AuditLog.LockAbandoned, Assert.Single(scan.Entries).WriterLock);
        Assert.Equal(1, scan.LinesAfterAbandonedLock);
        Assert.Equal(0, scan.LinesWithoutWriterLock);
        Assert.Equal(before + 1, AuditLog.WriterLockAbandoned);
    }

    // ================================================================== helpers

    private static AuditLogFilter All() => new(null, null, null, null);

    /// <summary>An append on a thread of its own, started; <paramref name="failure"/> reads what it threw.</summary>
    private Thread AppendOnAnotherThread(string operation, out Func<Exception?> failure)
    {
        Exception? thrown = null;
        Thread thread = new(() =>
        {
            try
            {
                AuditLog.AppendTo(_dir, operation, Array.Empty<(string, string?)>());
            }
            catch (Exception ex)
            {
                thrown = ex;
            }
        });
        thread.Start();
        failure = () => thrown;
        return thread;
    }

    private static string Seal(string covered) => covered + " crc=\"" + Crc(covered) + "\"";

    private static string Crc(string covered) => AuditCrc32.ToHex(BitwiseCrc32(new UTF8Encoding(false).GetBytes(covered)));

    /// <summary>CRC-32 bit by bit, independently of the table the product uses.</summary>
    private static uint BitwiseCrc32(byte[] bytes)
    {
        uint crc = 0xFFFFFFFFu;
        foreach (byte b in bytes)
        {
            crc ^= b;
            for (int k = 0; k < 8; k++)
            {
                crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320u : crc >> 1;
            }
        }

        return ~crc;
    }
}
