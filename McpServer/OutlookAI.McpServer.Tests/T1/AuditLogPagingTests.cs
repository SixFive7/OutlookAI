using System.Globalization;
using System.Text;
using OutlookAI.Core.Audit;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 for reading the audit log a page at a time (Q119, 2026-10-03). The position is a byte offset in
/// the file plus the file's identity and an anchor on the line it points at, so the tests here are the
/// three ways a timestamp-based position failed or could fail: lines that share a millisecond (the old
/// after/before "paging" skipped 122 of them over the real log at page size 25), lines appended between
/// two pages, and a log archived, replaced, truncated or edited under a position.
/// </summary>
public sealed class AuditLogPagingTests : IDisposable
{
    private static readonly DateTime Ts = new(2026, 10, 3, 9, 0, 0, 0, DateTimeKind.Utc);

    private readonly string _dir = Path.Combine(Path.GetTempPath(), "OutlookAI-AuditLogPagingTests-" + Guid.NewGuid().ToString("N"));
    private long _seq;

    public AuditLogPagingTests()
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

    [Fact]
    public void PagesWalkBackThroughLinesThatShareAMillisecond_SkippingNoneAndRepeatingNone()
    {
        // 130 lines in runs of ten that share one millisecond each.
        for (int i = 0; i < 130; i++)
        {
            Write(Ts.AddMilliseconds(i / 10), i);
        }

        List<int> seen = new();
        List<long> olderAfterEachPage = new();
        AuditLogPosition? next = null;
        int pages = 0;
        do
        {
            AuditLogScan page = AuditLogReader.Read(LogPath, All(), 25, next);
            Assert.Equal(next == null ? AuditLogResumeStatus.NotResumed : AuditLogResumeStatus.Resumed, page.Resume);
            Assert.Equal(130, page.Matched);
            seen.AddRange(page.Entries.Select(N));
            olderAfterEachPage.Add(page.OlderMatches);
            next = page.NextPosition;
            pages++;
        }
        while (next != null);

        Assert.Equal(Enumerable.Range(0, 130).Reverse(), seen);
        Assert.Equal(6, pages);
        Assert.Equal(new long[] { 105, 80, 55, 30, 5, 0 }, olderAfterEachPage);
    }

    [Fact]
    public void LinesAppendedBetweenPages_DoNotShiftTheChain_AndAreCountedInMatched()
    {
        for (int i = 0; i < 10; i++)
        {
            Write(Ts, i);
        }

        AuditLogScan first = AuditLogReader.Read(LogPath, All(), 4);
        Assert.Equal(new[] { 9, 8, 7, 6 }, first.Entries.Select(N));

        for (int i = 10; i < 15; i++)
        {
            Write(Ts, i);
        }

        AuditLogScan second = AuditLogReader.Read(LogPath, All(), 4, first.NextPosition);
        Assert.Equal(new[] { 5, 4, 3, 2 }, second.Entries.Select(N));
        Assert.Equal(15, second.Matched);
        Assert.Equal(2, second.OlderMatches);

        AuditLogScan third = AuditLogReader.Read(LogPath, All(), 4, second.NextPosition);
        Assert.Equal(new[] { 1, 0 }, third.Entries.Select(N));
        Assert.Null(third.NextPosition);
        Assert.Equal(0, third.OlderMatches);
    }

    [Fact]
    public void AFilteredChain_PagesTheMatchesOnly()
    {
        for (int i = 0; i < 20; i++)
        {
            Write(Ts.AddSeconds(i), i, i % 2 == 0 ? "even" : "odd");
        }

        AuditLogFilter even = new(null, null, new[] { "even" }, null);
        AuditLogScan first = AuditLogReader.Read(LogPath, even, 3);
        AuditLogScan second = AuditLogReader.Read(LogPath, even, 3, first.NextPosition);

        Assert.Equal(new[] { 18, 16, 14 }, first.Entries.Select(N));
        Assert.Equal(new[] { 12, 10, 8 }, second.Entries.Select(N));
        Assert.Equal(10, second.Matched);
        Assert.Equal(4, second.OlderMatches);
    }

    [Fact]
    public void APositionIntoALogArchivedSince_IsRefused_WhetherOrNotANewLogExists()
    {
        for (int i = 0; i < 10; i++)
        {
            Write(Ts, i);
        }

        AuditLogPosition next = AuditLogReader.Read(LogPath, All(), 4).NextPosition!;
        Assert.NotNull(next.FileIdentity);
        File.Move(LogPath, Path.Combine(_dir, "audit.until-test.log"));

        AuditLogScan missing = AuditLogReader.Read(LogPath, All(), 4, next);
        Assert.Equal(AuditLogResumeStatus.FileMissing, missing.Resume);
        Assert.Empty(missing.Entries);

        // A new log under the old name - long enough that the offset falls inside it.
        for (int i = 100; i < 140; i++)
        {
            Write(Ts, i);
        }

        AuditLogScan replaced = AuditLogReader.Read(LogPath, All(), 4, next);
        Assert.Equal(AuditLogResumeStatus.FileReplaced, replaced.Resume);
        Assert.Empty(replaced.Entries);
        Assert.Null(replaced.NextPosition);
    }

    [Fact]
    public void APositionIntoALogEditedInPlace_IsRefusedByItsAnchor()
    {
        for (int i = 0; i < 10; i++)
        {
            Write(Ts, i);
        }

        AuditLogPosition next = AuditLogReader.Read(LogPath, All(), 4).NextPosition!;

        // Same file, same identity - but the bytes moved: one line taken out at the start.
        string[] lines = File.ReadAllLines(LogPath);
        using (FileStream stream = new(LogPath, FileMode.Open, FileAccess.Write, FileShare.ReadWrite))
        {
            byte[] rest = Encoding.UTF8.GetBytes(string.Join("\n", lines.Skip(1)) + "\n");
            stream.Write(rest, 0, rest.Length);
            stream.SetLength(rest.Length);
        }

        Assert.Equal(AuditLogResumeStatus.PositionNotFound, AuditLogReader.Read(LogPath, All(), 4, next).Resume);

        // And truncated below the position altogether.
        using (FileStream stream = new(LogPath, FileMode.Open, FileAccess.Write, FileShare.ReadWrite))
        {
            stream.SetLength(10);
        }

        Assert.Equal(AuditLogResumeStatus.PositionNotFound, AuditLogReader.Read(LogPath, All(), 4, next).Resume);
    }

    [Fact]
    public void APositionWithoutAnIdentity_IsStillHeldToItsAnchor()
    {
        for (int i = 0; i < 10; i++)
        {
            Write(Ts, i);
        }

        AuditLogPosition next = AuditLogReader.Read(LogPath, All(), 4).NextPosition!;

        AuditLogScan good = AuditLogReader.Read(LogPath, All(), 4, new AuditLogPosition(null, next.Offset, next.Anchor));
        Assert.Equal(AuditLogResumeStatus.Resumed, good.Resume);
        Assert.Equal(new[] { 5, 4, 3, 2 }, good.Entries.Select(N));

        Assert.Equal(
            AuditLogResumeStatus.PositionNotFound,
            AuditLogReader.Read(LogPath, All(), 4, new AuditLogPosition(null, next.Offset, next.Anchor ^ 1)).Resume);
        Assert.Equal(
            AuditLogResumeStatus.PositionNotFound,
            AuditLogReader.Read(LogPath, All(), 4, new AuditLogPosition(null, next.Offset + 1, next.Anchor)).Resume);
    }

    // ================================================================== helpers

    private static AuditLogFilter All() => new(null, null, null, null);

    private static int N(AuditLogEntry entry) =>
        int.Parse(entry.Fields.Single(f => f.Key == "n").Value, CultureInfo.InvariantCulture);

    private void Write(DateTime ts, int n, string operation = "paged")
    {
        _seq++;
        string line = AuditLog.FormatSealedLine(
            ts, operation, new (string, string?)[] { ("n", n.ToString(CultureInfo.InvariantCulture)) }, 4242, "0a1b2c3d", _seq, null);
        using FileStream stream = new(LogPath, FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
        byte[] bytes = new UTF8Encoding(false).GetBytes(line + "\n");
        stream.Write(bytes, 0, bytes.Length);
    }
}
