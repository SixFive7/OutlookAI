using OutlookAI.Core.Audit;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 for the structured write-op audit log (v3.MD Phase 4: audit goes LIVE and
/// load-bearing). Format golden-shapes, escaping, validation, and real file appends
/// against a temp directory - never the shared %LOCALAPPDATA% location.
/// </summary>
public sealed class AuditLogTests
{
    private static readonly DateTime Ts = new(2026, 7, 23, 10, 11, 12, 345, DateTimeKind.Utc);

    [Fact]
    public void FormatLine_GoldenShape()
    {
        string line = AuditLog.FormatLine(Ts, "new_draft", new (string, string?)[]
        {
            ("entryId", "00AB"),
            ("store", "telefonie@xxlnet.nl"),
            ("displayed", "false"),
        });

        Assert.Equal(
            "ts=2026-07-23T10:11:12.345Z op=new_draft entryId=\"00AB\" store=\"telefonie@xxlnet.nl\" displayed=\"false\"",
            line);
    }

    [Fact]
    public void FormatLine_EscapesQuotesBackslashesAndControlChars()
    {
        string line = AuditLog.FormatLine(Ts, "op1", new (string, string?)[]
        {
            ("path", "C:\\dir\\file \"x\".txt"),
            ("note", "line1\r\nline2\tend"),
        });

        Assert.Contains("path=\"C:\\\\dir\\\\file \\\"x\\\".txt\"", line, StringComparison.Ordinal);
        Assert.Contains("note=\"line1\\r\\nline2\\tend\"", line, StringComparison.Ordinal);
        Assert.DoesNotContain('\n', line);
        Assert.DoesNotContain('\r', line);
    }

    [Fact]
    public void FormatLine_OmitsNullValues()
    {
        string line = AuditLog.FormatLine(Ts, "op", new (string, string?)[]
        {
            ("kept", "v"),
            ("dropped", null),
        });

        Assert.Contains("kept=\"v\"", line, StringComparison.Ordinal);
        Assert.DoesNotContain("dropped", line, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("")]
    [InlineData("  ")]
    [InlineData("bad op")]
    [InlineData("bad\"op")]
    public void FormatLine_RejectsInvalidOperationTokens(string operation)
    {
        Assert.Throws<ArgumentException>(() => AuditLog.FormatLine(Ts, operation, Array.Empty<(string, string?)>()));
    }

    [Fact]
    public void FormatLine_RejectsInvalidFieldKeys()
    {
        Assert.Throws<ArgumentException>(() =>
            AuditLog.FormatLine(Ts, "op", new (string, string?)[] { ("bad key", "v") }));
        Assert.Throws<ArgumentException>(() =>
            AuditLog.FormatLine(Ts, "op", new (string, string?)[] { ("", "v") }));
    }

    [Fact]
    public void AppendTo_CreatesDirectoryAndAppendsLines()
    {
        string dir = Path.Combine(Path.GetTempPath(), "OutlookAI-AuditLogTests-" + Guid.NewGuid().ToString("N"));
        try
        {
            AuditLog.AppendTo(dir, "first_op", new (string, string?)[] { ("k", "v1") });
            AuditLog.AppendTo(dir, "second_op", new (string, string?)[] { ("k", "v2") });

            string[] lines = File.ReadAllLines(Path.Combine(dir, "audit.log"));
            Assert.Equal(2, lines.Length);
            Assert.Contains("op=first_op", lines[0], StringComparison.Ordinal);
            Assert.Contains("k=\"v1\"", lines[0], StringComparison.Ordinal);
            Assert.Contains("op=second_op", lines[1], StringComparison.Ordinal);
        }
        finally
        {
            if (Directory.Exists(dir))
            {
                Directory.Delete(dir, recursive: true);
            }
        }
    }

    [Fact]
    public void AppendTo_FailureThrows_InsteadOfSwallowing()
    {
        // A FILE where the directory should be makes CreateDirectory/open fail - the
        // load-bearing contract is that the failure SURFACES.
        string parent = Path.Combine(Path.GetTempPath(), "OutlookAI-AuditLogTests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(parent);
        string fileAsDir = Path.Combine(parent, "not-a-directory");
        File.WriteAllText(fileAsDir, "x");
        try
        {
            Assert.Throws<InvalidOperationException>(() =>
                AuditLog.AppendTo(fileAsDir, "op", Array.Empty<(string, string?)>()));
        }
        finally
        {
            Directory.Delete(parent, recursive: true);
        }
    }

    [Fact]
    public void ConcurrentWriters_LoseNoLine_AndGarbleNone()
    {
        // Several server processes - one per agent session - append to the same log. Each append
        // opens at the end-of-file it sees and writes there, so two that overlap used to land on
        // the same offset and one line silently overwrote the other: measured 2026-10-03, two
        // processes appending 3,000 lines each kept 5,883, with no exception anywhere. The writer
        // lock serializes them. Threads stand in for processes here; the lock is a mutex NAMED in
        // the session's namespace, so it is the same object either way (checked across two real
        // processes when it was written: 6,000 of 6,000 lines, against 5,883 without it).
        string dir = Path.Combine(Path.GetTempPath(), "OutlookAI-AuditLogTests-" + Guid.NewGuid().ToString("N"));
        const int Writers = 2;
        const int PerWriter = 2000;
        try
        {
            using Barrier start = new(Writers);
            Exception?[] failures = new Exception?[Writers];
            Thread[] threads = Enumerable.Range(0, Writers).Select(w => new Thread(() =>
            {
                try
                {
                    start.SignalAndWait();
                    for (int i = 0; i < PerWriter; i++)
                    {
                        AuditLog.AppendTo(dir, "writer" + w.ToString(System.Globalization.CultureInfo.InvariantCulture), new (string, string?)[]
                        {
                            ("n", i.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                        });
                    }
                }
                catch (Exception ex)
                {
                    failures[w] = ex;
                }
            })).ToArray();

            foreach (Thread thread in threads)
            {
                thread.Start();
            }

            foreach (Thread thread in threads)
            {
                thread.Join();
            }

            Assert.All(failures, Assert.Null);

            string[] lines = File.ReadAllLines(Path.Combine(dir, AuditLog.LogFileName));
            Assert.Equal(Writers * PerWriter, lines.Length);
            for (int w = 0; w < Writers; w++)
            {
                List<AuditLogEntry> mine = lines
                    .Select(AuditLogReader.ParseLine)
                    .Where(e => e != null && e.Operation == "writer" + w.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    .Select(e => e!)
                    .ToList();

                // Every line of every writer, intact and in that writer's own order.
                Assert.Equal(
                    Enumerable.Range(0, PerWriter).Select(i => i.ToString(System.Globalization.CultureInfo.InvariantCulture)),
                    mine.Select(e => e.Fields.Single().Value));
            }
        }
        finally
        {
            if (Directory.Exists(dir))
            {
                Directory.Delete(dir, recursive: true);
            }
        }
    }

    [Fact]
    public void TheWriterLock_IsOnePerLogFile_WhateverTheSpelling()
    {
        string temp = Path.GetTempPath();
        string log = Path.Combine(temp, "OutlookAI-AuditLogTests-lock", AuditLog.LogFileName);

        // Same file, same lock: case, separators and a dot-dot detour do not split it.
        Assert.Equal(AuditLog.WriterMutexName(log), AuditLog.WriterMutexName(log.ToUpperInvariant()));
        Assert.Equal(
            AuditLog.WriterMutexName(log),
            AuditLog.WriterMutexName(Path.Combine(temp, "OutlookAI-AuditLogTests-lock", "..", "OutlookAI-AuditLogTests-lock", AuditLog.LogFileName)));

        // Another file, another lock - which is what keeps a test run's throwaway log from ever
        // waiting on the real one.
        Assert.NotEqual(
            AuditLog.WriterMutexName(log),
            AuditLog.WriterMutexName(Path.Combine(temp, "OutlookAI-AuditLogTests-other", AuditLog.LogFileName)));

        // A session-local name the OS accepts: one backslash, after the namespace.
        string name = AuditLog.WriterMutexName(log);
        Assert.StartsWith(@"Local\OutlookAI.AuditLog.", name, StringComparison.Ordinal);
        Assert.Equal(1, name.Count(c => c == '\\'));
    }

    [Fact]
    public void DefaultPaths_LiveUnderTheSharedStateRoot()
    {
        string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        Assert.Equal(Path.Combine(localAppData, "OutlookAI"), AuditLog.DefaultDirectory);
        Assert.Equal(Path.Combine(localAppData, "OutlookAI", "audit.log"), AuditLog.DefaultLogPath);
    }
}
