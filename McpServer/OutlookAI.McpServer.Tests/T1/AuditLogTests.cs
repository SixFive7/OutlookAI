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
    public void DefaultPaths_LiveUnderTheSharedStateRoot()
    {
        string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        Assert.Equal(Path.Combine(localAppData, "OutlookAI"), AuditLog.DefaultDirectory);
        Assert.Equal(Path.Combine(localAppData, "OutlookAI", "audit.log"), AuditLog.DefaultLogPath);
    }
}
