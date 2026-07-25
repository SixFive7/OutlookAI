using System.Text.Json;
using OutlookAI.Core.Audit;
using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tests.T2;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// D38 wire acceptance: the manage_signature tool over REAL stdio MCP against the
/// real Signatures folder - create/update/delete of ONE "OutlookAI-McpTest-" named
/// signature (the only granted scope), asserting the JSON outcome shapes (camelCase
/// fields, backupPath under %LOCALAPPDATA%\OutlookAI\signature-backups, the
/// set_default_for {account,scope} object rejected pre-registry for an unknown
/// account) and the audit line per operation written by the SERVER process. Runs in
/// the guarded LiveSignatureManage collection: the fixture's snapshot proves the real
/// signatures bit-identical afterwards.
/// </summary>
[Collection("LiveSignatureManage")]
[Trait("Category", "Live")]
public sealed class LiveManageSignatureMcpToolTests
{
    private readonly LiveSignatureManageFixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveManageSignatureMcpToolTests(LiveSignatureManageFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    public async Task ManageSignature_FullCycle_OverRealStdio_WithBackupPathsAndAudit()
    {
        string name = _fixture.TestSignatureName("Wire");
        string directory = SignatureCatalog.DefaultSignatureDirectory;
        int auditBefore = CountAuditLines();
        List<string> backupsToCleanUp = new();

        try
        {
            await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(TimeSpan.FromMinutes(5));

            // CREATE over the wire.
            JsonElement created = await client.CallToolAsync("manage_signature", new
            {
                action = "create",
                name,
                body_text = "Wire lifecycle " + _fixture.RunMarker,
            });
            Assert.Equal("create", created.GetProperty("action").GetString());
            Assert.Equal(name, created.GetProperty("name").GetString());
            Assert.Equal(3, created.GetProperty("filesWritten").GetArrayLength());
            Assert.False(created.TryGetProperty("backupPath", out _), "create must not report a backup");
            Assert.True(File.Exists(Path.Combine(directory, name + ".htm")));

            // list_signatures over the wire sees it.
            JsonElement listed = await client.CallToolAsync("list_signatures", new { });
            Assert.Contains(listed.GetProperty("signatures").EnumerateArray(),
                s => s.GetProperty("name").GetString() == name);

            // UPDATE: backupPath present, under the documented root.
            JsonElement updated = await client.CallToolAsync("manage_signature", new
            {
                action = "update",
                name,
                body_html = "<p>Wire updated " + _fixture.RunMarker + "</p>",
            });
            string? updateBackup = updated.GetProperty("backupPath").GetString();
            Assert.NotNull(updateBackup);
            backupsToCleanUp.Add(updateBackup!);
            Assert.StartsWith(SignatureManager.DefaultBackupRoot, updateBackup!, StringComparison.OrdinalIgnoreCase);
            Assert.True(Directory.Exists(updateBackup!), "the update backup directory must exist");

            // Unknown account in the object-shaped set_default_for: rejected, files intact.
            JsonElement rejected = await client.CallToolAsync("manage_signature", new
            {
                action = "update",
                name,
                body_text = "should not land",
                set_default_for = new { account = "stranger@nowhere.test", scope = "new" },
            });
            Assert.Equal("InvalidArgument", rejected.GetProperty("error").GetProperty("type").GetString());
            Assert.Contains("not found", rejected.GetProperty("error").GetProperty("message").GetString(),
                StringComparison.OrdinalIgnoreCase);

            // DELETE: backup again, files gone.
            JsonElement deleted = await client.CallToolAsync("manage_signature", new { action = "delete", name });
            string? deleteBackup = deleted.GetProperty("backupPath").GetString();
            Assert.NotNull(deleteBackup);
            backupsToCleanUp.Add(deleteBackup!);
            Assert.True(deleted.GetProperty("filesDeleted").GetArrayLength() >= 3);
            Assert.False(File.Exists(Path.Combine(directory, name + ".htm")));

            // Audit lines written by the SERVER process (load-bearing contract).
            IReadOnlyList<string> lines = ReadAuditLinesAfter(auditBefore);
            Assert.Contains(lines, l => l.Contains(" op=manage_signature ", StringComparison.Ordinal)
                && l.Contains("action=\"create\"", StringComparison.Ordinal)
                && l.Contains("name=\"" + name + "\"", StringComparison.Ordinal));
            Assert.Contains(lines, l => l.Contains(" op=manage_signature ", StringComparison.Ordinal)
                && l.Contains("action=\"delete\"", StringComparison.Ordinal)
                && l.Contains("name=\"" + name + "\"", StringComparison.Ordinal));
            _output.WriteLine("wire cycle ok: shapes + 2 backups + audit lines from the server process");
        }
        finally
        {
            foreach (string extension in new[] { ".htm", ".html", ".rtf", ".txt" })
            {
                string path = Path.Combine(directory, name + extension);
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }

            string resources = Path.Combine(directory, name + "_files");
            if (Directory.Exists(resources))
            {
                Directory.Delete(resources, recursive: true);
            }

            foreach (string backup in backupsToCleanUp)
            {
                try
                {
                    if (Directory.Exists(backup))
                    {
                        Directory.Delete(backup, recursive: true);
                    }
                }
                catch (IOException)
                {
                }
            }
        }

        Assert.Empty(Directory.GetFileSystemEntries(directory, SignatureCatalog.TestSignaturePrefix + "*Wire*"));
    }

    private static int CountAuditLines()
    {
        string path = AuditLog.DefaultLogPath;
        return File.Exists(path) ? File.ReadAllLines(path).Length : 0;
    }

    private static IReadOnlyList<string> ReadAuditLinesAfter(int skip)
    {
        string path = AuditLog.DefaultLogPath;
        return File.Exists(path) ? File.ReadAllLines(path).Skip(skip).ToList() : new List<string>();
    }
}
