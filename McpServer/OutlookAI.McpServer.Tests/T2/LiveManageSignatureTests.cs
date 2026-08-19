using OutlookAI.Core.Audit;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Soak fix D38 (manage_signature) live acceptance. ABSOLUTE scope rule (user-ordered):
/// live tests create/update/delete ONLY signatures named with the
/// "OutlookAI-McpTest-" prefix; the fixture's directory snapshot proves the user's
/// real signatures bit-identical afterwards. Coverage: the full create -> list ->
/// update (backup) -> delete (backup) lifecycle against the REAL Signatures folder
/// with audit lines per operation, and the default-assignment flow against the HUB
/// account only - original registry values read first, restored exactly in finally,
/// and the restoration asserted.
/// </summary>
[Collection(LiveCollections.SignatureManage)]
[Trait("Category", "Live")]
public sealed class LiveManageSignatureTests
{
    private readonly LiveSignatureManageFixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveManageSignatureTests(LiveSignatureManageFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    [Fact]
    [Trait("LiveTier", "Portable")]
    public void FullLifecycle_CreateUpdateDelete_WithAlwaysOnBackups_AndAuditLines()
    {
        string name = _fixture.TestSignatureName("Cycle");
        string directory = SignatureCatalog.DefaultSignatureDirectory;
        int auditBefore = CountAuditLines();
        List<string> backupsToCleanUp = new();

        try
        {
            // CREATE: no backup (nothing existed), three renditions written.
            ManageSignatureOutcome created = Service.ManageSignature(new ManageSignatureRequest
            {
                Action = "create",
                Name = name,
                BodyText = "Met vriendelijke groet,\r\nD38 lifecycle " + _fixture.RunMarker,
            });
            Assert.Null(created.BackupPath);
            Assert.NotNull(created.FilesWritten);
            Assert.Equal(3, created.FilesWritten!.Count);
            Assert.True(File.Exists(Path.Combine(directory, name + ".htm")));
            Assert.True(File.Exists(Path.Combine(directory, name + ".txt")));
            Assert.True(File.Exists(Path.Combine(directory, name + ".rtf")));

            SignatureView listed = Assert.Single(Service.ListSignatures().Signatures, s => s.Name == name);
            Assert.NotNull(listed.Excerpt);
            Assert.Contains("D38 lifecycle", listed.Excerpt, StringComparison.Ordinal);

            // UPDATE: pre-update bytes must land byte-identical in the returned backup.
            byte[] htmBefore = File.ReadAllBytes(Path.Combine(directory, name + ".htm"));
            byte[] txtBefore = File.ReadAllBytes(Path.Combine(directory, name + ".txt"));
            ManageSignatureOutcome updated = Service.ManageSignature(new ManageSignatureRequest
            {
                Action = "update",
                Name = name,
                BodyHtml = "<p>Updated rendition " + _fixture.RunMarker + "</p>",
            });
            Assert.NotNull(updated.BackupPath);
            backupsToCleanUp.Add(updated.BackupPath!);
            Assert.StartsWith(SignatureManager.DefaultBackupRoot, updated.BackupPath!, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(htmBefore, File.ReadAllBytes(Path.Combine(updated.BackupPath!, name + ".htm")));
            Assert.Equal(txtBefore, File.ReadAllBytes(Path.Combine(updated.BackupPath!, name + ".txt")));
            string derivedText = File.ReadAllText(Path.Combine(directory, name + ".txt"));
            Assert.Contains("Updated rendition", derivedText, StringComparison.Ordinal);

            // DELETE: second backup, file set gone, catalog no longer lists it.
            ManageSignatureOutcome deleted = Service.ManageSignature(new ManageSignatureRequest
            {
                Action = "delete",
                Name = name,
            });
            Assert.NotNull(deleted.BackupPath);
            backupsToCleanUp.Add(deleted.BackupPath!);
            Assert.NotEqual(updated.BackupPath, deleted.BackupPath);
            Assert.NotNull(deleted.FilesDeleted);
            Assert.False(File.Exists(Path.Combine(directory, name + ".htm")));
            Assert.False(File.Exists(Path.Combine(directory, name + ".rtf")));
            Assert.DoesNotContain(Service.ListSignatures().Signatures, s => s.Name == name);

            // Audit: one load-bearing line per operation.
            IReadOnlyList<string> lines = ReadAuditLinesAfter(auditBefore);
            Assert.True(IndexOfManageLine(lines, "create", name) >= 0, "create audit line missing");
            Assert.True(IndexOfManageLine(lines, "update", name) >= 0, "update audit line missing");
            Assert.True(IndexOfManageLine(lines, "delete", name) >= 0, "delete audit line missing");
            _output.WriteLine($"lifecycle ok: create/update/delete audited, backups at 2 distinct paths");
        }
        finally
        {
            DeleteTestSignatureFiles(directory, name);
            foreach (string backup in backupsToCleanUp)
            {
                TryDeleteDirectory(backup);
            }
        }

        Assert.Empty(Directory.GetFileSystemEntries(directory, SignatureCatalog.TestSignaturePrefix + "*Cycle*"));
    }

    [Fact]
    [Trait("LiveTier", "ProfileBound")]
    [Trait("Requires", "MailAccount")]
    public void DefaultAssignment_HubAccountOnly_SetThenCleared_OriginalsRestoredExactly()
    {
        string name = _fixture.TestSignatureName("Dflt");
        string directory = SignatureCatalog.DefaultSignatureDirectory;
        ProfileSignatureDefaultsStore store = new ProfileSignatureDefaultsStore();

        SignatureDefaultsRow? hubRow = store.ReadAccounts()
            .FirstOrDefault(r => string.Equals(r.Account, Hub, StringComparison.OrdinalIgnoreCase));
        if (hubRow == null)
        {
            // Same rule as the delegate probe: on a real profile the hub account IS in the
            // signature registry, so its absence is a fault to report rather than a reason to
            // pass. A machine with no mail accounts has no rows at all, which is why this test
            // is LiveTier=ProfileBound.
            _fixture.Settings.RequireProductionPopulation(
                "the hub account's row in the profile signature registry");
            _output.WriteLine(
                "PROVED NOTHING: hub account not present in the profile signature registry, so the "
                + "set-then-restore contract did not run.");
            return;
        }

        // Read-first (restore-exact contract): remember the ORIGINAL values verbatim.
        string? originalNew = hubRow.NewMessage;
        string? originalReply = hubRow.ReplyForward;
        _output.WriteLine($"hub originals: new={(originalNew ?? "<absent>")} reply={(originalReply ?? "<absent>")}");
        List<string> backupsToCleanUp = new();

        try
        {
            // CREATE with set_default_for both on the HUB (the only granted account).
            ManageSignatureOutcome created = Service.ManageSignature(new ManageSignatureRequest
            {
                Action = "create",
                Name = name,
                BodyText = "Default-assignment test " + _fixture.RunMarker,
                DefaultForAccount = hubRow.Account,
                DefaultForScope = "both",
            });
            Assert.Equal(hubRow.Account, created.DefaultSetForAccount);
            Assert.Equal("both", created.DefaultSetScope);
            Assert.Contains("next start", created.Advice, StringComparison.OrdinalIgnoreCase);

            SignatureDefaultsRow afterSet = RequireHubRow(store);
            Assert.Equal(name, afterSet.NewMessage);
            Assert.Equal(name, afterSet.ReplyForward);

            // list_signatures reflects the assignment (registry-read path).
            SignaturesOutcome outcome = Service.ListSignatures();
            SignatureAccountView hubView = Assert.Single(
                outcome.Accounts ?? Array.Empty<SignatureAccountView>().ToList(),
                a => string.Equals(a.Account, hubRow.Account, StringComparison.OrdinalIgnoreCase));
            Assert.Equal(name, hubView.NewMessage);

            // DELETE clears the now-dangling assignment.
            ManageSignatureOutcome deleted = Service.ManageSignature(new ManageSignatureRequest
            {
                Action = "delete",
                Name = name,
            });
            if (deleted.BackupPath != null)
            {
                backupsToCleanUp.Add(deleted.BackupPath);
            }

            Assert.NotNull(deleted.DefaultsClearedForAccounts);
            Assert.Contains(hubRow.Account, deleted.DefaultsClearedForAccounts!,
                StringComparer.OrdinalIgnoreCase);

            SignatureDefaultsRow afterDelete = RequireHubRow(store);
            Assert.Null(afterDelete.NewMessage);
            Assert.Null(afterDelete.ReplyForward);
            _output.WriteLine("set -> asserted -> delete -> dangling assignment cleared");
        }
        finally
        {
            // Restore-exact: write the original values back (or remove what we added
            // when the original was absent) - THEN assert the restoration held.
            RestoreValue(store, hubRow.AccountKey, SignatureManager.NewSignatureValueName, originalNew);
            RestoreValue(store, hubRow.AccountKey, SignatureManager.ReplyForwardSignatureValueName, originalReply);
            DeleteTestSignatureFiles(directory, name);
            foreach (string backup in backupsToCleanUp)
            {
                TryDeleteDirectory(backup);
            }
        }

        SignatureDefaultsRow restored = RequireHubRow(store);
        Assert.Equal(originalNew, restored.NewMessage);
        Assert.Equal(originalReply, restored.ReplyForward);
        _output.WriteLine("hub registry values restored exactly and asserted");
    }

    // ------------------------------------------------------------------ helpers

    private SignatureDefaultsRow RequireHubRow(ProfileSignatureDefaultsStore store)
    {
        return store.ReadAccounts()
                .FirstOrDefault(r => string.Equals(r.Account, Hub, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException("hub account row vanished from the profile registry");
    }

    private static void RestoreValue(
        ProfileSignatureDefaultsStore store, string accountKey, string valueName, string? original)
    {
        if (original != null)
        {
            store.WriteDefault(accountKey, valueName, original);
        }
        else
        {
            store.ClearDefault(accountKey, valueName);
        }
    }

    /// <summary>Exact-name deletion of the test signature's possible file set (no patterns beyond the pinned prefix name).</summary>
    private static void DeleteTestSignatureFiles(string directory, string name)
    {
        foreach (string extension in new[] { ".htm", ".html", ".rtf", ".txt" })
        {
            string path = Path.Combine(directory, name + extension);
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }

        TryDeleteDirectory(Path.Combine(directory, name + "_files"));
    }

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
            }
        }
        catch (IOException)
        {
        }
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

    private static int IndexOfManageLine(IReadOnlyList<string> lines, string action, string name)
    {
        for (int i = 0; i < lines.Count; i++)
        {
            if (lines[i].Contains(" op=manage_signature ", StringComparison.Ordinal)
                && lines[i].Contains("action=\"" + action + "\"", StringComparison.Ordinal)
                && lines[i].Contains("name=\"" + name + "\"", StringComparison.Ordinal))
            {
                return i;
            }
        }

        return -1;
    }
}
