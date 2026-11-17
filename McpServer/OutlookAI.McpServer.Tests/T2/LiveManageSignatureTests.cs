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
    [Trait("Requires", "OutlookInstance")]
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
    [Trait("Requires", "MailAccount")]
    public void DefaultAssignment_HubAccountOnly_SetAndClearedWithoutRewriting_OriginalsRestoredExactly()
    {
        string name = _fixture.TestSignatureName("Dflt");
        string directory = SignatureCatalog.DefaultSignatureDirectory;
        ProfileSignatureDefaultsStore store = new ProfileSignatureDefaultsStore();

        // The profile Outlook runs - named, never guessed (Q155): the one list_signatures marks.
        SignaturesOutcome listing = Service.ListSignatures();
        string? profile = listing.RunningProfile ?? listing.DefaultProfile;
        SignatureDefaultsRow? hubRow = store.ReadAccounts().FirstOrDefault(r =>
            ProfileAccountEntries.ProfileEquals(r.Profile, profile) && ProfileAccountEntries.AddressEquals(r.Account, Hub));
        if (hubRow == null)
        {
            // Same rule as the delegate probe: on a real profile the hub account IS in the
            // signature registry, so its absence is a fault to report rather than a reason to
            // pass. A machine with no mail accounts has no rows at all, which is why this test
            // declares Requires=MailAccount.
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
        _output.WriteLine($"hub originals in '{profile}': new={(originalNew ?? "<absent>")} reply={(originalReply ?? "<absent>")}");
        List<string> backupsToCleanUp = new();

        try
        {
            Service.ManageSignature(new ManageSignatureRequest
            {
                Action = "create",
                Name = name,
                BodyText = "Default-assignment test " + _fixture.RunMarker,
            });
            Dictionary<string, byte[]> files = Directory.GetFiles(directory, name + ".*").ToDictionary(f => f, File.ReadAllBytes);

            // set_default ON ITS OWN (the defect: a default could only be set by rewriting the signature).
            ManageSignatureOutcome set = Service.ManageSignature(new ManageSignatureRequest
            {
                Action = "set_default",
                Name = name,
                DefaultForProfile = profile,
                DefaultForAccount = hubRow.Account,
                DefaultForScope = "both",
            });
            Assert.Equal(hubRow.Account, set.DefaultSetForAccount);
            Assert.Equal(profile, set.Profile, ignoreCase: true);
            Assert.Null(set.FilesWritten);
            Assert.Null(set.BackupPath);
            Assert.Equal(files, Directory.GetFiles(directory, name + ".*").ToDictionary(f => f, File.ReadAllBytes));
            Assert.Contains("next start", set.Advice, StringComparison.OrdinalIgnoreCase);

            SignatureDefaultsRow afterSet = RequireHubRow(store, hubRow.AccountKey);
            Assert.Equal(name, afterSet.NewMessage);
            Assert.Equal(name, afterSet.ReplyForward);

            // list_signatures reflects the assignment, in that profile, as own and as effective default.
            SignatureAccountView hubView = Service.ListSignatures().Profiles
                .Single(p => ProfileAccountEntries.ProfileEquals(p.Name, profile)).Accounts
                .Single(a => ProfileAccountEntries.AddressEquals(a.Account, hubRow.Account));
            Assert.Equal(name, hubView.NewMessage);
            Assert.Equal(name, hubView.EffectiveReplyForward);

            // clear_default clears only the default it names.
            ManageSignatureOutcome cleared = Service.ManageSignature(new ManageSignatureRequest
            {
                Action = "clear_default",
                Name = name,
                DefaultForProfile = profile,
                DefaultForAccount = hubRow.Account,
                DefaultForScope = "reply",
            });
            Assert.Equal(name, Assert.Single(cleared.DefaultsChanged!).Previous);
            Assert.Null(RequireHubRow(store, hubRow.AccountKey).ReplyForward);

            // DELETE clears the default still naming it, in whatever profile.
            ManageSignatureOutcome deleted = Service.ManageSignature(new ManageSignatureRequest
            {
                Action = "delete",
                Name = name,
            });
            if (deleted.BackupPath != null)
            {
                backupsToCleanUp.Add(deleted.BackupPath);
            }

            SignatureDefaultChange swept = Assert.Single(deleted.DefaultsCleared!);
            Assert.Equal(("new", hubRow.Account), (swept.Scope, swept.Account));
            Assert.Null(RequireHubRow(store, hubRow.AccountKey).NewMessage);
            _output.WriteLine("create -> set_default (files untouched) -> clear_default (reply) -> delete (new swept)");
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

        SignatureDefaultsRow restored = RequireHubRow(store, hubRow.AccountKey);
        Assert.Equal(originalNew, restored.NewMessage);
        Assert.Equal(originalReply, restored.ReplyForward);
        _output.WriteLine("hub registry values restored exactly and asserted");
    }

    [Fact]
    [Trait("Requires", "OutlookInstance")]
    public void Update_EditsInPlace_PicturesAndUntouchedRenditionsSurviveByteForByte()
    {
        // Q159 on the real folder: a signature with a logo in its picture folder AND embedded in its
        // RTF is tweaked - one phone number - and nothing else may move. The update used to delete the
        // picture folder and rebuild the RTF from plain text.
        string name = _fixture.TestSignatureName("Fid");
        string directory = SignatureCatalog.DefaultSignatureDirectory;
        string logo = Path.Combine(Path.GetTempPath(), "OutlookAI-McpTest-logo-" + _fixture.RunMarker + ".png");
        File.WriteAllBytes(logo, TinyPng());
        List<string> backupsToCleanUp = new();
        try
        {
            ManageSignatureOutcome created = Service.ManageSignature(new ManageSignatureRequest
            {
                Action = "create",
                Name = name,
                BodyHtml = "<p>Jan Modaal</p><p>Tel 010 123 " + _fixture.RunMarker + "</p><p><img src=\""
                    + SignatureHtml.ReferenceFor(name, "logo.png") + "\" width=\"1\" height=\"1\"></p>",
                Pictures = new[] { new SignaturePictureChange { File = "logo.png", Path = logo } },
            });
            Assert.Equal(new[] { "logo.png" }, created.PicturesAdded);
            string picture = Path.Combine(directory, name + "_files", "logo.png");
            byte[] pictureBefore = File.ReadAllBytes(picture);
            Assert.Single(SignatureRtf.FindPictures(File.ReadAllText(Path.Combine(directory, name + ".rtf"))));

            ManageSignatureOutcome updated = Service.ManageSignature(new ManageSignatureRequest
            {
                Action = "update",
                Name = name,
                Replace = new[] { new SignatureTextEdit { Find = "Tel 010 123", Replace = "Tel 010 999" } },
            });
            backupsToCleanUp.Add(updated.BackupPath!);

            SignatureReplacementView replaced = Assert.Single(updated.Replacements!);
            Assert.Equal((1, 1, 1), (replaced.Html, replaced.Rtf, replaced.Text));
            Assert.Equal(pictureBefore, File.ReadAllBytes(picture));
            string rtf = File.ReadAllText(Path.Combine(directory, name + ".rtf"));
            Assert.Contains("Tel 010 999", rtf, StringComparison.Ordinal);
            Assert.Single(SignatureRtf.FindPictures(rtf));
            SignatureDetailView detail = Service.ListSignatures(name).Signature!;
            Assert.Contains("Tel 010 999", detail.Html, StringComparison.Ordinal);
            Assert.True(Assert.Single(detail.Pictures!).ReferencedByHtml);
            Assert.Null(detail.MissingPictures);
            _output.WriteLine("in-place update: three renditions edited, the picture file and the embedded RTF picture intact");
        }
        finally
        {
            DeleteTestSignatureFiles(directory, name);
            foreach (string backup in backupsToCleanUp)
            {
                TryDeleteDirectory(backup);
            }

            File.Delete(logo);
        }
    }

    // ------------------------------------------------------------------ helpers

    private static SignatureDefaultsRow RequireHubRow(ProfileSignatureDefaultsStore store, string accountKey)
    {
        return store.ReadAccounts()
                .FirstOrDefault(r => ProfileAccountEntries.KeyEquals(r.AccountKey, accountKey))
            ?? throw new InvalidOperationException("hub account row vanished from the profile registry");
    }

    private static byte[] TinyPng()
    {
        return new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
            0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9C, 0x62, 0x00, 0x01, 0x00, 0x00,
            0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
            0x42, 0x60, 0x82,
        };
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

    // The log THIS process appends to. The service runs in-process here, so since Q86 its lines
    // go to the test run's throwaway directory (AuditIsolation), never the machine's real log.
    private static int CountAuditLines()
    {
        string path = AuditLog.EffectiveLogPath;
        return File.Exists(path) ? File.ReadAllLines(path).Length : 0;
    }

    private static IReadOnlyList<string> ReadAuditLinesAfter(int skip)
    {
        string path = AuditLog.EffectiveLogPath;
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
