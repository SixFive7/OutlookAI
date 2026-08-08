using System.Text;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Soak fix D38 (manage_signature): pure-logic coverage of validation, rendition
/// derivation (.htm/.txt/.rtf encodings), the ALWAYS-ON pre-modification backup
/// (path shape + byte-identical content + abort-on-failure), file-set semantics of
/// create/update/delete, and the default-assignment registry writes/dangling-clear
/// against a fake defaults store. Everything runs in per-test temp directories -
/// the real Signatures folder is never touched by T1.
/// </summary>
public sealed class SignatureManagerTests : IDisposable
{
    private readonly string _dir;
    private readonly string _backupRoot;

    public SignatureManagerTests()
    {
        string root = Path.Combine(Path.GetTempPath(), "OutlookAI-SigManagerTests-" + Guid.NewGuid().ToString("N"));
        _dir = Path.Combine(root, "Signatures");
        _backupRoot = Path.Combine(root, "backups");
        Directory.CreateDirectory(_dir);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(Path.GetDirectoryName(_dir)!, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    private ManageSignatureOutcome Manage(
        ManageSignatureRequest request, ISignatureDefaultsStore? store = null, Func<DateTime>? utcNow = null)
    {
        return SignatureManager.Manage(request, _dir, _backupRoot, store ?? new FakeDefaultsStore(), utcNow);
    }

    // ------------------------------------------------------------------ create + derivation

    [Fact]
    public void Create_FromTextOnly_WritesAllThreeRenditions_WithTheDocumentedEncodings()
    {
        ManageSignatureOutcome outcome = Manage(new ManageSignatureRequest
        {
            Action = "create",
            Name = "My Sig",
            BodyText = "Met vriendelijke groet,\r\nJan Modaal - Café",
        });

        Assert.Equal("create", outcome.Action);
        Assert.Null(outcome.BackupPath);
        Assert.NotNull(outcome.FilesWritten);
        Assert.Equal(3, outcome.FilesWritten!.Count);

        string htmPath = Path.Combine(_dir, "My Sig.htm");
        string txtPath = Path.Combine(_dir, "My Sig.txt");
        string rtfPath = Path.Combine(_dir, "My Sig.rtf");
        Assert.True(File.Exists(htmPath));
        Assert.True(File.Exists(txtPath));
        Assert.True(File.Exists(rtfPath));

        // .htm: UTF-8 WITHOUT BOM + explicit charset meta (research-backed convention).
        byte[] htmBytes = File.ReadAllBytes(htmPath);
        Assert.False(htmBytes.Length >= 3 && htmBytes[0] == 0xEF && htmBytes[1] == 0xBB && htmBytes[2] == 0xBF,
            ".htm must not carry a UTF-8 BOM");
        string html = Encoding.UTF8.GetString(htmBytes);
        Assert.Contains("charset=utf-8", html, StringComparison.OrdinalIgnoreCase);
        // Non-ASCII is entity-encoded by the derivation (renders identically, immune
        // to charset mishaps in HTML mail clients).
        Assert.Contains("Caf&#233;", html, StringComparison.Ordinal);

        // .txt: UTF-16 LE WITH BOM (what Outlook itself accepts).
        byte[] txtBytes = File.ReadAllBytes(txtPath);
        Assert.True(txtBytes.Length >= 2 && txtBytes[0] == 0xFF && txtBytes[1] == 0xFE, ".txt must be UTF-16 LE with BOM");

        // .rtf: ASCII with escapes; needed because RTF-format mail reads ONLY .rtf.
        string rtf = File.ReadAllText(rtfPath);
        Assert.StartsWith(@"{\rtf1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\u233?", rtf, StringComparison.Ordinal); // é
        Assert.Contains(@"\par", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Create_FromHtmlOnly_DerivesThePlainTextRendition()
    {
        Manage(new ManageSignatureRequest
        {
            Action = "create",
            Name = "HtmlOnly",
            BodyHtml = "<p>Kind regards,</p><p>Derived Person</p>",
        });

        string text = File.ReadAllText(Path.Combine(_dir, "HtmlOnly.txt"));
        Assert.Contains("Kind regards,", text, StringComparison.Ordinal);
        Assert.Contains("Derived Person", text, StringComparison.Ordinal);
        Assert.DoesNotContain("<p>", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Create_FullHtmlDocument_IsWrittenAsIs_FragmentGetsWrapped()
    {
        string fullDoc = "<html><head><meta charset=\"utf-8\"></head><body><p>Doc</p></body></html>";
        Manage(new ManageSignatureRequest { Action = "create", Name = "FullDoc", BodyHtml = fullDoc });
        Assert.Equal(fullDoc, File.ReadAllText(Path.Combine(_dir, "FullDoc.htm")));

        Manage(new ManageSignatureRequest { Action = "create", Name = "Fragment", BodyHtml = "<p>Frag</p>" });
        string wrapped = File.ReadAllText(Path.Combine(_dir, "Fragment.htm"));
        Assert.Contains("<html>", wrapped, StringComparison.Ordinal);
        Assert.Contains("charset=utf-8", wrapped, StringComparison.Ordinal);
        Assert.Contains("<p>Frag</p>", wrapped, StringComparison.Ordinal);
    }

    [Fact]
    public void Create_ExistingName_IsRejected_PointingAtUpdate()
    {
        File.WriteAllText(Path.Combine(_dir, "Taken.htm"), "<p>x</p>");

        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            Manage(new ManageSignatureRequest { Action = "create", Name = "Taken", BodyText = "x" }));
        Assert.Contains("update", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ------------------------------------------------------------------ backup (ALWAYS-ON safety)

    [Fact]
    public void BackupPath_HasTheDocumentedShape_UtcTimestampDashName()
    {
        SeedSignature("Shaped", withResources: false);
        DateTime fixedUtc = new DateTime(2026, 7, 25, 3, 4, 5, 678, DateTimeKind.Utc);

        ManageSignatureOutcome outcome = Manage(
            new ManageSignatureRequest { Action = "update", Name = "Shaped", BodyText = "new" },
            utcNow: () => fixedUtc);

        Assert.Equal(Path.Combine(_backupRoot, "20260725T030405678Z-Shaped"), outcome.BackupPath);
        Assert.True(Directory.Exists(outcome.BackupPath!));
    }

    [Fact]
    public void Update_BacksUpThePreviousFileSet_ByteIdentical_IncludingResources()
    {
        byte[] resourceBytes = { 1, 2, 3, 4, 5 };
        SeedSignature("Rich", withResources: true, resourceBytes);
        string oldHtml = File.ReadAllText(Path.Combine(_dir, "Rich.htm"));

        ManageSignatureOutcome outcome = Manage(new ManageSignatureRequest
        {
            Action = "update",
            Name = "Rich",
            BodyText = "Replaced content",
        });

        Assert.NotNull(outcome.BackupPath);
        Assert.Equal(oldHtml, File.ReadAllText(Path.Combine(outcome.BackupPath!, "Rich.htm")));
        Assert.Equal("old text", File.ReadAllText(Path.Combine(outcome.BackupPath!, "Rich.txt")));
        Assert.Equal(resourceBytes, File.ReadAllBytes(Path.Combine(outcome.BackupPath!, "Rich_files", "img.png")));

        // The update replaced the renditions and removed the now-orphaned resources
        // (all preserved in the backup).
        Assert.Contains("Replaced content", File.ReadAllText(Path.Combine(_dir, "Rich.htm")), StringComparison.Ordinal);
        Assert.False(Directory.Exists(Path.Combine(_dir, "Rich_files")));
        Assert.True(File.Exists(Path.Combine(_dir, "Rich.rtf")));
    }

    [Fact]
    public void Delete_RemovesTheFileSet_AfterBackingItUp()
    {
        SeedSignature("Doomed", withResources: true, new byte[] { 9, 9 });

        ManageSignatureOutcome outcome = Manage(new ManageSignatureRequest { Action = "delete", Name = "Doomed" });

        Assert.NotNull(outcome.BackupPath);
        Assert.NotNull(outcome.FilesDeleted);
        Assert.Equal(3, outcome.FilesDeleted!.Count); // .htm + .txt + _files
        Assert.False(File.Exists(Path.Combine(_dir, "Doomed.htm")));
        Assert.False(File.Exists(Path.Combine(_dir, "Doomed.txt")));
        Assert.False(Directory.Exists(Path.Combine(_dir, "Doomed_files")));
        Assert.True(File.Exists(Path.Combine(outcome.BackupPath!, "Doomed.htm")));
        Assert.Equal(new byte[] { 9, 9 }, File.ReadAllBytes(Path.Combine(outcome.BackupPath!, "Doomed_files", "img.png")));
    }

    [Fact]
    public void FailingBackup_AbortsTheOperation_NothingModified()
    {
        SeedSignature("Protected", withResources: false);

        // Make the backup root unusable: a FILE where the directory must go.
        Directory.CreateDirectory(Path.GetDirectoryName(_backupRoot)!);
        File.WriteAllText(_backupRoot, "not a directory");

        Assert.Throws<InvalidOperationException>(() =>
            Manage(new ManageSignatureRequest { Action = "delete", Name = "Protected" }));

        Assert.True(File.Exists(Path.Combine(_dir, "Protected.htm")), "a failed backup must leave the signature untouched");
    }

    [Fact]
    public void UpdateAndDelete_OfAMissingSignature_AreRejected()
    {
        ArgumentException update = Assert.Throws<ArgumentException>(() =>
            Manage(new ManageSignatureRequest { Action = "update", Name = "Ghost", BodyText = "x" }));
        Assert.Contains("not found", update.Message, StringComparison.OrdinalIgnoreCase);

        ArgumentException delete = Assert.Throws<ArgumentException>(() =>
            Manage(new ManageSignatureRequest { Action = "delete", Name = "Ghost" }));
        Assert.Contains("not found", delete.Message, StringComparison.OrdinalIgnoreCase);
    }

    // ------------------------------------------------------------------ default assignments

    [Fact]
    public void SetDefaultFor_WritesTheRequestedScopes_AsRegSzValues()
    {
        FakeDefaultsStore store = new FakeDefaultsStore();
        store.Rows.Add(new SignatureDefaultsRow("key-hub", "hub@example.com", null, null));

        ManageSignatureOutcome outcome = Manage(new ManageSignatureRequest
        {
            Action = "create",
            Name = "Steered",
            BodyText = "x",
            DefaultForAccount = "HUB@example.com", // case-insensitive account match
            DefaultForScope = "both",
        }, store);

        Assert.Equal("hub@example.com", outcome.DefaultSetForAccount);
        Assert.Equal("both", outcome.DefaultSetScope);
        Assert.Contains(("key-hub", SignatureManager.NewSignatureValueName, "Steered"), store.Writes);
        Assert.Contains(("key-hub", SignatureManager.ReplyForwardSignatureValueName, "Steered"), store.Writes);
        Assert.Contains("next start", outcome.Advice, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("new", SignatureManager.NewSignatureValueName)]
    [InlineData("reply", SignatureManager.ReplyForwardSignatureValueName)]
    public void SetDefaultFor_SingleScope_WritesOnlyThatValue(string scope, string expectedValueName)
    {
        FakeDefaultsStore store = new FakeDefaultsStore();
        store.Rows.Add(new SignatureDefaultsRow("k", "a@b.nl", null, null));

        Manage(new ManageSignatureRequest
        {
            Action = "create",
            Name = "One" + scope,
            BodyText = "x",
            DefaultForAccount = "a@b.nl",
            DefaultForScope = scope,
        }, store);

        (string, string, string) write = Assert.Single(store.Writes);
        Assert.Equal(expectedValueName, write.Item2);
    }

    [Fact]
    public void SetDefaultFor_UnknownAccount_IsRejectedBeforeAnyFileWork()
    {
        Assert.Throws<ArgumentException>(() => Manage(new ManageSignatureRequest
        {
            Action = "create",
            Name = "Never",
            BodyText = "x",
            DefaultForAccount = "stranger@nowhere.test",
            DefaultForScope = "new",
        }));

        Assert.False(File.Exists(Path.Combine(_dir, "Never.htm")), "validation must reject before writing files");
    }

    [Fact]
    public void Delete_ClearsDanglingDefaults_OnEveryAccountThatReferencedIt()
    {
        SeedSignature("Assigned", withResources: false);
        FakeDefaultsStore store = new FakeDefaultsStore();
        store.Rows.Add(new SignatureDefaultsRow("k1", "one@example.com", "Assigned", null));
        store.Rows.Add(new SignatureDefaultsRow("k2", "two@example.com", "Other", "assigned")); // case-insensitive
        store.Rows.Add(new SignatureDefaultsRow("k3", "three@example.com", "Other", null));

        ManageSignatureOutcome outcome = Manage(
            new ManageSignatureRequest { Action = "delete", Name = "Assigned" }, store);

        Assert.NotNull(outcome.DefaultsClearedForAccounts);
        Assert.Equal(new[] { "one@example.com", "two@example.com" }, outcome.DefaultsClearedForAccounts);
        Assert.Contains(("k1", SignatureManager.NewSignatureValueName), store.Clears);
        Assert.Contains(("k2", SignatureManager.ReplyForwardSignatureValueName), store.Clears);
        Assert.DoesNotContain(store.Clears, c => c.Item1 == "k3");
        Assert.Contains("next start", outcome.Advice, StringComparison.OrdinalIgnoreCase);
    }

    // ------------------------------------------------------------------ validation matrix

    [Theory]
    [InlineData("destroy", "X", "b", null, null, null)]               // unknown action
    [InlineData("", "X", "b", null, null, null)]                      // empty action
    [InlineData("create", "", "b", null, null, null)]                 // empty name
    [InlineData("create", "a/b", "b", null, null, null)]              // path separator
    [InlineData("create", "a\\b", "b", null, null, null)]             // path separator
    [InlineData("create", "a?b", "b", null, null, null)]              // invalid file char
    [InlineData("create", "CON", "b", null, null, null)]              // reserved device name
    [InlineData("create", "dot.", "b", null, null, null)]             // trailing dot
    [InlineData("create", "X", null, null, null, null)]               // create without any body
    [InlineData("delete", "X", "b", null, null, null)]                // delete with body
    [InlineData("create", "X", "b", null, "a@b.nl", null)]            // account without scope
    [InlineData("create", "X", "b", null, null, "new")]               // scope without account
    [InlineData("create", "X", "b", null, "a@b.nl", "everything")]    // bad scope
    [InlineData("delete", "X", null, null, "a@b.nl", "new")]          // set_default_for on delete
    public void InvalidRequests_AreRejected_WithArgumentException(
        string action, string name, string? bodyText, string? bodyHtml, string? account, string? scope)
    {
        Assert.Throws<ArgumentException>(() => Manage(new ManageSignatureRequest
        {
            Action = action,
            Name = name,
            BodyText = bodyText,
            BodyHtml = bodyHtml,
            DefaultForAccount = account,
            DefaultForScope = scope,
        }));
    }

    [Fact]
    public void NameLength_IsCapped()
    {
        Assert.Throws<ArgumentException>(() => Manage(new ManageSignatureRequest
        {
            Action = "create",
            Name = new string('n', SignatureManager.NameMaxChars + 1),
            BodyText = "x",
        }));
    }

    [Fact]
    public void BodySize_IsCappedAtTheSharedBodyCap()
    {
        Assert.Throws<ArgumentException>(() => Manage(new ManageSignatureRequest
        {
            Action = "create",
            Name = "Big",
            BodyText = new string('x', MailService.BodyCharsCap + 1),
        }));
    }

    [Fact]
    public void RtfDerivation_EscapesSpecialsAndNonAscii()
    {
        string rtf = SignatureManager.BuildRtfFromText("a\\b {c} é\r\nsecond");
        Assert.Contains(@"a\\b", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\{c\}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\u233?", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\par second", rtf, StringComparison.Ordinal);
        Assert.EndsWith("}", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void ListedAfterCreate_GoneAfterDelete_ViaTheCatalog()
    {
        Manage(new ManageSignatureRequest { Action = "create", Name = "RoundTrip", BodyText = "Hello / World" });
        SignatureInfo? listed = SignatureCatalog.TryResolve("RoundTrip", _dir);
        Assert.NotNull(listed);
        Assert.NotNull(listed!.HtmlPath);
        Assert.NotNull(listed.TextPath);
        Assert.NotNull(listed.RtfPath);
        Assert.Contains("Hello", listed.Excerpt, StringComparison.Ordinal);

        Manage(new ManageSignatureRequest { Action = "delete", Name = "RoundTrip" });
        Assert.Null(SignatureCatalog.TryResolve("RoundTrip", _dir));
    }

    // ------------------------------------------------------------------ helpers

    private void SeedSignature(string name, bool withResources, byte[]? resourceBytes = null)
    {
        File.WriteAllText(Path.Combine(_dir, name + ".htm"), "<html><body><p>old html of " + name + "</p></body></html>");
        File.WriteAllText(Path.Combine(_dir, name + ".txt"), "old text");
        if (withResources)
        {
            string resources = Path.Combine(_dir, name + "_files");
            Directory.CreateDirectory(resources);
            File.WriteAllBytes(Path.Combine(resources, "img.png"), resourceBytes ?? new byte[] { 0 });
        }
    }

    private sealed class FakeDefaultsStore : ISignatureDefaultsStore
    {
        public List<SignatureDefaultsRow> Rows { get; } = new List<SignatureDefaultsRow>();

        public List<(string, string, string)> Writes { get; } = new List<(string, string, string)>();

        public List<(string, string)> Clears { get; } = new List<(string, string)>();

        public IReadOnlyList<SignatureDefaultsRow> ReadAccounts()
        {
            return Rows;
        }

        public void WriteDefault(string accountKey, string valueName, string signatureName)
        {
            Writes.Add((accountKey, valueName, signatureName));
        }

        public void ClearDefault(string accountKey, string valueName)
        {
            Clears.Add((accountKey, valueName));
        }
    }
}
