using System.Text;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// manage_signature (soak fix D38; Q147, Q155, Q156, Q159): validation, the renditions a create
/// writes and derives, the ALWAYS-ON backup, in-place editing that touches only what was asked
/// (each rendition, its encoding and its pictures), defaults set and cleared on their own for an
/// explicit profile and account, and the policy and roaming refusals. Everything runs in per-test
/// temp directories against fake registries - the real Signatures folder and registry are never touched.
/// </summary>
public sealed class SignatureManagerTests : IDisposable
{
    private const string Profile = "Outlook";

    private readonly TempSignatureFolder _folder = new TempSignatureFolder();

    public void Dispose() => _folder.Dispose();

    private string Dir => _folder.Directory;

    private ManageSignatureOutcome Manage(
        ManageSignatureRequest request,
        ISignatureDefaultsStore? store = null,
        Func<DateTime>? utcNow = null,
        FakeSignatureSettings? settings = null)
    {
        return SignatureManager.Manage(request, Dir, _folder.BackupRoot, store ?? new FakeDefaultsStore(), utcNow, settings ?? new FakeSignatureSettings());
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
        Assert.Equal(new[] { "html", "rtf", "text" }, outcome.FormatsWritten);
        Assert.Equal(new[] { "html", "rtf" }, outcome.FormatsDerived);

        byte[] htmBytes = File.ReadAllBytes(Path.Combine(Dir, "My Sig.htm"));
        Assert.False(htmBytes.Length >= 3 && htmBytes[0] == 0xEF && htmBytes[1] == 0xBB && htmBytes[2] == 0xBF,
            ".htm must not carry a UTF-8 BOM");
        string html = Encoding.UTF8.GetString(htmBytes);
        Assert.Contains("charset=utf-8", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Caf&#233;", html, StringComparison.Ordinal);

        byte[] txtBytes = File.ReadAllBytes(Path.Combine(Dir, "My Sig.txt"));
        Assert.True(txtBytes.Length >= 2 && txtBytes[0] == 0xFF && txtBytes[1] == 0xFE, ".txt must be UTF-16 LE with BOM");

        string rtf = File.ReadAllText(Path.Combine(Dir, "My Sig.rtf"));
        Assert.StartsWith(@"{\rtf1", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\'e9", rtf, StringComparison.Ordinal); // é, in the code page the document names
        Assert.Contains(@"\par", rtf, StringComparison.Ordinal);
        Assert.True(File.ReadAllBytes(Path.Combine(Dir, "My Sig.rtf")).All(b => b < 0x80), ".rtf must be 7-bit");
    }

    [Fact]
    public void Create_FromHtmlOnly_DerivesTheTextAndAnRtfWithTheFormatting()
    {
        Manage(new ManageSignatureRequest
        {
            Action = "create",
            Name = "HtmlOnly",
            BodyHtml = "<p>Kind regards,</p><p><b>Derived</b> Person</p>",
        });

        string text = File.ReadAllText(Path.Combine(Dir, "HtmlOnly.txt"));
        Assert.Contains("Kind regards,", text, StringComparison.Ordinal);
        Assert.Contains("Derived Person", text, StringComparison.Ordinal);
        Assert.DoesNotContain("<p>", text, StringComparison.Ordinal);

        string rtf = File.ReadAllText(Path.Combine(Dir, "HtmlOnly.rtf"));
        Assert.Contains(@"{\b Derived}", rtf, StringComparison.Ordinal);
        Assert.Contains("Person", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void Create_WithAPicture_PutsItInThePictureFolder_AndEmbedsItInTheDerivedRtf()
    {
        byte[] logo = TestPictures.Png(40, 20);
        string source = _folder.Outside("logo.png", logo);

        ManageSignatureOutcome outcome = Manage(new ManageSignatureRequest
        {
            Action = "create",
            Name = "With Logo",
            BodyHtml = "<p>Jan</p><p><img src=\"With%20Logo_files/logo.png\" width=\"40\" height=\"20\"></p>",
            Pictures = new[] { new SignaturePictureChange { File = "logo.png", Path = source } },
        });

        Assert.Equal(new[] { "logo.png" }, outcome.PicturesAdded);
        Assert.Equal(logo, File.ReadAllBytes(Path.Combine(Dir, "With Logo_files", "logo.png")));
        string rtf = File.ReadAllText(Path.Combine(Dir, "With Logo.rtf"));
        Assert.Contains(@"\pngblip", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\picwgoal600", rtf, StringComparison.Ordinal); // 40 px at 15 twips a pixel
        Assert.Single(SignatureRtf.FindPictures(rtf));
        Assert.Equal(SignatureRtf.Sha256Hex(logo), SignatureRtf.FindPictures(rtf)[0].Sha256);
    }

    [Fact]
    public void Create_TheHtmlMayNotPointAtAPictureThatWillNotExist()
    {
        ArgumentException refusal = Assert.Throws<ArgumentException>(() => Manage(new ManageSignatureRequest
        {
            Action = "create",
            Name = "Broken",
            BodyHtml = "<p><img src=\"Broken_files/missing.png\"></p>",
        }));
        Assert.Contains("missing.png", refusal.Message, StringComparison.Ordinal);

        ArgumentException stray = Assert.Throws<ArgumentException>(() => Manage(new ManageSignatureRequest
        {
            Action = "create",
            Name = "Stray",
            BodyHtml = "<p><img src=\"logo.png\"></p>",
        }));
        Assert.Contains("Stray_files/<file>", stray.Message, StringComparison.Ordinal);
        Assert.Empty(Directory.GetFileSystemEntries(Dir));
    }

    [Fact]
    public void Create_FullHtmlDocument_IsWrittenAsIs_FragmentGetsWrapped()
    {
        string fullDoc = "<html><head><meta charset=\"utf-8\"></head><body><p>Doc</p></body></html>";
        Manage(new ManageSignatureRequest { Action = "create", Name = "FullDoc", BodyHtml = fullDoc });
        Assert.Equal(fullDoc, File.ReadAllText(Path.Combine(Dir, "FullDoc.htm")));

        Manage(new ManageSignatureRequest { Action = "create", Name = "Fragment", BodyHtml = "<p>Frag</p>" });
        string wrapped = File.ReadAllText(Path.Combine(Dir, "Fragment.htm"));
        Assert.Contains("<html>", wrapped, StringComparison.Ordinal);
        Assert.Contains("charset=utf-8", wrapped, StringComparison.Ordinal);
        Assert.Contains("<p>Frag</p>", wrapped, StringComparison.Ordinal);
    }

    [Fact]
    public void Create_ExistingName_IsRejected_PointingAtUpdate()
    {
        File.WriteAllText(Path.Combine(Dir, "Taken.htm"), "<p>x</p>");

        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            Manage(new ManageSignatureRequest { Action = "create", Name = "Taken", BodyText = "x" }));
        Assert.Contains("update", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TheReservedWordNone_IsNeverASignatureName()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            Manage(new ManageSignatureRequest { Action = "create", Name = "None", BodyText = "x" }));
        Assert.Contains("reserved", ex.Message, StringComparison.Ordinal);
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

        Assert.Equal(Path.Combine(_folder.BackupRoot, "20260725T030405678Z-Shaped"), outcome.BackupPath);
        Assert.True(Directory.Exists(outcome.BackupPath!));
    }

    [Fact]
    public void Update_BacksUpThePreviousFileSet_ByteIdentical_AndTouchesOnlyTheRenditionGiven()
    {
        // The Q159 defect: update used to delete the picture folder and rebuild the .rtf from plain text.
        byte[] resourceBytes = { 1, 2, 3, 4, 5 };
        SeedSignature("Rich", withResources: true, resourceBytes);
        File.WriteAllText(Path.Combine(Dir, "Rich.rtf"), @"{\rtf1\ansi Old {\b rich} text}");
        byte[] htmBefore = File.ReadAllBytes(Path.Combine(Dir, "Rich.htm"));
        byte[] rtfBefore = File.ReadAllBytes(Path.Combine(Dir, "Rich.rtf"));

        ManageSignatureOutcome outcome = Manage(new ManageSignatureRequest
        {
            Action = "update",
            Name = "Rich",
            BodyText = "Replaced content",
        });

        Assert.NotNull(outcome.BackupPath);
        Assert.Equal(htmBefore, File.ReadAllBytes(Path.Combine(outcome.BackupPath!, "Rich.htm")));
        Assert.Equal("old text", File.ReadAllText(Path.Combine(outcome.BackupPath!, "Rich.txt")));
        Assert.Equal(resourceBytes, File.ReadAllBytes(Path.Combine(outcome.BackupPath!, "Rich_files", "img.png")));

        Assert.Equal(new[] { "text" }, outcome.FormatsWritten);
        Assert.Equal("Replaced content", File.ReadAllText(Path.Combine(Dir, "Rich.txt")));
        Assert.Equal(htmBefore, File.ReadAllBytes(Path.Combine(Dir, "Rich.htm")));
        Assert.Equal(rtfBefore, File.ReadAllBytes(Path.Combine(Dir, "Rich.rtf")));
        Assert.Equal(resourceBytes, File.ReadAllBytes(Path.Combine(Dir, "Rich_files", "img.png")));
    }

    [Fact]
    public void Update_Replace_EditsEveryRenditionInPlace_KeepingEachOnesEncoding()
    {
        // As Word writes them: windows-1252 HTML with a charset meta and a line wrap inside the
        // text, RTF with \'xx escapes and a picture, a UTF-16 text file.
        Encoding w1252 = CodePage(1252);
        string html = "<html><head><meta http-equiv=Content-Type content=\"text/html; charset=windows-1252\"></head>"
            + "<body><p class=MsoNormal>Jan Modaal &#8211; Café de\r\nBuren</p><p>Tel 010 123</p></body></html>";
        _folder.Write("Word.htm", html, w1252);
        byte[] png = TestPictures.Png(2, 2);
        string rtf = @"{\rtf1\ansi\ansicpg1252{\fonttbl{\f0 Calibri;}}\f0 Jan Modaal - Caf\'e9 de Buren\par Tel 010 123\par "
            + @"{\*\shppict{\pict\pngblip " + TestPictures.Hex(png) + "}}}";
        _folder.Write("Word.rtf", rtf, Encoding.ASCII);
        _folder.Write("Word.txt", "Jan Modaal - Café de Buren\r\nTel 010 123\r\n", Encoding.Unicode);

        ManageSignatureOutcome outcome = Manage(new ManageSignatureRequest
        {
            Action = "update",
            Name = "Word",
            Replace = new[]
            {
                new SignatureTextEdit { Find = "Café de Buren", Replace = "Café De Gouden Leeuw" },
                new SignatureTextEdit { Find = "Tel 010 123", Replace = "Tel 010 999" },
            },
        });

        SignatureReplacementView cafe = outcome.Replacements![0];
        Assert.Equal(1, cafe.Html);
        Assert.Equal(1, cafe.Rtf);
        Assert.Equal(1, cafe.Text);
        Assert.Equal(new[] { "html", "rtf", "text" }, outcome.FormatsWritten);

        // HTML: still windows-1252, markup untouched, the text replaced where Word wrapped it.
        byte[] htmlAfter = File.ReadAllBytes(Path.Combine(Dir, "Word.htm"));
        string decoded = w1252.GetString(htmlAfter);
        Assert.Contains("Café De Gouden Leeuw", decoded, StringComparison.Ordinal);
        Assert.Contains("Tel 010 999", decoded, StringComparison.Ordinal);
        Assert.Contains("<p class=MsoNormal>Jan Modaal &#8211; ", decoded, StringComparison.Ordinal);
        Assert.Contains("charset=windows-1252", decoded, StringComparison.Ordinal);
        Assert.Contains((byte)0xE9, htmlAfter);

        // RTF: replaced in its own escaping, the picture byte for byte.
        string rtfAfter = File.ReadAllText(Path.Combine(Dir, "Word.rtf"));
        Assert.Contains(@"Caf\'e9 De Gouden Leeuw", rtfAfter, StringComparison.Ordinal);
        Assert.Contains(TestPictures.Hex(png), rtfAfter, StringComparison.Ordinal);

        // Text: still UTF-16 with its mark.
        byte[] textAfter = File.ReadAllBytes(Path.Combine(Dir, "Word.txt"));
        Assert.True(textAfter[0] == 0xFF && textAfter[1] == 0xFE);
        Assert.Contains("Café De Gouden Leeuw", Encoding.Unicode.GetString(textAfter, 2, textAfter.Length - 2), StringComparison.Ordinal);
    }

    [Fact]
    public void Update_Replace_SplitByFormatting_IsRefused_NothingWritten()
    {
        _folder.Write("Split.htm", "<html><body><p>Jan <b>Modaal</b></p></body></html>", Encoding.ASCII);
        _folder.Write("Split.txt", "Jan Modaal", Encoding.ASCII);

        ArgumentException refusal = Assert.Throws<ArgumentException>(() => Manage(new ManageSignatureRequest
        {
            Action = "update",
            Name = "Split",
            Replace = new[] { new SignatureTextEdit { Find = "Jan Modaal", Replace = "J. Modaal" } },
        }));

        Assert.Contains("formatting", refusal.Message, StringComparison.Ordinal);
        Assert.Contains("body_html", refusal.Message, StringComparison.Ordinal);
        Assert.Equal("Jan Modaal", File.ReadAllText(Path.Combine(Dir, "Split.txt")));
        Assert.False(Directory.Exists(_folder.BackupRoot), "a refusal comes before the backup");
    }

    [Fact]
    public void Update_Replace_NotFoundAnywhere_IsRefused()
    {
        SeedSignature("Plain", withResources: false);

        ArgumentException refusal = Assert.Throws<ArgumentException>(() => Manage(new ManageSignatureRequest
        {
            Action = "update",
            Name = "Plain",
            Replace = new[] { new SignatureTextEdit { Find = "Nowhere", Replace = "x" } },
        }));

        Assert.Contains("not found in any rendition", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Update_Pictures_AddReplaceRemove_KeepWordsFileListInStep()
    {
        _folder.Write("Logo.htm", "<html><body><p><img src=\"Logo_files/image001.png\"></p><p><img src=\"Logo_files/old.png\"></p></body></html>", Encoding.ASCII);
        byte[] oldLogo = TestPictures.Png(10, 10, 1);
        _folder.WriteBytes(Path.Combine("Logo_files", "image001.png"), oldLogo);
        _folder.WriteBytes(Path.Combine("Logo_files", "old.png"), TestPictures.Png(5, 5, 2));
        _folder.Write(Path.Combine("Logo_files", "filelist.xml"),
            "<xml xmlns:o=\"urn:schemas-microsoft-com:office:office\">\r\n <o:MainFile HRef=\"../Logo.htm\"/>\r\n"
            + " <o:File HRef=\"image001.png\"/>\r\n <o:File HRef=\"old.png\"/>\r\n <o:File HRef=\"filelist.xml\"/>\r\n</xml>", Encoding.ASCII);
        byte[] newLogo = TestPictures.Png(10, 10, 3);
        byte[] banner = TestPictures.Jpeg(300, 50);

        ManageSignatureOutcome outcome = Manage(new ManageSignatureRequest
        {
            Action = "update",
            Name = "Logo",
            BodyHtml = "<html><body><p><img src=\"Logo_files/image001.png\"></p><p><img src=\"Logo_files/banner.jpg\"></p></body></html>",
            Pictures = new[]
            {
                new SignaturePictureChange { File = "image001.png", Path = _folder.Outside("new.png", newLogo) },
                new SignaturePictureChange { File = "banner.jpg", Path = _folder.Outside("banner.jpg", banner) },
                new SignaturePictureChange { File = "old.png", Remove = true },
            },
        });

        Assert.Equal(new[] { "banner.jpg" }, outcome.PicturesAdded);
        Assert.Equal(new[] { "image001.png" }, outcome.PicturesReplaced);
        Assert.Equal(new[] { "old.png" }, outcome.PicturesRemoved);
        Assert.Equal(newLogo, File.ReadAllBytes(Path.Combine(Dir, "Logo_files", "image001.png")));
        Assert.Equal(banner, File.ReadAllBytes(Path.Combine(Dir, "Logo_files", "banner.jpg")));
        Assert.False(File.Exists(Path.Combine(Dir, "Logo_files", "old.png")));
        string fileList = File.ReadAllText(Path.Combine(Dir, "Logo_files", "filelist.xml"));
        Assert.Contains("banner.jpg", fileList, StringComparison.Ordinal);
        Assert.DoesNotContain("old.png", fileList, StringComparison.Ordinal);
        Assert.Contains("image001.png", fileList, StringComparison.Ordinal);
        Assert.Equal(oldLogo, File.ReadAllBytes(Path.Combine(outcome.BackupPath!, "Logo_files", "image001.png")));
    }

    [Fact]
    public void Update_APictureTheHtmlStillShows_CannotBeRemoved_AndAPictureMustBeWhatItsNameSays()
    {
        _folder.Write("Shown.htm", "<html><body><img src=\"Shown_files/logo.png\"></body></html>", Encoding.ASCII);
        _folder.WriteBytes(Path.Combine("Shown_files", "logo.png"), TestPictures.Png(1, 1));

        ArgumentException shown = Assert.Throws<ArgumentException>(() => Manage(new ManageSignatureRequest
        {
            Action = "update",
            Name = "Shown",
            Pictures = new[] { new SignaturePictureChange { File = "logo.png", Remove = true } },
        }));
        Assert.Contains("still shown", shown.Message, StringComparison.Ordinal);

        ArgumentException mislabelled = Assert.Throws<ArgumentException>(() => Manage(new ManageSignatureRequest
        {
            Action = "update",
            Name = "Shown",
            Pictures = new[] { new SignaturePictureChange { File = "logo.png", Path = _folder.Outside("photo.jpg", TestPictures.Jpeg(4, 4)) } },
        }));
        Assert.Contains("not a PNG picture", mislabelled.Message, StringComparison.Ordinal);
        Assert.True(File.Exists(Path.Combine(Dir, "Shown_files", "logo.png")));
    }

    [Fact]
    public void Update_BodyRtf_KeepsEachTokensPictureByteForByte_AndEmbedsANewOne()
    {
        byte[] png = TestPictures.Png(3, 3, 7);
        string pictureGroup = @"{\*\shppict{\pict\picwgoal450\pichgoal450\pngblip" + "\r\n" + TestPictures.Hex(png) + "}}"
            + "\r\n" + @"{\nonshppict{\pict\wmetafile8 0102}}";
        string original = @"{\rtf1\ansi Old line\par " + pictureGroup + "}";
        _folder.Write("Tokens.rtf", original, Encoding.ASCII);
        _folder.Write("Tokens.txt", "Old line", Encoding.ASCII);
        string elided = SignatureRtf.Elide(original);
        Assert.Equal(@"{\rtf1\ansi Old line\par [[picture:1]]}", elided);
        byte[] fresh = TestPictures.Png(8, 4, 9);

        Manage(new ManageSignatureRequest
        {
            Action = "update",
            Name = "Tokens",
            BodyRtf = elided.Replace("Old line", "New line") .Replace("[[picture:1]]", "[[picture:1]][[picture-file:" + _folder.Outside("fresh.png", fresh) + "|16x8]]"),
        });

        string after = File.ReadAllText(Path.Combine(Dir, "Tokens.rtf"));
        Assert.Contains("New line", after, StringComparison.Ordinal);
        Assert.Contains(pictureGroup, after, StringComparison.Ordinal);
        IReadOnlyList<RtfPicture> pictures = SignatureRtf.FindPictures(after);
        Assert.Equal(2, pictures.Count);
        Assert.Equal(SignatureRtf.Sha256Hex(fresh), pictures[1].Sha256);
        Assert.Equal(240, pictures[1].WidthTwips);

        ArgumentException unknown = Assert.Throws<ArgumentException>(() => Manage(new ManageSignatureRequest
        {
            Action = "update",
            Name = "Tokens",
            BodyRtf = @"{\rtf1 [[picture:9]]}",
        }));
        Assert.Contains("[[picture:9]]", unknown.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Delete_RemovesTheFileSet_AfterBackingItUp()
    {
        SeedSignature("Doomed", withResources: true, new byte[] { 9, 9 });

        ManageSignatureOutcome outcome = Manage(new ManageSignatureRequest { Action = "delete", Name = "Doomed" });

        Assert.NotNull(outcome.BackupPath);
        Assert.NotNull(outcome.FilesDeleted);
        Assert.Equal(3, outcome.FilesDeleted!.Count); // .htm + .txt + _files
        Assert.False(File.Exists(Path.Combine(Dir, "Doomed.htm")));
        Assert.False(File.Exists(Path.Combine(Dir, "Doomed.txt")));
        Assert.False(Directory.Exists(Path.Combine(Dir, "Doomed_files")));
        Assert.True(File.Exists(Path.Combine(outcome.BackupPath!, "Doomed.htm")));
        Assert.Equal(new byte[] { 9, 9 }, File.ReadAllBytes(Path.Combine(outcome.BackupPath!, "Doomed_files", "img.png")));
    }

    [Fact]
    public void FailingBackup_AbortsTheOperation_NothingModified()
    {
        SeedSignature("Protected", withResources: false);

        Directory.CreateDirectory(Path.GetDirectoryName(_folder.BackupRoot)!);
        File.WriteAllText(_folder.BackupRoot, "not a directory");

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(() =>
            Manage(new ManageSignatureRequest { Action = "delete", Name = "Protected" }));

        Assert.True(File.Exists(Path.Combine(Dir, "Protected.htm")), "a failed backup must leave the signature untouched");
        Assert.Equal(MutationOutcome.Unchanged, failure.Outcome);
        Assert.Contains("your signatures were not modified", failure.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("INCOMPLETE BACKUP", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ABackupThatFailedPartWay_NamesTheFolderItLeftBehind()
    {
        SeedSignature("Protected", withResources: false);

        string source = Path.Combine(Dir, "Protected.htm");
        using (File.Open(source, FileMode.Open, FileAccess.Read, FileShare.None))
        {
            OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(() =>
                Manage(new ManageSignatureRequest { Action = "delete", Name = "Protected" }));

            Assert.Equal(MutationOutcome.Unchanged, failure.Outcome);
            Assert.Contains("INCOMPLETE BACKUP", failure.Message, StringComparison.Ordinal);
            Assert.Contains(_folder.BackupRoot, failure.Message, StringComparison.OrdinalIgnoreCase);
        }

        Assert.True(File.Exists(source), "a failed backup must leave the signature untouched");
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

    // ------------------------------------------------------------------ defaults (Q147, Q155)

    [Fact]
    public void SetDefault_OnItsOwn_WritesTheRegistry_AndLeavesTheSignatureFilesAlone()
    {
        // The defect: a default could only be set by rewriting the signature.
        SeedSignature("Steered", withResources: true);
        Dictionary<string, byte[]> before = Snapshot("Steered");
        FakeDefaultsStore store = new FakeDefaultsStore().Account(Profile, "key-hub", "hub@example.com", newMessage: "Previous");

        ManageSignatureOutcome outcome = Manage(new ManageSignatureRequest
        {
            Action = "set_default",
            Name = "Steered",
            DefaultForProfile = "OUTLOOK",
            DefaultForAccount = "HUB@example.com",
            DefaultForScope = "both",
        }, store);

        Assert.Equal(Profile, outcome.Profile);
        Assert.Equal("hub@example.com", outcome.DefaultSetForAccount);
        Assert.Equal("both", outcome.DefaultSetScope);
        Assert.Null(outcome.BackupPath);
        Assert.Null(outcome.FilesWritten);
        Assert.Contains(("key-hub", SignatureManager.NewSignatureValueName, "Steered"), store.Writes);
        Assert.Contains(("key-hub", SignatureManager.ReplyForwardSignatureValueName, "Steered"), store.Writes);
        SignatureDefaultChange newChange = Assert.Single(outcome.DefaultsChanged!, c => c.Scope == "new");
        Assert.Equal("Previous", newChange.Previous);
        Assert.Equal("Steered", newChange.Now);
        Assert.Contains("next start", outcome.Advice, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(before, Snapshot("Steered"));
    }

    [Theory]
    [InlineData("new", SignatureManager.NewSignatureValueName)]
    [InlineData("reply", SignatureManager.ReplyForwardSignatureValueName)]
    public void SetDefault_WithACreate_SingleScope_WritesOnlyThatValue(string scope, string expectedValueName)
    {
        FakeDefaultsStore store = new FakeDefaultsStore().Account(Profile, "k", "a@b.nl");

        Manage(new ManageSignatureRequest
        {
            Action = "create",
            Name = "One" + scope,
            BodyText = "x",
            DefaultForProfile = Profile,
            DefaultForAccount = "a@b.nl",
            DefaultForScope = scope,
        }, store);

        (string, string, string) write = Assert.Single(store.Writes);
        Assert.Equal(expectedValueName, write.Item2);
    }

    [Fact]
    public void SetDefault_NeedsAnExplicitProfile_ThatExists_AndAnAccountOfThatProfile()
    {
        SeedSignature("S", withResources: false);
        FakeDefaultsStore store = new FakeDefaultsStore()
            .Account("Home", "k1", "jan@example.com")
            .Account("Work", "k2", "jan@example.com");

        ArgumentException noProfile = Assert.Throws<ArgumentException>(() => Manage(SetDefault("S", null, "jan@example.com"), store));
        Assert.Contains("profile is required", noProfile.Message, StringComparison.Ordinal);
        Assert.Contains("'Home', 'Work'", noProfile.Message, StringComparison.Ordinal);

        ArgumentException unknownProfile = Assert.Throws<ArgumentException>(() => Manage(SetDefault("S", "Elsewhere", "jan@example.com"), store));
        Assert.Contains("does not exist", unknownProfile.Message, StringComparison.Ordinal);

        ArgumentException otherAccount = Assert.Throws<ArgumentException>(() => Manage(SetDefault("S", "Home", "piet@example.com"), store));
        Assert.Contains("not a mail account of Outlook profile 'Home'", otherAccount.Message, StringComparison.Ordinal);
        Assert.Empty(store.Writes);

        // The same address in two profiles: only the named profile's entry is written - never merged.
        Manage(SetDefault("S", "Work", "jan@example.com"), store);
        Assert.Equal(new[] { ("k2", SignatureManager.NewSignatureValueName, "S") }, store.Writes);
    }

    [Fact]
    public void SetDefault_UnknownAccount_IsRejectedBeforeAnyFileWork()
    {
        Assert.Throws<ArgumentException>(() => Manage(new ManageSignatureRequest
        {
            Action = "create",
            Name = "Never",
            BodyText = "x",
            DefaultForProfile = Profile,
            DefaultForAccount = "stranger@nowhere.test",
            DefaultForScope = "new",
        }, new FakeDefaultsStore().Account(Profile, "k", "a@b.nl")));

        Assert.False(File.Exists(Path.Combine(Dir, "Never.htm")), "validation must reject before writing files");
    }

    [Fact]
    public void ClearDefault_ClearsOnlyTheDefaultItNames()
    {
        FakeDefaultsStore store = new FakeDefaultsStore().Account(Profile, "k", "jan@example.com", newMessage: "Current", replyForward: "Other");

        ArgumentException wrong = Assert.Throws<ArgumentException>(() => Manage(new ManageSignatureRequest
        {
            Action = "clear_default", Name = "Other", DefaultForProfile = Profile, DefaultForAccount = "jan@example.com", DefaultForScope = "new",
        }, store));
        Assert.Contains("is 'Current', not 'Other'", wrong.Message, StringComparison.Ordinal);

        ManageSignatureOutcome cleared = Manage(new ManageSignatureRequest
        {
            Action = "clear_default", Name = "current", DefaultForProfile = Profile, DefaultForAccount = "jan@example.com", DefaultForScope = "new",
        }, store);

        Assert.Equal(new[] { ("k", SignatureManager.NewSignatureValueName) }, store.Clears);
        SignatureDefaultChange change = Assert.Single(cleared.DefaultsChanged!);
        Assert.Equal("Current", change.Previous);
        Assert.Null(change.Now);
        Assert.Null(cleared.DefaultSetForAccount);
    }

    [Fact]
    public void Delete_ClearsDanglingDefaults_InEveryProfile()
    {
        SeedSignature("Assigned", withResources: false);
        FakeDefaultsStore store = new FakeDefaultsStore()
            .Account("Home", "k1", "one@example.com", newMessage: "Assigned")
            .Account("Work", "k2", "two@example.com", newMessage: "Other", replyForward: "assigned")
            .Account("Work", "k3", "three@example.com", newMessage: "Other");

        ManageSignatureOutcome outcome = Manage(new ManageSignatureRequest { Action = "delete", Name = "Assigned" }, store);

        Assert.Equal(
            new[] { ("Home", "one@example.com", "new"), ("Work", "two@example.com", "reply") },
            outcome.DefaultsCleared!.Select(c => (c.Profile, c.Account, c.Scope)));
        Assert.DoesNotContain(store.Clears, c => c.Item1 == "k3");
        Assert.Contains("next start", outcome.Advice, StringComparison.OrdinalIgnoreCase);
    }

    // ------------------------------------------------------------------ policy (Q156)

    [Fact]
    public void Policy_SignaturesDisabled_RefusesEveryWrite()
    {
        SeedSignature("Any", withResources: false);
        FakeSignatureSettings disabled = new FakeSignatureSettings().Policy(SignaturePolicySource.GroupPolicy, disable: true);

        foreach (ManageSignatureRequest request in new[]
        {
            new ManageSignatureRequest { Action = "create", Name = "New", BodyText = "x" },
            new ManageSignatureRequest { Action = "update", Name = "Any", BodyText = "x" },
            new ManageSignatureRequest { Action = "delete", Name = "Any" },
        })
        {
            SignatureRefusedException refusal = Assert.Throws<SignatureRefusedException>(() => Manage(request, settings: disabled));
            Assert.Equal("signatures_disabled_by_policy", refusal.Reason);
        }

        Assert.True(File.Exists(Path.Combine(Dir, "Any.htm")));
        Assert.False(File.Exists(Path.Combine(Dir, "New.htm")));
    }

    [Theory]
    [InlineData(SignaturePolicySource.GroupPolicy, "cannot be changed")]
    [InlineData(SignaturePolicySource.Preference, "would do nothing")]
    public void Policy_ADefaultSetForEveryAccount_RefusesADefaultChange(SignaturePolicySource source, string why)
    {
        SeedSignature("S", withResources: false);
        FakeDefaultsStore store = new FakeDefaultsStore().Account(Profile, "k", "jan@example.com");

        SignatureRefusedException refusal = Assert.Throws<SignatureRefusedException>(() => Manage(
            SetDefault("S", Profile, "jan@example.com"), store, settings: new FakeSignatureSettings().Policy(source, newMessage: "Corporate")));

        Assert.Equal("default_set_by_policy", refusal.Reason);
        Assert.Contains(why, refusal.Message, StringComparison.Ordinal);
        Assert.Empty(store.Writes);
    }

    [Fact]
    public void Policy_TheForcedSignature_CannotBeDeleted()
    {
        SeedSignature("Corporate", withResources: false);

        SignatureRefusedException refusal = Assert.Throws<SignatureRefusedException>(() => Manage(
            new ManageSignatureRequest { Action = "delete", Name = "corporate" },
            settings: new FakeSignatureSettings().Policy(SignaturePolicySource.GroupPolicy, reply: "Corporate")));

        Assert.Equal("signature_forced_by_policy", refusal.Reason);
        Assert.True(File.Exists(Path.Combine(Dir, "Corporate.htm")));
    }

    // ------------------------------------------------------------------ roaming (Q152)

    [Fact]
    public void Roaming_ALocalCopyOfAMailboxSignature_IsNeverEdited_WhileRoamingIsPossible()
    {
        SeedSignature("Work (jan@example.com)", withResources: false);
        FakeSignatureSettings roams = new FakeSignatureSettings().Microsoft365();

        foreach (ManageSignatureRequest request in new[]
        {
            new ManageSignatureRequest { Action = "update", Name = "Work (jan@example.com)", BodyText = "x" },
            new ManageSignatureRequest { Action = "delete", Name = "Work (jan@example.com)" },
            new ManageSignatureRequest { Action = "create", Name = "Other (jan@example.com)", BodyText = "x" },
        })
        {
            SignatureRefusedException refusal = Assert.Throws<SignatureRefusedException>(() => Manage(request, settings: roams));
            Assert.Equal("managed_in_mailbox", refusal.Reason);
        }

        // With roaming switched off the same file is just a local signature.
        Manage(new ManageSignatureRequest { Action = "update", Name = "Work (jan@example.com)", BodyText = "edited" },
            settings: new FakeSignatureSettings().Microsoft365().RoamingSwitchedOff());
        Assert.Equal("edited", File.ReadAllText(Path.Combine(Dir, "Work (jan@example.com).txt")));
    }

    [Fact]
    public void Roaming_AnExchangeAccountsDefaults_AreLeftToOutlook_WhileRoamingIsPossible()
    {
        SeedSignature("S", withResources: false);
        FakeDefaultsStore store = new FakeDefaultsStore()
            .Account(Profile, "k1", "jan@example.com", kind: MailAccountKind.Exchange)
            .Account(Profile, "k2", "pop@example.com", kind: MailAccountKind.Pop3);
        FakeSignatureSettings ltsc2024 = new FakeSignatureSettings().Ltsc2024();

        SignatureRefusedException refusal = Assert.Throws<SignatureRefusedException>(
            () => Manage(SetDefault("S", Profile, "jan@example.com"), store, settings: ltsc2024));
        Assert.Equal("account_may_roam", refusal.Reason);

        Manage(SetDefault("S", Profile, "pop@example.com"), store, settings: ltsc2024);
        Assert.Equal(new[] { ("k2", SignatureManager.NewSignatureValueName, "S") }, store.Writes);
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
    [InlineData("create", "X", "b", null, "a@b.nl", null)]            // account without scope or profile
    [InlineData("create", "X", "b", null, null, "new")]               // scope without account
    [InlineData("create", "X", "b", null, "a@b.nl", "everything")]    // bad scope
    [InlineData("delete", "X", null, null, "a@b.nl", "new")]          // a default with delete
    [InlineData("set_default", "X", "b", null, "a@b.nl", "new")]      // content with set_default
    public void InvalidRequests_AreRejected_WithArgumentException(
        string action, string name, string? bodyText, string? bodyHtml, string? account, string? scope)
    {
        Assert.ThrowsAny<ArgumentException>(() => Manage(new ManageSignatureRequest
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
    public void Update_WithOnlyADefault_PointsAtSetDefault()
    {
        SeedSignature("S", withResources: false);

        ArgumentException ex = Assert.Throws<ArgumentException>(() => Manage(new ManageSignatureRequest
        {
            Action = "update", Name = "S", DefaultForProfile = Profile, DefaultForAccount = "a@b.nl", DefaultForScope = "new",
        }));
        Assert.Contains("set_default", ex.Message, StringComparison.Ordinal);
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
        string rtf = SignatureManager.BuildRtfFromText("a\\b {c} é 中\r\nsecond");
        Assert.Contains(@"a\\b", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\{c\}", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\'e9", rtf, StringComparison.Ordinal);
        Assert.Contains(@" 3?", rtf, StringComparison.Ordinal);
        Assert.Contains(@"\par second", rtf, StringComparison.Ordinal);
        Assert.EndsWith("}", rtf, StringComparison.Ordinal);
    }

    [Fact]
    public void ListedAfterCreate_GoneAfterDelete_ViaTheCatalog()
    {
        Manage(new ManageSignatureRequest { Action = "create", Name = "RoundTrip", BodyText = "Hello / World" });
        SignatureInfo? listed = SignatureCatalog.TryResolve("RoundTrip", Dir);
        Assert.NotNull(listed);
        Assert.NotNull(listed!.HtmlPath);
        Assert.NotNull(listed.TextPath);
        Assert.NotNull(listed.RtfPath);
        Assert.Contains("Hello", listed.Excerpt, StringComparison.Ordinal);

        Manage(new ManageSignatureRequest { Action = "delete", Name = "RoundTrip" });
        Assert.Null(SignatureCatalog.TryResolve("RoundTrip", Dir));
    }

    // ------------------------------------------------------------------ helpers

    private static ManageSignatureRequest SetDefault(string name, string? profile, string account, string scope = "new")
    {
        return new ManageSignatureRequest
        {
            Action = "set_default", Name = name, DefaultForProfile = profile, DefaultForAccount = account, DefaultForScope = scope,
        };
    }

    private void SeedSignature(string name, bool withResources, byte[]? resourceBytes = null)
    {
        File.WriteAllText(Path.Combine(Dir, name + ".htm"), "<html><body><p>old html of " + name + "</p></body></html>");
        File.WriteAllText(Path.Combine(Dir, name + ".txt"), "old text");
        if (withResources)
        {
            string resources = Path.Combine(Dir, name + "_files");
            Directory.CreateDirectory(resources);
            File.WriteAllBytes(Path.Combine(resources, "img.png"), resourceBytes ?? new byte[] { 0 });
        }
    }

    private Dictionary<string, byte[]> Snapshot(string name)
    {
        return Directory.GetFiles(Dir, "*", SearchOption.AllDirectories)
            .Where(f => Path.GetFileName(f).StartsWith(name, StringComparison.Ordinal) || f.Contains(name + "_files", StringComparison.Ordinal))
            .ToDictionary(f => f, File.ReadAllBytes);
    }

    private static Encoding CodePage(int codePage)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding(codePage);
    }

    /// <summary>
    /// Row-level fake: every row IS a mail account by construction, so this exercises Manage's
    /// orchestration only. Which registry entries count as accounts - and the 2026-09-24 data-file
    /// defect - is SignatureDefaultTargetTests, over a registry fake. Stateful, because Manage reads
    /// a default back before reporting it set.
    /// </summary>
    private sealed class FakeDefaultsStore : ISignatureDefaultsStore
    {
        public List<SignatureDefaultsRow> Rows { get; } = new List<SignatureDefaultsRow>();

        public List<(string, string, string)> Writes { get; } = new List<(string, string, string)>();

        public List<(string, string)> Clears { get; } = new List<(string, string)>();

        public FakeDefaultsStore Account(
            string profile, string key, string address, string? newMessage = null, string? replyForward = null, MailAccountKind kind = MailAccountKind.Pop3)
        {
            Rows.Add(new SignatureDefaultsRow(key, address, newMessage, replyForward, profile, kind));
            return this;
        }

        public IReadOnlyList<SignatureDefaultsRow> ReadAccounts()
        {
            // A snapshot: the delete sweep clears while it iterates.
            return Rows.ToList();
        }

        public ProfileList ReadProfiles()
        {
            return new ProfileList(null, Rows.Select(r => r.Profile!).Distinct(StringComparer.OrdinalIgnoreCase).ToList());
        }

        public ProfileAccountEntry? ReadEntry(string accountKey)
        {
            SignatureDefaultsRow? row = Rows.FirstOrDefault(r => r.AccountKey == accountKey);
            return row == null
                ? null
                : new ProfileAccountEntry(row.AccountKey, ProfileEntryKind.MailAccount, row.Account, row.Account, row.NewMessage,
                    row.ReplyForward, row.Profile, row.AccountKind);
        }

        public void WriteDefault(string accountKey, string valueName, string signatureName)
        {
            Writes.Add((accountKey, valueName, signatureName));
            Replace(accountKey, valueName, signatureName);
        }

        public void ClearDefault(string accountKey, string valueName)
        {
            Clears.Add((accountKey, valueName));
            Replace(accountKey, valueName, null);
        }

        private void Replace(string accountKey, string valueName, string? value)
        {
            int index = Rows.FindIndex(r => r.AccountKey == accountKey);
            if (index < 0)
            {
                return;
            }

            SignatureDefaultsRow row = Rows[index];
            Rows[index] = valueName == SignatureManager.NewSignatureValueName
                ? new SignatureDefaultsRow(row.AccountKey, row.Account, value, row.ReplyForward, row.Profile, row.AccountKind)
                : new SignatureDefaultsRow(row.AccountKey, row.Account, row.NewMessage, value, row.Profile, row.AccountKind);
        }
    }
}
