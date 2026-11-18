using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 pins for the PRE-COM attachment gate (v3.MD D46/C3). The user granted NO path
/// restrictions, so these tests deliberately do NOT assert on where a file lives - they
/// assert that a file is genuinely ATTACHABLE (rooted, present, a file, readable,
/// non-empty) and that a bad set is refused WHOLE, with every offending path named.
/// Temp files live under the OS temp dir following the suite convention.
/// </summary>
public sealed class DraftAttachmentsTests : IDisposable
{
    private readonly string _directory;

    public DraftAttachmentsTests()
    {
        _directory = Path.Combine(Path.GetTempPath(), "OutlookAI-McpTest-t1attach-" + Guid.NewGuid().ToString("N")[..12]);
        Directory.CreateDirectory(_directory);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_directory, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    private string WriteFile(string name, string content = "attachment payload")
    {
        string path = Path.Combine(_directory, name);
        File.WriteAllText(path, content);
        return path;
    }

    [Fact]
    public void NullOrEmpty_YieldsNoFiles_BecauseTheParameterIsOptionalEverywhere()
    {
        Assert.Empty(DraftAttachments.Validate(null));
        Assert.Empty(DraftAttachments.Validate(Array.Empty<string>()));
    }

    [Fact]
    public void ValidFiles_AreAcceptedWithTheirRealNameAndSize()
    {
        string a = WriteFile("offer.pdf", "12345");
        string b = WriteFile("terms.txt", "1234567890");

        IReadOnlyList<DraftAttachmentFile> files = DraftAttachments.Validate(new[] { a, b });

        Assert.Equal(2, files.Count);
        Assert.Equal("offer.pdf", files[0].FileName);
        Assert.Equal(5, files[0].SizeBytes);
        Assert.Equal("terms.txt", files[1].FileName);
        Assert.Equal(10, files[1].SizeBytes);
        Assert.Equal(15, DraftAttachments.TotalBytes(files));
    }

    [Fact]
    public void AnyPathOutsideAFolderIsFine_TheGrantHasNoDirectoryRestriction()
    {
        // Explicit pin of the user's grant: acceptance depends on the FILE, never on
        // where it sits. A file directly under the temp root is as acceptable as one in
        // a nested folder.
        string nested = Path.Combine(_directory, "sub", "deep");
        Directory.CreateDirectory(nested);
        string deepFile = Path.Combine(nested, "deep.bin");
        File.WriteAllBytes(deepFile, new byte[] { 1, 2, 3 });

        IReadOnlyList<DraftAttachmentFile> files = DraftAttachments.Validate(new[] { WriteFile("flat.txt"), deepFile });

        Assert.Equal(2, files.Count);
    }

    [Theory]
    [InlineData("documents\\offer.pdf")]
    [InlineData("offer.pdf")]
    [InlineData("./offer.pdf")]
    public void RelativePaths_AreRejected_BecauseTheServerHasNoDependableWorkingDirectory(string relative)
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() => DraftAttachments.Validate(new[] { relative }));

        Assert.Contains(relative, ex.Message, StringComparison.Ordinal);
        Assert.Contains("absolute path", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void MissingFile_IsRejected_AndNamed()
    {
        string missing = Path.Combine(_directory, "not-there.pdf");

        ArgumentException ex = Assert.Throws<ArgumentException>(() => DraftAttachments.Validate(new[] { missing }));

        Assert.Contains(missing, ex.Message, StringComparison.Ordinal);
        Assert.Contains("no such file", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Directory_IsRejectedWithADirectorySpecificReason()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() => DraftAttachments.Validate(new[] { _directory }));

        Assert.Contains("is a directory", ex.Message, StringComparison.Ordinal);
        Assert.Contains("name the individual files", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void EmptyFile_IsRejected_BecauseOutlookDropsZeroByteAttachments()
    {
        string empty = Path.Combine(_directory, "empty.txt");
        File.WriteAllText(empty, string.Empty);

        ArgumentException ex = Assert.Throws<ArgumentException>(() => DraftAttachments.Validate(new[] { empty }));

        Assert.Contains("is empty", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void BlankEntry_IsRejected_ByPosition()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(
            () => DraftAttachments.Validate(new[] { WriteFile("ok.txt"), "   " }));

        Assert.Contains("entry 2", ex.Message, StringComparison.Ordinal);
        Assert.Contains("blank path", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DuplicateFile_IsRejected_SoTheSameFileCannotBeAttachedTwiceByAccident()
    {
        string a = WriteFile("dup.txt");

        ArgumentException ex = Assert.Throws<ArgumentException>(() => DraftAttachments.Validate(new[] { a, a }));

        Assert.Contains("listed more than once", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void EveryBadPathIsNamedInOneMessage_NotJustTheFirst()
    {
        // THE fail-closed contract: one retry must be able to fix everything, and a
        // partially attached draft is never produced.
        string good = WriteFile("good.txt");
        string missingA = Path.Combine(_directory, "gone-a.pdf");
        string missingB = Path.Combine(_directory, "gone-b.pdf");

        ArgumentException ex = Assert.Throws<ArgumentException>(
            () => DraftAttachments.Validate(new[] { good, missingA, missingB, "relative.txt" }));

        Assert.Contains(missingA, ex.Message, StringComparison.Ordinal);
        Assert.Contains(missingB, ex.Message, StringComparison.Ordinal);
        Assert.Contains("relative.txt", ex.Message, StringComparison.Ordinal);
        Assert.Contains("3 unusable entries", ex.Message, StringComparison.Ordinal);
        Assert.Contains("no draft was changed", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TooManyFiles_AreRejectedBeforeAnyDiskWork()
    {
        string[] paths = Enumerable.Range(0, DraftAttachments.MaxFiles + 1)
            .Select(i => Path.Combine(_directory, "f" + i + ".txt"))
            .ToArray();

        ArgumentException ex = Assert.Throws<ArgumentException>(() => DraftAttachments.Validate(paths));

        Assert.Contains("at most 20 files", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void FormatBytes_ReadsAsHumanSizes()
    {
        Assert.Equal("512 B", DraftAttachments.FormatBytes(512));
        Assert.Equal("1.0 KB", DraftAttachments.FormatBytes(1024));
        Assert.Equal("150.0 MB", DraftAttachments.FormatBytes(DraftAttachments.MaxTotalBytes));
    }

    [Fact]
    public void RemoveNames_AreTrimmedAndDeduplicated()
    {
        IReadOnlyList<string> names = DraftAttachments.ValidateRemoveNames(new[] { " a.pdf ", "A.PDF", "b.txt" });

        Assert.Equal(new[] { "a.pdf", "b.txt" }, names);
    }

    [Fact]
    public void RemoveNames_RejectBlankEntries()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(
            () => DraftAttachments.ValidateRemoveNames(new[] { "a.pdf", " " }));

        Assert.Contains("blank name", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RemoveNames_NullIsEmpty()
    {
        Assert.Empty(DraftAttachments.ValidateRemoveNames(null));
    }

    // ------------------------------------------------------------------ the paths Outlook cannot be handed
    //
    // Measured on the unindexed test guest 2026-11-18 (runbook section 4.10): a \\?\ path crashed
    // OUTLOOK.EXE inside Attachments.Add four times out of four, whatever the file was called; Outlook
    // refused a file called CON.txt, an alternate data stream and a 304-character path; it attached the
    // right file from \\.\C:\... and \??\C:\... . Every spelling below is judged on its text, so most
    // need no file at all.

    [Theory]
    [InlineData(@"\\?\C:\Users\me\Documents\offer.pdf", @"C:\Users\me\Documents\offer.pdf")]
    [InlineData("//?/C:/Users/me/offer.pdf", "C:/Users/me/offer.pdf")]
    [InlineData(@"\\?/C:\offer.pdf", @"C:\offer.pdf")]
    [InlineData(@"\\?\UNC\server\share\offer.pdf", @"\\server\share\offer.pdf")]
    [InlineData(@"\\?\unc\server\share\offer.pdf", @"\\server\share\offer.pdf")]
    [InlineData(@"\\?\C:\x\CON.txt", @"C:\x\CON.txt")]
    [InlineData(@"\\?\C:\x\offer.pdf. . .", @"C:\x\offer.pdf. . .")]
    public void AQuestionMarkPath_IsRefused_BecauseOutlookCrashesOnIt_AndTheRefusalGivesItsOrdinaryPath(string asked, string ordinary)
    {
        string? reason = DraftAttachments.RefuseOnItsText(asked);

        Assert.NotNull(reason);
        Assert.Contains(@"is a \\?\ path, and Outlook crashes", reason, StringComparison.Ordinal);
        Assert.EndsWith("give the file's ordinary path instead, '" + ordinary + "'", reason, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(@"\\?\Volume{12345678-1234-1234-1234-123456789012}\offer.pdf")]
    [InlineData(@"\\?\GLOBALROOT\Device\HarddiskVolume3\offer.pdf")]
    public void AQuestionMarkPathWithNoDriveOrShare_IsRefused_WithAnExampleOfAnOrdinaryPath(string asked)
    {
        string? reason = DraftAttachments.RefuseOnItsText(asked);

        Assert.NotNull(reason);
        Assert.Contains("Outlook crashes", reason, StringComparison.Ordinal);
        Assert.EndsWith(@"such as 'C:\folder\file.pdf'", reason, StringComparison.Ordinal);
    }

    [Fact]
    public void AQuestionMarkPath_ToARealFile_IsRefusedWhole_BeforeOutlookCouldBeAsked()
    {
        // The file is there and readable: until 2026-11-18 this path passed every check and reached
        // Attachments.Add, which crashed Outlook.
        string file = WriteFile("plain.txt");
        string asked = @"\\?\" + file;

        ArgumentException ex = Assert.Throws<ArgumentException>(() => DraftAttachments.Validate(new[] { WriteFile("ok.txt"), asked }));

        Assert.Contains("'" + asked + "': is a \\\\?\\ path, and Outlook crashes", ex.Message, StringComparison.Ordinal);
        Assert.Contains("'" + file + "'", ex.Message, StringComparison.Ordinal);
        Assert.Contains("1 unusable entry", ex.Message, StringComparison.Ordinal);
        Assert.Contains("no draft was changed", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(@"\\.\")]
    [InlineData(@"\??\")]
    [InlineData("//./")]
    public void ADotOrNtPathToAFileOnADrive_IsAccepted_AsGiven_BecauseOutlookAttachedTheRightFileFromIt(string prefix)
    {
        string file = WriteFile("plain.txt", "12345");

        DraftAttachmentFile accepted = Assert.Single(DraftAttachments.Validate(new[] { prefix + file }));

        Assert.Equal("plain.txt", accepted.FileName);
        Assert.Equal(5, accepted.SizeBytes);
        Assert.Equal(file, accepted.Path.Substring(4), ignoreCase: true);
        Assert.Null(DraftAttachments.RefuseOnItsText(prefix + file));
    }

    [Theory]
    [InlineData(@"\\.\")]
    [InlineData(@"\??\")]
    public void TheSameFile_ByItsOrdinaryPathAndThroughTheDeviceNamespace_IsListedTwice(string prefix)
    {
        string file = WriteFile("twice.txt");

        ArgumentException ex = Assert.Throws<ArgumentException>(() => DraftAttachments.Validate(new[] { file, prefix + file }));

        Assert.Contains("'" + prefix + file + "': the same file is listed more than once", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(@"\\.\pipe\outlookai")]
    [InlineData(@"\\.\COM1")]
    [InlineData(@"\\.\PhysicalDrive0")]
    [InlineData(@"\\.\GLOBALROOT\Device\HarddiskVolume3\offer.pdf")]
    [InlineData(@"\\.\UNC\server\share\offer.pdf")]
    [InlineData(@"\??\UNC\server\share\offer.pdf")]
    [InlineData(@"\??\Volume{12345678-1234-1234-1234-123456789012}\offer.pdf")]
    [InlineData(@"\\.\C:")]
    [InlineData("//./pipe/outlookai")]
    public void ADevicePathThatNamesNoFileOnADrive_IsRefusedOnItsText_SoNothingOpensIt(string asked)
    {
        string? reason = DraftAttachments.RefuseOnItsText(asked);

        Assert.NotNull(reason);
        Assert.Contains("names no file on a drive", reason, StringComparison.Ordinal);
        Assert.Contains(@"'\\server\share\file.pdf'", reason, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("offer.pdf.")]
    [InlineData("offer.pdf ")]
    [InlineData("offer.pdf. . .")]
    public void AnNtPathToANameEndingInADotOrASpace_IsRefused(string name)
    {
        string? reason = DraftAttachments.RefuseOnItsText(@"\??\C:\x\" + name);

        Assert.NotNull(reason);
        Assert.Contains("ends in a dot or a space", reason, StringComparison.Ordinal);
    }

    [Fact]
    public void ADotPathToANameEndingInADot_IsNormalizedLikeAnOrdinaryPath_ToTheFileWithoutIt()
    {
        // \\.\ is normalized as an ordinary path is: Windows drops the dots, and the file without them
        // is the one attached - from \\.\ as from C:\.
        string file = WriteFile("normalized.txt");

        DraftAttachmentFile accepted = Assert.Single(DraftAttachments.Validate(new[] { @"\\.\" + file + ". ." }));

        Assert.Equal("normalized.txt", accepted.FileName);
    }

    [Theory]
    [InlineData("CON")]
    [InlineData("CON.txt")]
    [InlineData("con.TXT")]
    [InlineData("NUL")]
    [InlineData("nul.tar.gz")]
    [InlineData("PRN .pdf")]
    [InlineData("AUX.doc")]
    [InlineData("COM1.log")]
    [InlineData("COM\u00B9.txt")]
    [InlineData("LPT9")]
    [InlineData("CONIN$")]
    [InlineData("conout$.txt")]
    public void AFileCalledAfterAWindowsDevice_IsRefusedOnItsName_ThroughEveryKindOfPath(string name)
    {
        foreach (string asked in new[]
        {
            Path.Combine(_directory, name), @"\\.\" + Path.Combine(_directory, name), @"\??\" + Path.Combine(_directory, name),
            @"\\server\share\" + name, "C:/x/" + name,
        })
        {
            string? reason = DraftAttachments.RefuseOnItsText(asked);

            Assert.NotNull(reason);
            Assert.Contains("is called after a Windows device ('" + name + "')", reason, StringComparison.Ordinal);
            Assert.Contains("such as '_" + name + "'", reason, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void AFileCalledConTxt_ThatReallyExists_IsRefused_NotAttachedAndNotReportedMissing()
    {
        // Windows 11 keeps C:\x\CON.txt as an ordinary file; Windows 10 reads it as the console. Outlook
        // refused it on the guest, and crashed on it from a \\?\ path - so it is refused, on every Windows.
        string con = Path.Combine(_directory, "CON.txt");
        File.WriteAllText(@"\\?\" + con, "a file called CON.txt");
        try
        {
            foreach (string asked in new[] { con, @"\\.\" + con, @"\??\" + con, @"\\?\" + con })
            {
                ArgumentException ex = Assert.Throws<ArgumentException>(() => DraftAttachments.Validate(new[] { asked }));

                Assert.Contains("'" + asked + "': is ", ex.Message, StringComparison.Ordinal);
                Assert.DoesNotContain("no such file", ex.Message, StringComparison.Ordinal);
            }
        }
        finally
        {
            ScratchFiles.TryDeleteFile(@"\\?\" + con);
        }
    }

    [Fact]
    public void AnAlternateDataStream_IsRefused_EvenWhenItExists_ThroughEveryKindOfPath()
    {
        string file = WriteFile("host.txt");
        string stream = file + ":hidden";
        File.WriteAllText(stream, "the stream's own content");

        foreach (string asked in new[] { stream, file + "::$DATA", @"\??\" + stream, @"\\.\" + stream, @"\\server\share\host.txt:hidden" })
        {
            string? reason = DraftAttachments.RefuseOnItsText(asked);

            Assert.NotNull(reason);
            Assert.Contains("alternate data stream", reason, StringComparison.Ordinal);
        }

        ArgumentException ex = Assert.Throws<ArgumentException>(() => DraftAttachments.Validate(new[] { stream }));
        Assert.Contains("'" + stream + "': holds a ':' that is not a drive letter's", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ADriveRelativePath_IsRefused_ForItsColon()
    {
        string? reason = DraftAttachments.RefuseOnItsText("C:offer.pdf");

        Assert.NotNull(reason);
        Assert.Contains("relative to a drive's current folder", reason, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(@"NUL\.")]
    [InlineData(@"nul\x\..")]
    [InlineData(@"CON.txt\.")]
    public void APathWhoseTextHidesADevice_IsRefusedOnTheFullPathWindowsMakesOfIt(string below)
    {
        // The last name on the text is '.' or '..', so the text check passes; the full path Windows makes
        // of it ends in a device name, or is the device itself (C:\x\NUL is \\.\NUL on Windows 11).
        string asked = Path.Combine(_directory, below);
        Assert.Null(DraftAttachments.RefuseOnItsText(asked));

        ArgumentException ex = Assert.Throws<ArgumentException>(() => DraftAttachments.Validate(new[] { asked }));

        Assert.Contains("'" + asked + "': ", ex.Message, StringComparison.Ordinal);
        Assert.True(
            ex.Message.Contains("is called after a Windows device", StringComparison.Ordinal)
                || ex.Message.Contains("names no file on a drive", StringComparison.Ordinal),
            ex.Message);
    }

    [Fact]
    public void APathOverWindowsOrdinaryLimit_IsRefused_NamingItsLength_AndOneAtTheLimitIsAccepted()
    {
        string atTheLimit = WriteFileWithPathLength(WindowsFileNames.MaxPathChars, "at");
        string over = WriteFileWithPathLength(WindowsFileNames.MaxPathChars + 1, "over");

        DraftAttachmentFile accepted = Assert.Single(DraftAttachments.Validate(new[] { atTheLimit }));
        Assert.Equal(WindowsFileNames.MaxPathChars, accepted.Path.Length);

        ArgumentException ex = Assert.Throws<ArgumentException>(() => DraftAttachments.Validate(new[] { over }));
        Assert.Contains(
            "'" + over + "': is " + (WindowsFileNames.MaxPathChars + 1) + " characters long, and Outlook attaches a file only from a path "
            + "of at most " + WindowsFileNames.MaxPathChars,
            ex.Message,
            StringComparison.Ordinal);

        // Through the device namespace the prefix counts too: Outlook is handed the whole string.
        Assert.Throws<ArgumentException>(() => DraftAttachments.Validate(new[] { @"\\.\" + atTheLimit }));
    }

    [Theory]
    [InlineData(@"C:\Users\me\Documents\offer.pdf")]
    [InlineData("C:/Users/me/offer.pdf")]
    [InlineData(@"\\server\share\offer.pdf")]
    [InlineData(@"\\server\share\CONSOLE.txt")]
    [InlineData(@"C:\x\con-notes.txt")]
    [InlineData(@"C:\CON\offer.pdf")]
    [InlineData(@"C:\x\offer.pdf.")]
    [InlineData(@"C:\x\.hidden")]
    public void AnOrdinaryPath_PassesTheTextChecks_WhereverItIs(string asked)
    {
        // The grant stands (D46): no folder is refused, a share included - only the spellings above.
        Assert.Null(DraftAttachments.RefuseOnItsText(asked));
    }

    [Fact]
    public void EveryPathOutlookCannotBeHanded_IsNamedWithItsOwnReason_InOneRefusal()
    {
        string good = WriteFile("good.txt");
        string verbatim = @"\\?\" + WriteFile("verbatim.txt");
        string device = Path.Combine(_directory, "NUL.txt");
        string stream = good + ":hidden";

        ArgumentException ex = Assert.Throws<ArgumentException>(
            () => DraftAttachments.Validate(new[] { good, verbatim, device, stream }));

        Assert.Contains("3 unusable entries", ex.Message, StringComparison.Ordinal);
        Assert.Contains("'" + verbatim + "': is a \\\\?\\ path", ex.Message, StringComparison.Ordinal);
        Assert.Contains("'" + device + "': is called after a Windows device", ex.Message, StringComparison.Ordinal);
        Assert.Contains("'" + stream + "': holds a ':'", ex.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// A file of this test whose full path is exactly <paramref name="length"/> characters, in folders
    /// of at most 200 characters each.
    /// </summary>
    private string WriteFileWithPathLength(int length, string label)
    {
        string folder = Path.Combine(_directory, label);
        string name = label + ".txt";
        while (folder.Length + 1 + name.Length < length)
        {
            int room = length - folder.Length - 1 - name.Length;
            if (room <= 200)
            {
                name = label + new string('n', room) + ".txt";
                break;
            }

            folder = Path.Combine(folder, new string('d', 150));
        }

        string path = Path.Combine(folder, name);
        Assert.Equal(length, path.Length);
        Directory.CreateDirectory(folder);
        File.WriteAllText(path, "a file whose path is " + length + " characters long");
        return path;
    }
}
