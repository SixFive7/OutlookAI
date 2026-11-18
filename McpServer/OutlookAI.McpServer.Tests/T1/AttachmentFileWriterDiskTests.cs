using System.Runtime.InteropServices;

using OutlookAI.Core.Services;

using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// save_attachment's writer against the real disk (Q198 (a)), in a temporary folder of its own that
/// stands in for the download folder: what only Windows can answer. Junctions made for the test,
/// leading out and leading in; a directory symbolic link and a file link where Windows lets this
/// account make them; an 8.3 short name where the volume keeps them; the name claim that never
/// replaces a file, folder or link; and a write that fails before or part-way. The same checks on
/// a disk described by hand, for every spelling of a folder, are <see cref="AttachmentFileWriterTests"/>.
/// <para>
/// The attachment is <see cref="FakeOutlookAttachment"/>: what the writer asks of Outlook's
/// <c>Attachment</c> - its <c>FileName</c> and <c>SaveAsFile</c> - and nothing else. Outlook's own
/// <c>SaveAsFile</c> into the claimed file is proven on the test guests
/// (<c>T2.LiveSaveAttachmentFolderTests</c>).
/// </para>
/// </summary>
public sealed class AttachmentFileWriterDiskTests : IDisposable
{
    private readonly ITestOutputHelper _output;
    private readonly string _scratch;
    private readonly string _home;
    private readonly string _outside;
    private readonly List<string> _links = new();

    public AttachmentFileWriterDiskTests(ITestOutputHelper output)
    {
        _output = output;
        _scratch = Path.Combine(Path.GetTempPath(), "OutlookAI-AttachmentFileWriterDiskTests-" + Guid.NewGuid().ToString("N"));
        _home = Path.Combine(_scratch, "download");
        _outside = Path.Combine(_scratch, "outside");
        Directory.CreateDirectory(_home);
        Directory.CreateDirectory(_outside);
    }

    /// <summary>The download folder's long name, worked out without the code under test.</summary>
    private string RealHome => LongName(_home);

    public void Dispose()
    {
        // Links first, as links: deleting a folder recursively through one could follow it.
        foreach (string link in _links)
        {
            try
            {
                if (File.Exists(link))
                {
                    File.Delete(link);
                }
                else if (Directory.Exists(link))
                {
                    Directory.Delete(link, recursive: false);
                }
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }
        }

        try
        {
            Directory.Delete(_scratch, recursive: true);
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }

    // ------------------------------------------------------------------ the write

    [Fact]
    public void Save_WritesTheAttachmentIntoItsClaimedFile_AndReturnsItsRealPathAndSize()
    {
        FakeOutlookAttachment attachment = new FakeOutlookAttachment("report.pdf", Bytes("%PDF-1.7 report"));
        string? claimed = null;

        (string path, long size) = Save(attachment, null, c => claimed = c);

        Assert.Equal(Path.Combine(RealHome, "report.pdf"), path, ignoreCase: true);
        Assert.Equal(path, claimed);
        Assert.Equal(attachment.Content, File.ReadAllBytes(path));
        Assert.Equal(attachment.Content.Length, size);

        // The name was claimed - an empty file of its own - before Outlook was asked to write.
        Assert.Equal(path, attachment.WrittenTo);
        Assert.True(attachment.ClaimWasEmptyFile);
    }

    [Fact]
    public void Save_MakesAFolderThatIsMissing_InsideTheDownloadFolder()
    {
        (string path, _) = Save(new FakeOutlookAttachment("a.txt", Bytes("a")), @"project-x\2026");

        Assert.Equal(Path.Combine(RealHome, "project-x", "2026", "a.txt"), path, ignoreCase: true);
        Assert.True(File.Exists(path));
    }

    [Fact]
    public void Save_MakesTheDownloadFolderItself_WhenItIsNotThereYet()
    {
        string notYet = Path.Combine(_scratch, "not-yet", "scratch", "attachments");

        (string path, _) = AttachmentFileWriter.Save(
            new FakeOutlookAttachment("a.txt", Bytes("a")), null, null, notYet, AttachmentFileWriter.WindowsDisk);

        Assert.Equal(Path.Combine(LongName(_scratch), "not-yet", "scratch", "attachments", "a.txt"), path, ignoreCase: true);
    }

    [Fact]
    public void Save_NeverReplacesAFile_ItNumbersTheNewOne()
    {
        string existing = Path.Combine(_home, "report.pdf");
        File.WriteAllText(existing, "the user's own report");

        (string second, _) = Save(new FakeOutlookAttachment("report.pdf", Bytes("new")), null);
        (string third, _) = Save(new FakeOutlookAttachment("report.pdf", Bytes("newer")), null);

        Assert.Equal("report (2).pdf", Path.GetFileName(second));
        Assert.Equal("report (3).pdf", Path.GetFileName(third));
        Assert.Equal("the user's own report", File.ReadAllText(existing));
        Assert.Equal("new", File.ReadAllText(second));
        Assert.Equal("newer", File.ReadAllText(third));
    }

    [Fact]
    public void Save_CountsAFolderOfThatName_AsTaken()
    {
        Directory.CreateDirectory(Path.Combine(_home, "report.pdf"));

        (string path, _) = Save(new FakeOutlookAttachment("report.pdf", Bytes("x")), null);

        Assert.Equal("report (2).pdf", Path.GetFileName(path));
    }

    [Fact]
    public void Save_CountsAJunctionOfThatName_AsTaken_AndWritesNothingThroughIt()
    {
        MakeJunction(Path.Combine(_home, "report.pdf"), _outside);

        (string path, _) = Save(new FakeOutlookAttachment("report.pdf", Bytes("x")), null);

        Assert.Equal("report (2).pdf", Path.GetFileName(path));
        Assert.Empty(Directory.EnumerateFileSystemEntries(_outside));
    }

    [Fact]
    public void Save_CountsAFileLinkOfThatName_AsTaken_AndCreatesNothingWhereItPoints()
    {
        // The classic: a link at the name, pointing where nothing is yet. A create that followed it
        // would make the file at the far end - in the Startup folder, say.
        string planted = Path.Combine(_outside, "planted.cmd");
        string link = Path.Combine(_home, "report.pdf");
        if (!TryMakeFileLink(link, planted))
        {
            return;
        }

        (string path, _) = Save(new FakeOutlookAttachment("report.pdf", Bytes("x")), null);

        Assert.Equal("report (2).pdf", Path.GetFileName(path));
        Assert.False(File.Exists(planted), "the claim followed the link and created the file it points at");
    }

    [Fact]
    public void Save_ShortensANameTooLongForThePath_KeepingItsExtension()
    {
        FakeOutlookAttachment attachment = new FakeOutlookAttachment(new string('n', 300) + ".pdf", Bytes("x"));

        (string path, _) = Save(attachment, null);

        Assert.EndsWith(".pdf", path, StringComparison.Ordinal);
        Assert.True(path.Length <= AttachmentFileWriter.MaxPathChars, path.Length + " characters");
        Assert.True(File.Exists(path));
    }

    [Theory]
    [InlineData(@"..\..\..\Start Menu\Programs\Startup\evil.cmd", ".._.._.._Start Menu_Programs_Startup_evil.cmd")]
    [InlineData("CON", "_CON")]
    [InlineData("nul.txt", "_nul.txt")]
    [InlineData("report.pdf. . .", "report.pdf")]
    [InlineData("a:b.txt", "a_b.txt")]
    [InlineData("invoice‮fdp.exe", "invoice_fdp.exe")]
    [InlineData("", AttachmentFileWriter.FallbackFileName)]
    public void Save_WritesTheSendersFile_UnderItsSafeName_DirectlyInTheFolder(string sent, string saved)
    {
        (string path, _) = Save(new FakeOutlookAttachment(sent, Bytes("payload")), null);

        Assert.Equal(Path.Combine(RealHome, saved), path, ignoreCase: true);
        Assert.Equal("payload", File.ReadAllText(path));
        Assert.Single(Directory.EnumerateFileSystemEntries(_home));
        Assert.Empty(Directory.EnumerateFileSystemEntries(_outside));
    }

    [Fact]
    public void Save_RemovesTheEmptyClaim_WhenTheWriteFailsBeforeWriting()
    {
        FakeOutlookAttachment attachment = new FakeOutlookAttachment("report.pdf", Bytes("x"))
        {
            FailBeforeWriting = new COMException("Outlook refused.", unchecked((int)0x80004005)),
        };
        string? claimed = null;

        COMException failure = Assert.Throws<COMException>(() => Save(attachment, null, c => claimed = c));

        Assert.Equal(unchecked((int)0x80004005), failure.HResult);
        Assert.NotNull(claimed);
        Assert.False(File.Exists(claimed), "an empty claim was left behind");
    }

    [Fact]
    public void Save_KeepsAPartlyWrittenFile_WhenTheWriteFailsPartWay()
    {
        // A partial file is reported to the caller, who is told where it is - never deleted here.
        FakeOutlookAttachment attachment = new FakeOutlookAttachment("report.pdf", Bytes("partial"))
        {
            FailAfterWriting = new COMException("Outlook stopped.", unchecked((int)0x800706BE)),
        };
        string? claimed = null;

        Assert.Throws<COMException>(() => Save(attachment, null, c => claimed = c));

        Assert.Equal("partial", File.ReadAllText(claimed!));
    }

    // ------------------------------------------------------------------ links and short names

    [Fact]
    public void AJunctionThatLeadsOut_IsRefused_AndNothingIsMadeOrWrittenWhereItLeads()
    {
        MakeJunction(Path.Combine(_home, "j"), _outside);
        FakeOutlookAttachment attachment = new FakeOutlookAttachment("evil.cmd", Bytes("x"));

        foreach (string asked in new[] { "j", @"j\deeper\still", Path.Combine(_home, "j", "x") })
        {
            AttachmentFolderRefusedException checkRefused = Assert.Throws<AttachmentFolderRefusedException>(
                () => AttachmentFileWriter.Resolve(asked, _home, AttachmentFileWriter.WindowsDisk));
            Assert.Contains("it is really '" + LongName(_outside), checkRefused.Message, StringComparison.OrdinalIgnoreCase);

            Assert.Throws<AttachmentFolderRefusedException>(() => Save(attachment, asked));
        }

        Assert.Null(attachment.WrittenTo);
        Assert.Empty(Directory.EnumerateFileSystemEntries(_outside));
    }

    [Fact]
    public void AJunctionThatStaysInside_IsFollowed_AndTheFileIsWrittenInTheFolderItLeadsTo()
    {
        string real = Path.Combine(_home, "2026", "week-47");
        Directory.CreateDirectory(real);
        MakeJunction(Path.Combine(_home, "latest"), real);

        (string path, _) = Save(new FakeOutlookAttachment("a.txt", Bytes("a")), "latest");

        Assert.Equal(Path.Combine(LongName(real), "a.txt"), path, ignoreCase: true);
    }

    [Fact]
    public void ADirectorySymbolicLinkThatLeadsOut_IsRefused()
    {
        string link = Path.Combine(_home, "symlink-out");
        try
        {
            Directory.CreateSymbolicLink(link, _outside);
            _links.Add(link);
        }
        catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
        {
            // The junction tests above exercise the same resolution: Windows follows both kinds
            // when the folder is opened. Said out loud, so a pass here is never read as more.
            _output.WriteLine("NOT EXERCISED: this account may not create a symbolic link here (" + ex.Message + ").");
            return;
        }

        Assert.Throws<AttachmentFolderRefusedException>(() => Save(new FakeOutlookAttachment("x.txt", Bytes("x")), @"symlink-out\x"));
        Assert.Empty(Directory.EnumerateFileSystemEntries(_outside));
    }

    [Fact]
    public void AShortNameOfAFolderInside_IsAccepted_AsItsLongName()
    {
        string longFolder = Path.Combine(_home, "A Long Folder Name For Short");
        Directory.CreateDirectory(longFolder);
        string? shortFolder = ShortName(longFolder);
        if (shortFolder == null || string.Equals(shortFolder, longFolder, StringComparison.OrdinalIgnoreCase))
        {
            // AttachmentFileWriterTests pins the same case on a disk described by hand.
            _output.WriteLine("NOT EXERCISED: this volume keeps no 8.3 short name for '" + longFolder + "'.");
            return;
        }

        _output.WriteLine("short name: " + shortFolder);
        (string path, _) = Save(new FakeOutlookAttachment("a.txt", Bytes("a")), Path.Combine(shortFolder, "sub"));

        Assert.Equal(Path.Combine(LongName(longFolder), "sub", "a.txt"), path, ignoreCase: true);
    }

    [Fact]
    public void TheRealStartupFolder_IsRefused()
    {
        string startup = Environment.GetFolderPath(Environment.SpecialFolder.Startup, Environment.SpecialFolderOption.DoNotVerify);
        Assert.False(string.IsNullOrEmpty(startup), "Windows names no Startup folder for this account");

        Assert.Throws<AttachmentFolderRefusedException>(
            () => Save(new FakeOutlookAttachment("evil.cmd", Bytes("x")), startup));
        Assert.Throws<AttachmentFolderRefusedException>(
            () => AttachmentFileWriter.Resolve(Path.Combine(startup, "evil"), _home, AttachmentFileWriter.WindowsDisk));
    }

    [Fact]
    public void AFullPathOutside_IsRefused_AndNothingIsWritten()
    {
        FakeOutlookAttachment attachment = new FakeOutlookAttachment("x.txt", Bytes("x"));

        Assert.Throws<AttachmentFolderRefusedException>(() => Save(attachment, _outside));
        Assert.Throws<AttachmentFolderRefusedException>(() => Save(attachment, Path.Combine(_home, "..", "outside")));
        Assert.Throws<AttachmentFolderRefusedException>(() => Save(attachment, @"..\outside"));

        Assert.Null(attachment.WrittenTo);
        Assert.Empty(Directory.EnumerateFileSystemEntries(_outside));
    }

    [Fact]
    public void AFileInTheWay_IsRefused()
    {
        File.WriteAllText(Path.Combine(_home, "report.pdf"), "a file");

        AttachmentFolderRefusedException refused = Assert.Throws<AttachmentFolderRefusedException>(
            () => Save(new FakeOutlookAttachment("x.txt", Bytes("x")), @"report.pdf\x"));

        Assert.Contains("is a file, not a folder", refused.Message, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ support

    private (string Path, long SizeBytes) Save(FakeOutlookAttachment attachment, string? asked, Action<string>? reserved = null)
    {
        return AttachmentFileWriter.Save(attachment, asked, reserved, _home, AttachmentFileWriter.WindowsDisk);
    }

    private static byte[] Bytes(string text)
    {
        return System.Text.Encoding.UTF8.GetBytes(text);
    }

    private void MakeJunction(string link, string target)
    {
        TestJunction.Make(link, target);
        _links.Add(link);
    }

    private bool TryMakeFileLink(string link, string target)
    {
        try
        {
            File.CreateSymbolicLink(link, target);
            _links.Add(link);
            return true;
        }
        catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
        {
            _output.WriteLine("NOT EXERCISED: this account may not create a symbolic link here (" + ex.Message + ").");
            return false;
        }
    }

    /// <summary>A path's long name, from <c>GetLongPathNameW</c> - an answer the code under test does not give.</summary>
    private static string LongName(string path)
    {
        char[] buffer = new char[1024];
        uint length = GetLongPathNameW(path, buffer, (uint)buffer.Length);
        Assert.True(length > 0 && length < buffer.Length, "GetLongPathNameW failed for " + path);
        return new string(buffer, 0, (int)length).TrimEnd('\\');
    }

    private static string? ShortName(string path)
    {
        char[] buffer = new char[1024];
        uint length = GetShortPathNameW(path, buffer, (uint)buffer.Length);
        return length == 0 || length >= buffer.Length ? null : new string(buffer, 0, (int)length);
    }

    [DllImport("kernel32.dll", EntryPoint = "GetLongPathNameW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetLongPathNameW(string lpszShortPath, [Out] char[] lpszLongPath, uint cchBuffer);

    [DllImport("kernel32.dll", EntryPoint = "GetShortPathNameW", CharSet = CharSet.Unicode, SetLastError = true)]
    private static extern uint GetShortPathNameW(string lpszLongPath, [Out] char[] lpszShortPath, uint cchBuffer);
}

/// <summary>
/// What <see cref="AttachmentFileWriter"/> asks of Outlook's <c>Attachment</c>: its <c>FileName</c>
/// and <c>SaveAsFile</c>, which - like Outlook's - writes over the file at the path it is given.
/// Public, because the writer calls it the way it calls Outlook, through <c>dynamic</c>, from
/// another assembly.
/// </summary>
public sealed class FakeOutlookAttachment
{
    public FakeOutlookAttachment(string? fileName, byte[] content)
    {
        FileName = fileName;
        Content = content;
    }

    /// <summary>The name the sender gave the attachment.</summary>
    public string? FileName { get; }

    /// <summary>What the attachment holds.</summary>
    public byte[] Content { get; }

    /// <summary>Thrown before anything is written, when set.</summary>
    public Exception? FailBeforeWriting { get; set; }

    /// <summary>Thrown after <see cref="Content"/> is written, when set.</summary>
    public Exception? FailAfterWriting { get; set; }

    /// <summary>Where <see cref="SaveAsFile"/> was asked to write.</summary>
    public string? WrittenTo { get; private set; }

    /// <summary>Whether that path held an empty file of its own when the write began - the claim.</summary>
    public bool ClaimWasEmptyFile { get; private set; }

    public void SaveAsFile(string path)
    {
        WrittenTo = path;
        FileInfo claim = new FileInfo(path);
        ClaimWasEmptyFile = claim.Exists && claim.Length == 0 && (claim.Attributes & FileAttributes.ReparsePoint) == 0;
        if (FailBeforeWriting != null)
        {
            throw FailBeforeWriting;
        }

        File.WriteAllBytes(path, Content);
        if (FailAfterWriting != null)
        {
            throw FailAfterWriting;
        }
    }
}
