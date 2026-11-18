using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// save_attachment writes only into OutlookAI's download folder (Q198 (a)) - every way of naming a
/// folder pinned against a disk described by hand, so each case answers the same on any machine:
/// relative segments, junctions and symbolic links, 8.3 short names, network and device paths,
/// alternate data streams, and the names Windows will not keep. The same checks against the real
/// disk - links made for the test, the name claim, the write - are
/// <see cref="AttachmentFileWriterDiskTests"/>.
/// <para>
/// The download folder here sits under a placeholder profile (AGENTS.md, Personal data), and nothing
/// in this file touches the disk: <see cref="FakeDisk"/> answers what the product would ask Windows.
/// </para>
/// </summary>
public sealed class AttachmentFileWriterTests
{
    private const string Profile = @"C:\Users\tester";

    private const string Local = Profile + @"\AppData\Local";

    private const string Home = Local + @"\OutlookAI\scratch\attachments";

    private const string Startup = Profile + @"\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup";

    // ------------------------------------------------------------------ the folder: inside

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void NoFolder_IsTheDownloadFolderItself(string? asked)
    {
        Assert.Equal(Home, Resolve(asked, Disk()));
    }

    [Theory]
    [InlineData("invoices", @"\invoices")]
    [InlineData("project-x/2026", @"\project-x\2026")]
    [InlineData(@"project-x\2026\", @"\project-x\2026")]
    [InlineData(@"a\.\b\..\c", @"\a\c")]
    [InlineData(@"a\\b", @"\a\b")]
    [InlineData(" invoices ", @"\invoices")]
    [InlineData("Offerte 2026 (def)", @"\Offerte 2026 (def)")]
    [InlineData(".hidden", @"\.hidden")]
    public void ARelativeFolder_IsAFolderInsideIt(string asked, string below)
    {
        Assert.Equal(Home + below, Resolve(asked, Disk()));
    }

    [Theory]
    [InlineData(Home)]
    [InlineData(Home + @"\")]
    [InlineData(Home + @"\sub")]
    [InlineData(@"c:\users\TESTER\appdata\local\outlookai\SCRATCH\attachments\sub")]
    [InlineData("C:/Users/tester/AppData/Local/OutlookAI/scratch/attachments/sub")]
    [InlineData(Home + @"\a\..\sub")]
    [InlineData(Local + @"\OutlookAI\scratch\x\..\attachments\sub")]
    public void AFullPathInsideIt_IsAccepted_HoweverItIsSpelled(string asked)
    {
        string resolved = Resolve(asked, Disk());

        Assert.True(AttachmentFileWriter.IsWithin(resolved, Home), resolved);
    }

    // ------------------------------------------------------------------ the folder: relative segments

    [Theory]
    [InlineData("..")]
    [InlineData(@"..\..\..\Roaming\Microsoft\Windows\Start Menu\Programs\Startup")]
    [InlineData(@"a\..\..\x")]
    [InlineData(@"..\attachments-old")]
    [InlineData(Home + @"\..\attachments-old")]
    [InlineData(Home + @"\..\..\..\..")]
    public void DotDotThatLeavesIt_IsRefused(string asked)
    {
        AttachmentFolderRefusedException refused = AssertRefused(asked, Disk());

        Assert.Contains("is outside OutlookAI's download folder", refused.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(@"C:\..\Windows")]
    [InlineData(@"..\..\..\..\..\..\..\..\..\..\..\Windows")]
    public void DotDotAboveTheDrive_IsRefused(string asked)
    {
        AttachmentFolderRefusedException refused = AssertRefused(asked, Disk());

        Assert.Contains("climbs above the root of its drive", refused.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(Home + "-old")]
    [InlineData(Home + "2")]
    [InlineData(Home + @"2\x")]
    public void ASiblingThatStartsWithTheSameLetters_IsOutside(string asked)
    {
        AssertRefused(asked, Disk().Folder(Home + "-old").Folder(Home + "2"));
    }

    [Theory]
    [InlineData(Startup)]
    [InlineData(@"C:\Windows\System32")]
    [InlineData(@"C:\")]
    [InlineData(@"D:\attachments")]
    [InlineData(Profile + @"\Documents")]
    [InlineData(Local + @"\OutlookAI\scratch")]
    [InlineData(Local + @"\OutlookAI")]
    public void AnyOtherFolderOrDrive_IsRefused(string asked)
    {
        AttachmentFolderRefusedException refused = AssertRefused(asked, Disk());

        Assert.Contains("is outside OutlookAI's download folder", refused.Message, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ the folder: refused on its text

    [Theory]
    [InlineData(@"\\server\share\inbox")]
    [InlineData("//server/share/inbox")]
    [InlineData(@"\\?\" + Home)]
    [InlineData(@"\\?\" + Home + @"\..\..\..\Roaming")]
    [InlineData(@"\\.\" + Home)]
    [InlineData(@"\\?\UNC\server\share\inbox")]
    [InlineData(@"\\.\pipe\outlookai")]
    [InlineData("//?/C:/Windows")]
    [InlineData(@"\??\C:\Windows")]
    [InlineData(@"\Windows")]
    [InlineData("/Windows")]
    public void ANetworkOrDevicePath_IsRefusedBeforeTheDiskIsAsked(string asked)
    {
        FakeDisk disk = Disk();

        AssertRefused(asked, disk);

        // The text alone decides it: \\?\ and \\.\ skip Windows' own path rules, so a check that
        // asked the disk about such a path would be checking something Windows does not do.
        Assert.Equal(0, disk.Questions);
    }

    [Theory]
    [InlineData("sub:stream")]
    [InlineData(Home + ":stream")]
    [InlineData(Home + @"\sub::$INDEX_ALLOCATION")]
    [InlineData(Home + @"\sub:$I30:$INDEX_ALLOCATION")]
    [InlineData("C:sub")]
    [InlineData("C:")]
    [InlineData(@"sub\C:\x")]
    public void AColonThatIsNotADrivesLetter_IsRefused(string asked)
    {
        FakeDisk disk = Disk();

        AttachmentFolderRefusedException refused = AssertRefused(asked, disk);

        Assert.Contains("':'", refused.Message, StringComparison.Ordinal);
        Assert.Equal(0, disk.Questions);
    }

    [Fact]
    public void ACharacterNoFolderCanHave_IsRefused()
    {
        // In a Fact rather than as theory data: a control character in a test's name is a character
        // the results file cannot hold.
        foreach (string asked in new[] { "a*b", "a?b", "a<b", "a>b", "a|b", "a\"b", "a\u0001b", "a\tb", "a\u0000b", "x\\a\u001Fb" })
        {
            FakeDisk disk = Disk();

            AttachmentFolderRefusedException refused = AssertRefused(asked, disk);

            Assert.Contains("no folder name can have", refused.Message, StringComparison.Ordinal);
            Assert.Equal(0, disk.Questions);
        }
    }

    [Fact]
    public void OuterSpaces_AreDropped_AsAroundAnyArgument()
    {
        Assert.Equal(Home + @"\sub", Resolve("  sub  ", Disk()));
    }

    [Theory]
    [InlineData("sub.")]
    [InlineData(@"sub.\x")]
    [InlineData(@"sub \x")]
    [InlineData("...")]
    [InlineData(@"a\...\b")]
    [InlineData(@"a\. \b")]
    public void AFolderNameEndingInADotOrASpace_IsRefused(string asked)
    {
        // Windows drops those, so "sub." would be made as "sub" - not the folder that was named.
        AttachmentFolderRefusedException refused = AssertRefused(asked, Disk());

        Assert.Contains("ending in a dot or a space", refused.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("CON")]
    [InlineData("nul")]
    [InlineData("Com1")]
    [InlineData("LPT9")]
    [InlineData("nul.txt")]
    [InlineData(@"x\AUX\y")]
    [InlineData("COM\u00B9")]
    [InlineData("CONIN$")]
    [InlineData(Home + @"\PRN")]
    public void AWindowsDeviceName_IsRefused(string asked)
    {
        AttachmentFolderRefusedException refused = AssertRefused(asked, Disk());

        Assert.Contains("names a Windows device", refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AFolderNameLongerThanNtfsKeeps_IsRefused()
    {
        AttachmentFolderRefusedException refused = AssertRefused(new string('x', 256), Disk());

        Assert.Contains("longer than 255 characters", refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AFolderTooDeepForAFileNameToFit_IsRefused()
    {
        string deep = string.Join(@"\", Enumerable.Repeat("abcdefghij", 20));

        AttachmentFolderRefusedException refused = AssertRefused(deep, Disk());

        Assert.Contains("is too long", refused.Message, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ the folder: links and short names

    [Fact]
    public void AJunctionInsideThatLeadsOut_IsRefused_AndTheRefusalSaysWhereItLeads()
    {
        FakeDisk disk = Disk().Folder(Home + @"\startup-link", really: Startup);

        foreach (string asked in new[] { "startup-link", @"startup-link\deeper\still", Home + @"\startup-link" })
        {
            AttachmentFolderRefusedException refused = AssertRefused(asked, disk);
            Assert.Contains("it is really '" + Startup, refused.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void AJunctionInsideThatStaysInside_IsAccepted_AsTheFolderItLeadsTo()
    {
        FakeDisk disk = Disk().Folder(Home + @"\latest", really: Home + @"\2026\week-47");

        Assert.Equal(Home + @"\2026\week-47\new", Resolve(@"latest\new", disk));
    }

    [Fact]
    public void ALinkHigherUpThatLeadsIn_IsAccepted()
    {
        // Somebody's own shortcut to the download folder, used as the full path.
        FakeDisk disk = Disk().Folder(@"C:\shortcut-to-downloads", really: Home);

        Assert.Equal(Home + @"\x", Resolve(@"C:\shortcut-to-downloads\x", disk));
    }

    [Fact]
    public void AShortNameOfAFolderInside_IsAccepted_AsItsLongName()
    {
        FakeDisk disk = Disk().Folder(Local + @"\OUTLOO~1\scratch\ATTACH~1", really: Home);

        Assert.Equal(Home + @"\new", Resolve(Local + @"\OUTLOO~1\scratch\ATTACH~1\new", disk));
    }

    [Fact]
    public void AShortNameOfAFolderOutside_IsRefused_AndTheRefusalNamesItsLongName()
    {
        FakeDisk disk = Disk().Folder(@"C:\PROGRA~1", really: @"C:\Program Files");

        AttachmentFolderRefusedException refused = AssertRefused(@"C:\PROGRA~1\x", disk);

        Assert.Contains(@"it is really 'C:\Program Files\x'", refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ADownloadFolderAJunctionMoved_IsFollowed()
    {
        // OutlookAI's whole state folder moved to another drive with a junction, by its user. The
        // download folder is where it really is, and a folder is inside it however it is named.
        FakeDisk disk = new FakeDisk()
            .Folder(@"C:\")
            .Folder(Local + @"\OutlookAI", really: @"E:\OutlookAI")
            .Folder(@"E:\OutlookAI");

        Assert.Equal(@"E:\OutlookAI\scratch\attachments", Resolve(null, disk));
        Assert.Equal(@"E:\OutlookAI\scratch\attachments\x", Resolve("x", disk));
        Assert.Equal(@"E:\OutlookAI\scratch\attachments\y", Resolve(@"E:\OutlookAI\scratch\attachments\y", disk));
        Assert.Equal(@"E:\OutlookAI\scratch\attachments\z", Resolve(Home + @"\z", disk));
        AssertRefused(@"E:\OutlookAI\other", disk);
        AssertRefused(Home + @"\..\..\other", disk);
    }

    [Fact]
    public void ADownloadFolderThatDoesNotExistYet_IsDecidedFromTheFolderAboveIt()
    {
        FakeDisk disk = new FakeDisk().Folder(Local);

        Assert.Equal(Home, Resolve(null, disk));
        Assert.Equal(Home + @"\sub", Resolve("sub", disk));
        AssertRefused(Local + @"\elsewhere", disk);
    }

    [Fact]
    public void ALinkThatLeadsNowhere_IsRefused()
    {
        FakeDisk disk = Disk().Dangling(Home + @"\gone");

        AttachmentFolderRefusedException refused = AssertRefused(@"gone\x", disk);

        Assert.Contains("cannot be checked", refused.Message, StringComparison.Ordinal);
        Assert.Contains("a link that leads nowhere", refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AFolderWindowsWillNotDescribe_IsRefused()
    {
        FakeDisk disk = Disk().Unreadable(Home + @"\locked");

        AttachmentFolderRefusedException refused = AssertRefused(@"locked\x", disk);

        Assert.Contains("cannot be checked", refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AFileInTheWay_IsRefused()
    {
        FakeDisk disk = Disk().File(Home + @"\report.pdf");

        AttachmentFolderRefusedException refused = AssertRefused(@"report.pdf\x", disk);

        Assert.Contains("is a file, not a folder", refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ADriveThatIsNotThere_IsRefused()
    {
        AttachmentFolderRefusedException refused = AssertRefused(@"Q:\attachments", Disk());

        Assert.Contains("drive that is not there", refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ADownloadFolderThatCannotBeResolved_IsRefusedAsSuch()
    {
        FakeDisk disk = new FakeDisk().Folder(@"C:\").Dangling(Profile);

        AttachmentFolderRefusedException refused = AssertRefused(null, disk);

        Assert.StartsWith("OutlookAI's download folder cannot be checked", refused.Message, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ what every refusal says

    [Theory]
    [InlineData(@"\\server\share")]
    [InlineData(Startup)]
    [InlineData("..")]
    [InlineData("sub:stream")]
    [InlineData("CON")]
    public void EveryRefusal_SaysNothingWasWritten_WhereTheFolderIs_AndWhatToDoInstead(string asked)
    {
        AttachmentFolderRefusedException refused = AssertRefused(asked, Disk());

        Assert.StartsWith("target_dir '" + asked + "' ", refused.Message, StringComparison.Ordinal);
        Assert.Contains("Nothing was written.", refused.Message, StringComparison.Ordinal);
        Assert.Contains("'" + Home + "'", refused.Message, StringComparison.Ordinal);
        Assert.Contains("copy it with your own file tools, under their own permission rules", refused.Message, StringComparison.Ordinal);

        // An ArgumentException, as every refused argument is - but its own type, so the tool layer
        // can say "unchanged" about it, which it cannot about an argument error in general.
        Assert.IsAssignableFrom<ArgumentException>(refused);
    }

    // ------------------------------------------------------------------ the sender's file name

    [Theory]
    [InlineData("report.pdf", "report.pdf")]
    [InlineData(@"..\..\..\Start Menu\Programs\Startup\evil.cmd", ".._.._.._Start Menu_Programs_Startup_evil.cmd")]
    [InlineData("../../evil.txt", ".._.._evil.txt")]
    [InlineData(@"C:\Windows\evil.txt", "C__Windows_evil.txt")]
    [InlineData("a:b.txt", "a_b.txt")]
    [InlineData("evil.txt::$DATA", "evil.txt__$DATA")]
    [InlineData("CON", "_CON")]
    [InlineData("con.txt", "_con.txt")]
    [InlineData("NUL.tar.gz", "_NUL.tar.gz")]
    [InlineData("COM1", "_COM1")]
    [InlineData("LPT9.log", "_LPT9.log")]
    [InlineData("COM\u00B9.txt", "_COM\u00B9.txt")]
    [InlineData("CONIN$", "_CONIN$")]
    [InlineData("CON .txt", "_CON .txt")]
    [InlineData("CONSOLE.txt", "CONSOLE.txt")]
    [InlineData("report.pdf. . .", "report.pdf")]
    [InlineData("report.pdf   ", "report.pdf")]
    [InlineData("  spaced  .txt", "spaced  .txt")]
    [InlineData("a\u007Fb.txt", "a_b.txt")]
    [InlineData("invoice\u202Efdp.exe", "invoice_fdp.exe")]
    [InlineData("x\u200Fy\u2066z.txt", "x_y_z.txt")]
    [InlineData("a*b?c\"d<e>f|g.txt", "a_b_c_d_e_f_g.txt")]
    [InlineData(".bashrc", ".bashrc")]
    [InlineData("\uFF0E\uFF0E", "\uFF0E\uFF0E")]
    public void TheSendersFileName_IsMadeOneNameWindowsKeepsAsWritten(string sent, string saved)
    {
        Assert.Equal(saved, AttachmentFileWriter.SafeFileName(sent));
    }

    [Fact]
    public void AControlCharacterInTheSendersName_BecomesAnUnderscore()
    {
        // In a Fact: a control character in a test's name is one the results file cannot hold.
        Assert.Equal("a_b.txt", AttachmentFileWriter.SafeFileName("a\tb.txt"));
        Assert.Equal("a_b.txt", AttachmentFileWriter.SafeFileName("a\u0000b.txt"));
        Assert.Equal("a__b.txt", AttachmentFileWriter.SafeFileName("a\r\nb.txt"));
        Assert.Equal("_report.pdf", AttachmentFileWriter.SafeFileName("\u001Freport.pdf"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData(".")]
    [InlineData("..")]
    [InlineData("...")]
    [InlineData(". . .")]
    public void ANameWithNothingLeft_BecomesTheFallback(string? sent)
    {
        Assert.Equal(AttachmentFileWriter.FallbackFileName, AttachmentFileWriter.SafeFileName(sent));
    }

    [Fact]
    public void ASafeName_HasNothingInItThatWindowsTreatsSpecially()
    {
        string[] hostile =
        {
            @"..\..\x", "a/b", "a:b", "CON", "nul.txt", "x. ", "\u202Eexe.pdf", "a\u0001b", "*", "?", "<>", "|", "\"",
        };

        foreach (string sent in hostile)
        {
            string name = AttachmentFileWriter.SafeFileName(sent);
            Assert.True(name.IndexOfAny(new[] { '\\', '/', ':', '*', '?', '"', '<', '>', '|' }) < 0, sent + " -> " + name);
            Assert.DoesNotContain(name, c => c < ' ' || c == '\u202E');
            Assert.False(name.EndsWith(".", StringComparison.Ordinal) || name.EndsWith(" ", StringComparison.Ordinal), name);
            Assert.False(WindowsFileNames.IsReservedDeviceName(name), name);
            Assert.NotEqual(".", name);
            Assert.NotEqual("..", name);
        }
    }

    // ------------------------------------------------------------------ device names, shared with manage_signature

    [Theory]
    [InlineData("CON")]
    [InlineData("con")]
    [InlineData("Prn")]
    [InlineData("AUX")]
    [InlineData("NUL")]
    [InlineData("COM0")]
    [InlineData("COM9")]
    [InlineData("LPT0")]
    [InlineData("lpt5")]
    [InlineData("COM\u00B2")]
    [InlineData("LPT\u00B3")]
    [InlineData("CONIN$")]
    [InlineData("conout$")]
    [InlineData("nul.txt")]
    [InlineData("NUL.tar.gz")]
    [InlineData("CON ")]
    [InlineData("con  .txt")]
    public void ADeviceName_IsReserved_AloneBeforeAnExtensionOrBeforeSpaces(string name)
    {
        Assert.True(WindowsFileNames.IsReservedDeviceName(name));
    }

    [Theory]
    [InlineData("CONSOLE")]
    [InlineData("COM10")]
    [InlineData("LPT")]
    [InlineData("NULL")]
    [InlineData("aux_file")]
    [InlineData("xCON")]
    [InlineData("_CON")]
    [InlineData(" CON")]
    [InlineData("com1x.txt")]
    [InlineData("")]
    [InlineData(null)]
    public void AnythingElse_IsNot(string? name)
    {
        Assert.False(WindowsFileNames.IsReservedDeviceName(name));
    }

    // ------------------------------------------------------------------ the pieces

    [Theory]
    [InlineData(Home, Home, true)]
    [InlineData(Home + @"\x", Home, true)]
    [InlineData(@"C:\USERS\TESTER\APPDATA\LOCAL\OUTLOOKAI\SCRATCH\ATTACHMENTS\X", Home, true)]
    [InlineData(Home + "-old", Home, false)]
    [InlineData(Home + "x", Home, false)]
    [InlineData(Local, Home, false)]
    [InlineData(@"C:\", @"C:\", true)]
    [InlineData(@"C:\x", @"C:\", true)]
    [InlineData(@"D:\x", @"C:\", false)]
    public void IsWithin_IsTheFolderOrInsideIt_NeverASiblingThatStartsTheSame(string path, string folder, bool within)
    {
        Assert.Equal(within, AttachmentFileWriter.IsWithin(path, folder));
    }

    [Fact]
    public void Fit_ShortensTheStem_KeepsTheEnding_AndNeverEndsInADotOrASpace()
    {
        Assert.Equal("report.pdf", AttachmentFileWriter.Fit("report", ".pdf", 255));

        string fitted = AttachmentFileWriter.Fit(new string('a', 300), " (2).pdf", 50);
        Assert.Equal(50, fitted.Length);
        Assert.EndsWith(" (2).pdf", fitted, StringComparison.Ordinal);

        Assert.Equal("abc.pdf", AttachmentFileWriter.Fit("abc. . . . . . ." + new string('z', 40), ".pdf", 13));
    }

    [Fact]
    public void Fit_NeverSplitsASurrogatePair()
    {
        // U+1F4CE (a paperclip) is two UTF-16 units; cutting between them leaves half a character.
        string stem = new string('a', 9) + "\U0001F4CE" + new string('b', 40);

        string fitted = AttachmentFileWriter.Fit(stem, ".pdf", 14);

        Assert.Equal(new string('a', 9) + ".pdf", fitted);
        Assert.DoesNotContain(fitted, char.IsSurrogate);
    }

    // ------------------------------------------------------------------ support

    private static FakeDisk Disk()
    {
        return new FakeDisk().Folder(@"C:\").Folder(@"D:\").Folder(Home);
    }

    private static string Resolve(string? asked, FakeDisk disk)
    {
        return AttachmentFileWriter.Resolve(asked, Home, disk).Folder;
    }

    private static AttachmentFolderRefusedException AssertRefused(string? asked, FakeDisk disk)
    {
        return Assert.Throws<AttachmentFolderRefusedException>(() => AttachmentFileWriter.Resolve(asked, Home, disk));
    }
}

/// <summary>
/// A disk described by hand, answering what <see cref="AttachmentFileWriter"/> asks Windows: which
/// spelled paths are folders and where each really is, which are files, which Windows will not
/// describe, and which are links that lead nowhere. A spelled path it was not told about is missing.
/// </summary>
internal sealed class FakeDisk : IFolderProbe
{
    private readonly Dictionary<string, string> _folders = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _files = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _unreadable = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _dangling = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>How many questions the disk was asked.</summary>
    internal int Questions { get; private set; }

    /// <summary>A folder at <paramref name="path"/> - or, with <paramref name="really"/>, a link or short name for that folder.</summary>
    internal FakeDisk Folder(string path, string? really = null)
    {
        _folders[path] = really ?? path;
        return this;
    }

    internal FakeDisk File(string path)
    {
        _files.Add(path);
        return this;
    }

    internal FakeDisk Unreadable(string path)
    {
        _unreadable.Add(path);
        return this;
    }

    /// <summary>A link at <paramref name="path"/> whose target is gone: it reads as a folder, and has no final path.</summary>
    internal FakeDisk Dangling(string path)
    {
        _folders[path] = path;
        _dangling.Add(path);
        return this;
    }

    public FolderEntry Inspect(string path, out int win32Error)
    {
        Questions++;
        win32Error = 0;
        if (_unreadable.Contains(path))
        {
            win32Error = 5;
            return FolderEntry.Unreadable;
        }

        if (_files.Contains(path))
        {
            return FolderEntry.NotADirectory;
        }

        if (_folders.ContainsKey(path))
        {
            return FolderEntry.Directory;
        }

        win32Error = 2;
        return FolderEntry.Missing;
    }

    public string FinalPath(string directory)
    {
        Questions++;
        if (_dangling.Contains(directory) || !_folders.TryGetValue(directory, out string? really))
        {
            throw new FileNotFoundException("The system cannot find the file specified. (" + directory + ")");
        }

        return really;
    }
}
