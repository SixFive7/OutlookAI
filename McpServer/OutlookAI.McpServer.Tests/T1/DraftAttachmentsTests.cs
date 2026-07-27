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
}
