using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The path rules <see cref="WindowsFileNames"/> holds once for save_attachment and the draft tools
/// (Q154): which paths are written in Windows' device namespace, which ':' names a stream, which names
/// end in a dot or a space. Its device names are pinned beside save_attachment's tests
/// (<see cref="AttachmentFileWriterTests"/>); what the draft tools make of all of it,
/// <see cref="DraftAttachmentsTests"/>.
/// </summary>
public sealed class WindowsFileNamesTests
{
    [Theory]
    [InlineData(@"\\?\C:\x\offer.pdf")]
    [InlineData(@"\\.\C:\x\offer.pdf")]
    [InlineData(@"\??\C:\x\offer.pdf")]
    [InlineData("//?/C:/x/offer.pdf")]
    [InlineData("//./pipe/outlookai")]
    [InlineData(@"\\?/C:\x")]
    [InlineData(@"/\.\C:\x")]
    [InlineData(@"\\?\UNC\server\share\offer.pdf")]
    [InlineData(@"\\.\")]
    public void ADeviceNamespacePath_IsOne_WithEitherSeparator(string path)
    {
        Assert.True(WindowsFileNames.IsDeviceNamespacePath(path));
    }

    [Theory]
    [InlineData(@"\\server\share\offer.pdf")]
    [InlineData(@"C:\x\offer.pdf")]
    [InlineData(@"\x\offer.pdf")]
    [InlineData("offer.pdf")]
    [InlineData(@"\\?")]
    [InlineData("/??/C:/x")]
    [InlineData(@"\\??\C:\x")]
    [InlineData(@"\\..\x")]
    [InlineData(@"\.\x")]
    [InlineData("")]
    [InlineData(null)]
    public void AnythingElse_IsNot(string? path)
    {
        Assert.False(WindowsFileNames.IsDeviceNamespacePath(path));
    }

    [Theory]
    [InlineData(@"C:\x\offer.pdf:hidden")]
    [InlineData(@"C:\x\offer.pdf::$DATA")]
    [InlineData(@"C:\x:y\offer.pdf")]
    [InlineData("C:offer.pdf")]
    [InlineData("C:")]
    [InlineData("offer.pdf:hidden")]
    [InlineData(@"\\server\share\offer.pdf:hidden")]
    [InlineData(@"\\.\C:\x\offer.pdf")]
    public void AColonThatIsNotAFullPathsDrive_IsBeyondTheDrive(string path)
    {
        // The last one is why the draft tools judge a device-namespace path without its prefix.
        Assert.True(WindowsFileNames.HasColonBeyondTheDrive(path));
    }

    [Theory]
    [InlineData(@"C:\x\offer.pdf")]
    [InlineData("C:/x/offer.pdf")]
    [InlineData(@"z:\")]
    [InlineData(@"\\server\share\offer.pdf")]
    [InlineData("offer.pdf")]
    [InlineData("")]
    public void ADrivesOwnColon_OrNone_IsNot(string path)
    {
        Assert.False(WindowsFileNames.HasColonBeyondTheDrive(path));
    }

    [Theory]
    [InlineData("offer.pdf.", true)]
    [InlineData("offer.pdf ", true)]
    [InlineData("offer.pdf. . .", true)]
    [InlineData(".", true)]
    [InlineData("..", true)]
    [InlineData("offer.pdf", false)]
    [InlineData(".hidden", false)]
    [InlineData(" offer.pdf", false)]
    [InlineData("", false)]
    public void ANameEndingInADotOrASpace_IsOne(string name, bool ends)
    {
        Assert.Equal(ends, WindowsFileNames.EndsInDotOrSpace(name));
    }

    [Fact]
    public void TheLongestOrdinaryPath_IsMaxPathLessItsNull_AndSaveAttachmentUsesTheSameNumber()
    {
        Assert.Equal(259, WindowsFileNames.MaxPathChars);
        Assert.Equal(WindowsFileNames.MaxPathChars, AttachmentFileWriter.MaxPathChars);
    }
}
