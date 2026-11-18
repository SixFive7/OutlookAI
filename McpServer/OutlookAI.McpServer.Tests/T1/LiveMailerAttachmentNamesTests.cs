using OutlookAI.McpServer.Tests.T2;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The live mailer's guard on the file names it may give a test draft's attachments
/// (<see cref="LiveOutlookTestMailer.SaveTaggedDraftWithAttachments"/>, for
/// <c>T2.LiveSaveAttachmentFolderTests</c>): one entry per attachment, none blank. Pinned here
/// because the live tier that uses it cannot run in this one.
/// </summary>
public sealed class LiveMailerAttachmentNamesTests
{
    private static readonly string[] Paths = { @"C:\Temp\a.txt", @"C:\Temp\b.txt" };

    [Fact]
    public void NoNames_OrOnePerAttachment_AreAccepted()
    {
        Assert.Null(LiveOutlookTestMailer.RefuseAttachmentFileNames(Paths, null));
        Assert.Null(LiveOutlookTestMailer.RefuseAttachmentFileNames(Paths, new[] { @"..\..\x.txt", null }));
        Assert.Null(LiveOutlookTestMailer.RefuseAttachmentFileNames(Paths, new[] { "CON.txt", "a:b.txt" }));
    }

    [Fact]
    public void AWrongCount_OrABlankName_IsRefused()
    {
        Assert.NotNull(LiveOutlookTestMailer.RefuseAttachmentFileNames(Paths, new[] { "a.txt" }));
        Assert.NotNull(LiveOutlookTestMailer.RefuseAttachmentFileNames(Paths, new string?[] { "a.txt", "b.txt", "c.txt" }));
        Assert.NotNull(LiveOutlookTestMailer.RefuseAttachmentFileNames(Paths, new[] { "a.txt", " " }));
        Assert.NotNull(LiveOutlookTestMailer.RefuseAttachmentFileNames(Array.Empty<string>(), new[] { "a.txt" }));
    }
}
