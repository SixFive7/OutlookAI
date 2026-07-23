using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Phase-7 payload-size review (v3.MD section 12 compact-payload discipline): PINS
/// every tool-facing cap so accidental cap creep fails the build, and covers the
/// capped list mappers (recipients/attachments) that keep read/draft/send payloads
/// bounded with has-more indicators.
/// </summary>
public sealed class PayloadDisciplineTests
{
    [Fact]
    public void Caps_ArePinned()
    {
        // Search hits: small pages, iterate instead of dumping (section 8/12).
        Assert.Equal(100, MailService.SearchTopCap);
        Assert.Equal(25, MailService.SearchTopDefault);
        Assert.Equal(1000, MailService.SnippetCharsCap);
        Assert.Equal(200, MailService.SnippetCharsDefault);

        // Thread members.
        Assert.Equal(200, MailService.ThreadTopCap);
        Assert.Equal(50, MailService.ThreadTopDefault);

        // Read: body/header text caps (flags: bodyTruncated/headersTruncated).
        Assert.Equal(500_000, MailService.BodyCharsCap);
        Assert.Equal(20_000, MailService.BodyCharsDefault);
        Assert.Equal(256, MailService.HeaderCharsMin);
        Assert.Equal(65_536, MailService.HeaderCharsCap);
        Assert.Equal(8_192, MailService.HeaderCharsDefault);

        // List caps with has-more indicators (Phase-7 additions).
        Assert.Equal(100, MailService.RecipientsCap);
        Assert.Equal(100, MailService.AttachmentsCap);

        // Folder listing bounds (no unbounded folder walks).
        Assert.Equal(1000, MailService.FoldersCap);
        Assert.Equal(300, MailService.FoldersDefault);
        Assert.Equal(6, MailService.FolderDepthCap);
    }

    [Fact]
    public void SearchRequest_Defaults_AreCompact()
    {
        SearchRequest request = new();

        Assert.Equal(MailService.SearchTopDefault, request.Top);
        Assert.Equal(MailService.SnippetCharsDefault, request.SnippetChars);
        Assert.Equal(SearchMode.Fresh, request.Mode);
        Assert.True(request.IncludeAttachmentHits);
    }

    [Fact]
    public void CapRecipients_UnderCap_ReturnsAll_NotTruncated()
    {
        var recipients = MakeRecipients(3);

        var views = MailService.CapRecipients(recipients, out int total, out bool truncated);

        Assert.Equal(3, views.Count);
        Assert.Equal(3, total);
        Assert.False(truncated);
        Assert.Equal("to", views[0].Kind);
        Assert.Equal("r1@example.com", views[0].Address);
    }

    [Fact]
    public void CapRecipients_OverCap_CapsWithHasMore()
    {
        var recipients = MakeRecipients(MailService.RecipientsCap + 57);

        var views = MailService.CapRecipients(recipients, out int total, out bool truncated);

        Assert.Equal(MailService.RecipientsCap, views.Count);
        Assert.Equal(MailService.RecipientsCap + 57, total);
        Assert.True(truncated);
        // Order preserved: the first cap-many recipients are listed.
        Assert.Equal("r1@example.com", views[0].Address);
        Assert.Equal($"r{MailService.RecipientsCap}@example.com", views[^1].Address);
    }

    [Fact]
    public void CapRecipients_ExactlyCap_NotTruncated()
    {
        var recipients = MakeRecipients(MailService.RecipientsCap);

        var views = MailService.CapRecipients(recipients, out int total, out bool truncated);

        Assert.Equal(MailService.RecipientsCap, views.Count);
        Assert.Equal(MailService.RecipientsCap, total);
        Assert.False(truncated);
    }

    [Fact]
    public void CapAttachments_OverCap_CapsAndKeepsOriginalIndexes()
    {
        var attachments = new List<ComAttachmentInfo>();
        for (int i = 1; i <= MailService.AttachmentsCap + 5; i++)
        {
            attachments.Add(new ComAttachmentInfo(i, $"file{i}.txt", i * 10));
        }

        var views = MailService.CapAttachments(attachments, out int total, out bool truncated);

        Assert.Equal(MailService.AttachmentsCap, views.Count);
        Assert.Equal(MailService.AttachmentsCap + 5, total);
        Assert.True(truncated);
        // 1-based ORIGINAL indexes survive the cap - save_attachment still addresses
        // unlisted attachments by index.
        Assert.Equal(1, views[0].Index);
        Assert.Equal(MailService.AttachmentsCap, views[^1].Index);
    }

    [Fact]
    public void CapAttachments_UnderCap_ReturnsAll_NotTruncated()
    {
        var attachments = new List<ComAttachmentInfo> { new(1, "a.pdf", 100), new(2, "b.pdf", null) };

        var views = MailService.CapAttachments(attachments, out int total, out bool truncated);

        Assert.Equal(2, views.Count);
        Assert.Equal(2, total);
        Assert.False(truncated);
        Assert.Null(views[1].SizeBytes);
    }

    private static List<ComRecipientInfo> MakeRecipients(int count)
    {
        var list = new List<ComRecipientInfo>(count);
        for (int i = 1; i <= count; i++)
        {
            list.Add(new ComRecipientInfo("to", $"Recipient {i}", $"r{i}@example.com"));
        }

        return list;
    }
}
