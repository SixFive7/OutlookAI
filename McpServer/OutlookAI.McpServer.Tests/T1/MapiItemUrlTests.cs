using OutlookAI.Core.Mapi;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 URL-grammar tests on synthetic URLs (fabricated SID, example.com store, fabricated
/// hash - v3.MD S6). Grammar per v3.MD section 4:
/// mapi16://{SID}/Store($hash)/StoreType/Folders.../EncodedEntryID[/at=EncodedAttachId:FileName]
/// </summary>
public sealed class MapiItemUrlTests
{
    private const string Sid = "{S-1-5-21-1111111111-2222222222-3333333333-1001}";
    private const string StoreSegment = "alice@example.com($deadbeef)";
    private static readonly string StorePrefix = $"mapi16://{Sid}/{StoreSegment}";

    private static string MessageUrl(string folders = "Inbox")
        => $"{StorePrefix}/0/{folders}/{EntryIdCodecTests.SyntheticEncodedTail()}";

    [Fact]
    public void TryParse_MessageUrl_ParsesAllComponents()
    {
        Assert.True(MapiItemUrl.TryParse(MessageUrl("Inbox/Sub Folder"), out MapiItemUrl? url));

        Assert.NotNull(url);
        Assert.Equal("mapi16", url!.Scheme);
        Assert.Equal(Sid, url.SidSegment);
        Assert.Equal(StoreSegment, url.StoreSegment);
        Assert.Equal("alice@example.com", url.StoreDisplayName);
        Assert.Equal("deadbeef", url.StoreUrlHash);
        Assert.Equal(StorePrefix, url.StorePrefix);
        Assert.Equal(0, url.StoreType);
        Assert.Equal(new[] { "Inbox", "Sub Folder" }, url.FolderSegments);
        Assert.False(url.IsAttachment);
        Assert.Null(url.ParentItemUrl);
        Assert.NotNull(url.EncodedItemSegment);

        Assert.True(url.TryDecodeEntryId(out DecodedEntryId? decoded));
        Assert.Equal(EntryIdCodecTests.SyntheticEntryIdHex, decoded!.EntryIdHex);
    }

    [Fact]
    public void TryParse_AttachmentUrl_MapsParentAndFileName()
    {
        string parent = MessageUrl("Archive");
        string encodedAttachId = EntryIdCodec.EncodeBytes(new byte[] { 0x05, 0x14, 0x21, 0x00 });
        string attachmentUrl = $"{parent}/at={encodedAttachId}:report 2024.pdf";

        Assert.True(MapiItemUrl.TryParse(attachmentUrl, out MapiItemUrl? url));

        Assert.True(url!.IsAttachment);
        Assert.Equal("report 2024.pdf", url.AttachmentFileName);
        Assert.Equal(encodedAttachId, url.EncodedAttachmentId);
        Assert.Equal(parent, url.ParentItemUrl);
        Assert.Equal(new[] { "Archive" }, url.FolderSegments);

        // The decoded EntryID of an attachment URL is the PARENT message id.
        Assert.True(url.TryDecodeEntryId(out DecodedEntryId? decoded));
        Assert.Equal(EntryIdCodecTests.SyntheticEntryIdHex, decoded!.EntryIdHex);

        Assert.True(EntryIdCodec.TryDecodeAttachmentId(url.EncodedAttachmentId, out byte[] attachBytes));
        Assert.Equal(new byte[] { 0x05, 0x14, 0x21, 0x00 }, attachBytes);
    }

    [Fact]
    public void TryParse_AttachmentFileNameMayContainColons()
    {
        string parent = MessageUrl();
        string encodedAttachId = EntryIdCodec.EncodeBytes(new byte[] { 0x01 });
        string attachmentUrl = $"{parent}/at={encodedAttachId}:odd:name.pdf";

        Assert.True(MapiItemUrl.TryParse(attachmentUrl, out MapiItemUrl? url));
        Assert.Equal("odd:name.pdf", url!.AttachmentFileName);
        Assert.Equal(encodedAttachId, url.EncodedAttachmentId);
    }

    [Fact]
    public void TryParse_DelegateStoreTypeSegment()
    {
        string url = $"{StorePrefix}/1/Bob Delegate/Inbox/{EntryIdCodecTests.SyntheticEncodedTail()}";

        Assert.True(MapiItemUrl.TryParse(url, out MapiItemUrl? parsed));
        Assert.Equal(1, parsed!.StoreType);
        Assert.Equal(new[] { "Bob Delegate", "Inbox" }, parsed.FolderSegments);
        Assert.True(parsed.TryDecodeEntryId(out _));
    }

    [Fact]
    public void TryParse_StoreSegmentWithSpaceBeforeHash_TrimsDisplayName()
    {
        string url = $"mapi16://{Sid}/Shared Mailbox ($0abc1234)/0/Inbox/{EntryIdCodecTests.SyntheticEncodedTail()}";

        Assert.True(MapiItemUrl.TryParse(url, out MapiItemUrl? parsed));
        Assert.Equal("Shared Mailbox", parsed!.StoreDisplayName);
        Assert.Equal("0abc1234", parsed.StoreUrlHash);
    }

    [Fact]
    public void TryParse_StoreSegmentWithoutHash_KeepsFullDisplayName()
    {
        string url = $"mapi16://{Sid}/alice@example.com/0/Inbox/{EntryIdCodecTests.SyntheticEncodedTail()}";

        Assert.True(MapiItemUrl.TryParse(url, out MapiItemUrl? parsed));
        Assert.Equal("alice@example.com", parsed!.StoreDisplayName);
        Assert.Null(parsed.StoreUrlHash);
    }

    [Fact]
    public void TryParse_FolderUrlWithoutEntryTail_HasNoItemSegment()
    {
        string url = $"{StorePrefix}/0/Inbox/Sub Folder";

        Assert.True(MapiItemUrl.TryParse(url, out MapiItemUrl? parsed));
        Assert.Null(parsed!.EncodedItemSegment);
        Assert.False(parsed.IsAttachment);
        Assert.Equal(new[] { "Inbox", "Sub Folder" }, parsed.FolderSegments);
        Assert.False(parsed.TryDecodeEntryId(out _));
    }

    [Fact]
    public void TryParse_StorePrefixOnly_Parses()
    {
        Assert.True(MapiItemUrl.TryParse(StorePrefix, out MapiItemUrl? parsed));
        Assert.Equal(StorePrefix, parsed!.StorePrefix);
        Assert.Empty(parsed.FolderSegments);
        Assert.Null(parsed.StoreType);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("file:///c:/temp/x.txt")]
    [InlineData("https://example.com/a")]
    [InlineData("no-scheme-at-all")]
    [InlineData("mapi16://")]
    public void TryParse_RejectsNonMapiOrMalformedUrls(string? url)
    {
        Assert.False(MapiItemUrl.TryParse(url, out MapiItemUrl? parsed));
        Assert.Null(parsed);
    }

    [Fact]
    public void TryParse_AttachmentWithoutDecodableParent_IsRejected()
    {
        string url = $"{StorePrefix}/0/Inbox/at={EntryIdCodec.EncodeBytes(new byte[] { 0x01 })}:file.pdf";

        Assert.False(MapiItemUrl.TryParse(url, out _));
    }
}
