using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 row-mapper tests with synthetic provider rows (S6): the mapper must tolerate
/// DBNull/missing columns, provider array shapes, and derive EntryID/attachment mapping
/// from the item URL.
/// </summary>
public sealed class IndexRowMapperTests
{
    private const string Sid = "{S-1-5-21-1111111111-2222222222-3333333333-1001}";
    private static readonly string StorePrefix = $"mapi16://{Sid}/alice@example.com($deadbeef)";

    private static Dictionary<string, object?> MessageRow()
    {
        string itemUrl = $"{StorePrefix}/0/Inbox/{EntryIdCodecTests.SyntheticEncodedTail()}";
        return new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = itemUrl,
            ["System.Subject"] = "synthetic subject",
            ["System.Message.FromAddress"] = "billing@example.com",
            ["System.Message.FromName"] = "Billing Dept",
            ["System.Message.ToAddress"] = new object[] { "alice@example.com", "bob@example.com" },
            ["System.Message.DateReceived"] = new DateTime(2026, 7, 20, 10, 15, 30, DateTimeKind.Unspecified),
            ["System.ItemPathDisplay"] = "/alice@example.com/Inbox/synthetic subject",
            ["System.ItemNameDisplay"] = "synthetic subject",
            ["System.Kind"] = new object[] { "email" },
            ["System.Search.AutoSummary"] = "snippet text",
            ["System.Size"] = 12345L,
            ["System.IsRead"] = true,
            ["System.Message.HasAttachments"] = false,
            ["System.Message.ConversationID"] = "CAFEBABE",
        };
    }

    [Fact]
    public void Map_MessageRow_MapsAllFieldsAndDecodesEntryId()
    {
        IndexHit hit = IndexRowMapper.Map(MessageRow());

        Assert.Equal("synthetic subject", hit.Subject);
        Assert.Equal("billing@example.com", hit.FromAddress);
        Assert.Equal("Billing Dept", hit.FromName);
        Assert.Equal(new[] { "alice@example.com", "bob@example.com" }, hit.ToAddresses);
        Assert.Equal(new[] { "email" }, hit.Kinds);
        Assert.Equal("snippet text", hit.AutoSummary);
        Assert.Equal(12345L, hit.SizeBytes);
        Assert.True(hit.IsRead);
        Assert.False(hit.HasAttachments);
        Assert.Equal("CAFEBABE", hit.ConversationId);

        Assert.NotNull(hit.DateReceivedUtc);
        Assert.Equal(DateTimeKind.Utc, hit.DateReceivedUtc!.Value.Kind);
        Assert.Equal(new DateTime(2026, 7, 20, 10, 15, 30, DateTimeKind.Utc), hit.DateReceivedUtc.Value);

        Assert.Equal(EntryIdCodecTests.SyntheticEntryIdHex, hit.EntryIdHex);
        Assert.Equal(EntryIdCodecTests.SyntheticStoreUidHex, hit.StoreUidHex);
        Assert.Equal("alice@example.com", hit.StoreDisplayName);
        Assert.Equal(StorePrefix, hit.StorePrefix);
        Assert.Equal(0, hit.StoreType);
        Assert.Equal(new[] { "Inbox" }, hit.FolderSegments);
        Assert.False(hit.IsAttachmentHit);
        Assert.Null(hit.ParentItemUrl);
    }

    [Fact]
    public void Map_AttachmentRow_MapsParentEntryIdAndFileName()
    {
        string parentUrl = $"{StorePrefix}/0/Archive/{EntryIdCodecTests.SyntheticEncodedTail()}";
        string attachUrl = $"{parentUrl}/at={EntryIdCodec.EncodeBytes(new byte[] { 0x05, 0x14 })}:invoice.pdf";
        var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = attachUrl,
            ["System.Kind"] = new object[] { "document" },
            ["System.ItemNameDisplay"] = "invoice.pdf",
        };

        IndexHit hit = IndexRowMapper.Map(row);

        Assert.True(hit.IsAttachmentHit);
        Assert.Equal("invoice.pdf", hit.AttachmentFileName);
        Assert.Equal(parentUrl, hit.ParentItemUrl);
        Assert.Equal(new[] { "document" }, hit.Kinds);
        // Attachment hits resolve to the PARENT message EntryID (v3.MD section 4).
        Assert.Equal(EntryIdCodecTests.SyntheticEntryIdHex, hit.EntryIdHex);
    }

    [Fact]
    public void Map_DbNullAndMissingColumns_YieldNullsAndEmptyLists()
    {
        var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = DBNull.Value,
            ["System.Subject"] = DBNull.Value,
            ["System.Message.DateReceived"] = DBNull.Value,
        };

        IndexHit hit = IndexRowMapper.Map(row);

        Assert.Equal(string.Empty, hit.ItemUrl);
        Assert.Null(hit.Subject);
        Assert.Null(hit.DateReceivedUtc);
        Assert.Empty(hit.ToAddresses);
        Assert.Empty(hit.Kinds);
        Assert.Null(hit.EntryIdHex);
        Assert.Null(hit.StorePrefix);
        Assert.False(hit.IsAttachmentHit);
    }

    [Fact]
    public void Map_ScalarKindAndStringArrays_Normalize()
    {
        var row = MessageRow();
        row["System.Kind"] = "email";
        row["System.Message.ToAddress"] = new[] { "alice@example.com" };

        IndexHit hit = IndexRowMapper.Map(row);

        Assert.Equal(new[] { "email" }, hit.Kinds);
        Assert.Equal(new[] { "alice@example.com" }, hit.ToAddresses);
    }

    [Fact]
    public void Map_NumericConversions_AreTolerant()
    {
        var row = MessageRow();
        row["System.Size"] = (decimal)987654;
        row["System.IsRead"] = 0;

        IndexHit hit = IndexRowMapper.Map(row);

        Assert.Equal(987654L, hit.SizeBytes);
        Assert.False(hit.IsRead);
    }

    [Fact]
    public void Map_ConversationIdBytes_HexEncoded()
    {
        var row = MessageRow();
        row["System.Message.ConversationID"] = new byte[] { 0xCA, 0xFE };

        IndexHit hit = IndexRowMapper.Map(row);

        Assert.Equal("CAFE", hit.ConversationId);
    }

    [Fact]
    public void Map_NullRow_Throws()
    {
        Assert.Throws<ArgumentNullException>(() => IndexRowMapper.Map(null!));
    }
}
