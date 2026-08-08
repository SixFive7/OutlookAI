using OutlookAI.Core.Mapi;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 codec tests against SYNTHETIC fixtures only - the encoding algorithm is
/// deterministic (byte -> WCHAR b + 0xAC00, v3.MD section 4), so fixtures are fabricated
/// in code. No live-recorded EntryIDs, store UIDs or subjects may ever be committed
/// (v3.MD S6 - public repo).
/// </summary>
public sealed class EntryIdCodecTests
{
    /// <summary>Fabricates a 24-byte message EntryID tail: marker flag byte + patterned UID + NID.</summary>
    internal static byte[] SyntheticEntryIdBytes(byte firstByte = 0xEF)
    {
        var bytes = new byte[EntryIdCodec.MessageEntryIdLength];
        bytes[0] = firstByte;
        // bytes 1..3: zero flags. bytes 4..19: patterned 16-byte store UID.
        for (int i = 0; i < 16; i++)
        {
            bytes[4 + i] = (byte)(0x10 + i);
        }

        // bytes 20..23: little-endian NID whose low 5 bits are 0x04 (NID_TYPE_NORMAL_MESSAGE).
        bytes[20] = 0x24; // 0x24 & 0x1F == 4
        bytes[21] = 0x01;
        bytes[22] = 0x00;
        bytes[23] = 0x00;
        return bytes;
    }

    internal const string SyntheticStoreUidHex = "101112131415161718191A1B1C1D1E1F";
    internal const string SyntheticEntryIdHex = "00000000" + SyntheticStoreUidHex + "24010000";

    internal static string SyntheticEncodedTail() => EntryIdCodec.EncodeBytes(SyntheticEntryIdBytes());

    [Fact]
    public void EncodeBytes_MapsBytesIntoHangulRange()
    {
        string encoded = EntryIdCodec.EncodeBytes(new byte[] { 0x00, 0x7F, 0xFF });

        Assert.Equal(3, encoded.Length);
        Assert.Equal((char)0xAC00, encoded[0]);
        Assert.Equal((char)(0xAC00 + 0x7F), encoded[1]);
        Assert.Equal((char)0xACFF, encoded[2]);
    }

    [Fact]
    public void TryDecodeEncodedBytes_RoundTripsEncode()
    {
        byte[] original = SyntheticEntryIdBytes();

        Assert.True(EntryIdCodec.TryDecodeEncodedBytes(EntryIdCodec.EncodeBytes(original), out byte[] decoded));
        Assert.Equal(original, decoded);
    }

    [Theory]
    [InlineData("plain-ascii")]
    [InlineData("")]
    [InlineData(null)]
    public void TryDecodeEncodedBytes_RejectsNonEncodedInput(string? input)
    {
        Assert.False(EntryIdCodec.TryDecodeEncodedBytes(input, out _));
    }

    [Fact]
    public void TryDecodeEncodedBytes_RejectsCharBeyondByteRange()
    {
        // 0xAD00 would be byte 0x100 - outside the encoded range.
        string invalid = EntryIdCodec.EncodeBytes(SyntheticEntryIdBytes()) + (char)0xAD00;

        Assert.False(EntryIdCodec.TryDecodeEncodedBytes(invalid, out _));
    }

    [Fact]
    public void TryDecodeMessageEntryId_RebuildsMarkerFlagsAsZeros()
    {
        Assert.True(EntryIdCodec.TryDecodeMessageEntryId(SyntheticEncodedTail(), out DecodedEntryId? decoded));

        Assert.NotNull(decoded);
        Assert.Equal(SyntheticEntryIdHex, decoded!.EntryIdHex);
        Assert.Equal(SyntheticStoreUidHex, decoded.StoreUidHex);
        Assert.Equal("24010000", decoded.NidHex);
        Assert.Equal(4, decoded.NidLowFiveBits);
        Assert.Equal(0xEF, decoded.RawFirstByte);
        Assert.StartsWith("00000000", decoded.EntryIdHex, StringComparison.Ordinal);
        Assert.Equal(48, decoded.EntryIdHex.Length);
    }

    [Fact]
    public void TryDecodeMessageEntryId_AcceptsAlreadyZeroFirstByte()
    {
        string encoded = EntryIdCodec.EncodeBytes(SyntheticEntryIdBytes(firstByte: 0x00));

        Assert.True(EntryIdCodec.TryDecodeMessageEntryId(encoded, out DecodedEntryId? decoded));
        Assert.Equal(SyntheticEntryIdHex, decoded!.EntryIdHex);
        Assert.Equal(0x00, decoded.RawFirstByte);
    }

    [Theory]
    [InlineData(23)]
    [InlineData(25)]
    [InlineData(4)]
    public void TryDecodeMessageEntryId_RejectsNonMessageLengths(int length)
    {
        var bytes = new byte[length];
        for (int i = 0; i < length; i++)
        {
            bytes[i] = (byte)i;
        }

        Assert.False(EntryIdCodec.TryDecodeMessageEntryId(EntryIdCodec.EncodeBytes(bytes), out _));
    }

    [Fact]
    public void TryDecodeAttachmentId_DecodesShortIds()
    {
        var attachmentId = new byte[] { 0x05, 0x14, 0x21, 0x00 };

        Assert.True(EntryIdCodec.TryDecodeAttachmentId(EntryIdCodec.EncodeBytes(attachmentId), out byte[] decoded));
        Assert.Equal(attachmentId, decoded);
    }

    [Fact]
    public void EncodeBytes_NullThrows()
    {
        Assert.Throws<ArgumentNullException>(() => EntryIdCodec.EncodeBytes(null!));
    }
}
