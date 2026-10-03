using System.Text;
using OutlookAI.Core.Mapi;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins <see cref="StoreHash"/> - Microsoft's documented store hash, the <c>($hash)</c> Outlook
/// writes into every index URL - against values MEASURED on the indexed test guest (Q92/Q99,
/// Office LTSC 2024, Docs/live-tier-on-the-vm.md section 8 item 24). Every vector below is a
/// store a real Outlook pushed and the index then carried under exactly this hash; the paths are
/// the guest's, synthetic test data. The two Store.StoreID literals are byte for byte what
/// Outlook reported over COM, so the layout helper below is pinned to Outlook's bytes rather
/// than to this file's reading of the documentation.
/// </summary>
public sealed class StoreHashTests
{
    /// <summary>Store.StoreID of the tier store, C:\OutlookAI-Tier\Outlook.pst, exactly as Outlook reported it.</summary>
    private const string TierStoreId =
        "0000000038A1BB1005E5101AA1BB08002B2A56C200006D737073742E646C6C00000000004E495441F9BFB80100AA0037D96E00000000"
        + "43003A005C004F00750074006C006F006F006B00410049002D0054006900650072005C004F00750074006C006F006F006B002E007000730074000000";

    /// <summary>Store.StoreID of an ANSI (97-2002) PST, C:\OutlookAI-Tier\q99-ansi.pst: an 8-bit path, one reserved byte before it.</summary>
    private const string AnsiStoreId =
        "0000000038A1BB1005E5101AA1BB08002B2A56C200006D737073742E646C6C00000000004E495441F9BFB80100AA0037D96E0000"
        + "00433A5C4F75746C6F6F6B41492D546965725C7139392D616E73692E70737400";

    /// <summary>
    /// The wrapped entry ID Outlook gives a Unicode-path PST store: the MAPI store wrapper, "mspst.dll",
    /// the PST provider's UID, a reserved byte, an empty 8-bit path, then the UTF-16 path and its NUL.
    /// </summary>
    private static byte[] UnicodePstStoreEntryId(string path)
    {
        byte[] header = Convert.FromHexString(
            "0000000038A1BB1005E5101AA1BB08002B2A56C200006D737073742E646C6C00000000004E495441F9BFB80100AA0037D96E0000" + "0000");
        byte[] wide = Encoding.Unicode.GetBytes(path + "\0");
        return header.Concat(wide).ToArray();
    }

    [Theory]
    [InlineData(@"C:\OutlookAI-Tier\Outlook.pst", 0x93F42B43u)]
    [InlineData(@"C:\OutlookAI-Tier\identity.pst", 0xB25AC20Au)]
    [InlineData(@"C:\OutlookAI-Tier\Outlook Data File - CorpusProfile.pst", 0x23A27F0Du)]
    [InlineData(@"C:\OutlookAI-Tier\Outlook Data File - IdentityMint.pst", 0xBE889D8Bu)]
    [InlineData(@"C:\OutlookAI-Tier\q98-scratch.pst", 0x65D10200u)]
    [InlineData(@"C:\OutlookAI-Tier\moved\q98-scratch.pst", 0xBEFBE850u)]
    [InlineData(@"C:\OutlookAI-Tier\q99-lz-12.pst", 0x0CE9D6E4u)]
    [InlineData(@"C:\OutlookAI-Tier\q99-special.pst", 0x5159380Du)]
    public void EveryMeasuredUnicodePst_HashesToTheValueItsIndexUrlCarried(string path, uint measured)
    {
        Assert.Equal(measured, StoreHash.Compute(UnicodePstStoreEntryId(path)));
    }

    [Fact]
    public void TheHashIsOverOutlooksOwnStoreId_ForAUnicodeAndAnAnsiPst()
    {
        Assert.True(StoreHash.TryDecodeHex(TierStoreId, out byte[] tier));
        Assert.Equal(UnicodePstStoreEntryId(@"C:\OutlookAI-Tier\Outlook.pst"), tier);
        Assert.Equal(0x93F42B43u, StoreHash.Compute(tier));

        Assert.True(StoreHash.TryDecodeHex(AnsiStoreId, out byte[] ansi));
        Assert.Equal(84, ansi.Length);
        Assert.Equal(0x54556EE0u, StoreHash.Compute(ansi));
    }

    [Fact]
    public void OneChangedByte_GivesAnotherHash_AndSoDoesTheSamePathAtAnotherPlace()
    {
        // The controls for the vectors above: the hash is a function of every byte, so a store
        // copied or moved to another path is - correctly - another index store.
        byte[] tier = UnicodePstStoreEntryId(@"C:\OutlookAI-Tier\Outlook.pst");
        tier[tier.Length - 6] ^= 0x01;
        Assert.NotEqual(0x93F42B43u, StoreHash.Compute(tier));
        Assert.NotEqual(
            StoreHash.Compute(UnicodePstStoreEntryId(@"C:\OutlookAI-Tier\q98-scratch.pst")),
            StoreHash.Compute(UnicodePstStoreEntryId(@"C:\OutlookAI-Tier\moved\q98-scratch.pst")));
    }

    [Fact]
    public void TheAlgorithm_IsMicrosoftsComputeHash_DwordsLittleEndianThenTailThenPubThenPath()
    {
        // Worked by hand from the documented code, so a wrong word order, a dropped tail or the
        // '.PUB' constant in the wrong place each changes the value:
        //   dword 0x04030201 -> h = 0x04030201; tail byte 0x05 -> h*33 + 5.
        uint h = 0x04030201u;
        h = unchecked((h << 5) + h + 0x05u);
        Assert.Equal(h, StoreHash.Compute(new byte[] { 0x01, 0x02, 0x03, 0x04, 0x05 }));

        uint pub = unchecked((h << 5) + h + StoreHash.PublicStoreConstant);
        Assert.Equal(pub, StoreHash.Compute(new byte[] { 0x01, 0x02, 0x03, 0x04, 0x05 }, publicStore: true));

        uint path = unchecked((pub << 5) + pub + 'C');
        path = unchecked((path << 5) + path + ':');
        Assert.Equal(path, StoreHash.Compute(new byte[] { 0x01, 0x02, 0x03, 0x04, 0x05 }, publicStore: true, fileName: "C:"));
        Assert.Equal(0x2E505542u, StoreHash.PublicStoreConstant); // ".PUB" read as a little-endian DWORD
        Assert.Equal(0u, StoreHash.Compute(Array.Empty<byte>()));
    }

    [Theory]
    [InlineData("ce9d6e4", 0x0CE9D6E4u)] // measured: the URL drops the leading zero
    [InlineData("0ce9d6e4", 0x0CE9D6E4u)]
    [InlineData("93F42B43", 0x93F42B43u)]
    public void UrlHashes_CompareAsNumbers_SoALeadingZeroCannotSplitOneStoreInTwo(string digits, uint value)
    {
        Assert.True(StoreHash.TryParseUrlHash(digits, out uint parsed));
        Assert.Equal(value, parsed);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("123456789")]
    [InlineData("93f42b4g")]
    [InlineData("-1")]
    public void AnythingButOneToEightHexDigits_IsNoHash(string? digits)
    {
        Assert.False(StoreHash.TryParseUrlHash(digits, out _));
    }

    [Fact]
    public void HexThatDoesNotDecode_YieldsNoBytes()
    {
        Assert.False(StoreHash.TryDecodeHex("store-pst", out byte[] none));
        Assert.Empty(none);
        Assert.False(StoreHash.TryDecodeHex("ABC", out _));
        Assert.False(StoreHash.TryDecodeHex(null, out _));
        Assert.True(StoreHash.TryDecodeHex("00ff7F", out byte[] three));
        Assert.Equal(new byte[] { 0x00, 0xFF, 0x7F }, three);
    }

    [Fact]
    public void ANonExchangeStore_HasOneCandidate_ItsEntryId()
    {
        IReadOnlyList<StoreHashCandidate> candidates = StoreHash.Candidates(
            TierStoreId, exchangeStore: false, publicStore: false, "AABBCCDD", "11223344", @"C:\x.ost");

        StoreHashCandidate only = Assert.Single(candidates);
        Assert.Equal(StoreHashInput.EntryId, only.Input);
        Assert.Equal(0x93F42B43u, only.Hash);
    }

    [Fact]
    public void AnExchangeStore_HasEveryDocumentedInput_AndPublicFoldersMixInPub()
    {
        const string signature = "0102030405060708090A0B0C0D0E0F10";
        IReadOnlyList<StoreHashCandidate> candidates = StoreHash.Candidates(
            TierStoreId, exchangeStore: true, publicStore: false, signature, signature, @"C:\Users\u\x.ost");

        Assert.Equal(
            new[] { StoreHashInput.EntryId, StoreHashInput.ProfileMappingSignature, StoreHashInput.MappingSignature, StoreHashInput.EntryIdAndPath },
            candidates.Select(c => c.Input));
        StoreHash.TryDecodeHex(signature, out byte[] sig);
        Assert.Equal(StoreHash.Compute(sig), candidates[1].Hash);
        Assert.Equal(candidates[1].Hash, candidates[2].Hash);

        IReadOnlyList<StoreHashCandidate> pub = StoreHash.Candidates(
            TierStoreId, exchangeStore: true, publicStore: true, signature, null, null);
        Assert.Equal(new[] { StoreHashInput.EntryId, StoreHashInput.MappingSignature }, pub.Select(c => c.Input));
        Assert.Equal(StoreHash.Compute(sig, publicStore: true), pub[1].Hash);
        Assert.NotEqual(StoreHash.Compute(sig), pub[1].Hash);
    }

    [Fact]
    public void AnUnreadableStoreId_LeavesNoEntryIdCandidate()
    {
        Assert.Empty(StoreHash.Candidates("store-pst", exchangeStore: false, publicStore: false, null, null, null));
        Assert.Empty(StoreHash.Candidates(null, exchangeStore: true, publicStore: false, null, null, @"C:\x.ost"));
    }
}
