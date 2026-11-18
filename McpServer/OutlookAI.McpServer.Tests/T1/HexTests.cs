using System.Security.Cryptography;
using System.Text;
using OutlookAI.Core.AddIn;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 for Core's <c>Hex</c> (<c>McpServer/OutlookAI.Core/AddIn/Hex.cs</c>), the one writer and reader of
/// hexadecimal the server and the add-in share since the code-sharing refactor's hex stage (Q154). Every
/// writer is held to .NET's own writer or to a value known outside OutlookAI, never to Hex itself, so the
/// case each one names is the case it writes. The outputs whose case a reader depends on are pinned where
/// they are made: <c>SendConfirmationTests</c>, <c>DraftUpdateReentrancyTests</c>, <c>PagingTests</c>,
/// <c>AuditLogTests</c> and <c>SignatureListingTests</c>.
/// </summary>
public sealed class HexTests
{
    private static readonly byte[] EveryByte = Enumerable.Range(0, 256).Select(i => (byte)i).ToArray();

    [Fact]
    public void TheWriters_WriteEveryByte_InTheCaseTheyName()
    {
        Assert.Equal(Convert.ToHexString(EveryByte), Hex.Upper(EveryByte));
        Assert.Equal(Convert.ToHexStringLower(EveryByte), Hex.Lower(EveryByte));
        Assert.Equal("00FF7F", Hex.Upper(new byte[] { 0x00, 0xFF, 0x7F }));
        Assert.Equal("00ff7f", Hex.Lower(new byte[] { 0x00, 0xFF, 0x7F }));
        Assert.Equal(string.Empty, Hex.Upper(Array.Empty<byte>()));
    }

    [Fact]
    public void ARange_IsWrittenAlone()
    {
        Assert.Equal(Convert.ToHexString(EveryByte, 250, 6), Hex.Upper(EveryByte, 250, 6));
        Assert.Equal(Convert.ToHexStringLower(EveryByte, 10, 4), Hex.Lower(EveryByte, 10, 4));
        Assert.Equal(string.Empty, Hex.Lower(EveryByte, 256, 0));
    }

    [Fact]
    public void ANumber_IsWrittenInLowercase_PaddedToItsDigits()
    {
        // CRC-32's published check value - the CRC of "123456789" - as an audit line carries a CRC.
        Assert.Equal("cbf43926", Hex.Lower(0xCBF43926u, 8));
        Assert.Equal("0000000000000001", Hex.Lower(1UL, 16));
        Assert.Equal("ffffffffffffffff", Hex.Lower(ulong.MaxValue, 16));
        Assert.Equal("00000000", Hex.Lower(0UL, 8));
    }

    [Fact]
    public void TheHashWriters_WriteSha256_InTheCaseTheyName()
    {
        // FIPS 180-2's own example: the SHA-256 of "abc".
        const string Abc = "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad";
        byte[] abc = Encoding.ASCII.GetBytes("abc");

        Assert.Equal(Abc, Hex.Sha256Lower(abc));
        Assert.Equal(Abc.ToUpperInvariant(), Hex.Sha256Upper(abc));
        Assert.Equal(Convert.ToHexString(SHA256.HashData(EveryByte)), Hex.Sha256Upper(EveryByte));
    }

    [Fact]
    public void TheReader_ReadsEitherCase_AndEveryByteComesBack()
    {
        Assert.True(Hex.TryDecode("00ff7F", out byte[] three));
        Assert.Equal(new byte[] { 0x00, 0xFF, 0x7F }, three);
        Assert.True(Hex.TryDecode(Convert.ToHexString(EveryByte), out byte[] upper));
        Assert.Equal(EveryByte, upper);
        Assert.True(Hex.TryDecode(Convert.ToHexStringLower(EveryByte), out byte[] lower));
        Assert.Equal(EveryByte, lower);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("ABC")]
    [InlineData("store-pst")]
    [InlineData("0G")]
    [InlineData(" F")]
    [InlineData("F ")]
    [InlineData("0x0F")]
    public void TheReader_RefusesAnythingButPairsOfDigits_ASpaceIncluded(string? text)
    {
        // A space inside a pair is refused: the conversation index's own reader let one through until the
        // hex stage, as .NET's hex number parsing does, and no value it read could have held one.
        Assert.False(Hex.TryDecode(text, out byte[] none));
        Assert.Empty(none);
    }

    [Fact]
    public void TheDigits_ReadInEitherCase_AndTheirNeighboursAreNotDigits()
    {
        Assert.Equal(new[] { 0, 9, 10, 15, 10, 15 }, "09afAF".Select(Hex.DigitValue));

        // The characters either side of each run of digits in ASCII.
        Assert.All("/:@G`g ", c => Assert.Equal(-1, Hex.DigitValue(c)));
        Assert.True(Hex.IsDigit('F'));
        Assert.False(Hex.IsDigit('g'));
        Assert.True(Hex.IsLowerDigit('f'));
        Assert.False(Hex.IsLowerDigit('F'));
        Assert.True(Hex.IsAllDigits(string.Empty));
        Assert.False(Hex.IsAllDigits("12 4"));
    }

    [Fact]
    public void ARandomToken_IsLowercase_TwoDigitsAByte()
    {
        Assert.Matches("^[0-9a-f]{32}$", Hex.NewRandom(16));
    }
}
