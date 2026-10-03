using OutlookAI.Core.Mapi;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1: the percent-encoding Outlook applies to a store or folder NAME inside a Windows Search
/// MAPI URL (<see cref="MapiUrlSegment"/>) - five characters, documented by Microsoft and
/// measured on the indexed test guest, for a store (Q99) and for folders
/// (<c>T2/LiveFolderNameEncodingTests</c>; <c>Docs/live-tier-on-the-vm.md</c> section 8 item 26).
/// <para>
/// The pins that matter most are the two quiet ones: a name WITHOUT any of the five characters
/// must come back byte for byte - that is every folder name the product has ever searched - and
/// both directions must be single-pass, so a name that merely looks like an escape survives.
/// </para>
/// </summary>
public sealed class MapiUrlSegmentTests
{
    [Theory]
    [InlineData("%", "%25")]
    [InlineData("/", "%2F")]
    [InlineData("\\", "%5C")]
    [InlineData("*", "%2A")]
    [InlineData("?", "%3F")]
    public void EachDocumentedCharacter_IsEncodedAsMicrosoftWritesIt(string character, string escape)
    {
        Assert.Equal(escape, MapiUrlSegment.Encode(character));
        Assert.Equal("a" + escape + "b", MapiUrlSegment.Encode("a" + character + "b"));
        Assert.Equal(character, MapiUrlSegment.Decode(escape));
    }

    [Fact]
    public void TheFiveCharacters_AreExactlyTheDocumentedOnes()
    {
        Assert.Equal("%/\\*?", MapiUrlSegment.EncodedCharacters);
    }

    [Theory]
    [InlineData("Inbox")]
    [InlineData("Sent Items")]
    [InlineData("Auto ongeluk")]
    [InlineData("O'Brien")]
    [InlineData("Projects [2026] {draft} #1 & co + more ~ ok")]
    [InlineData("Postvak IN")]
    [InlineData("Ελληνικά")]
    [InlineData("가나다")]
    [InlineData("alice@example.com")]
    [InlineData("a.b-c_d,e;f=g!h(i)j")]
    [InlineData("")]
    public void ANameWithNoneOfTheFive_IsReturnedUnchanged_InBothDirections(string name)
    {
        // Byte-identical, and the SAME instance: every ordinary name builds exactly the scope it
        // always did, with nothing even copied.
        Assert.Same(name, MapiUrlSegment.Encode(name));
        Assert.Same(name, MapiUrlSegment.Decode(name));
        Assert.Same(name, MapiUrlSegment.EncodePath(name));
    }

    [Fact]
    public void TheMeasuredStoreName_IsSpelledAsTheIndexFiledIt()
    {
        // Q99, measured on OutlookAI-Indexed: 'q99 50% off*?x' is filed 'q99 50%25 off%2A%3Fx($5159380d)'.
        Assert.Equal("q99 50%25 off%2A%3Fx", MapiUrlSegment.Encode("q99 50% off*?x"));
        Assert.Equal("q99 50% off*?x", MapiUrlSegment.Decode("q99 50%25 off%2A%3Fx"));
    }

    [Theory]
    [InlineData("100%*? mix", "100%25%2A%3F mix")]
    [InlineData("a/b\\c", "a%2Fb%5Cc")]
    [InlineData("%%", "%25%25")]
    [InlineData("?*?*", "%3F%2A%3F%2A")]
    [InlineData("50% off / 2026\\Q1 *final*?", "50%25 off %2F 2026%5CQ1 %2Afinal%2A%3F")]
    public void ANameMixingSeveral_EncodesEveryOneOfThem(string name, string encoded)
    {
        Assert.Equal(encoded, MapiUrlSegment.Encode(name));
        Assert.Equal(name, MapiUrlSegment.Decode(encoded));
    }

    [Theory]
    [InlineData("%2A", "%252A")]
    [InlineData("%25", "%2525")]
    [InlineData("50%2F off", "50%252F off")]
    [InlineData("%5C%3F", "%255C%253F")]
    public void ANameThatLooksLikeAnEscape_RoundTripsAsItself(string name, string encoded)
    {
        // The chain-of-Replace bug, pinned in both directions: '%' must be encoded in the SAME pass
        // as the others, and a decoded '%25' must never be decoded again.
        Assert.Equal(encoded, MapiUrlSegment.Encode(name));
        Assert.Equal(name, MapiUrlSegment.Decode(encoded));
    }

    [Theory]
    [InlineData("%2a", "*")]
    [InlineData("%3f", "?")]
    [InlineData("%2f", "/")]
    [InlineData("%5c", "\\")]
    public void LowerCaseHex_DecodesToo(string escape, string character)
    {
        Assert.Equal(character, MapiUrlSegment.Decode(escape));
    }

    [Theory]
    [InlineData("%20")]
    [InlineData("%41")]
    [InlineData("%")]
    [InlineData("50%")]
    [InlineData("%2")]
    [InlineData("%G1")]
    [InlineData("%2G")]
    [InlineData("100% sure")]
    public void AnythingButTheFiveEscapes_IsLeftAsItIs(string segment)
    {
        // Outlook is documented and measured to write those five and nothing else, so a '%' that
        // starts no such escape is kept rather than guessed at - no worse than before the decode.
        Assert.Equal(segment, MapiUrlSegment.Decode(segment));
    }

    [Fact]
    public void DecodeInvertsEncode_ForEveryStringOverTheTrickyAlphabet()
    {
        // Deterministic sweep: every string up to length 4 over an alphabet holding the five
        // characters and the letters and digits their escapes are spelled with.
        char[] alphabet = { '%', '/', '\\', '*', '?', '2', '5', 'A', 'F', 'C', '3', 'a', ' ' };
        int checkedCount = 0;
        foreach (string s in Strings(alphabet, 4))
        {
            Assert.Equal(s, MapiUrlSegment.Decode(MapiUrlSegment.Encode(s)));
            checkedCount++;
        }

        Assert.True(checkedCount > 30000, $"the sweep checked only {checkedCount} strings");
    }

    [Theory]
    [InlineData("Inbox/Fun", "Inbox/Fun")]
    [InlineData("Inbox//Fun", "Inbox//Fun")]
    [InlineData("Clients/50% off/Q1*?", "Clients/50%25 off/Q1%2A%3F")]
    [InlineData("back\\slash/why?", "back%5Cslash/why%3F")]
    [InlineData("%2A/%25", "%252A/%2525")]
    public void APath_IsEncodedSegmentBySegment_AndKeepsItsSeparators(string path, string encoded)
    {
        Assert.Equal(encoded, MapiUrlSegment.EncodePath(path));
    }

    [Fact]
    public void Null_IsRefused()
    {
        Assert.Throws<ArgumentNullException>(() => MapiUrlSegment.Encode(null!));
        Assert.Throws<ArgumentNullException>(() => MapiUrlSegment.Decode(null!));
        Assert.Throws<ArgumentNullException>(() => MapiUrlSegment.EncodePath(null!));
    }

    private static IEnumerable<string> Strings(char[] alphabet, int maxLength)
    {
        List<string> current = new() { string.Empty };
        yield return string.Empty;
        for (int length = 1; length <= maxLength; length++)
        {
            List<string> next = new(current.Count * alphabet.Length);
            foreach (string prefix in current)
            {
                foreach (char c in alphabet)
                {
                    string s = prefix + c;
                    next.Add(s);
                    yield return s;
                }
            }

            current = next;
        }
    }
}
