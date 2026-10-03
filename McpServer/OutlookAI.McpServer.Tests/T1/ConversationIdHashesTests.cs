using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The live drafts test's reading of which string a ConversationId is a hash of
/// (<see cref="ConversationIdHashes"/>) - pinned here because the live test prints only its verdict.
/// </summary>
public sealed class ConversationIdHashesTests
{
    [Fact]
    public void Md5Hex_IsTheStandardDigest_UpperCaseHex()
    {
        // RFC 1321's own test vector: MD5("abc").
        Assert.Equal("900150983CD24FB0D6963F7D28E17F72", ConversationIdHashes.Md5Hex("abc", "utf8"));
        Assert.Equal(ConversationIdHashes.Md5Hex("ABC", "utf8"), ConversationIdHashes.Md5Hex("abc", "utf8-upper"));
        Assert.NotEqual(ConversationIdHashes.Md5Hex("abc", "utf8"), ConversationIdHashes.Md5Hex("abc", "utf16"));
    }

    [Fact]
    public void Describe_NamesEveryCandidateAndVariantThatMatches_AndNothingElse()
    {
        string id = ConversationIdHashes.Md5Hex("[OutlookAI-McpTest] renamed", "utf16-upper");

        Assert.Equal(
            "newSubject/utf16-upper",
            ConversationIdHashes.Describe(id, ("newSubject", "[OutlookAI-McpTest] renamed"), ("keptTopic", "[OutlookAI-McpTest] seed")));
        Assert.Equal("none", ConversationIdHashes.Describe(id, ("keptTopic", "[OutlookAI-McpTest] seed")));
        Assert.Equal("-", ConversationIdHashes.Describe(null, ("newSubject", "x")));
        Assert.Equal("none", ConversationIdHashes.Describe(id, ("missing", null)));
    }
}
