using OutlookAI.Core.Com;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the orphan-index-row messages (soak fix 16, part B2). The shipped text said
/// "Re-run search - the item may have moved" for EVERY locate failure, which is the one
/// remedy that provably cannot help a stale row: the next search returns the same row.
/// </summary>
public sealed class LocateFailureAdviceTests
{
    [Fact]
    public void FolderNotFound_SaysStaleIndexRow_AndDoesNotPromiseARetry()
    {
        // The measured case: ~458 rows in one delegate store filed under a leaf path with
        // no COM folder - returned by search, openable by nothing.
        string message = LocateFailureAdvice.Describe("url:FolderNotFound fallback:FolderNotFound");

        Assert.Contains("no longer exists in Outlook", message, StringComparison.Ordinal);
        Assert.Contains("stale index row", message, StringComparison.Ordinal);
        Assert.Contains("exhaustive:true", message, StringComparison.Ordinal);
        Assert.Contains("url:FolderNotFound", message, StringComparison.Ordinal);
        Assert.DoesNotContain("may have moved", message, StringComparison.Ordinal);
    }

    [Fact]
    public void StoreNotFound_PointsAtTheMailboxNotTheItem()
    {
        string message = LocateFailureAdvice.Describe("url:StoreNotFound fallback:StoreNotFound");

        Assert.Contains("not open in this Outlook profile", message, StringComparison.Ordinal);
        Assert.Contains("list_accounts", message, StringComparison.Ordinal);
    }

    [Fact]
    public void ItemGoneFromAnExistingFolder_StillRecommendsRerunning()
    {
        // Here re-running IS the right answer: the folder is fine, the item moved.
        string message = LocateFailureAdvice.Describe("url:NoSubjectTimeMatch fallback:NoSubjectTimeMatch");

        Assert.Contains("moved or deleted", message, StringComparison.Ordinal);
        Assert.Contains("Re-run the search", message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("url:FolderTooLargeForTimeOnlyProbe fallback:NoItemPathDisplay", "narrower folder")]
    [InlineData("url:RootFolderUnavailable fallback:RootFolderUnavailable", "outlook_health")]
    [InlineData("url:UrlNotParsable fallback:NoItemPathDisplay", "no usable location")]
    public void EveryKnownTokenGetsItsOwnRemedy(string token, string expected)
    {
        Assert.Contains(expected, LocateFailureAdvice.Describe(token), StringComparison.Ordinal);
    }

    [Fact]
    public void UnknownOrMissingTokens_StillProduceAnActionableSentence()
    {
        foreach (string? token in new[] { null, string.Empty, "something new" })
        {
            string message = LocateFailureAdvice.Describe(token);
            Assert.Contains("could not be opened in Outlook", message, StringComparison.Ordinal);
            Assert.Contains("exhaustive:true", message, StringComparison.Ordinal);
        }

        Assert.Contains("(unknown)", LocateFailureAdvice.Describe(null), StringComparison.Ordinal);
    }

    [Fact]
    public void IsMissingLocation_IdentifiesTheOrphanClassOnly()
    {
        Assert.True(LocateFailureAdvice.IsMissingLocation("url:FolderNotFound fallback:FolderNotFound"));
        Assert.True(LocateFailureAdvice.IsMissingLocation("url:StoreNotFound fallback:NoItemPathDisplay"));
        Assert.False(LocateFailureAdvice.IsMissingLocation("url:NoSubjectTimeMatch fallback:NoSubjectTimeMatch"));
        Assert.False(LocateFailureAdvice.IsMissingLocation(null));
    }
}
