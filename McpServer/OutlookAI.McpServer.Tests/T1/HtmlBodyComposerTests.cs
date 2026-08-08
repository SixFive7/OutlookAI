using OutlookAI.Core.Text;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 for the draft body composition rules (v3.MD sections 3/12): agent text becomes an
/// escaped HTML fragment and is inserted at the TOP of the existing HTML body - above
/// the signature Outlook injected and above the quoted reply/forward history.
/// </summary>
public sealed class HtmlBodyComposerTests
{
    [Fact]
    public void ToHtmlFragment_EscapesHtmlAndConvertsLineBreaks()
    {
        string fragment = HtmlBodyComposer.ToHtmlFragment("a<b & \"c\">\r\nnext\nlast\rend");

        Assert.Equal("<div>a&lt;b &amp; &quot;c&quot;&gt;<br>next<br>last<br>end</div>", fragment);
    }

    [Fact]
    public void ToHtmlFragment_EmptyTextYieldsEmptyDiv()
    {
        Assert.Equal("<div></div>", HtmlBodyComposer.ToHtmlFragment(string.Empty));
    }

    [Fact]
    public void InsertAtBodyTop_InsertsRightAfterOutlookStyleBodyTag()
    {
        string html = "<html><head><style>p{}</style></head>"
            + "<body lang=EN-US link=blue style='word-wrap:break-word'><p>signature</p><p>quote</p></body></html>";

        string result = HtmlBodyComposer.InsertAtBodyTop(html, "<div>agent</div>");

        Assert.Equal(
            "<html><head><style>p{}</style></head>"
            + "<body lang=EN-US link=blue style='word-wrap:break-word'><div>agent</div><p>signature</p><p>quote</p></body></html>",
            result);

        // The agent fragment must precede the pre-existing content (above the quote).
        Assert.True(result.IndexOf("agent", StringComparison.Ordinal) < result.IndexOf("signature", StringComparison.Ordinal));
    }

    [Fact]
    public void InsertAtBodyTop_MatchesBodyTagCaseInsensitively()
    {
        string result = HtmlBodyComposer.InsertAtBodyTop("<HTML><BODY><p>x</p></BODY></HTML>", "<div>a</div>");

        Assert.Equal("<HTML><BODY><div>a</div><p>x</p></BODY></HTML>", result);
    }

    [Fact]
    public void InsertAtBodyTop_DoesNotMatchTagsMerelyStartingWithBody()
    {
        // <bodyguard> is not a body tag; with no real body tag the fragment is prepended.
        string result = HtmlBodyComposer.InsertAtBodyTop("<bodyguard><p>x</p></bodyguard>", "<div>a</div>");

        Assert.Equal("<div>a</div><bodyguard><p>x</p></bodyguard>", result);
    }

    [Fact]
    public void InsertAtBodyTop_NoBodyTagPrepends()
    {
        Assert.Equal("<div>a</div><p>x</p>", HtmlBodyComposer.InsertAtBodyTop("<p>x</p>", "<div>a</div>"));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void InsertAtBodyTop_BlankExistingYieldsMinimalDocument(string? existing)
    {
        Assert.Equal("<html><body><div>a</div></body></html>", HtmlBodyComposer.InsertAtBodyTop(existing, "<div>a</div>"));
    }

    [Theory]
    [InlineData(null, 0)]
    [InlineData("", 0)]
    [InlineData("  ", 0)]
    [InlineData("a@b.example", 1)]
    [InlineData("a@b.example; c@d.example", 2)]
    [InlineData("a@b.example,c@d.example ,, ; e@f.example", 3)]
    public void SplitRecipients_SplitsOnBothSeparatorsAndTrims(string? input, int expectedCount)
    {
        IReadOnlyList<string> result = HtmlBodyComposer.SplitRecipients(input);

        Assert.Equal(expectedCount, result.Count);
        Assert.All(result, r =>
        {
            Assert.False(string.IsNullOrWhiteSpace(r));
            Assert.Equal(r, r.Trim());
        });
    }

    // ================================== D47: inline-image accounting for update_draft

    [Theory]
    [InlineData(null, 0)]
    [InlineData("", 0)]
    [InlineData("<p>no pictures here</p>", 0)]
    [InlineData("<img src=\"cid:image001.png@01D\" alt=logo>", 1)]
    [InlineData("<IMG SRC=x><img/><img\r\nsrc=y>", 3)]
    [InlineData("<image src=x><imgfoo>", 0)]
    [InlineData("<img", 0)]
    public void CountInlineImages_CountsImgElementsOnly(string? html, int expected)
    {
        Assert.Equal(expected, HtmlBodyComposer.CountInlineImages(html));
    }

    [Fact]
    public void CountInlineImages_SeesTheLossThatMadeD47_LinkedImageBecomesAPlaceholderShape()
    {
        // The measured shapes, before and after Word re-serializes a LINKED picture: the
        // <img> is replaced by a VML placeholder AutoShape carrying only the alt text.
        const string linked = "<p><img width=1 height=1 src=\"file:///C:/x/sig_files/logo.png\" alt=logo></p>";
        const string afterRerender = "<p><!--[if gte vml 1]><v:rect id=\"Picture_x0020_1\" alt=\"logo\"/><![endif]--></p>";

        Assert.Equal(1, HtmlBodyComposer.CountInlineImages(linked));
        Assert.Equal(0, HtmlBodyComposer.CountInlineImages(afterRerender));
    }

    [Fact]
    public void CountInlineImages_AnEmbeddedImageSurvivesTheSameRoundTrip()
    {
        const string embedded = "<p><img width=1 height=1 src=\"cid:image001.png@01DD202F.4A7E0EF0\" alt=logo></p>";

        Assert.Equal(1, HtmlBodyComposer.CountInlineImages(embedded));
    }
}
