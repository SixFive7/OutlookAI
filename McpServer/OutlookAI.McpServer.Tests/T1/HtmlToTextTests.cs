using OutlookAI.Core.Text;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>T1: HTML-to-text fallback used by the read tool when Outlook has no plain-text rendering.</summary>
public sealed class HtmlToTextTests
{
    [Fact]
    public void NullAndEmpty_ReturnEmpty()
    {
        Assert.Equal(string.Empty, HtmlToText.Convert(null));
        Assert.Equal(string.Empty, HtmlToText.Convert(string.Empty));
    }

    [Fact]
    public void StripsTags_KeepsText()
    {
        string text = HtmlToText.Convert("<html><body><b>Hello</b> <i>world</i></body></html>");
        Assert.Equal("Hello world", text);
    }

    [Fact]
    public void RemovesScriptStyleAndComments_Entirely()
    {
        string html = "<style>p{color:red}</style><script>alert('x')</script><!-- hidden -->Visible";
        Assert.Equal("Visible", HtmlToText.Convert(html));
    }

    [Fact]
    public void BreakAndParagraphTags_BecomeNewlines()
    {
        string text = HtmlToText.Convert("line1<br>line2<p>para</p>done");
        Assert.Contains("line1\nline2", text.Replace("\r", ""), StringComparison.Ordinal);
        Assert.Contains("para\n", text.Replace("\r", "") + "\n", StringComparison.Ordinal);
    }

    [Fact]
    public void ListItems_GetDashes()
    {
        string text = HtmlToText.Convert("<ul><li>alpha</li><li>beta</li></ul>");
        Assert.Contains("- alpha", text, StringComparison.Ordinal);
        Assert.Contains("- beta", text, StringComparison.Ordinal);
    }

    [Fact]
    public void Entities_AreDecoded()
    {
        string text = HtmlToText.Convert("a&nbsp;&amp;&nbsp;b &lt;tag&gt; &quot;q&quot; &#65; &#x42; &euro;100");
        Assert.Equal("a & b <tag> \"q\" A B €100", text);
    }

    [Fact]
    public void UnknownAndMalformedEntities_PassThrough()
    {
        string text = HtmlToText.Convert("AT&T &unknownentity; &#xZZ; end");
        Assert.Contains("AT&T", text, StringComparison.Ordinal);
        Assert.Contains("&unknownentity;", text, StringComparison.Ordinal);
    }

    [Fact]
    public void BlankLineRuns_CollapseToOneBlankLine()
    {
        string text = HtmlToText.Convert("a<br><br><br><br>b");
        Assert.DoesNotContain("\n\n\n", text.Replace("\r", ""), StringComparison.Ordinal);
    }

    [Fact]
    public void RealisticNewsletterFragment_YieldsReadableText()
    {
        string html = "<html><head><title>x</title><style>a{b:c}</style></head><body>"
            + "<div><h1>Invoice 42</h1><table><tr><td>Amount:</td><td>&euro;12,50</td></tr></table>"
            + "<p>Thanks &amp; regards,<br>The Team</p></div></body></html>";
        string text = HtmlToText.Convert(html);
        Assert.Contains("Invoice 42", text, StringComparison.Ordinal);
        Assert.Contains("€12,50", text, StringComparison.Ordinal);
        Assert.Contains("Thanks & regards,", text, StringComparison.Ordinal);
        Assert.DoesNotContain("<", text, StringComparison.Ordinal);
        Assert.DoesNotContain("a{b:c}", text, StringComparison.Ordinal);
    }
}
