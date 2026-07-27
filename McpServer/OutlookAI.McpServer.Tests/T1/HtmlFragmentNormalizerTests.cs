using System;
using System.Linq;
using OutlookAI.Core.Text;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 pins for the <c>body_html</c> normalizer (batch B - B1, v3.MD D45). Every allowed
/// construct, every dropped construct and every repair of malformed input is fixed here:
/// this markup goes into REAL mail, so "the model closed a tag differently today" must
/// never change what the recipient sees - and must never be able to swallow the signature
/// or the quoted thread.
/// </summary>
public sealed class HtmlFragmentNormalizerTests
{
    private static string Normalize(string html)
    {
        return HtmlFragmentNormalizer.Normalize(html).Html;
    }

    // ---------------------------------------------------------------- allowed constructs

    [Theory]
    [InlineData("<h1>Title</h1>", "<h1>Title</h1>")]
    [InlineData("<h2>A</h2><h3>B</h3><h4>C</h4><h5>D</h5><h6>E</h6>", "<h2>A</h2><h3>B</h3><h4>C</h4><h5>D</h5><h6>E</h6>")]
    [InlineData("<p>Dear Sir,</p>", "<p>Dear Sir,</p>")]
    [InlineData("<p>one<br>two</p>", "<p>one<br>two</p>")]
    [InlineData("<hr>", "<hr>")]
    [InlineData("<strong>bold</strong> and <b>bold</b>", "<strong>bold</strong> and <b>bold</b>")]
    [InlineData("<em>it</em> and <i>it</i> and <u>u</u>", "<em>it</em> and <i>it</i> and <u>u</u>")]
    [InlineData("<ul><li>a</li><li>b</li></ul>", "<ul><li>a</li><li>b</li></ul>")]
    [InlineData("<ol><li>a</li></ol>", "<ol><li>a</li></ol>")]
    [InlineData("<blockquote><p>quoted</p></blockquote>", "<blockquote><p>quoted</p></blockquote>")]
    [InlineData("<sub>x</sub><sup>y</sup><small>z</small>", "<sub>x</sub><sup>y</sup><small>z</small>")]
    [InlineData("<pre><code>x</code></pre>", "<pre><code>x</code></pre>")]
    [InlineData("<dl><dt>t</dt><dd>d</dd></dl>", "<dl><dt>t</dt><dd>d</dd></dl>")]
    public void AllowedElements_SurviveUnchanged(string input, string expected)
    {
        Assert.Equal(expected, Normalize(input));
    }

    [Fact]
    public void SimpleTable_SurvivesWithItsStructureAndCellAttributes()
    {
        const string Input = "<table border=\"1\"><thead><tr><th>Item</th><th>Price</th></tr></thead>"
            + "<tbody><tr><td colspan=\"2\" align=\"right\">Total</td></tr></tbody></table>";

        Assert.Equal(Input, Normalize(Input));
    }

    [Fact]
    public void Links_KeepSafeSchemesOnly()
    {
        Assert.Equal("<a href=\"https://example.com/x\">site</a>", Normalize("<a href=\"https://example.com/x\">site</a>"));
        Assert.Equal("<a href=\"mailto:a@b.example\">mail</a>", Normalize("<a href=\"mailto:a@b.example\">mail</a>"));
        Assert.Equal("<a href=\"tel:+3110\">call</a>", Normalize("<a href=\"tel:+3110\">call</a>"));
    }

    [Fact]
    public void InlineStyle_IsKept_BecauseFormattingIsThePoint()
    {
        Assert.Equal(
            "<p style=\"color: #123456; font-weight: bold\">x</p>",
            Normalize("<p style=\"color: #123456; font-weight: bold\">x</p>"));
        // Declarations are passed through as written - the normalizer filters, it does
        // not restyle.
        Assert.Equal(
            "<span style=\"font-family:Calibri\">x</span>",
            Normalize("<span style=\"font-family:Calibri\">x</span>"));
    }

    [Fact]
    public void EntitiesAndAmpersands_AreLeftValid()
    {
        Assert.Equal("a &amp; b", Normalize("a & b"));
        Assert.Equal("a &amp; b", Normalize("a &amp; b"));
        Assert.Equal("&#8364;5 &#x20AC;5 &nbsp;", Normalize("&#8364;5 &#x20AC;5 &nbsp;"));
    }

    // ---------------------------------------------------------------- dropped constructs

    [Fact]
    public void Wrappers_AreStripped_AndHeadContentGoesWithThem()
    {
        HtmlNormalizationResult result = HtmlFragmentNormalizer.Normalize(
            "<html><head><title>t</title><meta charset=\"utf-8\"></head><body><p>Hi</p></body></html>");

        Assert.Equal("<p>Hi</p>", result.Html);
        Assert.Contains(result.Adjustments, a => a.Contains("<html>", StringComparison.Ordinal));
        Assert.Contains(result.Adjustments, a => a.Contains("<body>", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("<script>alert('x')</script><p>Hi</p>", "<p>Hi</p>", "script")]
    [InlineData("<style>p{color:red}</style><p>Hi</p>", "<p>Hi</p>", "style")]
    [InlineData("<p>Hi</p><iframe src=\"http://evil\">fallback</iframe>", "<p>Hi</p>", "iframe")]
    [InlineData("<object data=\"x\">o</object><p>Hi</p>", "<p>Hi</p>", "object")]
    [InlineData("<p>Hi</p><link rel=\"stylesheet\" href=\"http://x\">", "<p>Hi</p>", "link")]
    public void CodeAndRemoteLoadingElements_AreDroppedWithTheirContent(string input, string expected, string reported)
    {
        HtmlNormalizationResult result = HtmlFragmentNormalizer.Normalize(input);

        Assert.Equal(expected, result.Html);
        Assert.Contains(result.Adjustments, a => a.Contains("<" + reported + ">", StringComparison.Ordinal));
    }

    [Fact]
    public void Images_AreRemoved_TheyWouldMeanARemoteOrAttachedResource()
    {
        HtmlNormalizationResult result = HtmlFragmentNormalizer.Normalize("<p>See <img src=\"https://x/y.png\"> here</p>");

        Assert.Equal("<p>See  here</p>", result.Html);
        Assert.Contains(result.Adjustments, a => a.Contains("<img>", StringComparison.Ordinal));
    }

    [Fact]
    public void UnsupportedTags_AreUnwrapped_ContentIsNeverLost()
    {
        HtmlNormalizationResult result = HtmlFragmentNormalizer.Normalize("<font size=\"3\">kept text</font><marquee>also kept</marquee>");

        Assert.Equal("kept text also kept".Replace(" ", string.Empty), result.Html.Replace(" ", string.Empty));
        Assert.Contains(result.Adjustments, a => a.Contains("unwrapped unsupported <font>", StringComparison.Ordinal));
        Assert.Contains(result.Adjustments, a => a.Contains("unwrapped unsupported <marquee>", StringComparison.Ordinal));
    }

    [Fact]
    public void EventHandlers_AreRemoved()
    {
        HtmlNormalizationResult result = HtmlFragmentNormalizer.Normalize("<p onclick=\"steal()\" onmouseover=\"x\">Hi</p>");

        Assert.Equal("<p>Hi</p>", result.Html);
        Assert.Contains(result.Adjustments, a => a.Contains("onclick", StringComparison.Ordinal));
    }

    [Fact]
    public void IdAndName_AreRemoved_SoOutlooksRegionMarkersCannotBeForged()
    {
        // A body carrying id="_MailAutoSig" would redraw the draft/signature boundary the
        // whole compose model (and the add-in's sidebar) depends on - see v3.MD D44.
        HtmlNormalizationResult result = HtmlFragmentNormalizer.Normalize(
            "<div id=\"_MailAutoSig\"><a name=\"_MailOriginal\">x</a></div>");

        Assert.DoesNotContain("_MailAutoSig", result.Html, StringComparison.Ordinal);
        Assert.DoesNotContain("_MailOriginal", result.Html, StringComparison.Ordinal);
        Assert.Equal("<div><a>x</a></div>", result.Html);
    }

    [Fact]
    public void DangerousHrefSchemes_LoseTheLinkButKeepTheText()
    {
        HtmlNormalizationResult result = HtmlFragmentNormalizer.Normalize("<a href=\"javascript:steal()\">click</a>");

        Assert.Equal("<a>click</a>", result.Html);
        Assert.Contains(result.Adjustments, a => a.Contains("unsupported URL scheme", StringComparison.Ordinal));
    }

    [Fact]
    public void StyleDeclarationsThatFetchOrExecute_AreDropped_TheRestOfTheStyleSurvives()
    {
        HtmlNormalizationResult result = HtmlFragmentNormalizer.Normalize(
            "<p style=\"color:red; background:url(http://evil/x.png); font-size:12pt\">x</p>");

        Assert.Equal("<p style=\"color:red; font-size:12pt\">x</p>", result.Html);
        Assert.Contains(result.Adjustments, a => a.Contains("loads or executes", StringComparison.Ordinal));
    }

    [Fact]
    public void ClassAttributesAndComments_AreRemoved()
    {
        HtmlNormalizationResult result = HtmlFragmentNormalizer.Normalize("<!-- note --><p class=\"MsoNormal\">Hi</p>");

        Assert.Equal("<p>Hi</p>", result.Html);
        Assert.Contains(result.Adjustments, a => a.Contains("comment", StringComparison.Ordinal));
        Assert.Contains(result.Adjustments, a => a.Contains("class", StringComparison.Ordinal));
    }

    // ---------------------------------------------------------------- malformed input is REPAIRED

    [Fact]
    public void UnclosedTags_AreClosed_SoTheyCannotSwallowTheSignature()
    {
        HtmlNormalizationResult result = HtmlFragmentNormalizer.Normalize("<p><b>bold text");

        Assert.Equal("<p><b>bold text</b></p>", result.Html);
        Assert.Contains(result.Adjustments, a => a.Contains("unclosed <b>", StringComparison.Ordinal));
    }

    [Fact]
    public void MisNestedTags_AreUnwound()
    {
        Assert.Equal("<b>a<i>b</i></b>", Normalize("<b>a<i>b</b></i>"));
    }

    [Fact]
    public void StrayEndTags_AreDropped()
    {
        HtmlNormalizationResult result = HtmlFragmentNormalizer.Normalize("</div><p>Hi</p></span>");

        Assert.Equal("<p>Hi</p>", result.Html);
        Assert.Contains(result.Adjustments, a => a.Contains("stray </div>", StringComparison.Ordinal));
    }

    [Fact]
    public void StrayLessThan_IsEscaped_NotTreatedAsATag()
    {
        HtmlNormalizationResult result = HtmlFragmentNormalizer.Normalize("<p>5 < 6 and 10 <= 20</p>");

        Assert.Equal("<p>5 &lt; 6 and 10 &lt;= 20</p>", result.Html);
        Assert.Contains(result.Adjustments, a => a.Contains("stray", StringComparison.Ordinal));
    }

    [Fact]
    public void LessThanFollowedByALetter_IsATag_LikeInEveryHtmlParser_AndStaysWellFormed()
    {
        // "a<b" reads as an opening <b> to any HTML parser; the documented rule is the
        // browser rule. What matters here is that the output cannot be malformed and
        // cannot leak past the body: the invented element is closed for us.
        Assert.Equal("<p>a<b></b></p>", Normalize("<p>a<b</p>"));
    }

    [Theory]
    [InlineData("<p>a<p>b", "<p>a</p><p>b</p>")]
    [InlineData("<ul><li>a<li>b</ul>", "<ul><li>a</li><li>b</li></ul>")]
    [InlineData("<table><tr><td>a<td>b</tr></table>", "<table><tr><td>a</td><td>b</td></tr></table>")]
    [InlineData("<p>a<h1>b</h1>", "<p>a</p><h1>b</h1>")]
    public void ImplicitCloses_MatchWhatABrowserWouldDo(string input, string expected)
    {
        Assert.Equal(expected, Normalize(input));
    }

    [Fact]
    public void StrayListAndTableParts_GetTheAncestorsTheyNeed()
    {
        Assert.Equal("<ul><li>a</li></ul>", Normalize("<li>a</li>"));
        Assert.Equal("<table><tr><td>a</td></tr></table>", Normalize("<tr><td>a</td></tr>"));
        Assert.Equal("<table><tr><td>a</td></tr></table>", Normalize("<td>a</td>"));
    }

    [Fact]
    public void LooseTextInsideAListOrTable_IsGivenACellOrItem()
    {
        Assert.Equal("<ul><li>loose</li></ul>", Normalize("<ul>loose</ul>"));
        Assert.Equal("<table><tr><td>loose</td></tr></table>", Normalize("<table>loose</table>"));
    }

    [Fact]
    public void SelfClosingForms_AreHandled()
    {
        Assert.Equal("<p>a<br>b</p>", Normalize("<p>a<br/>b</p>"));
        Assert.Equal("<div></div>", Normalize("<div/>"));
    }

    [Fact]
    public void AttributeQuotingVariants_AreReSerializedSafely()
    {
        Assert.Equal("<td colspan=\"2\">x</td>", Normalize("<table><tr><td colspan=2>x</td></tr></table>")
            .Replace("<table><tr>", string.Empty).Replace("</tr></table>", string.Empty));
        Assert.Equal("<a href=\"https://x/?a=1&amp;b=2\">l</a>", Normalize("<a href='https://x/?a=1&b=2'>l</a>"));
    }

    // ---------------------------------------------------------------- reporting + limits

    [Fact]
    public void NothingChanged_MeansNoAdjustmentsToReport()
    {
        Assert.Empty(HtmlFragmentNormalizer.Normalize("<p>Hello</p>").Adjustments);
    }

    [Fact]
    public void RepeatedAdjustments_AreAggregatedWithACount()
    {
        HtmlNormalizationResult result = HtmlFragmentNormalizer.Normalize("<p id=\"a\">1</p><p id=\"b\">2</p><p id=\"c\">3</p>");

        string reported = Assert.Single(result.Adjustments, a => a.Contains("id attribute", StringComparison.Ordinal));
        Assert.Contains("(x3)", reported, StringComparison.Ordinal);
    }

    [Fact]
    public void AdjustmentListIsCapped_AndSaysHowManyItOmitted()
    {
        string input = string.Concat(Enumerable.Range(0, HtmlFragmentNormalizer.MaxReportedAdjustments + 5)
            .Select(i => "<x" + i + ">t</x" + i + ">"));

        HtmlNormalizationResult result = HtmlFragmentNormalizer.Normalize(input);

        Assert.Equal(HtmlFragmentNormalizer.MaxReportedAdjustments + 1, result.Adjustments.Count);
        Assert.Contains("further adjustment", result.Adjustments[result.Adjustments.Count - 1], StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("<script>alert(1)</script>")]
    [InlineData("<p>   </p>")]
    public void InputWithNothingVisible_IsReportedAsEmpty(string input)
    {
        Assert.False(HtmlFragmentNormalizer.Normalize(input).HasVisibleContent);
    }

    [Theory]
    [InlineData("<p>x</p>")]
    [InlineData("<hr>")]
    [InlineData("<table><tr><td></td></tr></table>")]
    public void InputWithSomethingVisible_IsReportedAsSuch(string input)
    {
        Assert.True(HtmlFragmentNormalizer.Normalize(input).HasVisibleContent);
    }

    [Fact]
    public void OversizeInput_IsRejectedBeforeAnythingElse()
    {
        string huge = new string('a', HtmlFragmentNormalizer.MaxInputChars + 1);

        ArgumentException ex = Assert.Throws<ArgumentException>(() => HtmlFragmentNormalizer.Normalize(huge));

        Assert.Contains("body_html", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void NullInput_IsTreatedAsEmpty_NotAsACrash()
    {
        Assert.False(HtmlFragmentNormalizer.Normalize(null).HasVisibleContent);
    }

    [Fact]
    public void AFullFormalLetter_SurvivesEndToEnd()
    {
        const string Letter = "<h1>Notice of termination</h1>"
            + "<p>Dear <strong>Mrs. Jansen</strong>,</p>"
            + "<p>Per the agreement we confirm the following:</p>"
            + "<ul><li>Contract <em>XY-12</em> ends on 1 September.</li><li>Final invoice follows.</li></ul>"
            + "<table><thead><tr><th>Item</th><th>Amount</th></tr></thead>"
            + "<tbody><tr><td>Licence</td><td>EUR 120</td></tr></tbody></table>"
            + "<p>See <a href=\"https://example.com/terms\">the terms</a>.</p>"
            + "<p style=\"color:#555\">Kind regards,</p>";

        HtmlNormalizationResult result = HtmlFragmentNormalizer.Normalize(Letter);

        Assert.Equal(Letter, result.Html);
        Assert.Empty(result.Adjustments);
        Assert.True(result.HasVisibleContent);
    }
}
