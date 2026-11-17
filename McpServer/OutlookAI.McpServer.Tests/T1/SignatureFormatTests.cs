using System.Text;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The three renditions read and written as they are stored (Q159: "without breaking or converting
/// anything"): encodings kept, RTF pictures elided to tokens and restored byte for byte, text replaced
/// in place in each rendition's own spelling - and never across formatting or into markup.
/// </summary>
public sealed class SignatureFormatTests
{
    private static readonly Encoding W1252 = CodePage(1252);

    // ------------------------------------------------------------------ encodings

    [Fact]
    public void Html_InTheCharsetItDeclares_RoundTripsByteForByte()
    {
        string html = "<html><head><meta http-equiv=Content-Type content=\"text/html; charset=windows-1252\"></head><body>Café</body></html>";
        byte[] stored = W1252.GetBytes(html);

        DecodedSignatureText decoded = SignatureText.DecodeHtml(stored);

        Assert.Equal(html, decoded.Text);
        Assert.Equal(1252, decoded.CodePage);
        Assert.Equal(stored, SignatureText.EncodeHtml(decoded.Text, decoded));
    }

    [Fact]
    public void Html_ACharacterItsCharsetCannotHold_BecomesAReference_TheDeclarationStaysTrue()
    {
        string html = "<html><head><meta charset=\"windows-1252\"></head><body>Café 中 \U0001F600</body></html>";

        byte[] stored = SignatureText.EncodeHtml(html, null);
        string back = W1252.GetString(stored);

        Assert.Contains("Café &#20013; &#128512;", back, StringComparison.Ordinal);
        Assert.Contains("charset=\"windows-1252\"", back, StringComparison.Ordinal);
    }

    [Fact]
    public void Html_WithAMark_KeepsIt_WithoutADeclaration_GetsUtf8OnlyWhenItNeedsIt()
    {
        byte[] marked = new byte[] { 0xEF, 0xBB, 0xBF }.Concat(Encoding.UTF8.GetBytes("<p>é</p>")).ToArray();
        DecodedSignatureText decoded = SignatureText.DecodeHtml(marked);
        Assert.True(decoded.HasBom);
        Assert.Equal(marked, SignatureText.EncodeHtml(decoded.Text, decoded));

        Assert.Equal(Encoding.ASCII.GetBytes("<p>plain</p>"), SignatureText.EncodeHtml("<p>plain</p>", null));
        string declared = Encoding.UTF8.GetString(SignatureText.EncodeHtml("<html><head></head><body>é</body></html>", null));
        Assert.Contains("charset=utf-8", declared, StringComparison.Ordinal);
        Assert.Contains("é", declared, StringComparison.Ordinal);
    }

    [Fact]
    public void Text_KeepsItsEncoding_WhileTheTextFits_ThenMovesToOutlooksOwn()
    {
        DecodedSignatureText ansi = SignatureText.DecodePlain(W1252.GetBytes("Café"));
        Assert.Equal(W1252.GetBytes("Café bar"), SignatureText.EncodeText("Café bar", ansi));

        byte[] wider = SignatureText.EncodeText("中", ansi);
        Assert.Equal(new byte[] { 0xFF, 0xFE, 0x2D, 0x4E }, wider);

        DecodedSignatureText utf16 = SignatureText.DecodePlain(new byte[] { 0xFF, 0xFE, 0x61, 0x00 });
        Assert.Equal(new byte[] { 0xFF, 0xFE, 0x62, 0x00 }, SignatureText.EncodeText("b", utf16));

        DecodedSignatureText utf8 = SignatureText.DecodePlain(Encoding.UTF8.GetBytes("é"));
        Assert.Equal("utf-8", utf8.EncodingName);
        Assert.Equal(Encoding.UTF8.GetBytes("è"), SignatureText.EncodeText("è", utf8));
    }

    [Fact]
    public void Rtf_IsWrittenSevenBit_ExceptBinaryData_WhichSurvivesAsItIs()
    {
        string rtf = "{\\rtf1\\ansi\\ansicpg1252 Café 中{\\pict\\bin3 " + "ÿ\u0000\u0080" + "}}";

        byte[] stored = SignatureText.EncodeRtf(rtf);
        string back = SignatureText.Latin1.GetString(stored);

        Assert.Contains("Caf\\'e9 \\u20013?", back, StringComparison.Ordinal);
        Assert.EndsWith("\\bin3 ÿ\u0000\u0080}}", back, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ RTF pictures

    /// <summary>The construct Word writes for an inline picture: the PNG for current readers, a WMF for old ones.</summary>
    private static string WordPicture(byte[] png, string wmfHex = "0100090000")
    {
        return "{\\*\\shppict{\\pict{\\*\\picprop\\shplid1025{\\sp{\\sn shapeType}{\\sv 75}}}\\picscalex100\\picscaley100"
            + "\\picw2117\\pich1058\\picwgoal1200\\pichgoal600\\pngblip\\bliptag-1\r\n" + TestPictures.Hex(png) + "}}"
            + "{\\nonshppict{\\pict\\picscalex100\\picscaley100\\picw2117\\pich1058\\picwgoal1200\\pichgoal600\\wmetafile8\r\n" + wmfHex + "}}";
    }

    [Fact]
    public void RtfPictures_AreFoundWithTheirFallback_AndElidedToTokens_ThatExpandBackByteForByte()
    {
        byte[] png = TestPictures.Png(80, 40, 5);
        string rtf = "{\\rtf1\\ansi\\ansicpg1252{\\fonttbl{\\f0 Calibri;}}\\f0 Jan Modaal\\par " + WordPicture(png) + "\\par Bye}";

        RtfPicture picture = Assert.Single(SignatureRtf.FindPictures(rtf));
        Assert.Equal("png", picture.Kind);
        Assert.Equal(png.Length, picture.ByteCount);
        Assert.Equal(SignatureRtf.Sha256Hex(png), picture.Sha256);
        Assert.Equal(1200, picture.WidthTwips);
        Assert.True(picture.HasFallback);

        string elided = SignatureRtf.Elide(rtf);
        Assert.Equal("{\\rtf1\\ansi\\ansicpg1252{\\fonttbl{\\f0 Calibri;}}\\f0 Jan Modaal\\par [[picture:1]]\\par Bye}", elided);
        Assert.Equal(rtf, SignatureRtf.Expand(elided, rtf, _ => null));
        Assert.Equal(rtf.Replace(WordPicture(png), string.Empty), SignatureRtf.Expand(elided.Replace("[[picture:1]]", string.Empty), rtf, _ => null));
    }

    [Fact]
    public void RtfPictures_ANewOne_IsAPngOrJpegBlip_SizedAsAsked()
    {
        byte[] jpeg = TestPictures.Jpeg(120, 60);

        string group = SignatureRtf.PictureGroup(jpeg, null, null);
        Assert.Contains("\\jpegblip", group, StringComparison.Ordinal);
        Assert.Contains("\\picwgoal1800\\pichgoal900", group, StringComparison.Ordinal);

        string expanded = SignatureRtf.Expand("{\\rtf1 [[picture-file:logo.jpg|60x30]]}", string.Empty,
            file => file == "logo.jpg" ? jpeg : null);
        Assert.Contains("\\picwgoal900\\pichgoal450", expanded, StringComparison.Ordinal);

        Assert.Throws<ArgumentException>(() => SignatureRtf.PictureGroup(Encoding.ASCII.GetBytes("GIF89a......"), null, null));
        Assert.Throws<ArgumentException>(() => SignatureRtf.Expand("{\\rtf1 [[picture-file:nope.png]]}", string.Empty, _ => null));
    }

    // ------------------------------------------------------------------ RTF text

    [Fact]
    public void RtfText_IsWhatAReaderShows()
    {
        string rtf = "{\\rtf1\\ansi\\ansicpg1252\\uc1{\\fonttbl{\\f0 Calibri;}}{\\colortbl;\\red0\\green0\\blue0;}"
            + "{\\*\\generator Riched20;}\\f0 Caf\\'e9\\u8211\\'96x\\par Tel\\tab 1\\line"
            + "{\\field{\\*\\fldinst HYPERLINK \"https://example.com\"}{\\fldrslt example}}" + WordPicture(TestPictures.Png(1, 1)) + "}";

        Assert.Equal("Café–x\nTel\t1\nexample", SignatureRtf.ToPlainText(rtf));
    }

    [Fact]
    public void RtfReplace_InsideARun_InTheDocumentsOwnEscaping_AcrossSourceLineBreaks()
    {
        string rtf = "{\\rtf1\\ansi\\ansicpg1252 Jan Modaal\\par Caf\\'e9 de\r\nBuren\\par Tel 010 123}";

        SignatureTextReplacement cafe = SignatureRtf.ReplaceText(rtf, "Café de Buren", "Bar é中");
        Assert.Equal(1, cafe.Replaced);
        Assert.Equal("{\\rtf1\\ansi\\ansicpg1252 Jan Modaal\\par Bar \\'e9\\u20013?\\par Tel 010 123}", cafe.Content);

        SignatureTextReplacement unicode = SignatureRtf.ReplaceText("{\\rtf1 A\\u8211?B}", "A–B", "AB");
        Assert.Equal(1, unicode.Replaced);
        Assert.Equal("{\\rtf1 AB}", unicode.Content);
    }

    [Fact]
    public void RtfReplace_NeverAcrossFormatting_NorInsideTablesOrPictures()
    {
        SignatureTextReplacement split = SignatureRtf.ReplaceText("{\\rtf1 Jan {\\b Modaal}}", "Jan Modaal", "X");
        Assert.Equal(0, split.Replaced);
        Assert.True(split.FoundButSplit);

        SignatureTextReplacement fonts = SignatureRtf.ReplaceText("{\\rtf1{\\fonttbl{\\f0 Calibri;}} text}", "Calibri", "Arial");
        Assert.Equal(0, fonts.Replaced);
        Assert.False(fonts.FoundButSplit);

        SignatureTextReplacement control = SignatureRtf.ReplaceText("{\\rtf1\\par par}", "par", "x");
        Assert.Equal(1, control.Replaced);
        Assert.Equal("{\\rtf1\\par x}", control.Content);
    }

    // ------------------------------------------------------------------ HTML

    [Fact]
    public void HtmlReferences_AreTheFilesOfTheSignaturesOwnFolder_AsWordEncodesThem()
    {
        string html = "<link rel=File-List href=\"My%20Sig_files/filelist.xml\"><v:imagedata src=\"My%20Sig_files/image001.png\" o:title=\"\"/>"
            + "<img width=120 height=40 src=\"My%20Sig_files/image002.png\" v:shapes=\"Picture_x0020_1\">"
            + "<a href=\"https://example.com\">x</a><img src=\"logo.png\"><td background='My Sig_files\\bg.gif'>";

        IReadOnlyList<SignatureHtmlReference> references = SignatureHtml.References(html, "My Sig");

        Assert.Equal(new[] { "filelist.xml", "image001.png", "image002.png", null, null, "bg.gif" }, references.Select(r => r.FileName));
        Assert.True(SignatureHtml.IsStrayLocalReference(references[4]));
        Assert.False(SignatureHtml.IsStrayLocalReference(references[3]));
        Assert.Equal("My%20Sig_files/logo.png", SignatureHtml.ReferenceFor("My Sig", "logo.png"));
        Assert.Null(SignatureHtml.ResourceFileOf("My%20Sig_files/../escape.png", "My Sig"));
    }

    [Theory]
    [InlineData("<p>Caf&eacute; de Buren</p>")]
    [InlineData("<p>Caf&#233; de Buren</p>")]
    [InlineData("<p>Caf&#xE9; de Buren</p>")]
    [InlineData("<p>Café&nbsp;de\r\nBuren</p>")]
    public void HtmlReplace_MatchesTheTextHoweverHtmlSpellsIt(string html)
    {
        SignatureTextReplacement replaced = SignatureHtml.ReplaceText(html, "Café de Buren", "A & B");

        Assert.Equal(1, replaced.Replaced);
        Assert.Equal("<p>A &amp; B</p>", replaced.Content);
    }

    [Fact]
    public void HtmlReplace_NeverInMarkupCommentsOrStyles_AndSplitByMarkupIsReported()
    {
        string html = "<html><head><style>p.Jan { color: red }</style><title>Jan</title></head>"
            + "<body><!--[if gte mso 9]>Jan<![endif]--><p class=\"Jan\" title=\"Jan\">Jan</p></body></html>";

        SignatureTextReplacement replaced = SignatureHtml.ReplaceText(html, "Jan", "Piet");

        Assert.Equal(1, replaced.Replaced);
        Assert.Equal(html.Replace(">Jan</p>", ">Piet</p>"), replaced.Content);

        SignatureTextReplacement split = SignatureHtml.ReplaceText("<p>Jan <b>Modaal</b></p>", "Jan Modaal", "X");
        Assert.Equal(0, split.Replaced);
        Assert.True(split.FoundButSplit);
    }

    // ------------------------------------------------------------------ pictures

    [Fact]
    public void Pictures_SizesFromTheirHeaders()
    {
        Assert.True(SignaturePictures.TryReadSize(TestPictures.Png(80, 40), out int pw, out int ph));
        Assert.Equal((80, 40), (pw, ph));
        Assert.True(SignaturePictures.TryReadSize(TestPictures.Jpeg(300, 50), out int jw, out int jh));
        Assert.Equal((300, 50), (jw, jh));
        Assert.False(SignaturePictures.TryReadSize(new byte[] { 1, 2, 3 }, out _, out _));
        Assert.True(SignaturePictures.MatchesExtension(TestPictures.Png(1, 1), ".PNG"));
        Assert.False(SignaturePictures.MatchesExtension(TestPictures.Png(1, 1), ".jpg"));
    }

    private static Encoding CodePage(int codePage)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        return Encoding.GetEncoding(codePage);
    }
}
