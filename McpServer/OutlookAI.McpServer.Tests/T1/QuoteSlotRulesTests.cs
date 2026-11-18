using OutlookAI.Core.AddIn;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Q209 (a), decided 2026-11-18: a reply's or forward's own text and signature go into the empty paragraphs
/// <c>Reply()</c> and <c>Forward()</c> leave above the quoted message's header block - never into the block, whose
/// first paragraph carries the separator line as its top border. Measured by Q149 on 2026-11-17: every one of the 12
/// replies and forwards the product wrote with a signature had the body and the signature under that line, because
/// both went in at <c>_MailOriginal</c>'s start, inside the bordered paragraph. These are the pure parts of the rule;
/// the T2 tests (<c>LiveSignatureTests</c>, <c>LiveDraftTests</c>) read Word's layout of real drafts.
/// </summary>
public sealed class QuoteSlotRulesTests
{
    [Theory]
    [InlineData("\r")]
    [InlineData(" \r")]
    [InlineData(" \r")]
    [InlineData("\t  \v\r")]
    [InlineData("")]
    public void AParagraphHoldingNothingAPersonWrote_IsBlank(string text)
    {
        Assert.True(QuoteSlotRules.IsBlankParagraphText(text));
    }

    [Theory]
    [InlineData("a\r")]
    [InlineData(" Met vriendelijke groet\r")]
    [InlineData(" .\r")]
    public void AParagraphWithText_IsNotBlank(string text)
    {
        Assert.False(QuoteSlotRules.IsBlankParagraphText(text));
    }

    [Fact]
    public void NoText_IsNotBlank_BecauseItCouldNotBeRead()
    {
        Assert.False(QuoteSlotRules.IsBlankParagraphText(null));
    }

    [Fact]
    public void AFreshReplyOrForward_UsesTheEmptyParagraphAboveTheHeaderBlock()
    {
        // Reply() and Forward() with no inspector: two empty paragraphs, then the bordered header block (C4, 6 of 6).
        Assert.Equal(QuoteSlotAction.UseLastParagraph, QuoteSlotRules.Decide(headerParagraphStart: 2, paragraphAboveIsBlank: true));
    }

    [Fact]
    public void TextRightAboveTheHeaderBlock_GetsAnEmptyParagraphSplitOffItsEnd()
    {
        Assert.Equal(QuoteSlotAction.SplitLastParagraph, QuoteSlotRules.Decide(headerParagraphStart: 40, paragraphAboveIsBlank: false));
    }

    [Fact]
    public void AHeaderBlockAtTheTop_GetsAnUnformattedParagraphPutBeforeIt()
    {
        Assert.Equal(QuoteSlotAction.InsertAtTop, QuoteSlotRules.Decide(headerParagraphStart: 0, paragraphAboveIsBlank: false));
        Assert.Equal(QuoteSlotAction.InsertAtTop, QuoteSlotRules.Decide(headerParagraphStart: 0, paragraphAboveIsBlank: true));
    }
}
