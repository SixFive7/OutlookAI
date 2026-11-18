namespace OutlookAI.Core.AddIn
{
    /// <summary>How the empty paragraph above a quoted message's header block is had (<see cref="QuoteSlotRules.Decide"/>).</summary>
    public enum QuoteSlotAction
    {
        /// <summary>The paragraph directly above the header block is empty: the text goes at its start, and it stays.</summary>
        UseLastParagraph = 0,

        /// <summary>The paragraph above holds text: an empty paragraph is split off its end, with its formatting.</summary>
        SplitLastParagraph = 1,

        /// <summary>The header block is the document's first paragraph: an empty one is put before it, unformatted.</summary>
        InsertAtTop = 2,
    }

    /// <summary>
    /// Q209 (a), decided 2026-11-18: a reply's or forward's own text and signature go into the empty paragraphs
    /// <c>Reply()</c> and <c>Forward()</c> leave above the quoted message's header block, never into the block itself,
    /// whose first paragraph carries the separator line as a top border. These are the pure parts of that rule, so T1
    /// pins them without Word; <c>ComposeDocument</c> applies them to the document.
    /// </summary>
    public static class QuoteSlotRules
    {
        /// <summary>
        /// True when a paragraph's text - as Word's <c>Range.Text</c> gives it, its paragraph mark included - holds
        /// nothing a person wrote: only whitespace, non-breaking spaces and paragraph or line marks. Outlook's own empty
        /// paragraph (<c>&lt;o:p&gt;&amp;nbsp;&lt;/o:p&gt;</c>) is one.
        /// </summary>
        public static bool IsBlankParagraphText(string? text)
        {
            if (text == null)
            {
                return false;
            }

            foreach (char c in text)
            {
                if (!char.IsWhiteSpace(c) && c != ' ' && c != '​' && c != '﻿')
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// How to have the empty paragraph above the header block, given where the block's first paragraph starts and
        /// whether the paragraph right above it is blank.
        /// </summary>
        public static QuoteSlotAction Decide(int headerParagraphStart, bool paragraphAboveIsBlank)
        {
            if (headerParagraphStart <= 0)
            {
                return QuoteSlotAction.InsertAtTop;
            }

            return paragraphAboveIsBlank ? QuoteSlotAction.UseLastParagraph : QuoteSlotAction.SplitLastParagraph;
        }
    }
}
