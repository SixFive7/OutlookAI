using System.Globalization;
using Xunit;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Q209 (a), decided 2026-11-18: a reply's or forward's own text and signature sit in the empty paragraphs
/// <c>Reply()</c> and <c>Forward()</c> leave ABOVE the quoted message's header block - never inside it, under its
/// separator line. Measured by Q149 on 2026-11-17: all 12 replies and forwards the product wrote with a signature
/// had their first eleven paragraphs laid out with the header block's top border. Read off Word's own layout of the
/// draft (<see cref="LiveOutlookTestMailer.ReadDraftLook"/>), because that border is what a person sees and the HTML
/// alone does not say where Word draws it.
/// </summary>
internal static class QuotedDraftLayout
{
    /// <summary>
    /// Asserts the layout: the FIRST paragraph has no top border; the paragraph holding <paramref name="agentText"/>,
    /// and the one holding <paramref name="signatureText"/> when given, come before the first bordered paragraph -
    /// the header block's separator line, which must still be there, or the check would prove nothing.
    /// </summary>
    internal static void AssertTextAboveTheHeaderBlock(
        IReadOnlyList<WordParagraphLook> look,
        string agentText,
        string? signatureText,
        string what,
        Action<string> log)
    {
        Assert.NotEmpty(look);
        int header = IndexOf(look, p => p.TopBorder > 0);
        log(what + " layout: " + look.Count.ToString(CultureInfo.InvariantCulture) + " paragraphs; first top border at p"
            + (header + 1).ToString(CultureInfo.InvariantCulture) + "; "
            + string.Join(" | ", look.Take(Math.Max(1, Math.Min(look.Count, header + 2))).Select((p, i) => "p" + (i + 1).ToString(CultureInfo.InvariantCulture) + " " + p.Describe())));

        Assert.True(
            look[0].TopBorder == 0,
            what + ": the first paragraph has a top border (line style " + look[0].TopBorder.ToString(CultureInfo.InvariantCulture)
            + ") - the text went under the quoted message's separator line, inside its header block (Q209 (a))");
        Assert.True(
            header > 0,
            what + ": no paragraph has a top border, so the quoted message's header block was not found and this check proves nothing");

        int agent = IndexOf(look, p => p.Text.Contains(Normalize(agentText), StringComparison.Ordinal));
        Assert.True(agent >= 0, what + ": the agent's text is in no paragraph");
        Assert.True(agent < header, what + ": the agent's text (p" + (agent + 1).ToString(CultureInfo.InvariantCulture)
            + ") is not above the header block (p" + (header + 1).ToString(CultureInfo.InvariantCulture) + ")");
        if (signatureText != null)
        {
            int signature = IndexOf(look, p => p.Text.Contains(Normalize(signatureText), StringComparison.Ordinal));
            Assert.True(signature > agent, what + ": the signature does not follow the agent's text");
            Assert.True(signature < header, what + ": the signature (p" + (signature + 1).ToString(CultureInfo.InvariantCulture)
                + ") is not above the header block (p" + (header + 1).ToString(CultureInfo.InvariantCulture) + ")");
        }

        for (int i = 0; i < header; i++)
        {
            Assert.True(look[i].TopBorder == 0, what + ": p" + (i + 1).ToString(CultureInfo.InvariantCulture) + " above the header block has a top border");
        }
    }

    private static string Normalize(string text)
    {
        return Q149WordLook.NormalizeText(text);
    }

    private static int IndexOf(IReadOnlyList<WordParagraphLook> look, Func<WordParagraphLook, bool> match)
    {
        for (int i = 0; i < look.Count; i++)
        {
            if (match(look[i]))
            {
                return i;
            }
        }

        return -1;
    }
}
