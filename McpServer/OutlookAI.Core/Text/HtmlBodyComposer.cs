using System;
using System.Text;

namespace OutlookAI.Core.Text
{
    /// <summary>
    /// HTML composition helpers for the draft tools (v3.MD sections 3/8 L4). The write
    /// path sets HTMLBody exactly ONCE per draft (section 12: .Body vs .HTMLBody -
    /// last one set wins) with the agent's text placed at the TOP of the existing body,
    /// i.e. ABOVE the signature Outlook injected and ABOVE the quoted reply/forward
    /// history. Pure logic - T1-tested.
    /// </summary>
    public static class HtmlBodyComposer
    {
        /// <summary>
        /// Converts agent-authored plain text into an HTML fragment: HTML-escapes the
        /// text and turns line breaks into &lt;br&gt; tags, wrapped in a div.
        /// </summary>
        public static string ToHtmlFragment(string plainText)
        {
            if (plainText == null)
            {
                throw new ArgumentNullException(nameof(plainText));
            }

            StringBuilder sb = new StringBuilder(plainText.Length + 32);
            sb.Append("<div>");
            for (int i = 0; i < plainText.Length; i++)
            {
                char c = plainText[i];
                switch (c)
                {
                    case '&':
                        sb.Append("&amp;");
                        break;
                    case '<':
                        sb.Append("&lt;");
                        break;
                    case '>':
                        sb.Append("&gt;");
                        break;
                    case '"':
                        sb.Append("&quot;");
                        break;
                    case '\r':
                        if (i + 1 < plainText.Length && plainText[i + 1] == '\n')
                        {
                            i++;
                        }

                        sb.Append("<br>");
                        break;
                    case '\n':
                        sb.Append("<br>");
                        break;
                    default:
                        sb.Append(c);
                        break;
                }
            }

            sb.Append("</div>");
            return sb.ToString();
        }

        /// <summary>
        /// Inserts <paramref name="fragmentHtml"/> at the top of an HTML document's
        /// body: right after the opening &lt;body ...&gt; tag when one exists (case
        /// insensitive, attributes tolerated), else prepended to the raw content. A
        /// null/blank existing body yields a minimal HTML document around the fragment.
        /// </summary>
        public static string InsertAtBodyTop(string? existingHtml, string fragmentHtml)
        {
            if (fragmentHtml == null)
            {
                throw new ArgumentNullException(nameof(fragmentHtml));
            }

            if (string.IsNullOrWhiteSpace(existingHtml))
            {
                return "<html><body>" + fragmentHtml + "</body></html>";
            }

            string html = existingHtml!;
            int bodyOpen = FindBodyTagStart(html);
            if (bodyOpen >= 0)
            {
                int tagEnd = html.IndexOf('>', bodyOpen);
                if (tagEnd >= 0)
                {
                    return html.Substring(0, tagEnd + 1) + fragmentHtml + html.Substring(tagEnd + 1);
                }
            }

            return fragmentHtml + html;
        }

        private static int FindBodyTagStart(string html)
        {
            // First occurrence of "<body" followed by whitespace or '>' - enough for
            // Outlook-generated HTML (its body tag carries plain attributes only).
            int from = 0;
            while (true)
            {
                int idx = html.IndexOf("<body", from, StringComparison.OrdinalIgnoreCase);
                if (idx < 0)
                {
                    return -1;
                }

                int after = idx + 5;
                if (after >= html.Length)
                {
                    return -1;
                }

                char next = html[after];
                if (next == '>' || char.IsWhiteSpace(next))
                {
                    return idx;
                }

                from = idx + 1;
            }
        }

        /// <summary>
        /// Splits an agent-provided recipient list on ';' and ',' into trimmed,
        /// non-empty addresses (the MCP draft tools accept both separators).
        /// </summary>
        public static System.Collections.Generic.IReadOnlyList<string> SplitRecipients(string? recipients)
        {
            if (string.IsNullOrWhiteSpace(recipients))
            {
                return Array.Empty<string>();
            }

            string[] raw = recipients!.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries);
            System.Collections.Generic.List<string> result = new System.Collections.Generic.List<string>(raw.Length);
            foreach (string entry in raw)
            {
                string trimmed = entry.Trim();
                if (trimmed.Length > 0)
                {
                    result.Add(trimmed);
                }
            }

            return result;
        }
    }
}
