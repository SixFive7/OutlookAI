using System;
using System.Text;
using System.Text.RegularExpressions;

namespace OutlookAI.Core.Text
{
    /// <summary>
    /// Minimal HTML-to-plain-text conversion for mail bodies whose plain-text rendering
    /// is unavailable (the MCP <c>read</c> tool always returns text - v3.MD section 8).
    /// Outlook itself maintains <c>MailItem.Body</c> as the plain-text rendering, so this
    /// is the fallback path only. Pure logic, no dependencies; every regex carries a
    /// timeout so hostile input degrades to raw-ish text instead of hanging.
    /// </summary>
    public static class HtmlToText
    {
        private static readonly TimeSpan RegexTimeout = TimeSpan.FromSeconds(2);

        private static readonly Regex RemovableBlocks = new Regex(
            @"<(script|style|head|title)\b[^>]*>.*?</\1\s*>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.Compiled,
            RegexTimeout);

        private static readonly Regex Comments = new Regex(
            @"<!--.*?-->",
            RegexOptions.Singleline | RegexOptions.Compiled,
            RegexTimeout);

        private static readonly Regex LineBreakTags = new Regex(
            @"<\s*(br|/p|/div|/tr|/li|/h[1-6]|/table|/blockquote)\b[^>]*>",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            RegexTimeout);

        private static readonly Regex ListItemTags = new Regex(
            @"<\s*li\b[^>]*>",
            RegexOptions.IgnoreCase | RegexOptions.Compiled,
            RegexTimeout);

        private static readonly Regex AnyTag = new Regex(
            @"<[^>]+>",
            RegexOptions.Compiled,
            RegexTimeout);

        private static readonly Regex SpaceRuns = new Regex(
            @"[ \t ]{2,}",
            RegexOptions.Compiled,
            RegexTimeout);

        private static readonly Regex BlankLineRuns = new Regex(
            @"(\r?\n[ \t]*){3,}",
            RegexOptions.Compiled,
            RegexTimeout);

        /// <summary>Converts an HTML fragment/document to readable plain text.</summary>
        public static string Convert(string? html)
        {
            if (string.IsNullOrEmpty(html))
            {
                return string.Empty;
            }

            string text = html!;
            try
            {
                text = RemovableBlocks.Replace(text, string.Empty);
                text = Comments.Replace(text, string.Empty);
                text = LineBreakTags.Replace(text, "\n");
                text = ListItemTags.Replace(text, "\n- ");
                text = AnyTag.Replace(text, " ");
                text = DecodeEntities(text);
                text = SpaceRuns.Replace(text, " ");
                text = TrimLineEnds(text);
                text = BlankLineRuns.Replace(text, "\n\n");
            }
            catch (RegexMatchTimeoutException)
            {
                // Pathological input: fall through with whatever transformations applied.
            }

            return text.Trim();
        }

        private static string TrimLineEnds(string text)
        {
            string[] lines = text.Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                lines[i] = lines[i].TrimEnd().TrimEnd('\r');
            }

            return string.Join("\n", lines);
        }

        private static string DecodeEntities(string text)
        {
            StringBuilder sb = new StringBuilder(text.Length);
            int i = 0;
            while (i < text.Length)
            {
                char c = text[i];
                if (c != '&')
                {
                    sb.Append(c);
                    i++;
                    continue;
                }

                int semicolon = text.IndexOf(';', i + 1);
                if (semicolon < 0 || semicolon - i > 10)
                {
                    sb.Append(c);
                    i++;
                    continue;
                }

                string entity = text.Substring(i + 1, semicolon - i - 1);
                string? decoded = DecodeEntity(entity);
                if (decoded == null)
                {
                    sb.Append(c);
                    i++;
                    continue;
                }

                sb.Append(decoded);
                i = semicolon + 1;
            }

            return sb.ToString();
        }

        private static string? DecodeEntity(string entity)
        {
            switch (entity.ToLowerInvariant())
            {
                case "nbsp":
                    return " ";
                case "amp":
                    return "&";
                case "lt":
                    return "<";
                case "gt":
                    return ">";
                case "quot":
                    return "\"";
                case "apos":
                    return "'";
                case "ndash":
                    return "–";
                case "mdash":
                    return "—";
                case "hellip":
                    return "…";
                case "copy":
                    return "©";
                case "euro":
                    return "€";
            }

            if (entity.Length > 1 && entity[0] == '#')
            {
                string number = entity.Substring(1);
                bool isHex = number.Length > 1 && (number[0] == 'x' || number[0] == 'X');
                if (isHex)
                {
                    number = number.Substring(1);
                }

                if (int.TryParse(
                        number,
                        isHex ? System.Globalization.NumberStyles.HexNumber : System.Globalization.NumberStyles.Integer,
                        System.Globalization.CultureInfo.InvariantCulture,
                        out int codePoint)
                    && codePoint > 0 && codePoint <= 0x10FFFF
                    && (codePoint < 0xD800 || codePoint > 0xDFFF))
                {
                    return char.ConvertFromUtf32(codePoint);
                }
            }

            return null;
        }
    }
}
