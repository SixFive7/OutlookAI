using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using OutlookAI.Core.AddIn;

namespace OutlookAI.Core.Services
{
    /// <summary>One reference an HTML signature makes to a file next to it.</summary>
    public sealed class SignatureHtmlReference
    {
        /// <summary>Creates the description.</summary>
        public SignatureHtmlReference(string attribute, string value, string? fileName)
        {
            Attribute = attribute;
            Value = value;
            FileName = fileName;
        }

        /// <summary>src, href or background.</summary>
        public string Attribute { get; }

        /// <summary>The value as written (percent-encoded as Word writes it).</summary>
        public string Value { get; }

        /// <summary>The file in the signature's <c>_files</c> folder it names, or null when it names something else.</summary>
        public string? FileName { get; }
    }

    /// <summary>
    /// The HTML rendition, edited in place (Q159): its references to the pictures in
    /// <c>&lt;name&gt;_files</c>, and text replacement that understands how HTML spells text
    /// (entities, collapsed white space) and never reaches into markup, comments, styles or scripts.
    /// </summary>
    public static class SignatureHtml
    {
        private static readonly Regex Reference = new Regex(
            @"(?<attr>\b(?:src|href|background))\s*=\s*(?:""(?<v>[^""]*)""|'(?<v>[^']*)'|(?<v>[^\s>]+))",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(5));

        private static readonly Regex Protected = new Regex(
            @"<!--.*?-->|<(?<block>style|script|head|title|xml)\b[^>]*>.*?</\k<block>\s*>|<[^>]*>",
            RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(5));

        private static readonly string[] Latin1Entities =
        {
            "nbsp", "iexcl", "cent", "pound", "curren", "yen", "brvbar", "sect", "uml", "copy", "ordf", "laquo", "not", "shy", "reg",
            "macr", "deg", "plusmn", "sup2", "sup3", "acute", "micro", "para", "middot", "cedil", "sup1", "ordm", "raquo", "frac14",
            "frac12", "frac34", "iquest", "Agrave", "Aacute", "Acirc", "Atilde", "Auml", "Aring", "AElig", "Ccedil", "Egrave",
            "Eacute", "Ecirc", "Euml", "Igrave", "Iacute", "Icirc", "Iuml", "ETH", "Ntilde", "Ograve", "Oacute", "Ocirc", "Otilde",
            "Ouml", "times", "Oslash", "Ugrave", "Uacute", "Ucirc", "Uuml", "Yacute", "THORN", "szlig", "agrave", "aacute", "acirc",
            "atilde", "auml", "aring", "aelig", "ccedil", "egrave", "eacute", "ecirc", "euml", "igrave", "iacute", "icirc", "iuml",
            "eth", "ntilde", "ograve", "oacute", "ocirc", "otilde", "ouml", "divide", "oslash", "ugrave", "uacute", "ucirc", "uuml",
            "yacute", "thorn", "yuml",
        };

        private static readonly Dictionary<char, string> OtherEntities = new Dictionary<char, string>
        {
            ['\u20AC'] = "euro",
            ['\u2013'] = "ndash",
            ['\u2014'] = "mdash",
            ['\u2018'] = "lsquo",
            ['\u2019'] = "rsquo",
            ['\u201A'] = "sbquo",
            ['\u201C'] = "ldquo",
            ['\u201D'] = "rdquo",
            ['\u201E'] = "bdquo",
            ['\u2022'] = "bull",
            ['\u2026'] = "hellip",
            ['\u2122'] = "trade",
        };

        /// <summary>
        /// The references a signature's HTML makes - <c>src</c>, <c>href</c> and <c>background</c>
        /// values - with the file each names in <paramref name="signatureName"/>'s <c>_files</c>
        /// folder, decoded the way Word encodes them (<c>My%20Sig_files/image001.png</c>). Pure.
        /// </summary>
        public static IReadOnlyList<SignatureHtmlReference> References(string html, string signatureName)
        {
            List<SignatureHtmlReference> references = new List<SignatureHtmlReference>();
            foreach (Match match in Reference.Matches(html ?? string.Empty))
            {
                string value = WebUtility.HtmlDecode(match.Groups["v"].Value.Trim());
                references.Add(new SignatureHtmlReference(match.Groups["attr"].Value.ToLowerInvariant(), value, ResourceFileOf(value, signatureName)));
            }

            return references;
        }

        /// <summary>
        /// The file a reference names in the signature's <c>_files</c> folder, or null. Accepts
        /// <c>Name_files/x</c> and <c>Name_files\x</c>, percent-encoded or not; never a path that leaves
        /// the folder.
        /// </summary>
        public static string? ResourceFileOf(string reference, string signatureName)
        {
            if (string.IsNullOrWhiteSpace(reference) || reference.IndexOf(':') >= 0)
            {
                return null;
            }

            string decoded;
            try
            {
                decoded = Uri.UnescapeDataString(reference.Trim());
            }
            catch (UriFormatException)
            {
                decoded = reference.Trim();
            }

            decoded = decoded.Replace('\\', '/');
            string prefix = signatureName + SignatureCatalog.ResourceDirectorySuffix + "/";
            if (!decoded.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                return null;
            }

            string file = decoded.Substring(prefix.Length);
            int query = file.IndexOfAny(new[] { '?', '#' });
            if (query >= 0)
            {
                file = file.Substring(0, query);
            }

            return file.Length == 0 || file.IndexOf('/') >= 0 || file == "." || file == ".." ? null : file;
        }

        /// <summary>True when a reference points at a local file outside the signature's own folder (a bare <c>logo.png</c>).</summary>
        public static bool IsStrayLocalReference(SignatureHtmlReference reference)
        {
            if (reference.FileName != null || reference.Attribute == "href")
            {
                return false;
            }

            string value = reference.Value;
            return value.Length > 0 && value.IndexOf(':') < 0 && !value.StartsWith("#", StringComparison.Ordinal)
                && !value.StartsWith("//", StringComparison.Ordinal);
        }

        /// <summary>
        /// The relative reference Word writes for a file of the signature's folder:
        /// <c>Name_files/file</c>, spaces and other unsafe characters percent-encoded.
        /// </summary>
        public static string ReferenceFor(string signatureName, string fileName)
        {
            return Uri.EscapeDataString(signatureName + SignatureCatalog.ResourceDirectorySuffix) + "/" + Uri.EscapeDataString(fileName);
        }

        /// <summary>
        /// Replaces visible text in place: every occurrence of <paramref name="find"/> in the
        /// document's text - however HTML spells it (entities, white space wrapped across lines) -
        /// that is not split by markup, outside tags, comments, styles and scripts. An occurrence split
        /// by markup is not touched and is reported. The replacement is HTML-encoded, line breaks as
        /// <c>&lt;br&gt;</c>. Pure.
        /// </summary>
        public static SignatureTextReplacement ReplaceText(string html, string find, string replace)
        {
            if (html == null)
            {
                throw new ArgumentNullException(nameof(html));
            }

            if (string.IsNullOrEmpty(find))
            {
                throw new ArgumentException("find must not be empty.", nameof(find));
            }

            List<KeyValuePair<int, int>> blocked = Protected.Matches(html).Cast<Match>()
                .Select(m => new KeyValuePair<int, int>(m.Index, m.Index + m.Length))
                .ToList();
            Regex pattern = new Regex(FindPattern(find), RegexOptions.CultureInvariant, TimeSpan.FromSeconds(5));
            List<Match> accepted = new List<Match>();
            foreach (Match match in pattern.Matches(html))
            {
                int start = match.Index;
                int end = match.Index + match.Length;
                if (match.Length > 0 && !blocked.Any(b => start < b.Value && end > b.Key))
                {
                    accepted.Add(match);
                }
            }

            if (accepted.Count == 0)
            {
                string visible = AddIn.HtmlToText.Convert(html);
                bool present = Collapse(visible).IndexOf(Collapse(find), StringComparison.Ordinal) >= 0;
                return new SignatureTextReplacement(html, 0, present);
            }

            // Only what HTML itself reserves is escaped: other characters go in as they are, and the
            // file's own charset decides how they are stored (SignatureText.EncodeHtml), as Word does it.
            string encoded = (replace ?? string.Empty)
                .Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;")
                .Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", "<br>\r\n");
            StringBuilder result = new StringBuilder(html.Length + accepted.Count * encoded.Length);
            int at = 0;
            foreach (Match match in accepted)
            {
                result.Append(html, at, match.Index - at).Append(encoded);
                at = match.Index + match.Length;
            }

            result.Append(html, at, html.Length - at);
            return new SignatureTextReplacement(result.ToString(), accepted.Count, false);
        }

        /// <summary>A regex matching <paramref name="find"/> however HTML would spell it.</summary>
        private static string FindPattern(string find)
        {
            StringBuilder pattern = new StringBuilder(find.Length * 12);
            bool inSpace = false;
            foreach (char ch in find)
            {
                if (ch == ' ' || ch == '\u00A0' || ch == '\t')
                {
                    if (!inSpace)
                    {
                        pattern.Append(@"(?:[ \t\r\n\u00A0]|&nbsp;|&#160;|&#[xX]0*[aA]0;)+");
                        inSpace = true;
                    }

                    continue;
                }

                inSpace = false;
                List<string> forms = new List<string>();
                switch (ch)
                {
                    case '&':
                        forms.Add("&amp;");
                        break;
                    case '<':
                        forms.Add("&lt;");
                        break;
                    case '>':
                        forms.Add("&gt;");
                        forms.Add(">");
                        break;
                    case '"':
                        forms.Add("\"");
                        forms.Add("&quot;");
                        break;
                    case '\'':
                        forms.Add("'");
                        forms.Add("&apos;");
                        break;
                    default:
                        forms.Add(Regex.Escape(ch.ToString()));
                        break;
                }

                int code = ch;
                forms.Add("&#0*" + code.ToString(CultureInfo.InvariantCulture) + ";");
                forms.Add("&#[xX]0*(?i:" + code.ToString("x", CultureInfo.InvariantCulture) + ");");
                if (code >= 160 && code <= 255)
                {
                    forms.Add("&" + Latin1Entities[code - 160] + ";");
                }
                else if (OtherEntities.TryGetValue(ch, out string? name))
                {
                    forms.Add("&" + name + ";");
                }

                pattern.Append("(?:").Append(string.Join("|", forms)).Append(')');
            }

            return pattern.ToString();
        }

        private static string Collapse(string text)
        {
            return Regex.Replace(text.Replace('\u00A0', ' '), @"\s+", " ", RegexOptions.None, TimeSpan.FromSeconds(2));
        }
    }
}
