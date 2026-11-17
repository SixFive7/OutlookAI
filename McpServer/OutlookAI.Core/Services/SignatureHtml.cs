using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace OutlookAI.Core.Services
{
    /// <summary>Pictures a signature can carry: sniffing, sizes and the limits a new one is held to.</summary>
    public static class SignaturePictures
    {
        /// <summary>The largest picture accepted into a signature: 2 MB. A signature rides on every message.</summary>
        public const int MaxPictureBytes = 2 * 1024 * 1024;

        /// <summary>The picture formats a new picture may be: PNG, JPEG, GIF, BMP.</summary>
        public static readonly IReadOnlyList<string> AcceptedExtensions = Array.AsReadOnly(new[] { ".png", ".jpg", ".jpeg", ".gif", ".bmp" });

        /// <summary>PNG signature bytes.</summary>
        public static bool IsPng(byte[] data)
        {
            return data != null && data.Length >= 8 && data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47
                && data[4] == 0x0D && data[5] == 0x0A && data[6] == 0x1A && data[7] == 0x0A;
        }

        /// <summary>JPEG start-of-image.</summary>
        public static bool IsJpeg(byte[] data)
        {
            return data != null && data.Length >= 3 && data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF;
        }

        /// <summary>GIF87a / GIF89a.</summary>
        public static bool IsGif(byte[] data)
        {
            return data != null && data.Length >= 6 && data[0] == (byte)'G' && data[1] == (byte)'I' && data[2] == (byte)'F'
                && data[3] == (byte)'8' && (data[4] == (byte)'7' || data[4] == (byte)'9') && data[5] == (byte)'a';
        }

        /// <summary>A Windows bitmap.</summary>
        public static bool IsBmp(byte[] data)
        {
            return data != null && data.Length >= 26 && data[0] == (byte)'B' && data[1] == (byte)'M';
        }

        /// <summary>True when the bytes are the kind of picture the extension says.</summary>
        public static bool MatchesExtension(byte[] data, string extension)
        {
            switch ((extension ?? string.Empty).ToLowerInvariant())
            {
                case ".png":
                    return IsPng(data);
                case ".jpg":
                case ".jpeg":
                    return IsJpeg(data);
                case ".gif":
                    return IsGif(data);
                case ".bmp":
                    return IsBmp(data);
                default:
                    return false;
            }
        }

        /// <summary>The MIME type for a picture file name.</summary>
        public static string ContentType(string fileName)
        {
            switch (Path.GetExtension(fileName ?? string.Empty).ToLowerInvariant())
            {
                case ".png":
                    return "image/png";
                case ".jpg":
                case ".jpeg":
                    return "image/jpeg";
                case ".gif":
                    return "image/gif";
                case ".bmp":
                    return "image/bmp";
                case ".svg":
                    return "image/svg+xml";
                case ".webp":
                    return "image/webp";
                case ".tif":
                case ".tiff":
                    return "image/tiff";
                case ".emf":
                case ".emz":
                    return "image/x-emf";
                case ".wmf":
                case ".wmz":
                    return "image/x-wmf";
                case ".xml":
                    return "text/xml";
                case ".thmx":
                    return "application/vnd.ms-officetheme";
                default:
                    return "application/octet-stream";
            }
        }

        /// <summary>The picture's size in pixels from its header (PNG, JPEG, GIF, BMP); false when unreadable.</summary>
        public static bool TryReadSize(byte[] data, out int width, out int height)
        {
            width = 0;
            height = 0;
            if (data == null)
            {
                return false;
            }

            if (IsPng(data) && data.Length >= 24)
            {
                width = (data[16] << 24) | (data[17] << 16) | (data[18] << 8) | data[19];
                height = (data[20] << 24) | (data[21] << 16) | (data[22] << 8) | data[23];
                return width > 0 && height > 0;
            }

            if (IsGif(data) && data.Length >= 10)
            {
                width = data[6] | (data[7] << 8);
                height = data[8] | (data[9] << 8);
                return width > 0 && height > 0;
            }

            if (IsBmp(data))
            {
                width = BitConverter.ToInt32(data, 18);
                height = Math.Abs(BitConverter.ToInt32(data, 22));
                return width > 0 && height > 0;
            }

            if (IsJpeg(data))
            {
                int i = 2;
                while (i + 9 < data.Length)
                {
                    if (data[i] != 0xFF)
                    {
                        i++;
                        continue;
                    }

                    byte marker = data[i + 1];
                    int length = (data[i + 2] << 8) | data[i + 3];
                    bool startOfFrame = marker >= 0xC0 && marker <= 0xCF && marker != 0xC4 && marker != 0xC8 && marker != 0xCC;
                    if (startOfFrame)
                    {
                        height = (data[i + 5] << 8) | data[i + 6];
                        width = (data[i + 7] << 8) | data[i + 8];
                        return width > 0 && height > 0;
                    }

                    if (length < 2)
                    {
                        return false;
                    }

                    i += 2 + length;
                }
            }

            return false;
        }

        /// <summary>
        /// Reads a picture an agent named by absolute path, held to the rules every new picture is:
        /// an existing file, not empty, at most <see cref="MaxPictureBytes"/>, a PNG, JPEG, GIF or BMP
        /// whose bytes are what its extension says. Throws ArgumentException naming the path and the
        /// rule broken; nothing has been written when it does.
        /// </summary>
        public static byte[] ReadNewPicture(string path, string targetFileName)
        {
            if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path))
            {
                throw new ArgumentException("picture path '" + path + "' must be an absolute path to a file on this machine. Nothing was written.");
            }

            if (Directory.Exists(path))
            {
                throw new ArgumentException("picture path '" + path + "' is a folder, not a file. Nothing was written.");
            }

            if (!File.Exists(path))
            {
                throw new ArgumentException("picture file '" + path + "' does not exist. Nothing was written.");
            }

            long length = new FileInfo(path).Length;
            if (length == 0)
            {
                throw new ArgumentException("picture file '" + path + "' is empty. Nothing was written.");
            }

            if (length > MaxPictureBytes)
            {
                throw new ArgumentException("picture file '" + path + "' is " + length.ToString(CultureInfo.InvariantCulture)
                    + " bytes; a signature picture may be at most " + MaxPictureBytes.ToString(CultureInfo.InvariantCulture)
                    + " bytes, because it travels with every message. Nothing was written.");
            }

            string extension = Path.GetExtension(targetFileName);
            if (!AcceptedExtensions.Contains(extension, StringComparer.OrdinalIgnoreCase))
            {
                throw new ArgumentException("picture '" + targetFileName + "' must be a .png, .jpg, .jpeg, .gif or .bmp file. Nothing was written.");
            }

            byte[] data = File.ReadAllBytes(path);
            if (!MatchesExtension(data, extension))
            {
                throw new ArgumentException("picture file '" + path + "' is not a " + extension.TrimStart('.').ToUpperInvariant()
                    + " picture, which its target name '" + targetFileName + "' says it is. Name it after its real format. Nothing was written.");
            }

            return data;
        }
    }

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
            ['€'] = "euro",
            ['–'] = "ndash",
            ['—'] = "mdash",
            ['‘'] = "lsquo",
            ['’'] = "rsquo",
            ['‚'] = "sbquo",
            ['“'] = "ldquo",
            ['”'] = "rdquo",
            ['„'] = "bdquo",
            ['•'] = "bull",
            ['…'] = "hellip",
            ['™'] = "trade",
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
                string visible = Text.HtmlToText.Convert(html);
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
                if (ch == ' ' || ch == ' ' || ch == '\t')
                {
                    if (!inSpace)
                    {
                        pattern.Append(@"(?:[ \t\r\n ]|&nbsp;|&#160;|&#[xX]0*[aA]0;)+");
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
            return Regex.Replace(text.Replace(' ', ' '), @"\s+", " ", RegexOptions.None, TimeSpan.FromSeconds(2));
        }
    }
}
