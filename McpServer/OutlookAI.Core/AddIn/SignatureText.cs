using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace OutlookAI.Core.AddIn
{
    /// <summary>A signature file decoded: its text and how it was stored, so a rewrite can store it the same way.</summary>
    public sealed class DecodedSignatureText
    {
        /// <summary>Creates the description.</summary>
        public DecodedSignatureText(string text, string encodingName, int codePage, bool hasBom, long byteLength)
        {
            Text = text;
            EncodingName = encodingName;
            CodePage = codePage;
            HasBom = hasBom;
            ByteLength = byteLength;
        }

        /// <summary>The decoded text.</summary>
        public string Text { get; }

        /// <summary>The encoding's web name: utf-8, utf-16, windows-1252, us-ascii...</summary>
        public string EncodingName { get; }

        /// <summary>The encoding's code page (65001 UTF-8, 1200 UTF-16 LE, 1252 ...).</summary>
        public int CodePage { get; }

        /// <summary>True when the file started with a byte order mark.</summary>
        public bool HasBom { get; }

        /// <summary>The file's size.</summary>
        public long ByteLength { get; }
    }

    /// <summary>
    /// Reads and writes the three signature renditions in the encodings they are stored in (Q159:
    /// "without breaking or converting anything"). The rules, per file:
    /// <list type="bullet">
    /// <item><b>.htm</b> - a byte order mark wins; else the <c>charset</c> the document declares
    /// (Word writes <c>windows-1252</c> for a Western signature); else UTF-8 when the bytes are valid
    /// UTF-8; else the ANSI code page. A rewrite keeps the declared charset, writing any character it
    /// cannot hold as a numeric character reference rather than switching encodings; a document that
    /// declares none gets a <c>utf-8</c> declaration only when it needs one (non-ASCII text).</item>
    /// <item><b>.txt</b> - a byte order mark wins; else UTF-8 when valid; else ANSI. A rewrite keeps
    /// the encoding, and moves to UTF-16 LE with a mark (what Outlook writes) only when the text no
    /// longer fits the old one.</item>
    /// <item><b>.rtf</b> - 7-bit by definition; anything else is written as <c>\uN?</c>.</item>
    /// </list>
    /// </summary>
    public static class SignatureText
    {
        private static readonly Regex MetaCharset = new Regex(
            @"<meta\b[^>]*?charset\s*=\s*[""']?\s*(?<cs>[A-Za-z0-9_\-:.]+)",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));

#if !NETFRAMEWORK
        static SignatureText()
        {
            // The ANSI code pages are not in .NET's box until the provider is registered; on .NET
            // Framework they always are.
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
#endif

        /// <summary>The ANSI code page a file without a mark is read in: the installed UI culture's, 1252 when it has none.</summary>
        public static int AnsiCodePage
        {
            get
            {
                int codePage = CultureInfo.InstalledUICulture.TextInfo.ANSICodePage;
                return codePage > 0 ? codePage : 1252;
            }
        }

        /// <summary>Reads an .htm rendition.</summary>
        public static DecodedSignatureText ReadHtml(string path)
        {
            byte[] bytes = File.ReadAllBytes(path);
            return DecodeHtml(bytes);
        }

        /// <summary>Reads a .txt rendition.</summary>
        public static DecodedSignatureText ReadText(string path)
        {
            byte[] bytes = File.ReadAllBytes(path);
            return DecodePlain(bytes);
        }

        /// <summary>
        /// Reads an .rtf rendition byte for byte: RTF is 7-bit text, and the one exception - the raw
        /// data after <c>\binN</c> - must survive a rewrite unchanged, so every byte maps to the one
        /// character of the same value (ISO-8859-1) and back.
        /// </summary>
        public static DecodedSignatureText ReadRtf(string path)
        {
            byte[] bytes = File.ReadAllBytes(path);
            return new DecodedSignatureText(Latin1.GetString(bytes), Latin1.WebName, Latin1.CodePage, false, bytes.Length);
        }

        /// <summary>ISO-8859-1: the byte-for-byte mapping the RTF rendition is read and written through.</summary>
        public static Encoding Latin1 => Encoding.GetEncoding(28591);

        /// <summary>Decodes .htm bytes by mark, declared charset, UTF-8 validity, then ANSI. Pure.</summary>
        public static DecodedSignatureText DecodeHtml(byte[] bytes)
        {
            if (bytes == null)
            {
                throw new ArgumentNullException(nameof(bytes));
            }

            if (TryDecodeBom(bytes, out DecodedSignatureText? marked))
            {
                return marked!;
            }

            string sniff = Encoding.ASCII.GetString(bytes, 0, Math.Min(bytes.Length, 4096));
            Encoding? declared = DeclaredEncoding(sniff);
            if (declared != null)
            {
                return new DecodedSignatureText(declared.GetString(bytes), declared.WebName, declared.CodePage, false, bytes.Length);
            }

            return DecodePlain(bytes);
        }

        /// <summary>Decodes .txt bytes by mark, UTF-8 validity, then ANSI. Pure.</summary>
        public static DecodedSignatureText DecodePlain(byte[] bytes)
        {
            if (bytes == null)
            {
                throw new ArgumentNullException(nameof(bytes));
            }

            if (TryDecodeBom(bytes, out DecodedSignatureText? marked))
            {
                return marked!;
            }

            if (IsAscii(bytes))
            {
                return new DecodedSignatureText(Encoding.ASCII.GetString(bytes), "us-ascii", 20127, false, bytes.Length);
            }

            try
            {
                string utf8 = new UTF8Encoding(false, throwOnInvalidBytes: true).GetString(bytes);
                return new DecodedSignatureText(utf8, "utf-8", 65001, false, bytes.Length);
            }
            catch (DecoderFallbackException)
            {
                Encoding ansi = Encoding.GetEncoding(AnsiCodePage);
                return new DecodedSignatureText(ansi.GetString(bytes), ansi.WebName, ansi.CodePage, false, bytes.Length);
            }
        }

        /// <summary>
        /// The bytes to store an .htm rendition as. <paramref name="previous"/> is how the file was
        /// stored before (null for a new signature). Pure.
        /// </summary>
        public static byte[] EncodeHtml(string html, DecodedSignatureText? previous)
        {
            if (html == null)
            {
                throw new ArgumentNullException(nameof(html));
            }

            if (previous != null && previous.HasBom)
            {
                Encoding marked = Encoding.GetEncoding(previous.CodePage);
                return WithPreamble(marked, marked.GetBytes(html));
            }

            Encoding? declared = DeclaredEncoding(html);
            if (declared != null)
            {
                if (declared.CodePage == 65001)
                {
                    return new UTF8Encoding(false).GetBytes(html);
                }

                if (declared.IsSingleByte)
                {
                    return EncodeWithReferences(html, declared);
                }

                return declared.GetBytes(html);
            }

            if (IsAscii(html))
            {
                return Encoding.ASCII.GetBytes(html);
            }

            return new UTF8Encoding(false).GetBytes(EnsureUtf8Declaration(html));
        }

        /// <summary>
        /// The bytes to store a .txt rendition as: the old encoding when the text still fits it,
        /// UTF-16 LE with a mark (Outlook's own choice) otherwise and for a new signature. Pure.
        /// </summary>
        public static byte[] EncodeText(string text, DecodedSignatureText? previous)
        {
            if (text == null)
            {
                throw new ArgumentNullException(nameof(text));
            }

            if (previous != null)
            {
                if (previous.HasBom)
                {
                    Encoding marked = Encoding.GetEncoding(previous.CodePage);
                    return WithPreamble(marked, marked.GetBytes(text));
                }

                if (previous.CodePage == 65001)
                {
                    return new UTF8Encoding(false).GetBytes(text);
                }

                Encoding ansi = Encoding.GetEncoding(
                    previous.CodePage == 20127 ? AnsiCodePage : previous.CodePage,
                    EncoderFallback.ExceptionFallback,
                    DecoderFallback.ExceptionFallback);
                try
                {
                    return ansi.GetBytes(text);
                }
                catch (EncoderFallbackException)
                {
                    // The text no longer fits: Outlook's own encoding, marked, which every reader takes.
                }
            }

            return WithPreamble(Encoding.Unicode, Encoding.Unicode.GetBytes(text));
        }

        /// <summary>The bytes of an .rtf rendition: 7-bit, anything else as <c>\uN?</c>. Pure.</summary>
        public static byte[] EncodeRtf(string rtf)
        {
            if (rtf == null)
            {
                throw new ArgumentNullException(nameof(rtf));
            }

            return Latin1.GetBytes(IsAscii(rtf) ? rtf : SignatureRtf.EscapeNonAscii(rtf));
        }

        /// <summary>The charset an HTML document declares, as an encoding .NET has, or null.</summary>
        public static Encoding? DeclaredEncoding(string html)
        {
            Match match = MetaCharset.Match(html ?? string.Empty);
            if (!match.Success)
            {
                return null;
            }

            try
            {
                return Encoding.GetEncoding(match.Groups["cs"].Value.Trim());
            }
            catch (ArgumentException)
            {
                return null;
            }
        }

        /// <summary>
        /// An HTML document that declares UTF-8: one that declares nothing gets the declaration in its
        /// head (or a minimal head when it has none); one that already declares a charset keeps its
        /// own. Pure.
        /// </summary>
        public static string EnsureUtf8Declaration(string html)
        {
            if (DeclaredEncoding(html) != null)
            {
                return html;
            }

            const string meta = "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\">";
            int head = IndexOfTag(html, "head");
            if (head >= 0)
            {
                int close = html.IndexOf('>', head);
                if (close > head)
                {
                    return html.Substring(0, close + 1) + meta + html.Substring(close + 1);
                }
            }

            int root = IndexOfTag(html, "html");
            if (root >= 0)
            {
                int close = html.IndexOf('>', root);
                if (close > root)
                {
                    return html.Substring(0, close + 1) + "<head>" + meta + "</head>" + html.Substring(close + 1);
                }
            }

            return "<html><head>" + meta + "</head><body>\r\n" + html + "\r\n</body></html>";
        }

        /// <summary>True when every character is 7-bit.</summary>
        public static bool IsAscii(string text)
        {
            foreach (char c in text)
            {
                if (c > 0x7F)
                {
                    return false;
                }
            }

            return true;
        }

        private static bool IsAscii(byte[] bytes)
        {
            foreach (byte b in bytes)
            {
                if (b > 0x7F)
                {
                    return false;
                }
            }

            return true;
        }

        private static int IndexOfTag(string html, string tag)
        {
            Match match = Regex.Match(html, "<" + tag + @"(\s|>)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
            return match.Success ? match.Index : -1;
        }

        private static bool TryDecodeBom(byte[] bytes, out DecodedSignatureText? decoded)
        {
            decoded = null;
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF)
            {
                decoded = new DecodedSignatureText(Encoding.UTF8.GetString(bytes, 3, bytes.Length - 3), "utf-8", 65001, true, bytes.Length);
            }
            else if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE)
            {
                decoded = new DecodedSignatureText(Encoding.Unicode.GetString(bytes, 2, bytes.Length - 2), "utf-16", 1200, true, bytes.Length);
            }
            else if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF)
            {
                decoded = new DecodedSignatureText(
                    Encoding.BigEndianUnicode.GetString(bytes, 2, bytes.Length - 2), "unicodeFFFE", 1201, true, bytes.Length);
            }

            return decoded != null;
        }

        private static byte[] WithPreamble(Encoding encoding, byte[] body)
        {
            byte[] preamble = encoding.GetPreamble();
            if (preamble.Length == 0)
            {
                preamble = encoding.CodePage == 65001 ? new byte[] { 0xEF, 0xBB, 0xBF }
                    : encoding.CodePage == 1200 ? new byte[] { 0xFF, 0xFE }
                    : encoding.CodePage == 1201 ? new byte[] { 0xFE, 0xFF }
                    : preamble;
            }

            byte[] result = new byte[preamble.Length + body.Length];
            Buffer.BlockCopy(preamble, 0, result, 0, preamble.Length);
            Buffer.BlockCopy(body, 0, result, preamble.Length, body.Length);
            return result;
        }

        /// <summary>
        /// Encodes HTML in a single-byte charset it declares, writing any character the charset cannot
        /// hold as a numeric character reference - so the declaration stays true and nothing renders
        /// as a question mark.
        /// </summary>
        private static byte[] EncodeWithReferences(string html, Encoding declared)
        {
            Encoding strict = Encoding.GetEncoding(declared.CodePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
            try
            {
                return strict.GetBytes(html);
            }
            catch (EncoderFallbackException)
            {
                StringBuilder safe = new StringBuilder(html.Length + 64);
                for (int i = 0; i < html.Length; i++)
                {
                    char c = html[i];
                    string unit = char.IsHighSurrogate(c) && i + 1 < html.Length && char.IsLowSurrogate(html[i + 1])
                        ? html.Substring(i, 2)
                        : c.ToString();
                    if (CanEncode(strict, unit))
                    {
                        safe.Append(unit);
                    }
                    else
                    {
                        int code = unit.Length == 2 ? char.ConvertToUtf32(unit[0], unit[1]) : unit[0];
                        safe.Append("&#").Append(code.ToString(CultureInfo.InvariantCulture)).Append(';');
                    }

                    i += unit.Length - 1;
                }

                return strict.GetBytes(safe.ToString());
            }
        }

        private static bool CanEncode(Encoding strict, string unit)
        {
            try
            {
                strict.GetBytes(unit);
                return true;
            }
            catch (EncoderFallbackException)
            {
                return false;
            }
        }
    }
}
