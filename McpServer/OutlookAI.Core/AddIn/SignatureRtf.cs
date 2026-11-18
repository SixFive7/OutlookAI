using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

namespace OutlookAI.Core.AddIn
{
    /// <summary>One picture inside an RTF signature (data only).</summary>
    public sealed class RtfPicture
    {
        /// <summary>Creates the description.</summary>
        public RtfPicture(int number, int start, int length, string kind, int byteCount, int? widthTwips, int? heightTwips, string sha256, bool hasFallback)
        {
            Number = number;
            Start = start;
            Length = length;
            Kind = kind;
            ByteCount = byteCount;
            WidthTwips = widthTwips;
            HeightTwips = heightTwips;
            Sha256 = sha256;
            HasFallback = hasFallback;
        }

        /// <summary>1-based position among the document's pictures - the N of <c>[[picture:N]]</c>.</summary>
        public int Number { get; }

        /// <summary>Where the picture's whole construct starts in the RTF.</summary>
        public int Start { get; }

        /// <summary>The construct's length - a <c>{\pict}</c> group, or a <c>{\*\shppict}</c> group with its <c>{\nonshppict}</c> fallback.</summary>
        public int Length { get; }

        /// <summary>png, jpeg, emf, wmf, bitmap or unknown.</summary>
        public string Kind { get; }

        /// <summary>The picture's own bytes (decoded from hex).</summary>
        public int ByteCount { get; }

        /// <summary>Displayed width (<c>\picwgoal</c>, twips), when recorded.</summary>
        public int? WidthTwips { get; }

        /// <summary>Displayed height (<c>\pichgoal</c>, twips), when recorded.</summary>
        public int? HeightTwips { get; }

        /// <summary>SHA-256 of the picture's bytes, hex.</summary>
        public string Sha256 { get; }

        /// <summary>True when the construct carries a <c>\nonshppict</c> fallback copy for old readers.</summary>
        public bool HasFallback { get; }
    }

    /// <summary>The result of one find/replace over one rendition.</summary>
    public sealed class SignatureTextReplacement
    {
        /// <summary>Creates the result.</summary>
        public SignatureTextReplacement(string content, int replaced, bool foundButSplit)
        {
            Content = content;
            Replaced = replaced;
            FoundButSplit = foundButSplit;
        }

        /// <summary>The rendition after the replacement (unchanged when nothing was replaced).</summary>
        public string Content { get; }

        /// <summary>How many occurrences were replaced.</summary>
        public int Replaced { get; }

        /// <summary>
        /// True when the visible text DOES contain the phrase but formatting splits it in the source
        /// (a bold word in the middle, a spelling mark): nothing is replaced, because replacing across
        /// the formatting would move or drop it.
        /// </summary>
        public bool FoundButSplit { get; }
    }

    /// <summary>
    /// The RTF rendition, edited in place (Q159). Outlook's .rtf is what Rich Text mail inserts, with
    /// its pictures EMBEDDED as hex - typically a <c>{\*\shppict{\pict ... \pngblip ...}}</c> followed by
    /// a <c>{\nonshppict{\pict ... \wmetafile8 ...}}</c> fallback for old readers. Nothing here rebuilds
    /// the RTF from another format: reading elides each picture to a <c>[[picture:N]]</c> token so the
    /// text can be read and edited without shipping kilobytes of hex, and writing puts the original
    /// construct back byte for byte wherever its token still stands.
    /// <para>
    /// Tokens: <c>[[picture:N]]</c> is the document's N-th picture as it was; <c>[[picture-file:NAME]]</c>
    /// or <c>[[picture-file:NAME|WIDTHxHEIGHT]]</c> is a NEW picture from a PNG or JPEG file (in the
    /// signature's <c>_files</c> folder, or an absolute path), at that size in pixels or its own.
    /// Leaving a token out removes the picture; moving it moves it.
    /// </para>
    /// </summary>
    public static class SignatureRtf
    {
        private static readonly Regex PictureToken = new Regex(
            @"\[\[picture:(?<n>\d+)(?:[^\]]*)\]\]",
            RegexOptions.CultureInvariant, TimeSpan.FromSeconds(5));

        private static readonly Regex PictureFileToken = new Regex(
            @"\[\[picture-file:(?<file>[^\]\|]+?)(?:\|(?<w>\d+)x(?<h>\d+))?\]\]",
            RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, TimeSpan.FromSeconds(5));

        /// <summary>
        /// Destinations whose text is not the document's text: tables, document info, pictures,
        /// objects, field instructions, shapes. A <c>{\*\...}</c> group is skipped whatever it is.
        /// </summary>
        private static readonly HashSet<string> SkippedDestinations = new HashSet<string>(StringComparer.Ordinal)
        {
            "fonttbl", "colortbl", "stylesheet", "info", "pict", "object", "fldinst", "header", "headerl", "headerr",
            "headerf", "footer", "footerl", "footerr", "footerf", "listtable", "listoverridetable", "rsidtbl", "generator",
            "themedata", "colorschememapping", "latentstyles", "datastore", "xmlnstbl", "mmathPr", "nonshppict", "shpinst",
            "shprslt", "filetbl", "revtbl", "pgdsctbl", "private", "userprops", "docvar", "bkmkstart", "bkmkend", "xe", "tc",
        };

        // ------------------------------------------------------------------ pictures

        /// <summary>The pictures of an RTF document, in order. Pure.</summary>
        public static IReadOnlyList<RtfPicture> FindPictures(string rtf)
        {
            if (rtf == null)
            {
                throw new ArgumentNullException(nameof(rtf));
            }

            List<RtfPicture> pictures = new List<RtfPicture>();
            int index = 0;
            while (true)
            {
                int pict = FindGroupStartingWith(rtf, index, "pict");
                if (pict < 0)
                {
                    break;
                }

                int pictEnd = FindGroupEnd(rtf, pict);
                if (pictEnd < 0)
                {
                    break;
                }

                // The construct: a {\*\shppict ...} parent holding only this picture, plus a
                // {\nonshppict ...} fallback right after it; or the bare {\pict ...} group.
                int start = pict;
                int end = pictEnd;
                bool hasFallback = false;
                int parent = EnclosingShppict(rtf, pict, pictEnd);
                if (parent >= 0)
                {
                    start = parent;
                    end = FindGroupEnd(rtf, parent);
                    int next = SkipLineBreaks(rtf, end + 1);
                    if (StartsGroupWith(rtf, next, "nonshppict"))
                    {
                        int fallbackEnd = FindGroupEnd(rtf, next);
                        if (fallbackEnd > next)
                        {
                            end = fallbackEnd;
                            hasFallback = true;
                        }
                    }
                }

                byte[] data = PictureBytes(rtf, pict, pictEnd);
                pictures.Add(new RtfPicture(
                    pictures.Count + 1,
                    start,
                    end - start + 1,
                    PictureKind(rtf, pict, pictEnd),
                    data.Length,
                    ControlParameter(rtf, pict, pictEnd, "picwgoal"),
                    ControlParameter(rtf, pict, pictEnd, "pichgoal"),
                    Sha256Hex(data),
                    hasFallback));
                index = end + 1;
            }

            return pictures;
        }

        /// <summary>The RTF with every picture construct replaced by its <c>[[picture:N]]</c> token. Pure.</summary>
        public static string Elide(string rtf)
        {
            IReadOnlyList<RtfPicture> pictures = FindPictures(rtf);
            if (pictures.Count == 0)
            {
                return rtf;
            }

            StringBuilder result = new StringBuilder(rtf.Length);
            int at = 0;
            foreach (RtfPicture picture in pictures)
            {
                result.Append(rtf, at, picture.Start - at);
                result.Append("[[picture:").Append(picture.Number.ToString(CultureInfo.InvariantCulture)).Append("]]");
                at = picture.Start + picture.Length;
            }

            result.Append(rtf, at, rtf.Length - at);
            return result.ToString();
        }

        /// <summary>
        /// Puts pictures back into an edited RTF: <c>[[picture:N]]</c> becomes the original's N-th
        /// construct byte for byte, <c>[[picture-file:...]]</c> a new picture group from
        /// <paramref name="pictureFile"/> (the file's bytes, or null when there is no such file).
        /// Refuses (ArgumentException) a token naming a picture or file that does not exist, or a new
        /// picture that is not PNG or JPEG. Pure over the resolver.
        /// </summary>
        public static string Expand(string editedRtf, string originalRtf, Func<string, byte[]?> pictureFile)
        {
            if (editedRtf == null)
            {
                throw new ArgumentNullException(nameof(editedRtf));
            }

            IReadOnlyList<RtfPicture> pictures = originalRtf == null ? Array.Empty<RtfPicture>() : FindPictures(originalRtf);
            string withOriginals = PictureToken.Replace(editedRtf, match =>
            {
                int number = int.Parse(match.Groups["n"].Value, CultureInfo.InvariantCulture);
                RtfPicture? picture = pictures.FirstOrDefault(p => p.Number == number);
                if (picture == null)
                {
                    throw new ArgumentException(
                        "body_rtf names [[picture:" + number.ToString(CultureInfo.InvariantCulture) + "]], but the signature's RTF has "
                        + pictures.Count.ToString(CultureInfo.InvariantCulture) + " picture(s). Nothing was written.");
                }

                return originalRtf!.Substring(picture.Start, picture.Length);
            });

            return PictureFileToken.Replace(withOriginals, match =>
            {
                string file = match.Groups["file"].Value.Trim();
                byte[]? bytes = pictureFile(file);
                if (bytes == null)
                {
                    throw new ArgumentException(
                        "body_rtf names [[picture-file:" + file + "]], but no such picture exists - give a file of the signature's "
                        + "pictures (add it with 'pictures' in the same call) or an absolute path. Nothing was written.");
                }

                int? width = match.Groups["w"].Success ? int.Parse(match.Groups["w"].Value, CultureInfo.InvariantCulture) : (int?)null;
                int? height = match.Groups["h"].Success ? int.Parse(match.Groups["h"].Value, CultureInfo.InvariantCulture) : (int?)null;
                return PictureGroup(bytes, width, height);
            });
        }

        /// <summary>
        /// An RTF picture group for a PNG or JPEG: <c>\picw</c>/<c>\pich</c> in HIMETRIC and
        /// <c>\picwgoal</c>/<c>\pichgoal</c> in twips (at 96 dpi), as Word writes them. The size is the
        /// one asked for in pixels, else the picture's own. Refuses any other format.
        /// </summary>
        public static string PictureGroup(byte[] image, int? widthPixels, int? heightPixels)
        {
            if (image == null)
            {
                throw new ArgumentNullException(nameof(image));
            }

            string? blip = SignaturePictures.IsPng(image) ? "pngblip" : SignaturePictures.IsJpeg(image) ? "jpegblip" : null;
            if (blip == null)
            {
                throw new ArgumentException("A picture placed in the RTF rendition must be a PNG or a JPEG. Nothing was written.");
            }

            SignaturePictures.TryReadSize(image, out int ownWidth, out int ownHeight);
            int width = widthPixels ?? (ownWidth > 0 ? ownWidth : 100);
            int height = heightPixels ?? (ownHeight > 0 ? ownHeight : 100);
            StringBuilder group = new StringBuilder(image.Length * 2 + 200);
            group.Append(@"{\pict\picscalex100\picscaley100")
                .Append(@"\picw").Append(ToHimetric(ownWidth > 0 ? ownWidth : width))
                .Append(@"\pich").Append(ToHimetric(ownHeight > 0 ? ownHeight : height))
                .Append(@"\picwgoal").Append((width * 15).ToString(CultureInfo.InvariantCulture))
                .Append(@"\pichgoal").Append((height * 15).ToString(CultureInfo.InvariantCulture))
                .Append('\\').Append(blip).Append("\r\n");
            for (int i = 0; i < image.Length; i++)
            {
                group.Append(image[i].ToString("x2", CultureInfo.InvariantCulture));
                if ((i + 1) % 64 == 0)
                {
                    group.Append("\r\n");
                }
            }

            group.Append('}');
            return group.ToString();
        }

        // ------------------------------------------------------------------ text

        /// <summary>The document's visible text: paragraphs as line breaks, tables and pictures skipped. Pure.</summary>
        public static string ToPlainText(string rtf)
        {
            if (rtf == null)
            {
                throw new ArgumentNullException(nameof(rtf));
            }

            StringBuilder text = new StringBuilder(rtf.Length / 4);
            Encoding ansi = DocumentEncoding(rtf);
            Stack<GroupState> groups = new Stack<GroupState>();
            GroupState state = new GroupState(false, 1);
            int skipFallback = 0;
            int i = 0;
            while (i < rtf.Length)
            {
                char c = rtf[i];
                if (c == '{')
                {
                    groups.Push(state);
                    state = new GroupState(state.Skipping, state.UnicodeSkip);
                    i++;
                    if (i + 1 < rtf.Length && rtf[i] == '\\' && rtf[i + 1] == '*')
                    {
                        state = new GroupState(true, state.UnicodeSkip);
                    }

                    continue;
                }

                if (c == '}')
                {
                    state = groups.Count > 0 ? groups.Pop() : new GroupState(false, 1);
                    i++;
                    continue;
                }

                if (c == '\r' || c == '\n')
                {
                    i++;
                    continue;
                }

                if (c == '\\')
                {
                    ControlToken token = ReadControl(rtf, i);
                    if (token.Word == "bin" && token.Parameter is int binary && binary > 0)
                    {
                        i = Math.Min(rtf.Length, token.End + binary);
                        continue;
                    }

                    if (token.Word != null && SkippedDestinations.Contains(token.Word))
                    {
                        state = new GroupState(true, state.UnicodeSkip);
                    }
                    else if (token.Word == "uc" && token.Parameter is int uc)
                    {
                        state = new GroupState(state.Skipping, uc);
                    }

                    if (!state.Skipping)
                    {
                        if (skipFallback > 0 && token.Symbol == '\'')
                        {
                            skipFallback--;
                        }
                        else
                        {
                            string? produced = Produced(token, ansi);
                            if (produced != null)
                            {
                                text.Append(produced);
                            }

                            if (token.Word == "u")
                            {
                                skipFallback = state.UnicodeSkip;
                            }
                        }
                    }

                    i = token.End;
                    continue;
                }

                if (!state.Skipping)
                {
                    if (skipFallback > 0)
                    {
                        skipFallback--;
                    }
                    else
                    {
                        text.Append(c < 0x80 ? c : DecodeByte(ansi, (byte)c));
                    }
                }

                i++;
            }

            return text.ToString();
        }

        /// <summary>
        /// Replaces visible text in place: every occurrence of <paramref name="find"/> that sits whole
        /// inside one run of text (line breaks in the source allowed) becomes <paramref name="replace"/>,
        /// escaped for RTF. An occurrence split by formatting is not touched and is reported. Pure.
        /// </summary>
        public static SignatureTextReplacement ReplaceText(string rtf, string find, string replace)
        {
            if (rtf == null)
            {
                throw new ArgumentNullException(nameof(rtf));
            }

            if (string.IsNullOrEmpty(find))
            {
                throw new ArgumentException("find must not be empty.", nameof(find));
            }

            Encoding ansi = DocumentEncoding(rtf);
            bool[] visible = VisibleTextMap(rtf);
            Regex pattern = new Regex(FindPattern(find, ansi), RegexOptions.CultureInvariant, TimeSpan.FromSeconds(5));
            List<Match> accepted = new List<Match>();
            foreach (Match match in pattern.Matches(rtf))
            {
                if (match.Length > 0 && Covered(visible, match.Index, match.Length))
                {
                    accepted.Add(match);
                }
            }

            if (accepted.Count == 0)
            {
                bool present = ToPlainText(rtf).IndexOf(find, StringComparison.Ordinal) >= 0;
                return new SignatureTextReplacement(rtf, 0, present);
            }

            string escaped = EscapeText(replace ?? string.Empty, ansi);
            StringBuilder result = new StringBuilder(rtf.Length + accepted.Count * escaped.Length);
            int at = 0;
            foreach (Match match in accepted)
            {
                result.Append(rtf, at, match.Index - at).Append(escaped);
                at = match.Index + match.Length;
            }

            result.Append(rtf, at, rtf.Length - at);
            return new SignatureTextReplacement(result.ToString(), accepted.Count, false);
        }

        /// <summary>
        /// Text as RTF source: <c>\ \{ \}</c> escaped, line breaks as <c>\par</c>, tabs as <c>\tab</c>,
        /// characters the document's code page holds in one byte as <c>\'xx</c>, the rest as <c>\uN?</c>.
        /// </summary>
        public static string EscapeText(string text, Encoding? ansi = null)
        {
            Encoding codePage = StrictSingleByte(ansi ?? Encoding.GetEncoding(1252));
            StringBuilder rtf = new StringBuilder(text.Length + 16);
            string normalized = text.Replace("\r\n", "\n").Replace('\r', '\n');
            for (int i = 0; i < normalized.Length; i++)
            {
                char ch = normalized[i];
                switch (ch)
                {
                    case '\\':
                    case '{':
                    case '}':
                        rtf.Append('\\').Append(ch);
                        break;
                    case '\n':
                        rtf.Append(@"\par ");
                        break;
                    case '\t':
                        rtf.Append(@"\tab ");
                        break;
                    default:
                        if (ch < 0x80)
                        {
                            rtf.Append(ch);
                        }
                        else if (TryEncodeByte(codePage, ch, out byte b))
                        {
                            rtf.Append(@"\'").Append(b.ToString("x2", CultureInfo.InvariantCulture));
                        }
                        else
                        {
                            rtf.Append(@"\u").Append(((short)ch).ToString(CultureInfo.InvariantCulture)).Append('?');
                        }

                        break;
                }
            }

            return rtf.ToString();
        }

        /// <summary>
        /// A whole RTF document made 7-bit: binary (<c>\bin</c>) data kept as it is, every other character
        /// above 0x7F escaped. What <see cref="SignatureText.EncodeRtf"/> writes.
        /// </summary>
        public static string EscapeNonAscii(string rtf)
        {
            Encoding codePage = StrictSingleByte(DocumentEncoding(rtf));
            StringBuilder result = new StringBuilder(rtf.Length + 64);
            int i = 0;
            while (i < rtf.Length)
            {
                char c = rtf[i];
                if (c == '\\')
                {
                    ControlToken token = ReadControl(rtf, i);
                    result.Append(rtf, i, token.End - i);
                    i = token.End;
                    if (token.Word == "bin" && token.Parameter is int binary && binary > 0)
                    {
                        int take = Math.Min(binary, rtf.Length - i);
                        result.Append(rtf, i, take);
                        i += take;
                    }

                    continue;
                }

                if (c < 0x80)
                {
                    result.Append(c);
                }
                else if (TryEncodeByte(codePage, c, out byte b))
                {
                    result.Append(@"\'").Append(b.ToString("x2", CultureInfo.InvariantCulture));
                }
                else
                {
                    result.Append(@"\u").Append(((short)c).ToString(CultureInfo.InvariantCulture)).Append('?');
                }

                i++;
            }

            return result.ToString();
        }

        /// <summary>A minimal RTF document from plain text (Calibri 11), used only to create a missing rendition.</summary>
        public static string FromText(string text)
        {
            return Header + EscapeText(text ?? string.Empty) + "}";
        }

        /// <summary>
        /// A minimal RTF document from HTML - used only when a NEW signature is created without an RTF
        /// rendition, so Rich Text mail does not lose the signature. Keeps paragraphs and line breaks,
        /// bold, italic and underline, links (as hyperlink fields) and pictures (embedded from
        /// <paramref name="pictureFile"/>, PNG and JPEG); drops the rest. Never used on an existing
        /// signature: an update edits each rendition in place.
        /// </summary>
        public static string FromHtml(string html, Func<string, byte[]?> pictureFile)
        {
            if (html == null)
            {
                throw new ArgumentNullException(nameof(html));
            }

            return Header + HtmlToRtfBody(html, pictureFile) + "}";
        }

        private const string Header = @"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Calibri;}}\uc1\f0\fs22 ";

        // ------------------------------------------------------------------ the scanner

        private readonly struct GroupState
        {
            public GroupState(bool skipping, int unicodeSkip)
            {
                Skipping = skipping;
                UnicodeSkip = unicodeSkip;
            }

            public bool Skipping { get; }

            public int UnicodeSkip { get; }
        }

        private readonly struct ControlToken
        {
            public ControlToken(string? word, int? parameter, char? symbol, int end, int hexValue)
            {
                Word = word;
                Parameter = parameter;
                Symbol = symbol;
                End = end;
                HexValue = hexValue;
            }

            /// <summary>The control word's letters, or null for a control symbol.</summary>
            public string? Word { get; }

            public int? Parameter { get; }

            /// <summary>The control symbol's character (<c>'</c> for a hex escape), or null for a word.</summary>
            public char? Symbol { get; }

            /// <summary>The index just past the token, its delimiter space included.</summary>
            public int End { get; }

            /// <summary>For <c>\'xx</c>, the byte; otherwise -1.</summary>
            public int HexValue { get; }
        }

        /// <summary>Reads the control word or symbol at <paramref name="start"/> (which holds a backslash).</summary>
        private static ControlToken ReadControl(string rtf, int start)
        {
            int i = start + 1;
            if (i >= rtf.Length)
            {
                return new ControlToken(null, null, '\\', rtf.Length, -1);
            }

            char first = rtf[i];
            if (IsLetter(first))
            {
                int wordStart = i;
                while (i < rtf.Length && IsLetter(rtf[i]))
                {
                    i++;
                }

                string word = rtf.Substring(wordStart, i - wordStart);
                int? parameter = null;
                int numberStart = i;
                if (i < rtf.Length && (rtf[i] == '-' || char.IsDigit(rtf[i])))
                {
                    i++;
                    while (i < rtf.Length && char.IsDigit(rtf[i]))
                    {
                        i++;
                    }

                    if (int.TryParse(rtf.Substring(numberStart, i - numberStart), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int value))
                    {
                        parameter = value;
                    }
                }

                if (i < rtf.Length && rtf[i] == ' ')
                {
                    i++;
                }

                return new ControlToken(word, parameter, null, i, -1);
            }

            if (first == '\'')
            {
                if (i + 2 < rtf.Length
                    && int.TryParse(rtf.Substring(i + 1, 2), NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out int value))
                {
                    return new ControlToken(null, null, '\'', i + 3, value);
                }

                return new ControlToken(null, null, '\'', i + 1, -1);
            }

            return new ControlToken(null, null, first, i + 1, -1);
        }

        private static bool IsLetter(char c)
        {
            return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
        }

        /// <summary>What a control token contributes to the visible text, or null.</summary>
        private static string? Produced(ControlToken token, Encoding ansi)
        {
            if (token.Symbol == '\'' && token.HexValue >= 0)
            {
                return DecodeByte(ansi, (byte)token.HexValue).ToString();
            }

            if (token.Symbol != null)
            {
                switch (token.Symbol.Value)
                {
                    case '\\':
                    case '{':
                    case '}':
                        return token.Symbol.Value.ToString();
                    case '~':
                        return "\u00A0";
                    case '_':
                        return "\u2011";
                    case '\r':
                    case '\n':
                        return "\n";
                    default:
                        return null;
                }
            }

            switch (token.Word)
            {
                case "par":
                case "line":
                case "row":
                case "sect":
                case "page":
                    return "\n";
                case "tab":
                case "cell":
                    return "\t";
                case "emdash":
                    return "\u2014";
                case "endash":
                    return "\u2013";
                case "bullet":
                    return "\u2022";
                case "lquote":
                    return "\u2018";
                case "rquote":
                    return "\u2019";
                case "ldblquote":
                    return "\u201C";
                case "rdblquote":
                    return "\u201D";
                case "emspace":
                case "enspace":
                case "qmspace":
                    return " ";
                case "u":
                    return token.Parameter is int code ? ((char)(code < 0 ? code + 65536 : code)).ToString() : null;
                default:
                    return null;
            }
        }

        /// <summary>
        /// For every character of the RTF, whether it belongs to visible text in a way a replacement
        /// may rewrite: plain text characters, whole <c>\'xx</c> and <c>\uN?</c> escapes, and source line
        /// breaks between them. Control words, groups and skipped destinations are false.
        /// </summary>
        private static bool[] VisibleTextMap(string rtf)
        {
            bool[] visible = new bool[rtf.Length];
            Stack<GroupState> groups = new Stack<GroupState>();
            GroupState state = new GroupState(false, 1);
            int i = 0;
            while (i < rtf.Length)
            {
                char c = rtf[i];
                if (c == '{')
                {
                    groups.Push(state);
                    state = new GroupState(state.Skipping || (i + 2 < rtf.Length && rtf[i + 1] == '\\' && rtf[i + 2] == '*'), state.UnicodeSkip);
                    i++;
                    continue;
                }

                if (c == '}')
                {
                    state = groups.Count > 0 ? groups.Pop() : new GroupState(false, 1);
                    i++;
                    continue;
                }

                if (c == '\r' || c == '\n')
                {
                    visible[i] = !state.Skipping;
                    i++;
                    continue;
                }

                if (c == '\\')
                {
                    ControlToken token = ReadControl(rtf, i);
                    if (token.Word == "bin" && token.Parameter is int binary && binary > 0)
                    {
                        i = Math.Min(rtf.Length, token.End + binary);
                        continue;
                    }

                    if (token.Word != null && SkippedDestinations.Contains(token.Word))
                    {
                        state = new GroupState(true, state.UnicodeSkip);
                    }

                    bool textEscape = (token.Symbol == '\'' && token.HexValue >= 0)
                        || token.Symbol == '\\' || token.Symbol == '{' || token.Symbol == '}';
                    if (!state.Skipping && textEscape)
                    {
                        for (int k = i; k < token.End; k++)
                        {
                            visible[k] = true;
                        }
                    }
                    else if (!state.Skipping && token.Word == "u")
                    {
                        // \uN, its delimiter and its one fallback character (\'xx or a plain character).
                        int end = token.End;
                        if (end < rtf.Length && rtf[end] == '\\')
                        {
                            ControlToken fallback = ReadControl(rtf, end);
                            if (fallback.Symbol == '\'' && fallback.HexValue >= 0)
                            {
                                end = fallback.End;
                            }
                        }
                        else if (end < rtf.Length && rtf[end] != '{' && rtf[end] != '}')
                        {
                            end++;
                        }

                        for (int k = i; k < end; k++)
                        {
                            visible[k] = true;
                        }

                        i = end;
                        continue;
                    }

                    i = token.End;
                    continue;
                }

                visible[i] = !state.Skipping;
                i++;
            }

            return visible;
        }

        private static bool Covered(bool[] visible, int start, int length)
        {
            for (int k = start; k < start + length; k++)
            {
                if (!visible[k])
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>A regex matching <paramref name="find"/> as RTF source would spell it.</summary>
        private static string FindPattern(string find, Encoding ansi)
        {
            Encoding strict = StrictSingleByte(ansi);
            StringBuilder pattern = new StringBuilder(find.Length * 8);
            for (int i = 0; i < find.Length; i++)
            {
                if (i > 0)
                {
                    pattern.Append(@"(?:\r?\n)?");
                }

                char ch = find[i];
                switch (ch)
                {
                    case '\\':
                        pattern.Append(@"\\\\");
                        break;
                    case '{':
                        pattern.Append(@"\\\{");
                        break;
                    case '}':
                        pattern.Append(@"\\\}");
                        break;
                    default:
                        if (ch < 0x80)
                        {
                            pattern.Append(Regex.Escape(ch.ToString()));
                        }
                        else
                        {
                            List<string> forms = new List<string>
                            {
                                @"\\u" + ((short)ch).ToString(CultureInfo.InvariantCulture) + @" ?(?:\\'[0-9a-fA-F]{2}|[^\\{}\r\n])",
                            };
                            if (TryEncodeByte(strict, ch, out byte b))
                            {
                                forms.Add(@"\\'" + "(?i:" + b.ToString("x2", CultureInfo.InvariantCulture) + ")");
                                forms.Add(Regex.Escape(((char)b).ToString()));
                            }

                            pattern.Append("(?:").Append(string.Join("|", forms)).Append(')');
                        }

                        break;
                }
            }

            return pattern.ToString();
        }

        /// <summary>The document's ANSI code page (<c>\ansicpgN</c>), 1252 when it names none or one .NET lacks.</summary>
        private static Encoding DocumentEncoding(string rtf)
        {
            Match match = Regex.Match(rtf, @"\\ansicpg(\d+)", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
            if (match.Success && int.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int codePage))
            {
                try
                {
                    _ = SignatureText.AnsiCodePage; // registers the code-page provider on .NET
                    return Encoding.GetEncoding(codePage);
                }
                catch (Exception ex) when (ex is ArgumentException or NotSupportedException)
                {
                }
            }

            _ = SignatureText.AnsiCodePage;
            return Encoding.GetEncoding(1252);
        }

        private static Encoding StrictSingleByte(Encoding encoding)
        {
            _ = SignatureText.AnsiCodePage;
            return Encoding.GetEncoding(encoding.CodePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
        }

        private static bool TryEncodeByte(Encoding strict, char ch, out byte value)
        {
            value = 0;
            if (!strict.IsSingleByte || char.IsSurrogate(ch))
            {
                return false;
            }

            try
            {
                byte[] bytes = strict.GetBytes(new[] { ch });
                if (bytes.Length == 1)
                {
                    value = bytes[0];
                    return true;
                }
            }
            catch (EncoderFallbackException)
            {
            }

            return false;
        }

        private static char DecodeByte(Encoding ansi, byte value)
        {
            string decoded = ansi.GetString(new[] { value });
            return decoded.Length > 0 ? decoded[0] : '?';
        }

        /// <summary>The start of the next group whose first control word is <paramref name="word"/>, at or after <paramref name="from"/>.</summary>
        private static int FindGroupStartingWith(string rtf, int from, string word)
        {
            int i = from;
            while (i < rtf.Length)
            {
                char c = rtf[i];
                if (c == '\\')
                {
                    ControlToken token = ReadControl(rtf, i);
                    if (token.Word == "bin" && token.Parameter is int binary && binary > 0)
                    {
                        i = Math.Min(rtf.Length, token.End + binary);
                        continue;
                    }

                    i = token.End;
                    continue;
                }

                if (c == '{' && StartsGroupWith(rtf, i, word))
                {
                    return i;
                }

                i++;
            }

            return -1;
        }

        /// <summary>True when the group opening at <paramref name="open"/> begins with <c>\word</c> (after an optional <c>\*</c>).</summary>
        private static bool StartsGroupWith(string rtf, int open, string word)
        {
            if (open < 0 || open >= rtf.Length || rtf[open] != '{')
            {
                return false;
            }

            int i = open + 1;
            if (i + 1 < rtf.Length && rtf[i] == '\\' && rtf[i + 1] == '*')
            {
                i += 2;
                while (i < rtf.Length && (rtf[i] == '\r' || rtf[i] == '\n' || rtf[i] == ' '))
                {
                    i++;
                }
            }

            if (i >= rtf.Length || rtf[i] != '\\')
            {
                return false;
            }

            ControlToken token = ReadControl(rtf, i);
            return string.Equals(token.Word, word, StringComparison.Ordinal);
        }

        /// <summary>The index of the brace closing the group that opens at <paramref name="open"/>, or -1.</summary>
        private static int FindGroupEnd(string rtf, int open)
        {
            int depth = 0;
            int i = open;
            while (i < rtf.Length)
            {
                char c = rtf[i];
                if (c == '\\')
                {
                    ControlToken token = ReadControl(rtf, i);
                    if (token.Word == "bin" && token.Parameter is int binary && binary > 0)
                    {
                        i = Math.Min(rtf.Length, token.End + binary);
                        continue;
                    }

                    i = token.End;
                    continue;
                }

                if (c == '{')
                {
                    depth++;
                }
                else if (c == '}')
                {
                    depth--;
                    if (depth == 0)
                    {
                        return i;
                    }
                }

                i++;
            }

            return -1;
        }

        /// <summary>
        /// The <c>{\*\shppict</c> group directly enclosing a picture group - only when the picture is
        /// its sole content, so eliding the parent elides nothing else - or -1.
        /// </summary>
        private static int EnclosingShppict(string rtf, int pictStart, int pictEnd)
        {
            int i = pictStart - 1;
            while (i >= 0 && (rtf[i] == '\r' || rtf[i] == '\n' || rtf[i] == ' '))
            {
                i--;
            }

            // Expect "{\*\shppict" right before, possibly with whitespace.
            const string head = @"\shppict";
            int headStart = i - head.Length + 1;
            if (headStart < 0 || string.CompareOrdinal(rtf, headStart, head, 0, head.Length) != 0)
            {
                return -1;
            }

            int star = headStart - 1;
            while (star >= 0 && (rtf[star] == '\r' || rtf[star] == '\n' || rtf[star] == ' '))
            {
                star--;
            }

            if (star < 2 || rtf[star] != '*' || rtf[star - 1] != '\\' || rtf[star - 2] != '{')
            {
                return -1;
            }

            int open = star - 2;
            int close = FindGroupEnd(rtf, open);
            int after = SkipLineBreaks(rtf, pictEnd + 1);
            return close == after ? open : -1;
        }

        private static int SkipLineBreaks(string rtf, int from)
        {
            int i = from;
            while (i < rtf.Length && (rtf[i] == '\r' || rtf[i] == '\n' || rtf[i] == ' '))
            {
                i++;
            }

            return i;
        }

        /// <summary>The picture's bytes: the hex digits at the picture group's own level (nested groups skipped), or its \bin data.</summary>
        private static byte[] PictureBytes(string rtf, int pictStart, int pictEnd)
        {
            List<byte> bytes = new List<byte>();
            int depth = 0;
            int i = pictStart;
            int nibble = -1;
            while (i <= pictEnd && i < rtf.Length)
            {
                char c = rtf[i];
                if (c == '{')
                {
                    depth++;
                    i++;
                    continue;
                }

                if (c == '}')
                {
                    depth--;
                    i++;
                    continue;
                }

                if (c == '\\')
                {
                    ControlToken token = ReadControl(rtf, i);
                    if (token.Word == "bin" && token.Parameter is int binary && binary > 0 && depth == 1)
                    {
                        int take = Math.Min(binary, rtf.Length - token.End);
                        for (int k = 0; k < take; k++)
                        {
                            bytes.Add((byte)rtf[token.End + k]);
                        }

                        i = token.End + take;
                        continue;
                    }

                    i = token.End;
                    continue;
                }

                if (depth == 1)
                {
                    int value = HexValue(c);
                    if (value >= 0)
                    {
                        if (nibble < 0)
                        {
                            nibble = value;
                        }
                        else
                        {
                            bytes.Add((byte)((nibble << 4) | value));
                            nibble = -1;
                        }
                    }
                }

                i++;
            }

            return bytes.ToArray();
        }

        private static int HexValue(char c)
        {
            if (c >= '0' && c <= '9')
            {
                return c - '0';
            }

            if (c >= 'a' && c <= 'f')
            {
                return c - 'a' + 10;
            }

            if (c >= 'A' && c <= 'F')
            {
                return c - 'A' + 10;
            }

            return -1;
        }

        private static string PictureKind(string rtf, int pictStart, int pictEnd)
        {
            string group = rtf.Substring(pictStart, pictEnd - pictStart + 1);
            if (group.IndexOf(@"\pngblip", StringComparison.Ordinal) >= 0)
            {
                return "png";
            }

            if (group.IndexOf(@"\jpegblip", StringComparison.Ordinal) >= 0)
            {
                return "jpeg";
            }

            if (group.IndexOf(@"\emfblip", StringComparison.Ordinal) >= 0)
            {
                return "emf";
            }

            if (group.IndexOf(@"\wmetafile", StringComparison.Ordinal) >= 0)
            {
                return "wmf";
            }

            if (group.IndexOf(@"\dibitmap", StringComparison.Ordinal) >= 0 || group.IndexOf(@"\wbitmap", StringComparison.Ordinal) >= 0)
            {
                return "bitmap";
            }

            return "unknown";
        }

        private static int? ControlParameter(string rtf, int pictStart, int pictEnd, string word)
        {
            Match match = Regex.Match(
                rtf.Substring(pictStart, pictEnd - pictStart + 1), @"\\" + word + @"(-?\d+)", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
            return match.Success && int.TryParse(match.Groups[1].Value, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int value)
                ? value
                : (int?)null;
        }

        private static string ToHimetric(int pixels)
        {
            return ((int)Math.Round(pixels * 2540.0 / 96.0)).ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>SHA-256 of bytes, upper-case hex.</summary>
        public static string Sha256Hex(byte[] data)
        {
            using SHA256 sha = SHA256.Create();
            byte[] hash = sha.ComputeHash(data ?? Array.Empty<byte>());
            StringBuilder hex = new StringBuilder(hash.Length * 2);
            foreach (byte b in hash)
            {
                hex.Append(b.ToString("X2", CultureInfo.InvariantCulture));
            }

            return hex.ToString();
        }

        // ------------------------------------------------------------------ HTML to RTF (create only)

        private static readonly Regex HtmlToken = new Regex(
            @"<!--.*?-->|<![^>]*>|<(?<close>/)?(?<tag>[A-Za-z][A-Za-z0-9:]*)(?<attrs>[^>]*)>|(?<text>[^<]+)|<",
            RegexOptions.Singleline | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(5));

        private static readonly Regex HtmlAttribute = new Regex(
            @"(?<name>[A-Za-z_:][-A-Za-z0-9_:.]*)\s*=\s*(?:""(?<v>[^""]*)""|'(?<v>[^']*)'|(?<v>[^\s>]+))",
            RegexOptions.CultureInvariant, TimeSpan.FromSeconds(5));

        private static string HtmlToRtfBody(string html, Func<string, byte[]?> pictureFile)
        {
            StringBuilder rtf = new StringBuilder(html.Length);
            Stack<string> open = new Stack<string>();
            int skipDepth = 0;
            bool lastWasBreak = true;
            foreach (Match token in HtmlToken.Matches(html))
            {
                if (token.Value.StartsWith("<!", StringComparison.Ordinal))
                {
                    continue;
                }

                if (token.Groups["text"].Success)
                {
                    if (skipDepth > 0)
                    {
                        continue;
                    }

                    string text = Regex.Replace(WebUtility.HtmlDecode(token.Groups["text"].Value), @"\s+", " ", RegexOptions.None, TimeSpan.FromSeconds(2));
                    if (lastWasBreak)
                    {
                        text = text.TrimStart();
                    }

                    if (text.Length > 0)
                    {
                        rtf.Append(EscapeText(text));
                        lastWasBreak = false;
                    }

                    continue;
                }

                if (!token.Groups["tag"].Success)
                {
                    continue;
                }

                string tag = token.Groups["tag"].Value.ToLowerInvariant();
                bool closing = token.Groups["close"].Success;
                if (tag == "head" || tag == "style" || tag == "script" || tag == "title" || tag == "xml")
                {
                    skipDepth = Math.Max(0, skipDepth + (closing ? -1 : 1));
                    continue;
                }

                if (skipDepth > 0)
                {
                    continue;
                }

                switch (tag)
                {
                    case "br":
                        rtf.Append(@"\line ");
                        lastWasBreak = true;
                        break;
                    case "p":
                    case "div":
                    case "h1":
                    case "h2":
                    case "h3":
                    case "h4":
                    case "h5":
                    case "h6":
                    case "li":
                    case "tr":
                    case "table":
                        if (closing && !lastWasBreak)
                        {
                            rtf.Append(@"\par ");
                            lastWasBreak = true;
                        }

                        break;
                    case "td":
                    case "th":
                        if (closing)
                        {
                            rtf.Append(@"\tab ");
                        }

                        break;
                    case "b":
                    case "strong":
                    case "i":
                    case "em":
                    case "u":
                    case "a":
                        if (!closing)
                        {
                            string attrs = token.Groups["attrs"].Value;
                            string? href = tag == "a" ? Attribute(attrs, "href") : null;
                            if (href != null && (href.StartsWith("http", StringComparison.OrdinalIgnoreCase)
                                || href.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase)))
                            {
                                rtf.Append(@"{\field{\*\fldinst HYPERLINK """).Append(EscapeText(href).Replace("\"", "%22"))
                                    .Append(@"""}{\fldrslt {\ul ");
                                open.Push("}}}");
                            }
                            else
                            {
                                string control = tag == "b" || tag == "strong" ? @"{\b " : tag == "i" || tag == "em" ? @"{\i " : tag == "u" ? @"{\ul " : "{";
                                rtf.Append(control);
                                open.Push("}");
                            }
                        }
                        else if (open.Count > 0)
                        {
                            rtf.Append(open.Pop());
                        }

                        break;
                    case "img":
                        string? src = Attribute(token.Groups["attrs"].Value, "src");
                        byte[]? image = src == null ? null : pictureFile(src);
                        if (image != null && (SignaturePictures.IsPng(image) || SignaturePictures.IsJpeg(image)))
                        {
                            int? width = PixelAttribute(token.Groups["attrs"].Value, "width");
                            int? height = PixelAttribute(token.Groups["attrs"].Value, "height");
                            rtf.Append(PictureGroup(image, width, height));
                            lastWasBreak = false;
                        }

                        break;
                }
            }

            while (open.Count > 0)
            {
                rtf.Append(open.Pop());
            }

            return rtf.ToString();
        }

        private static string? Attribute(string attributes, string name)
        {
            foreach (Match attribute in HtmlAttribute.Matches(attributes))
            {
                if (string.Equals(attribute.Groups["name"].Value, name, StringComparison.OrdinalIgnoreCase))
                {
                    return WebUtility.HtmlDecode(attribute.Groups["v"].Value);
                }
            }

            return null;
        }

        private static int? PixelAttribute(string attributes, string name)
        {
            string? value = Attribute(attributes, name);
            return value != null && int.TryParse(value.Trim().Replace("px", string.Empty), NumberStyles.Integer, CultureInfo.InvariantCulture, out int pixels)
                && pixels > 0
                ? pixels
                : (int?)null;
        }
    }
}
