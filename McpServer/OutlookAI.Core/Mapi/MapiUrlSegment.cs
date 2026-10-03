using System;
using System.Text;

namespace OutlookAI.Core.Mapi
{
    /// <summary>
    /// The percent-encoding Outlook applies to a store or folder NAME when it writes the name
    /// into a Windows Search MAPI URL - and its inverse, for reading a name back out of one.
    /// <para>
    /// Microsoft documents five characters, "encoded if they are in the store or folder display
    /// name": <c>%</c> as <c>%25</c>, <c>/</c> as <c>%2F</c>, <c>\</c> as <c>%5C</c>, <c>*</c> as
    /// <c>%2A</c> and <c>?</c> as <c>%3F</c> (<i>About MAPI URLs for Notification-Based
    /// Indexing</i>, MAPI reference, "Special characters"; the original Outlook 2007 note by
    /// Stephen Griffin says "store or folder display name" in so many words). Nothing else is
    /// encoded - a space, an apostrophe and every non-ASCII letter stay literal, because MAPI
    /// URLs are Unicode.
    /// </para>
    /// <para>
    /// MEASURED, both halves, on the indexed test guest. A STORE named <c>q99 50% off*?x</c> is
    /// filed as <c>q99 50%25 off%2A%3Fx($5159380d)</c> (Q99, 2026-10-03,
    /// <c>Docs/live-tier-on-the-vm.md</c> section 8 item 24). FOLDERS the same day
    /// (<c>T2/LiveFolderNameEncodingTests</c>, section 8 item 26): Outlook accepts all five in a
    /// folder name, the URL spells each exactly as above - <c>50%25 off</c>, <c>star%2A</c>,
    /// <c>why%3F</c>, <c>back%5Cslash</c>, <c>a%2Fb</c>, <c>%252A not a star</c> - and
    /// <c>System.ItemFolderPathDisplay</c> and <c>System.ItemPathDisplay</c> hold the NAMES, never
    /// this spelling.
    /// </para>
    /// <para>
    /// ⚠ WHY THIS EXISTS. A folder scope used to be built from the raw name -
    /// <c>storePrefix + "/0/" + folder</c> - so a folder named <c>50% off</c> was scoped as
    /// <c>.../0/50% off</c> against an index that files it as <c>.../0/50%25 off</c>: a SCOPE that
    /// addresses nothing, zero rows with no error, and only the freshness sweep's last few days to
    /// hide it. And in the other direction a hit from that folder carried <c>50%25 off</c> as its
    /// folder name - the name no folder in Outlook has, so the URL route to open it failed and the
    /// sweep's copy of the same mail was never recognised as a duplicate.
    /// </para>
    /// <para>
    /// Both directions are exact inverses, character by character, in ONE pass each - never a
    /// chain of <c>Replace</c> calls, which turns a name that merely looks like an escape (a
    /// folder literally called <c>%2A</c>, filed as <c>%252A</c>) into a different name. A
    /// segment holding none of the five characters is returned unchanged, so every ordinary
    /// name produces byte for byte the URL it always did.
    /// </para>
    /// </summary>
    public static class MapiUrlSegment
    {
        /// <summary>The five characters Outlook percent-encodes in a store or folder name, in Microsoft's order.</summary>
        public const string EncodedCharacters = "%/\\*?";

        /// <summary>The five, as an array for <see cref="string.IndexOfAny(char[])"/>.</summary>
        private static readonly char[] Encoded = EncodedCharacters.ToCharArray();

        /// <summary>The four of them a segment of a <c>/</c>-separated path can hold.</summary>
        private static readonly char[] InSegmentCharacters = { '%', '\\', '*', '?' };

        /// <summary>
        /// A name as the index spells it inside a MAPI URL: each of
        /// <see cref="EncodedCharacters"/> replaced by <c>%</c> and its two UPPER-case hex digits
        /// (the spelling measured: <c>%2A</c>, <c>%3F</c>), every other character as it is.
        /// </summary>
        public static string Encode(string name)
        {
            if (name == null)
            {
                throw new ArgumentNullException(nameof(name));
            }

            if (name.IndexOfAny(Encoded) < 0)
            {
                return name;
            }

            StringBuilder encoded = new StringBuilder(name.Length + 8);
            foreach (char c in name)
            {
                switch (c)
                {
                    case '%':
                        encoded.Append("%25");
                        break;
                    case '/':
                        encoded.Append("%2F");
                        break;
                    case '\\':
                        encoded.Append("%5C");
                        break;
                    case '*':
                        encoded.Append("%2A");
                        break;
                    case '?':
                        encoded.Append("%3F");
                        break;
                    default:
                        encoded.Append(c);
                        break;
                }
            }

            return encoded.ToString();
        }

        /// <summary>
        /// A store-relative folder PATH as the index spells it: every <c>/</c>-separated segment
        /// through <see cref="Encode"/>, the separators kept. The product's folder paths are
        /// <c>/</c>-separated (<c>Inbox/Fun</c>), so a segment of one never holds a <c>/</c> and the
        /// <c>%2F</c> escape cannot arise here; it is the other four that do. Empty segments are kept
        /// where they are, so a path with none of the five characters comes back unchanged.
        /// </summary>
        public static string EncodePath(string path)
        {
            if (path == null)
            {
                throw new ArgumentNullException(nameof(path));
            }

            if (path.IndexOfAny(InSegmentCharacters) < 0)
            {
                return path;
            }

            string[] segments = path.Split('/');
            for (int i = 0; i < segments.Length; i++)
            {
                segments[i] = Encode(segments[i]);
            }

            return string.Join("/", segments);
        }

        /// <summary>
        /// The name a MAPI URL segment spells: each of the five escapes <see cref="Encode"/>
        /// writes - its hex digits in either case - turned back into its character, in one
        /// left-to-right pass. Anything else is kept exactly as it is, including a <c>%</c> that
        /// does not begin one of those five escapes: this inverts what Outlook is documented and
        /// measured to write, and guesses at nothing else.
        /// </summary>
        public static string Decode(string segment)
        {
            if (segment == null)
            {
                throw new ArgumentNullException(nameof(segment));
            }

            int first = segment.IndexOf('%');
            if (first < 0)
            {
                return segment;
            }

            StringBuilder decoded = new StringBuilder(segment.Length);
            decoded.Append(segment, 0, first);
            int i = first;
            while (i < segment.Length)
            {
                char c = segment[i];
                if (c == '%' && i + 2 < segment.Length && TryDecodeEscape(segment[i + 1], segment[i + 2], out char escaped))
                {
                    decoded.Append(escaped);
                    i += 3;
                    continue;
                }

                decoded.Append(c);
                i++;
            }

            return decoded.ToString();
        }

        private static bool TryDecodeEscape(char high, char low, out char escaped)
        {
            escaped = '\0';
            switch (char.ToUpperInvariant(high))
            {
                case '2':
                    switch (char.ToUpperInvariant(low))
                    {
                        case '5':
                            escaped = '%';
                            return true;
                        case 'F':
                            escaped = '/';
                            return true;
                        case 'A':
                            escaped = '*';
                            return true;
                        default:
                            return false;
                    }

                case '3':
                    if (char.ToUpperInvariant(low) == 'F')
                    {
                        escaped = '?';
                        return true;
                    }

                    return false;

                case '5':
                    if (char.ToUpperInvariant(low) == 'C')
                    {
                        escaped = '\\';
                        return true;
                    }

                    return false;

                default:
                    return false;
            }
        }
    }
}
