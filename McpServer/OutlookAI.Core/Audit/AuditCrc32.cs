using System;

namespace OutlookAI.Core.Audit
{
    /// <summary>
    /// CRC-32 (IEEE 802.3: reflected, polynomial 0xEDB88320, initial and final value 0xFFFFFFFF -
    /// the zlib/PNG/Ethernet one) for the audit log's per-line <c>crc</c> field (Q117), and as the
    /// anchor a resume position checks its line against.
    /// <para>
    /// Written out rather than taken from <c>System.IO.Hashing</c>: that is a NuGet package, this is
    /// twenty lines, and the same code has to run unchanged in the .NET Framework 4.8 add-in that will
    /// write this log too. A checksum, not a signature - it detects a line that was cut short,
    /// overwritten or spliced, not one forged on purpose by someone who can edit the file anyway.
    /// </para>
    /// </summary>
    internal static class AuditCrc32
    {
        private static readonly uint[] Table = BuildTable();

        /// <summary>The CRC-32 of <paramref name="count"/> bytes of <paramref name="bytes"/> from <paramref name="offset"/>.</summary>
        internal static uint Compute(byte[] bytes, int offset, int count)
        {
            if (bytes == null)
            {
                throw new ArgumentNullException(nameof(bytes));
            }

            if (offset < 0 || count < 0 || offset > bytes.Length - count)
            {
                throw new ArgumentOutOfRangeException(nameof(count));
            }

            uint crc = 0xFFFFFFFFu;
            for (int i = offset; i < offset + count; i++)
            {
                crc = Table[(crc ^ bytes[i]) & 0xFF] ^ (crc >> 8);
            }

            return ~crc;
        }

        /// <summary>Eight lowercase hex digits, the form a line carries.</summary>
        internal static string ToHex(uint crc)
        {
            char[] hex = new char[8];
            for (int i = 7; i >= 0; i--)
            {
                int nibble = (int)(crc & 0xF);
                hex[i] = (char)(nibble < 10 ? '0' + nibble : 'a' + (nibble - 10));
                crc >>= 4;
            }

            return new string(hex);
        }

        /// <summary>Parses exactly eight lowercase hex digits; anything else is not a value the writer wrote.</summary>
        internal static bool TryParseHex(string text, out uint crc)
        {
            crc = 0;
            if (text == null || text.Length != 8)
            {
                return false;
            }

            foreach (char c in text)
            {
                int nibble;
                if (c >= '0' && c <= '9')
                {
                    nibble = c - '0';
                }
                else if (c >= 'a' && c <= 'f')
                {
                    nibble = c - 'a' + 10;
                }
                else
                {
                    return false;
                }

                crc = (crc << 4) | (uint)nibble;
            }

            return true;
        }

        private static uint[] BuildTable()
        {
            uint[] table = new uint[256];
            for (uint n = 0; n < 256; n++)
            {
                uint c = n;
                for (int k = 0; k < 8; k++)
                {
                    c = (c & 1) != 0 ? 0xEDB88320u ^ (c >> 1) : c >> 1;
                }

                table[n] = c;
            }

            return table;
        }
    }
}
