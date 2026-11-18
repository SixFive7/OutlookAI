using System;
using System.Globalization;
using System.Security.Cryptography;

namespace OutlookAI.Core.AddIn
{
    /// <summary>
    /// Hexadecimal, read and written - what the server and the add-in each wrote over and over until
    /// the code-sharing refactor (Q154): stage S6 folded four checks of a string's digits, one of a
    /// single character and two writers of a random token into it, and the hex stage after it nine
    /// writers of bytes, six of a SHA-256, two of a number and three readers. Every writer
    /// names its case, and each caller keeps the case its readers compare: uppercase for an id, as
    /// Outlook writes an entry id, and lowercase for the hashes and tokens that were lowercase before.
    /// </summary>
    public static class Hex
    {
        /// <summary>Whether <paramref name="c"/> is a hexadecimal digit, in either case.</summary>
        public static bool IsDigit(char c) => DigitValue(c) >= 0;

        /// <summary>
        /// Whether <paramref name="c"/> is one of the sixteen lowercase hexadecimal digits: all a token
        /// or a CRC this code wrote holds, so a reader of one refuses an uppercase digit.
        /// </summary>
        public static bool IsLowerDigit(char c) => (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f');

        /// <summary>
        /// Whether every character of <paramref name="value"/> is a hexadecimal digit. True for an
        /// empty string: a caller that needs at least one digit, or an even count, says so itself.
        /// </summary>
        public static bool IsAllDigits(string value)
        {
            foreach (char c in value)
            {
                if (!IsDigit(c))
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>The value of a hexadecimal digit in either case, 0 to 15; -1 for any other character.</summary>
        public static int DigitValue(char c)
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

        /// <summary>
        /// <paramref name="bytes"/> as uppercase hexadecimal, two digits a byte: the case Outlook writes an
        /// entry id in, and every id this code makes of bytes.
        /// </summary>
        public static string Upper(byte[] bytes) => Write(bytes, 0, bytes.Length, 'A');

        /// <summary>
        /// <paramref name="count"/> bytes of <paramref name="bytes"/> from <paramref name="offset"/>, as
        /// uppercase hexadecimal.
        /// </summary>
        public static string Upper(byte[] bytes, int offset, int count) => Write(bytes, offset, count, 'A');

        /// <summary><paramref name="bytes"/> as lowercase hexadecimal, two digits a byte.</summary>
        public static string Lower(byte[] bytes) => Write(bytes, 0, bytes.Length, 'a');

        /// <summary>
        /// <paramref name="count"/> bytes of <paramref name="bytes"/> from <paramref name="offset"/>, as
        /// lowercase hexadecimal.
        /// </summary>
        public static string Lower(byte[] bytes, int offset, int count) => Write(bytes, offset, count, 'a');

        /// <summary>
        /// <paramref name="value"/> as lowercase hexadecimal, padded with zeros to
        /// <paramref name="digits"/> digits: a CRC-32 in eight, a file's id in sixteen.
        /// </summary>
        public static string Lower(ulong value, int digits) =>
            value.ToString("x" + digits.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);

        /// <summary>The SHA-256 of <paramref name="data"/>, as uppercase hexadecimal.</summary>
        public static string Sha256Upper(byte[] data) => Upper(Sha256(data));

        /// <summary>The SHA-256 of <paramref name="data"/>, as lowercase hexadecimal.</summary>
        public static string Sha256Lower(byte[] data) => Lower(Sha256(data));

        /// <summary>
        /// Reads hexadecimal back into bytes, two digits a byte, in either case. False - and no bytes -
        /// for null, an empty string, an odd number of characters, or any character that is not a
        /// digit, a space included, so a value that would not read is not matched on garbage.
        /// </summary>
        public static bool TryDecode(string? hex, out byte[] bytes)
        {
            bytes = Array.Empty<byte>();
            if (string.IsNullOrEmpty(hex) || (hex!.Length % 2) != 0)
            {
                return false;
            }

            byte[] result = new byte[hex.Length / 2];
            for (int i = 0; i < result.Length; i++)
            {
                int high = DigitValue(hex[i * 2]);
                int low = DigitValue(hex[(i * 2) + 1]);
                if (high < 0 || low < 0)
                {
                    return false;
                }

                result[i] = (byte)((high << 4) | low);
            }

            bytes = result;
            return true;
        }

        /// <summary>
        /// <paramref name="byteCount"/> random bytes from the system's cryptographic generator,
        /// written as lowercase hexadecimal - twice as many characters. What a send confirmation
        /// and a resumable scan's tokens are made of.
        /// </summary>
        public static string NewRandom(int byteCount)
        {
            byte[] bytes = new byte[byteCount];
            using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(bytes);
            }

            return Lower(bytes);
        }

        private static byte[] Sha256(byte[] data)
        {
            using (SHA256 sha = SHA256.Create())
            {
                return sha.ComputeHash(data);
            }
        }

        private static string Write(byte[] bytes, int offset, int count, char letterA)
        {
            char[] hex = new char[count * 2];
            for (int i = 0; i < count; i++)
            {
                int b = bytes[offset + i];
                hex[i * 2] = Digit(b >> 4, letterA);
                hex[(i * 2) + 1] = Digit(b & 0xF, letterA);
            }

            return new string(hex);
        }

        private static char Digit(int value, char letterA) => (char)(value < 10 ? '0' + value : letterA + (value - 10));
    }
}
