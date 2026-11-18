using System.Security.Cryptography;

namespace OutlookAI.Core.AddIn
{
    /// <summary>
    /// Hexadecimal digits, read and written - what the server and the add-in each wrote several
    /// times over until the code-sharing refactor's stage S6 (Q154): four checks of a string's
    /// digits, one of a single character, and two writers of a random token.
    /// </summary>
    public static class Hex
    {
        /// <summary>Whether <paramref name="c"/> is a hexadecimal digit, in either case.</summary>
        public static bool IsDigit(char c) =>
            (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');

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

            char[] hex = new char[bytes.Length * 2];
            for (int i = 0; i < bytes.Length; i++)
            {
                hex[i * 2] = LowercaseDigit(bytes[i] >> 4);
                hex[(i * 2) + 1] = LowercaseDigit(bytes[i] & 0xF);
            }

            return new string(hex);
        }

        private static char LowercaseDigit(int value) => (char)(value < 10 ? '0' + value : 'a' + (value - 10));
    }
}
