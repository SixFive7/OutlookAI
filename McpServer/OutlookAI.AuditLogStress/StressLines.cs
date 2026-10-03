using System.Text;

namespace OutlookAI.AuditLogStress
{
    /// <summary>
    /// What a stress writer puts in each line, derived from (seed, writer, n) alone - so the test that
    /// reads the log back can recompute every line it expects and compare byte for byte, without the
    /// writers reporting anything. Compiled into the stress writer AND linked into the test project,
    /// one copy of the rule for both sides.
    /// <para>
    /// Lines up to 64 KB (Q117): every 25th payload is 60,000 to 64,000 bytes once escaped and encoded
    /// - with the line's other fields and trailer that stays under 65,536 - and the rest 40 to 3,040.
    /// The characters are drawn from a palette that exercises every escape the writer performs and
    /// UTF-8 sequences of two and three bytes, so a byte lost or moved anywhere shows.
    /// </para>
    /// </summary>
    internal static class StressLines
    {
        /// <summary>The operation every stress line records.</summary>
        internal const string Operation = "stress_append";

        /// <summary>The largest a stress line gets, terminator included.</summary>
        internal const int MaxLineBytes = 64 * 1024;

        private const string Palette =
            "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 .,;:-_/\\\"\t\né日";

        /// <summary>The payload writer <paramref name="writer"/> puts in its line number <paramref name="n"/>.</summary>
        internal static string Payload(int seed, string writer, long n)
        {
            uint state = Mix(seed, writer, n);
            int target = n % 25 == 0
                ? 60000 + (int)(Next(ref state) % 4000)
                : 40 + (int)(Next(ref state) % 3000);
            StringBuilder payload = new StringBuilder(target);
            int bytes = 0;
            while (bytes < target)
            {
                char c = Palette[(int)(Next(ref state) % (uint)Palette.Length)];
                payload.Append(c);
                bytes += EncodedLength(c);
            }

            return payload.ToString();
        }

        /// <summary>How many bytes <paramref name="c"/> takes in a line once escaped and UTF-8 encoded.</summary>
        private static int EncodedLength(char c)
        {
            switch (c)
            {
                case '\\':
                case '"':
                case '\t':
                case '\n':
                    return 2;
                default:
                    return c < 0x80 ? 1 : (c < 0x800 ? 2 : 3);
            }
        }

        private static uint Mix(int seed, string writer, long n)
        {
            // FNV-1a over the three inputs; never 0, which xorshift cannot leave.
            uint hash = 2166136261u;
            hash = Fold(hash, unchecked((uint)seed));
            foreach (char c in writer)
            {
                hash = Fold(hash, c);
            }

            hash = Fold(hash, unchecked((uint)n));
            hash = Fold(hash, unchecked((uint)(n >> 32)));
            return hash == 0 ? 0x9E3779B9u : hash;
        }

        private static uint Fold(uint hash, uint value)
        {
            for (int i = 0; i < 4; i++)
            {
                hash ^= (value >> (8 * i)) & 0xFF;
                hash = unchecked(hash * 16777619u);
            }

            return hash;
        }

        private static uint Next(ref uint state)
        {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            return state;
        }
    }
}
