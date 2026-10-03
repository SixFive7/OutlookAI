using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// The paging every tool shares (Q119, 2026-10-03). Three pieces, because the tools page three
    /// different things, and one rule over all of them: the order is stable, the continuation is
    /// deterministic, and "there is more" is said explicitly rather than inferred from a short page.
    /// <list type="bullet">
    /// <item><see cref="PageWindow"/> - an offset window over an order that is fully known when the
    /// page is cut: <c>list_folders</c> (<c>offset</c> / <c>nextOffset</c>) and <c>read</c>'s body
    /// (<c>body_offset</c> / <c>bodyTruncated</c>).</item>
    /// <item><see cref="PagingFingerprint"/> - the canonical text of the arguments that decide WHICH
    /// results a chain pages through, so a continuation asked with different arguments is refused by
    /// name instead of quietly answering another question: <c>search</c>'s exhaustive
    /// <c>resume_token</c> and <c>audit_log</c>'s.</item>
    /// <item><see cref="IssueToken"/> / <see cref="ResolveToken"/> - a STATELESS continuation token:
    /// the position and that fingerprint, sealed with a checksum, so any server process can continue
    /// any chain and a restart loses nothing. <c>audit_log</c> uses it (its position is a byte offset
    /// plus the file's identity). <c>search</c>'s exhaustive chain keeps its server-side session,
    /// because what it continues is a live Outlook cursor that cannot be written into a string.</item>
    /// </list>
    /// </summary>
    public static class Paging
    {
        /// <summary>What every token this helper issues starts with.</summary>
        public const string TokenPrefix = "pg1.";

        private const char PartSeparator = '\u001e';
        private const char ValueSeparator = '\u001f';

        /// <summary>
        /// A continuation token for the page after this one: <paramref name="kind"/> (the tool),
        /// the <see cref="PagingFingerprint"/> of the arguments the chain was opened with - kept as a
        /// hash per argument, so the token carries no argument values - and the position values,
        /// which mean something only to the tool that issued them. Opaque to the caller by contract;
        /// base64url, checksummed, so a damaged token is told apart from one for other arguments.
        /// </summary>
        public static string IssueToken(string kind, string fingerprint, IReadOnlyList<string> position)
        {
            if (string.IsNullOrEmpty(kind) || kind.IndexOf(PartSeparator) >= 0 || kind.IndexOf(ValueSeparator) >= 0)
            {
                throw new ArgumentException("A token kind must be a plain, non-empty name.", nameof(kind));
            }

            if (position == null)
            {
                throw new ArgumentNullException(nameof(position));
            }

            foreach (string value in position)
            {
                if (value == null || value.IndexOf(PartSeparator) >= 0 || value.IndexOf(ValueSeparator) >= 0)
                {
                    throw new ArgumentException("Position values must be non-null and free of separators.", nameof(position));
                }
            }

            string payload = kind + PartSeparator + PagingFingerprint.Compact(fingerprint ?? string.Empty)
                + PartSeparator + string.Join(ValueSeparator.ToString(), position);
            return TokenPrefix + ToBase64Url(Encoding.UTF8.GetBytes(payload)) + "." + Checksum(payload);
        }

        /// <summary>
        /// Reads a token <see cref="IssueToken"/> made. <see cref="PageTokenDecision.Valid"/> hands back
        /// the position; <see cref="PageTokenDecision.RequestChanged"/> names the arguments that differ
        /// from the ones the chain was opened with (in fingerprint order); the other two say the string
        /// is not a token of this kind at all. Nothing is stored anywhere, so nothing is consumed: the
        /// same token can be used again, and gives the same page if the data has not changed.
        /// </summary>
        public static PageTokenDecision ResolveToken(
            string? token,
            string kind,
            string fingerprint,
            out IReadOnlyList<string> position,
            out IReadOnlyList<string> changedArguments)
        {
            position = Array.Empty<string>();
            changedArguments = Array.Empty<string>();
            if (token == null || !token.StartsWith(TokenPrefix, StringComparison.Ordinal))
            {
                return PageTokenDecision.Malformed;
            }

            int dot = token.LastIndexOf('.');
            if (dot <= TokenPrefix.Length - 1)
            {
                return PageTokenDecision.Malformed;
            }

            string payload;
            try
            {
                payload = new UTF8Encoding(false, true).GetString(
                    FromBase64Url(token.Substring(TokenPrefix.Length, dot - TokenPrefix.Length)));
            }
            catch (FormatException)
            {
                return PageTokenDecision.Malformed;
            }
            catch (ArgumentException)
            {
                return PageTokenDecision.Malformed;
            }

            if (!string.Equals(Checksum(payload), token.Substring(dot + 1), StringComparison.Ordinal))
            {
                return PageTokenDecision.Malformed;
            }

            string[] parts = payload.Split(PartSeparator);
            if (parts.Length != 3)
            {
                return PageTokenDecision.Malformed;
            }

            if (!string.Equals(parts[0], kind, StringComparison.Ordinal))
            {
                return PageTokenDecision.OtherKind;
            }

            string current = PagingFingerprint.Compact(fingerprint ?? string.Empty);
            if (!string.Equals(parts[1], current, StringComparison.Ordinal))
            {
                changedArguments = PagingFingerprint.DifferingArguments(parts[1], current);
                return PageTokenDecision.RequestChanged;
            }

            position = parts[2].Length == 0 ? Array.Empty<string>() : parts[2].Split(ValueSeparator);
            return PageTokenDecision.Valid;
        }

        /// <summary>The first eight hex digits of the SHA-256 of <paramref name="text"/>'s UTF-8.</summary>
        internal static string ShortHash(string text)
        {
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(text ?? string.Empty));
                StringBuilder hex = new StringBuilder(8);
                for (int i = 0; i < 4; i++)
                {
                    hex.Append(hash[i].ToString("x2", CultureInfo.InvariantCulture));
                }

                return hex.ToString();
            }
        }

        private static string Checksum(string payload)
        {
            return ShortHash("pg1|" + payload);
        }

        private static string ToBase64Url(byte[] bytes)
        {
            return Convert.ToBase64String(bytes).TrimEnd('=').Replace('+', '-').Replace('/', '_');
        }

        private static byte[] FromBase64Url(string text)
        {
            string standard = text.Replace('-', '+').Replace('_', '/');
            switch (standard.Length % 4)
            {
                case 2:
                    standard += "==";
                    break;
                case 3:
                    standard += "=";
                    break;
                case 1:
                    throw new FormatException("Not base64url.");
            }

            return Convert.FromBase64String(standard);
        }
    }

    /// <summary>What <see cref="Paging.ResolveToken"/> made of a token.</summary>
    public enum PageTokenDecision
    {
        /// <summary>A token of this kind for these arguments; its position is usable.</summary>
        Valid,

        /// <summary>Not a token this helper issued, or one that was damaged on the way.</summary>
        Malformed,

        /// <summary>A sound token, but another tool's.</summary>
        OtherKind,

        /// <summary>A token of this kind, issued for different arguments.</summary>
        RequestChanged,
    }

    /// <summary>
    /// One window of an order that is fully known when the page is cut: items (or characters)
    /// [<see cref="Start"/>, <see cref="End"/>) of <see cref="Total"/>, with the has-more contract
    /// derived in one place. The offset is clamped into [0, total] and the size floored at 0, which is
    /// what <c>list_folders</c> and <c>read</c> did each in their own words before (Q119).
    /// </summary>
    public readonly struct PageWindow
    {
        private PageWindow(int start, int count, int total)
        {
            Start = start;
            Count = count;
            Total = total;
        }

        /// <summary>Index of the first item in the window.</summary>
        public int Start { get; }

        /// <summary>How many items the window holds.</summary>
        public int Count { get; }

        /// <summary>How many items the whole order holds.</summary>
        public int Total { get; }

        /// <summary>Index just past the window.</summary>
        public int End => Start + Count;

        /// <summary>Whether items exist beyond the window - the explicit "more" signal.</summary>
        public bool HasMore => End < Total;

        /// <summary>The offset that continues the order, or null when the window reaches its end.</summary>
        public int? NextOffset => HasMore ? End : (int?)null;

        /// <summary>The window of at most <paramref name="size"/> items at <paramref name="offset"/> in an order of <paramref name="total"/>.</summary>
        public static PageWindow Of(int total, int offset, int size)
        {
            if (total < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(total), "total must not be negative.");
            }

            int start = offset < 0 ? 0 : (offset > total ? total : offset);
            int count = size <= 0 ? 0 : Math.Min(size, total - start);
            return new PageWindow(start, count, total);
        }
    }

    /// <summary>
    /// The canonical text of everything that decides WHICH results a paged chain returns, one line
    /// per argument, in a fixed order and presence-first: <c>label=0</c> for an argument that was not
    /// given, <c>label=1|value</c> for one that was - so an ABSENT argument and an empty one never hash
    /// alike, and no sentinel a caller could type can stand in for absence. Arguments that only shape a
    /// page (<c>top</c>, snippet sizes) are left out on purpose: a caller may change them mid-chain.
    /// Lifted out of <c>ExhaustiveScanCursors</c> (Q119) so <c>search</c> and <c>audit_log</c> refuse a
    /// changed continuation the same way, naming what changed.
    /// </summary>
    public sealed class PagingFingerprint
    {
        private readonly StringBuilder _canonical = new StringBuilder(256);

        /// <summary>One argument: absent when <paramref name="value"/> is null.</summary>
        public PagingFingerprint Add(string label, string? value)
        {
            _canonical.Append(label).Append('=');
            if (value == null)
            {
                _canonical.Append('0');
            }
            else
            {
                _canonical.Append('1').Append('|').Append(value);
            }

            _canonical.Append('\n');
            return this;
        }

        /// <summary>A list argument: absent when null; otherwise its count and its values in order.</summary>
        public PagingFingerprint AddList(string label, IReadOnlyList<string>? values)
        {
            if (values == null)
            {
                _canonical.Append(label).Append("=0").Append('\n');
                return this;
            }

            _canonical.Append(label).Append("=1|").Append(values.Count.ToString(CultureInfo.InvariantCulture));
            for (int i = 0; i < values.Count; i++)
            {
                _canonical.Append('|').Append(values[i] ?? string.Empty);
            }

            _canonical.Append('\n');
            return this;
        }

        /// <summary>The canonical text.</summary>
        public override string ToString()
        {
            return _canonical.ToString();
        }

        /// <summary>
        /// The same lines with every present value replaced by a short hash: still presence-first and
        /// still labelled, so <see cref="DifferingArguments"/> works on it, but it carries no argument
        /// values - what a token that leaves the server may hold.
        /// </summary>
        public static string Compact(string canonical)
        {
            string[] lines = (canonical ?? string.Empty).Split('\n');
            StringBuilder compact = new StringBuilder(canonical?.Length ?? 0);
            foreach (string line in lines)
            {
                if (line.Length == 0)
                {
                    continue;
                }

                int equals = line.IndexOf('=');
                if (equals > 0 && line.Length > equals + 2 && line[equals + 1] == '1' && line[equals + 2] == '|')
                {
                    compact.Append(line, 0, equals + 3).Append(Paging.ShortHash(line.Substring(equals + 3)));
                }
                else
                {
                    compact.Append(line);
                }

                compact.Append('\n');
            }

            return compact.ToString();
        }

        /// <summary>
        /// The argument labels two fingerprints disagree on, in fingerprint order. Empty when
        /// they agree. It is what lets a refusal name the thing the caller changed instead of
        /// asserting, unhelpfully, that something did.
        /// </summary>
        public static IReadOnlyList<string> DifferingArguments(string expected, string actual)
        {
            string[] left = (expected ?? string.Empty).Split('\n');
            string[] right = (actual ?? string.Empty).Split('\n');
            List<string> changed = new List<string>();
            int max = left.Length > right.Length ? left.Length : right.Length;
            for (int i = 0; i < max; i++)
            {
                string a = i < left.Length ? left[i] : string.Empty;
                string b = i < right.Length ? right[i] : string.Empty;
                if (string.Equals(a, b, StringComparison.Ordinal))
                {
                    continue;
                }

                string label = LabelOf(a.Length > 0 ? a : b);
                if (label.Length > 0 && !changed.Contains(label))
                {
                    changed.Add(label);
                }
            }

            return changed;
        }

        private static string LabelOf(string canonicalLine)
        {
            int equals = canonicalLine.IndexOf('=');
            return equals > 0 ? canonicalLine.Substring(0, equals) : string.Empty;
        }
    }
}
