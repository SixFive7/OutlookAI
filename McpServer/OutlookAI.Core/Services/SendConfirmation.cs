using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;

namespace OutlookAI.Core.Services
{
    /// <summary>Outcome of one confirm-token consumption attempt (v3.MD D4 send policy).</summary>
    public enum SendTokenDecision
    {
        /// <summary>Token matched the draft and its unchanged content - the send may proceed.</summary>
        Valid = 0,

        /// <summary>Token is unknown: never issued by this server process, already consumed, or replaced by a newer token.</summary>
        UnknownOrUsed = 1,

        /// <summary>Token existed but its time-to-live had elapsed.</summary>
        Expired = 2,

        /// <summary>Token was issued for a DIFFERENT draft than the one being sent.</summary>
        DraftMismatch = 3,

        /// <summary>The draft's content changed after the token was issued.</summary>
        ContentChanged = 4,
    }

    /// <summary>
    /// Thrown when the high-friction send policy refuses to send (D4): missing/invalid/
    /// expired token, changed draft, non-draft item, unverifiable identity. Nothing was
    /// sent when this is thrown. <see cref="Reason"/> is a machine-readable code that is
    /// also written to the audit log.
    /// </summary>
    public sealed class SendRefusedException : Exception
    {
        /// <summary>Creates the refusal.</summary>
        public SendRefusedException(string reason, string message)
            : base(message)
        {
            Reason = reason;
        }

        /// <summary>Machine-readable refusal code (e.g. "unknown_or_used_token", "draft_changed").</summary>
        public string Reason { get; }
    }

    /// <summary>
    /// In-memory one-time confirm-token store for the high-friction send flow (v3.MD
    /// D4/L5): <c>send(id)</c> without a token refuses and issues a token bound to that
    /// draft's EntryID + content hash; <c>send(id, confirm_token)</c> consumes it.
    /// Tokens are single-use (consumed by ANY attempt that references them, whatever the
    /// outcome), expire after <see cref="TimeToLive"/>, and at most one token is pending
    /// per draft (re-issuing replaces the previous one). Host-neutral and per-process:
    /// a server restart invalidates all tokens, exactly like hit ids.
    /// </summary>
    public sealed class SendConfirmationTokens
    {
        /// <summary>Default token time-to-live: short by design (D4 friction).</summary>
        public static readonly TimeSpan DefaultTimeToLive = TimeSpan.FromSeconds(120);

        private const int MaxPendingTokens = 32;
        private const string TokenPrefix = "confirm-";

        private readonly object _lock = new object();
        private readonly Dictionary<string, PendingToken> _tokens =
            new Dictionary<string, PendingToken>(StringComparer.Ordinal);
        private readonly Func<DateTime> _utcNow;

        /// <summary>Creates a store with the default 120 s time-to-live.</summary>
        public SendConfirmationTokens()
            : this(DefaultTimeToLive, null)
        {
        }

        /// <summary>
        /// Creates a store with an explicit time-to-live and optional test clock. The default
        /// is <see cref="MonotonicClock"/>, not <see cref="DateTime.UtcNow"/>: a token's whole
        /// purpose is that it expires, and on the wall clock a backwards jump would keep every
        /// pending send token valid for the size of the jump.
        /// </summary>
        public SendConfirmationTokens(TimeSpan timeToLive, Func<DateTime>? utcNowProvider = null)
        {
            if (timeToLive <= TimeSpan.Zero)
            {
                throw new ArgumentOutOfRangeException(nameof(timeToLive), "Token time-to-live must be positive.");
            }

            TimeToLive = timeToLive;
            _utcNow = utcNowProvider ?? (() => MonotonicClock.UtcNow);
        }

        /// <summary>How long an issued token stays valid.</summary>
        public TimeSpan TimeToLive { get; }

        /// <summary>Number of pending (unconsumed, possibly expired) tokens - test/diagnostic aid.</summary>
        public int PendingCount
        {
            get
            {
                lock (_lock)
                {
                    return _tokens.Count;
                }
            }
        }

        /// <summary>
        /// Issues a new one-time token bound to the draft's EntryID and current content
        /// hash. Any previously pending token for the same draft is invalidated.
        /// </summary>
        public string Issue(string draftEntryId, string contentHash)
        {
            if (string.IsNullOrWhiteSpace(draftEntryId))
            {
                throw new ArgumentException("Draft EntryID must not be blank.", nameof(draftEntryId));
            }

            if (string.IsNullOrWhiteSpace(contentHash))
            {
                throw new ArgumentException("Content hash must not be blank.", nameof(contentHash));
            }

            string token = TokenPrefix + NewRandomHex(16);
            DateTime now = _utcNow();
            lock (_lock)
            {
                PruneLocked(now);

                // One pending token per draft: a re-issue replaces the older token.
                List<string> replaced = _tokens
                    .Where(p => string.Equals(p.Value.DraftEntryId, draftEntryId, StringComparison.OrdinalIgnoreCase))
                    .Select(p => p.Key)
                    .ToList();
                foreach (string key in replaced)
                {
                    _tokens.Remove(key);
                }

                // Bound store: evict the oldest pending token when full.
                if (_tokens.Count >= MaxPendingTokens)
                {
                    string oldest = _tokens.OrderBy(p => p.Value.IssuedUtc).First().Key;
                    _tokens.Remove(oldest);
                }

                _tokens[token] = new PendingToken(draftEntryId, contentHash, now);
            }

            return token;
        }

        /// <summary>
        /// Consumes a token for a send attempt. SINGLE-USE IS STRICT: when the token is
        /// found it is removed from the store no matter the outcome - a mismatching or
        /// expired attempt burns it and a fresh token must be requested.
        /// </summary>
        public SendTokenDecision Consume(string token, string draftEntryId, string contentHash)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return SendTokenDecision.UnknownOrUsed;
            }

            lock (_lock)
            {
                if (!_tokens.TryGetValue(token, out PendingToken? pending))
                {
                    return SendTokenDecision.UnknownOrUsed;
                }

                _tokens.Remove(token);
                if (_utcNow() - pending.IssuedUtc > TimeToLive)
                {
                    return SendTokenDecision.Expired;
                }

                if (!string.Equals(pending.DraftEntryId, draftEntryId, StringComparison.OrdinalIgnoreCase))
                {
                    return SendTokenDecision.DraftMismatch;
                }

                if (!string.Equals(pending.ContentHash, contentHash, StringComparison.Ordinal))
                {
                    return SendTokenDecision.ContentChanged;
                }

                return SendTokenDecision.Valid;
            }
        }

        /// <summary>Removes a pending token (e.g. when its audit line could not be written).</summary>
        public void Invalidate(string token)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                return;
            }

            lock (_lock)
            {
                _tokens.Remove(token);
            }
        }

        private void PruneLocked(DateTime now)
        {
            List<string> expired = _tokens
                .Where(p => now - p.Value.IssuedUtc > TimeToLive)
                .Select(p => p.Key)
                .ToList();
            foreach (string key in expired)
            {
                _tokens.Remove(key);
            }
        }

        private static string NewRandomHex(int byteCount)
        {
            byte[] bytes = new byte[byteCount];
            using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(bytes);
            }

            char[] hex = new char[bytes.Length * 2];
            for (int i = 0; i < bytes.Length; i++)
            {
                hex[i * 2] = ToHexDigit(bytes[i] >> 4);
                hex[(i * 2) + 1] = ToHexDigit(bytes[i] & 0xF);
            }

            return new string(hex);
        }

        private static char ToHexDigit(int value)
        {
            return (char)(value < 10 ? '0' + value : 'a' + (value - 10));
        }

        private sealed class PendingToken
        {
            public PendingToken(string draftEntryId, string contentHash, DateTime issuedUtc)
            {
                DraftEntryId = draftEntryId;
                ContentHash = contentHash;
                IssuedUtc = issuedUtc;
            }

            public string DraftEntryId { get; }

            public string ContentHash { get; }

            public DateTime IssuedUtc { get; }
        }
    }
}
