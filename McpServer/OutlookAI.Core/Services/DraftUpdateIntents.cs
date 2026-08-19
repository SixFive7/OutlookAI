using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

using OutlookAI.Core.Com;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// One <c>update_draft</c> whose outcome nobody can state, plus the pre-image a repeat
    /// of the same request needs to finish it.
    /// </summary>
    public sealed class DraftUpdateIntent
    {
        /// <summary>Creates the record.</summary>
        public DraftUpdateIntent(string key, string entryId, ComDraftUpdateResume preImage, DateTime issuedUtc)
        {
            Key = key;
            EntryId = entryId;
            PreImage = preImage;
            IssuedUtc = issuedUtc;
        }

        /// <summary>Idempotence key: this draft plus this exact request.</summary>
        public string Key { get; }

        /// <summary>The draft the request was aimed at.</summary>
        public string EntryId { get; }

        /// <summary>The draft's state before the first attempt touched it.</summary>
        public ComDraftUpdateResume PreImage { get; }

        /// <summary>When the intent was recorded (monotonic clock - it is only ever used as an age).</summary>
        public DateTime IssuedUtc { get; }
    }

    /// <summary>
    /// The record that makes <c>update_draft</c> re-entrant: what a call INTENDED to do,
    /// written down before it was attempted, so a repeat of the same request completes the
    /// remaining steps instead of performing the whole thing a second time.
    /// <para>
    /// <b>Why it lives in this process and not on the draft.</b> The failure being
    /// defended against is the operation deadline expiring and the supervisor killing the
    /// COM host child. The child dies; THIS process does not. A marker written on the item
    /// itself would travel with the draft and be visible to a second server, but it can
    /// only be written by the process that dies - so it is exactly as unreliable as the
    /// thing it records - it needs a second mutation to clear, which is one more window
    /// for the same kill, and it puts server bookkeeping on the user's own mail, which is
    /// a line this product does not cross (the same instinct that makes discard a soft
    /// delete and signatures untouchable).
    /// </para>
    /// <para>
    /// <b>Why it is not persisted.</b> A server restart empties it, exactly like
    /// <see cref="ServerDraftRegistry"/> and <see cref="SendConfirmationTokens"/>. A fresh
    /// process did not observe the pre-image and cannot vouch for it; resuming from a
    /// record it inherited would be asserting the state of a draft it never read. When the
    /// record is gone the caller simply gets the pre-existing answer - the outcome is
    /// unknown, check the draft - which is a smaller guarantee, never a wrong one.
    /// </para>
    /// <para>
    /// <b>What counts as a repeat.</b> Only a request that is byte-for-byte the same
    /// request, against the same draft, while the earlier attempt's outcome is still
    /// unknown. Two identical calls are NOT automatically a retry: a call that ANSWERED
    /// (updated, or refused) settles its intent, so a second identical call after it runs
    /// as a fresh update. And any other update to the same draft drops the pending record,
    /// because the pre-image then describes a draft that no longer exists.
    /// </para>
    /// </summary>
    public sealed class DraftUpdateIntents
    {
        /// <summary>
        /// How long an unresolved intent may still be resumed.
        /// <para>
        /// It bounds a genuine hazard rather than tidiness: the pre-image describes the
        /// draft as it was, and the longer it sits the more likely the user has edited that
        /// draft in Outlook themselves - at which point completing the interrupted update
        /// would be reasoning from a state that no longer holds. Ten minutes is well past
        /// any operation deadline (the longest is 615 s) and well short of a working
        /// session, so it covers "the agent retried the call that just failed" and nothing
        /// else.
        /// </para>
        /// </summary>
        public static readonly TimeSpan DefaultTimeToLive = TimeSpan.FromMinutes(10);

        /// <summary>Maximum unresolved intents held; the oldest is evicted beyond it.</summary>
        public const int Capacity = 32;

        private readonly object _lock = new object();
        private readonly Dictionary<string, DraftUpdateIntent> _pending = new Dictionary<string, DraftUpdateIntent>(StringComparer.Ordinal);
        private readonly Func<DateTime> _utcNow;

        /// <summary>Creates a store with the default time-to-live.</summary>
        public DraftUpdateIntents()
            : this(DefaultTimeToLive, null)
        {
        }

        /// <summary>
        /// Creates a store with an explicit time-to-live and optional test clock. The
        /// default is <see cref="MonotonicClock"/> rather than <see cref="DateTime.UtcNow"/>
        /// for the reason that class states: only the AGE of a record is ever meant here,
        /// and a backwards wall-clock jump would keep every stale intent resumable.
        /// </summary>
        public DraftUpdateIntents(TimeSpan timeToLive, Func<DateTime>? utcNowProvider = null)
        {
            if (timeToLive <= TimeSpan.Zero)
            {
                throw new ArgumentOutOfRangeException(nameof(timeToLive), "Intent time-to-live must be positive.");
            }

            TimeToLive = timeToLive;
            _utcNow = utcNowProvider ?? (() => MonotonicClock.UtcNow);
        }

        /// <summary>How long an unresolved intent stays resumable.</summary>
        public TimeSpan TimeToLive { get; }

        /// <summary>Number of unresolved intents (test/diagnostic aid).</summary>
        public int PendingCount
        {
            get
            {
                lock (_lock)
                {
                    return _pending.Count;
                }
            }
        }

        /// <summary>
        /// Derives the idempotence key for one update request. THE SERVER DERIVES IT; the
        /// caller supplies nothing.
        /// <para>
        /// The send path's confirm token is caller-supplied because its whole purpose is
        /// friction - a human has to say yes. Re-entrancy wants the opposite: it has to
        /// work when the caller does nothing special, because the caller that most needs it
        /// is an agent re-issuing the call that just failed. A caller-supplied key would
        /// also let two DIFFERENT requests claim the same identity, which is a worse
        /// failure than the one it prevents.
        /// </para>
        /// <para>
        /// Everything that reaches Outlook is folded in, in a fixed order, so any
        /// difference at all - one changed character of body, one more recipient, a
        /// different file - is a different request rather than a resumption of this one.
        /// </para>
        /// </summary>
        public static string KeyFor(
            string entryId,
            ComDraftBody? body,
            string? subject,
            IReadOnlyList<string>? to,
            IReadOnlyList<string>? cc,
            IReadOnlyList<string>? bcc,
            int? importance,
            bool? requestReadReceipt,
            ComSignatureOverride? signature,
            IReadOnlyList<string> attachmentPaths,
            IReadOnlyList<string> removeNames,
            bool display)
        {
            if (string.IsNullOrWhiteSpace(entryId))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryId));
            }

            StringBuilder canonical = new StringBuilder(256);
            canonical.Append("entry=").Append(entryId.Trim()).Append('\n');
            canonical.Append("bodyKind=").Append(body == null ? "none" : body.IsHtml ? "html" : "text").Append('\n');
            Append(canonical, "body", body == null ? null : body.IsHtml ? body.Html : body.Text);
            Append(canonical, "subject", subject);
            AppendList(canonical, "to", to);
            AppendList(canonical, "cc", cc);
            AppendList(canonical, "bcc", bcc);
            Append(canonical, "importance", importance?.ToString(CultureInfo.InvariantCulture));
            Append(canonical, "readReceipt", requestReadReceipt?.ToString());
            Append(canonical, "signature", signature?.Name);
            Append(canonical, "signaturePath", signature?.FilePath);

            // Attachment paths and removal names keep their REQUEST ORDER: the plan a repeat
            // computes consumes them in that order, so two requests differing only in ordering
            // are genuinely different requests here.
            AppendList(canonical, "attach", attachmentPaths);
            AppendList(canonical, "remove", removeNames);
            canonical.Append("display=").Append(display ? "1" : "0").Append('\n');

            byte[] hash;
            using (SHA256 sha = SHA256.Create())
            {
                hash = sha.ComputeHash(Encoding.UTF8.GetBytes(canonical.ToString()));
            }

            StringBuilder hex = new StringBuilder(hash.Length * 2);
            foreach (byte b in hash)
            {
                _ = hex.Append(b.ToString("x2", CultureInfo.InvariantCulture));
            }

            return hex.ToString();
        }

        /// <summary>
        /// Records an intent BEFORE its first attempt runs, dropping any other unresolved
        /// intent for the same draft - once a different update touches the draft, an older
        /// pre-image describes a state that is gone.
        /// </summary>
        public void Begin(string key, string entryId, ComDraftUpdateResume preImage)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                throw new ArgumentException("Intent key must not be blank.", nameof(key));
            }

            if (string.IsNullOrWhiteSpace(entryId))
            {
                throw new ArgumentException("EntryID must not be blank.", nameof(entryId));
            }

            if (preImage == null)
            {
                throw new ArgumentNullException(nameof(preImage));
            }

            DateTime now = _utcNow();
            lock (_lock)
            {
                PruneLocked(now);
                foreach (string stale in _pending
                    .Where(p => string.Equals(p.Value.EntryId, entryId, StringComparison.OrdinalIgnoreCase)
                        && !string.Equals(p.Key, key, StringComparison.Ordinal))
                    .Select(p => p.Key)
                    .ToList())
                {
                    _ = _pending.Remove(stale);
                }

                // An existing record for the SAME key keeps its original pre-image: it
                // describes the draft before anything was applied, and re-reading it now
                // would capture the interrupted attempt's own damage.
                if (_pending.ContainsKey(key))
                {
                    return;
                }

                if (_pending.Count >= Capacity)
                {
                    string oldest = _pending.OrderBy(p => p.Value.IssuedUtc).First().Key;
                    _ = _pending.Remove(oldest);
                }

                _pending[key] = new DraftUpdateIntent(key, entryId, preImage, now);
            }
        }

        /// <summary>
        /// The pre-image to resume from, or null when this is not a repeat of an
        /// unresolved attempt. Does NOT consume the record: an interrupted retry has to
        /// stay resumable, and only an answer settles it.
        /// </summary>
        public ComDraftUpdateResume? Resume(string key, string entryId)
        {
            if (string.IsNullOrWhiteSpace(key) || string.IsNullOrWhiteSpace(entryId))
            {
                return null;
            }

            DateTime now = _utcNow();
            lock (_lock)
            {
                PruneLocked(now);
                if (!_pending.TryGetValue(key, out DraftUpdateIntent? intent))
                {
                    return null;
                }

                return string.Equals(intent.EntryId, entryId, StringComparison.OrdinalIgnoreCase)
                    ? intent.PreImage
                    : null;
            }
        }

        /// <summary>Settles an intent: the call answered, so nothing is left outstanding.</summary>
        public void Settle(string key)
        {
            if (string.IsNullOrWhiteSpace(key))
            {
                return;
            }

            lock (_lock)
            {
                _ = _pending.Remove(key);
            }
        }

        /// <summary>
        /// Drops every unresolved intent for a draft. Called when the draft is discarded or
        /// its id is re-keyed: a pre-image outlives neither.
        /// </summary>
        public void Forget(string? entryId)
        {
            if (string.IsNullOrWhiteSpace(entryId))
            {
                return;
            }

            lock (_lock)
            {
                foreach (string key in _pending
                    .Where(p => string.Equals(p.Value.EntryId, entryId, StringComparison.OrdinalIgnoreCase))
                    .Select(p => p.Key)
                    .ToList())
                {
                    _ = _pending.Remove(key);
                }
            }
        }

        private void PruneLocked(DateTime now)
        {
            foreach (string key in _pending
                .Where(p => now - p.Value.IssuedUtc > TimeToLive)
                .Select(p => p.Key)
                .ToList())
            {
                _ = _pending.Remove(key);
            }
        }

        /// <summary>
        /// Appends one optional value, PRESENCE FIRST. An absent argument and a supplied one
        /// mean different things to update_draft - "leave this alone" against "set it to
        /// this" - so a sentinel value the caller could also have typed would let one hash as
        /// the other. A leading 0/1 makes that impossible for any value at all.
        /// </summary>
        private static void Append(StringBuilder canonical, string label, string? value)
        {
            _ = canonical.Append(label).Append('=');
            _ = value == null ? canonical.Append('0') : canonical.Append('1').Append('|').Append(value);
            _ = canonical.Append('\n');
        }

        /// <summary>Appends an optional list, on the same presence-first rule and for the same reason.</summary>
        private static void AppendList(StringBuilder canonical, string label, IReadOnlyList<string>? values)
        {
            if (values == null)
            {
                _ = canonical.Append(label).Append("=0").Append('\n');
                return;
            }

            _ = canonical.Append(label).Append("=1|")
                .Append(values.Count.ToString(CultureInfo.InvariantCulture)).Append('\n');
            foreach (string value in values)
            {
                _ = canonical.Append(label).Append('=').Append(value ?? string.Empty).Append('\n');
            }
        }
    }
}
