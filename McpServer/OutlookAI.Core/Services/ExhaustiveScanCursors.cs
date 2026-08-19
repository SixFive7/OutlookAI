using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

using OutlookAI.Core.Com;

namespace OutlookAI.Core.Services
{
    /// <summary>Why a <c>resume_token</c> was not honoured. Five outcomes, five remedies.</summary>
    public enum ScanTokenDecision
    {
        /// <summary>The token names a live scan and the request matches it - carry on.</summary>
        Valid = 0,

        /// <summary>Not a token this server could ever have issued (wrong prefix, wrong length, not hex).</summary>
        Malformed = 1,

        /// <summary>Well formed, and this process has never issued it - or has restarted since.</summary>
        Unknown = 2,

        /// <summary>Issued by this process, and its time-to-live has elapsed.</summary>
        Expired = 3,

        /// <summary>
        /// A REAL token for a live scan, but a later page of the same scan has already been
        /// served, so this one names a position the chain has moved past.
        /// </summary>
        Superseded = 4,

        /// <summary>The scan is live and the request asks a different question than the one it continues.</summary>
        RequestChanged = 5,
    }

    /// <summary>One paged exhaustive scan, as the server parent remembers it between pages.</summary>
    public sealed class ExhaustiveScanSession
    {
        internal ExhaustiveScanSession(string sessionId, string fingerprint, DateTime issuedUtc)
        {
            SessionId = sessionId;
            Fingerprint = fingerprint;
            IssuedUtc = issuedUtc;
        }

        /// <summary>Identity of the scan, independent of which token currently addresses it.</summary>
        public string SessionId { get; }

        /// <summary>
        /// The canonical text of the request this chain answers. Stored as TEXT rather than
        /// as a hash on purpose: comparing is the same cost either way, and only the text can
        /// say WHICH argument changed when a resume is refused - which turns "start over"
        /// into "you moved the 'after' bound".
        /// </summary>
        public string Fingerprint { get; }

        /// <summary>When the last page was served (monotonic - only ever read as an age).</summary>
        public DateTime IssuedUtc { get; internal set; }

        /// <summary>The one token that currently addresses this scan; every earlier one is superseded.</summary>
        public string LiveToken { get; internal set; } = string.Empty;

        /// <summary>Where the walk stopped, handed straight back to the COM child.</summary>
        public ComScanCursor? Cursor { get; internal set; }

        /// <summary>Hits returned across every page of this chain so far.</summary>
        public int ItemsReturnedTotal { get; internal set; }

        /// <summary>Pages served so far, this one included.</summary>
        public int PagesServed { get; internal set; }

        /// <summary>Mail folders finished across the chain.</summary>
        public int FoldersDone { get; internal set; }

        /// <summary>Mail folders in scope, as the last page's enumeration counted them.</summary>
        public int FoldersTotal { get; internal set; }

        /// <summary>Store-relative path the next page starts in.</summary>
        public string? ResumeFolderPath { get; internal set; }

        /// <summary>The next page's inclusive date bound, when the folder sorted.</summary>
        public DateTime? ResumeCursorUtc { get; internal set; }

        /// <summary>Which rung the next page will use.</summary>
        public string? ResumeTier { get; internal set; }

        /// <summary>
        /// One sentence a caller can act on WITHOUT any token: the folder to aim at and the
        /// date to stop at, in the parameters <c>search</c> already has. It is what keeps a
        /// refusal from costing the work already done.
        /// </summary>
        public string DescribeRecovery()
        {
            StringBuilder text = new StringBuilder();
            _ = text.Append("That scan had finished ")
                .Append(FoldersDone.ToString(CultureInfo.InvariantCulture))
                .Append(" of ")
                .Append(FoldersTotal.ToString(CultureInfo.InvariantCulture))
                .Append(" folder(s) and returned ")
                .Append(ItemsReturnedTotal.ToString(CultureInfo.InvariantCulture))
                .Append(" item(s) over ")
                .Append(PagesServed.ToString(CultureInfo.InvariantCulture))
                .Append(" page(s).");

            if (!string.IsNullOrEmpty(ResumeFolderPath))
            {
                _ = text.Append(" Continue without a token by re-running the same search with folder:'")
                    .Append(ResumeFolderPath)
                    .Append('\'');
                if (ResumeCursorUtc.HasValue)
                {
                    _ = text.Append(" and before:'")
                        .Append(ResumeCursorUtc.Value.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture))
                        .Append('\'');
                }

                _ = text.Append('.');
            }

            return text.ToString();
        }
    }

    /// <summary>
    /// The continuation-token store for resumable exhaustive scans: it holds the walk state
    /// between pages, hands out the short opaque handles that address it, and refuses every
    /// way a handle can stop meaning what the caller thinks it means.
    /// <para>
    /// <b>Why the state is here and not on the wire.</b> Proving that no folder was skipped
    /// needs the SET of folders the chain has finished, and on a large store that set is far
    /// too big to put in a token an agent carries in its context - the maintainer's standing
    /// rule is that payload is context and context is the scarce resource. So the token is
    /// 37 characters and the state stays in this process.
    /// </para>
    /// <para>
    /// <b>Why not in the COM host child.</b> The event this design exists for is a scan that
    /// runs past its deadline, which ends with the supervisor killing that child. State kept
    /// there would be destroyed by exactly the failure it is meant to survive. In the parent
    /// it survives a host kill and a host restart, and the one thing it does NOT survive - an
    /// MCP-server restart - is covered without any state at all, because every payload
    /// carries the resume folder and date in plain fields a caller can pass back as
    /// <c>folder</c> and <c>before</c>.
    /// </para>
    /// <para>
    /// <b>Why per-process and unpersisted</b>, like <see cref="SendConfirmationTokens"/>,
    /// <see cref="ServerDraftRegistry"/> and <see cref="DraftUpdateIntents"/>: a fresh
    /// process never ran the earlier pages and cannot vouch for where they stopped. A caller
    /// whose token is gone gets a refusal that says so and names the way back, which is a
    /// smaller guarantee and never a wrong one.
    /// </para>
    /// </summary>
    public sealed class ExhaustiveScanCursors
    {
        /// <summary>
        /// How long a chain stays resumable after its last page.
        /// <para>
        /// Thirty minutes is roughly twice <c>ComOperationBudgets.ExhaustiveScanDeadlineMs</c>
        /// (615 s), so a caller who lets one FULL-BUDGET page run and then spends as long
        /// again deciding what to do with the answer still has a whole page of slack.
        /// <see cref="BodyCache"/>'s fifteen minutes is the right KIND of precedent - a paging
        /// session rather than send friction - but it is only about 1.5 page-lengths here, and
        /// a time-to-live shorter than one page is unusable by construction. T1 pins the
        /// relationship rather than the number.
        /// </para>
        /// </summary>
        public static readonly TimeSpan DefaultTimeToLive = TimeSpan.FromMinutes(30);

        /// <summary>
        /// Concurrent paged scans held; the least recently served is evicted beyond it.
        /// <para>
        /// Lower than <see cref="SendConfirmationTokens"/>'s 32 because a scan session is far
        /// heavier - it carries a finished-folder set and a per-folder duplicate-suppression
        /// set - and because paging four different exhaustive scans at once is already past
        /// any real use of a mode that costs minutes per page.
        /// </para>
        /// </summary>
        public const int Capacity = 4;

        private const string TokenPrefix = "scan-";
        private const int TokenHexLength = 32;

        private readonly object _lock = new object();
        private readonly Dictionary<string, ExhaustiveScanSession> _sessions =
            new Dictionary<string, ExhaustiveScanSession>(StringComparer.Ordinal);

        // Every token ever issued for a live session, not just the live one. A superseded
        // token has to be told apart from a token this server never issued, because the two
        // mean different things to the caller: one says "you already have a newer page", the
        // other says "start over".
        private readonly Dictionary<string, string> _tokenToSession =
            new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly Func<DateTime> _utcNow;

        /// <summary>Creates a store with the default time-to-live.</summary>
        public ExhaustiveScanCursors()
            : this(DefaultTimeToLive, null)
        {
        }

        /// <summary>
        /// Creates a store with an explicit time-to-live and optional test clock. The default
        /// is <see cref="MonotonicClock"/> rather than <see cref="DateTime.UtcNow"/> for the
        /// reason that class states: only the AGE of a session is ever meant here, and a
        /// backwards wall-clock jump would keep every dead chain resumable for the size of the
        /// jump.
        /// </summary>
        public ExhaustiveScanCursors(TimeSpan timeToLive, Func<DateTime>? utcNowProvider = null)
        {
            if (timeToLive <= TimeSpan.Zero)
            {
                throw new ArgumentOutOfRangeException(nameof(timeToLive), "Token time-to-live must be positive.");
            }

            TimeToLive = timeToLive;
            _utcNow = utcNowProvider ?? (() => MonotonicClock.UtcNow);
        }

        /// <summary>How long a chain stays resumable after its last page.</summary>
        public TimeSpan TimeToLive { get; }

        /// <summary>Live paged scans (test/diagnostic aid).</summary>
        public int SessionCount
        {
            get
            {
                lock (_lock)
                {
                    return _sessions.Count;
                }
            }
        }

        /// <summary>
        /// The canonical text of everything that decides WHICH MAIL a scan returns, in a fixed
        /// order and presence-first.
        /// <para>
        /// <c>top</c> and <c>snippet_chars</c> are deliberately absent: they shape the
        /// presentation of a page, not the question, so a caller may pull smaller pages
        /// part-way through a chain. Everything else is in, because a resume that quietly
        /// answered a different question under a continuity claim is the exact failure the
        /// token exists to prevent.
        /// </para>
        /// <para>
        /// Presence-first (a leading 0/1 per value) for the same reason
        /// <see cref="DraftUpdateIntents"/> uses it: an ABSENT argument and a SUPPLIED one
        /// mean different things, and no sentinel a caller could also have typed can be
        /// allowed to hash as the other.
        /// </para>
        /// </summary>
        public static string FingerprintOf(SearchRequest request, IReadOnlyList<string>? terms)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            StringBuilder canonical = new StringBuilder(256);
            AppendList(canonical, "terms", terms);
            Append(canonical, "searchIn", request.SearchIn.ToString());
            Append(canonical, "store", request.Store);
            Append(canonical, "folder", request.Folder);
            Append(canonical, "includeSubfolders", request.IncludeSubfolders ? "1" : "0");
            Append(canonical, "after", FormatInstant(request.AfterUtc));
            Append(canonical, "before", FormatInstant(request.BeforeUtc));
            Append(canonical, "from", request.From);
            Append(canonical, "unreadOnly", request.UnreadOnly?.ToString());
            Append(canonical, "hasAttachments", request.HasAttachments?.ToString());
            Append(canonical, "orderBySize", request.OrderBySizeDescending ? "1" : "0");
            return canonical.ToString();
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

        /// <summary>
        /// Resolves a caller's token. NOTHING IS CONSUMED: a page that fails after this point
        /// must leave the chain exactly as resumable as it found it, which is the opposite of
        /// the send path's single-use token and for the opposite reason - that one is friction
        /// by design, this one is the remedy for work that already cost minutes.
        /// </summary>
        public ScanTokenDecision Resolve(string? token, string fingerprint, out ExhaustiveScanSession? session)
        {
            session = null;
            if (!LooksLikeToken(token))
            {
                return ScanTokenDecision.Malformed;
            }

            DateTime now = _utcNow();
            lock (_lock)
            {
                // Deliberately NOT pruned before the lookup. Pruning first would delete the
                // very session this call is about and answer "unknown" - the message that
                // tells a caller the server restarted - for a chain that merely aged out,
                // whose message points them at exhaustive.position instead. Expiry is
                // therefore decided against the record and the record is dropped afterwards,
                // so a second attempt does correctly read as unknown.
                if (!_tokenToSession.TryGetValue(token!, out string? sessionId)
                    || !_sessions.TryGetValue(sessionId, out ExhaustiveScanSession? found))
                {
                    return ScanTokenDecision.Unknown;
                }

                if (now - found.IssuedUtc > TimeToLive)
                {
                    RemoveSessionLocked(sessionId);
                    return ScanTokenDecision.Expired;
                }

                session = found;
                if (!string.Equals(found.LiveToken, token, StringComparison.Ordinal))
                {
                    return ScanTokenDecision.Superseded;
                }

                return string.Equals(found.Fingerprint, fingerprint, StringComparison.Ordinal)
                    ? ScanTokenDecision.Valid
                    : ScanTokenDecision.RequestChanged;
            }
        }

        /// <summary>
        /// Opens a new chain for a scan that stopped early, or continues an existing one, and
        /// returns the token addressing its NEXT page. The previous token stops working the
        /// moment this returns, which is what keeps the finished-folder set coherent: two live
        /// tokens over one chain would let two callers advance the same state past each other.
        /// </summary>
        public string Issue(
            ExhaustiveScanSession? existing,
            string fingerprint,
            ComScanPosition position,
            int itemsReturnedThisPage,
            out ExhaustiveScanSession session)
        {
            if (position == null)
            {
                throw new ArgumentNullException(nameof(position));
            }

            if (string.IsNullOrEmpty(fingerprint))
            {
                throw new ArgumentException("Fingerprint must not be blank.", nameof(fingerprint));
            }

            DateTime now = _utcNow();
            string token = TokenPrefix + NewRandomHex(TokenHexLength / 2);
            lock (_lock)
            {
                PruneLocked(now);
                session = existing != null && _sessions.ContainsKey(existing.SessionId)
                    ? existing
                    : NewSessionLocked(fingerprint, now);

                session.IssuedUtc = now;
                session.LiveToken = token;
                session.Cursor = position.Cursor;
                session.ItemsReturnedTotal += itemsReturnedThisPage;
                session.PagesServed++;
                session.FoldersDone = position.FoldersDone;
                session.FoldersTotal = position.FoldersTotal;
                session.ResumeFolderPath = position.ResumeFolderPath;
                session.ResumeCursorUtc = position.ResumeCursorUtc;
                session.ResumeTier = position.ResumeTier;
                _tokenToSession[token] = session.SessionId;
                return token;
            }
        }

        /// <summary>
        /// Closes a chain that has covered its scope. Called on the page that completes, so a
        /// finished scan stops holding its finished-folder set and its emitted ids: the state
        /// exists to make the NEXT page possible, and there is no next page.
        /// </summary>
        public void Complete(ExhaustiveScanSession? session)
        {
            if (session == null)
            {
                return;
            }

            lock (_lock)
            {
                RemoveSessionLocked(session.SessionId);
            }
        }

        /// <summary>
        /// Records a page of a chain that stopped and could not be given a token, so the
        /// counters a later refusal reports stay true. Used when the walk stopped without a
        /// resumable position at all.
        /// </summary>
        public void NoteUnresumablePage(ExhaustiveScanSession? session, int itemsReturnedThisPage)
        {
            if (session == null)
            {
                return;
            }

            lock (_lock)
            {
                if (!_sessions.ContainsKey(session.SessionId))
                {
                    return;
                }

                session.ItemsReturnedTotal += itemsReturnedThisPage;
                session.PagesServed++;
                session.IssuedUtc = _utcNow();
            }
        }

        /// <summary>
        /// Whether a string could be a token this server issued. Shape only - it says nothing
        /// about whether the token is live - and it is what makes "you passed something that
        /// is not a token" a different answer from "that token has expired".
        /// </summary>
        public static bool LooksLikeToken(string? token)
        {
            if (token == null || token.Length != TokenPrefix.Length + TokenHexLength)
            {
                return false;
            }

            if (!token.StartsWith(TokenPrefix, StringComparison.Ordinal))
            {
                return false;
            }

            for (int i = TokenPrefix.Length; i < token.Length; i++)
            {
                char c = token[i];
                bool hex = (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f');
                if (!hex)
                {
                    return false;
                }
            }

            return true;
        }

        private ExhaustiveScanSession NewSessionLocked(string fingerprint, DateTime now)
        {
            if (_sessions.Count >= Capacity)
            {
                string? oldest = null;
                DateTime oldestAt = DateTime.MaxValue;
                foreach (KeyValuePair<string, ExhaustiveScanSession> pair in _sessions)
                {
                    if (pair.Value.IssuedUtc < oldestAt)
                    {
                        oldestAt = pair.Value.IssuedUtc;
                        oldest = pair.Key;
                    }
                }

                if (oldest != null)
                {
                    RemoveSessionLocked(oldest);
                }
            }

            ExhaustiveScanSession session = new ExhaustiveScanSession(NewRandomHex(8), fingerprint, now);
            _sessions[session.SessionId] = session;
            return session;
        }

        private void RemoveSessionLocked(string sessionId)
        {
            _ = _sessions.Remove(sessionId);
            List<string> orphans = new List<string>();
            foreach (KeyValuePair<string, string> pair in _tokenToSession)
            {
                if (string.Equals(pair.Value, sessionId, StringComparison.Ordinal))
                {
                    orphans.Add(pair.Key);
                }
            }

            for (int i = 0; i < orphans.Count; i++)
            {
                _ = _tokenToSession.Remove(orphans[i]);
            }
        }

        private void PruneLocked(DateTime now)
        {
            List<string> expired = new List<string>();
            foreach (KeyValuePair<string, ExhaustiveScanSession> pair in _sessions)
            {
                if (now - pair.Value.IssuedUtc > TimeToLive)
                {
                    expired.Add(pair.Key);
                }
            }

            for (int i = 0; i < expired.Count; i++)
            {
                // The SESSION goes; its tokens stay reachable only through it, so removing it
                // is what turns a later replay into "unknown" rather than a null dereference.
                RemoveSessionLocked(expired[i]);
            }
        }

        private static string FormatInstant(DateTime? value)
        {
            return value.HasValue
                ? value.Value.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ss.fffZ", CultureInfo.InvariantCulture)
                : null!;
        }

        private static string LabelOf(string canonicalLine)
        {
            int equals = canonicalLine.IndexOf('=');
            return equals > 0 ? canonicalLine.Substring(0, equals) : string.Empty;
        }

        private static void Append(StringBuilder canonical, string label, string? value)
        {
            _ = canonical.Append(label).Append('=');
            _ = value == null ? canonical.Append('0') : canonical.Append('1').Append('|').Append(value);
            _ = canonical.Append('\n');
        }

        private static void AppendList(StringBuilder canonical, string label, IReadOnlyList<string>? values)
        {
            if (values == null)
            {
                _ = canonical.Append(label).Append("=0").Append('\n');
                return;
            }

            _ = canonical.Append(label).Append("=1|")
                .Append(values.Count.ToString(CultureInfo.InvariantCulture));
            for (int i = 0; i < values.Count; i++)
            {
                _ = canonical.Append('|').Append(values[i] ?? string.Empty);
            }

            _ = canonical.Append('\n');
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
    }
}
