using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OutlookAI.Core.IndexSearch
{
    /// <summary>
    /// Builds Windows Search SQL for SystemIndex queries. Every emitted statement honors
    /// the v3.MD section-12 anti-pattern checklist:
    /// <list type="bullet">
    /// <item>Scoping only via <c>SCOPE='mapi...'</c> - never LIKE on System.ItemPathDisplay.</item>
    /// <item>Never selects System.Message.MessageId (0x80040E55 in combined queries),
    /// System.Search.Contents (query-only) or System.Search.EntryID (not the MAPI id).</item>
    /// <item>Rejects bare '*' terms (CONTAINS('*') = 0x80041605).</item>
    /// <item>Term predicates always NAME their column(s): a bare CONTAINS('term') searches
    /// System.Search.Contents alone (documented CONTAINS-predicate semantics), which
    /// carries no subject text - the SF-6 recall bug. Default shape is the
    /// Subject-OR-Contents pair; measured cost over the bare shape is ~0-2 ms on
    /// agent-sized (TOP 26 + ORDER BY) queries.</item>
    /// <item>Multi-term queries AND ACROSS the columns (one pair per term), never inside
    /// one column - the in-column shape missed mail with one term in the subject and
    /// another in the body (soak fix 13).</item>
    /// <item>Kind filter is 'email' or '(email OR document)' - never unfiltered.</item>
    /// <item>No aggregates, no JOINs (unsupported in WS-SQL).</item>
    /// <item>Sender/recipient filters use per-column CONTAINS - Phase-1 probes measured
    /// equality/LIKE on FromAddress at 1-10 s (property scan) vs ~60 ms for CONTAINS.</item>
    /// </list>
    /// </summary>
    public static class WsSqlBuilder
    {
        /// <summary>Columns selected for hit mapping. Keep in sync with <see cref="IndexRowMapper"/>.</summary>
        public static readonly IReadOnlyList<string> SelectColumns = new[]
        {
            "System.ItemUrl",
            "System.Subject",
            "System.Message.FromAddress",
            "System.Message.FromName",
            "System.Message.ToAddress",
            "System.Message.DateReceived",
            "System.ItemPathDisplay",
            "System.ItemNameDisplay",
            "System.Kind",
            "System.Search.AutoSummary",
            "System.Size",
            "System.IsRead",
            "System.Message.HasAttachments",
            "System.Message.ConversationID",
        };

        /// <summary>
        /// Columns that must never appear in a SELECT list (v3.MD sections 4/12): MessageId
        /// errors combined queries, Contents is query-only, Search.EntryID is a recyclable
        /// catalog-internal int32 unrelated to the MAPI EntryID.
        /// </summary>
        public static readonly IReadOnlyList<string> ForbiddenSelectColumns = new[]
        {
            "System.Message.MessageId",
            "System.Search.Contents",
            "System.Search.EntryID",
        };

        private const int MaxTop = 5000;
        private const int MaxTermLength = 128;

        /// <summary>Subject column of the term predicate (query-only for CONTAINS, selectable).</summary>
        private const string SubjectColumn = "System.Subject";

        /// <summary>Body/attachment-content stream. Query-only - never appears in a SELECT list.</summary>
        private const string ContentsColumn = "System.Search.Contents";

        /// <summary>Builds the search statement for <paramref name="query"/>.</summary>
        public static string Build(IndexQuery query)
        {
            if (query == null)
            {
                throw new ArgumentNullException(nameof(query));
            }

            if (query.Top < 1 || query.Top > MaxTop)
            {
                throw new ArgumentException(
                    string.Format(CultureInfo.InvariantCulture, "Top must be between 1 and {0}.", MaxTop),
                    nameof(query));
            }

            if (query.ReceivedOnOrAfterUtc.HasValue && query.ReceivedBeforeUtc.HasValue
                && query.ReceivedOnOrAfterUtc.Value >= query.ReceivedBeforeUtc.Value)
            {
                throw new ArgumentException("ReceivedOnOrAfterUtc must lie before ReceivedBeforeUtc.", nameof(query));
            }

            List<string> where = new List<string>();

            if (query.Scope != null)
            {
                where.Add("SCOPE='" + ValidateScope(query.Scope) + "'");
            }

            switch (query.Kinds)
            {
                case KindFilter.EmailOnly:
                    where.Add("System.Kind='email'");
                    break;
                case KindFilter.DocumentsOnly:
                    where.Add("System.Kind='document'");
                    break;
                case KindFilter.EmailAndDocuments:
                    where.Add("(System.Kind='email' OR System.Kind='document')");
                    break;
                default:
                    throw new ArgumentException("Unknown KindFilter value.", nameof(query));
            }

            if (query.Terms != null && query.Terms.Count > 0)
            {
                where.Add(BuildTermsPredicate(query.Terms, query.SearchIn));
            }

            if (query.FromAddressContains != null)
            {
                where.Add("CONTAINS(System.Message.FromAddress, '"
                    + QuotedContainsValue(ValidateTerm(query.FromAddressContains, "FromAddressContains")) + "')");
            }

            if (query.RecipientContains != null)
            {
                string recipient = QuotedContainsValue(ValidateTerm(query.RecipientContains, "RecipientContains"));
                where.Add("(CONTAINS(System.Message.ToAddress, '" + recipient
                    + "') OR CONTAINS(System.Message.CcAddress, '" + recipient + "'))");
            }

            if (query.ReceivedOnOrAfterUtc.HasValue)
            {
                where.Add("System.Message.DateReceived >= '" + FormatUtc(query.ReceivedOnOrAfterUtc.Value) + "'");
            }

            if (query.ReceivedBeforeUtc.HasValue)
            {
                where.Add("System.Message.DateReceived < '" + FormatUtc(query.ReceivedBeforeUtc.Value) + "'");
            }

            if (query.IsRead.HasValue)
            {
                where.Add("System.IsRead=" + (query.IsRead.Value ? "TRUE" : "FALSE"));
            }

            if (query.HasAttachments.HasValue)
            {
                where.Add("System.Message.HasAttachments=" + (query.HasAttachments.Value ? "TRUE" : "FALSE"));
            }

            if (query.ConversationIdEquals != null)
            {
                where.Add("System.Message.ConversationID='"
                    + ValidateConversationId(query.ConversationIdEquals).Replace("'", "''") + "'");
            }

            StringBuilder sql = new StringBuilder();
            sql.Append("SELECT TOP ").Append(query.Top.ToString(CultureInfo.InvariantCulture));
            sql.Append(' ').Append(string.Join(", ", SelectColumns));
            sql.Append(" FROM SystemIndex WHERE ");
            sql.Append(string.Join(" AND ", where));
            sql.Append(query.OrderBy == IndexOrder.SizeDescending
                ? " ORDER BY System.Size DESC"
                : " ORDER BY System.Message.DateReceived DESC");
            return sql.ToString();
        }

        private static string ValidateConversationId(string conversationId)
        {
            string trimmed = conversationId.Trim();
            if (trimmed.Length == 0 || trimmed.Length > 512)
            {
                throw new ArgumentException("ConversationIdEquals must be 1-512 characters.", nameof(conversationId));
            }

            foreach (char c in trimmed)
            {
                // Observed values are opaque id strings; reject anything that could
                // break out of the SQL literal beyond the escaped single quote.
                if (char.IsControl(c))
                {
                    throw new ArgumentException("ConversationIdEquals contains control characters.", nameof(conversationId));
                }
            }

            return trimmed;
        }

        /// <summary>
        /// Statement for the staleness probe: newest indexed DateReceived, optionally scoped.
        /// </summary>
        public static string BuildNewestReceivedProbe(string? scope)
        {
            string where = scope != null ? "SCOPE='" + ValidateScope(scope) + "' AND " : string.Empty;
            return "SELECT TOP 1 System.Message.DateReceived FROM SystemIndex WHERE " + where
                + "System.Kind='email' ORDER BY System.Message.DateReceived DESC";
        }

        /// <summary>
        /// Statement sampling email item URLs for store-scope discovery (URL prefixes are
        /// grouped client-side; WS-SQL has no DISTINCT/GROUP BY).
        /// </summary>
        public static string BuildStoreDiscoverySample(int top)
        {
            if (top < 1 || top > 100000)
            {
                throw new ArgumentException("Discovery sample size out of range.", nameof(top));
            }

            return "SELECT TOP " + top.ToString(CultureInfo.InvariantCulture)
                + " System.ItemUrl FROM SystemIndex WHERE System.Kind='email'";
        }

        /// <summary>Statement probing whether any item exists under a scope (TOP 1, URL only).</summary>
        public static string BuildScopeExistenceProbe(string scope)
        {
            return "SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE SCOPE='" + ValidateScope(scope) + "'";
        }

        /// <summary>
        /// Term predicate. Multi-term queries AND across the WHOLE matched text, not
        /// inside one column: each term gets its own Subject-OR-Contents pair and the
        /// pairs are ANDed, so mail carrying one term in the subject and another in the
        /// body matches (the pair-per-column shape
        /// <c>CONTAINS(Subject,'"a" AND "b"') OR CONTAINS(Contents,'"a" AND "b"')</c>
        /// silently missed exactly those - soak fix 13). Narrowed scopes stay
        /// single-column, where an in-column AND is equivalent and cheaper.
        /// Measured cost of the per-term pairs on this machine (warm best-of-3,
        /// agent-sized TOP 26 + ORDER BY): +0-2 ms over the old shape at 1-3 terms.
        /// </summary>
        private static string BuildTermsPredicate(IReadOnlyList<string> terms, SearchIn searchIn)
        {
            List<string> quoted = new List<string>(terms.Count);
            for (int i = 0; i < terms.Count; i++)
            {
                quoted.Add(QuotedContainsValue(ValidateTerm(terms[i], "Terms[" + i.ToString(CultureInfo.InvariantCulture) + "]")));
            }

            switch (searchIn)
            {
                case SearchIn.SubjectAndBody:
                    List<string> pairs = new List<string>(quoted.Count);
                    foreach (string term in quoted)
                    {
                        pairs.Add("(CONTAINS(" + SubjectColumn + ", '" + term + "') OR CONTAINS("
                            + ContentsColumn + ", '" + term + "'))");
                    }

                    return string.Join(" AND ", pairs);
                case SearchIn.SubjectOnly:
                    return "CONTAINS(" + SubjectColumn + ", '" + string.Join(" AND ", quoted) + "')";
                case SearchIn.BodyOnly:
                    return "CONTAINS(" + ContentsColumn + ", '" + string.Join(" AND ", quoted) + "')";
                default:
                    throw new ArgumentException("Unknown SearchIn value.", nameof(searchIn));
            }
        }

        private static string ValidateScope(string scope)
        {
            string trimmed = scope.Trim();
            if (trimmed.Length == 0)
            {
                throw new ArgumentException("Scope must not be empty.", nameof(scope));
            }

            int schemeEnd = trimmed.IndexOf("://", StringComparison.Ordinal);
            if (schemeEnd <= 0 || !trimmed.Substring(0, schemeEnd).StartsWith("mapi", StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException(
                    "Scope must be a mapi URL (mapi16://...). Path filtering outside SCOPE is a property scan (v3.MD section 12).",
                    nameof(scope));
            }

            if (trimmed.IndexOf('\'') >= 0)
            {
                throw new ArgumentException("Scope must not contain single quotes.", nameof(scope));
            }

            return trimmed;
        }

        /// <summary>
        /// Validates a free-text term. Rules: non-blank after trimming; no '"' (phrase
        /// quoting is added by the builder); '*' only as the final character (prefix
        /// match) and never alone - CONTAINS('*') is invalid (0x80041605); remaining
        /// characters restricted to letters/digits plus a small punctuation set.
        /// </summary>
        private static string ValidateTerm(string term, string parameterName)
        {
            if (term == null)
            {
                throw new ArgumentException("Term must not be null.", parameterName);
            }

            string trimmed = term.Trim();
            if (trimmed.Length == 0)
            {
                throw new ArgumentException("Term must not be blank.", parameterName);
            }

            if (trimmed.Length > MaxTermLength)
            {
                throw new ArgumentException("Term too long.", parameterName);
            }

            if (trimmed == "*")
            {
                throw new ArgumentException("Bare '*' is not a valid term (CONTAINS('*') fails with 0x80041605).", parameterName);
            }

            for (int i = 0; i < trimmed.Length; i++)
            {
                char c = trimmed[i];
                if (c == '*')
                {
                    if (i != trimmed.Length - 1)
                    {
                        throw new ArgumentException("'*' is only allowed as the final character (prefix match).", parameterName);
                    }

                    continue;
                }

                bool allowed = char.IsLetterOrDigit(c)
                    || c == ' ' || c == '@' || c == '.' || c == '_' || c == '-' || c == '\'' || c == '+';
                if (!allowed)
                {
                    throw new ArgumentException(
                        string.Format(CultureInfo.InvariantCulture, "Term contains unsupported character '{0}'.", c),
                        parameterName);
                }
            }

            return trimmed;
        }

        /// <summary>Wraps a validated term in CONTAINS phrase quotes, escaping embedded single quotes for SQL.</summary>
        private static string QuotedContainsValue(string validatedTerm)
        {
            return "\"" + validatedTerm.Replace("'", "''") + "\"";
        }

        private static string FormatUtc(DateTime value)
        {
            // Unspecified kinds are taken as already-UTC (the property names carry the
            // Utc suffix); only Local values are converted.
            DateTime utc = value.Kind == DateTimeKind.Local ? value.ToUniversalTime() : value;
            return utc.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }
    }
}
