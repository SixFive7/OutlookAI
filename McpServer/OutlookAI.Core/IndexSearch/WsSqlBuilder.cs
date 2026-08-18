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
    /// <item>Scoping only via <c>SCOPE='mapi...'</c> - never LIKE on System.ItemPathDisplay.
    /// Non-recursive narrowing adds <c>System.ItemFolderPathDisplay=</c> equality
    /// (index-backed) - never the shallow <c>DIRECTORY=</c> predicate, which silently
    /// drops every attachment-content row.</item>
    /// <item>String literals escape a single quote by doubling it, in EVERY position -
    /// a folder named <c>O'Brien</c> must be searchable, not an exception.</item>
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
    /// <item>Kind: 'email' narrows to messages; the attachment-bearing shapes emit NO
    /// Kind predicate under a mapi SCOPE (the namespace is the guard) and an enumerated
    /// kind list without one, because an attachment row carries the ATTACHMENT's kind -
    /// 'document' alone dropped 22.6% of them. Admission is decided by IndexRowFilter
    /// after the rows come back (v3.MD section 0.8 block (q)).</item>
    /// <item>No aggregates, no JOINs (unsupported in WS-SQL).</item>
    /// <item>Sender/recipient filters use per-column CONTAINS - Phase-1 probes measured
    /// equality/LIKE on FromAddress at 1-10 s (property scan) vs ~60 ms for CONTAINS. The
    /// sender filter names BOTH sender columns (address OR display name), because the other
    /// two search tiers always did and the address alone silently under-returned.</item>
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

        /// <summary>Hard ceiling on the emitted <c>SELECT TOP</c> (also bounds the post-filter over-fetch).</summary>
        public const int MaxTop = 5000;

        /// <summary>
        /// Longest accepted search term, in characters. Public because it is the definition
        /// of a valid TERM for the whole product, not a detail of the index statement: the
        /// COM tier's <c>ExhaustiveDaslFilter</c> derives its own limit from this one.
        /// <para>
        /// They were two private 128s. Raising either alone would have made indexed and
        /// exhaustive search disagree about which terms are valid - a term accepted by one
        /// mode and rejected by the other, for the same query, with nothing to notice it.
        /// </para>
        /// </summary>
        public const int MaxTermLength = 128;

        /// <summary>
        /// Hard ceiling on <see cref="IndexQuery.FolderPathsAnyOf"/> literals in one
        /// statement. MEASURED on this machine (read-only ADODB battery, 2026-07-27): a
        /// folder-path OR-set of 95 literals executes, 100 fails outright with
        /// "Catastrophic failure" (0x8000FFFF) - the provider has a hard restriction-count
        /// limit, so an uncapped OR-set is a crash, not a slowdown. Cost is linear and
        /// modest well below it (delegate store root + term, warm best-of-3: bare SCOPE
        /// 43 ms, x10 53 ms, x20 59 ms, x40 71 ms, x80 101 ms), which is why callers cap
        /// at a much lower value and WIDEN to the plain scope instead of truncating a set.
        /// This constant is the builder's last-resort guard: exceeding it throws rather
        /// than emitting a statement that would fail at execution time.
        /// </summary>
        public const int MaxFolderPaths = 64;

        /// <summary>The non-recursive folder column (see <see cref="IndexQuery.FolderPathsAnyOf"/>).</summary>
        private const string FolderPathColumn = "System.ItemFolderPathDisplay";

        /// <summary>Subject column of the term predicate (query-only for CONTAINS, selectable).</summary>
        private const string SubjectColumn = "System.Subject";

        /// <summary>Body/attachment-content stream. Query-only - never appears in a SELECT list.</summary>
        private const string ContentsColumn = "System.Search.Contents";

        /// <summary>Sender SMTP address column of the <c>from</c> predicate.</summary>
        private const string FromAddressColumn = "System.Message.FromAddress";

        /// <summary>Sender DISPLAY NAME column of the <c>from</c> predicate (see <see cref="IndexQuery.SenderContains"/>).</summary>
        private const string FromNameColumn = "System.Message.FromName";

        /// <summary>Builds the search statement for <paramref name="query"/>.</summary>
        public static string Build(IndexQuery query)
        {
            return Build(query, null);
        }

        /// <summary>
        /// Builds the search statement, optionally emitting a different <c>TOP</c> than
        /// <see cref="IndexQuery.Top"/>. The override exists for the post-filter over-fetch
        /// (<see cref="IndexRowFilter.ComputeSqlTop"/>): the caller still wants
        /// <c>query.Top</c> ADMITTED rows, but the statement must offer more candidates.
        /// </summary>
        public static string Build(IndexQuery query, int? topOverride)
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

            int effectiveTop = topOverride ?? query.Top;
            if (effectiveTop < 1 || effectiveTop > MaxTop)
            {
                throw new ArgumentException(
                    string.Format(CultureInfo.InvariantCulture, "Top must be between 1 and {0}.", MaxTop),
                    nameof(topOverride));
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

            if (query.FolderPathsAnyOf != null && query.FolderPathsAnyOf.Count > 0)
            {
                where.Add(BuildFolderPathPredicate(query.FolderPathsAnyOf));
            }

            switch (query.Kinds)
            {
                case KindFilter.EmailOnly:
                    // Messages only: 'email' is exactly the message-level kind.
                    where.Add("System.Kind='email'");
                    break;
                case KindFilter.DocumentsOnly:
                case KindFilter.EmailAndDocuments:
                case KindFilter.MessagesAnyClass:
                    // Shapes that must not be narrowed by kind. An attachment-content row
                    // carries the ATTACHMENT's kind (picture / communication / calendar /
                    // music / video, not just document) and a message-level row carries its
                    // ITEM CLASS's kind (a meeting request is 'calendar'), so no kind list
                    // can be both complete and future-proof - IndexRowFilter decides
                    // admission on the URL instead. Under a mapi SCOPE the namespace already
                    // fences the statement, so no Kind predicate is emitted at all; without
                    // one the enumerated kinds keep the provider from offering the whole
                    // file system.
                    if (query.Scope == null)
                    {
                        where.Add(BuildUnscopedKindPredicate());
                    }

                    break;
                default:
                    throw new ArgumentException("Unknown KindFilter value.", nameof(query));
            }

            if (query.Terms != null && query.Terms.Count > 0)
            {
                where.Add(BuildTermsPredicate(query.Terms, query.SearchIn));
            }

            if (query.SenderContains != null)
            {
                // ADDRESS *or* NAME, because that is what the tool promises and what the
                // other two tiers do. Matching the address alone made a name-fragment
                // filter return zero index rows and then report the sweep's few minutes of
                // mail as the whole answer (gap B1). See IndexQuery.SenderContains for the
                // measured recall gap; the added CONTAINS is index-backed and cheap
                // (measured on this index: FromName alone 18 ms against FromAddress alone
                // 42 ms for the same fragment; the OR pair costs +0-12 ms on agent-sized
                // TOP 26 statements and is faster than address-only on some).
                string sender = QuotedContainsValue(ValidateTerm(query.SenderContains, "SenderContains"));
                where.Add("(CONTAINS(" + FromAddressColumn + ", '" + sender
                    + "') OR CONTAINS(" + FromNameColumn + ", '" + sender + "'))");
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

            if (where.Count == 0)
            {
                // Cannot happen through the shipped paths (an unscoped query always keeps
                // the enumerated kind predicate), but an empty WHERE would be invalid SQL.
                throw new ArgumentException("A search statement needs at least one predicate.", nameof(query));
            }

            StringBuilder sql = new StringBuilder();
            sql.Append("SELECT TOP ").Append(effectiveTop.ToString(CultureInfo.InvariantCulture));
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
        /// Statement probing whether the index holds ANY row for a folder scope - the
        /// non-silent zero-row guard (v3.MD probe block, constraint C7). Deliberately
        /// carries no term/date/sender predicates: it answers "does this folder scope
        /// resolve at all", never "does the search match".
        /// </summary>
        public static string BuildFolderScopeExistenceProbe(string? scope, IReadOnlyList<string>? folderPaths)
        {
            List<string> where = new List<string>();
            if (scope != null)
            {
                where.Add("SCOPE='" + ValidateScope(scope) + "'");
            }

            if (folderPaths != null && folderPaths.Count > 0)
            {
                where.Add(BuildFolderPathPredicate(folderPaths));
            }

            if (where.Count == 0)
            {
                throw new ArgumentException("A folder-scope probe needs a scope or a folder path.", nameof(scope));
            }

            // No kind predicate: the question is "does this folder scope resolve", and a
            // folder holding only meeting requests or only attachment rows resolves just as
            // well as one holding mail. Filtering here would re-create the silent zero the
            // guard exists to prevent.
            return "SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE " + string.Join(" AND ", where);
        }

        /// <summary>
        /// Kind predicate for a statement with no SCOPE. Enumerates the kinds actually seen
        /// on message-level and attachment-content rows (v3.MD block (q)); it keeps the
        /// provider selective, while <see cref="IndexRowFilter"/> supplies correctness.
        /// </summary>
        private static string BuildUnscopedKindPredicate()
        {
            List<string> equalities = new List<string>(IndexRowFilter.UnscopedKinds.Count);
            foreach (string kind in IndexRowFilter.UnscopedKinds)
            {
                equalities.Add("System.Kind='" + kind + "'");
            }

            return "(" + string.Join(" OR ", equalities) + ")";
        }

        /// <summary>
        /// Non-recursive folder predicate: <c>System.ItemFolderPathDisplay</c> equality,
        /// ORed over the requested paths. Under <c>=</c> nothing but the single quote
        /// needs escaping - <c>%</c>, <c>_</c>, <c>[</c>, <c>]</c>, <c>{</c>, <c>}</c> and
        /// <c>"</c> are literal (proved live: '/store/Inbo_' and '/store/Inbox%' both
        /// return 0 against a real Inbox, so '=' is not LIKE), and spaces must stay
        /// literal (a %20-encoded space returns 0 - the MAPI handler already encoded its
        /// URLs at index time).
        /// </summary>
        private static string BuildFolderPathPredicate(IReadOnlyList<string> folderPaths)
        {
            if (folderPaths.Count > MaxFolderPaths)
            {
                throw new ArgumentException(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "At most {0} folder paths may be ORed in one statement (the provider fails outright near 100).",
                        MaxFolderPaths),
                    nameof(folderPaths));
            }

            List<string> equalities = new List<string>(folderPaths.Count);
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string path in folderPaths)
            {
                string validated = ValidateFolderPath(path);
                if (!seen.Add(validated))
                {
                    continue; // Duplicate leaf names collapse to one equality.
                }

                equalities.Add(FolderPathColumn + "='" + EscapeSqlLiteral(validated) + "'");
            }

            if (equalities.Count == 0)
            {
                throw new ArgumentException("Folder paths must contain at least one usable value.", nameof(folderPaths));
            }

            return equalities.Count == 1 ? equalities[0] : "(" + string.Join(" OR ", equalities) + ")";
        }

        private static string ValidateFolderPath(string folderPath)
        {
            if (folderPath == null)
            {
                throw new ArgumentException("Folder path must not be null.", nameof(folderPath));
            }

            string trimmed = folderPath.Trim();
            if (trimmed.Length == 0)
            {
                throw new ArgumentException("Folder path must not be blank.", nameof(folderPath));
            }

            if (trimmed[0] != '/')
            {
                throw new ArgumentException(
                    "Folder path must be a System.ItemFolderPathDisplay value starting with '/' (e.g. '/account/Inbox').",
                    nameof(folderPath));
            }

            if (trimmed.Length > 512)
            {
                // The property-store limit; real paths here top out around 47 characters.
                throw new ArgumentException("Folder path is too long.", nameof(folderPath));
            }

            foreach (char c in trimmed)
            {
                if (char.IsControl(c))
                {
                    throw new ArgumentException("Folder path contains control characters.", nameof(folderPath));
                }
            }

            return trimmed;
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

        /// <summary>
        /// Validates a SCOPE value and returns it as a READY-TO-EMBED SQL literal body
        /// (single quotes doubled).
        /// <para>
        /// This used to THROW on any scope containing <c>'</c>, which made a folder named
        /// <c>O'Brien</c> un-searchable by hard exception. Measured (2026-07-27): a raw
        /// <c>'</c> yields 0x80040E14 (syntax error) in both the SCOPE and the folder-path
        /// literal, while <c>''</c> parses in both - so doubling is the correct and only
        /// required escape. Nothing else needs escaping: <c>%</c>, <c>_</c>, <c>[</c>,
        /// <c>]</c>, <c>{</c>, <c>}</c> and <c>"</c> are literal under <c>=</c>, and
        /// spaces must stay literal (<c>%20</c> matches nothing).
        /// </para>
        /// </summary>
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

            foreach (char c in trimmed)
            {
                if (char.IsControl(c))
                {
                    throw new ArgumentException("Scope contains control characters.", nameof(scope));
                }
            }

            return EscapeSqlLiteral(trimmed);
        }

        /// <summary>Doubles single quotes - the only escape a WS-SQL string literal needs.</summary>
        private static string EscapeSqlLiteral(string value)
        {
            return value.IndexOf('\'') < 0 ? value : value.Replace("'", "''");
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
