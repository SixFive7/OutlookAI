using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

using OutlookAI.Core.IndexSearch;

namespace OutlookAI.Core.Com
{
    /// <summary>Term-matching engine for the exhaustive COM scan (v3.MD D19/section 12).</summary>
    public enum ExhaustiveEngine
    {
        /// <summary>
        /// Index-backed <c>ci_phrasematch</c> in the DASL restriction - whole-word
        /// semantics, fast even on big folders. Supported in Restrict/GetTable (NOT
        /// Find/FindRow); gated on <c>Store.IsInstantSearchEnabled</c> and
        /// feature-detected at runtime (v3.MD section 12).
        /// </summary>
        CiPhraseMatch = 0,

        /// <summary>
        /// Plain <c>like '%term%'</c> - substring semantics, provider-side scan (slow on
        /// big folders). The last-resort fallback.
        /// </summary>
        Like = 1,
    }

    /// <summary>
    /// Builds DASL <c>@SQL=</c> restrictions for the exhaustive folder/date-bounded COM
    /// scan (v3.MD section 0.6 Phase 3). Pure logic, unit-tested in T1. Shapes follow the
    /// section-12 rules: ci_* only in Restrict/GetTable, date literals UTC and year-first
    /// via <see cref="DaslDateLiteral"/> (Outlook parses them in the MACHINE locale, so a
    /// month-first literal transposes day and month on any day 12 or lower - see that
    /// class for the measurement), single quotes escaped by doubling. Terms are ANDed; each
    /// term matches subject OR body by default, narrowable to one of them via
    /// <see cref="SearchIn"/> (D40 - the same three scopes the index tier offers).
    /// A trailing '*' marks a prefix stem and is matched via LIKE substring in BOTH
    /// engines (ci_phrasematch is whole-word and would miss the stem's continuations).
    /// <para>
    /// IT NO LONGER CARRIES A MESSAGE-CLASS FILTER (audit gap B3, maintainer decision
    /// 2026-08-18). Every filter used to open with <c>PR_MESSAGE_CLASS like 'IPM.Note%'</c>,
    /// which quietly made the one mode chosen FOR completeness the only tier blind to
    /// bounce reports and read receipts (<c>REPORT.IPM.Note.*</c>), meeting requests and
    /// their responses, posts and sharing invitations - while the freshness sweep beside it
    /// returned all of them. Item class no longer excludes anything anywhere; see
    /// <see cref="OutlookAI.Core.Mapi.MailItemAdmission"/> for the rule and why it is
    /// stated in one place.
    /// </para>
    /// </summary>
    public static class ExhaustiveDaslFilter
    {
        private const string SubjectProp = "\"urn:schemas:httpmail:subject\"";
        private const string BodyProp = "\"urn:schemas:httpmail:textdescription\"";
        private const string DateReceivedProp = "\"urn:schemas:httpmail:datereceived\"";
        private const string MessageClassProp = "\"http://schemas.microsoft.com/mapi/proptag/0x001A001E\"";

        /// <summary>
        /// The predicate used when there is nothing else to restrict on: it matches every
        /// item class, so it narrows nothing.
        /// <para>
        /// PR_MESSAGE_CLASS is mandatory on every MAPI message, so <c>like '%'</c> over it
        /// is true for every row a mail folder can hand back - deliberately the SAME column
        /// the removed filter used, because that column is the one already proven to be
        /// present, selectable and filterable on every folder this scan opens. It is public
        /// so T1 can pin the shape: a predicate that silently stopped matching would empty
        /// out exactly the unbounded folder scan that has no terms to fall back on.
        /// </para>
        /// <para>
        /// REVIEWED 2026-08-18, and KEPT. The question was whether some other always-true
        /// predicate is more clearly correct. None is, and the reason is structural rather
        /// than a preference: MAPI documents the result of a restriction over a property the
        /// message does NOT have as UNDEFINED - not false - so a row whose property is absent
        /// is admitted or dropped at the provider's discretion, and no predicate over that
        /// property can promise either. The distinction is not pedantry (corrected
        /// 2026-08-19): "excluded" invites the inference that <c>NOT (...)</c> therefore
        /// ADMITS such a row, which is exactly the reasoning that made a broken fix for the
        /// sweep's date restriction look viable - negating an undefined value leaves it
        /// undefined. Swapping this predicate for a comparison, or for a negation, moves that
        /// risk around instead of removing it, and lands on syntax this codebase has never
        /// emitted. The only construction that
        /// removes it is no restriction at all (<c>Folder.GetTable()</c> with no argument),
        /// which was considered and not taken: PR_MESSAGE_CLASS is required on every MAPI
        /// message and is what Outlook itself reads to decide which item type to hand back,
        /// so the absent case is not reachable through the object model, while dropping the
        /// filter would change a COM call site that cannot be exercised outside a live
        /// profile and would make the reported scan engine ("like") a claim about matching
        /// that never happened. What stays unverified is narrow and loud: whether the
        /// provider reads the pattern <c>%</c> as "any string". If it does not, GetTable
        /// throws, the folder is counted skipped and a coverage gap is raised - the tier
        /// says it lost the folder rather than pretending it was empty.
        /// </para>
        /// </summary>
        public const string AdmitEveryClassClause = MessageClassProp + " like '%'";

        /// <summary>
        /// Longest accepted search term. THE SAME limit the index tier enforces, not a
        /// second copy of it: the two modes answer the same user query and must agree on
        /// which terms are valid.
        /// </summary>
        private const int MaxTermLength = IndexSearch.WsSqlBuilder.MaxTermLength;

        /// <summary>
        /// Builds the full restriction. Always a valid non-empty filter, even without terms
        /// or date bounds (the CALLER enforces the exhaustive-mode bounding rules): with
        /// nothing to restrict on it emits <see cref="AdmitEveryClassClause"/>, which
        /// selects every item in the folder.
        /// </summary>
        /// <param name="resumeAtOrBeforeUtc">
        /// The date cursor a resumed scan continues from, as an INCLUSIVE upper bound. It is
        /// separate from <paramref name="beforeUtc"/> and not a replacement for it: the
        /// caller's own bound is exclusive and stays exactly as they wrote it, while this one
        /// has to admit the cursor instant itself so that items sharing it are reachable at
        /// all. Everything already returned AT that instant is excluded by EntryID afterwards
        /// (the cursor's tie set), because a date alone cannot separate them.
        /// <para>
        /// Folding resumption into the restriction rather than into a row position is the
        /// whole reason the date rung is the preferred one: the provider evaluates it as part
        /// of the query it was going to evaluate anyway, so skipping costs nothing, and
        /// nothing depends on an unsorted table returning rows in the same order twice - which
        /// MAPI documents it need not do.
        /// </para>
        /// </param>
        public static string Build(
            IReadOnlyList<string>? terms,
            DateTime? sinceUtc,
            DateTime? beforeUtc,
            ExhaustiveEngine engine,
            SearchIn searchIn = SearchInValues.Default,
            DateTime? resumeAtOrBeforeUtc = null)
        {
            if (sinceUtc.HasValue && beforeUtc.HasValue && sinceUtc.Value >= beforeUtc.Value)
            {
                throw new ArgumentException("sinceUtc must lie before beforeUtc.", nameof(sinceUtc));
            }

            List<string> clauses = new List<string>();

            if (sinceUtc.HasValue)
            {
                clauses.Add(DateReceivedProp + " >= '" + DaslDateLiteral.FormatUtc(sinceUtc.Value) + "'");
            }

            if (beforeUtc.HasValue)
            {
                clauses.Add(DateReceivedProp + " < '" + DaslDateLiteral.FormatUtc(beforeUtc.Value) + "'");
            }

            if (resumeAtOrBeforeUtc.HasValue)
            {
                clauses.Add(DateReceivedProp + " <= '" + DaslDateLiteral.FormatUtc(resumeAtOrBeforeUtc.Value) + "'");
            }

            if (terms != null)
            {
                for (int i = 0; i < terms.Count; i++)
                {
                    clauses.Add(BuildTermClause(
                        ValidateTerm(terms[i], "terms[" + i.ToString(CultureInfo.InvariantCulture) + "]"),
                        engine,
                        searchIn));
                }
            }

            if (clauses.Count == 0)
            {
                // A folder-scoped scan with no terms and no dates ("show me this folder")
                // is a legal call, and `@SQL=` with no predicate is not a restriction
                // Outlook accepts. The class clause used to fill this slot by accident of
                // being a filter; now that it is gone, the slot needs a predicate that
                // excludes nothing.
                return "@SQL=(" + AdmitEveryClassClause + ")";
            }

            StringBuilder sql = new StringBuilder("@SQL=");
            for (int i = 0; i < clauses.Count; i++)
            {
                if (i > 0)
                {
                    sql.Append(" AND ");
                }

                sql.Append('(').Append(clauses[i]).Append(')');
            }

            return sql.ToString();
        }

        private static string BuildTermClause(string term, ExhaustiveEngine engine, SearchIn searchIn)
        {
            bool prefixStem = term.EndsWith("*", StringComparison.Ordinal);
            string value = prefixStem ? term.Substring(0, term.Length - 1) : term;

            string subjectClause;
            string bodyClause;
            if (prefixStem || engine == ExhaustiveEngine.Like)
            {
                string like = " like '%" + EscapeLikeValue(value) + "%'";
                subjectClause = SubjectProp + like;
                bodyClause = BodyProp + like;
            }
            else
            {
                string match = " ci_phrasematch '" + EscapeDaslValue(value) + "'";
                subjectClause = SubjectProp + match;
                bodyClause = BodyProp + match;
            }

            switch (searchIn)
            {
                case SearchIn.SubjectAndBody:
                    return subjectClause + " OR " + bodyClause;
                case SearchIn.SubjectOnly:
                    return subjectClause;
                case SearchIn.BodyOnly:
                    return bodyClause;
                default:
                    throw new ArgumentException("Unknown SearchIn value.", nameof(searchIn));
            }
        }

        /// <summary>
        /// Validates a term with the same charset contract as the index search
        /// (letters/digits plus space @ . _ - ' +; optional single trailing '*'), keeping
        /// agent-facing semantics identical across modes and the DASL literal safe.
        /// </summary>
        public static string ValidateTerm(string term, string parameterName)
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
                throw new ArgumentException("Bare '*' is not a valid term.", parameterName);
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

        private static string EscapeDaslValue(string value)
        {
            return value.Replace("'", "''");
        }

        private static string EscapeLikeValue(string value)
        {
            // '_' is a single-char LIKE wildcard and part of the allowed term charset -
            // bracket-escape it ('%'/'[' are already rejected by the charset).
            return value.Replace("'", "''").Replace("_", "[_]");
        }
    }
}
