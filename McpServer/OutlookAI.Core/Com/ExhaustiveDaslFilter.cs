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
    /// section-12 rules: ci_* only in Restrict/GetTable, date literals UTC invariant
    /// "MM/dd/yyyy HH:mm:ss", single quotes escaped by doubling. Terms are ANDed; each
    /// term matches subject OR body by default, narrowable to one of them via
    /// <see cref="TermScope"/> (D40 - the same three scopes the index tier offers).
    /// A trailing '*' marks a prefix stem and is matched via LIKE substring in BOTH
    /// engines (ci_phrasematch is whole-word and would miss the stem's continuations).
    /// Every filter carries an IPM.Note message-class clause so only mail items are
    /// enumerated.
    /// </summary>
    public static class ExhaustiveDaslFilter
    {
        private const string SubjectProp = "\"urn:schemas:httpmail:subject\"";
        private const string BodyProp = "\"urn:schemas:httpmail:textdescription\"";
        private const string DateReceivedProp = "\"urn:schemas:httpmail:datereceived\"";
        private const string MessageClassProp = "\"http://schemas.microsoft.com/mapi/proptag/0x001A001E\"";
        private const int MaxTermLength = 128;

        /// <summary>
        /// Builds the full restriction. At least the mail-only message-class clause is
        /// always present, so the result is a valid non-empty filter even without terms
        /// or date bounds (the CALLER enforces the exhaustive-mode bounding rules).
        /// </summary>
        public static string Build(
            IReadOnlyList<string>? terms,
            DateTime? sinceUtc,
            DateTime? beforeUtc,
            ExhaustiveEngine engine,
            TermScope termScope = TermScopes.Default)
        {
            if (sinceUtc.HasValue && beforeUtc.HasValue && sinceUtc.Value >= beforeUtc.Value)
            {
                throw new ArgumentException("sinceUtc must lie before beforeUtc.", nameof(sinceUtc));
            }

            List<string> clauses = new List<string>
            {
                MessageClassProp + " like 'IPM.Note%'",
            };

            if (sinceUtc.HasValue)
            {
                clauses.Add(DateReceivedProp + " >= '" + FormatDaslUtc(sinceUtc.Value) + "'");
            }

            if (beforeUtc.HasValue)
            {
                clauses.Add(DateReceivedProp + " < '" + FormatDaslUtc(beforeUtc.Value) + "'");
            }

            if (terms != null)
            {
                for (int i = 0; i < terms.Count; i++)
                {
                    clauses.Add(BuildTermClause(
                        ValidateTerm(terms[i], "terms[" + i.ToString(CultureInfo.InvariantCulture) + "]"),
                        engine,
                        termScope));
                }
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

        private static string BuildTermClause(string term, ExhaustiveEngine engine, TermScope termScope)
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

            switch (termScope)
            {
                case TermScope.SubjectAndBody:
                    return subjectClause + " OR " + bodyClause;
                case TermScope.SubjectOnly:
                    return subjectClause;
                case TermScope.BodyOnly:
                    return bodyClause;
                default:
                    throw new ArgumentException("Unknown TermScope value.", nameof(termScope));
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

        /// <summary>DASL date literal: UTC, invariant US format (documented DASL semantics).</summary>
        private static string FormatDaslUtc(DateTime value)
        {
            DateTime utc = value.Kind == DateTimeKind.Local ? value.ToUniversalTime() : value;
            return utc.ToString("MM/dd/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
        }
    }
}
