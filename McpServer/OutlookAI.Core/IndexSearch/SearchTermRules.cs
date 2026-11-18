using System;
using System.Globalization;

namespace OutlookAI.Core.IndexSearch
{
    /// <summary>
    /// What a search TERM may be, for every mode that answers a query: the index search
    /// (<see cref="WsSqlBuilder"/>) and the exhaustive COM scan (<c>ExhaustiveDaslFilter</c>) both
    /// validate their terms here, so a term one mode accepts the other accepts too, with the same
    /// words when it refuses one. Each tier kept its own copy of these rules until the code-sharing
    /// refactor's stage S6 (Q154); the copies differed only in how they refused a bare '*'.
    /// </summary>
    public static class SearchTermRules
    {
        /// <summary>
        /// Longest accepted search term, in characters, for every mode.
        /// <para>
        /// It was two private 128s, one per tier. Raising either alone would have made indexed and
        /// exhaustive search disagree about which terms are valid - a term accepted by one mode and
        /// rejected by the other, for the same query, with nothing to notice it.
        /// </para>
        /// </summary>
        public const int MaxTermLength = 128;

        /// <summary>
        /// The term, trimmed, or an <see cref="ArgumentException"/> naming
        /// <paramref name="parameterName"/>. Rules: not null and not blank after trimming; at most
        /// <see cref="MaxTermLength"/> characters; '*' only as the final character - a prefix match -
        /// and never alone, which the index refuses outright (<c>CONTAINS('*')</c> fails with
        /// 0x80041605); every other character a letter, a digit, a space or one of @ . _ - ' +.
        /// No '"': each builder adds its own phrase quoting, and escapes the single quote.
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
    }
}
