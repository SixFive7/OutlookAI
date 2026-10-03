using System;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// A service-layer failure that also STATES whether the operation took effect.
    /// <para>
    /// <b>Why the field and not just the sentence.</b> The 2026-08-19 atomicity audit found
    /// sixteen claims of non-effect that were wrong, and the cheapest of them to get wrong
    /// again is the one nothing can test: a sentence. A value from
    /// <see cref="Com.MutationOutcome"/> travelling beside the message is assertable - T1
    /// can require that an unclassified COM failure never reports
    /// <see cref="Com.MutationOutcome.Unchanged"/> - which is the assertion nobody could
    /// write while the claim lived only in prose.
    /// </para>
    /// <para>
    /// <b>Deliberately an <see cref="InvalidOperationException"/>.</b> Every site that
    /// throws this used to throw exactly that, and several callers catch it by that type
    /// (the move/archive batch reports such a failure per item rather than failing the
    /// call). Deriving keeps all of them working, so adding the field changed no control
    /// flow anywhere.
    /// </para>
    /// <para>
    /// It never crosses the COM host pipe: every site that raises it is in the SERVER
    /// parent, past the point where the child's answer has already been rebuilt.
    /// </para>
    /// </summary>
    public sealed class OperationOutcomeException : InvalidOperationException
    {
        /// <summary>Creates the failure.</summary>
        /// <param name="outcome">One of the <see cref="Com.MutationOutcome"/> values.</param>
        /// <param name="message">What the caller is told, remedy included.</param>
        public OperationOutcomeException(string outcome, string message)
            : base(message)
        {
            Outcome = outcome;
        }

        /// <summary>Creates the failure with an inner cause.</summary>
        public OperationOutcomeException(string outcome, string message, Exception inner)
            : base(message, inner)
        {
            Outcome = outcome;
        }

        /// <summary>
        /// Whether the user's mail was left alone, changed, or left in a state nobody can
        /// state: <c>unchanged</c> | <c>applied</c> | <c>unknown</c>.
        /// </summary>
        public string Outcome { get; }
    }
}
