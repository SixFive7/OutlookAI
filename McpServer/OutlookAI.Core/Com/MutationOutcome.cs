using System.Globalization;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// The one place this product decides what to SAY about whether an operation took
    /// effect, and the three-value vocabulary it says it in.
    /// <para>
    /// <b>Why this exists.</b> Every claim of atomicity in this server used to be a
    /// hand-written sentence that decided, locally, what to assert about effect - and a
    /// sentence written for the refusals it sat next to was then attached to a catch-all
    /// that reached much further. The 2026-08-19 audit found sixteen such claims, every
    /// one wrong the same way. The classification that answers the question already
    /// existed in <see cref="ComSessionOperations"/>, and one site already used it
    /// correctly (<c>ComHostSupervisor.DescribeInterruption</c>). This is that shape,
    /// lifted out so the other sites share it instead of re-deciding.
    /// </para>
    /// <para>
    /// <b>What it does NOT own.</b> Only the OPENING sentence - the part that is the same
    /// wherever it is said. The remedy differs sharply per path (update means repeat the
    /// call, send means check the Outbox, create means check Drafts, discard means check
    /// Deleted Items) and a shared sentence cannot carry it, so every caller appends its
    /// own. Specific remedies are what an agent can act on; that is what makes
    /// <c>MailService.DescribeSendOutcomeUnknown</c> good, and it is deliberately kept.
    /// </para>
    /// <para>
    /// <b>Fail-closed, inherited from the classification.</b> An operation name that was
    /// never classified reads as mutating, so the failure mode of forgetting is an
    /// over-cautious "check before you retry" and never a false "nothing happened".
    /// </para>
    /// </summary>
    public static class MutationOutcome
    {
        /// <summary>
        /// The call left the user's mail exactly as it was. Only ever asserted where the
        /// code PROVES it: a read, or a refusal decided before the first write.
        /// </summary>
        public const string Unchanged = "unchanged";

        /// <summary>
        /// The change happened and stands, even though the call is being reported as a
        /// failure. The move that could not be audited and the answer too large to return
        /// are both this: the work succeeded and only the report did not.
        /// </summary>
        public const string Applied = "applied";

        /// <summary>
        /// Nobody can state whether it took effect. This is the value the whole vocabulary
        /// exists for: a boolean cannot carry it, so before this field a three-state
        /// outcome was being reported as "not done".
        /// </summary>
        public const string Unknown = "unknown";

        /// <summary>
        /// The outcome of an operation that did NOT answer - killed, interrupted, or
        /// refused by Outlook part-way through. A read that never answered still changed
        /// nothing; a mutation that never answered is the unknown case.
        /// </summary>
        public static string ForInterrupted(string? operationName)
        {
            return ComSessionOperations.IsRetryable(operationName) ? Unchanged : Unknown;
        }

        /// <summary>
        /// The outcome of an operation that RAN TO COMPLETION even though the caller is
        /// being told about a failure. The failure is downstream of the effect - an answer
        /// too large to frame, an audit line that could not be written - so a mutation's
        /// effect stands.
        /// </summary>
        public static string ForCompleted(string? operationName)
        {
            return ComSessionOperations.IsRetryable(operationName) ? Unchanged : Applied;
        }

        /// <summary>
        /// The shared OPENING sentence for an operation that did not answer. Callers append
        /// their own remedy clause; this half only states what is known.
        /// </summary>
        public static string DescribeInterrupted(string? operationName)
        {
            return ComSessionOperations.IsRetryable(operationName)
                ? string.Format(
                    CultureInfo.InvariantCulture,
                    "'{0}' only READS mail, so nothing was changed by it.",
                    Name(operationName))
                : string.Format(
                    CultureInfo.InvariantCulture,
                    "'{0}' CHANGES mail and it did not answer, so whether it took effect is UNKNOWN.",
                    Name(operationName));
        }

        /// <summary>
        /// The shared OPENING sentence for an operation that COMPLETED and whose answer was
        /// then lost. The distinction from <see cref="DescribeInterrupted"/> is the whole
        /// point: here a repeat performs the change a SECOND time, so the advice inverts.
        /// </summary>
        public static string DescribeAnswerLost(string? operationName)
        {
            return ComSessionOperations.IsRetryable(operationName)
                ? string.Format(
                    CultureInfo.InvariantCulture,
                    "'{0}' only READS mail: it succeeded, nothing was changed, and only the answer was lost.",
                    Name(operationName))
                : string.Format(
                    CultureInfo.InvariantCulture,
                    "'{0}' CHANGES mail: the work SUCCEEDED and its effect STANDS - only the answer was lost, "
                    + "so do NOT repeat it.",
                    Name(operationName));
        }

        /// <summary>
        /// A name for the sentence even when the caller had none. An empty quoted string
        /// would read as a bug; "the operation" reads as what it is.
        /// </summary>
        private static string Name(string? operationName)
        {
            return string.IsNullOrWhiteSpace(operationName) ? "the operation" : operationName!;
        }
    }
}
