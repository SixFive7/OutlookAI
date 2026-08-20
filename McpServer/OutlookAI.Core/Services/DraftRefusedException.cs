using System;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// Thrown when <c>update_draft</c> or <c>discard_draft</c> refuses to touch an item
    /// (v3.MD D46, S1 v3): the item is not a mail item, has already been sent, does not
    /// live in a Drafts folder, or - for a discard - was not created or last updated by
    /// THIS server process.
    /// <para>
    /// <b>Read <see cref="Reason"/> before believing anything about effect.</b> This doc
    /// used to end "Nothing was changed or deleted when this is thrown", which is true of
    /// every NAMED reason - each is decided before the first write - and false of
    /// <c>com_failure</c>, which is raised by the catch-all around the entire COM sequence
    /// and can arrive after the body was rewritten or after the delete was issued. The
    /// message says which case it is; <c>MailService.ComFailureRefusal</c> is the code, and
    /// the tool layer branches on it to choose the advice and the outcome field.
    /// </para>
    /// <para>
    /// A refusal is NEVER a silent no-op: <see cref="Reason"/> is the machine-readable
    /// code (also written to the audit log and surfaced on the wire), and the message
    /// says what would make the call succeed.
    /// </para>
    /// </summary>
    public sealed class DraftRefusedException : Exception
    {
        /// <summary>Creates the refusal.</summary>
        public DraftRefusedException(string reason, string message)
            : base(message)
        {
            Reason = reason;
        }

        /// <summary>Machine-readable refusal code (e.g. "not_created_by_this_server").</summary>
        public string Reason { get; }
    }
}
