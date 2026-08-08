using System;
using System.Collections.Generic;

namespace OutlookAI.Core.IndexSearch
{
    /// <summary>
    /// Row admission for the index tier, applied in code AFTER the SQL returns.
    /// <para>
    /// Why this exists (measured read-only probe 2026-07-27, v3.MD section 0.8 block (q)):
    /// the shipped predicate <c>(System.Kind='email' OR System.Kind='document')</c> assumed
    /// every attachment-content row is a <c>document</c>. It is not. An attachment row
    /// carries the kind of the ATTACHMENT, so images index as <c>picture</c>, embedded
    /// <c>.msg</c>/<c>.eml</c> as <c>communication</c>, <c>.ics</c> invites as
    /// <c>calendar</c>, plus <c>music</c>/<c>video</c>. Measured across five stores,
    /// <b>709 of 3,139 attachment rows (22.6%)</b> were structurally unmatchable: a term
    /// living only inside an image, an embedded message or an invite could never find its
    /// parent mail.
    /// </para>
    /// <para>
    /// Shape shipped (the block-(q) recommendation): drop the Kind predicate from the SQL
    /// and decide here instead - a MESSAGE-level row (URL without <c>/at=</c>) is kept only
    /// when its kinds contain <c>email</c>; an ATTACHMENT-child row (<c>/at=</c>) is kept
    /// whatever its kind. Message-level non-mail rows (meeting requests index as
    /// <c>calendar</c>) are therefore still excluded exactly as before.
    /// </para>
    /// <para>
    /// The mapi-namespace check is load-bearing: without a Kind predicate an UNSCOPED
    /// statement would also match the file system, so only <c>mapi16://</c> rows are
    /// admitted here (an <c>.eml</c> file on disk indexes as kind <c>email</c> and would
    /// otherwise pass). <see cref="WsSqlBuilder"/> keeps an enumerated kind predicate on
    /// unscoped statements so the provider narrows too; this filter is the correctness
    /// guarantee in both cases.
    /// </para>
    /// </summary>
    public static class IndexRowFilter
    {
        /// <summary>Only rows in the MAPI namespace are mail; everything else is the file system.</summary>
        public const string MapiUrlPrefix = "mapi16://";

        /// <summary>URL marker of an attachment-content row: <c>&lt;messageUrl&gt;/at=...</c>.</summary>
        public const string AttachmentMarker = "/at=";

        /// <summary>The one kind a message-level row must carry to be mail.</summary>
        public const string EmailKind = "email";

        /// <summary>
        /// Kinds enumerated on an UNSCOPED statement, where SCOPE cannot fence the query to
        /// the MAPI namespace. This is the measured union of message-level and
        /// attachment-row kinds on this profile (block (q)); anything outside it is still
        /// caught by <see cref="Keep"/>, this list only keeps the provider selective.
        /// </summary>
        public static readonly IReadOnlyList<string> UnscopedKinds = new[]
        {
            "email",
            "document",
            "picture",
            "communication",
            "calendar",
            "music",
            "video",
        };

        /// <summary>True when <paramref name="itemUrl"/> addresses an Outlook item.</summary>
        public static bool IsMapiRow(string? itemUrl)
        {
            return itemUrl != null && itemUrl.StartsWith(MapiUrlPrefix, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// True when <paramref name="itemUrl"/> is an attachment-content row. Deliberately a
        /// raw URL test rather than <c>MapiItemUrl.TryParse</c>: a malformed
        /// attachment URL must still be recognised as an attachment, not promoted to a
        /// message-level row that would then be judged on its (attachment) kind.
        /// </summary>
        public static bool IsAttachmentRow(string? itemUrl)
        {
            return itemUrl != null && itemUrl.IndexOf(AttachmentMarker, StringComparison.Ordinal) >= 0;
        }

        /// <summary>True when the row's kinds contain <c>email</c> (System.Kind is case-insensitive).</summary>
        public static bool HasEmailKind(IReadOnlyList<string>? kinds)
        {
            if (kinds == null)
            {
                return false;
            }

            for (int i = 0; i < kinds.Count; i++)
            {
                if (string.Equals(kinds[i], EmailKind, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>Admits one mapped row under <paramref name="kinds"/>.</summary>
        public static bool Keep(IndexHit hit, KindFilter kinds)
        {
            if (hit == null)
            {
                throw new ArgumentNullException(nameof(hit));
            }

            if (!IsMapiRow(hit.ItemUrl))
            {
                return false;
            }

            bool attachment = IsAttachmentRow(hit.ItemUrl);
            switch (kinds)
            {
                case KindFilter.EmailOnly:
                    // Messages only - an attachment row is never a message.
                    return !attachment && HasEmailKind(hit.Kinds);
                case KindFilter.DocumentsOnly:
                    // attachment_hits_only: every attachment row, whatever its kind.
                    return attachment;
                case KindFilter.EmailAndDocuments:
                    return attachment || HasEmailKind(hit.Kinds);
                default:
                    throw new ArgumentException("Unknown KindFilter value.", nameof(kinds));
            }
        }

        /// <summary>
        /// SQL <c>TOP</c> to request so that post-filtering still yields
        /// <paramref name="requestedTop"/> admitted rows. Dropping the Kind predicate lets
        /// rows through that this filter then removes (message-level calendar rows under a
        /// scope; file-system rows without one), so a bare <c>TOP n</c> could return fewer
        /// than n mail rows while more existed. Over-fetch factors are deliberately
        /// generous relative to the measured drop rates (calendar-class message rows are
        /// 0.3-1.2% of a scoped folder) and cost little: the provider orders and caps the
        /// same way, and the drain stops as soon as enough rows are admitted.
        /// </summary>
        public static int ComputeSqlTop(int requestedTop, bool scoped, int maxTop)
        {
            if (requestedTop < 1)
            {
                throw new ArgumentException("requestedTop must be positive.", nameof(requestedTop));
            }

            if (maxTop < 1)
            {
                throw new ArgumentException("maxTop must be positive.", nameof(maxTop));
            }

            int factor = scoped ? 2 : 4;
            int floor = scoped ? 10 : 20;
            long widened = (long)requestedTop * factor + floor;
            if (widened > maxTop)
            {
                widened = maxTop;
            }

            return widened < requestedTop ? requestedTop : (int)widened;
        }
    }
}
