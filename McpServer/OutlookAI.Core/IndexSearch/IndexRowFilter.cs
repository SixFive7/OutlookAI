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
    /// and decide here instead - an ATTACHMENT-child row (<c>/at=</c>) is kept whatever its
    /// kind, and a MESSAGE-level row (URL without <c>/at=</c>) is kept because it is a
    /// message.
    /// </para>
    /// <para>
    /// MESSAGE ROWS ARE NO LONGER JUDGED ON THEIR KIND (gap B3, maintainer decision
    /// 2026-08-18). They used to need <c>email</c>, which dropped every meeting request and
    /// response (they index as <c>calendar</c>) from a search while the freshness sweep
    /// beside them returned all of them - so the same query gave a different answer
    /// depending on which tier reached the mail first. Item class narrows nothing anywhere
    /// now (<see cref="OutlookAI.Core.Mapi.MailItemAdmission"/>).
    /// </para>
    /// <para>
    /// WHAT THAT COSTS, stated rather than hidden. The index carries no folder-TYPE column,
    /// so this tier cannot draw the line the COM tiers draw structurally by only entering
    /// mail folders: a store-scoped statement emits no kind predicate at all, so an
    /// appointment in the Calendar or a card in Contacts is now admitted alongside the
    /// meeting request in the Inbox that the decision was about. That is over-return, which
    /// a caller can see (every such hit carries <c>itemClass</c>) and filter, in place of
    /// under-return, which nothing downstream can detect. An UNSCOPED statement is
    /// unaffected in practice: it still carries <see cref="UnscopedKinds"/>, which the
    /// provider needs anyway to stay off the file system.
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

        /// <summary>
        /// The one kind that marks a message-level row as ordinary mail. It is no longer an
        /// admission test - only <see cref="KindFilter.MailKindOnly"/>, the store-discovery
        /// shape, still reads it that way.
        /// </summary>
        public const string EmailKind = Mapi.MailItemAdmission.EmailKind;

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
                case KindFilter.MailKindOnly:
                    // The store-discovery shape, and the only surviving kind test: it wants
                    // a row it is CERTAIN is a mail message, to read a store prefix off it.
                    return !attachment && HasEmailKind(hit.Kinds);
                case KindFilter.AttachmentsOnly:
                    // attachment_hits_only: every attachment row, whatever its kind.
                    return attachment;
                case KindFilter.MessagesAndAttachments:
                    // Everything the statement offered in the mapi namespace.
                    return true;
                case KindFilter.MessagesOnly:
                    // Every message-level row, whatever its item class. The mapi-namespace
                    // check above is what keeps this honest: without a kind test it is the
                    // ONLY thing standing between this filter and the file system.
                    return !attachment;
                default:
                    throw new ArgumentException("Unknown KindFilter value.", nameof(kinds));
            }
        }

        /// <summary>
        /// SQL <c>TOP</c> to request so that post-filtering still yields
        /// <paramref name="requestedTop"/> admitted rows. Dropping the Kind predicate lets
        /// rows through that this filter then removes (file-system rows on an unscoped
        /// statement; the shape mismatches of <see cref="KindFilter.AttachmentsOnly"/> and
        /// <see cref="KindFilter.MailKindOnly"/>), so a bare <c>TOP n</c> could return fewer
        /// than n admitted rows while more existed. The over-fetch is unchanged by the
        /// widening of message-row admission, which only ever removes drops: the factors
        /// were already generous relative to the measured rates (the message-level calendar
        /// rows they were sized for, 0.3-1.2% of a scoped folder, are now KEPT) and cost
        /// little - the provider orders and caps the same way, and the drain stops as soon
        /// as enough rows are admitted.
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
