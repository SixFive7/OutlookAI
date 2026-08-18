using System;
using System.Collections.Generic;

namespace OutlookAI.Core.Mapi
{
    /// <summary>
    /// THE admission rule for every search tier: <b>an item's CLASS never excludes it</b>.
    /// What bounds a search is the folder it looks in, never what kind of thing it finds
    /// there.
    /// <para>
    /// WHY THIS EXISTS (audit gap B3, maintainer decision 2026-08-18 - "unify all three,
    /// and prefer returning EVERYTHING where possible"). The three tiers that can answer one
    /// search used to admit three different item sets, and nothing in the payload said so:
    /// the index tier required <c>System.Kind</c> to contain <c>email</c>, the freshness
    /// sweep filtered nothing at all, and the exhaustive COM scan filtered
    /// <c>PR_MESSAGE_CLASS like 'IPM.Note%'</c>. So a meeting request the sweep returned
    /// today vanished once it was indexed, and the one mode that exists BECAUSE
    /// completeness matters was the only one blind to NDRs and read receipts - "did my mail
    /// bounce?" was unanswerable exactly where a user looks hardest.
    /// </para>
    /// <para>
    /// THE RULE IS "NO RULE", DELIBERATELY. The widest of the three was the sweep, which
    /// filtered nothing, so unifying upwards means the other two stop filtering too. An
    /// allowlist of admitted classes was considered and rejected for a reason that decides
    /// it: the SystemIndex carries no message-class column at all, so an allowlist could
    /// only ever be enforced in the COM tiers - which would replace one asymmetry with
    /// another, in the same payload, for the same query. See
    /// <see cref="Admits(string?)"/>, which is that decision written as code so it has one
    /// location and one test.
    /// </para>
    /// <para>
    /// WHAT EACH TIER CAN STILL NOT REACH, since the rule cannot make a tier see what its
    /// engine does not carry:
    /// <list type="bullet">
    /// <item>The COM tiers (freshness sweep, exhaustive scan) only enter folders whose
    /// <c>DefaultItemType</c> is <c>olMailItem</c>. That is unchanged and is not a class
    /// filter - it is where mail lives. An appointment sitting in the Calendar is therefore
    /// out of their reach, while a meeting REQUEST sitting in the Inbox is in it.</item>
    /// <item>The index tier has no folder-type column and no message-class column, so it
    /// cannot draw that same line: a message-level row is admitted whatever its kind, which
    /// takes in the meeting requests this decision is about AND the calendar/contact items
    /// of the folders the COM tiers never open. That is a REPORTED fact, not a silent one -
    /// each hit carries <c>itemClass</c> (see <see cref="DescribeIndexRowClass"/>), and the
    /// direction of the error is the safe one: over-return, which a caller can see and
    /// filter, rather than under-return, which nothing downstream can detect.</item>
    /// </list>
    /// </para>
    /// </summary>
    public static class MailItemAdmission
    {
        /// <summary>
        /// The message class of ordinary mail. Everything from here on is a PREFIX test
        /// (<c>IPM.Note.SMIME.MultipartSigned</c> is ordinary mail that happens to be
        /// signed), never an equality test.
        /// </summary>
        public const string OrdinaryMailClass = "IPM.Note";

        /// <summary>
        /// Prefix that marks an index-tier class description apart from a real MAPI class:
        /// the index never opens an item, so the best it can say is the row's
        /// <c>System.Kind</c>. A caller must be able to tell "Outlook says this is a
        /// meeting request" from "the index says this row is calendar-shaped".
        /// </summary>
        public const string IndexKindPrefix = "kind:";

        /// <summary>What <see cref="DescribeIndexRowClass"/> says when the row carried no kind at all.</summary>
        public const string UnknownIndexKind = IndexKindPrefix + "unknown";

        /// <summary>The one <c>System.Kind</c> value that marks a message-level row as ordinary mail.</summary>
        public const string EmailKind = "email";

        /// <summary>
        /// Message classes this rule is REQUIRED to admit, named so the decision is testable
        /// rather than merely described. Every one of them was excluded by at least one tier
        /// before 2026-08-18, and each is mail a user asks about by name: bounce reports and
        /// read receipts (<c>REPORT.*</c>), meeting requests and their responses, posts, and
        /// sharing invitations.
        /// <para>
        /// It is documentation and a test corpus, NOT a filter - nothing reads it to decide
        /// admission, because <see cref="Admits(string?)"/> admits everything including the
        /// classes nobody thought to list here. Treating it as a filter is precisely the
        /// allowlist this decision rejected.
        /// </para>
        /// </summary>
        public static readonly IReadOnlyList<string> ClassesTheOldFiltersDropped = new[]
        {
            "REPORT.IPM.Note.NDR",
            "REPORT.IPM.Note.IPNRN",
            "REPORT.IPM.Note.DR",
            "IPM.Schedule.Meeting.Request",
            "IPM.Schedule.Meeting.Canceled",
            "IPM.Schedule.Meeting.Resp.Pos",
            "IPM.Schedule.Meeting.Resp.Neg",
            "IPM.Schedule.Meeting.Resp.Tent",
            "IPM.Post",
            "IPM.Sharing",
        };

        /// <summary>
        /// Whether a search tier admits an item of <paramref name="messageClass"/>. Always
        /// true, including for a class that could not be read at all.
        /// <para>
        /// A method that cannot return false looks like a placeholder and is not one: it is
        /// the decision itself, in the single place the three tiers can point at. Written as
        /// code rather than as a comment for two reasons. It is greppable - a future
        /// narrowing has to delete a call site that says what it is deleting, instead of
        /// quietly adding a class test next to the item loop, which is how the three tiers
        /// drifted apart in the first place. And it is testable: T1 drives every class in
        /// <see cref="ClassesTheOldFiltersDropped"/> through it, so re-narrowing the rule
        /// fails the build with the reason attached.
        /// </para>
        /// </summary>
        public static bool Admits(string? messageClass)
        {
            _ = messageClass;
            return true;
        }

        /// <summary>
        /// Whether <paramref name="messageClass"/> is ordinary mail - the case that needs no
        /// mention in a payload, because a mail search returning mail is not news.
        /// <para>
        /// Prefix-matched at a class-name boundary, so <c>IPM.Note.SMIME</c> is ordinary and
        /// <c>IPM.Notification.Meeting</c> is not. A null or blank class is NOT ordinary: it
        /// means the read failed, and reporting an unidentified item as ordinary mail would
        /// be the payload asserting something nobody established.
        /// </para>
        /// </summary>
        public static bool IsOrdinaryMailClass(string? messageClass)
        {
            if (string.IsNullOrWhiteSpace(messageClass))
            {
                return false;
            }

            string trimmed = messageClass!.Trim();
            if (!trimmed.StartsWith(OrdinaryMailClass, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return trimmed.Length == OrdinaryMailClass.Length || trimmed[OrdinaryMailClass.Length] == '.';
        }

        /// <summary>
        /// The <c>itemClass</c> a COM-tier hit (freshness sweep, exhaustive scan, thread
        /// walk) reports, or null when there is nothing worth saying.
        /// <para>
        /// Null for ordinary mail and null when the class could not be read, so the field is
        /// absent from the overwhelming majority of hits and costs nothing; its PRESENCE is
        /// the signal that this hit is not an ordinary mail. Widening admission without this
        /// would hand callers a result set they cannot reason about - which is the tidiness
        /// half of the price, paid here rather than by narrowing the answer.
        /// </para>
        /// </summary>
        public static string? DescribeComItemClass(string? messageClass)
        {
            if (string.IsNullOrWhiteSpace(messageClass) || IsOrdinaryMailClass(messageClass))
            {
                return null;
            }

            return messageClass!.Trim();
        }

        /// <summary>
        /// The <c>itemClass</c> an INDEX-tier hit reports, or null when there is nothing
        /// worth saying. The index never opens the item, so the answer is the row's
        /// <c>System.Kind</c> behind <see cref="IndexKindPrefix"/> - never a bare class name,
        /// which would claim an authority this tier does not have.
        /// <para>
        /// Null for an attachment-content row (<c>isAttachmentHit</c> already says what it
        /// is, and the kind describes the ATTACHMENT rather than the mail carrying it) and
        /// null when the kinds contain <c>email</c>. A message-level row with no kind at all
        /// reports <see cref="UnknownIndexKind"/>: such rows used to be dropped, so silence
        /// about them would be the old filter surviving as a blank field.
        /// </para>
        /// </summary>
        public static string? DescribeIndexRowClass(IReadOnlyList<string>? kinds, bool isAttachmentRow)
        {
            if (isAttachmentRow)
            {
                return null;
            }

            if (kinds == null || kinds.Count == 0)
            {
                return UnknownIndexKind;
            }

            List<string> named = new List<string>(kinds.Count);
            for (int i = 0; i < kinds.Count; i++)
            {
                string kind = kinds[i];
                if (string.IsNullOrWhiteSpace(kind))
                {
                    continue;
                }

                if (string.Equals(kind, EmailKind, StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                named.Add(kind.Trim());
            }

            return named.Count == 0 ? UnknownIndexKind : IndexKindPrefix + string.Join("+", named.ToArray());
        }
    }
}
