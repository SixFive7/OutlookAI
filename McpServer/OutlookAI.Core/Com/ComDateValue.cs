using System;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// The ONE place this solution decides what time zone a date coming out of Outlook is
    /// in. Two sources, two documented answers, and a single helper for each so that no two
    /// call sites can derive the same instant differently.
    /// <para>
    /// <b>Why this type exists.</b> It was written after exactly that happened. The live
    /// tripwire census read a table's date column as already-UTC while
    /// <c>OutlookComSession.ReadRowDate</c> called <c>ToUniversalTime</c> on the same
    /// variant, which treats it as local. A COM-marshalled <c>VT_DATE</c> always arrives
    /// with <see cref="DateTimeKind.Unspecified"/>, so the two could never both be right,
    /// and the disagreement was invisible because each side was internally consistent.
    /// </para>
    /// <para>
    /// <b>Why it is not one function with a flag.</b> A table value and an object-model
    /// property are genuinely different readings, not one reading with an option: Microsoft
    /// documents the <c>Table</c> object as returning date-time values in UTC and the
    /// object model as returning local time. Folding them together is how a caller ends up
    /// passing the wrong flag and getting an answer that is wrong by exactly the machine's
    /// UTC offset, which is the least visible size an error of this kind can have.
    /// </para>
    /// </summary>
    public static class ComDateValue
    {
        /// <summary>
        /// A value read out of an Outlook <c>Table</c> row, as UTC. Returns null for
        /// anything that is not a date, which is the ordinary reading for a row whose date
        /// property was never set.
        /// <para>
        /// An <see cref="DateTimeKind.Unspecified"/> kind is taken as ALREADY UTC. Three
        /// things point the same way: Microsoft documents the <c>Table</c> object as
        /// returning date-time values in UTC (unlike the object model, which returns local
        /// time); <see cref="DaslDateLiteral.FormatUtc"/> already treats an unspecified kind
        /// as UTC, so the restriction that selected the row and the value read back out of
        /// it agree; and it is the SAFE direction for the one caller where being wrong
        /// costs mail. A resumed exhaustive scan uses this value as an inclusive "at or
        /// before" bound, so reading a UTC instant as local moves the bound EARLIER by the
        /// local offset and silently skips the mail in that window, while reading a local
        /// instant as UTC moves it LATER and merely re-reads rows the chain already
        /// suppresses by EntryID.
        /// </para>
        /// <para>
        /// <b>Still to be confirmed by measurement</b> (QUESTIONS.md Q11): the
        /// <c>T2/LiveTableDateKindProbe</c> reading settles it against a real profile by
        /// comparing this value with the same item's <c>MailItem.ReceivedTime</c>. If that
        /// run shows tables reporting LOCAL time, this method is the single line to change
        /// and every caller follows.
        /// </para>
        /// </summary>
        public static DateTime? FromTableValue(object? value)
        {
            if (!(value is DateTime moment))
            {
                return null;
            }

            if (moment.Kind == DateTimeKind.Utc)
            {
                return moment;
            }

            if (moment.Kind == DateTimeKind.Local)
            {
                return moment.ToUniversalTime();
            }

            return DateTime.SpecifyKind(moment, DateTimeKind.Utc);
        }

        /// <summary>
        /// A date read off an OPENED Outlook item (<c>MailItem.ReceivedTime</c> and its
        /// siblings), as UTC. Null in, null out.
        /// <para>
        /// The opposite default to <see cref="FromTableValue"/>, and deliberately so: the
        /// object model returns local wall time, and COM hands it over with
        /// <see cref="DateTimeKind.Unspecified"/> just as it does a table value, so the kind
        /// alone cannot tell the two apart. Which method to call is decided by where the
        /// value CAME FROM, which is why they are separate names rather than one call that
        /// inspects the kind.
        /// </para>
        /// </summary>
        public static DateTime? FromItemValue(DateTime? value)
        {
            if (!value.HasValue)
            {
                return null;
            }

            DateTime moment = value.Value;
            if (moment.Kind == DateTimeKind.Utc)
            {
                return moment;
            }

            return DateTime.SpecifyKind(moment, DateTimeKind.Local).ToUniversalTime();
        }
    }
}
