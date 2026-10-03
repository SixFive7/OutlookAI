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
        /// <para>
        /// "Local" is this machine's zone, <see cref="TimeZoneInfo.Local"/> - the zone the
        /// Outlook beside this process converts into. This overload is the only one the product
        /// calls; the zone is a parameter of <see cref="FromItemValue(DateTime?, TimeZoneInfo)"/>
        /// only so that a test can hold it still.
        /// </para>
        /// </summary>
        public static DateTime? FromItemValue(DateTime? value)
        {
            return FromItemValue(value, TimeZoneInfo.Local);
        }

        /// <summary>
        /// <see cref="FromItemValue(DateTime?)"/> with the zone the wall time is read in
        /// passed in rather than taken from the machine. The product passes
        /// <see cref="TimeZoneInfo.Local"/> and nothing else.
        /// <para>
        /// <b>Why the seam exists (Q95, 2026-10-03).</b> The conversion is the identity on a
        /// machine whose zone is UTC, so a test that relied on the machine's own zone could not
        /// tell a converted value from an unconverted one there - and GitHub's runners ARE UTC,
        /// so CI failed on a test that passed on the maintainer's UTC+2 machine, while the
        /// product was right on both. A test now names a zone of its own and gets the same
        /// answer on every machine.
        /// </para>
        /// <para>
        /// <b>The same answer the product always gave.</b> Before the seam this was
        /// <c>DateTime.SpecifyKind(moment, DateTimeKind.Local).ToUniversalTime()</c>, which
        /// subtracts the offset <see cref="TimeZoneInfo.Local"/> holds for that wall time,
        /// never throws on a wall time the zone skips at a spring-forward, takes a wall time it
        /// repeats at a fall-back as standard time, and clamps a result that would leave the
        /// <see cref="DateTime"/> range to the end of it. <see cref="TimeZoneInfo.GetUtcOffset(DateTime)"/>
        /// on the same wall time is that same offset, gap and overlap included - so this is the
        /// same arithmetic made explicit, with the clamp kept. <c>ComDateValueTests</c> holds
        /// the two equal across the local zone's own transitions.
        /// </para>
        /// </summary>
        /// <param name="value">The item's wall time, as COM handed it over.</param>
        /// <param name="localZone">The zone that wall time is in.</param>
        public static DateTime? FromItemValue(DateTime? value, TimeZoneInfo localZone)
        {
            if (localZone == null)
            {
                throw new ArgumentNullException(nameof(localZone));
            }

            if (!value.HasValue)
            {
                return null;
            }

            DateTime moment = value.Value;
            if (moment.Kind == DateTimeKind.Utc)
            {
                return moment;
            }

            DateTime wallTime = DateTime.SpecifyKind(moment, DateTimeKind.Unspecified);
            long utcTicks = wallTime.Ticks - localZone.GetUtcOffset(wallTime).Ticks;
            if (utcTicks < DateTime.MinValue.Ticks)
            {
                utcTicks = DateTime.MinValue.Ticks;
            }
            else if (utcTicks > DateTime.MaxValue.Ticks)
            {
                utcTicks = DateTime.MaxValue.Ticks;
            }

            return new DateTime(utcTicks, DateTimeKind.Utc);
        }
    }
}
