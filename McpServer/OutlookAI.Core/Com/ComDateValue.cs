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
        /// A value read out of an Outlook <c>Table</c> column that was added by its NAMESPACE
        /// reference (<c>urn:schemas:...</c>, <c>http://schemas.microsoft.com/mapi/...</c>), as
        /// UTC. Returns null for anything that is not a date, which is the ordinary reading for
        /// a row whose date property was never set. A column added by its EXPLICIT built-in name
        /// reads differently - use <see cref="FromTableValue(object?, string)"/>, which decides by
        /// the column.
        /// <para>
        /// An <see cref="DateTimeKind.Unspecified"/> kind is taken as ALREADY UTC, which is what
        /// a namespace-referenced column holds: measured 2026-10-03 (below), and what
        /// <see cref="DaslDateLiteral.FormatUtc"/> assumes of the restriction that selected the
        /// row.
        /// </para>
        /// <para>
        /// <b>QUESTIONS.md Q11, settled by measurement 2026-10-03</b> (the first live run on a test
        /// guest, <c>OutlookAI-Unindexed</c>, Office LTSC 2024 16.0.17932, machine at UTC+2,
        /// <c>T2/LiveTableSortProbeTests</c>): a table reports a date in the zone its COLUMN
        /// SPELLING asks for, not one zone for every column. Under the explicit name
        /// <c>ReceivedTime</c> the raw value equalled the opened item's own local
        /// <c>ReceivedTime</c> (13:44:18 for an item received 11:44:18Z) on both stores read;
        /// under <c>urn:schemas:httpmail:datereceived</c> the same folder's rows read in UTC
        /// (11:44:18). The exhaustive scan adds the explicit name first, so treating every table
        /// value as UTC put its resume cursor one offset LATE east of UTC - which re-admitted a
        /// row the cursor's tie set does not suppress, and produced a duplicate on that run - and
        /// one offset EARLY west of it, which skips mail and calls the scan complete.
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
        /// A value read out of an Outlook <c>Table</c> row, as UTC, read in the zone the column
        /// it came from was added in: a NAMESPACE reference holds UTC
        /// (<see cref="FromTableValue(object?)"/>), an EXPLICIT built-in name such as
        /// <c>ReceivedTime</c> holds LOCAL wall time and is converted as an item value
        /// (<see cref="FromItemValue(DateTime?, TimeZoneInfo)"/>). Measured 2026-10-03 - see
        /// <see cref="FromTableValue(object?)"/> for the run and the numbers.
        /// </summary>
        /// <param name="value">The row's value for the column.</param>
        /// <param name="columnProperty">The spelling the column was added under, exactly as passed to <c>Columns.Add</c>.</param>
        public static DateTime? FromTableValue(object? value, string columnProperty)
        {
            return FromTableValue(value, columnProperty, TimeZoneInfo.Local);
        }

        /// <summary>
        /// <see cref="FromTableValue(object?, string)"/> with the zone an explicit-name column's
        /// wall time is in passed in rather than taken from the machine. The product passes
        /// <see cref="TimeZoneInfo.Local"/> and nothing else; the parameter exists so a test can
        /// hold the zone still (the Q95 seam <see cref="FromItemValue(DateTime?, TimeZoneInfo)"/>
        /// already has).
        /// </summary>
        public static DateTime? FromTableValue(object? value, string columnProperty, TimeZoneInfo localZone)
        {
            if (columnProperty == null)
            {
                throw new ArgumentNullException(nameof(columnProperty));
            }

            if (localZone == null)
            {
                throw new ArgumentNullException(nameof(localZone));
            }

            if (!(value is DateTime moment))
            {
                return null;
            }

            return IsNamespaceReference(columnProperty)
                ? FromTableValue(moment)
                : FromItemValue(moment, localZone);
        }

        /// <summary>
        /// True when a table column's spelling is a NAMESPACE reference - a DAV <c>urn:</c> name
        /// or a MAPI <c>http://schemas.microsoft.com/...</c> proptag - rather than an explicit
        /// built-in property name. It is the spelling, not the property, that decides the zone a
        /// table reports a date in (Q11, measured 2026-10-03).
        /// </summary>
        public static bool IsNamespaceReference(string columnProperty)
        {
            if (columnProperty == null)
            {
                throw new ArgumentNullException(nameof(columnProperty));
            }

            return columnProperty.StartsWith("urn:", StringComparison.OrdinalIgnoreCase)
                || columnProperty.StartsWith("http://", StringComparison.OrdinalIgnoreCase)
                || columnProperty.StartsWith("https://", StringComparison.OrdinalIgnoreCase);
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
