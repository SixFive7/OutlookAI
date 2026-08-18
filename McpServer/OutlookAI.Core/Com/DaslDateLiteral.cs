using System;
using System.Globalization;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// THE one place a <see cref="DateTime"/> becomes a DASL (<c>@SQL=</c>) date literal.
    /// The exhaustive-scan filter and the freshness sweep both come through here; before
    /// this class they each carried their own copy of the format string, which is how the
    /// same defect could ship twice and be fixed once.
    /// <para>
    /// WHY THE SHAPE MATTERS. Outlook does NOT parse a DASL date literal in a fixed
    /// culture: "Outlook evaluates date-time values according to the time format, short
    /// date format, and long date format settings in the Regional and Language Options
    /// applet in the Windows Control Panel" (Microsoft, "Filtering Items Using a Date-time
    /// Comparison"). Both formatters used to emit the invariant US <c>MM/dd/yyyy</c>, so on
    /// a day-first machine every date whose DAY is 12 or lower - roughly 40% of days - was
    /// read with day and month TRANSPOSED, and the query silently answered about a
    /// different window. Measured on this machine (nl-NL, short date <c>d-M-yyyy</c>),
    /// 2026-08-18: an exhaustive search bounded to 1-5 August returned 158 items received
    /// between 8 January and 7 June and ZERO inside the requested window, because
    /// <c>08/01/2026</c> read as 8 January and <c>08/06/2026</c> as 8 June. The same query
    /// for 13-15 August was exact, because a day above 12 cannot be a month.
    /// </para>
    /// <para>
    /// WHY YEAR-FIRST RATHER THAN THE CURRENT CULTURE. Formatting in
    /// <see cref="CultureInfo.CurrentCulture"/> is the shape the documentation implies, and
    /// it measured exact here - but it is only correct while the process culture still
    /// equals the Windows user locale (invariant-globalization mode, a host-set default
    /// thread culture, or a non-Gregorian calendar culture each break it, and each breaks
    /// it SILENTLY, the same way this defect did). A 4-digit leading year cannot be a day
    /// or a month, so <c>yyyy-MM-dd</c> is unambiguous under a day-first, month-first or
    /// year-first locale alike; it is also the literal the Windows Search tier already
    /// emits (<c>WsSqlBuilder.FormatUtc</c>), so one shape now covers both query languages,
    /// and it is a constant string that T1 can pin under ANY thread culture.
    /// </para>
    /// <para>
    /// MEASURED FORMAT MATRIX (read-only probe, this machine, <c>Folder.GetTable</c> AND
    /// <c>Items.Restrict</c> over one folder, ground truth from a full unfiltered table
    /// scan; ambiguous window 13 items, unambiguous window 22):
    /// <c>MM/dd/yyyy HH:mm:ss</c> and <c>MM/dd/yyyy HH:mm</c> returned 121 wrong items and
    /// 0 right ones; <c>yyyy-MM-dd HH:mm:ss</c>, <c>yyyy-MM-dd HH:mm</c>,
    /// <c>yyyy/MM/dd HH:mm</c>, <c>dd-MM-yyyy HH:mm</c> and the current-culture form all
    /// returned exactly the ground truth. Both APIs agreed on every row, so this literal is
    /// safe for both. ISO 8601 with a <c>T</c> separator (<c>yyyy-MM-ddTHH:mm:ss</c>) is a
    /// TRAP and is deliberately not used: it does not throw, it returns the whole folder
    /// (366 of 368 rows) - an unparsed date literal fails silently in BOTH directions.
    /// </para>
    /// <para>
    /// Seconds are kept even though the same Microsoft page warns that Outlook "evaluates
    /// time according to that specified time format without seconds": with the year-first
    /// literal, the with-seconds and without-seconds forms measured identical (13/13 and
    /// 22/22), and the item-probe window in <c>OutlookComSession</c> is expressed in
    /// seconds.
    /// </para>
    /// </summary>
    public static class DaslDateLiteral
    {
        /// <summary>
        /// The literal's format: year-first with a 4-digit year, so no locale can read the
        /// day as the month. Always rendered with <see cref="CultureInfo.InvariantCulture"/>
        /// - the digits and separators must not follow the calling thread's culture.
        /// </summary>
        public const string Format = "yyyy-MM-dd HH:mm:ss";

        /// <summary>
        /// Renders a DASL date literal (WITHOUT the surrounding single quotes). DASL
        /// compares date-time properties in UTC, so a <see cref="DateTimeKind.Local"/>
        /// value is converted first; <see cref="DateTimeKind.Unspecified"/> is taken as
        /// already-UTC, which is the contract every caller in this assembly follows.
        /// </summary>
        public static string FormatUtc(DateTime value)
        {
            DateTime utc = value.Kind == DateTimeKind.Local ? value.ToUniversalTime() : value;
            return utc.ToString(Format, CultureInfo.InvariantCulture);
        }
    }
}
