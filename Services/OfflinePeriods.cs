namespace OutlookAI.Services
{
    /// <summary>
    /// THE ELEVEN OFFLINE PERIODS - Q143, decided by the maintainer: Outlook's "Download email for
    /// the past" offers exactly the eleven windows Microsoft documents (KB 3115009's table, KB
    /// 2733062's policy values, the eleven items of the administrative template's "Cached Exchange
    /// Mode Sync Settings"), each written as one (months, days) pair, and nothing else is ever
    /// written. KB 3115009: a value that is missing or invalid makes Outlook use its default of 12
    /// months - so a stray 18 would silently be a year. Measured on the Exchange VM, 2026-11-17, Office
    /// LTSC 2024: the slider has exactly these eleven positions, an account's own pair is written by
    /// Outlook as these same eleven pairs, and 36 and 60 survive a restart as "3 years" and "5 years".
    ///
    /// <para>
    /// ONE TABLE FOR BOTH TREES (Q154: no duplicated code). The add-in compiles this file: its
    /// <c>CachedModeSettings</c> writes the user's own period for new accounts from it.
    /// <c>OutlookAI.Core</c> LINKS it: its <c>AccountOfflinePeriodWriter</c> writes an existing
    /// account's own pair from it. The test project links neither copy of its own - it reaches Core's
    /// through InternalsVisibleTo, as with <c>OfficeVersions</c>.
    /// </para>
    ///
    /// <para>
    /// FRAMEWORK-NEUTRAL and INTERNAL, under the rules <c>Services\OfficeVersions.cs</c> documents:
    /// net48 (the add-in, C# 7.3) and net48 AND net10 inside Core, nullable-enabled with warnings as
    /// errors - so no nullable annotations, nothing returned or assigned as null, nothing newer than
    /// C# 7.3. Internal, so the copies in the two assemblies the test assembly can see never collide.
    /// </para>
    /// </summary>
    internal static class OfflinePeriods
    {
        /// <summary>All mail: (0, 0). The shipped default, and what OutlookAI's search needs (D25).</summary>
        internal static readonly OfflinePeriod All = new OfflinePeriod(0, 0, "All");

        // In the slider's order, shortest first: the administrative template's eleven items.
        private static readonly OfflinePeriod[] Periods =
        {
            new OfflinePeriod(0, 3, "3 days"),
            new OfflinePeriod(0, 7, "1 week"),
            new OfflinePeriod(0, 14, "2 weeks"),
            new OfflinePeriod(1, 0, "1 month"),
            new OfflinePeriod(3, 0, "3 months"),
            new OfflinePeriod(6, 0, "6 months"),
            new OfflinePeriod(12, 0, "1 year"),
            new OfflinePeriod(24, 0, "2 years"),
            new OfflinePeriod(36, 0, "3 years"),
            new OfflinePeriod(60, 0, "5 years"),
            All,
        };

        /// <summary>How many there are: eleven.</summary>
        internal static int Count
        {
            get { return Periods.Length; }
        }

        /// <summary>The period at a position in the slider's order, 0 to 10. Throws outside it.</summary>
        internal static OfflinePeriod At(int index)
        {
            return Periods[index];
        }

        /// <summary>The position of (months, days) in the slider's order, or -1 for a pair Microsoft does not document.</summary>
        internal static int IndexOf(int months, int days)
        {
            for (int i = 0; i < Periods.Length; i++)
            {
                if (Periods[i].Months == months && Periods[i].Days == days)
                    return i;
            }
            return -1;
        }

        /// <summary>True for exactly the eleven (months, days) pairs, and for nothing else.</summary>
        internal static bool IsOffered(int months, int days)
        {
            return IndexOf(months, days) >= 0;
        }
    }

    /// <summary>One of the eleven offline periods: what Outlook calls it, and the pair it is written as.</summary>
    internal sealed class OfflinePeriod
    {
        internal OfflinePeriod(int months, int days, string label)
        {
            Months = months;
            Days = days;
            Label = label;
        }

        /// <summary><c>SyncWindowSetting</c> / profile property 0x6649: 0 for the three day periods and for All.</summary>
        internal int Months { get; }

        /// <summary><c>SyncWindowSettingDays</c> / profile property 0x665A: 3, 7 or 14 with months 0; 0 for every other period.</summary>
        internal int Days { get; }

        internal string Label { get; }

        public override string ToString()
        {
            return Label;
        }
    }
}
