using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;
using OutlookAI.Core.AddIn;

namespace OutlookAI.Core.Services
{
    /// <summary>One section of an Outlook profile, as read (data only).</summary>
    public sealed class ProfileSection
    {
        /// <summary>Creates a section.</summary>
        public ProfileSection(string key, IReadOnlyDictionary<string, object?> values)
        {
            Key = key ?? throw new ArgumentNullException(nameof(key));
            Values = values ?? throw new ArgumentNullException(nameof(values));
        }

        /// <summary>The section's handle: its HKCU-relative registry path for the live registry, any unique string for a fake.</summary>
        public string Key { get; }

        /// <summary>Value name to raw value, as stored - a profile property is REG_BINARY.</summary>
        public IReadOnlyDictionary<string, object?> Values { get; }
    }

    /// <summary>
    /// Raw access to the sections of the default Outlook profile - the T1 seam for
    /// <see cref="AccountOfflinePeriodWriter"/>. <see cref="LiveProfileSectionRegistry"/> is the real one.
    /// </summary>
    public interface IProfileSectionRegistry
    {
        /// <summary>Every section, read NOW - never cached, so a read-back sees what the registry holds.</summary>
        IReadOnlyList<ProfileSection> ReadSections();

        /// <summary>Sets a REG_BINARY value in an EXISTING section; never creates a section.</summary>
        void SetBinary(string key, string valueName, byte[] value);
    }

    /// <summary>
    /// The live registry: every subkey of
    /// HKCU\Software\Microsoft\Office\&lt;major&gt;\Outlook\Profiles\&lt;default profile&gt; (every
    /// profile's, when no <c>DefaultProfile</c> is recorded - the rule <see cref="LiveProfileAccountRegistry"/>
    /// reads by). Loads only what <see cref="AccountOfflinePeriodWriter"/> decides by: the two period
    /// values and the PT_UNICODE values an address can be in.
    /// </summary>
    public sealed class LiveProfileSectionRegistry : IProfileSectionRegistry
    {
        /// <inheritdoc />
        public IReadOnlyList<ProfileSection> ReadSections()
        {
            List<ProfileSection> sections = new List<ProfileSection>();
            string outlookRoot = OutlookProfileRegistry.OutlookRootKeyPath;
            using RegistryKey? outlook = Registry.CurrentUser.OpenSubKey(outlookRoot);
            if (outlook == null)
            {
                return sections;
            }

            string? defaultProfile = outlook.GetValue("DefaultProfile") as string;
            using RegistryKey? profiles = outlook.OpenSubKey(OutlookProfileRegistry.ProfilesSubKeyName);
            if (profiles == null)
            {
                return sections;
            }

            IEnumerable<string> profileNames = defaultProfile != null ? new[] { defaultProfile } : profiles.GetSubKeyNames();
            foreach (string profileName in profileNames)
            {
                using RegistryKey? profile = profiles.OpenSubKey(profileName);
                if (profile == null)
                {
                    continue;
                }

                foreach (string sectionName in profile.GetSubKeyNames())
                {
                    using RegistryKey? section = profile.OpenSubKey(sectionName);
                    if (section == null)
                    {
                        continue;
                    }

                    Dictionary<string, object?> values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                    foreach (string valueName in section.GetValueNames())
                    {
                        if (AccountOfflinePeriodWriter.IsValueRead(valueName))
                        {
                            values[valueName] = section.GetValue(valueName);
                        }
                    }

                    sections.Add(new ProfileSection(
                        outlookRoot + "\\" + OutlookProfileRegistry.ProfilesSubKeyName + "\\" + profileName + "\\" + sectionName,
                        values));
                }
            }

            return sections;
        }

        /// <inheritdoc />
        public void SetBinary(string key, string valueName, byte[] value)
        {
            using RegistryKey? section = Registry.CurrentUser.OpenSubKey(key, writable: true);
            if (section == null)
            {
                throw new InvalidOperationException("The profile section no longer exists: " + key);
            }

            section.SetValue(valueName, value, RegistryValueKind.Binary);
        }
    }

    /// <summary>How one <see cref="AccountOfflinePeriodWriter.Write"/> ended.</summary>
    public enum AccountOfflinePeriodOutcome
    {
        /// <summary>Both values written and read back equal. Outlook applies it at its next start.</summary>
        Written,

        /// <summary>The account already had that period: nothing written.</summary>
        AlreadySet,

        /// <summary>The pair is not one of the eleven: nothing read, nothing written.</summary>
        NotOffered,

        /// <summary>No section holds both the period and the address - no such Exchange account here. Nothing written.</summary>
        NoSuchAccount,

        /// <summary>More than one section holds both: which is the account's is not guessed. Nothing written.</summary>
        SeveralSections,

        /// <summary>A write failed; the months value was put back where it had been written alone.</summary>
        NotWritten,

        /// <summary>Written, but the values read back differently.</summary>
        ReadBackMismatch,
    }

    /// <summary>What one write did, and what it found before it.</summary>
    public sealed class AccountOfflinePeriodResult
    {
        /// <summary>Creates a result.</summary>
        public AccountOfflinePeriodResult(AccountOfflinePeriodOutcome outcome, string message, string? sectionKey, int? monthsBefore, int? daysBefore)
        {
            Outcome = outcome;
            Message = message ?? string.Empty;
            SectionKey = sectionKey;
            MonthsBefore = monthsBefore;
            DaysBefore = daysBefore;
        }

        /// <summary>How it ended.</summary>
        public AccountOfflinePeriodOutcome Outcome { get; }

        /// <summary>One sentence for the user, or for a refusal.</summary>
        public string Message { get; }

        /// <summary>The account's section, when exactly one was found.</summary>
        public string? SectionKey { get; }

        /// <summary>The account's months value before the write (null when absent or not read).</summary>
        public int? MonthsBefore { get; }

        /// <summary>The account's days value before the write (null when absent or not read).</summary>
        public int? DaysBefore { get; }

        /// <summary>The account now has the period asked for.</summary>
        public bool Succeeded => Outcome == AccountOfflinePeriodOutcome.Written || Outcome == AccountOfflinePeriodOutcome.AlreadySet;

        /// <summary>Outlook must restart before the new period takes effect (measured: it applies a written pair at its next start).</summary>
        public bool RestartNeeded => Outcome == AccountOfflinePeriodOutcome.Written;
    }

    /// <summary>
    /// WRITES AN EXISTING EXCHANGE ACCOUNT'S OWN OFFLINE PERIOD ("Download email for the past") INTO
    /// THE OUTLOOK PROFILE - the half of Q145 (a) that waited for a measurement. The user's own
    /// <c>SyncWindowSetting</c> reaches only accounts added later (KB 3115009); an account that exists
    /// keeps its period in its own Exchange provider section of the profile, which this writes.
    /// <para>
    /// <b>What it rests on - MEASURED on the Exchange VM, 2026-11-17, Office LTSC 2024, the test
    /// account only (Testbed/guest/Measure-OfflineWindow.ps1).</b> The period lives in the account's
    /// own Exchange provider section,
    /// HKCU\Software\Microsoft\Office\16.0\Outlook\Profiles\&lt;profile&gt;\&lt;section&gt;, as two
    /// PT_LONG properties stored REG_BINARY, four bytes little-endian: 0x6649 (value
    /// <c>00036649</c>) the months and 0x665A (<c>0003665a</c>) the days. Outlook always writes them
    /// as a pair, and its slider's eleven positions are exactly the eleven pairs of
    /// <see cref="OfflinePeriods"/> (Q143) - 36 and 60 survive a restart as "3 years" and "5 years" -
    /// and moving it changes no other value in the section. A pair written with Outlook CLOSED was
    /// honoured at the next start and not rewritten; so was one written while Outlook RAN: Outlook did
    /// not overwrite it on a graceful exit, and the next start showed it. Whether the running session
    /// applies it before a restart was not measured, and Outlook's own dialog asks for a restart - so a
    /// write says the period takes effect when Outlook next starts. Also not measured: whether Outlook
    /// then really downloads that much more mail.
    /// </para>
    /// <para>
    /// <b>Which section.</b> The one that holds 0x6649 AND carries the account's address as a whole
    /// PT_UNICODE value (UTF-16LE, NUL-terminated - the address, trimmed, case-insensitive, never a
    /// substring: <see cref="ProfileAccountEntries.AddressEquals"/>). None is a refusal - there is no
    /// such Exchange account in the profile, and POP3 and IMAP accounts have no such section; several
    /// is a refusal too, never a guess. Months and days are written together, months first, as one
    /// of the eleven pairs and nothing else; a days write that fails puts the months value back.
    /// </para>
    /// <para>
    /// <b>Who calls it.</b> Nobody yet. The add-in, which lives in Outlook and is where a user picks a
    /// period, references Core since the code-sharing refactor's first stage (Q167 (a)), but may call
    /// only the namespace OutlookAI.Core.AddIn (Q164 (a)), which this must move into first; until
    /// then the add-in's seam, <c>IAccountOfflinePeriods</c> in that namespace, keeps telling the
    /// user how to set the period in Account Settings.
    /// </para>
    /// </summary>
    public static class AccountOfflinePeriodWriter
    {
        /// <summary>PT_LONG 0x6649: the period's months.</summary>
        public const string MonthsValueName = "00036649";

        /// <summary>PT_LONG 0x665A: the period's days.</summary>
        public const string DaysValueName = "0003665a";

        /// <summary>The value-name prefix of a PT_UNICODE property (type 0x001F) - where an address can be.</summary>
        public const string UnicodeValuePrefix = "001f";

        /// <summary>True for exactly the eleven (months, days) pairs Outlook documents (Q143).</summary>
        public static bool IsOffered(int months, int days) => OfflinePeriods.IsOffered(months, days);

        /// <summary>The name Outlook's slider gives a pair, or null for a pair that is not one of the eleven.</summary>
        public static string? LabelOf(int months, int days)
        {
            int index = OfflinePeriods.IndexOf(months, days);
            return index < 0 ? null : OfflinePeriods.At(index).Label;
        }

        /// <summary>The values the live reader loads: the two period values and every PT_UNICODE one.</summary>
        public static bool IsValueRead(string valueName)
        {
            return string.Equals(valueName, MonthsValueName, StringComparison.OrdinalIgnoreCase)
                || string.Equals(valueName, DaysValueName, StringComparison.OrdinalIgnoreCase)
                || (valueName != null && valueName.StartsWith(UnicodeValuePrefix, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>The sections that hold 0x6649 and carry <paramref name="address"/> as a whole PT_UNICODE value. Pure.</summary>
        public static IReadOnlyList<ProfileSection> SectionsOf(IReadOnlyList<ProfileSection> sections, string address)
        {
            if (sections == null)
            {
                throw new ArgumentNullException(nameof(sections));
            }

            return sections.Where(s => Find(s.Values, MonthsValueName, out _) && CarriesAddress(s, address)).ToList();
        }

        /// <summary>
        /// Writes (<paramref name="months"/>, <paramref name="days"/>) as the own offline period of the
        /// Exchange account whose address is <paramref name="address"/> - see the type summary for every
        /// refusal. Never throws for a registry failure: that is the result's.
        /// </summary>
        public static AccountOfflinePeriodResult Write(IProfileSectionRegistry registry, string address, int months, int days)
        {
            if (registry == null)
            {
                throw new ArgumentNullException(nameof(registry));
            }

            string? label = LabelOf(months, days);
            if (label == null)
            {
                return new AccountOfflinePeriodResult(AccountOfflinePeriodOutcome.NotOffered,
                    "(" + months + ", " + days + ") is not one of Outlook's eleven offline periods, so nothing was written.", null, null, null);
            }

            if (string.IsNullOrWhiteSpace(address))
            {
                return new AccountOfflinePeriodResult(AccountOfflinePeriodOutcome.NoSuchAccount, "No account address was given, so nothing was written.", null, null, null);
            }

            IReadOnlyList<ProfileSection> found;
            try
            {
                found = SectionsOf(registry.ReadSections(), address);
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                return new AccountOfflinePeriodResult(AccountOfflinePeriodOutcome.NotWritten,
                    "The Outlook profile could not be read (" + ex.Message + "), so nothing was written.", null, null, null);
            }

            if (found.Count == 0)
            {
                return new AccountOfflinePeriodResult(AccountOfflinePeriodOutcome.NoSuchAccount,
                    "No section of the Outlook profile holds the offline period of " + address.Trim()
                    + ": it is not an Exchange account of this profile. Nothing was written.", null, null, null);
            }

            if (found.Count > 1)
            {
                return new AccountOfflinePeriodResult(AccountOfflinePeriodOutcome.SeveralSections,
                    found.Count + " sections of the Outlook profile hold the offline period of " + address.Trim()
                    + ", and which one is the account's is not guessed. Nothing was written.", null, null, null);
            }

            ProfileSection section = found[0];
            int? monthsBefore = DecodeLong(Value(section.Values, MonthsValueName));
            int? daysBefore = DecodeLong(Value(section.Values, DaysValueName));
            if (monthsBefore == months && daysBefore == days)
            {
                return new AccountOfflinePeriodResult(AccountOfflinePeriodOutcome.AlreadySet,
                    address.Trim() + " already keeps " + Phrase(label) + " offline.", section.Key, monthsBefore, daysBefore);
            }

            try
            {
                registry.SetBinary(section.Key, MonthsValueName, EncodeLong(months));
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                return new AccountOfflinePeriodResult(AccountOfflinePeriodOutcome.NotWritten,
                    "The offline period of " + address.Trim() + " could not be written (" + ex.Message + "). Nothing was changed.",
                    section.Key, monthsBefore, daysBefore);
            }

            try
            {
                registry.SetBinary(section.Key, DaysValueName, EncodeLong(days));
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                // Never leave half a pair: Outlook always writes both, and (12, 7) is no period at all.
                string putBack = "its months value could not be put back either";
                if (monthsBefore.HasValue)
                {
                    try
                    {
                        registry.SetBinary(section.Key, MonthsValueName, EncodeLong(monthsBefore.Value));
                        putBack = "its months value was put back";
                    }
                    catch (Exception again) when (!(again is OutOfMemoryException))
                    {
                        putBack = "its months value could not be put back either (" + again.Message + ")";
                    }
                }

                return new AccountOfflinePeriodResult(AccountOfflinePeriodOutcome.NotWritten,
                    "The days half of the offline period of " + address.Trim() + " could not be written (" + ex.Message + "); " + putBack + ".",
                    section.Key, monthsBefore, daysBefore);
            }

            // Read back from the registry itself, not from what was handed to it.
            ProfileSection? after = null;
            try
            {
                after = registry.ReadSections().FirstOrDefault(s => ProfileAccountEntries.KeyEquals(s.Key, section.Key));
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                return new AccountOfflinePeriodResult(AccountOfflinePeriodOutcome.ReadBackMismatch,
                    "The offline period of " + address.Trim() + " was written, but could not be read back (" + ex.Message + ").",
                    section.Key, monthsBefore, daysBefore);
            }

            int? monthsAfter = after == null ? null : DecodeLong(Value(after.Values, MonthsValueName));
            int? daysAfter = after == null ? null : DecodeLong(Value(after.Values, DaysValueName));
            if (monthsAfter != months || daysAfter != days)
            {
                return new AccountOfflinePeriodResult(AccountOfflinePeriodOutcome.ReadBackMismatch,
                    "The offline period of " + address.Trim() + " was written, but reads back as (" + Show(monthsAfter) + ", " + Show(daysAfter)
                    + ") instead of (" + months + ", " + days + ").", section.Key, monthsBefore, daysBefore);
            }

            return new AccountOfflinePeriodResult(AccountOfflinePeriodOutcome.Written,
                address.Trim() + " now keeps " + Phrase(label) + " offline. Restart Outlook for it to take effect.",
                section.Key, monthsBefore, daysBefore);
        }

        /// <summary>A PT_LONG as the profile stores it: four bytes, little-endian.</summary>
        public static byte[] EncodeLong(int value)
        {
            return new[] { (byte)value, (byte)(value >> 8), (byte)(value >> 16), (byte)(value >> 24) };
        }

        /// <summary>A PT_LONG read back: four bytes, little-endian, or a REG_DWORD. Null for anything else.</summary>
        public static int? DecodeLong(object? raw)
        {
            if (raw is int number)
            {
                return number;
            }

            if (raw is byte[] bytes && bytes.Length == 4)
            {
                return bytes[0] | (bytes[1] << 8) | (bytes[2] << 16) | (bytes[3] << 24);
            }

            return null;
        }

        /// <summary>A PT_UNICODE value as text: UTF-16LE bytes with their terminating NULs removed, or a string as stored.</summary>
        public static string? DecodeUnicode(object? raw)
        {
            if (raw is string text)
            {
                return text.TrimEnd('\0');
            }

            if (raw is byte[] bytes && bytes.Length % 2 == 0)
            {
                return Encoding.Unicode.GetString(bytes).TrimEnd('\0');
            }

            return null;
        }

        private static bool CarriesAddress(ProfileSection section, string address)
        {
            foreach (KeyValuePair<string, object?> value in section.Values)
            {
                if (value.Key.StartsWith(UnicodeValuePrefix, StringComparison.OrdinalIgnoreCase)
                    && ProfileAccountEntries.AddressEquals(DecodeUnicode(value.Value), address))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool Find(IReadOnlyDictionary<string, object?> values, string name, out object? value)
        {
            foreach (KeyValuePair<string, object?> pair in values)
            {
                if (string.Equals(pair.Key, name, StringComparison.OrdinalIgnoreCase))
                {
                    value = pair.Value;
                    return true;
                }
            }

            value = null;
            return false;
        }

        private static object? Value(IReadOnlyDictionary<string, object?> values, string name)
        {
            Find(values, name, out object? value);
            return value;
        }

        private static string Phrase(string label) => label == "All" ? "all its mail" : label + " of mail";

        private static string Show(int? value) => value.HasValue ? value.Value.ToString(System.Globalization.CultureInfo.InvariantCulture) : "absent";
    }
}
