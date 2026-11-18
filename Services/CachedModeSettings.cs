using System;
using System.Collections.Generic;
using System.Globalization;

namespace OutlookAI.Services
{
    /// <summary>
    /// OUTLOOK'S CACHED MODE SETTINGS THAT OUTLOOKAI KEEPS APPLIED (D25): WHERE EACH ONE IS READ FROM,
    /// WHO SETS IT, AND THE VALUES OUTLOOKAI WRITES - the one list the add-in's tuning catalog, its
    /// settings dialog and the test suite take them from.
    ///
    /// <para>
    /// THE USER'S OWN SETTINGS, AND NOTHING ELSE - Q145 (a), decided by the maintainer: OutlookAI writes
    /// them under <c>HKCU\Software\Microsoft\Office\&lt;major&gt;\Outlook\Cached Mode</c>, which the user
    /// may write, so nothing ever needs an administrator. Until then - 3.0.0.319 to 3.1.0.325, and the
    /// unreleased Q128 builds after them - five of them were written as Outlook POLICY, under
    /// <c>HKCU\Software\Policies</c>, which Windows lets a user read and not change: Q128 added a
    /// UAC-elevated helper to write them, and Q145 took both out again. The policy key is the
    /// organization's; a value there locks Outlook's own slider; and Microsoft 365 Cloud Policy
    /// outranks it, so a value written there could be in the registry and still not in effect.
    /// </para>
    ///
    /// <para>
    /// TWO POLICY KEYS ARE READ, AND NEVER WRITTEN, to say who sets a value. Outlook takes each value
    /// from the first of these that holds it - "Read the policy value first, and then if not found,
    /// read the preference. (Policy always takes priority)" (bb742499), and Cloud Policy "take[s]
    /// precedence over ... Group Policy" (the Cloud Policy overview):
    /// </para>
    /// <list type="number">
    ///   <item>Microsoft 365 Cloud Policy, <see cref="CloudPolicyKeyPath"/>;</item>
    ///   <item>Group Policy, <see cref="GroupPolicyKeyPath"/>;</item>
    ///   <item>the user's own setting, <see cref="UserKeyPath"/> - the one OutlookAI writes.</item>
    /// </list>
    /// <para>
    /// A value under either policy key is shown as "Set by your organization" and left alone - except
    /// a Group Policy value an earlier OutlookAI wrote, recognised by the same number in
    /// <c>Tuning\Applied\caching.policy.&lt;name&gt;</c>, which is shown as "Set by an earlier
    /// OutlookAI" and also left alone (Q151 (a): those values stay until the maintainer asks for them
    /// to go; <see cref="EarlierPolicyRemoval"/> is the way they go, and nothing runs it yet). The
    /// number cannot prove who wrote it: an organization's policy holding the very number the earlier
    /// OutlookAI applied reads the same, and on a machine where any process of the user can write the
    /// policy key - the maintainer's own, measured 2026-10-11 - "under Policies" does not prove an
    /// administrator either.
    /// </para>
    ///
    /// <para>
    /// WHAT THE USER'S OWN SETTINGS REACH, which the dialog says beside each one:
    /// </para>
    /// <list type="bullet">
    ///   <item>the offline period (<see cref="SyncWindowSetting"/> with <see cref="SyncWindowSettingDays"/>):
    ///   accounts added from now on (KB 3115009). An account that already exists keeps its own period
    ///   in the Outlook profile, which Account Settings changes - <see cref="ExistingAccountPeriodSteps"/>
    ///   says how, and writing it for the user is <see cref="IAccountOfflinePeriods"/>'s, not built yet;</item>
    ///   <item><see cref="DownloadSharedFolders"/>: new profiles (inferred - an existing account has its
    ///   own "Download shared folders" box);</item>
    ///   <item><see cref="CacheOthersMail"/>: new profiles (KB 982697: "If the registry value exists in
    ///   the non-Policies key, it only applies to new Outlook profiles");</item>
    ///   <item><see cref="DisableSyncSliderForSharedMailbox"/>: every shared mailbox, existing ones
    ///   included (KB 3140747, Method 2). It is registry-only - not in the administrative template.</item>
    /// </list>
    ///
    /// <para>
    /// THE OFFLINE PERIOD IS ONE SETTING WITH ELEVEN VALUES - Q143, decided by the maintainer: exactly
    /// the eleven windows Microsoft documents (KB 3115009's table, KB 2733062's policy values, and the
    /// eleven items of the administrative template's "Cached Exchange Mode Sync Settings"), each one
    /// written as its (months, days) pair, and nothing else. KB 3115009: a value that is missing or
    /// invalid makes Outlook use its default of 12 months - so a stray 18 would silently be a year.
    /// No source documents a period that combines months with days, or any other number of months.
    /// </para>
    ///
    /// <para>
    /// FRAMEWORK-NEUTRAL and INTERNAL, under the rules every linked file here follows: the add-in
    /// compiles it (net48, C# 15) and the test project LINKS it (net11, nullable-enabled with
    /// warnings as errors), so no nullable annotations, no null returned or assigned, and
    /// no language feature .NET Framework 4.8 cannot run (Properties\LanguageFeatureStandIns.cs). It reads and writes through <see cref="ITuningStore"/>, never the registry.
    /// </para>
    /// </summary>
    internal static class CachedModeSettings
    {
        /// <summary>The subkey under Outlook's key, in all three roots.</summary>
        internal const string SubKeyName = "Cached Mode";

        internal const string SyncWindowSetting = "SyncWindowSetting";
        internal const string SyncWindowSettingDays = "SyncWindowSettingDays";
        internal const string DownloadSharedFolders = "DownloadSharedFolders";
        internal const string CacheOthersMail = "CacheOthersMail";
        internal const string DisableSyncSliderForSharedMailbox = "DisableSyncSliderForSharedMailbox";

        /// <summary>The id prefix of the catalog's five entries; the rest of each id is the value name.</summary>
        internal const string EntryIdPrefix = "caching.user.";

        /// <summary>
        /// The id prefix under which 3.0.0.319 to 3.1.0.325, and the Q128 builds after them, kept the
        /// five POLICY values they wrote: <c>Tuning\Desired\caching.policy.&lt;name&gt;</c> and
        /// <c>Tuning\Applied\caching.policy.&lt;name&gt;</c>. Nothing writes these any more; the Applied
        /// ones are how a policy value an earlier OutlookAI wrote is told from an organization's.
        /// </summary>
        internal const string EarlierPolicyEntryIdPrefix = "caching.policy.";

        /// <summary>The setting key of the offline period, the one setting written as two values.</summary>
        internal const string OfflinePeriodKey = "OfflinePeriod";

        // ===== The states the dialog shows, one per setting =====

        internal const string StateSetByOrganization = "Set by your organization";
        internal const string StateSetByEarlierOutlookAI = "Set by an earlier OutlookAI";
        internal const string StateNotManaged = "Not managed";
        internal const string StateNotOffered = "Not a value OutlookAI writes";
        internal const string StateInEffect = "In effect";
        internal const string StateRefused = "Windows refused the change";
        internal const string StateNotApplied = "Not applied yet";

        /// <summary>
        /// How an EXISTING account's offline period is changed, in Outlook's own words - KB 2733062's
        /// steps, with the setting's newer name first: "later versions of Outlook renamed the Mail to
        /// keep offline setting to Download email for the past". Q145 (a): until OutlookAI can write
        /// an account's own period (<see cref="IAccountOfflinePeriods"/>), it tells the user how.
        /// </summary>
        internal const string ExistingAccountPeriodSteps =
            "An account you already have keeps its own offline period: the one above applies to accounts "
            + "you add later. To keep all of an existing account's mail offline, where OutlookAI can search "
            + "it: in Outlook, choose File > Account Settings > Account Settings, double-click the account on "
            + "the Email tab, drag \"Download email for the past\" (\"Mail to keep offline\" in older "
            + "versions) to All, choose Next, and restart Outlook.";

        internal const string ExistingAccountPeriodSetByOrganization =
            "Your organization sets how much mail Outlook keeps offline, for every account, so Outlook's "
            + "Account Settings shows that setting greyed out.";

        internal const string ExistingAccountPeriodSetByEarlierOutlookAI =
            "An earlier OutlookAI set how much mail Outlook keeps offline as a policy, for every account, "
            + "so Outlook's Account Settings shows that setting greyed out. OutlookAI leaves it in place.";

        // ===== The eleven offline periods (Q143) live in Services\OfflinePeriods.cs, which Core links too. =====

        private static readonly SwitchChoice[] OnOff = { new SwitchChoice(1, "On"), new SwitchChoice(0, "Off") };

        private static readonly CachedModeSetting[] Settings =
        {
            new CachedModeSetting(OfflinePeriodKey, "Offline period for new accounts",
                "SyncWindowSetting and SyncWindowSettingDays: how much of a mailbox Outlook keeps offline in an "
                + "account added from now on. Older mail stays on the server only - not in the offline file and "
                + "not in the Windows Search index - so OutlookAI's search does not find it.",
                new[] { SyncWindowSetting, SyncWindowSettingDays }),
            new CachedModeSetting(DownloadSharedFolders, "Cache shared folders in new profiles",
                "DownloadSharedFolders: Outlook keeps shared and delegate folders offline, where they can be "
                + "searched. An existing account has its own \"Download shared folders\" box, under More "
                + "Settings > Advanced in its Account Settings.",
                new[] { DownloadSharedFolders }),
            new CachedModeSetting(CacheOthersMail, "Cache other people's mail in new profiles",
                "CacheOthersMail: of the shared folders Outlook keeps offline, the mail folders too - not only "
                + "calendars, contacts and tasks.",
                new[] { CacheOthersMail }),
            new CachedModeSetting(DisableSyncSliderForSharedMailbox, "Keep shared mailboxes fully cached",
                "DisableSyncSliderForSharedMailbox: shared mailboxes and public-folder favourites keep all their "
                + "mail offline, whatever the offline period.",
                new[] { DisableSyncSliderForSharedMailbox }),
        };

        /// <summary>The five value names, in catalog order.</summary>
        internal static IReadOnlyList<string> ValueNames
        {
            get
            {
                return new[] { SyncWindowSetting, SyncWindowSettingDays, DownloadSharedFolders, CacheOthersMail, DisableSyncSliderForSharedMailbox };
            }
        }

        /// <summary>The settings the dialog lists, one row each, in this order.</summary>
        internal static IReadOnlyList<CachedModeSetting> AllSettings
        {
            get { return Array.AsReadOnly(Settings); }
        }

        /// <summary>The eleven periods, shortest first, All last (<see cref="OfflinePeriods"/>).</summary>
        internal static IReadOnlyList<OfflinePeriod> AllPeriods
        {
            get
            {
                var periods = new List<OfflinePeriod>(OfflinePeriods.Count);
                for (int i = 0; i < OfflinePeriods.Count; i++)
                    periods.Add(OfflinePeriods.At(i));
                return periods;
            }
        }

        /// <summary>On, then Off: the values a switch offers.</summary>
        internal static IReadOnlyList<SwitchChoice> SwitchChoices
        {
            get { return Array.AsReadOnly(OnOff); }
        }

        /// <summary>The position of (months, days) among the eleven periods, or -1 for a pair Outlook does not document.</summary>
        internal static int IndexOfPeriod(int months, int days)
        {
            return OfflinePeriods.IndexOf(months, days);
        }

        /// <summary>True for exactly the eleven (months, days) pairs Microsoft documents (Q143).</summary>
        internal static bool IsOfferedPeriod(int months, int days)
        {
            return OfflinePeriods.IsOffered(months, days);
        }

        /// <summary>True for 1 and 0, the two values a switch offers.</summary>
        internal static bool IsOfferedSwitch(int value)
        {
            return value == 1 || value == 0;
        }

        /// <summary>The setting with this key, spelled exactly; throws for any other key.</summary>
        internal static CachedModeSetting Setting(string key)
        {
            foreach (CachedModeSetting s in Settings)
            {
                if (string.Equals(s.Key, key, StringComparison.Ordinal))
                    return s;
            }
            throw new ArgumentException("'" + key + "' is not one of the Cached Mode settings.", nameof(key));
        }

        /// <summary>The labels a setting's list shows, in order: the eleven periods, or On and Off.</summary>
        internal static List<string> ChoiceLabels(CachedModeSetting setting)
        {
            var labels = new List<string>();
            if (setting.IsPeriod)
            {
                for (int i = 0; i < OfflinePeriods.Count; i++)
                    labels.Add(OfflinePeriods.At(i).Label);
            }
            else
            {
                foreach (SwitchChoice c in OnOff)
                    labels.Add(c.Label);
            }
            return labels;
        }

        /// <summary>
        /// The desired values one choice stands for, by catalog entry id: the period's two halves, or
        /// the one switch. Empty for an index the setting does not offer - the one way a caller can
        /// ask for anything else, and refused.
        /// </summary>
        internal static List<KeyValuePair<string, int>> DesiredFor(CachedModeSetting setting, int choiceIndex)
        {
            var values = new List<KeyValuePair<string, int>>();
            if (setting.IsPeriod)
            {
                if (choiceIndex < 0 || choiceIndex >= OfflinePeriods.Count)
                    return values;
                values.Add(new KeyValuePair<string, int>(EntryIdPrefix + SyncWindowSetting, OfflinePeriods.At(choiceIndex).Months));
                values.Add(new KeyValuePair<string, int>(EntryIdPrefix + SyncWindowSettingDays, OfflinePeriods.At(choiceIndex).Days));
                return values;
            }
            if (choiceIndex < 0 || choiceIndex >= OnOff.Length)
                return values;
            values.Add(new KeyValuePair<string, int>(EntryIdPrefix + setting.Key, OnOff[choiceIndex].Value));
            return values;
        }

        // ===== Where they live =====

        /// <summary>The user's own Cached Mode key - the one OutlookAI writes.</summary>
        internal static string UserKeyPath(string officeVersion)
        {
            return OfficeVersions.OutlookKeyPath(officeVersion) + @"\" + SubKeyName;
        }

        /// <summary>Group Policy's Cached Mode key: read, never written.</summary>
        internal static string GroupPolicyKeyPath(string officeVersion)
        {
            return OfficeVersions.PolicyOutlookKeyPath(officeVersion) + @"\" + SubKeyName;
        }

        /// <summary>Microsoft 365 Cloud Policy's Cached Mode key: read, never written. It outranks Group Policy.</summary>
        internal static string CloudPolicyKeyPath(string officeVersion)
        {
            return OfficeVersions.CloudPolicyOutlookKeyPath(officeVersion) + @"\" + SubKeyName;
        }

        // ===== The tuning catalog's five entries =====

        /// <summary>
        /// The caching group of the add-in's tuning catalog: the five values, each written under the
        /// user's own key and left alone while either policy key sets it - the period as a whole, when
        /// a policy sets either half - and written only while its desired number is one offered.
        /// Shipped defaults: All, and the three switches On.
        /// </summary>
        internal static List<TuningEntry> CatalogEntries(string officeVersion)
        {
            string userKey = UserKeyPath(officeVersion);
            string[] policyKeys = { CloudPolicyKeyPath(officeVersion), GroupPolicyKeyPath(officeVersion) };
            string[] periodNames = { SyncWindowSetting, SyncWindowSettingDays };
            Func<ITuningStore, bool> periodOffered = IsDesiredPeriodOffered;

            var entries = new List<TuningEntry>
            {
                new TuningEntry(EntryIdPrefix + SyncWindowSetting, TuningReconciler.GroupCaching, userKey,
                                SyncWindowSetting, OfflinePeriods.All.Months, policyKeys, periodNames, periodOffered),
                new TuningEntry(EntryIdPrefix + SyncWindowSettingDays, TuningReconciler.GroupCaching, userKey,
                                SyncWindowSettingDays, OfflinePeriods.All.Days, policyKeys, periodNames, periodOffered),
            };
            foreach (string name in new[] { DownloadSharedFolders, CacheOthersMail, DisableSyncSliderForSharedMailbox })
            {
                string id = EntryIdPrefix + name;
                entries.Add(new TuningEntry(id, TuningReconciler.GroupCaching, userKey, name, 1, policyKeys, new[] { name },
                                            store => IsOfferedSwitch(DesiredOf(store, id, 1))));
            }
            return entries;
        }

        /// <summary>True when the two desired halves of the period, read together, are one of the eleven.</summary>
        internal static bool IsDesiredPeriodOffered(ITuningStore store)
        {
            return IsOfferedPeriod(DesiredOf(store, EntryIdPrefix + SyncWindowSetting, OfflinePeriods.All.Months),
                                   DesiredOf(store, EntryIdPrefix + SyncWindowSettingDays, OfflinePeriods.All.Days));
        }

        // As TuningReconciler.ReadDesired reads it: the stored desired number, else the shipped default.
        private static int DesiredOf(ITuningStore store, string id, int shippedDefault)
        {
            int? stored = store.ReadDword(TuningReconciler.DesiredKeyPath, id);
            return stored.HasValue ? stored.Value : shippedDefault;
        }

        // ===== Who sets each value =====

        /// <summary>One value as the three keys hold it, and what an earlier OutlookAI recorded applying.</summary>
        internal static CachedModeValueReading Read(ITuningStore store, string officeVersion, string valueName)
        {
            return new CachedModeValueReading(
                valueName,
                store.ReadDword(CloudPolicyKeyPath(officeVersion), valueName),
                store.ReadDword(GroupPolicyKeyPath(officeVersion), valueName),
                store.ReadDword(UserKeyPath(officeVersion), valueName),
                store.ReadDword(TuningReconciler.AppliedKeyPath, EarlierPolicyEntryIdPrefix + valueName));
        }

        /// <summary>
        /// Every setting as the dialog shows it. <paramref name="managed"/> is whether the caching
        /// group (and the master switch) is on; <paramref name="refusedIds"/> the entry ids Windows
        /// refused the last reconcile (Tuning\NeedsAdministrator).
        /// </summary>
        internal static List<CachedModeSettingState> ReadStates(ITuningStore store, string officeVersion, bool managed, ICollection<string> refusedIds)
        {
            var states = new List<CachedModeSettingState>();
            foreach (CachedModeSetting setting in Settings)
            {
                var readings = new List<CachedModeValueReading>();
                foreach (string name in setting.ValueNames)
                    readings.Add(Read(store, officeVersion, name));
                states.Add(StateOf(store, setting, readings, managed, refusedIds));
            }
            return states;
        }

        private static CachedModeSettingState StateOf(ITuningStore store, CachedModeSetting setting, List<CachedModeValueReading> readings,
                                                      bool managed, ICollection<string> refusedIds)
        {
            // A pair is set by whoever sets either half, the stronger of the two winning.
            CachedModeSource source = CachedModeSource.NotSet;
            foreach (CachedModeValueReading r in readings)
            {
                if (r.Source > source)
                    source = r.Source;
            }

            string inEffect;
            int desiredIndex;
            bool userMatchesDesired;
            bool refused = false;
            if (setting.IsPeriod)
            {
                inEffect = DescribePeriod(readings[0].Effective, readings[1].Effective);
                int months = DesiredOf(store, EntryIdPrefix + SyncWindowSetting, OfflinePeriods.All.Months);
                int days = DesiredOf(store, EntryIdPrefix + SyncWindowSettingDays, OfflinePeriods.All.Days);
                desiredIndex = IndexOfPeriod(months, days);
                userMatchesDesired = readings[0].User == months && readings[1].User == days;
                refused = refusedIds.Contains(EntryIdPrefix + SyncWindowSetting) || refusedIds.Contains(EntryIdPrefix + SyncWindowSettingDays);
            }
            else
            {
                inEffect = DescribeSwitch(readings[0].Effective);
                int desired = DesiredOf(store, EntryIdPrefix + setting.Key, 1);
                desiredIndex = -1;
                for (int i = 0; i < OnOff.Length; i++)
                {
                    if (OnOff[i].Value == desired)
                        desiredIndex = i;
                }
                userMatchesDesired = readings[0].User == desired;
                refused = refusedIds.Contains(EntryIdPrefix + setting.Key);
            }

            return new CachedModeSettingState(setting, source, inEffect, desiredIndex, managed,
                                              DescribeState(source, managed, desiredIndex >= 0, userMatchesDesired, refused));
        }

        /// <summary>
        /// The State column, decided in this order: who sets it outranks whether OutlookAI manages it,
        /// because a policy holds whether OutlookAI manages the value or not; then whether OutlookAI's
        /// desired value is one it writes; then whether it is in effect.
        /// </summary>
        internal static string DescribeState(CachedModeSource source, bool managed, bool desiredOffered, bool userMatchesDesired, bool refused)
        {
            if (source == CachedModeSource.CloudPolicy || source == CachedModeSource.GroupPolicy)
                return StateSetByOrganization;
            if (source == CachedModeSource.EarlierOutlookAI)
                return StateSetByEarlierOutlookAI;
            if (!managed)
                return StateNotManaged;
            if (!desiredOffered)
                return StateNotOffered;
            if (userMatchesDesired)
                return StateInEffect;
            if (refused)
                return StateRefused;
            return StateNotApplied;
        }

        /// <summary>What the dialog says under the settings about an account that already exists.</summary>
        internal static string ExistingAccountAdvice(CachedModeSource periodSource)
        {
            switch (periodSource)
            {
                case CachedModeSource.CloudPolicy:
                case CachedModeSource.GroupPolicy:
                    return ExistingAccountPeriodSetByOrganization;
                case CachedModeSource.EarlierOutlookAI:
                    return ExistingAccountPeriodSetByEarlierOutlookAI;
                default:
                    return ExistingAccountPeriodSteps;
            }
        }

        /// <summary>
        /// A period in words: the name of one of the eleven, "(not set)" when neither half is, and
        /// the bare numbers, flagged, for anything else - a policy may hold anything. A months value
        /// with no days value beside it reads as months alone, as Outlook reads it (KB 3115009: a
        /// missing days value is 0) and as Outlook 2013's template writes it.
        /// </summary>
        internal static string DescribePeriod(int? months, int? days)
        {
            if (!months.HasValue && !days.HasValue)
                return "(not set)";
            if (months.HasValue)
            {
                int i = IndexOfPeriod(months.Value, days.HasValue ? days.Value : 0);
                if (i >= 0)
                    return OfflinePeriods.At(i).Label;
            }
            return "months " + Number(months) + ", days " + Number(days) + " (not one of Outlook's periods)";
        }

        /// <summary>A switch in words: On, Off, "(not set)", or the bare number, flagged.</summary>
        internal static string DescribeSwitch(int? value)
        {
            if (!value.HasValue)
                return "(not set)";
            foreach (SwitchChoice c in OnOff)
            {
                if (c.Value == value.Value)
                    return c.Label;
            }
            return Number(value) + " (neither On nor Off)";
        }

        private static string Number(int? value)
        {
            return value.HasValue ? value.Value.ToString(CultureInfo.InvariantCulture) : "not set";
        }
    }

    /// <summary>One value a switch offers: 1 or 0, and its name.</summary>
    internal sealed class SwitchChoice
    {
        internal SwitchChoice(int value, string label)
        {
            Value = value;
            Label = label;
        }

        internal int Value { get; }

        internal string Label { get; }

        public override string ToString()
        {
            return Label;
        }
    }

    /// <summary>One row of the dialog: a switch, or the offline period's pair.</summary>
    internal sealed class CachedModeSetting
    {
        internal CachedModeSetting(string key, string caption, string detail, string[] valueNames)
        {
            Key = key;
            Caption = caption;
            Detail = detail;
            ValueNames = Array.AsReadOnly(valueNames);
        }

        /// <summary><see cref="CachedModeSettings.OfflinePeriodKey"/>, or the switch's value name.</summary>
        internal string Key { get; }

        internal string Caption { get; }

        /// <summary>What it does, and which value(s) it is - the row's tooltip.</summary>
        internal string Detail { get; }

        /// <summary>The Outlook value or values it is written as.</summary>
        internal IReadOnlyList<string> ValueNames { get; }

        internal bool IsPeriod
        {
            get { return ValueNames.Count == 2; }
        }
    }

    /// <summary>
    /// Who sets a value, weakest first: a pair's source is the stronger of its halves', and the
    /// numbers order them. Outlook reads the strongest one present.
    /// </summary>
    internal enum CachedModeSource
    {
        /// <summary>Nothing sets it: Outlook's own default applies.</summary>
        NotSet = 0,

        /// <summary>The user's own setting - the one OutlookAI writes.</summary>
        UserSetting = 1,

        /// <summary>A Group Policy value that Tuning\Applied says an earlier OutlookAI wrote (Q151 (a)).</summary>
        EarlierOutlookAI = 2,

        /// <summary>A Group Policy value OutlookAI did not write: the organization's.</summary>
        GroupPolicy = 3,

        /// <summary>Microsoft 365 Cloud Policy: the organization's, and above Group Policy.</summary>
        CloudPolicy = 4,
    }

    /// <summary>One value in the three keys Outlook reads it from, and the earlier OutlookAI's record of it.</summary>
    internal sealed class CachedModeValueReading
    {
        internal CachedModeValueReading(string valueName, int? cloudPolicy, int? groupPolicy, int? user, int? earlierOutlookAIApplied)
        {
            ValueName = valueName;
            CloudPolicy = cloudPolicy;
            GroupPolicy = groupPolicy;
            User = user;
            EarlierOutlookAIApplied = earlierOutlookAIApplied;
        }

        internal string ValueName { get; }

        internal int? CloudPolicy { get; }

        internal int? GroupPolicy { get; }

        internal int? User { get; }

        /// <summary>Tuning\Applied\caching.policy.&lt;name&gt;: what an earlier OutlookAI applied, if it recorded one.</summary>
        internal int? EarlierOutlookAIApplied { get; }

        /// <summary>What Outlook uses: Cloud Policy, else Group Policy, else the user's own setting.</summary>
        internal int? Effective
        {
            get { return CloudPolicy.HasValue ? CloudPolicy : (GroupPolicy.HasValue ? GroupPolicy : User); }
        }

        /// <summary>The Group Policy value is the one an earlier OutlookAI recorded applying - whether or not Cloud Policy outranks it.</summary>
        internal bool GroupPolicyIsEarlierOutlookAI
        {
            get { return GroupPolicy.HasValue && EarlierOutlookAIApplied.HasValue && GroupPolicy.Value == EarlierOutlookAIApplied.Value; }
        }

        /// <summary>Who sets what Outlook uses.</summary>
        internal CachedModeSource Source
        {
            get
            {
                if (CloudPolicy.HasValue)
                    return CachedModeSource.CloudPolicy;
                if (GroupPolicy.HasValue)
                    return GroupPolicyIsEarlierOutlookAI ? CachedModeSource.EarlierOutlookAI : CachedModeSource.GroupPolicy;
                if (User.HasValue)
                    return CachedModeSource.UserSetting;
                return CachedModeSource.NotSet;
            }
        }
    }

    /// <summary>One row of the dialog, read.</summary>
    internal sealed class CachedModeSettingState
    {
        internal CachedModeSettingState(CachedModeSetting setting, CachedModeSource source, string inEffect, int desiredIndex, bool managed, string state)
        {
            Setting = setting;
            Source = source;
            InEffect = inEffect;
            DesiredIndex = desiredIndex;
            Managed = managed;
            State = state;
        }

        internal CachedModeSetting Setting { get; }

        /// <summary>Who sets what Outlook uses.</summary>
        internal CachedModeSource Source { get; }

        /// <summary>What Outlook uses, in words.</summary>
        internal string InEffect { get; }

        /// <summary>OutlookAI's desired value, as a position in the setting's list; -1 when it is not one the list offers.</summary>
        internal int DesiredIndex { get; }

        /// <summary>The caching group and the master switch are on.</summary>
        internal bool Managed { get; }

        internal string State { get; }

        /// <summary>A policy - the organization's, or an earlier OutlookAI's - sets it: shown, never changed here.</summary>
        internal bool SetByPolicy
        {
            get { return Source >= CachedModeSource.EarlierOutlookAI; }
        }

        /// <summary>The user may pick another value: OutlookAI manages it and no policy sets it.</summary>
        internal bool Editable
        {
            get { return Managed && !SetByPolicy; }
        }
    }
}
