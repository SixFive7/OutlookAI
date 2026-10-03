using System;
using System.Collections.Generic;
using System.Globalization;

namespace OutlookAI.Services
{
    /// <summary>
    /// THE FIVE OUTLOOK CACHED MODE POLICY VALUES (D25), AND THE ONLY VALUES ANYTHING ELEVATED IN
    /// THIS PRODUCT MAY EVER WRITE.
    ///
    /// <para>
    /// They live under <c>HKCU\Software\Policies\Microsoft\Office\&lt;major&gt;\Outlook\Cached Mode</c>,
    /// and <c>HKCU\Software\Policies</c> grants the user ReadKey only - Administrators and SYSTEM hold
    /// FullControl, and a filtered token does not use the Administrators group. So in the Outlook a
    /// user actually runs, NOT elevated, the add-in can read them and cannot write them. Measured on a
    /// guest on 2026-10-03 (Docs/live-tier-on-the-vm.md section 2.3); decided by the maintainer the
    /// same day (Q128): "Allow reading it and changing it from the gui. If the change requires admin
    /// and the user is not admin generate a uac prompt."
    /// </para>
    ///
    /// <para>
    /// WHY ONE FILE. Three things must agree about exactly these five names and exactly these values,
    /// and two of them sit on opposite sides of a privilege boundary:
    ///  - the add-in's tuning catalog takes the names and the shipped defaults from here;
    ///  - the settings dialog offers the choices below and nothing else;
    ///  - the elevated helper, <c>OutlookAI.PolicyWriter.exe</c>, REFUSES any name or value this file
    ///    does not allow. It runs with an administrator's token, so this list is the whole of what a
    ///    UAC approval can be used for.
    /// The add-in and the helper compile this file directly; the test project LINKS it, so the suite
    /// pins the list that ships rather than a copy of it.
    /// </para>
    ///
    /// <para>
    /// THE VALUES are the ones Outlook's own administrative template (outlk16.admx, "Cached Exchange
    /// Mode Sync Settings" and the shared-folder settings) documents: the sync slider in months, the
    /// sync slider in days, and three switches. A closed list on purpose - a number Outlook may read
    /// differently in a later build is something to add here deliberately, not something a request can
    /// smuggle through an elevated process.
    /// </para>
    ///
    /// <para>
    /// FRAMEWORK-NEUTRAL and INTERNAL, under the rules every linked file here follows: it compiles as
    /// net48 (the add-in and the helper, C# 7.3) and as net10 (the test host, nullable-enabled with
    /// warnings as errors), so no nullable annotations, no null returned or assigned, and nothing
    /// newer than C# 7.3.
    /// </para>
    /// </summary>
    internal static class CachedModePolicy
    {
        /// <summary>The subkey under Outlook's key - in the Policies hive here, and in the user hive for D25's two user values.</summary>
        internal const string SubKeyName = "Cached Mode";

        internal const string SyncWindowSetting = "SyncWindowSetting";
        internal const string SyncWindowSettingDays = "SyncWindowSettingDays";
        internal const string DownloadSharedFolders = "DownloadSharedFolders";
        internal const string CacheOthersMail = "CacheOthersMail";
        internal const string DisableSyncSliderForSharedMailbox = "DisableSyncSliderForSharedMailbox";

        /// <summary>One value Outlook accepts for a setting, and what the settings dialog calls it.</summary>
        internal sealed class Choice
        {
            internal Choice(int value, string label)
            {
                Value = value;
                Label = label;
            }

            internal int Value { get; }

            internal string Label { get; }

            /// <summary>What a list shows: the meaning, then the number Outlook stores.</summary>
            public override string ToString()
            {
                return Label + " (" + Value.ToString(CultureInfo.InvariantCulture) + ")";
            }
        }

        private sealed class Setting
        {
            internal Setting(string name, int shippedDefault, string meaning, Choice[] choices)
            {
                Name = name;
                ShippedDefault = shippedDefault;
                Meaning = meaning;
                Choices = choices;
            }

            internal string Name { get; }

            internal int ShippedDefault { get; }

            internal string Meaning { get; }

            internal Choice[] Choices { get; }
        }

        private static Choice[] Switch(string on, string off)
        {
            return new[] { new Choice(1, on), new Choice(0, off) };
        }

        // In the catalog's order, which is the order the reconcile walks them and the dialog lists
        // them. The shipped defaults are D25's: sync everything, cache shared and delegate folders,
        // and keep the slider off shared mailboxes - the configuration the agent's search needs.
        private static readonly Setting[] Settings =
        {
            new Setting(SyncWindowSetting, 0, "Mail to keep offline, in months",
                new[]
                {
                    new Choice(0, "All"),
                    new Choice(1, "1 month"),
                    new Choice(3, "3 months"),
                    new Choice(6, "6 months"),
                    new Choice(12, "12 months"),
                    new Choice(24, "24 months"),
                }),
            new Setting(SyncWindowSettingDays, 0, "Mail to keep offline, in days (overrides months when set)",
                new[]
                {
                    new Choice(0, "Not used"),
                    new Choice(3, "3 days"),
                    new Choice(7, "1 week"),
                    new Choice(14, "2 weeks"),
                }),
            new Setting(DownloadSharedFolders, 1, "Cache shared folders", Switch("On", "Off")),
            new Setting(CacheOthersMail, 1, "Cache other people's mail folders", Switch("On", "Off")),
            new Setting(DisableSyncSliderForSharedMailbox, 1, "Keep shared mailboxes fully cached", Switch("On", "Off")),
        };

        /// <summary>The five value names, in catalog order.</summary>
        internal static IReadOnlyList<string> ValueNames
        {
            get
            {
                var names = new List<string>(Settings.Length);
                foreach (Setting s in Settings)
                    names.Add(s.Name);
                return names;
            }
        }

        /// <summary>
        /// True for one of the five names, spelled exactly. ORDINAL, not case-insensitive: the
        /// registry would accept "syncwindowsetting" as the same value, and that is precisely why a
        /// request spelling it that way is not one this product made.
        /// </summary>
        internal static bool IsKnownName(string name)
        {
            return IndexOf(name) >= 0;
        }

        /// <summary>True when <paramref name="value"/> is one of the values offered for <paramref name="name"/>.</summary>
        internal static bool IsAllowed(string name, int value)
        {
            int i = IndexOf(name);
            if (i < 0)
                return false;
            foreach (Choice c in Settings[i].Choices)
            {
                if (c.Value == value)
                    return true;
            }
            return false;
        }

        /// <summary>The value the product ships as desired. Throws for a name that is not one of the five.</summary>
        internal static int ShippedDefault(string name)
        {
            return Require(name).ShippedDefault;
        }

        /// <summary>What the setting does, in a few words. Throws for a name that is not one of the five.</summary>
        internal static string Meaning(string name)
        {
            return Require(name).Meaning;
        }

        /// <summary>The values offered for one setting, in the order a list shows them. Throws for an unknown name.</summary>
        internal static IReadOnlyList<Choice> Choices(string name)
        {
            return Require(name).Choices;
        }

        /// <summary>
        /// A live or desired number in words: "All (0)", "(not set)" for an absent value, and the bare
        /// number, flagged, for one this product does not offer (a GPO may set anything).
        /// </summary>
        internal static string Describe(string name, int? value)
        {
            if (!value.HasValue)
                return "(not set)";
            int i = IndexOf(name);
            if (i >= 0)
            {
                foreach (Choice c in Settings[i].Choices)
                {
                    if (c.Value == value.Value)
                        return c.ToString();
                }
            }
            return value.Value.ToString(CultureInfo.InvariantCulture) + " (not a value OutlookAI offers)";
        }

        /// <summary>
        /// The key, relative to a user's hive root, that holds the five values for one Office major:
        /// <c>Software\Policies\Microsoft\Office\&lt;major&gt;\Outlook\Cached Mode</c>. Built through
        /// <see cref="OfficeVersions.PolicyOutlookKeyPath"/>, the one spelling of the Policies root the
        /// whole product shares (Tools/Checks/check-pinned-constants.ps1 #13).
        /// </summary>
        internal static string KeyPath(string officeVersion)
        {
            return OfficeVersions.PolicyOutlookKeyPath(officeVersion) + @"\" + SubKeyName;
        }

        /// <summary>True for one of <see cref="OfficeVersions.Supported"/>, spelled exactly.</summary>
        internal static bool IsSupportedOfficeVersion(string version)
        {
            if (string.IsNullOrEmpty(version))
                return false;
            foreach (string v in OfficeVersions.Supported)
            {
                if (string.Equals(v, version, StringComparison.Ordinal))
                    return true;
            }
            return false;
        }

        /// <summary>The position of a name in <see cref="Settings"/>, or -1. Ordinal.</summary>
        private static int IndexOf(string name)
        {
            if (string.IsNullOrEmpty(name))
                return -1;
            for (int i = 0; i < Settings.Length; i++)
            {
                if (string.Equals(Settings[i].Name, name, StringComparison.Ordinal))
                    return i;
            }
            return -1;
        }

        private static Setting Require(string name)
        {
            int i = IndexOf(name);
            if (i < 0)
                throw new ArgumentException("'" + name + "' is not one of the five Cached Mode policy values.", nameof(name));
            return Settings[i];
        }
    }
}
