using System;
using System.Collections.Generic;
using Microsoft.Win32;

namespace OutlookAI.Services
{
    /// <summary>
    /// Keeps the user's proven Outlook tuning applied (R12 / v3 plan section 0.5.3): local
    /// search behavior, full Cached-Mode sync (slider = All), and OST size headroom.
    ///
    /// Desired state, per-group toggles, and reconcile bookkeeping live under
    /// HKCU\Software\OutlookAI\Tuning. On every Outlook start (and on demand from the
    /// settings dialog) the live registry is reconciled against desired state:
    ///  - idempotent — only actual diffs are written;
    ///  - GPO-respecting — a HKCU\Software\Policies\... value that reverts after we applied
    ///    it is never fought: we back off and flag the value so the settings dialog can
    ///    surface "managed by policy";
    ///  - restart-aware — writes take effect on the NEXT Outlook start, so any write sets a
    ///    persisted "restart needed" flag; the flag clears on the first startup reconcile
    ///    that finds everything already in sync (i.e. Outlook booted with the values live).
    ///
    /// Everything is HKCU — no elevation is ever required. Disabling a group (or the master
    /// switch) stops managing its values; already-written Outlook values are left in place,
    /// mirroring the uninstall behavior described in the plan.
    /// </summary>
    internal static class OutlookTuningService
    {
        internal const string TuningKeyPath = @"Software\OutlookAI\Tuning";
        private const string DesiredKeyPath = TuningKeyPath + @"\Desired";
        private const string AppliedKeyPath = TuningKeyPath + @"\Applied";

        private const string SearchKeyPath = @"Software\Microsoft\Office\16.0\Outlook\Search";
        private const string CachedModePolicyKeyPath = @"Software\Policies\Microsoft\Office\16.0\Outlook\Cached Mode";
        private const string CachedModeUserKeyPath = @"Software\Microsoft\Office\16.0\Outlook\Cached Mode";
        private const string PstKeyPath = @"Software\Microsoft\Office\16.0\Outlook\PST";

        internal const string GroupSearch = "search";
        internal const string GroupCaching = "caching";
        internal const string GroupOst = "ost";

        private static readonly object _gate = new object();

        internal sealed class TuningEntry
        {
            public string Id { get; }
            public string GroupId { get; }
            public string KeyPath { get; }
            public string ValueName { get; }
            public int DefaultDesired { get; }
            public bool IsPolicyHive { get; }

            public TuningEntry(string id, string groupId, string keyPath, string valueName, int defaultDesired, bool isPolicyHive)
            {
                Id = id;
                GroupId = groupId;
                KeyPath = keyPath;
                ValueName = valueName;
                DefaultDesired = defaultDesired;
                IsPolicyHive = isPolicyHive;
            }
        }

        internal sealed class ValueState
        {
            public TuningEntry Entry { get; internal set; }
            public int Desired { get; internal set; }
            public int? Live { get; internal set; }
            public bool InSync { get; internal set; }
            public bool BackedOff { get; internal set; }
            public bool GroupEnabled { get; internal set; }
        }

        internal sealed class TuningSnapshot
        {
            public bool MasterEnabled { get; internal set; }
            public bool SearchEnabled { get; internal set; }
            public bool CachingEnabled { get; internal set; }
            public bool OstEnabled { get; internal set; }
            public bool RestartNeeded { get; internal set; }
            public List<ValueState> Values { get; internal set; }
            public List<string> PolicyConflicts { get; internal set; }
        }

        internal sealed class ReconcileResult
        {
            public bool WroteAny { get; internal set; }
            public bool RestartNeeded { get; internal set; }
            public List<string> PolicyConflicts { get; internal set; }
        }

        // The full desired-state catalog (defaults per v3 plan D22/D24/D25). The registry
        // stores the desired NUMBERS (self-healing, future-tunable); this catalog is the
        // authoritative structure: which value lives in which key, and its default.
        private static readonly TuningEntry[] Catalog = new[]
        {
            // Search (D22) — the user's proven local-search setup.
            new TuningEntry("search.DisableServerAssistedSearch", GroupSearch, SearchKeyPath, "DisableServerAssistedSearch", 1, false),
            new TuningEntry("search.SearchResultsCap",            GroupSearch, SearchKeyPath, "SearchResultsCap",            0, false),
            new TuningEntry("search.IncludeDeletedItems",         GroupSearch, SearchKeyPath, "IncludeDeletedItems",         1, false),
            new TuningEntry("search.DefaultSearchScope",          GroupSearch, SearchKeyPath, "DefaultSearchScope",          2, false),

            // Full caching (D25) — sync slider = All for existing accounts (Policies hive)
            // and future accounts (user hive), plus shared-folder caching.
            new TuningEntry("caching.policy.SyncWindowSetting",                  GroupCaching, CachedModePolicyKeyPath, "SyncWindowSetting",                  0, true),
            new TuningEntry("caching.policy.SyncWindowSettingDays",              GroupCaching, CachedModePolicyKeyPath, "SyncWindowSettingDays",              0, true),
            new TuningEntry("caching.policy.DownloadSharedFolders",              GroupCaching, CachedModePolicyKeyPath, "DownloadSharedFolders",              1, true),
            new TuningEntry("caching.policy.CacheOthersMail",                    GroupCaching, CachedModePolicyKeyPath, "CacheOthersMail",                    1, true),
            new TuningEntry("caching.policy.DisableSyncSliderForSharedMailbox",  GroupCaching, CachedModePolicyKeyPath, "DisableSyncSliderForSharedMailbox",  1, true),
            new TuningEntry("caching.user.SyncWindowSetting",                    GroupCaching, CachedModeUserKeyPath,   "SyncWindowSetting",                  0, false),
            new TuningEntry("caching.user.SyncWindowSettingDays",                GroupCaching, CachedModeUserKeyPath,   "SyncWindowSettingDays",              0, false),

            // OST headroom (D25) — 100 GB max / ~94 GB warn so full caching never stalls at
            // the default 50 GB cap.
            new TuningEntry("ost.MaxLargeFileSize",  GroupOst, PstKeyPath, "MaxLargeFileSize",  102400, false),
            new TuningEntry("ost.WarnLargeFileSize", GroupOst, PstKeyPath, "WarnLargeFileSize", 96256,  false),
        };

        internal static IReadOnlyList<TuningEntry> Entries
        {
            get { return Catalog; }
        }

        // ===== Public operations =====

        /// <summary>Startup reconcile: applies diffs and maintains the restart-needed flag
        /// (clears it when Outlook booted with everything already in sync).</summary>
        public static ReconcileResult ReconcileOnStartup()
        {
            return Reconcile(true);
        }

        /// <summary>Mid-session reconcile (settings dialog): applies diffs; only ever SETS the
        /// restart-needed flag — a mid-session "everything matches" must not clear it because
        /// the running Outlook may still be on pre-change values.</summary>
        public static ReconcileResult ReconcileFromUi()
        {
            return Reconcile(false);
        }

        /// <summary>Read-only view of desired vs live state for the settings dialog.</summary>
        public static TuningSnapshot GetSnapshot()
        {
            lock (_gate)
            {
                try
                {
                    EnsureInitialized();
                    var conflicts = GetPolicyConflictsInternal();
                    var snapshot = new TuningSnapshot
                    {
                        MasterEnabled = GetToggleInternal("Enabled"),
                        SearchEnabled = GetToggleInternal("SearchEnabled"),
                        CachingEnabled = GetToggleInternal("CachingEnabled"),
                        OstEnabled = GetToggleInternal("OstEnabled"),
                        RestartNeeded = ReadDword(TuningKeyPath, "RestartNeeded") == 1,
                        Values = new List<ValueState>(),
                        PolicyConflicts = conflicts,
                    };
                    foreach (var entry in Catalog)
                    {
                        int desired = ReadDesired(entry);
                        int? live = ReadDword(entry.KeyPath, entry.ValueName);
                        snapshot.Values.Add(new ValueState
                        {
                            Entry = entry,
                            Desired = desired,
                            Live = live,
                            InSync = live.HasValue && live.Value == desired,
                            BackedOff = conflicts.Contains(entry.Id),
                            GroupEnabled = snapshot.MasterEnabled && GetToggleInternal(ToggleName(entry.GroupId)),
                        });
                    }
                    return snapshot;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Tuning snapshot failed: " + ex.Message);
                    return new TuningSnapshot
                    {
                        Values = new List<ValueState>(),
                        PolicyConflicts = new List<string>(),
                    };
                }
            }
        }

        public static bool GetMasterEnabled()
        {
            lock (_gate) { try { EnsureInitialized(); return GetToggleInternal("Enabled"); } catch { return true; } }
        }

        public static void SetMasterEnabled(bool enabled)
        {
            lock (_gate) { try { EnsureInitialized(); WriteDword(TuningKeyPath, "Enabled", enabled ? 1 : 0); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine("SetMasterEnabled: " + ex.Message); } }
        }

        public static bool GetGroupEnabled(string groupId)
        {
            lock (_gate) { try { EnsureInitialized(); return GetToggleInternal(ToggleName(groupId)); } catch { return true; } }
        }

        public static void SetGroupEnabled(string groupId, bool enabled)
        {
            lock (_gate) { try { EnsureInitialized(); WriteDword(TuningKeyPath, ToggleName(groupId), enabled ? 1 : 0); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine("SetGroupEnabled: " + ex.Message); } }
        }

        public static bool GetRestartNeeded()
        {
            lock (_gate) { try { return ReadDword(TuningKeyPath, "RestartNeeded") == 1; } catch { return false; } }
        }

        // ===== Core reconcile =====

        private static ReconcileResult Reconcile(bool isStartup)
        {
            lock (_gate)
            {
                var result = new ReconcileResult { PolicyConflicts = new List<string>() };
                try
                {
                    EnsureInitialized();

                    bool masterOn = GetToggleInternal("Enabled");
                    bool wroteAny = false;
                    var conflicts = new List<string>();

                    foreach (var entry in Catalog)
                    {
                        bool groupOn = masterOn && GetToggleInternal(ToggleName(entry.GroupId));
                        if (!groupOn)
                            continue;

                        int desired = ReadDesired(entry);
                        int? live = ReadDword(entry.KeyPath, entry.ValueName);

                        if (live.HasValue && live.Value == desired)
                        {
                            // Desired state is in effect. Record it as applied so a later
                            // revert of a Policies-hive value is recognized as an external
                            // (GPO) override we must not fight.
                            RecordApplied(entry, desired);
                            continue;
                        }

                        if (entry.IsPolicyHive)
                        {
                            int? applied = ReadDword(AppliedKeyPath, entry.Id);
                            if (applied.HasValue && applied.Value == desired)
                            {
                                // We had this value applied before and something reverted it:
                                // real policy management. Back off and flag; never re-fight.
                                conflicts.Add(entry.Id);
                                continue;
                            }
                        }

                        WriteDword(entry.KeyPath, entry.ValueName, desired);
                        RecordApplied(entry, desired);
                        wroteAny = true;
                    }

                    // Restart-needed bookkeeping (persisted so the dialog can show it and a
                    // later startup can clear it).
                    bool restart = ReadDword(TuningKeyPath, "RestartNeeded") == 1;
                    if (wroteAny)
                        restart = true;
                    else if (isStartup)
                        restart = false; // Outlook just booted with everything in sync.
                    WriteDword(TuningKeyPath, "RestartNeeded", restart ? 1 : 0);

                    WriteString(TuningKeyPath, "PolicyConflicts", string.Join(";", conflicts));
                    WriteString(TuningKeyPath, "LastReconcileUtc", DateTime.UtcNow.ToString("o"));

                    result.WroteAny = wroteAny;
                    result.RestartNeeded = restart;
                    result.PolicyConflicts = conflicts;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Tuning reconcile failed: " + ex.Message);
                }
                return result;
            }
        }

        // ===== Desired-state storage =====

        private static void EnsureInitialized()
        {
            using (var key = Registry.CurrentUser.CreateSubKey(TuningKeyPath))
            {
                if (key == null)
                    return;
                bool initialized = ReadDword(TuningKeyPath, "Initialized") == 1;
                if (!initialized)
                {
                    WriteDword(TuningKeyPath, "Enabled", 1);
                    WriteDword(TuningKeyPath, "SearchEnabled", 1);
                    WriteDword(TuningKeyPath, "CachingEnabled", 1);
                    WriteDword(TuningKeyPath, "OstEnabled", 1);
                    WriteDword(TuningKeyPath, "RestartNeeded", 0);
                    WriteDword(TuningKeyPath, "Initialized", 1);
                }
            }
            // Self-heal missing desired values (first run writes all of them).
            foreach (var entry in Catalog)
            {
                if (!ReadDword(DesiredKeyPath, entry.Id).HasValue)
                    WriteDword(DesiredKeyPath, entry.Id, entry.DefaultDesired);
            }
        }

        private static int ReadDesired(TuningEntry entry)
        {
            int? stored = ReadDword(DesiredKeyPath, entry.Id);
            return stored.HasValue ? stored.Value : entry.DefaultDesired;
        }

        private static void RecordApplied(TuningEntry entry, int value)
        {
            int? existing = ReadDword(AppliedKeyPath, entry.Id);
            if (!existing.HasValue || existing.Value != value)
                WriteDword(AppliedKeyPath, entry.Id, value);
        }

        private static string ToggleName(string groupId)
        {
            switch (groupId)
            {
                case GroupSearch: return "SearchEnabled";
                case GroupCaching: return "CachingEnabled";
                case GroupOst: return "OstEnabled";
                default: return "Enabled";
            }
        }

        private static bool GetToggleInternal(string name)
        {
            int? value = ReadDword(TuningKeyPath, name);
            return !value.HasValue || value.Value != 0; // missing = enabled (defaults ON)
        }

        private static List<string> GetPolicyConflictsInternal()
        {
            var list = new List<string>();
            try
            {
                string raw = ReadString(TuningKeyPath, "PolicyConflicts");
                if (!string.IsNullOrEmpty(raw))
                {
                    foreach (var part in raw.Split(';'))
                    {
                        if (part.Length > 0)
                            list.Add(part);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("GetPolicyConflicts: " + ex.Message);
            }
            return list;
        }

        // ===== Registry helpers (HKCU only; never throw out of public surface) =====

        private static int? ReadDword(string keyPath, string valueName)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(keyPath))
                {
                    var value = key?.GetValue(valueName);
                    if (value is int i)
                        return i;
                    return null;
                }
            }
            catch { return null; }
        }

        private static string ReadString(string keyPath, string valueName)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(keyPath))
                {
                    return key?.GetValue(valueName) as string;
                }
            }
            catch { return null; }
        }

        private static void WriteDword(string keyPath, string valueName, int value)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(keyPath))
            {
                key?.SetValue(valueName, value, RegistryValueKind.DWord);
            }
        }

        private static void WriteString(string keyPath, string valueName, string value)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(keyPath))
            {
                key?.SetValue(valueName, value ?? string.Empty, RegistryValueKind.String);
            }
        }
    }
}
