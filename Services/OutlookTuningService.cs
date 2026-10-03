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
    /// settings dialog) the live registry is reconciled against desired state - by
    /// <see cref="TuningReconciler"/>, which holds the walk and says what it guarantees:
    /// idempotent, GPO-respecting, restart-aware, and - since Q128 - never stopped early by a
    /// value it cannot write.
    ///
    /// THIS CLASS is the add-in's half: the catalog for the Office version actually installed,
    /// the HKEY_CURRENT_USER store the reconcile runs against, the lock, and the surface the
    /// settings dialog calls. Nothing here throws out of its public surface.
    ///
    /// NOT EVERYTHING IS WRITABLE. Four groups of values live in the user's own hive and the add-in
    /// can always write them; D25's five Cached Mode values live under HKCU\Software\Policies, which
    /// a NOT elevated token may only read - and the Outlook a user runs is NOT elevated. Those are
    /// skipped and recorded as needing an administrator; the settings dialog applies them through
    /// the elevated helper (<see cref="PolicyElevation"/>, OutlookAI.PolicyWriter.exe) when the user
    /// asks. Disabling a group (or the master switch) stops managing its values; already-written
    /// Outlook values are left in place, mirroring the uninstall behavior described in the plan.
    /// </summary>
    internal static class OutlookTuningService
    {
        /// <summary>
        /// The key this service owns. The MCP server's <c>outlook_health</c> reads it to report
        /// the tuning state, so the path and every value name under it live in
        /// <see cref="AddInServerContract"/> - one definition, compiled into both, rather than a
        /// comment on each side claiming to mirror the other.
        /// </summary>
        internal const string TuningKeyPath = AddInServerContract.TuningKeyPath;

        /// <summary>
        /// The Office major version whose hives these values are written into, detected once
        /// per session. It used to be the literal "16.0" in all four paths below, which meant
        /// that on Outlook 2013 (15.0) or a future 17.0 every write landed in a hive Outlook
        /// never reads: the settings dialog showed every value as "(not set)" for ever,
        /// RestartNeeded never cleared, and the user was told to restart Outlook indefinitely.
        /// <see cref="OfficeVersions"/> is the one list of versions the whole add-in agrees on.
        /// </summary>
        private static readonly string OfficeVersion = OfficeVersions.DetectOutlookVersion();

        /// <summary>
        /// Outlook's Search key, built by <see cref="OfficeVersions.OutlookSearchKeyPath"/>
        /// rather than concatenated here. THE SERVER READS THIS KEY: <c>HealthReporting</c> takes
        /// <c>DisableServerAssistedSearch</c> out of it and reports it as
        /// <c>uiSearchBackend</c>, which is what tells an agent whether the user's Outlook search
        /// box is looking at the same corpus the agent searched. Both sides therefore build one
        /// address from one expression - hand-building it here was the last surviving copy of a
        /// path spelled twice across a boundary no compiler crosses, and a typo in either copy
        /// aims at a key Outlook never touches while both halves go on working.
        /// </summary>
        private static readonly string SearchKeyPath = OfficeVersions.OutlookSearchKeyPath(OfficeVersion);

        // The other three are the add-in's alone - the server neither reads nor reports them -
        // but the HIVE ROOTS are not: the Policies prefix in particular is built on both sides
        // (here for D25's sync-slider values, through CachedModePolicy.KeyPath, which the
        // elevated helper builds its path with too; in HealthReporting for the Search policy).
        // So the roots come from OfficeVersions, and this file spells no Office hive path.
        private static readonly string CachedModePolicyKeyPath = CachedModePolicy.KeyPath(OfficeVersion);
        private static readonly string CachedModeUserKeyPath =
            OfficeVersions.OutlookKeyPath(OfficeVersion) + @"\" + CachedModePolicy.SubKeyName;
        private static readonly string PstKeyPath =
            OfficeVersions.OutlookKeyPath(OfficeVersion) + @"\PST";

        internal const string GroupSearch = TuningReconciler.GroupSearch;
        internal const string GroupCaching = TuningReconciler.GroupCaching;
        internal const string GroupOst = TuningReconciler.GroupOst;

        /// <summary>The id prefix of D25's five Policies-hive entries; the rest of each id is the value name.</summary>
        internal const string PolicyEntryIdPrefix = "caching.policy.";

        /// <summary>
        /// The OST size cap, in the megabytes Outlook stores it as. Named because the settings
        /// dialog's tick box has to say what the cap is, and a caption that spells the number
        /// out by hand states a figure the product may no longer be applying - the desired
        /// values live in the registry and are meant to be tunable. See
        /// <see cref="DescribeOstMaxSize"/>, which the caption is built from.
        /// </summary>
        internal const string OstMaxEntryId = "ost.MaxLargeFileSize";
        private const int OstMaxDefaultMb = 102400;   // 100 GB
        private const int OstWarnDefaultMb = 96256;   // ~94 GB, Outlook's own warn-below-max gap
        private const int MegabytesPerGigabyte = 1024;

        private static readonly object _gate = new object();

        private static readonly ITuningStore Store = new CurrentUserStore();

        // The full desired-state catalog (defaults per v3 plan D22/D24/D25). The registry
        // stores the desired NUMBERS (self-healing, future-tunable); this catalog is the
        // authoritative structure: which value lives in which key, and its default.
        private static readonly TuningEntry[] Catalog = new[]
        {
            // Search (D22) — the user's proven local-search setup.
            new TuningEntry("search.DisableServerAssistedSearch", GroupSearch, SearchKeyPath, AddInServerContract.DisableServerAssistedSearchValueName, 1, false),
            new TuningEntry("search.SearchResultsCap",            GroupSearch, SearchKeyPath, "SearchResultsCap",            0, false),
            new TuningEntry("search.IncludeDeletedItems",         GroupSearch, SearchKeyPath, "IncludeDeletedItems",         1, false),
            new TuningEntry("search.DefaultSearchScope",          GroupSearch, SearchKeyPath, "DefaultSearchScope",          2, false),

            // Full caching (D25) — sync slider = All for existing accounts (Policies hive)
            // and future accounts (user hive), plus shared-folder caching. The five policy values,
            // their defaults and the values a user may choose come from CachedModePolicy, the list
            // the elevated helper enforces.
            PolicyEntry(CachedModePolicy.SyncWindowSetting),
            PolicyEntry(CachedModePolicy.SyncWindowSettingDays),
            PolicyEntry(CachedModePolicy.DownloadSharedFolders),
            PolicyEntry(CachedModePolicy.CacheOthersMail),
            PolicyEntry(CachedModePolicy.DisableSyncSliderForSharedMailbox),
            new TuningEntry("caching.user.SyncWindowSetting",     GroupCaching, CachedModeUserKeyPath, CachedModePolicy.SyncWindowSetting,     0, false),
            new TuningEntry("caching.user.SyncWindowSettingDays", GroupCaching, CachedModeUserKeyPath, CachedModePolicy.SyncWindowSettingDays, 0, false),

            // OST headroom (D25) — 100 GB max / ~94 GB warn so full caching never stalls at
            // the default 50 GB cap. Outlook stores both as megabytes.
            new TuningEntry(OstMaxEntryId,           GroupOst, PstKeyPath, "MaxLargeFileSize",  OstMaxDefaultMb,  false),
            new TuningEntry("ost.WarnLargeFileSize", GroupOst, PstKeyPath, "WarnLargeFileSize", OstWarnDefaultMb, false),
        };

        private static TuningEntry PolicyEntry(string valueName)
        {
            return new TuningEntry(PolicyEntryIdPrefix + valueName, GroupCaching, CachedModePolicyKeyPath, valueName,
                                   CachedModePolicy.ShippedDefault(valueName), true);
        }

        internal static IReadOnlyList<TuningEntry> Entries
        {
            get { return Catalog; }
        }

        /// <summary>The Office major the policy values are written for - what the elevated helper is told.</summary>
        internal static string PolicyOfficeVersion
        {
            get { return OfficeVersion; }
        }

        // ===== Public operations =====

        /// <summary>Startup reconcile: applies diffs and maintains the restart-needed flag
        /// (clears it when Outlook booted with nothing left to write).</summary>
        public static TuningReconcileResult ReconcileOnStartup()
        {
            return Reconcile(true);
        }

        /// <summary>Mid-session reconcile (settings dialog): applies diffs; only ever SETS the
        /// restart-needed flag — a mid-session "everything matches" must not clear it because
        /// the running Outlook may still be on pre-change values.</summary>
        public static TuningReconcileResult ReconcileFromUi()
        {
            return Reconcile(false);
        }

        private static TuningReconcileResult Reconcile(bool isStartup)
        {
            lock (_gate)
            {
                try
                {
                    TuningReconcileResult result = TuningReconciler.Reconcile(Store, Catalog, isStartup, DateTime.UtcNow);
                    if (result.Problem.Length > 0)
                        System.Diagnostics.Debug.WriteLine("Tuning reconcile: " + result.Problem);
                    if (result.NeedsAdministrator.Count > 0)
                        System.Diagnostics.Debug.WriteLine("Tuning reconcile: needs an administrator for " + string.Join(", ", result.NeedsAdministrator));
                    return result;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Tuning reconcile failed: " + ex.Message);
                    return new TuningReconcileResult();
                }
            }
        }

        /// <summary>Read-only view of desired vs live state for the settings dialog.</summary>
        public static TuningSnapshot GetSnapshot()
        {
            lock (_gate)
            {
                try
                {
                    return TuningReconciler.ReadSnapshot(Store, Catalog);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Tuning snapshot failed: " + ex.Message);
                    return new TuningSnapshot();
                }
            }
        }

        public static bool GetMasterEnabled()
        {
            lock (_gate) { try { TuningReconciler.EnsureInitialized(Store, Catalog); return TuningReconciler.IsToggleOn(Store, AddInServerContract.TuningEnabledValueName); } catch { return true; } }
        }

        public static void SetMasterEnabled(bool enabled)
        {
            lock (_gate) { try { TuningReconciler.EnsureInitialized(Store, Catalog); Store.WriteDword(TuningKeyPath, AddInServerContract.TuningEnabledValueName, enabled ? 1 : 0); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine("SetMasterEnabled: " + ex.Message); } }
        }

        public static bool GetGroupEnabled(string groupId)
        {
            lock (_gate) { try { TuningReconciler.EnsureInitialized(Store, Catalog); return TuningReconciler.IsToggleOn(Store, TuningReconciler.ToggleName(groupId)); } catch { return true; } }
        }

        public static void SetGroupEnabled(string groupId, bool enabled)
        {
            lock (_gate) { try { TuningReconciler.EnsureInitialized(Store, Catalog); Store.WriteDword(TuningKeyPath, TuningReconciler.ToggleName(groupId), enabled ? 1 : 0); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine("SetGroupEnabled: " + ex.Message); } }
        }

        public static bool GetRestartNeeded()
        {
            lock (_gate) { try { return Store.ReadDword(TuningKeyPath, AddInServerContract.TuningRestartNeededValueName) == 1; } catch { return false; } }
        }

        /// <summary>
        /// Something other than the reconcile changed a live value under a running Outlook - the
        /// elevated helper - so the restart flag goes up, exactly as for a write of the reconcile's
        /// own. The next startup reconcile that writes nothing clears it, as always.
        /// </summary>
        public static void MarkRestartNeeded()
        {
            lock (_gate) { try { Store.WriteDword(TuningKeyPath, AddInServerContract.TuningRestartNeededValueName, 1); } catch (Exception ex) { System.Diagnostics.Debug.WriteLine("MarkRestartNeeded: " + ex.Message); } }
        }

        /// <summary>
        /// The settings dialog's choice for one of the five policy values: stored as its DESIRED
        /// value, which is the add-in's own key and always writable. Writing the live value is the
        /// reconcile's job - and, where it needs an administrator, the helper's. Refuses (false) a
        /// name or value <see cref="CachedModePolicy"/> does not offer, so the dialog can never set
        /// up a request the helper would refuse.
        /// </summary>
        public static bool SetPolicyDesired(string valueName, int value)
        {
            if (!CachedModePolicy.IsAllowed(valueName, value))
                return false;
            lock (_gate)
            {
                try
                {
                    TuningReconciler.EnsureInitialized(Store, Catalog);
                    Store.WriteDword(TuningReconciler.DesiredKeyPath, PolicyEntryIdPrefix + valueName, value);
                    return true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("SetPolicyDesired: " + ex.Message);
                    return false;
                }
            }
        }

        /// <summary>
        /// What the elevated helper should write: each policy value the last reconcile could not
        /// write, whose group is on and which is still not in effect, with its desired number - in
        /// catalog order. A desired number the helper would refuse (one typed into the registry by
        /// hand) is left out and named in <paramref name="notOffered"/> instead.
        /// </summary>
        public static List<KeyValuePair<string, int>> PolicyValuesNeedingAdministrator(out List<string> notOffered)
        {
            var request = new List<KeyValuePair<string, int>>();
            notOffered = new List<string>();
            foreach (TuningValueState v in GetSnapshot().Values)
            {
                if (!v.Entry.IsPolicyHive || !v.GroupEnabled || !v.NeedsAdministrator)
                    continue;
                if (CachedModePolicy.IsAllowed(v.Entry.ValueName, v.Desired))
                    request.Add(new KeyValuePair<string, int>(v.Entry.ValueName, v.Desired));
                else
                    notOffered.Add(v.Entry.ValueName + "=" + v.Desired);
            }
            return request;
        }

        /// <summary>
        /// The OST size cap this machine is actually being given, worded for a caption:
        /// "100 GB". Read from the DESIRED value rather than from the shipped default, because
        /// the desired numbers live in the registry and are meant to be tunable - a caption
        /// with "100 GB" typed into it would go on saying so after somebody changed the number.
        /// </summary>
        public static string DescribeOstMaxSize()
        {
            int megabytes = OstMaxDefaultMb;
            lock (_gate)
            {
                try
                {
                    foreach (var entry in Catalog)
                    {
                        if (entry.Id == OstMaxEntryId)
                        {
                            megabytes = TuningReconciler.ReadDesired(Store, entry);
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("DescribeOstMaxSize: " + ex.Message);
                }
            }

            double gigabytes = (double)megabytes / MegabytesPerGigabyte;
            return gigabytes >= 1 && gigabytes == Math.Floor(gigabytes)
                ? ((int)gigabytes).ToString(System.Globalization.CultureInfo.CurrentCulture) + " GB"
                : megabytes.ToString(System.Globalization.CultureInfo.CurrentCulture) + " MB";
        }

        // ===== The store: HKEY_CURRENT_USER =====

        /// <summary>
        /// HKCU, as the reconcile sees it. Reads never throw; writes DO, so the reconcile can tell a
        /// value Windows refused for lack of rights (HKCU\Software\Policies under a filtered token)
        /// from one it wrote.
        /// </summary>
        private sealed class CurrentUserStore : ITuningStore
        {
            public int? ReadDword(string keyPath, string valueName)
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

            public string ReadString(string keyPath, string valueName)
            {
                try
                {
                    using (var key = Registry.CurrentUser.OpenSubKey(keyPath))
                    {
                        return key?.GetValue(valueName) as string ?? string.Empty;
                    }
                }
                catch { return string.Empty; }
            }

            public void WriteDword(string keyPath, string valueName, int value)
            {
                using (var key = Registry.CurrentUser.CreateSubKey(keyPath))
                {
                    if (key == null)
                        throw new InvalidOperationException("HKCU\\" + keyPath + " could not be opened for writing.");
                    key.SetValue(valueName, value, RegistryValueKind.DWord);
                }
            }

            public void WriteString(string keyPath, string valueName, string value)
            {
                using (var key = Registry.CurrentUser.CreateSubKey(keyPath))
                {
                    if (key == null)
                        throw new InvalidOperationException("HKCU\\" + keyPath + " could not be opened for writing.");
                    key.SetValue(valueName, value ?? string.Empty, RegistryValueKind.String);
                }
            }
        }
    }
}
