using System;
using System.Collections.Generic;
using Microsoft.Win32;
using OutlookAI.Core.AddIn;

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
    /// idempotent, restart-aware, never stopped early by a value it cannot write (Q128), and
    /// never writing over a value a policy sets (Q145 (a)).
    ///
    /// THIS CLASS is the add-in's half: the catalog for the Office version actually installed,
    /// the HKEY_CURRENT_USER store the reconcile runs against, the lock, and the surface the
    /// settings dialog calls. Nothing here throws out of its public surface.
    ///
    /// EVERYTHING IT WRITES IS THE USER'S OWN - Q145 (a), decided by the maintainer: no
    /// administrator, no UAC prompt, no elevated helper. The Cached Mode values are written as the
    /// user's own settings, and Microsoft 365 Cloud Policy and Group Policy are only READ, to show a
    /// value an organization sets - or one an earlier OutlookAI wrote as policy - as such and leave
    /// it alone (<see cref="CachedModeSettings"/>). Disabling a group (or the master switch) stops
    /// managing its values; already-written Outlook values are left in place, mirroring the
    /// uninstall behavior described in the plan.
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

        // The other keys are the add-in's alone - the server neither reads nor reports them - but
        // the HIVE ROOTS are not: the Policies roots in particular are read on both sides (here,
        // through CachedModeSettings, for the Cached Mode values; in HealthReporting for the Search
        // policy). So the roots come from OfficeVersions, and this file spells no Office hive path.
        private static readonly string PstKeyPath =
            OfficeVersions.OutlookKeyPath(OfficeVersion) + @"\PST";

        internal const string GroupSearch = TuningReconciler.GroupSearch;
        internal const string GroupCaching = TuningReconciler.GroupCaching;
        internal const string GroupOst = TuningReconciler.GroupOst;

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

        /// <summary>
        /// THE SEAM for an existing account's own offline period (Q145 (a)): this build cannot write
        /// it, so the dialog tells the user how and the removal of an earlier OutlookAI's policy
        /// values refuses. Replace it with the writer the Exchange VM measurement makes possible -
        /// <see cref="IAccountOfflinePeriods"/> says what it must do.
        /// </summary>
        private static readonly IAccountOfflinePeriods AccountPeriods = AccountOfflinePeriodsNotYetBuilt.Instance;

        // The full desired-state catalog (defaults per v3 plan D22/D24/D25). The registry
        // stores the desired NUMBERS (self-healing, future-tunable); this catalog is the
        // authoritative structure: which value lives in which key, and its default.
        private static readonly TuningEntry[] Catalog = BuildCatalog();

        private static TuningEntry[] BuildCatalog()
        {
            var catalog = new List<TuningEntry>
            {
                // Search (D22) - the user's proven local-search setup.
                new TuningEntry("search.DisableServerAssistedSearch", GroupSearch, SearchKeyPath, AddInServerContract.DisableServerAssistedSearchValueName, 1),
                new TuningEntry("search.SearchResultsCap",            GroupSearch, SearchKeyPath, "SearchResultsCap",            0),
                new TuningEntry("search.IncludeDeletedItems",         GroupSearch, SearchKeyPath, "IncludeDeletedItems",         1),
                new TuningEntry("search.DefaultSearchScope",          GroupSearch, SearchKeyPath, "DefaultSearchScope",          2),
            };

            // Full caching (D25) - the offline period All for accounts added from now on, and
            // shared-folder caching, as the user's own settings; each left alone while a policy sets
            // it. The five values, their defaults and the values a user may choose come from
            // CachedModeSettings, which the test suite pins.
            catalog.AddRange(CachedModeSettings.CatalogEntries(OfficeVersion));

            // OST headroom (D25) - 100 GB max / ~94 GB warn so full caching never stalls at
            // the default 50 GB cap. Outlook stores both as megabytes.
            catalog.Add(new TuningEntry(OstMaxEntryId,           GroupOst, PstKeyPath, "MaxLargeFileSize",  OstMaxDefaultMb));
            catalog.Add(new TuningEntry("ost.WarnLargeFileSize", GroupOst, PstKeyPath, "WarnLargeFileSize", OstWarnDefaultMb));
            return catalog.ToArray();
        }

        internal static IReadOnlyList<TuningEntry> Entries
        {
            get { return Catalog; }
        }

        // ===== Public operations =====

        /// <summary>Startup reconcile: applies diffs and maintains the restart-needed flag
        /// (clears it when Outlook booted with nothing left to write).</summary>
        public static TuningReconcileResult ReconcileOnStartup()
        {
            return Reconcile(true);
        }

        /// <summary>Mid-session reconcile (settings dialog): applies diffs; only ever SETS the
        /// restart-needed flag - a mid-session "everything matches" must not clear it because
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
                        System.Diagnostics.Debug.WriteLine("Tuning reconcile: Windows refused " + string.Join(", ", result.NeedsAdministrator));
                    if (result.NotOffered.Count > 0)
                        System.Diagnostics.Debug.WriteLine("Tuning reconcile: desired value not offered, left unwritten: " + string.Join(", ", result.NotOffered));
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

        // ===== The Cached Mode settings (Q145 (a), Q143, Q151 (a)) =====

        /// <summary>
        /// The Cached Mode settings as the dialog lists them: what Outlook uses and who sets it, the
        /// value OutlookAI writes, and the state - read now, from the user's own key and both policy
        /// keys. Empty if even that read fails.
        /// </summary>
        public static List<CachedModeSettingState> GetCachedModeStates(TuningSnapshot snapshot)
        {
            lock (_gate)
            {
                try
                {
                    bool managed = snapshot.MasterEnabled && snapshot.CachingEnabled;
                    return CachedModeSettings.ReadStates(Store, OfficeVersion, managed, snapshot.NeedsAdministrator);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Cached Mode states: " + ex.Message);
                    return new List<CachedModeSettingState>();
                }
            }
        }

        /// <summary>
        /// The settings dialog's choice for one Cached Mode setting, stored as its DESIRED value -
        /// the add-in's own key, always writable; the period as its (months, days) pair. Writing the
        /// live value is the reconcile's job. Refuses (false) a setting or a position the setting
        /// does not offer, so nothing but one of its values is ever stored from here (Q143).
        /// </summary>
        public static bool SetCachedModeChoice(string settingKey, int choiceIndex)
        {
            lock (_gate)
            {
                try
                {
                    List<KeyValuePair<string, int>> desired = CachedModeSettings.DesiredFor(CachedModeSettings.Setting(settingKey), choiceIndex);
                    if (desired.Count == 0)
                        return false;
                    TuningReconciler.EnsureInitialized(Store, Catalog);
                    foreach (KeyValuePair<string, int> value in desired)
                        Store.WriteDword(TuningReconciler.DesiredKeyPath, value.Key, value.Value);
                    return true;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("SetCachedModeChoice: " + ex.Message);
                    return false;
                }
            }
        }

        /// <summary>Whether this build can set an existing account's own offline period. False until the seam is filled.</summary>
        public static bool CanSetExistingAccountPeriods
        {
            get { return AccountPeriods.CanSetEveryAccount; }
        }

        /// <summary>
        /// REMOVES THE CACHED MODE POLICY VALUES AN EARLIER OUTLOOKAI WROTE - and NOTHING CALLS IT
        /// (Q151 (a)): those values stay until the maintainer asks for their removal, and wiring this
        /// to a button or a startup is that request's change. What it does, and why it refuses while
        /// the seam above cannot set every account's own period to All first, is
        /// <see cref="EarlierPolicyRemoval"/>'s to say. Reconciles afterwards, so the user's own
        /// settings the policy values stood over are written at once.
        /// </summary>
        public static EarlierPolicyRemovalResult RemoveEarlierPolicyValues()
        {
            EarlierPolicyRemovalResult result;
            lock (_gate)
            {
                result = EarlierPolicyRemoval.Remove(Store, OfficeVersion, AccountPeriods);
            }
            if (result.Removed.Count > 0)
                ReconcileFromUi();
            return result;
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
        /// value Windows refused for lack of rights from one it wrote.
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

            public void DeleteValue(string keyPath, string valueName)
            {
                // Opened writable only when it exists: a key that is not there holds nothing to
                // delete, and opening one writable that the user may only read throws
                // SecurityException - which the caller reads as a refusal, as it should.
                using (var key = Registry.CurrentUser.OpenSubKey(keyPath, writable: true))
                {
                    if (key != null)
                        key.DeleteValue(valueName, throwOnMissingValue: false);
                }
            }
        }
    }
}
