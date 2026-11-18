using System;
using System.Collections.Generic;
using System.Globalization;
using System.Security;

namespace OutlookAI.Core.AddIn
{
    /// <summary>One Outlook value the tuning service keeps applied: where it lives, which group owns it, and its shipped default.</summary>
    public sealed class TuningEntry
    {
        private static readonly string[] NoPolicy = new string[0];

        private static readonly Func<ITuningStore, bool> AnyDesiredValue = store => true;

        /// <summary>A value no policy governs, any desired number of which may be written.</summary>
        public TuningEntry(string id, string groupId, string keyPath, string valueName, int defaultDesired)
            : this(id, groupId, keyPath, valueName, defaultDesired, NoPolicy, NoPolicy, AnyDesiredValue)
        {
        }

        /// <summary>
        /// A value a policy may set: while any of <paramref name="policyValueNames"/> holds a DWORD
        /// under any of <paramref name="policyKeyPaths"/>, the walk leaves this value alone. And one
        /// written only while <paramref name="desiredIsOffered"/> says its desired number - read
        /// from the store it is handed, with whatever else it belongs with - is one the product
        /// offers.
        /// </summary>
        public TuningEntry(string id, string groupId, string keyPath, string valueName, int defaultDesired,
                           IList<string> policyKeyPaths, IList<string> policyValueNames, Func<ITuningStore, bool> desiredIsOffered)
        {
            Id = id;
            GroupId = groupId;
            KeyPath = keyPath;
            ValueName = valueName;
            DefaultDesired = defaultDesired;
            PolicyKeyPaths = new List<string>(policyKeyPaths ?? NoPolicy);
            PolicyValueNames = new List<string>(policyValueNames ?? NoPolicy);
            DesiredIsOffered = desiredIsOffered ?? AnyDesiredValue;
        }

        public string Id { get; }

        public string GroupId { get; }

        /// <summary>Relative to HKEY_CURRENT_USER.</summary>
        public string KeyPath { get; }

        public string ValueName { get; }

        public int DefaultDesired { get; }

        /// <summary>
        /// The policy keys that outrank this value - Microsoft 365 Cloud Policy's and Group Policy's
        /// mirror of its key - relative to HKEY_CURRENT_USER; empty when no policy is consulted. Read
        /// only: since Q145 (a) nothing in this product writes under <c>HKCU\Software\Policies</c>.
        /// </summary>
        public IReadOnlyList<string> PolicyKeyPaths { get; }

        /// <summary>
        /// The value names under <see cref="PolicyKeyPaths"/> that set this value. More than one
        /// where Outlook reads a pair as one setting: either half of the offline period, set by a
        /// policy, takes the whole period out of the user's hands.
        /// </summary>
        public IReadOnlyList<string> PolicyValueNames { get; }

        /// <summary>
        /// Whether the desired number is one the product writes, asked of the store before every
        /// write. A desired value typed into the registry by hand that is not one of them is
        /// skipped and listed in <see cref="TuningReconcileResult.NotOffered"/>, never written.
        /// </summary>
        public Func<ITuningStore, bool> DesiredIsOffered { get; }
    }

    /// <summary>One catalog value as the settings dialog shows it: desired against live, and why they differ.</summary>
    public sealed class TuningValueState
    {
        public TuningValueState(TuningEntry entry, int desired, int? live, bool groupEnabled, bool setByPolicy, bool needsAdministrator)
        {
            Entry = entry;
            Desired = desired;
            Live = live;
            GroupEnabled = groupEnabled;
            SetByPolicy = setByPolicy;
            NeedsAdministrator = needsAdministrator;
        }

        public TuningEntry Entry { get; }

        public int Desired { get; }

        public int? Live { get; }

        public bool InSync
        {
            get { return Live.HasValue && Live.Value == Desired; }
        }

        /// <summary>A policy sets this value now: the reconcile leaves it alone.</summary>
        public bool SetByPolicy { get; }

        /// <summary>The last reconcile was refused the write for lack of rights, and it is still not in effect.</summary>
        public bool NeedsAdministrator { get; }

        public bool GroupEnabled { get; }
    }

    /// <summary>The tuning state the settings dialog paints from.</summary>
    public sealed class TuningSnapshot
    {
        public bool MasterEnabled { get; internal set; }

        public bool SearchEnabled { get; internal set; }

        public bool CachingEnabled { get; internal set; }

        public bool OstEnabled { get; internal set; }

        public bool RestartNeeded { get; internal set; }

        public List<TuningValueState> Values { get; } = new List<TuningValueState>();

        /// <summary>Entry ids the last reconcile left alone because a policy sets them.</summary>
        public List<string> PolicyConflicts { get; } = new List<string>();

        /// <summary>Entry ids whose write Windows refused the last reconcile, for lack of rights.</summary>
        public List<string> NeedsAdministrator { get; } = new List<string>();
    }

    /// <summary>What one reconcile did.</summary>
    public sealed class TuningReconcileResult
    {
        public bool WroteAny { get; internal set; }

        public bool RestartNeeded { get; internal set; }

        /// <summary>
        /// Values a policy sets - Microsoft 365 Cloud Policy or Group Policy, an organization's or
        /// one an earlier OutlookAI wrote: left alone, never written over.
        /// </summary>
        public List<string> PolicyConflicts { get; } = new List<string>();

        /// <summary>Values whose write Windows refused for lack of rights: skipped, and the walk went on.</summary>
        public List<string> NeedsAdministrator { get; } = new List<string>();

        /// <summary>Values whose desired number is not one the product writes (typed into the registry by hand): skipped.</summary>
        public List<string> NotOffered { get; } = new List<string>();

        /// <summary>Values whose write failed for any other reason: skipped, and the walk went on.</summary>
        public List<string> Failed { get; } = new List<string>();

        /// <summary>The timestamp written as LastReconcileUtc; empty only if even that write failed.</summary>
        public string LastReconcileUtc { get; internal set; } = string.Empty;

        /// <summary>The first thing that went wrong outside a single value's write, for the debugger; empty when nothing did.</summary>
        public string Problem { get; internal set; } = string.Empty;
    }

    /// <summary>
    /// The registry the reconcile works against - HKEY_CURRENT_USER in the add-in
    /// (<c>OutlookTuningService</c>), a dictionary in the tests. Key paths are relative to the hive.
    /// </summary>
    public interface ITuningStore
    {
        /// <summary>The value when it is a REG_DWORD; null when it is absent, not a DWORD, or unreadable. Never throws.</summary>
        int? ReadDword(string keyPath, string valueName);

        /// <summary>The value when it is a string; empty when absent or unreadable. Never throws.</summary>
        string ReadString(string keyPath, string valueName);

        /// <summary>Writes a REG_DWORD, creating the key. THROWS when Windows refuses: <see cref="UnauthorizedAccessException"/> or <see cref="SecurityException"/> for rights.</summary>
        void WriteDword(string keyPath, string valueName, int value);

        /// <summary>Writes a REG_SZ, creating the key. Throws like <see cref="WriteDword"/>.</summary>
        void WriteString(string keyPath, string valueName, string value);

        /// <summary>
        /// Deletes a value; nothing when the key or the value is absent. Throws like
        /// <see cref="WriteDword"/> when Windows refuses. Only the removal of an earlier OutlookAI's
        /// policy values calls it (<c>EarlierPolicyRemoval</c>), and nothing calls that yet (Q151 (a)).
        /// </summary>
        void DeleteValue(string keyPath, string valueName);
    }

    /// <summary>
    /// THE TUNING RECONCILE, as logic over a registry it is handed (R12 / v3 plan section 0.5.3).
    ///
    /// <para>
    /// IT NEVER STOPS EARLY - decided by the maintainer 2026-10-03 (Q128). Until then the whole walk
    /// sat in one try block, and the first write that threw ended it: D25's five Cached Mode values
    /// were then written under <c>HKCU\Software\Policies</c>, a NOT elevated token may only read that
    /// key, so in the Outlook a user actually runs the first of them threw - and the two user Cached
    /// Mode values, the two OST sizes, the restart flag, PolicyConflicts and LastReconcileUtc were
    /// never written (Docs/live-tier-on-the-vm.md section 2.3, measured on a guest with a control).
    /// Now:
    /// </para>
    /// <list type="bullet">
    ///   <item>a value Windows refuses to write for lack of rights is SKIPPED and recorded in
    ///   <c>NeedsAdministrator</c>, which the settings dialog and <c>outlook_health</c> report;</item>
    ///   <item>a value that fails for any other reason is skipped too, and the walk goes on;</item>
    ///   <item>the bookkeeping - RestartNeeded, PolicyConflicts, NeedsAdministrator, and
    ///   LastReconcileUtc LAST - is written in a <c>finally</c>, each write on its own, so a
    ///   reconcile always says when it ran.</item>
    /// </list>
    ///
    /// <para>
    /// IT WRITES ONLY THE USER'S OWN SETTINGS, AND LEAVES A VALUE A POLICY SETS ALONE - Q145 (a),
    /// decided by the maintainer: no administrator, ever. Nothing in the catalog lives under
    /// <c>HKCU\Software\Policies</c> any more, so in practice nothing is refused for lack of rights.
    /// An entry that names policy keys (<see cref="TuningEntry.PolicyKeyPaths"/>) is skipped while a
    /// policy - Microsoft 365 Cloud Policy or Group Policy, an organization's or one an earlier
    /// OutlookAI wrote - sets its value, and listed in PolicyConflicts: the policy outranks the
    /// user's setting, and Outlook keeps the user's setting for when the policy goes. Found at once,
    /// by reading the policy keys - not, as before, only after the reconcile had applied a value and
    /// seen it reverted. And a desired number that is not one the product writes is skipped too
    /// (<see cref="TuningEntry.DesiredIsOffered"/>; Q143 for the offline period).
    /// </para>
    ///
    /// <para>
    /// The rest is as it always was: idempotent (only actual differences are written); restart-aware
    /// (a write takes effect at the NEXT Outlook start, so it sets RestartNeeded, which the first
    /// startup reconcile that writes nothing clears).
    /// </para>
    ///
    /// <para>
    /// SPLIT OUT OF the add-in's <c>OutlookTuningService</c> so that the T1 suite can pin the walk, and
    /// since the code-sharing refactor's stage S3 OutlookAI.Core's, in the add-in namespace
    /// (Q163 (a)): the suite pins the one compiled copy the add-in loads. Every value name comes from
    /// <see cref="AddInServerContract"/>, which the server's <c>HealthReporting</c> reads with.
    /// </para>
    /// </summary>
    public static class TuningReconciler
    {
        public const string GroupSearch = "search";
        public const string GroupCaching = "caching";
        public const string GroupOst = "ost";

        public const string TuningKeyPath = AddInServerContract.TuningKeyPath;
        public const string DesiredKeyPath = TuningKeyPath + @"\Desired";
        public const string AppliedKeyPath = TuningKeyPath + @"\Applied";

        /// <summary>Applies diffs, skipping what cannot be written, then writes the bookkeeping - always.</summary>
        public static TuningReconcileResult Reconcile(ITuningStore store, IReadOnlyList<TuningEntry> catalog, bool isStartup, DateTime utcNow)
        {
            var result = new TuningReconcileResult();
            try
            {
                EnsureInitialized(store, catalog);

                bool masterOn = IsToggleOn(store, AddInServerContract.TuningEnabledValueName);
                foreach (TuningEntry entry in catalog)
                {
                    if (!masterOn || !IsToggleOn(store, ToggleName(entry.GroupId)))
                        continue;
                    ReconcileEntry(store, entry, result);
                }
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                // Only the seeding of the add-in's OWN key can get here - every value's write is
                // guarded on its own below. The bookkeeping still runs.
                Note(result, ex);
            }
            finally
            {
                WriteBookkeeping(store, result, isStartup, utcNow);
            }
            return result;
        }

        private static void ReconcileEntry(ITuningStore store, TuningEntry entry, TuningReconcileResult result)
        {
            if (IsSetByPolicy(store, entry))
            {
                // A policy outranks the user's setting this would write: leave both alone, and say
                // so. Never fought, never written over (Q145 (a)).
                result.PolicyConflicts.Add(entry.Id);
                return;
            }

            if (!entry.DesiredIsOffered(store))
            {
                // A desired number typed into the registry by hand that the product does not
                // write - an offline period that is not one of Outlook's eleven, say (Q143).
                result.NotOffered.Add(entry.Id);
                return;
            }

            int desired = ReadDesired(store, entry);
            int? live = store.ReadDword(entry.KeyPath, entry.ValueName);

            if (live.HasValue && live.Value == desired)
            {
                // Desired state is in effect.
                TryRecordApplied(store, entry, desired, result);
                return;
            }

            try
            {
                store.WriteDword(entry.KeyPath, entry.ValueName, desired);
            }
            catch (Exception ex) when (IsAccessDenied(ex))
            {
                // Windows refused the write for lack of rights - as under HKCU\Software\Policies, the
                // case Q128 measured before Q145 (a) took every catalog value out of that key. Skip
                // it, say so, and go on: nothing after it depends on it.
                result.NeedsAdministrator.Add(entry.Id);
                return;
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                result.Failed.Add(entry.Id);
                Note(result, ex);
                return;
            }

            result.WroteAny = true;
            TryRecordApplied(store, entry, desired, result);
        }

        // Each write on its own, LastReconcileUtc last: one that fails does not stop the next.
        private static void WriteBookkeeping(ITuningStore store, TuningReconcileResult result, bool isStartup, DateTime utcNow)
        {
            bool restart = store.ReadDword(TuningKeyPath, AddInServerContract.TuningRestartNeededValueName) == 1;
            if (result.WroteAny)
                restart = true;
            else if (isStartup)
                restart = false; // Outlook just booted after every write it had been given.
            result.RestartNeeded = restart;

            TryWrite(result, () => store.WriteDword(TuningKeyPath, AddInServerContract.TuningRestartNeededValueName, restart ? 1 : 0));
            TryWrite(result, () => store.WriteString(TuningKeyPath, AddInServerContract.TuningPolicyConflictsValueName, JoinIdList(result.PolicyConflicts)));
            TryWrite(result, () => store.WriteString(TuningKeyPath, AddInServerContract.TuningNeedsAdministratorValueName, JoinIdList(result.NeedsAdministrator)));

            // Last, and always: the one value that says a reconcile ran to the end.
            string stamp = DateTime.SpecifyKind(utcNow, DateTimeKind.Utc).ToString("o", CultureInfo.InvariantCulture);
            if (TryWrite(result, () => store.WriteString(TuningKeyPath, AddInServerContract.TuningLastReconcileUtcValueName, stamp)))
                result.LastReconcileUtc = stamp;
        }

        /// <summary>Seeds the toggles once, and any missing desired value, every time. Writes only the add-in's own key.</summary>
        public static void EnsureInitialized(ITuningStore store, IReadOnlyList<TuningEntry> catalog)
        {
            bool initialized = store.ReadDword(TuningKeyPath, AddInServerContract.TuningInitializedValueName) == 1;
            if (!initialized)
            {
                store.WriteDword(TuningKeyPath, AddInServerContract.TuningEnabledValueName, 1);
                store.WriteDword(TuningKeyPath, AddInServerContract.TuningSearchEnabledValueName, 1);
                store.WriteDword(TuningKeyPath, AddInServerContract.TuningCachingEnabledValueName, 1);
                store.WriteDword(TuningKeyPath, AddInServerContract.TuningOstEnabledValueName, 1);
                store.WriteDword(TuningKeyPath, AddInServerContract.TuningRestartNeededValueName, 0);
                store.WriteDword(TuningKeyPath, AddInServerContract.TuningInitializedValueName, 1);
            }

            // Self-heal missing desired values (first run writes all of them).
            foreach (TuningEntry entry in catalog)
            {
                if (!store.ReadDword(DesiredKeyPath, entry.Id).HasValue)
                    store.WriteDword(DesiredKeyPath, entry.Id, entry.DefaultDesired);
            }
        }

        /// <summary>Desired against live for every catalog value, plus the toggles and the last reconcile's lists.</summary>
        public static TuningSnapshot ReadSnapshot(ITuningStore store, IReadOnlyList<TuningEntry> catalog)
        {
            EnsureInitialized(store, catalog);
            List<string> conflicts = ParseIdList(store.ReadString(TuningKeyPath, AddInServerContract.TuningPolicyConflictsValueName));
            List<string> needsAdministrator = ParseIdList(store.ReadString(TuningKeyPath, AddInServerContract.TuningNeedsAdministratorValueName));

            var snapshot = new TuningSnapshot
            {
                MasterEnabled = IsToggleOn(store, AddInServerContract.TuningEnabledValueName),
                SearchEnabled = IsToggleOn(store, AddInServerContract.TuningSearchEnabledValueName),
                CachingEnabled = IsToggleOn(store, AddInServerContract.TuningCachingEnabledValueName),
                OstEnabled = IsToggleOn(store, AddInServerContract.TuningOstEnabledValueName),
                RestartNeeded = store.ReadDword(TuningKeyPath, AddInServerContract.TuningRestartNeededValueName) == 1,
            };
            snapshot.PolicyConflicts.AddRange(conflicts);
            snapshot.NeedsAdministrator.AddRange(needsAdministrator);

            foreach (TuningEntry entry in catalog)
            {
                int desired = ReadDesired(store, entry);
                int? live = store.ReadDword(entry.KeyPath, entry.ValueName);
                bool inSync = live.HasValue && live.Value == desired;
                snapshot.Values.Add(new TuningValueState(
                    entry,
                    desired,
                    live,
                    snapshot.MasterEnabled && IsToggleOn(store, ToggleName(entry.GroupId)),
                    // Read now, not from the last reconcile's list: a policy may have come or gone since.
                    IsSetByPolicy(store, entry),
                    // Recorded by the last reconcile AND still not in effect: somebody may have set
                    // it since.
                    needsAdministrator.Contains(entry.Id) && !inSync));
            }
            return snapshot;
        }

        /// <summary>
        /// True while any of the entry's <see cref="TuningEntry.PolicyValueNames"/> holds a DWORD
        /// under any of its <see cref="TuningEntry.PolicyKeyPaths"/> - whoever wrote it. Reads only.
        /// </summary>
        public static bool IsSetByPolicy(ITuningStore store, TuningEntry entry)
        {
            foreach (string keyPath in entry.PolicyKeyPaths)
            {
                foreach (string valueName in entry.PolicyValueNames)
                {
                    if (store.ReadDword(keyPath, valueName).HasValue)
                        return true;
                }
            }
            return false;
        }

        /// <summary>
        /// True for the two ways .NET reports "the registry refused this for lack of rights":
        /// <see cref="UnauthorizedAccessException"/> (CreateSubKey, SetValue) and
        /// <see cref="SecurityException"/> (opening a key writable that is not).
        /// </summary>
        public static bool IsAccessDenied(Exception ex)
        {
            return ex is UnauthorizedAccessException || ex is SecurityException;
        }

        public static int ReadDesired(ITuningStore store, TuningEntry entry)
        {
            int? stored = store.ReadDword(DesiredKeyPath, entry.Id);
            return stored.HasValue ? stored.Value : entry.DefaultDesired;
        }

        public static string ToggleName(string groupId)
        {
            switch (groupId)
            {
                case GroupSearch: return AddInServerContract.TuningSearchEnabledValueName;
                case GroupCaching: return AddInServerContract.TuningCachingEnabledValueName;
                case GroupOst: return AddInServerContract.TuningOstEnabledValueName;
                default: return AddInServerContract.TuningEnabledValueName;
            }
        }

        /// <summary>Missing means on: the toggles default ON.</summary>
        public static bool IsToggleOn(ITuningStore store, string name)
        {
            int? value = store.ReadDword(TuningKeyPath, name);
            return !value.HasValue || value.Value != 0;
        }

        /// <summary>';'-joined entry ids, as PolicyConflicts and NeedsAdministrator store them: empties dropped, duplicates once.</summary>
        public static List<string> ParseIdList(string raw)
        {
            var list = new List<string>();
            if (string.IsNullOrEmpty(raw))
                return list;
            foreach (string part in raw.Split(';'))
            {
                if (part.Length > 0 && !list.Contains(part))
                    list.Add(part);
            }
            return list;
        }

        public static string JoinIdList(IEnumerable<string> ids)
        {
            return string.Join(";", ids);
        }

        private static void TryRecordApplied(ITuningStore store, TuningEntry entry, int value, TuningReconcileResult result)
        {
            int? existing = store.ReadDword(AppliedKeyPath, entry.Id);
            if (existing.HasValue && existing.Value == value)
                return;
            TryWrite(result, () => store.WriteDword(AppliedKeyPath, entry.Id, value));
        }

        private static bool TryWrite(TuningReconcileResult result, Action write)
        {
            try
            {
                write();
                return true;
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                Note(result, ex);
                return false;
            }
        }

        private static void Note(TuningReconcileResult result, Exception ex)
        {
            if (result.Problem.Length == 0)
                result.Problem = ex.GetType().Name + ": " + ex.Message;
        }
    }
}
