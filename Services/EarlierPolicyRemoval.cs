using System;
using System.Collections.Generic;

namespace OutlookAI.Services
{
    /// <summary>
    /// ===================================================================================
    /// REMOVES THE CACHED MODE POLICY VALUES AN EARLIER OUTLOOKAI WROTE. NOTHING CALLS IT.
    /// ===================================================================================
    ///
    /// <para>
    /// Q151 (a), decided by the maintainer: the policy values OutlookAI 3.0.0.319 to 3.1.0.325 wrote
    /// under <c>HKCU\Software\Policies\Microsoft\Office\&lt;major&gt;\Outlook\Cached Mode</c> - where
    /// the user's token could write them - STAY, labelled "Set by an earlier OutlookAI" in OutlookAI
    /// Settings, until he asks for their removal. This is that removal, built and tested so it is
    /// ready, and run nowhere: no startup, no reconcile, no button calls it. Wiring it to something
    /// is the change his request will be.
    /// </para>
    ///
    /// <para>
    /// WHAT IT REMOVES: a Group Policy value only where <c>Tuning\Applied\caching.policy.&lt;name&gt;</c>
    /// holds the same number (<see cref="CachedModeValueReading.GroupPolicyIsEarlierOutlookAI"/>) - the
    /// test the dialog labels by - and that record with it, and the earlier OutlookAI's
    /// <c>Tuning\Desired\caching.policy.&lt;name&gt;</c> beside it. Never a Cloud Policy value, never
    /// a Group Policy value with any other number, never the user's own settings.
    /// </para>
    ///
    /// <para>
    /// EVERY ACCOUNT'S OWN PERIOD GOES TO ALL FIRST, or nothing is removed. A policy period holds every
    /// account at that period whatever the account's own says, and leaves the account's own unchanged
    /// (measured on the maintainer's workstation: a policy of All over an account set to 12 months).
    /// Removing the policy hands each account back its own period - and a shorter period makes Outlook
    /// delete the older mail from the offline file (KB 2733062, "a local-only deletion"), out of
    /// OutlookAI's search. So when the period is among the values to remove, every account's own
    /// period is set to All first, through <see cref="IAccountOfflinePeriods"/> - which this build
    /// cannot do yet, so today such a removal is REFUSED, whole, before anything changes.
    /// </para>
    ///
    /// <para>
    /// A value Windows will not let the user delete - one an ELEVATED Outlook wrote, so the key is
    /// the administrators' - is reported, not retried with more rights: Q145 (a), no administrator.
    /// Outlook reads policy values when it starts, so a removal raises RestartNeeded.
    /// </para>
    ///
    /// <para>
    /// FRAMEWORK-NEUTRAL and INTERNAL, like <see cref="CachedModeSettings"/>, which the test project
    /// links beside it.
    /// </para>
    /// </summary>
    internal static class EarlierPolicyRemoval
    {
        /// <summary>The values an earlier OutlookAI wrote as Group Policy and that still hold its number, in catalog order.</summary>
        internal static List<string> Plan(ITuningStore store, string officeVersion)
        {
            var names = new List<string>();
            foreach (string name in CachedModeSettings.ValueNames)
            {
                if (CachedModeSettings.Read(store, officeVersion, name).GroupPolicyIsEarlierOutlookAI)
                    names.Add(name);
            }
            return names;
        }

        /// <summary>
        /// Removes what <see cref="Plan"/> names - after every account's own period is All, when the
        /// period is among them - or refuses before changing anything. Never throws.
        /// </summary>
        internal static EarlierPolicyRemovalResult Remove(ITuningStore store, string officeVersion, IAccountOfflinePeriods accounts)
        {
            var result = new EarlierPolicyRemovalResult();
            try
            {
                List<string> plan = Plan(store, officeVersion);
                if (plan.Count == 0)
                {
                    result.Refusal = "No Cached Mode policy value here holds what an earlier OutlookAI wrote, so there is nothing to remove.";
                    return result;
                }

                if (plan.Contains(CachedModeSettings.SyncWindowSetting) || plan.Contains(CachedModeSettings.SyncWindowSettingDays))
                {
                    if (!accounts.CanSetEveryAccount)
                    {
                        result.Refusal = "Nothing was removed: the offline period is among them, and every account's own period "
                                         + "must be All before it goes, or Outlook drops older mail from the offline copy. "
                                         + accounts.WhyNot;
                        return result;
                    }
                    AccountPeriodChange change = accounts.SetEveryAccount(CachedModeSettings.All);
                    if (!change.Done)
                    {
                        result.Refusal = "Nothing was removed: setting every account's own period to All did not succeed. " + change.Detail;
                        return result;
                    }
                    result.AccountsSetToAll = true;
                }

                string policyKey = CachedModeSettings.GroupPolicyKeyPath(officeVersion);
                foreach (string name in plan)
                {
                    try
                    {
                        store.DeleteValue(policyKey, name);
                    }
                    catch (Exception ex) when (TuningReconciler.IsAccessDenied(ex))
                    {
                        result.Denied.Add(name);
                        continue;
                    }
                    catch (Exception ex) when (!(ex is OutOfMemoryException))
                    {
                        result.Failed.Add(name);
                        Note(result, ex);
                        continue;
                    }

                    result.Removed.Add(name);
                    // Its record goes with it, so nothing reads the value as the earlier OutlookAI's again.
                    TryDelete(store, TuningReconciler.AppliedKeyPath, CachedModeSettings.EarlierPolicyEntryIdPrefix + name, result);
                    TryDelete(store, TuningReconciler.DesiredKeyPath, CachedModeSettings.EarlierPolicyEntryIdPrefix + name, result);
                }

                if (result.Removed.Count > 0)
                {
                    try
                    {
                        store.WriteDword(TuningReconciler.TuningKeyPath, AddInServerContract.TuningRestartNeededValueName, 1);
                    }
                    catch (Exception ex) when (!(ex is OutOfMemoryException))
                    {
                        Note(result, ex);
                    }
                }
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                Note(result, ex);
            }
            return result;
        }

        private static void TryDelete(ITuningStore store, string keyPath, string valueName, EarlierPolicyRemovalResult result)
        {
            try
            {
                store.DeleteValue(keyPath, valueName);
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                Note(result, ex);
            }
        }

        private static void Note(EarlierPolicyRemovalResult result, Exception ex)
        {
            if (result.Problem.Length == 0)
                result.Problem = ex.GetType().Name + ": " + ex.Message;
        }
    }

    /// <summary>What one <see cref="EarlierPolicyRemoval.Remove"/> did.</summary>
    internal sealed class EarlierPolicyRemovalResult
    {
        /// <summary>Why nothing was removed; empty when the removal ran.</summary>
        public string Refusal { get; internal set; } = string.Empty;

        /// <summary>Every account's own period was set to All before anything was removed.</summary>
        public bool AccountsSetToAll { get; internal set; }

        /// <summary>The values removed, by name.</summary>
        public List<string> Removed { get; } = new List<string>();

        /// <summary>The values Windows would not let the user delete - written by an elevated Outlook.</summary>
        public List<string> Denied { get; } = new List<string>();

        /// <summary>The values whose deletion failed for any other reason.</summary>
        public List<string> Failed { get; } = new List<string>();

        /// <summary>The first thing that went wrong, for the debugger; empty when nothing did.</summary>
        public string Problem { get; internal set; } = string.Empty;
    }
}
