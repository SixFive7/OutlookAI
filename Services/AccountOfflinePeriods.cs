namespace OutlookAI.Services
{
    /// <summary>
    /// ===================================================================================
    /// THE SEAM: AN EXISTING ACCOUNT'S OWN OFFLINE PERIOD. NOT BUILT YET, ON PURPOSE.
    /// ===================================================================================
    ///
    /// <para>
    /// WHAT IT IS FOR. Outlook keeps the offline period ("Download email for the past") of an account
    /// that already exists in that account's section of the Outlook profile, and the user's own
    /// <c>SyncWindowSetting</c> reaches only accounts added later (KB 3115009). Q145 (a), decided by
    /// the maintainer: until it is measured, OutlookAI TELLS the user how to set it in Account
    /// Settings (<see cref="CachedModeSettings.ExistingAccountPeriodSteps"/>), and writing it waits
    /// for that measurement. Two callers are waiting on an implementation of this interface:
    /// </para>
    /// <list type="bullet">
    ///   <item>the settings dialog, which can then set every account to the chosen period instead of
    ///   telling the user how (<c>OutlookTuningService.CanSetExistingAccountPeriods</c>);</item>
    ///   <item><see cref="EarlierPolicyRemoval"/> (Q151 (a)), which must set every account's own
    ///   period to All BEFORE it removes an earlier OutlookAI's policy values, or Outlook drops the
    ///   mail outside each account's own, shorter period from the offline file when the policy goes
    ///   (KB 2733062: a lower window is "a local-only deletion of excess data").</item>
    /// </list>
    ///
    /// <para>
    /// WHAT IS KNOWN, AND WHAT IS NOT. Read on the maintainer's workstation, read-only, 2026-10-11
    /// (<c>.work\research-2026-10-11\q145-admin\</c>): each Exchange account's EMSMDB profile section,
    /// <c>HKCU\Software\Microsoft\Office\&lt;major&gt;\Outlook\Profiles\&lt;profile&gt;\&lt;section&gt;</c>,
    /// holds property 0x6649 (<c>00036649</c>, PR_RULE_ACTION_TYPE's tag) as the period in months - 0
    /// on three accounts, 12 on the fourth - which a Microsoft escalation engineer's MAPIToolkit
    /// (<c>CachedMode.cpp</c>) sets through the profile section. NOT known: where a day period is kept
    /// (0x665A is a candidate), whether Outlook honours a value written with Outlook closed, and
    /// whether a policy's value overrides it without changing it (measured here: it does, for 12 vs a
    /// policy of 0). The measurement belongs to the Exchange VM, the one test machine with an
    /// Exchange account (TODO.md).
    /// </para>
    ///
    /// <para>
    /// FRAMEWORK-NEUTRAL and INTERNAL, like <see cref="CachedModeSettings"/>, which the test project
    /// links beside it.
    /// </para>
    /// </summary>
    internal interface IAccountOfflinePeriods
    {
        /// <summary>Whether this build can set an existing account's own period. False until the measurement.</summary>
        bool CanSetEveryAccount { get; }

        /// <summary>Why it cannot, in a sentence; empty when it can.</summary>
        string WhyNot { get; }

        /// <summary>
        /// Sets the own period of every Exchange account in the current Outlook profile to
        /// <paramref name="period"/>, which is one of <see cref="CachedModeSettings.OfflinePeriods"/>.
        /// Never throws: a failure is the answer's.
        /// </summary>
        AccountPeriodChange SetEveryAccount(OfflinePeriod period);
    }

    /// <summary>What <see cref="IAccountOfflinePeriods.SetEveryAccount"/> did.</summary>
    internal sealed class AccountPeriodChange
    {
        private AccountPeriodChange(bool done, string detail)
        {
            Done = done;
            Detail = detail;
        }

        /// <summary>Every account now has the period asked for.</summary>
        internal bool Done { get; }

        /// <summary>What happened, or why nothing did.</summary>
        internal string Detail { get; }

        internal static AccountPeriodChange Succeeded(string detail)
        {
            return new AccountPeriodChange(true, detail ?? string.Empty);
        }

        internal static AccountPeriodChange NotDone(string why)
        {
            return new AccountPeriodChange(false, why ?? string.Empty);
        }
    }

    /// <summary>
    /// THIS BUILD'S ANSWER: it cannot. The one implementation there is until the Exchange VM
    /// measurement says how to write an account's own period; replace it then, in
    /// <c>OutlookTuningService.AccountPeriods</c>, and the dialog and the removal path follow.
    /// </summary>
    internal sealed class AccountOfflinePeriodsNotYetBuilt : IAccountOfflinePeriods
    {
        internal static readonly AccountOfflinePeriodsNotYetBuilt Instance = new AccountOfflinePeriodsNotYetBuilt();

        internal const string Reason =
            "OutlookAI cannot set an existing account's own offline period yet: where Outlook keeps it in "
            + "the profile has not been measured. Set it in Outlook's Account Settings.";

        private AccountOfflinePeriodsNotYetBuilt()
        {
        }

        public bool CanSetEveryAccount
        {
            get { return false; }
        }

        public string WhyNot
        {
            get { return Reason; }
        }

        public AccountPeriodChange SetEveryAccount(OfflinePeriod period)
        {
            return AccountPeriodChange.NotDone(Reason);
        }
    }
}
