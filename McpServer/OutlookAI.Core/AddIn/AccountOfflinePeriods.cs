namespace OutlookAI.Core.AddIn
{
    /// <summary>
    /// ===================================================================================
    /// THE SEAM: AN EXISTING ACCOUNT'S OWN OFFLINE PERIOD. THE WRITER IS BUILT, IN CORE;
    /// THE ADD-IN CANNOT CALL IT YET.
    /// ===================================================================================
    ///
    /// <para>
    /// WHAT IT IS FOR. Outlook keeps the offline period ("Download email for the past") of an account
    /// that already exists in that account's section of the Outlook profile, and the user's own
    /// <c>SyncWindowSetting</c> reaches only accounts added later (KB 3115009). Q145 (a), decided by
    /// the maintainer: until OutlookAI can write it, it TELLS the user how to set it in Account
    /// Settings (<see cref="CachedModeSettings.ExistingAccountPeriodSteps"/>). Two callers are
    /// waiting on an implementation of this interface:
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
    /// WHAT IS KNOWN. MEASURED on the Exchange VM, 2026-11-17, Office LTSC 2024, the test account
    /// only (Testbed/guest/Measure-OfflineWindow.ps1): the period lives in the account's own Exchange
    /// provider section of the profile,
    /// <c>HKCU\Software\Microsoft\Office\&lt;major&gt;\Outlook\Profiles\&lt;profile&gt;\&lt;section&gt;</c>,
    /// the one that also carries the account's address, as property 0x6649 (months) and 0x665A (days),
    /// which Outlook always writes as a pair - one of the eleven (<see cref="OfflinePeriods"/>). A
    /// pair written with Outlook closed, and one written while it RAN, were both honoured at the next
    /// start, and Outlook did not overwrite either on exit; whether a running Outlook applies one
    /// before it restarts was not measured, and its own dialog asks for a restart. Read before that
    /// on the maintainer's workstation, read-only (2026-10-11): a policy period stands over an
    /// account's own without changing it.
    /// </para>
    ///
    /// <para>
    /// WHAT IS BUILT, AND WHAT WAITS. The writer is <c>OutlookAI.Core.Services.AccountOfflinePeriodWriter</c>
    /// - in Core, beside the profile's account reader, because logic both sides need lives there (Q154);
    /// T1 AccountOfflinePeriodWriterTests pins it. The add-in references Core since the code-sharing
    /// refactor's first stage (Q167 (a)), but may call only Core's namespace OutlookAI.Core.AddIn
    /// (Q164 (a)), so the writer moves there first. Then this interface gets its implementation:
    /// every Exchange account of the default profile (<c>ProfileAccountEntries</c>), each written
    /// with <c>AccountOfflinePeriodWriter.Write</c>, and the user told that the new period takes
    /// effect when Outlook next starts. Until then, Account Settings is the user's way.
    /// </para>
    ///
    /// <para>
    /// In OutlookAI.Core's add-in namespace, like <see cref="CachedModeSettings"/> beside it (Q163 (a)).
    /// </para>
    /// </summary>
    public interface IAccountOfflinePeriods
    {
        /// <summary>Whether this build can set an existing account's own period. False until the add-in can call Core's writer.</summary>
        bool CanSetEveryAccount { get; }

        /// <summary>Why it cannot, in a sentence; empty when it can.</summary>
        string WhyNot { get; }

        /// <summary>
        /// Sets the own period of every Exchange account in the current Outlook profile to
        /// <paramref name="period"/>, which is one of the eleven (<see cref="OfflinePeriods"/>).
        /// Never throws: a failure is the answer's.
        /// </summary>
        AccountPeriodChange SetEveryAccount(OfflinePeriod period);
    }

    /// <summary>What <see cref="IAccountOfflinePeriods.SetEveryAccount"/> did.</summary>
    public sealed class AccountPeriodChange
    {
        private AccountPeriodChange(bool done, string detail)
        {
            Done = done;
            Detail = detail;
        }

        /// <summary>Every account now has the period asked for.</summary>
        public bool Done { get; }

        /// <summary>What happened, or why nothing did.</summary>
        public string Detail { get; }

        public static AccountPeriodChange Succeeded(string detail)
        {
            return new AccountPeriodChange(true, detail ?? string.Empty);
        }

        public static AccountPeriodChange NotDone(string why)
        {
            return new AccountPeriodChange(false, why ?? string.Empty);
        }
    }

    /// <summary>
    /// THIS BUILD'S ANSWER: it cannot. The one implementation there is until the add-in can call
    /// Core's <c>AccountOfflinePeriodWriter</c>; replace it then, in
    /// <c>OutlookTuningService.AccountPeriods</c>, and the dialog and the removal path follow.
    /// </summary>
    public sealed class AccountOfflinePeriodsNotYetBuilt : IAccountOfflinePeriods
    {
        public static readonly AccountOfflinePeriodsNotYetBuilt Instance = new AccountOfflinePeriodsNotYetBuilt();

        public const string Reason =
            "OutlookAI cannot set an existing account's own offline period yet. Set it in Outlook's Account Settings.";

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
