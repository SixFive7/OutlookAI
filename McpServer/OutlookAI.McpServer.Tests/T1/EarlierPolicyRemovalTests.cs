using OutlookAI.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The removal of the Cached Mode policy values an earlier OutlookAI wrote,
/// <c>Services\EarlierPolicyRemoval.cs</c>, LINKED, against a fabricated registry.
///
/// <para>
/// Q151 (a), decided by the maintainer: the values OutlookAI 3.0.0.319 to 3.1.0.325 wrote as Group
/// Policy stay, labelled "Set by an earlier OutlookAI", until he asks for their removal - and when
/// that comes, every account's own offline period is set to All first, so no mail drops out of the
/// offline copy. So this pins three things: what counts as the earlier OutlookAI's (and nothing
/// else is ever removed), that the removal REFUSES, whole, while every account's own period cannot
/// be set to All first - which is this build, until the Exchange VM measurement - and that nothing
/// in the add-in calls it.
/// </para>
/// </summary>
public sealed class EarlierPolicyRemovalTests
{
    private const string Office = "16.0";
    private const string UserKey = @"Software\Microsoft\Office\16.0\Outlook\Cached Mode";
    private const string GroupPolicyKey = @"Software\Policies\Microsoft\Office\16.0\Outlook\Cached Mode";
    private const string CloudPolicyKey = @"Software\Policies\Microsoft\Cloud\Office\16.0\Outlook\Cached Mode";

    private static readonly string[] Names = CachedModeSettings.ValueNames.ToArray();
    private static readonly int[] EarlierValues = { 0, 0, 1, 1, 1 };

    /// <summary>A seam that can, recording what it was asked and the moment it was asked.</summary>
    private sealed class FakeAccounts : IAccountOfflinePeriods
    {
        private readonly FakeTuningStore _store;
        private readonly bool _succeeds;

        public FakeAccounts(FakeTuningStore store, bool succeeds = true)
        {
            _store = store;
            _succeeds = succeeds;
        }

        public List<OfflinePeriod> Asked { get; } = new();

        /// <summary>How many deletes had been attempted when the accounts were set.</summary>
        public int DeletesBeforeAsked { get; private set; } = -1;

        public bool CanSetEveryAccount => true;

        public string WhyNot => string.Empty;

        public AccountPeriodChange SetEveryAccount(OfflinePeriod period)
        {
            Asked.Add(period);
            DeletesBeforeAsked = _store.DeleteAttempts.Count;
            return _succeeds ? AccountPeriodChange.Succeeded("3 accounts set") : AccountPeriodChange.NotDone("the profile is locked");
        }
    }

    /// <summary>What 3.1.0.325 left where the user's token could write the policy key: all five, and their records.</summary>
    private static FakeTuningStore WithEarlierValues()
    {
        var store = new FakeTuningStore();
        for (int i = 0; i < Names.Length; i++)
        {
            store.Set(GroupPolicyKey, Names[i], EarlierValues[i]);
            store.Set(TuningReconciler.AppliedKeyPath, "caching.policy." + Names[i], EarlierValues[i]);
            store.Set(TuningReconciler.DesiredKeyPath, "caching.policy." + Names[i], EarlierValues[i]);
        }
        return store;
    }

    [Fact]
    public void ThePlan_IsEveryGroupPolicyValueStillHoldingWhatAnEarlierOutlookAIApplied()
    {
        FakeTuningStore store = WithEarlierValues();

        Assert.Equal(Names, EarlierPolicyRemoval.Plan(store, Office));
    }

    [Fact]
    public void ThePlan_LeavesOutAnythingTheOrganizationSetOrChanged()
    {
        FakeTuningStore store = WithEarlierValues();
        store.Set(GroupPolicyKey, "SyncWindowSetting", 12);       // changed since: the organization's now
        store.Unset(TuningReconciler.AppliedKeyPath, "caching.policy.CacheOthersMail"); // never recorded: not ours
        store.Set(CloudPolicyKey, "DownloadSharedFolders", 1);     // a Cloud Policy is never ours - but the GP value beneath it still is

        Assert.Equal(new[] { "SyncWindowSettingDays", "DownloadSharedFolders", "DisableSyncSliderForSharedMailbox" },
                     EarlierPolicyRemoval.Plan(store, Office));
    }

    [Fact]
    public void ThisBuild_RefusesWhole_BecauseItCannotSetEveryAccountsOwnPeriodToAllFirst()
    {
        FakeTuningStore store = WithEarlierValues();

        EarlierPolicyRemovalResult result = EarlierPolicyRemoval.Remove(store, Office, AccountOfflinePeriodsNotYetBuilt.Instance);

        Assert.Contains("every account's own period must be All", result.Refusal, StringComparison.Ordinal);
        Assert.Contains(AccountOfflinePeriodsNotYetBuilt.Reason, result.Refusal, StringComparison.Ordinal);
        Assert.Empty(result.Removed);
        Assert.False(result.AccountsSetToAll);
        Assert.Empty(store.DeleteAttempts);
        Assert.Empty(store.WriteAttempts);
        Assert.All(Names, n => Assert.True(store.Has(GroupPolicyKey, n)));
    }

    [Fact]
    public void WithAWayToSetTheAccounts_EveryAccountGoesToAllFirst_ThenTheValuesAndTheirRecordsGo()
    {
        FakeTuningStore store = WithEarlierValues();
        store.Set(UserKey, "SyncWindowSetting", 0);
        var accounts = new FakeAccounts(store);

        EarlierPolicyRemovalResult result = EarlierPolicyRemoval.Remove(store, Office, accounts);

        Assert.Equal(string.Empty, result.Refusal);
        Assert.Same(OfflinePeriods.All, Assert.Single(accounts.Asked));
        Assert.Equal(0, accounts.DeletesBeforeAsked);
        Assert.True(result.AccountsSetToAll);
        Assert.Equal(Names, result.Removed);
        Assert.Empty(result.Denied);
        Assert.Empty(result.Failed);
        Assert.Empty(store.AddressesUnder(GroupPolicyKey));
        Assert.All(Names, n =>
        {
            Assert.False(store.Has(TuningReconciler.AppliedKeyPath, "caching.policy." + n));
            Assert.False(store.Has(TuningReconciler.DesiredKeyPath, "caching.policy." + n));
        });
        Assert.Equal(1, store.ReadDword(TuningReconciler.TuningKeyPath, "RestartNeeded"));
        Assert.Equal(0, store.ReadDword(UserKey, "SyncWindowSetting")); // the user's own settings are never touched
    }

    [Fact]
    public void WhenSettingTheAccountsFails_NothingIsRemoved()
    {
        FakeTuningStore store = WithEarlierValues();

        EarlierPolicyRemovalResult result = EarlierPolicyRemoval.Remove(store, Office, new FakeAccounts(store, succeeds: false));

        Assert.Contains("the profile is locked", result.Refusal, StringComparison.Ordinal);
        Assert.Empty(store.DeleteAttempts);
        Assert.Empty(result.Removed);
    }

    [Fact]
    public void WithoutThePeriodAmongThem_TheAccountsAreNotTouched_AndTheRestGoes()
    {
        FakeTuningStore store = WithEarlierValues();
        store.Set(GroupPolicyKey, "SyncWindowSetting", 12);     // the organization's
        store.Set(GroupPolicyKey, "SyncWindowSettingDays", 0);
        store.Set(TuningReconciler.AppliedKeyPath, "caching.policy.SyncWindowSettingDays", 7);

        EarlierPolicyRemovalResult result = EarlierPolicyRemoval.Remove(store, Office, AccountOfflinePeriodsNotYetBuilt.Instance);

        Assert.Equal(string.Empty, result.Refusal);
        Assert.False(result.AccountsSetToAll);
        Assert.Equal(new[] { "DownloadSharedFolders", "CacheOthersMail", "DisableSyncSliderForSharedMailbox" }, result.Removed);
        Assert.Equal(12, store.ReadDword(GroupPolicyKey, "SyncWindowSetting"));
        Assert.Equal(0, store.ReadDword(GroupPolicyKey, "SyncWindowSettingDays"));
    }

    [Fact]
    public void AValueTheUserMayNotDelete_IsReported_NeverRetriedWithMoreRights()
    {
        // What an ELEVATED Outlook left: the key is the administrators', and Q145 (a) asks for none.
        FakeTuningStore store = WithEarlierValues();
        store.DeniedPrefixes.Add(@"Software\Policies");

        EarlierPolicyRemovalResult result = EarlierPolicyRemoval.Remove(store, Office, new FakeAccounts(store));

        Assert.Equal(Names, result.Denied);
        Assert.Empty(result.Removed);
        Assert.All(Names, n => Assert.True(store.Has(TuningReconciler.AppliedKeyPath, "caching.policy." + n), "a value still there keeps its record"));
        Assert.Null(store.ReadDword(TuningReconciler.TuningKeyPath, "RestartNeeded"));
    }

    [Fact]
    public void NothingToRemove_SaysSo()
    {
        var store = new FakeTuningStore();
        store.Set(GroupPolicyKey, "SyncWindowSetting", 12);

        EarlierPolicyRemovalResult result = EarlierPolicyRemoval.Remove(store, Office, new FakeAccounts(store));

        Assert.Contains("nothing to remove", result.Refusal, StringComparison.Ordinal);
        Assert.Empty(store.DeleteAttempts);
    }

    [Fact]
    public void NothingInTheAddInCallsTheRemoval()
    {
        // Q151 (a): built, and run nowhere. The one way in is OutlookTuningService's
        // RemoveEarlierPolicyValues, and nothing calls that - wiring it to a button or a startup is
        // the maintainer's request, and this test changes with it.
        string root = CachedModeSettingsTests.RepoRoot();
        var sources = Directory.GetFiles(root, "*.cs", SearchOption.TopDirectoryOnly)
            .Concat(Directory.GetFiles(Path.Combine(root, "Services"), "*.cs", SearchOption.AllDirectories))
            .Concat(Directory.GetFiles(Path.Combine(root, "TaskPane"), "*.cs", SearchOption.AllDirectories))
            .ToList();

        var callsOfRemove = sources.Where(f => File.ReadAllText(f).Contains("EarlierPolicyRemoval.Remove(", StringComparison.Ordinal)).Select(Path.GetFileName).ToList();
        Assert.Equal(new[] { "OutlookTuningService.cs" }, callsOfRemove);

        int mentions = sources.Sum(f => CountOf(File.ReadAllText(f), "RemoveEarlierPolicyValues("));
        Assert.Equal(1, mentions); // its own declaration
    }

    private static int CountOf(string text, string needle)
    {
        int count = 0;
        for (int at = text.IndexOf(needle, StringComparison.Ordinal); at >= 0; at = text.IndexOf(needle, at + needle.Length, StringComparison.Ordinal))
            count++;
        return count;
    }
}
