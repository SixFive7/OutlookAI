using System.Security;

using OutlookAI.Core.AddIn;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The add-in's tuning reconcile, Core's <c>AddIn\TuningReconciler.cs</c> - the copy the add-in
/// loads - run against a fabricated registry.
///
/// <para>
/// WHAT IT PINS. (1) Q128, decided 2026-10-03: the walk used to sit in one try block, so the first
/// write that threw ended it - measured on a guest, where the first write under
/// <c>HKCU\Software\Policies</c> was refused and nothing after it was written. A refused value is now
/// skipped and listed in NeedsAdministrator, every other value is still applied, and the bookkeeping
/// - LastReconcileUtc last - is always written. (2) Q145 (a), decided by the maintainer: no
/// administrator, ever. The walk writes only the user's own settings and NEVER anything under
/// <c>HKCU\Software\Policies</c>; a value Microsoft 365 Cloud Policy or Group Policy sets - an
/// organization's, or one an earlier OutlookAI wrote - is left alone and listed in PolicyConflicts,
/// found by reading both policy keys rather than by seeing an applied value reverted. (3) Q143: a
/// desired number that is not one the product writes is never written.
/// </para>
///
/// <para>
/// The catalog below has the add-in's shape: four search values, the five Cached Mode entries
/// exactly as the add-in builds them (<see cref="CachedModeSettings.CatalogEntries"/>), and two OST
/// sizes, in that order. The add-in's own catalog lives in its <c>OutlookTuningService</c>, which
/// no test can load (it is the add-in's, net48/VSTO); what is pinned here is the walk.
/// </para>
/// </summary>
public sealed class TuningReconcilerTests
{
    private const string Office = "16.0";
    private const string SearchKey = @"Software\Microsoft\Office\16.0\Outlook\Search";
    private const string UserCachedModeKey = @"Software\Microsoft\Office\16.0\Outlook\Cached Mode";
    private const string GroupPolicyKey = @"Software\Policies\Microsoft\Office\16.0\Outlook\Cached Mode";
    private const string CloudPolicyKey = @"Software\Policies\Microsoft\Cloud\Office\16.0\Outlook\Cached Mode";
    private const string PstKey = @"Software\Microsoft\Office\16.0\Outlook\PST";
    private const string PoliciesRoot = @"Software\Policies";
    private static readonly DateTime Now = new DateTime(2026, 10, 3, 18, 30, 0, DateTimeKind.Utc);
    private const string NowText = "2026-10-03T18:30:00.0000000Z";

    private static readonly string[] CachingIds = CachedModeSettings.ValueNames.Select(n => "caching.user." + n).ToArray();
    private static readonly string[] PeriodIds = { "caching.user.SyncWindowSetting", "caching.user.SyncWindowSettingDays" };

    private static List<TuningEntry> Catalog()
    {
        var catalog = new List<TuningEntry>
        {
            new TuningEntry("search.DisableServerAssistedSearch", TuningReconciler.GroupSearch, SearchKey, AddInServerContract.DisableServerAssistedSearchValueName, 1),
            new TuningEntry("search.SearchResultsCap", TuningReconciler.GroupSearch, SearchKey, "SearchResultsCap", 0),
            new TuningEntry("search.IncludeDeletedItems", TuningReconciler.GroupSearch, SearchKey, "IncludeDeletedItems", 1),
            new TuningEntry("search.DefaultSearchScope", TuningReconciler.GroupSearch, SearchKey, "DefaultSearchScope", 2),
        };
        catalog.AddRange(CachedModeSettings.CatalogEntries(Office));
        catalog.Add(new TuningEntry("ost.MaxLargeFileSize", TuningReconciler.GroupOst, PstKey, "MaxLargeFileSize", 102400));
        catalog.Add(new TuningEntry("ost.WarnLargeFileSize", TuningReconciler.GroupOst, PstKey, "WarnLargeFileSize", 96256));
        return catalog;
    }

    /// <summary>What a filtered token meets: HKCU\Software\Policies is ReadKey for the user.</summary>
    private static FakeTuningStore NotElevated()
    {
        var store = new FakeTuningStore();
        store.DeniedPrefixes.Add(PoliciesRoot);
        return store;
    }

    private static void AssertNothingWrittenUnderPolicies(FakeTuningStore store) =>
        Assert.DoesNotContain(store.WriteAttempts, a => a.StartsWith(PoliciesRoot + "\\", StringComparison.OrdinalIgnoreCase));

    [Fact]
    public void NotElevated_EveryValueIsWritten_NoneUnderPolicies_AndNothingNeedsAnAdministrator()
    {
        FakeTuningStore store = NotElevated();

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Empty(result.NeedsAdministrator);
        Assert.Empty(result.PolicyConflicts);
        Assert.Empty(result.NotOffered);
        Assert.Empty(result.Failed);
        Assert.All(Catalog(), e => Assert.True(store.Has(e.KeyPath, e.ValueName), e.Id + " must be written"));
        Assert.Equal(0, store.ReadDword(UserCachedModeKey, "SyncWindowSetting"));
        Assert.Equal(0, store.ReadDword(UserCachedModeKey, "SyncWindowSettingDays"));
        Assert.Equal(1, store.ReadDword(UserCachedModeKey, "DownloadSharedFolders"));
        Assert.Equal(1, store.ReadDword(UserCachedModeKey, "CacheOthersMail"));
        Assert.Equal(1, store.ReadDword(UserCachedModeKey, "DisableSyncSliderForSharedMailbox"));
        Assert.Equal(102400, store.ReadDword(PstKey, "MaxLargeFileSize"));
        AssertNothingWrittenUnderPolicies(store);
    }

    [Fact]
    public void TheBookkeepingIsAlwaysWritten_LastReconcileUtcLast()
    {
        FakeTuningStore store = NotElevated();

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Equal(NowText, store.ReadString(TuningReconciler.TuningKeyPath, "LastReconcileUtc"));
        Assert.Equal(NowText, result.LastReconcileUtc);
        Assert.Equal(1, store.ReadDword(TuningReconciler.TuningKeyPath, "RestartNeeded"));
        Assert.True(result.RestartNeeded, "eleven values were written under a running Outlook");
        Assert.True(store.Has(TuningReconciler.TuningKeyPath, "NeedsAdministrator"), "NeedsAdministrator is written every time, empty when there are none");
        Assert.Equal(string.Empty, store.ReadString(TuningReconciler.TuningKeyPath, "NeedsAdministrator"));
        Assert.True(store.Has(TuningReconciler.TuningKeyPath, "PolicyConflicts"), "PolicyConflicts is written every time, empty when there are none");
        Assert.Equal(string.Empty, store.ReadString(TuningReconciler.TuningKeyPath, "PolicyConflicts"));

        // LastReconcileUtc is the LAST write of the reconcile: a reader that sees it may trust the rest.
        Assert.Equal(TuningReconciler.TuningKeyPath + "\\LastReconcileUtc", store.WriteAttempts.Last());
    }

    [Fact]
    public void ARefusedWrite_IsSkippedAndListed_AndEverythingAfterItIsStillApplied()
    {
        // Q128's walk, with what Windows may still refuse: an ACL on the user's own Cached Mode key.
        var store = new FakeTuningStore();
        store.DeniedPrefixes.Add(UserCachedModeKey);

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Equal(CachingIds, result.NeedsAdministrator);
        Assert.Empty(result.Failed);
        Assert.Equal(102400, store.ReadDword(PstKey, "MaxLargeFileSize"));
        Assert.Equal(96256, store.ReadDword(PstKey, "WarnLargeFileSize"));
        Assert.Equal(string.Join(";", CachingIds), store.ReadString(TuningReconciler.TuningKeyPath, "NeedsAdministrator"));
        Assert.Equal(NowText, store.ReadString(TuningReconciler.TuningKeyPath, "LastReconcileUtc"));

        var applied = Catalog().Where(e => store.Has(TuningReconciler.AppliedKeyPath, e.Id)).Select(e => e.Id).ToList();
        Assert.Equal(6, applied.Count);
        Assert.Empty(applied.Intersect(CachingIds));
    }

    [Fact]
    public void ASecurityException_IsTheOtherFaceOfAccessDenied()
    {
        var store = new FakeTuningStore();
        store.SecurityDeniedPrefixes.Add(UserCachedModeKey);

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: false, Now);

        Assert.Equal(CachingIds, result.NeedsAdministrator);
        Assert.Equal(NowText, store.ReadString(TuningReconciler.TuningKeyPath, "LastReconcileUtc"));
    }

    [Fact]
    public void AnyOtherWriteFailure_SkipsThatValue_IsNotCalledNeedsAdministrator_AndTheWalkGoesOn()
    {
        var store = new FakeTuningStore();
        store.BrokenPrefixes.Add(PstKey);

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Equal(new[] { "ost.MaxLargeFileSize", "ost.WarnLargeFileSize" }, result.Failed);
        Assert.Empty(result.NeedsAdministrator);
        Assert.Contains("IOException", result.Problem);
        Assert.Equal(0, store.ReadDword(UserCachedModeKey, "SyncWindowSetting"));
        Assert.Equal(NowText, store.ReadString(TuningReconciler.TuningKeyPath, "LastReconcileUtc"));
    }

    [Fact]
    public void TheFirstValuesOfTheCatalogRefused_StillLeaveTheOtherSevenWritten()
    {
        // The walk does not stop wherever the refusal lands.
        var store = new FakeTuningStore();
        store.DeniedPrefixes.Add(SearchKey);

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Equal(4, result.NeedsAdministrator.Count);
        Assert.Equal(7, Catalog().Count(e => store.Has(e.KeyPath, e.ValueName)));
        Assert.Equal(NowText, store.ReadString(TuningReconciler.TuningKeyPath, "LastReconcileUtc"));
    }

    [Fact]
    public void ABookkeepingWriteThatFails_DoesNotStopLastReconcileUtc()
    {
        var store = new FakeTuningStore();
        store.BrokenValueNames.Add("PolicyConflicts");
        store.BrokenValueNames.Add("NeedsAdministrator");

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Equal(NowText, store.ReadString(TuningReconciler.TuningKeyPath, "LastReconcileUtc"));
        Assert.Contains("IOException", result.Problem);
    }

    [Fact]
    public void SeedingTheAddInsOwnKeyFails_TheReconcileStillSaysWhenItRan()
    {
        var store = new FakeTuningStore();
        store.DeniedPrefixes.Add(TuningReconciler.DesiredKeyPath);

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Contains("UnauthorizedAccessException", result.Problem);
        Assert.Equal(NowText, store.ReadString(TuningReconciler.TuningKeyPath, "LastReconcileUtc"));
    }

    // ===== Q145 (a): a value a policy sets is left alone =====

    [Theory]
    [InlineData(GroupPolicyKey)]
    [InlineData(CloudPolicyKey)]
    public void AValueAPolicySets_IsNotWritten_AndIsListedAsAPolicyConflict(string policyKey)
    {
        FakeTuningStore store = NotElevated();
        store.Set(policyKey, "CacheOthersMail", 0); // the organization says: no shared mail offline

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Equal(new[] { "caching.user.CacheOthersMail" }, result.PolicyConflicts);
        Assert.False(store.Has(UserCachedModeKey, "CacheOthersMail"), "the user's own value under a policy is left as it was");
        Assert.False(store.Has(TuningReconciler.AppliedKeyPath, "caching.user.CacheOthersMail"));
        Assert.Equal(0, store.ReadDword(policyKey, "CacheOthersMail"));
        Assert.Equal("caching.user.CacheOthersMail", store.ReadString(TuningReconciler.TuningKeyPath, "PolicyConflicts"));
        // Everything else is still applied, and nothing at all is written under the Policies hive.
        Assert.Equal(1, store.ReadDword(UserCachedModeKey, "DownloadSharedFolders"));
        Assert.Equal(0, store.ReadDword(UserCachedModeKey, "SyncWindowSetting"));
        Assert.Empty(result.NeedsAdministrator);
        AssertNothingWrittenUnderPolicies(store);
    }

    [Theory]
    [InlineData(GroupPolicyKey, "SyncWindowSetting", 12)]
    [InlineData(GroupPolicyKey, "SyncWindowSettingDays", 7)]
    [InlineData(CloudPolicyKey, "SyncWindowSetting", 3)]
    [InlineData(CloudPolicyKey, "SyncWindowSettingDays", 14)]
    public void EitherHalfOfTheOfflinePeriodSetByAPolicy_LeavesTheWholePeriodAlone(string policyKey, string valueName, int value)
    {
        FakeTuningStore store = NotElevated();
        store.Set(policyKey, valueName, value);

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Equal(PeriodIds, result.PolicyConflicts);
        Assert.False(store.Has(UserCachedModeKey, "SyncWindowSetting"));
        Assert.False(store.Has(UserCachedModeKey, "SyncWindowSettingDays"));
        Assert.Equal(1, store.ReadDword(UserCachedModeKey, "DisableSyncSliderForSharedMailbox"));
        AssertNothingWrittenUnderPolicies(store);
    }

    [Fact]
    public void AnOrganizationsPolicySetBeforeOutlookAIEverRan_IsRecognisedAtOnce_NotOverwritten()
    {
        // The defect Q145 fixed: the back-off only worked for a value OutlookAI had applied and then
        // seen reverted, so an organization's value present from the start read as one to write.
        // Here the token CAN write the Policies hive - as on the maintainer's workstation - and still
        // nothing is written there.
        var store = new FakeTuningStore();
        store.Set(GroupPolicyKey, "SyncWindowSetting", 12);
        store.Set(GroupPolicyKey, "SyncWindowSettingDays", 0);

        TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Equal(12, store.ReadDword(GroupPolicyKey, "SyncWindowSetting"));
        AssertNothingWrittenUnderPolicies(store);
    }

    [Fact]
    public void AnEarlierOutlookAIsPolicyValues_AreLeftAlone_AndTheirRecordsKept()
    {
        // What 3.0.0.319 to 3.1.0.325 left on the maintainer's workstation: the five as Group Policy,
        // and Tuning\Applied\caching.policy.* recording them (Q151 (a): they stay).
        FakeTuningStore store = NotElevated();
        int[] values = { 0, 0, 1, 1, 1 };
        string[] names = CachedModeSettings.ValueNames.ToArray();
        for (int i = 0; i < names.Length; i++)
        {
            store.Set(GroupPolicyKey, names[i], values[i]);
            store.Set(TuningReconciler.AppliedKeyPath, "caching.policy." + names[i], values[i]);
        }

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Equal(CachingIds, result.PolicyConflicts);
        Assert.All(names, n => Assert.False(store.Has(UserCachedModeKey, n)));
        for (int i = 0; i < names.Length; i++)
        {
            Assert.Equal(values[i], store.ReadDword(GroupPolicyKey, names[i]));
            Assert.Equal(values[i], store.ReadDword(TuningReconciler.AppliedKeyPath, "caching.policy." + names[i]));
        }
        AssertNothingWrittenUnderPolicies(store);
        Assert.Empty(store.DeleteAttempts);
    }

    [Fact]
    public void OnceThePolicyGoes_TheNextReconcileWritesTheUsersOwnValue()
    {
        FakeTuningStore store = NotElevated();
        store.Set(GroupPolicyKey, "DisableSyncSliderForSharedMailbox", 0);
        TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);
        Assert.False(store.Has(UserCachedModeKey, "DisableSyncSliderForSharedMailbox"));

        store.Unset(GroupPolicyKey, "DisableSyncSliderForSharedMailbox"); // the organization withdrew it
        TuningReconcileResult second = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now.AddMinutes(5));

        Assert.Empty(second.PolicyConflicts);
        Assert.Equal(1, store.ReadDword(UserCachedModeKey, "DisableSyncSliderForSharedMailbox"));
        Assert.True(second.WroteAny);
    }

    [Fact]
    public void TheSnapshot_ReadsSetByPolicyNow_NotFromTheLastReconcile()
    {
        FakeTuningStore store = NotElevated();
        TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);
        store.Set(CloudPolicyKey, "DownloadSharedFolders", 0); // a Cloud Policy arrived since

        TuningSnapshot snap = TuningReconciler.ReadSnapshot(store, Catalog());

        Assert.Empty(snap.PolicyConflicts); // the last reconcile's list
        Assert.Equal(new[] { "DownloadSharedFolders" }, snap.Values.Where(v => v.SetByPolicy).Select(v => v.Entry.ValueName));
    }

    // ===== Q143: only values the product offers are written =====

    [Theory]
    [InlineData(18, 0)]
    [InlineData(12, 7)]
    [InlineData(0, 1)]
    [InlineData(120, 0)]
    public void AnOfflinePeriodThatIsNotOneOfTheEleven_IsNeverWritten(int months, int days)
    {
        var store = new FakeTuningStore();
        store.Set(TuningReconciler.DesiredKeyPath, "caching.user.SyncWindowSetting", months);
        store.Set(TuningReconciler.DesiredKeyPath, "caching.user.SyncWindowSettingDays", days);

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Equal(PeriodIds, result.NotOffered);
        Assert.False(store.Has(UserCachedModeKey, "SyncWindowSetting"));
        Assert.False(store.Has(UserCachedModeKey, "SyncWindowSettingDays"));
        Assert.Equal(1, store.ReadDword(UserCachedModeKey, "CacheOthersMail"));
    }

    [Theory]
    [InlineData(36, 0)]
    [InlineData(60, 0)]
    [InlineData(0, 14)]
    public void EveryOneOfTheElevenIsWrittenAsItsPair(int months, int days)
    {
        var store = new FakeTuningStore();
        store.Set(TuningReconciler.DesiredKeyPath, "caching.user.SyncWindowSetting", months);
        store.Set(TuningReconciler.DesiredKeyPath, "caching.user.SyncWindowSettingDays", days);

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Empty(result.NotOffered);
        Assert.Equal(months, store.ReadDword(UserCachedModeKey, "SyncWindowSetting"));
        Assert.Equal(days, store.ReadDword(UserCachedModeKey, "SyncWindowSettingDays"));
    }

    [Fact]
    public void ASwitchDesiredAsAnythingButOneOrZero_IsNeverWritten()
    {
        var store = new FakeTuningStore();
        store.Set(TuningReconciler.DesiredKeyPath, "caching.user.DownloadSharedFolders", 2);

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Equal(new[] { "caching.user.DownloadSharedFolders" }, result.NotOffered);
        Assert.False(store.Has(UserCachedModeKey, "DownloadSharedFolders"));
    }

    // ===== As it always was =====

    [Fact]
    public void AGroupThatIsOff_IsNeitherWrittenNorFlagged()
    {
        FakeTuningStore store = NotElevated();
        store.Set(TuningReconciler.TuningKeyPath, "Initialized", 1);
        store.Set(TuningReconciler.TuningKeyPath, "CachingEnabled", 0);
        store.Set(GroupPolicyKey, "SyncWindowSetting", 12);

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Empty(result.NeedsAdministrator);
        Assert.Empty(result.PolicyConflicts);
        Assert.False(store.Has(UserCachedModeKey, "SyncWindowSetting"));
        Assert.Equal(102400, store.ReadDword(PstKey, "MaxLargeFileSize"));
    }

    [Fact]
    public void AMidSessionReconcileNeverClearsRestartNeeded_AStartupThatWritesNothingDoes()
    {
        var store = new FakeTuningStore();
        TuningReconciler.Reconcile(store, Catalog(), isStartup: false, Now);
        Assert.Equal(1, store.ReadDword(TuningReconciler.TuningKeyPath, "RestartNeeded"));

        TuningReconciler.Reconcile(store, Catalog(), isStartup: false, Now);
        Assert.Equal(1, store.ReadDword(TuningReconciler.TuningKeyPath, "RestartNeeded"));

        TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);
        Assert.Equal(0, store.ReadDword(TuningReconciler.TuningKeyPath, "RestartNeeded"));
    }

    [Fact]
    public void AValueInEffect_IsRecordedAsApplied_AndNotWrittenAgain()
    {
        var store = new FakeTuningStore();
        TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);
        store.WriteAttempts.Clear();

        TuningReconcileResult second = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now.AddMinutes(5));

        Assert.False(second.WroteAny);
        Assert.False(second.RestartNeeded, "a startup that writes nothing clears the flag, as always");
        Assert.Equal(11, Catalog().Count(e => store.Has(TuningReconciler.AppliedKeyPath, e.Id)));
        Assert.DoesNotContain(store.WriteAttempts, a => Catalog().Any(e => a == e.KeyPath + "\\" + e.ValueName));
    }

    [Fact]
    public void AccessDenied_IsExactlyUnauthorizedAccessOrSecurity()
    {
        Assert.True(TuningReconciler.IsAccessDenied(new UnauthorizedAccessException()));
        Assert.True(TuningReconciler.IsAccessDenied(new SecurityException()));
        Assert.False(TuningReconciler.IsAccessDenied(new IOException()));
        Assert.False(TuningReconciler.IsAccessDenied(new InvalidOperationException()));
    }

    [Theory]
    [InlineData("", new string[0])]
    [InlineData("a", new[] { "a" })]
    [InlineData("a;b", new[] { "a", "b" })]
    [InlineData(";a;;b;", new[] { "a", "b" })]
    [InlineData("a;a;b", new[] { "a", "b" })]
    public void IdLists_DropEmptiesAndDuplicates(string raw, string[] expected)
    {
        Assert.Equal(expected, TuningReconciler.ParseIdList(raw));
        Assert.Equal(raw.Length == 0 ? string.Empty : string.Join(";", expected), TuningReconciler.JoinIdList(TuningReconciler.ParseIdList(raw)));
    }
}
