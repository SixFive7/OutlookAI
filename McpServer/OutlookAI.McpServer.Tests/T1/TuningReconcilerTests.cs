using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;

using OutlookAI.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The add-in's tuning reconcile, <c>Services\TuningReconciler.cs</c>, LINKED into this assembly and
/// run against a fabricated registry.
///
/// <para>
/// WHAT IT PINS (Q128, decided 2026-10-03). The walk used to sit in one try block, so the first
/// write that threw ended it - and in the Outlook a user actually runs, NOT elevated, the first
/// write under <c>HKCU\Software\Policies</c> always throws: the user may read that key, not write
/// it. Measured on a guest: the two user Cached Mode values, both OST sizes, RestartNeeded,
/// PolicyConflicts and LastReconcileUtc were never written, and <c>outlook_health</c> reported
/// <c>tuning.lastReconcileUtc</c> null (Docs/live-tier-on-the-vm.md section 2.3). Now a refused
/// value is skipped and recorded as needing an administrator, every other value is still applied,
/// and the bookkeeping - LastReconcileUtc last - is always written.
/// </para>
///
/// <para>
/// The catalog below has the add-in's shape: four search values, D25's five policy values (from
/// <see cref="CachedModePolicy"/>, as the add-in builds them), two user Cached Mode values and two
/// OST sizes, in that order. The add-in's own catalog lives in <c>OutlookTuningService</c>, which is
/// net48/VSTO and cannot be linked; what is pinned here is the walk.
/// </para>
/// </summary>
public sealed class TuningReconcilerTests
{
    private const string Office = "16.0";
    private const string SearchKey = @"Software\Microsoft\Office\16.0\Outlook\Search";
    private const string UserCachedModeKey = @"Software\Microsoft\Office\16.0\Outlook\Cached Mode";
    private const string PstKey = @"Software\Microsoft\Office\16.0\Outlook\PST";
    private const string PoliciesRoot = @"Software\Policies";
    private static readonly string PolicyKey = CachedModePolicy.KeyPath(Office);
    private static readonly DateTime Now = new DateTime(2026, 10, 3, 18, 30, 0, DateTimeKind.Utc);
    private const string NowText = "2026-10-03T18:30:00.0000000Z";

    private static readonly string[] PolicyIds = CachedModePolicy.ValueNames.Select(n => "caching.policy." + n).ToArray();

    private static List<TuningEntry> Catalog()
    {
        var catalog = new List<TuningEntry>
        {
            new TuningEntry("search.DisableServerAssistedSearch", TuningReconciler.GroupSearch, SearchKey, AddInServerContract.DisableServerAssistedSearchValueName, 1, false),
            new TuningEntry("search.SearchResultsCap", TuningReconciler.GroupSearch, SearchKey, "SearchResultsCap", 0, false),
            new TuningEntry("search.IncludeDeletedItems", TuningReconciler.GroupSearch, SearchKey, "IncludeDeletedItems", 1, false),
            new TuningEntry("search.DefaultSearchScope", TuningReconciler.GroupSearch, SearchKey, "DefaultSearchScope", 2, false),
        };
        foreach (string name in CachedModePolicy.ValueNames)
            catalog.Add(new TuningEntry("caching.policy." + name, TuningReconciler.GroupCaching, PolicyKey, name, CachedModePolicy.ShippedDefault(name), true));
        catalog.Add(new TuningEntry("caching.user.SyncWindowSetting", TuningReconciler.GroupCaching, UserCachedModeKey, "SyncWindowSetting", 0, false));
        catalog.Add(new TuningEntry("caching.user.SyncWindowSettingDays", TuningReconciler.GroupCaching, UserCachedModeKey, "SyncWindowSettingDays", 0, false));
        catalog.Add(new TuningEntry("ost.MaxLargeFileSize", TuningReconciler.GroupOst, PstKey, "MaxLargeFileSize", 102400, false));
        catalog.Add(new TuningEntry("ost.WarnLargeFileSize", TuningReconciler.GroupOst, PstKey, "WarnLargeFileSize", 96256, false));
        return catalog;
    }

    /// <summary>
    /// A registry in a dictionary. Writes under a denied prefix throw what .NET throws for a key the
    /// token may not write; reads never throw, like the add-in's store.
    /// </summary>
    private sealed class FakeStore : ITuningStore
    {
        private readonly Dictionary<string, object> _values = new(StringComparer.OrdinalIgnoreCase);

        public List<string> DeniedPrefixes { get; } = new();

        public List<string> SecurityDeniedPrefixes { get; } = new();

        public List<string> BrokenPrefixes { get; } = new();

        public HashSet<string> BrokenValueNames { get; } = new(StringComparer.Ordinal);

        public List<string> WriteAttempts { get; } = new();

        private static string Address(string keyPath, string valueName) => keyPath + "\\" + valueName;

        public void Set(string keyPath, string valueName, object value) => _values[Address(keyPath, valueName)] = value;

        public object? Get(string keyPath, string valueName) => _values.TryGetValue(Address(keyPath, valueName), out object? v) ? v : null;

        public bool Has(string keyPath, string valueName) => _values.ContainsKey(Address(keyPath, valueName));

        public int? ReadDword(string keyPath, string valueName) => Get(keyPath, valueName) is int i ? i : null;

        public string ReadString(string keyPath, string valueName) => Get(keyPath, valueName) as string ?? string.Empty;

        public void WriteDword(string keyPath, string valueName, int value) => Write(keyPath, valueName, value);

        public void WriteString(string keyPath, string valueName, string value) => Write(keyPath, valueName, value);

        private void Write(string keyPath, string valueName, object value)
        {
            WriteAttempts.Add(Address(keyPath, valueName));
            if (DeniedPrefixes.Any(p => keyPath.StartsWith(p, StringComparison.OrdinalIgnoreCase)))
                throw new UnauthorizedAccessException("Access to the registry key 'HKEY_CURRENT_USER\\" + keyPath + "' is denied.");
            if (SecurityDeniedPrefixes.Any(p => keyPath.StartsWith(p, StringComparison.OrdinalIgnoreCase)))
                throw new SecurityException("Requested registry access is not allowed.");
            if (BrokenPrefixes.Any(p => keyPath.StartsWith(p, StringComparison.OrdinalIgnoreCase)) || BrokenValueNames.Contains(valueName))
                throw new IOException("The device is not ready.");
            _values[Address(keyPath, valueName)] = value;
        }
    }

    private static FakeStore NotElevated()
    {
        var store = new FakeStore();
        // What a filtered token meets: HKCU\Software\Policies is ReadKey for the user.
        store.DeniedPrefixes.Add(PoliciesRoot);
        return store;
    }

    [Fact]
    public void NotElevated_ThePolicyValuesAreSkipped_AndEverythingAfterThemIsStillApplied()
    {
        FakeStore store = NotElevated();

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        // The five refused writes, named, in catalog order.
        Assert.Equal(PolicyIds, result.NeedsAdministrator);
        Assert.Empty(result.Failed);
        Assert.Empty(result.PolicyConflicts);

        // Every value AFTER them in the catalog - the defect left all four unwritten.
        Assert.Equal(0, store.ReadDword(UserCachedModeKey, "SyncWindowSetting"));
        Assert.Equal(0, store.ReadDword(UserCachedModeKey, "SyncWindowSettingDays"));
        Assert.Equal(102400, store.ReadDword(PstKey, "MaxLargeFileSize"));
        Assert.Equal(96256, store.ReadDword(PstKey, "WarnLargeFileSize"));
        // And the four before them.
        Assert.Equal(1, store.ReadDword(SearchKey, AddInServerContract.DisableServerAssistedSearchValueName));
        Assert.Equal(2, store.ReadDword(SearchKey, "DefaultSearchScope"));

        // Nothing at all under the Policies hive.
        foreach (string name in CachedModePolicy.ValueNames)
            Assert.False(store.Has(PolicyKey, name), name + " must not appear written");
    }

    [Fact]
    public void NotElevated_TheBookkeepingIsAlwaysWritten_LastReconcileUtcIncluded()
    {
        FakeStore store = NotElevated();

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Equal(NowText, store.ReadString(TuningReconciler.TuningKeyPath, "LastReconcileUtc"));
        Assert.Equal(NowText, result.LastReconcileUtc);
        Assert.Equal(1, store.ReadDword(TuningReconciler.TuningKeyPath, "RestartNeeded"));
        Assert.True(result.RestartNeeded, "eight values were written under a running Outlook");
        Assert.Equal(string.Join(";", PolicyIds), store.ReadString(TuningReconciler.TuningKeyPath, "NeedsAdministrator"));
        Assert.True(store.Has(TuningReconciler.TuningKeyPath, "PolicyConflicts"), "PolicyConflicts is written every time, empty when there are none");
        Assert.Equal(string.Empty, store.ReadString(TuningReconciler.TuningKeyPath, "PolicyConflicts"));

        // LastReconcileUtc is the LAST write of the reconcile: a reader that sees it may trust the rest.
        Assert.Equal(TuningReconciler.TuningKeyPath + "\\LastReconcileUtc", store.WriteAttempts.Last());
    }

    [Fact]
    public void NotElevated_AppliedRecordsEightValues_NeverThePolicyOnes()
    {
        FakeStore store = NotElevated();

        TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        var applied = Catalog().Where(e => store.Has(TuningReconciler.AppliedKeyPath, e.Id)).Select(e => e.Id).ToList();
        Assert.Equal(8, applied.Count);
        Assert.Empty(applied.Intersect(PolicyIds));
    }

    [Fact]
    public void ASecurityException_IsTheOtherFaceOfAccessDenied()
    {
        var store = new FakeStore();
        store.SecurityDeniedPrefixes.Add(PoliciesRoot);

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: false, Now);

        Assert.Equal(PolicyIds, result.NeedsAdministrator);
        Assert.Equal(NowText, store.ReadString(TuningReconciler.TuningKeyPath, "LastReconcileUtc"));
    }

    [Fact]
    public void AnyOtherWriteFailure_SkipsThatValue_IsNotCalledNeedsAdministrator_AndTheWalkGoesOn()
    {
        var store = new FakeStore();
        store.BrokenPrefixes.Add(PstKey);

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Equal(new[] { "ost.MaxLargeFileSize", "ost.WarnLargeFileSize" }, result.Failed);
        Assert.Empty(result.NeedsAdministrator);
        Assert.Contains("IOException", result.Problem);
        Assert.Equal(0, store.ReadDword(PolicyKey, CachedModePolicy.SyncWindowSetting));
        Assert.Equal(NowText, store.ReadString(TuningReconciler.TuningKeyPath, "LastReconcileUtc"));
    }

    [Fact]
    public void TheFirstValuesOfTheCatalogRefused_StillLeaveTheOtherNineWritten()
    {
        // The walk does not stop wherever the refusal lands, not only at the policy values.
        var store = new FakeStore();
        store.DeniedPrefixes.Add(SearchKey);

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Equal(4, result.NeedsAdministrator.Count);
        Assert.Equal(9, Catalog().Count(e => store.Has(e.KeyPath, e.ValueName)));
        Assert.Equal(NowText, store.ReadString(TuningReconciler.TuningKeyPath, "LastReconcileUtc"));
    }

    [Fact]
    public void ABookkeepingWriteThatFails_DoesNotStopLastReconcileUtc()
    {
        var store = new FakeStore();
        store.BrokenValueNames.Add("PolicyConflicts");
        store.BrokenValueNames.Add("NeedsAdministrator");

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Equal(NowText, store.ReadString(TuningReconciler.TuningKeyPath, "LastReconcileUtc"));
        Assert.Contains("IOException", result.Problem);
    }

    [Fact]
    public void SeedingTheAddInsOwnKeyFails_TheReconcileStillSaysWhenItRan()
    {
        var store = new FakeStore();
        store.DeniedPrefixes.Add(TuningReconciler.DesiredKeyPath);

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Contains("UnauthorizedAccessException", result.Problem);
        Assert.Equal(NowText, store.ReadString(TuningReconciler.TuningKeyPath, "LastReconcileUtc"));
    }

    [Fact]
    public void OnceAnAdministratorHasWrittenThem_TheNextReconcileFindsThemInEffect_AndClearsTheList()
    {
        FakeStore store = NotElevated();
        TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        // What OutlookAI.PolicyWriter.exe does, elevated, from outside this token's reach.
        foreach (string name in CachedModePolicy.ValueNames)
            store.Set(PolicyKey, name, CachedModePolicy.ShippedDefault(name));
        store.WriteAttempts.Clear();

        TuningReconcileResult second = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now.AddMinutes(5));

        Assert.Empty(second.NeedsAdministrator);
        Assert.Equal(string.Empty, store.ReadString(TuningReconciler.TuningKeyPath, "NeedsAdministrator"));
        Assert.False(second.WroteAny);
        Assert.False(second.RestartNeeded, "a startup that writes nothing clears the flag, as always");
        Assert.Equal(13, Catalog().Count(e => store.Has(TuningReconciler.AppliedKeyPath, e.Id)));
        // No write was even attempted under the Policies hive: the values were in effect.
        Assert.DoesNotContain(store.WriteAttempts, a => a.StartsWith(PoliciesRoot, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void AGroupThatIsOff_IsNeitherWrittenNorFlagged()
    {
        FakeStore store = NotElevated();
        store.Set(TuningReconciler.TuningKeyPath, "Initialized", 1);
        store.Set(TuningReconciler.TuningKeyPath, "CachingEnabled", 0);

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Empty(result.NeedsAdministrator);
        Assert.False(store.Has(UserCachedModeKey, "SyncWindowSetting"));
        Assert.Equal(102400, store.ReadDword(PstKey, "MaxLargeFileSize"));
    }

    [Fact]
    public void APolicyValueRevertedAfterItWasApplied_IsAConflict_NotAWriteAndNotNeedsAdministrator()
    {
        FakeStore store = NotElevated();
        string id = "caching.policy." + CachedModePolicy.SyncWindowSetting;
        store.Set(TuningReconciler.AppliedKeyPath, id, 0);
        store.Set(PolicyKey, CachedModePolicy.SyncWindowSetting, 12); // a GPO said 12 months

        TuningReconcileResult result = TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);

        Assert.Equal(new[] { id }, result.PolicyConflicts);
        Assert.DoesNotContain(id, result.NeedsAdministrator);
        Assert.DoesNotContain(PolicyKey + "\\" + CachedModePolicy.SyncWindowSetting, store.WriteAttempts);
        Assert.Equal(id, store.ReadString(TuningReconciler.TuningKeyPath, "PolicyConflicts"));
    }

    [Fact]
    public void AMidSessionReconcileNeverClearsRestartNeeded_AStartupThatWritesNothingDoes()
    {
        var store = new FakeStore();
        TuningReconciler.Reconcile(store, Catalog(), isStartup: false, Now);
        Assert.Equal(1, store.ReadDword(TuningReconciler.TuningKeyPath, "RestartNeeded"));

        TuningReconciler.Reconcile(store, Catalog(), isStartup: false, Now);
        Assert.Equal(1, store.ReadDword(TuningReconciler.TuningKeyPath, "RestartNeeded"));

        TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);
        Assert.Equal(0, store.ReadDword(TuningReconciler.TuningKeyPath, "RestartNeeded"));
    }

    [Fact]
    public void TheSnapshot_SaysNeedsAdministratorOnlyWhileTheValueIsStillNotInEffect()
    {
        FakeStore store = NotElevated();
        TuningReconciler.Reconcile(store, Catalog(), isStartup: true, Now);
        // One of the five set since - by an administrator, a GPO or the helper.
        store.Set(PolicyKey, CachedModePolicy.CacheOthersMail, 1);

        TuningSnapshot snap = TuningReconciler.ReadSnapshot(store, Catalog());

        var flagged = snap.Values.Where(v => v.NeedsAdministrator).Select(v => v.Entry.ValueName).ToList();
        Assert.Equal(CachedModePolicy.ValueNames.Where(n => n != CachedModePolicy.CacheOthersMail), flagged);
        Assert.Equal(PolicyIds, snap.NeedsAdministrator);
        Assert.True(snap.Values.Single(v => v.Entry.ValueName == CachedModePolicy.CacheOthersMail && v.Entry.IsPolicyHive).InSync);
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
