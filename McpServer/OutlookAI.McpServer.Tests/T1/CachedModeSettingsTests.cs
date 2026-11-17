using System.Reflection;
using System.Xml.Linq;

using OutlookAI.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The Cached Mode settings the add-in keeps applied - <c>Services\CachedModeSettings.cs</c>, LINKED -
/// and the guarantee that nothing in the add-in needs an administrator any more.
///
/// <para>
/// WHAT IT PINS. Q143: the offline period offers exactly the eleven windows Microsoft documents, each
/// as its (months, days) pair, and nothing else. Q145 (a): who sets a value is read from Microsoft 365
/// Cloud Policy, then Group Policy, then the user's own setting - in that order, because that is the
/// order Outlook reads them - and a policy value is shown, never changed. Q151 (a): a Group Policy
/// value an earlier OutlookAI wrote is told from an organization's by Tuning\Applied. And the defects
/// the Q145 research found in the Q128 code: Cloud Policy ignored; an organization's existing Group
/// Policy value not recognised as theirs; 36 and 60 months shown as "not a value OutlookAI offers";
/// an undocumented "days overrides months" label.
/// </para>
///
/// <para>
/// The packaging half reads the repository's own files, because the add-in is MSBuild-only and
/// nothing in this suite can build it: no second program ships, none asks for elevation.
/// </para>
/// </summary>
public sealed class CachedModeSettingsTests
{
    private const string Office = "16.0";
    private const string UserKey = @"Software\Microsoft\Office\16.0\Outlook\Cached Mode";
    private const string GroupPolicyKey = @"Software\Policies\Microsoft\Office\16.0\Outlook\Cached Mode";
    private const string CloudPolicyKey = @"Software\Policies\Microsoft\Cloud\Office\16.0\Outlook\Cached Mode";

    /// <summary>Q143's table, as the maintainer decided it: window, months, days - in the slider's order.</summary>
    private static readonly (string Label, int Months, int Days)[] Eleven =
    {
        ("3 days", 0, 3),
        ("1 week", 0, 7),
        ("2 weeks", 0, 14),
        ("1 month", 1, 0),
        ("3 months", 3, 0),
        ("6 months", 6, 0),
        ("1 year", 12, 0),
        ("2 years", 24, 0),
        ("3 years", 36, 0),
        ("5 years", 60, 0),
        ("All", 0, 0),
    };

    public static TheoryData<string, int, int> TheEleven
    {
        get
        {
            var data = new TheoryData<string, int, int>();
            foreach ((string label, int months, int days) in Eleven)
                data.Add(label, months, days);
            return data;
        }
    }

    [Fact]
    public void TheFiveNames_InCatalogOrder()
    {
        Assert.Equal(
            new[] { "SyncWindowSetting", "SyncWindowSettingDays", "DownloadSharedFolders", "CacheOthersMail", "DisableSyncSliderForSharedMailbox" },
            CachedModeSettings.ValueNames);
    }

    // ===== Q143: exactly eleven offline periods =====

    [Fact]
    public void TheOfflinePeriods_AreExactlyTheElevenMicrosoftDocuments_InTheSlidersOrder()
    {
        Assert.Equal(Eleven, CachedModeSettings.OfflinePeriods.Select(p => (p.Label, p.Months, p.Days)).ToArray());
        Assert.Equal(Eleven.Select(e => e.Label), CachedModeSettings.ChoiceLabels(CachedModeSettings.Setting(CachedModeSettings.OfflinePeriodKey)));
        Assert.Same(CachedModeSettings.All, CachedModeSettings.OfflinePeriods.Last());
    }

    [Theory]
    [MemberData(nameof(TheEleven))]
    public void EachOfTheEleven_IsOffered_AndNamedByOutlooksOwnWords(string label, int months, int days)
    {
        Assert.True(CachedModeSettings.IsOfferedPeriod(months, days));
        Assert.Equal(label, CachedModeSettings.DescribePeriod(months, days));
    }

    [Fact]
    public void NothingElseIsOffered_AcrossEveryPairUpToTenYearsAndAMonthOfDays()
    {
        var offered = new List<(int, int)>();
        for (int months = -1; months <= 120; months++)
        {
            for (int days = -1; days <= 31; days++)
            {
                if (CachedModeSettings.IsOfferedPeriod(months, days))
                    offered.Add((months, days));
            }
        }

        Assert.Equal(Eleven.Select(e => (e.Months, e.Days)).OrderBy(p => p).ToList(), offered.OrderBy(p => p).ToList());
    }

    [Theory]
    [InlineData(12, 7)]   // a months value with a days value: no source documents one
    [InlineData(18, 0)]   // KB 3115009: Outlook would silently make it 12
    [InlineData(0, 1)]
    [InlineData(0, 30)]
    [InlineData(48, 0)]
    [InlineData(120, 0)]  // no source has a 10-year option
    [InlineData(-1, 0)]
    [InlineData(int.MaxValue, 0)]
    public void AnyOtherPair_IsRefused(int months, int days)
    {
        Assert.False(CachedModeSettings.IsOfferedPeriod(months, days));
        Assert.Equal(-1, CachedModeSettings.IndexOfPeriod(months, days));
    }

    [Fact]
    public void AChoiceIsWrittenAsItsPair_AndAPositionOutsideTheListIsRefused()
    {
        CachedModeSetting period = CachedModeSettings.Setting(CachedModeSettings.OfflinePeriodKey);
        for (int i = 0; i < 11; i++)
        {
            OfflinePeriod p = CachedModeSettings.OfflinePeriods[i];
            Assert.Equal(
                new[]
                {
                    new KeyValuePair<string, int>("caching.user.SyncWindowSetting", p.Months),
                    new KeyValuePair<string, int>("caching.user.SyncWindowSettingDays", p.Days),
                },
                CachedModeSettings.DesiredFor(period, i));
        }
        Assert.Empty(CachedModeSettings.DesiredFor(period, -1));
        Assert.Empty(CachedModeSettings.DesiredFor(period, 11));
    }

    [Theory]
    [InlineData("DownloadSharedFolders")]
    [InlineData("CacheOthersMail")]
    [InlineData("DisableSyncSliderForSharedMailbox")]
    public void ASwitch_OffersOnAndOff_AndNothingElse(string name)
    {
        CachedModeSetting setting = CachedModeSettings.Setting(name);

        Assert.False(setting.IsPeriod);
        Assert.Equal(new[] { name }, setting.ValueNames);
        Assert.Equal(new[] { "On", "Off" }, CachedModeSettings.ChoiceLabels(setting));
        Assert.Equal(new[] { new KeyValuePair<string, int>("caching.user." + name, 1) }, CachedModeSettings.DesiredFor(setting, 0));
        Assert.Equal(new[] { new KeyValuePair<string, int>("caching.user." + name, 0) }, CachedModeSettings.DesiredFor(setting, 1));
        Assert.Empty(CachedModeSettings.DesiredFor(setting, 2));
        Assert.True(CachedModeSettings.IsOfferedSwitch(1));
        Assert.True(CachedModeSettings.IsOfferedSwitch(0));
        Assert.False(CachedModeSettings.IsOfferedSwitch(2));
        Assert.False(CachedModeSettings.IsOfferedSwitch(-1));
    }

    [Theory]
    [InlineData("SyncWindowSetting")]       // a half of the period is not a setting of its own
    [InlineData("syncwindowsetting")]
    [InlineData("offlineperiod")]
    [InlineData("")]
    public void ASettingKeyIsOrdinal_AndAnUnknownOneThrows(string key)
    {
        Assert.Throws<ArgumentException>(() => CachedModeSettings.Setting(key));
    }

    [Fact]
    public void TheDialogsFourRows_AndNoneOfThemClaimsDaysOverrideMonths()
    {
        Assert.Equal(
            new[] { "OfflinePeriod", "DownloadSharedFolders", "CacheOthersMail", "DisableSyncSliderForSharedMailbox" },
            CachedModeSettings.AllSettings.Select(s => s.Key));
        Assert.True(CachedModeSettings.AllSettings[0].IsPeriod);
        Assert.Equal(new[] { "SyncWindowSetting", "SyncWindowSettingDays" }, CachedModeSettings.AllSettings[0].ValueNames);
        // Microsoft documents no such rule; the Q128 dialog's "days overrides months" label is gone.
        Assert.All(CachedModeSettings.AllSettings, s =>
        {
            Assert.DoesNotContain("overrid", s.Caption + " " + s.Detail, StringComparison.OrdinalIgnoreCase);
            Assert.False(string.IsNullOrWhiteSpace(s.Caption));
            Assert.Contains(s.ValueNames[0], s.Detail, StringComparison.Ordinal);
        });
    }

    // ===== Where they live =====

    [Theory]
    [InlineData("16.0")]
    [InlineData("15.0")]
    [InlineData("17.0")]
    public void TheThreeKeys_UserGroupPolicyAndCloudPolicy(string version)
    {
        Assert.Equal(@"Software\Microsoft\Office\" + version + @"\Outlook\Cached Mode", CachedModeSettings.UserKeyPath(version));
        Assert.Equal(@"Software\Policies\Microsoft\Office\" + version + @"\Outlook\Cached Mode", CachedModeSettings.GroupPolicyKeyPath(version));
        Assert.Equal(@"Software\Policies\Microsoft\Cloud\Office\" + version + @"\Outlook\Cached Mode", CachedModeSettings.CloudPolicyKeyPath(version));
    }

    [Fact]
    public void TheCatalogsFiveEntries_WriteOnlyTheUsersKey_AndConsultBothPolicyKeys()
    {
        List<TuningEntry> entries = CachedModeSettings.CatalogEntries(Office);

        Assert.Equal(CachedModeSettings.ValueNames.Select(n => "caching.user." + n), entries.Select(e => e.Id));
        Assert.Equal(CachedModeSettings.ValueNames, entries.Select(e => e.ValueName));
        Assert.Equal(new[] { 0, 0, 1, 1, 1 }, entries.Select(e => e.DefaultDesired));
        Assert.All(entries, e =>
        {
            Assert.Equal(UserKey, e.KeyPath);
            Assert.Equal(TuningReconciler.GroupCaching, e.GroupId);
            Assert.DoesNotContain("Policies", e.KeyPath, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(new[] { CloudPolicyKey, GroupPolicyKey }, e.PolicyKeyPaths);
        });
        Assert.Equal(new[] { "SyncWindowSetting", "SyncWindowSettingDays" }, entries[0].PolicyValueNames);
        Assert.Equal(new[] { "SyncWindowSetting", "SyncWindowSettingDays" }, entries[1].PolicyValueNames);
        Assert.Equal(new[] { "DownloadSharedFolders" }, entries[2].PolicyValueNames);
        Assert.Equal(new[] { "CacheOthersMail" }, entries[3].PolicyValueNames);
        Assert.Equal(new[] { "DisableSyncSliderForSharedMailbox" }, entries[4].PolicyValueNames);
    }

    [Fact]
    public void TheCatalogsOfferedCheck_ReadsThePeriodAsAPair()
    {
        List<TuningEntry> entries = CachedModeSettings.CatalogEntries(Office);
        var store = new FakeTuningStore();

        Assert.True(entries[0].DesiredIsOffered(store), "nothing stored reads as the shipped default, All");
        store.Set(TuningReconciler.DesiredKeyPath, "caching.user.SyncWindowSetting", 36);
        Assert.True(entries[0].DesiredIsOffered(store));
        Assert.True(entries[1].DesiredIsOffered(store));
        store.Set(TuningReconciler.DesiredKeyPath, "caching.user.SyncWindowSettingDays", 7);
        Assert.False(entries[0].DesiredIsOffered(store), "36 months with 7 days is not one of the eleven");
        Assert.False(entries[1].DesiredIsOffered(store));
        store.Set(TuningReconciler.DesiredKeyPath, "caching.user.CacheOthersMail", 3);
        Assert.False(entries[3].DesiredIsOffered(store));
        Assert.True(entries[2].DesiredIsOffered(store));
    }

    // ===== Q145 (a): who sets a value =====

    [Theory]
    // cloud, group policy, user, earlier OutlookAI's Applied record -> effective, source
    [InlineData(null, null, null, null, null, "NotSet")]
    [InlineData(null, null, 0, null, 0, "UserSetting")]
    [InlineData(null, 12, 0, null, 12, "GroupPolicy")]
    [InlineData(3, 12, 0, null, 3, "CloudPolicy")]
    [InlineData(3, null, null, null, 3, "CloudPolicy")]
    [InlineData(null, 0, 12, 0, 0, "EarlierOutlookAI")]
    [InlineData(null, 12, 0, 0, 12, "GroupPolicy")]      // OutlookAI applied 0; the organization changed it to 12
    [InlineData(6, 0, null, 0, 6, "CloudPolicy")]        // an earlier OutlookAI's value under a Cloud Policy
    [InlineData(null, null, 0, 0, 0, "UserSetting")]     // an Applied record with no policy value is no policy
    public void CloudPolicyOutranksGroupPolicy_WhichOutranksTheUsersOwn(int? cloud, int? groupPolicy, int? user, int? applied, int? effective, string source)
    {
        var store = new FakeTuningStore();
        if (cloud.HasValue) store.Set(CloudPolicyKey, "SyncWindowSetting", cloud.Value);
        if (groupPolicy.HasValue) store.Set(GroupPolicyKey, "SyncWindowSetting", groupPolicy.Value);
        if (user.HasValue) store.Set(UserKey, "SyncWindowSetting", user.Value);
        if (applied.HasValue) store.Set(TuningReconciler.AppliedKeyPath, "caching.policy.SyncWindowSetting", applied.Value);

        CachedModeValueReading reading = CachedModeSettings.Read(store, Office, "SyncWindowSetting");

        Assert.Equal(effective, reading.Effective);
        Assert.Equal(Enum.Parse<CachedModeSource>(source), reading.Source);
        Assert.Equal(groupPolicy.HasValue && applied == groupPolicy, reading.GroupPolicyIsEarlierOutlookAI);
    }

    [Fact]
    public void AnOrganizationsGroupPolicy_PresentBeforeOutlookAIRan_IsTheirs()
    {
        // The Q128 code recognised a Group Policy value as the organization's only after it had
        // applied the value itself and seen it reverted; one present from the start read as
        // "Needs administrator" and "Apply as administrator" would have overwritten it.
        var store = new FakeTuningStore();
        store.Set(GroupPolicyKey, "SyncWindowSetting", 12);
        store.Set(GroupPolicyKey, "SyncWindowSettingDays", 0);

        CachedModeSettingState period = States(store).Single(s => s.Setting.IsPeriod);

        Assert.Equal(CachedModeSource.GroupPolicy, period.Source);
        Assert.Equal("Set by your organization", period.State);
        Assert.Equal("1 year", period.InEffect);
        Assert.False(period.Editable);
    }

    [Fact]
    public void ACloudPolicy_IsTheOrganizations_AndWhatItSaysIsWhatIsInEffect()
    {
        // The Q128 code read the Group Policy key alone, and called OutlookAI's value in effect
        // while Outlook used the Cloud Policy's.
        var store = new FakeTuningStore();
        store.Set(GroupPolicyKey, "DownloadSharedFolders", 1);
        store.Set(TuningReconciler.AppliedKeyPath, "caching.policy.DownloadSharedFolders", 1);
        store.Set(CloudPolicyKey, "DownloadSharedFolders", 0);

        CachedModeSettingState row = States(store).Single(s => s.Setting.Key == "DownloadSharedFolders");

        Assert.Equal(CachedModeSource.CloudPolicy, row.Source);
        Assert.Equal("Set by your organization", row.State);
        Assert.Equal("Off", row.InEffect);
        Assert.False(row.Editable);
    }

    [Fact]
    public void AnEarlierOutlookAIsPolicyValues_AreLabelledAsSuch_AndCannotBeChangedHere()
    {
        var store = new FakeTuningStore();
        int[] values = { 0, 0, 1, 1, 1 };
        string[] names = CachedModeSettings.ValueNames.ToArray();
        for (int i = 0; i < names.Length; i++)
        {
            store.Set(GroupPolicyKey, names[i], values[i]);
            store.Set(TuningReconciler.AppliedKeyPath, "caching.policy." + names[i], values[i]);
        }

        List<CachedModeSettingState> states = States(store);

        Assert.All(states, s =>
        {
            Assert.Equal(CachedModeSource.EarlierOutlookAI, s.Source);
            Assert.Equal("Set by an earlier OutlookAI", s.State);
            Assert.True(s.SetByPolicy);
            Assert.False(s.Editable);
        });
        Assert.Equal(new[] { "All", "On", "On", "On" }, states.Select(s => s.InEffect));
        Assert.Equal(CachedModeSettings.ExistingAccountPeriodSetByEarlierOutlookAI,
                     CachedModeSettings.ExistingAccountAdvice(states[0].Source));
    }

    [Theory]
    // months source, days source -> the pair's
    [InlineData("gp-org", "earlier", "GroupPolicy")]
    [InlineData("earlier", "earlier", "EarlierOutlookAI")]
    [InlineData("cloud", "user", "CloudPolicy")]
    [InlineData("user", "none", "UserSetting")]
    [InlineData("none", "earlier", "EarlierOutlookAI")]
    public void APairIsSetByTheStrongerOfItsHalves(string monthsBy, string daysBy, string expected)
    {
        var store = new FakeTuningStore();
        SetBy(store, "SyncWindowSetting", monthsBy, 0);
        SetBy(store, "SyncWindowSettingDays", daysBy, 0);

        Assert.Equal(Enum.Parse<CachedModeSource>(expected), States(store).Single(s => s.Setting.IsPeriod).Source);
    }

    [Fact]
    public void TheStateColumn_InTheOrderItIsDecided()
    {
        // Who sets it first: a policy holds whether OutlookAI manages the value or not.
        Assert.Equal("Set by your organization", CachedModeSettings.DescribeState(CachedModeSource.CloudPolicy, false, false, false, true));
        Assert.Equal("Set by your organization", CachedModeSettings.DescribeState(CachedModeSource.GroupPolicy, true, true, true, false));
        Assert.Equal("Set by an earlier OutlookAI", CachedModeSettings.DescribeState(CachedModeSource.EarlierOutlookAI, false, true, false, false));
        Assert.Equal("Not managed", CachedModeSettings.DescribeState(CachedModeSource.UserSetting, false, true, true, false));
        Assert.Equal("Not a value OutlookAI writes", CachedModeSettings.DescribeState(CachedModeSource.UserSetting, true, false, false, false));
        Assert.Equal("In effect", CachedModeSettings.DescribeState(CachedModeSource.UserSetting, true, true, true, true));
        Assert.Equal("Windows refused the change", CachedModeSettings.DescribeState(CachedModeSource.NotSet, true, true, false, true));
        Assert.Equal("Not applied yet", CachedModeSettings.DescribeState(CachedModeSource.NotSet, true, true, false, false));
    }

    [Fact]
    public void TheUsersOwnSettings_ReadAsInEffect_OnceTheyHoldWhatOutlookAISets()
    {
        var store = new FakeTuningStore();
        Assert.All(States(store), s => Assert.Equal("Not applied yet", s.State));

        store.Set(TuningReconciler.DesiredKeyPath, "caching.user.SyncWindowSetting", 60);
        store.Set(UserKey, "SyncWindowSetting", 60);
        store.Set(UserKey, "SyncWindowSettingDays", 0);
        store.Set(UserKey, "DownloadSharedFolders", 1);
        store.Set(UserKey, "CacheOthersMail", 1);
        store.Set(UserKey, "DisableSyncSliderForSharedMailbox", 1);
        List<CachedModeSettingState> states = States(store);

        Assert.All(states, s =>
        {
            Assert.Equal("In effect", s.State);
            Assert.Equal(CachedModeSource.UserSetting, s.Source);
            Assert.True(s.Editable);
        });
        Assert.Equal("5 years", states[0].InEffect); // the Q128 dialog said "60 (not a value OutlookAI offers)"
        Assert.Equal(9, states[0].DesiredIndex);
        Assert.Equal(CachedModeSettings.ExistingAccountPeriodSteps, CachedModeSettings.ExistingAccountAdvice(states[0].Source));
    }

    [Fact]
    public void ADesiredPeriodTypedInByHand_HasNoPlaceInTheList_AndSaysSo()
    {
        var store = new FakeTuningStore();
        store.Set(TuningReconciler.DesiredKeyPath, "caching.user.SyncWindowSetting", 18);

        CachedModeSettingState period = States(store).Single(s => s.Setting.IsPeriod);

        Assert.Equal(-1, period.DesiredIndex);
        Assert.Equal("Not a value OutlookAI writes", period.State);
    }

    [Fact]
    public void AValueWindowsRefused_ReadsAsRefused_UntilItIsInEffect()
    {
        var store = new FakeTuningStore();

        List<CachedModeSettingState> states = CachedModeSettings.ReadStates(store, Office, managed: true,
            new List<string> { "caching.user.SyncWindowSettingDays", "caching.user.CacheOthersMail" });

        Assert.Equal(new[] { "Windows refused the change", "Not applied yet", "Windows refused the change", "Not applied yet" },
                     states.Select(s => s.State));
    }

    [Fact]
    public void WithTheGroupOff_NothingIsEditable_ButWhoSetsAPolicyValueIsStillSaid()
    {
        var store = new FakeTuningStore();
        store.Set(CloudPolicyKey, "CacheOthersMail", 0);

        List<CachedModeSettingState> states = CachedModeSettings.ReadStates(store, Office, managed: false, new List<string>());

        Assert.All(states, s => Assert.False(s.Editable));
        Assert.Equal(new[] { "Not managed", "Not managed", "Set by your organization", "Not managed" }, states.Select(s => s.State));
    }

    [Theory]
    [InlineData(null, null, "(not set)")]
    [InlineData(0, 0, "All")]
    [InlineData(36, 0, "3 years")]
    [InlineData(60, 0, "5 years")]
    [InlineData(12, null, "1 year")]   // a months value alone, as Outlook 2013's template writes it
    [InlineData(0, null, "All")]
    [InlineData(null, 3, "months not set, days 3 (not one of Outlook's periods)")]
    [InlineData(18, 0, "months 18, days 0 (not one of Outlook's periods)")]
    [InlineData(12, 7, "months 12, days 7 (not one of Outlook's periods)")]
    public void APeriodInWords(int? months, int? days, string expected)
    {
        Assert.Equal(expected, CachedModeSettings.DescribePeriod(months, days));
    }

    [Theory]
    [InlineData(null, "(not set)")]
    [InlineData(1, "On")]
    [InlineData(0, "Off")]
    [InlineData(2, "2 (neither On nor Off)")]
    public void ASwitchInWords(int? value, string expected)
    {
        Assert.Equal(expected, CachedModeSettings.DescribeSwitch(value));
    }

    [Fact]
    public void AnExistingAccount_TheUserIsToldHowToSetItsOwnPeriod_InOutlooksWords()
    {
        string steps = CachedModeSettings.ExistingAccountAdvice(CachedModeSource.NotSet);

        Assert.Equal(steps, CachedModeSettings.ExistingAccountAdvice(CachedModeSource.UserSetting));
        Assert.Contains("File > Account Settings > Account Settings", steps, StringComparison.Ordinal);
        Assert.Contains("\"Download email for the past\"", steps, StringComparison.Ordinal);
        Assert.Contains("\"Mail to keep offline\"", steps, StringComparison.Ordinal);
        Assert.Contains("to All", steps, StringComparison.Ordinal);
        Assert.Contains("restart Outlook", steps, StringComparison.Ordinal);
        Assert.Equal(CachedModeSettings.ExistingAccountPeriodSetByOrganization, CachedModeSettings.ExistingAccountAdvice(CachedModeSource.CloudPolicy));
        Assert.Equal(CachedModeSettings.ExistingAccountPeriodSetByOrganization, CachedModeSettings.ExistingAccountAdvice(CachedModeSource.GroupPolicy));
    }

    [Fact]
    public void TheSeam_ThisBuildCannotSetAnExistingAccountsPeriod_AndSaysSo()
    {
        IAccountOfflinePeriods seam = AccountOfflinePeriodsNotYetBuilt.Instance;

        Assert.False(seam.CanSetEveryAccount);
        Assert.Contains("Account Settings", seam.WhyNot, StringComparison.Ordinal);
        AccountPeriodChange change = seam.SetEveryAccount(CachedModeSettings.All);
        Assert.False(change.Done);
        Assert.Equal(seam.WhyNot, change.Detail);
    }

    // ===== Q145 (a): nothing needs an administrator =====

    [Fact]
    public void TheSolution_HoldsTheAddInAlone_AndTheElevatedHelperIsGone()
    {
        XDocument solution = XDocument.Load(Path.Combine(RepoRoot(), "OutlookAI.slnx"));

        Assert.Equal(new[] { "OutlookAI.csproj" }, solution.Descendants("Project").Select(p => (string?)p.Attribute("Path")));
        // The project file, not the folder: a checkout that built the helper before keeps its bin and obj.
        Assert.False(File.Exists(Path.Combine(RepoRoot(), "PolicyWriter", "OutlookAI.PolicyWriter.csproj")), "PolicyWriter/ was the elevated helper");
    }

    [Fact]
    public void TheAddInBuildsNoOtherProject_AndShipsNoProgramBesideItself()
    {
        XDocument addIn = XDocument.Load(Path.Combine(RepoRoot(), "OutlookAI.csproj"));
        XNamespace ns = addIn.Root!.Name.Namespace;

        Assert.Empty(addIn.Descendants(ns + "ProjectReference"));
        Assert.DoesNotContain(addIn.Descendants(ns + "Content"),
            c => ((string?)c.Attribute("Include") ?? "").EndsWith(".exe", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void NoAddInSource_AsksWindowsForElevation()
    {
        string root = RepoRoot();
        var sources = Directory.GetFiles(root, "*.cs", SearchOption.TopDirectoryOnly)
            .Concat(Directory.GetFiles(Path.Combine(root, "Services"), "*.cs", SearchOption.AllDirectories))
            .Concat(Directory.GetFiles(Path.Combine(root, "TaskPane"), "*.cs", SearchOption.AllDirectories))
            .Concat(Directory.GetFiles(Path.Combine(root, "Properties"), "*", SearchOption.AllDirectories))
            .ToList();
        Assert.Contains(sources, s => s.EndsWith("OutlookTuningService.cs", StringComparison.Ordinal));

        foreach (string file in sources)
        {
            string text = File.ReadAllText(file);
            Assert.False(text.Contains("\"runas\"", StringComparison.OrdinalIgnoreCase), file + " asks for the runas verb");
            Assert.False(text.Contains("requireAdministrator", StringComparison.Ordinal), file + " requires an administrator");
            Assert.False(text.Contains("ShellExecuteEx", StringComparison.Ordinal), file + " calls ShellExecuteEx");
        }
    }

    [Fact]
    public void NothingWritesUnderThePoliciesHive_OnlyReadsIt()
    {
        // Every path built under a policy root is handed to a read: CachedModeSettings builds the
        // two policy keys only for the catalog's PolicyKeyPaths and for Read, and the one delete
        // there is the removal nothing calls (EarlierPolicyRemoval).
        string source = File.ReadAllText(Path.Combine(RepoRoot(), "Services", "OutlookTuningService.cs"));

        Assert.DoesNotContain("PolicyOutlookKeyPath", source, StringComparison.Ordinal);
        Assert.DoesNotContain(@"Software\Policies", source, StringComparison.Ordinal);
    }

    private static List<CachedModeSettingState> States(FakeTuningStore store) =>
        CachedModeSettings.ReadStates(store, Office, managed: true, new List<string>());

    private static void SetBy(FakeTuningStore store, string name, string by, int value)
    {
        switch (by)
        {
            case "cloud":
                store.Set(CloudPolicyKey, name, value);
                break;
            case "gp-org":
                store.Set(GroupPolicyKey, name, value);
                break;
            case "earlier":
                store.Set(GroupPolicyKey, name, value);
                store.Set(TuningReconciler.AppliedKeyPath, "caching.policy." + name, value);
                break;
            case "user":
                store.Set(UserKey, name, value);
                break;
        }
    }

    internal static string RepoRoot()
    {
        string testProjectDir =
            typeof(CachedModeSettingsTests).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
        return Path.GetFullPath(Path.Combine(testProjectDir, "..", ".."));
    }
}
