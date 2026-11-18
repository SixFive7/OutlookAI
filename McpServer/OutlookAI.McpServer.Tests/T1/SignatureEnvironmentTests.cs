using OutlookAI.Core.Services;
using Xunit;
using SignatureFolderLocation = global::OutlookAI.Services.SignatureFolderLocation;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The signature read model (Q147): where the files are (G3's rule), every profile and its accounts
/// (no merge), the administrator's rules and their precedence (Q156), and whether signatures can
/// roam (Q152, Q153). Fakes only - the real registry is never read.
/// </summary>
public sealed class SignatureEnvironmentTests
{
    private const string AppData = @"C:\Users\someone\AppData\Roaming";

    /// <summary>A folder that does not exist: no signatures, and nothing of the machine's own read.</summary>
    private static readonly string NoFolder = Path.Combine(Path.GetTempPath(), "OutlookAI-NoSignatures-" + Guid.NewGuid().ToString("N"));

    // ------------------------------------------------------------------ the folder

    [Theory]
    [InlineData(null, @"C:\Users\someone\AppData\Roaming\Microsoft\Signatures", false)]
    [InlineData("", @"C:\Users\someone\AppData\Roaming\Microsoft\Signatures", false)]
    [InlineData("   ", @"C:\Users\someone\AppData\Roaming\Microsoft\Signatures", false)]
    [InlineData("Handtekeningen", @"C:\Users\someone\AppData\Roaming\Microsoft\Handtekeningen", true)]  // a Dutch Outlook
    [InlineData("Signaturen", @"C:\Users\someone\AppData\Roaming\Microsoft\Signaturen", true)]          // a German one
    [InlineData(@"Corp\Signatures", @"C:\Users\someone\AppData\Roaming\Microsoft\Corp\Signatures", true)]
    [InlineData(@"D:\Shared\Signatures", @"D:\Shared\Signatures", true)]                               // absolute stands
    [InlineData("Signatures\0", @"C:\Users\someone\AppData\Roaming\Microsoft\Signatures", true)]
    public void TheFolder_IsTheRecordedValue_RelativeToAppDataMicrosoft_OrTheDefault(string? recorded, string expected, bool fromRegistry)
    {
        string path = SignatureFolderLocation.Resolve(AppData, recorded!, out bool fromRecorded);

        Assert.Equal(expected, path, ignoreCase: true);
        Assert.Equal(fromRegistry, fromRecorded);
    }

    [Fact]
    public void TheFolder_IsReadFromCommonGeneralSignatures_OfTheOfficeMajor()
    {
        SignatureFolderInfo folder = SignatureFolder.Read(new FakeSignatureSettings().FolderValue("Handtekeningen"), "16.0", AppData);

        Assert.True(folder.FromRegistry);
        Assert.Equal("Handtekeningen", folder.RecordedValue);
        Assert.EndsWith(@"\Microsoft\Handtekeningen", folder.Path, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(@"Software\Microsoft\Office\16.0\Common\General", SignatureFolderLocation.GeneralKeyPath("16.0"));

        SignatureFolderInfo fallback = SignatureFolder.Read(new FakeSignatureSettings(), "16.0", AppData);
        Assert.False(fallback.FromRegistry);
        Assert.Null(fallback.RecordedValue);
        Assert.EndsWith(@"\Microsoft\Signatures", fallback.Path, StringComparison.OrdinalIgnoreCase);
    }

    // ------------------------------------------------------------------ policy (Q156)

    [Fact]
    public void Policy_None_LeavesTheDefaultsToTheAccounts()
    {
        SignaturePolicy policy = SignaturePolicy.Read(new FakeSignatureSettings(), "16.0");

        Assert.False(policy.SignaturesDisabled);
        Assert.Null(policy.NewMessage);
        Assert.Null(policy.ReplyForward);
        Assert.False(policy.AnyDefaultSet);
    }

    [Theory]
    [InlineData(SignaturePolicySource.GroupPolicy)]
    [InlineData(SignaturePolicySource.CloudPolicy)]
    public void Policy_DisableSignatures_FromEitherPolicy_DisablesThem(SignaturePolicySource source)
    {
        SignaturePolicy policy = SignaturePolicy.Read(new FakeSignatureSettings().Policy(source, disable: true), "16.0");

        Assert.True(policy.SignaturesDisabled);
        Assert.Equal(source, policy.DisabledBy);
    }

    [Fact]
    public void Policy_DisableSignatures_InThePreferenceKey_IsNotAPolicy_AndZeroIsOff()
    {
        Assert.False(SignaturePolicy.Read(
            new FakeSignatureSettings().Policy(SignaturePolicySource.Preference, disable: true), "16.0").SignaturesDisabled);
        Assert.False(SignaturePolicy.Read(
            new FakeSignatureSettings().Policy(SignaturePolicySource.GroupPolicy, disable: false), "16.0").SignaturesDisabled);
    }

    [Fact]
    public void Policy_Precedence_CloudOverGroupOverPreference_PerValue()
    {
        FakeSignatureSettings settings = new FakeSignatureSettings()
            .Policy(SignaturePolicySource.Preference, newMessage: "Pref New", reply: "Pref Reply")
            .Policy(SignaturePolicySource.GroupPolicy, newMessage: "GPO New")
            .Policy(SignaturePolicySource.CloudPolicy, reply: "Cloud Reply");

        SignaturePolicy policy = SignaturePolicy.Read(settings, "16.0");

        Assert.Equal("GPO New", policy.NewMessage);
        Assert.Equal(SignaturePolicySource.GroupPolicy, policy.NewMessageSource);
        Assert.True(policy.NewMessageForced);
        Assert.Equal("Cloud Reply", policy.ReplyForward);
        Assert.Equal(SignaturePolicySource.CloudPolicy, policy.ReplyForwardSource);
        Assert.True(policy.ReplyForwardForced);
    }

    [Fact]
    public void Policy_APreferenceDefault_IsADefaultForEveryAccount_ButNotForced()
    {
        SignaturePolicy policy = SignaturePolicy.Read(
            new FakeSignatureSettings().Policy(SignaturePolicySource.Preference, newMessage: "Deployed"), "16.0");

        Assert.Equal("Deployed", policy.NewMessage);
        Assert.Equal(SignaturePolicySource.Preference, policy.NewMessageSource);
        Assert.False(policy.NewMessageForced);
        Assert.True(policy.AnyDefaultSet);
        Assert.Equal("preference", SignaturePolicy.Token(policy.NewMessageSource));
    }

    [Fact]
    public void Policy_KeyPaths_AreTheDocumentedOnes()
    {
        Assert.Equal(@"Software\Policies\Microsoft\Cloud\Office\16.0\Common\MailSettings",
            SignaturePolicy.KeyPath(SignaturePolicySource.CloudPolicy, "16.0"));
        Assert.Equal(@"Software\Policies\Microsoft\Office\16.0\Common\MailSettings",
            SignaturePolicy.KeyPath(SignaturePolicySource.GroupPolicy, "16.0"));
        Assert.Equal(@"Software\Microsoft\Office\16.0\Common\MailSettings",
            SignaturePolicy.KeyPath(SignaturePolicySource.Preference, "16.0"));
    }

    // ------------------------------------------------------------------ roaming (Q152, Q153)

    [Fact]
    public void Roaming_TheMaintainersLtsc2021_CannotRoam()
    {
        SignatureRoaming roaming = SignatureRoaming.Read(new FakeSignatureSettings(), "16.0");

        Assert.Equal(SignatureRoamingState.Off, roaming.State);
        Assert.False(roaming.MayRoam);
        Assert.Contains("predates Version 2303", roaming.Reason, StringComparison.Ordinal);
    }

    [Fact]
    public void Roaming_Microsoft365OnARoamingBuild_IsPossible()
    {
        SignatureRoaming roaming = SignatureRoaming.Read(new FakeSignatureSettings().Microsoft365(), "16.0");

        Assert.Equal(SignatureRoamingState.Possible, roaming.State);
        Assert.True(roaming.MayRoam);
    }

    [Fact]
    public void Roaming_Ltsc2024_IsOff_UnlessOutlooksOwnCopiesShowIt_ForThoseExchangeAccountsOnly()
    {
        // R1 on the Exchange test machine (2026-11-17): Office LTSC 2024 showed no sign of roaming. Not a
        // proof, so the verdict stays open to Outlook's own evidence.
        SignatureRoaming roaming = SignatureRoaming.Read(new FakeSignatureSettings().Ltsc2024(), "16.0");
        Assert.Equal(SignatureRoamingState.Off, roaming.State);
        Assert.True(roaming.EvidenceDecides);
        Assert.Contains("authoritative", roaming.Reason, StringComparison.Ordinal);

        SignatureRoaming evident = roaming.WithEvidence(
            new[] { "Work (jan@example.com)", "Home", "Old (pop@example.com)" }, new[] { "jan@example.com", "piet@example.com" });
        Assert.Equal(SignatureRoamingState.Evident, evident.State);
        Assert.Equal("evident", evident.Token);
        Assert.Equal(new[] { "jan@example.com" }, evident.EvidentAccounts);
        Assert.True(evident.MayRoamFor("JAN@example.com"));
        Assert.False(evident.MayRoamFor("piet@example.com"));
        Assert.False(evident.MayRoamFor("pop@example.com"));

        // A copy named after a POP3 or IMAP account is just a name: only Exchange accounts roam.
        Assert.Equal(SignatureRoamingState.Off, roaming.WithEvidence(new[] { "Old (pop@example.com)" }, Array.Empty<string>()).State);
    }

    [Fact]
    public void Roaming_AFinalVerdict_IsNotChangedByFiles()
    {
        string[] copies = { "Work (jan@example.com)" };
        string[] exchange = { "jan@example.com" };

        Assert.Equal(SignatureRoamingState.Off, SignatureRoaming.Read(
            new FakeSignatureSettings().Ltsc2024().RoamingSwitchedOff(), "16.0").WithEvidence(copies, exchange).State);
        Assert.Equal(SignatureRoamingState.Off, SignatureRoaming.Read(new FakeSignatureSettings(), "16.0").WithEvidence(copies, exchange).State);
        SignatureRoaming possible = SignatureRoaming.Read(new FakeSignatureSettings().Microsoft365(), "16.0").WithEvidence(copies, exchange);
        Assert.Equal(SignatureRoamingState.Possible, possible.State);
        Assert.True(possible.MayRoamFor("anyone@example.com"));
    }

    [Fact]
    public void Roaming_TheEnvironmentWeighsItsOwnFolder_AgainstItsProfilesExchangeAccounts()
    {
        string folder = Path.Combine(Path.GetTempPath(), "OutlookAI-RoamEvidence-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            File.WriteAllText(Path.Combine(folder, "Work (jan@example.com).htm"), "<p>x</p>");
            FakeProfileAccountRegistry accounts = new FakeProfileAccountRegistry()
                .Add("Outlook", 1, FakeProfileAccountRegistry.Exchange("jan@example.com"));

            SignatureEnvironment environment = SignatureEnvironment.Read(new FakeSignatureSettings().Ltsc2024(), accounts, "16.0", AppData, folder);

            Assert.Equal(SignatureRoamingState.Evident, environment.Roaming.State);
            Assert.True(environment.IsManagedInMailbox(Assert.Single(environment.Signatures)));
        }
        finally
        {
            Directory.Delete(folder, recursive: true);
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Roaming_SwitchedOff_ByEitherKeyAndEitherName_IsOff_EvenOnMicrosoft365(bool policy, bool temporaryToggle)
    {
        SignatureRoaming roaming = SignatureRoaming.Read(
            new FakeSignatureSettings().Microsoft365().RoamingSwitchedOff(policy, temporaryToggle), "16.0");

        Assert.Equal(SignatureRoamingState.Off, roaming.State);
        Assert.NotNull(roaming.SwitchedOffAt);
        Assert.Contains(temporaryToggle ? "TemporaryToggle" : "DisableRoamingSignatures", roaming.SwitchedOffAt, StringComparison.Ordinal);
    }

    [Fact]
    public void Roaming_Outlook2013_AndAnUnreadableBuild()
    {
        Assert.Equal(SignatureRoamingState.Off, SignatureRoaming.Decide("15.0", OfficeBuildInfo.Unknown, null).State);
        Assert.Equal(SignatureRoamingState.Off, SignatureRoaming.Decide("16.0", OfficeBuildInfo.Unknown, null).State);
        Assert.True(SignatureRoaming.Decide("16.0", OfficeBuildInfo.Unknown, null).EvidenceDecides);
        Assert.Equal(SignatureRoamingState.Off, SignatureRoaming.Decide("16.0", new OfficeBuildInfo("16.0.16226.20000", "O365ProPlusRetail"), null).State);
        Assert.Equal(SignatureRoamingState.Possible, SignatureRoaming.Decide("16.0", new OfficeBuildInfo("16.0.16227.20000", "O365ProPlusRetail"), null).State);
    }

    [Theory]
    [InlineData("Work (jan@example.com)", "jan@example.com")]
    [InlineData("Werk kort (j.modaal@example.com)", "j.modaal@example.com")]
    [InlineData("Work", null)]
    [InlineData("Work (short)", null)]
    [InlineData("Work (jan@example.com", null)]
    [InlineData("(jan@example.com)", null)]
    [InlineData("Work (two words@example.com)", null)]
    [InlineData("Work (a@b@c)", null)]
    [InlineData(null, null)]
    public void Roaming_ALocalCopyIsNamedAfterItsAccount(string? name, string? account)
    {
        Assert.Equal(account, SignatureRoaming.AccountSuffixOf(name));
    }

    // ------------------------------------------------------------------ profiles and the whole model

    [Fact]
    public void Environment_ListsEveryProfile_EachEntryWithItsOwnProfile_NeverMerged()
    {
        FakeProfileAccountRegistry accounts = new FakeProfileAccountRegistry()
            .Add("Outlook", 1, FakeProfileAccountRegistry.Pop3("jan@example.com", newSignature: "Home"))
            .Add("Work", 1, FakeProfileAccountRegistry.Pop3("jan@example.com", newSignature: "Office"))
            .Add("Work", 2, FakeProfileAccountRegistry.DataFile("jan@example.com"));

        SignatureEnvironment environment = SignatureEnvironment.Read(new FakeSignatureSettings(), accounts, "16.0", AppData, NoFolder);

        Assert.Equal(new[] { "Outlook", "Work" }, environment.ProfileNames);
        Assert.Null(environment.Profiles.DefaultProfile);
        Assert.Equal("Home", Assert.Single(environment.MailAccountsOf("Outlook")).NewSignature);
        Assert.Equal("Office", Assert.Single(environment.MailAccountsOf("WORK")).NewSignature);
        Assert.Equal("Work", environment.FindProfile("work"));
        Assert.Null(environment.FindProfile("Nope"));
        Assert.Empty(environment.Problems);
    }

    [Fact]
    public void Environment_APartThatCannotBeRead_IsNamed_AndTheRestStillAnswers()
    {
        SignatureEnvironment environment = SignatureEnvironment.Read(
            new FakeSignatureSettings(), new ThrowingAccounts(), "16.0", AppData, NoFolder);

        Assert.Contains(environment.Problems, p => p.StartsWith("the Outlook profile list", StringComparison.Ordinal));
        Assert.Contains(environment.Problems, p => p.StartsWith("the profiles' account lists", StringComparison.Ordinal));
        Assert.Empty(environment.Entries);
        Assert.Equal(SignatureRoamingState.Off, environment.Roaming.State);
    }

    [Fact]
    public void ProfileNameOf_ReadsTheLiveKeyShape()
    {
        Assert.Equal("Outlook", ProfileAccountEntries.ProfileNameOf(FakeProfileAccountRegistry.KeyFor("Outlook", 3)));
        Assert.Equal("OutlookAI-Tier", ProfileAccountEntries.ProfileNameOf(
            @"Software\Microsoft\Office\16.0\Outlook\Profiles\OutlookAI-Tier\9375CFF0413111d3B88A00104B2A6676\00000004"));
        Assert.Null(ProfileAccountEntries.ProfileNameOf("key-hub"));
        Assert.Null(ProfileAccountEntries.ProfileNameOf(null));
    }

    [Fact]
    public void Classify_TellsTheAccountKinds()
    {
        Assert.Equal(MailAccountKind.Pop3, ProfileAccountEntries.Classify("k", FakeProfileAccountRegistry.Pop3("a@b.example")).AccountKind);
        Assert.Equal(MailAccountKind.Imap, ProfileAccountEntries.Classify("k", FakeProfileAccountRegistry.Imap("a@b.example")).AccountKind);
        Assert.Equal(MailAccountKind.Exchange, ProfileAccountEntries.Classify("k", FakeProfileAccountRegistry.Exchange("a@b.example")).AccountKind);
        Assert.Null(ProfileAccountEntries.Classify("k", FakeProfileAccountRegistry.DataFile("a@b.example")).AccountKind);
    }

    private sealed class ThrowingAccounts : IProfileAccountRegistry
    {
        public IReadOnlyList<ProfileRegistryEntry> ReadEntries() => throw new UnauthorizedAccessException("denied");

        public ProfileList ReadProfiles() => throw new UnauthorizedAccessException("denied");

        public void SetString(string key, string valueName, string value) => throw new NotSupportedException();

        public void DeleteValue(string key, string valueName) => throw new NotSupportedException();
    }
}
