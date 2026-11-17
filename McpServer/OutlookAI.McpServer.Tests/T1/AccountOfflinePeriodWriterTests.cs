using System.Text;

using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// <see cref="AccountOfflinePeriodWriter"/>: an existing Exchange account's own offline period,
/// written into its section of the Outlook profile - against a fake profile holding the shapes the
/// Exchange VM measured on 2026-11-17 (Office LTSC 2024): 0x6649 months and 0x665A days as four
/// little-endian bytes, the account's address as a NUL-terminated UTF-16 value in the same section.
/// </summary>
public sealed class AccountOfflinePeriodWriterTests
{
    private const string Hub = "hub-test@example.invalid";
    private const string Other = "other@example.invalid";
    private const string ProfileRoot = @"Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\";

    private sealed class FakeProfile : IProfileSectionRegistry
    {
        private readonly Dictionary<string, Dictionary<string, object?>> _sections = new(StringComparer.OrdinalIgnoreCase);

        public List<string> Writes { get; } = new();

        /// <summary>Value names a write of throws, as Windows would refuse it.</summary>
        public HashSet<string> Refused { get; } = new(StringComparer.OrdinalIgnoreCase);

        /// <summary>When set, writes are accepted and silently dropped.</summary>
        public bool DropsWrites { get; set; }

        public string Add(string section, params (string Name, object? Value)[] values)
        {
            string key = ProfileRoot + section;
            var map = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach ((string name, object? value) in values)
                map[name] = value;
            _sections[key] = map;
            return key;
        }

        public byte[]? GetBytes(string key, string name) => _sections[key].TryGetValue(name, out object? v) ? v as byte[] : null;

        public IReadOnlyList<ProfileSection> ReadSections() =>
            _sections.Select(s => new ProfileSection(s.Key, new Dictionary<string, object?>(s.Value, StringComparer.OrdinalIgnoreCase))).ToList();

        public void SetBinary(string key, string valueName, byte[] value)
        {
            Writes.Add(key + "\\" + valueName + "=" + Convert.ToHexString(value));
            if (Refused.Contains(valueName))
                throw new UnauthorizedAccessException("Access to the registry key is denied.");
            if (!_sections.ContainsKey(key))
                throw new InvalidOperationException("The profile section no longer exists: " + key);
            if (!DropsWrites)
                _sections[key][valueName] = value;
        }
    }

    private static byte[] Long(int value) => BitConverter.GetBytes(value);

    private static byte[] Unicode(string text) => Encoding.Unicode.GetBytes(text + "\0");

    /// <summary>A profile like the Exchange VM's: the hub's section, another account's, and sections without the period.</summary>
    private static (FakeProfile Profile, string HubSection, string OtherSection) Profile(int months = 12, int days = 0)
    {
        var profile = new FakeProfile();
        string hub = profile.Add("3ba28d588c5522468fd9b33564c7122b",
            ("00036649", Long(months)), ("0003665a", Long(days)), ("00036601", Long(0x984)), ("001f3001", Unicode(Hub)));
        string other = profile.Add("ee89c689f03af149a5e236779aece5f8",
            ("00036649", Long(0)), ("0003665a", Long(0)), ("001f3001", Unicode(Other)));
        profile.Add("9375CFF0413111d3B88A00104B2A6676");
        profile.Add("0a0d020000000000c000000000000046", ("001f3001", Unicode(Hub)));
        return (profile, hub, other);
    }

    [Theory]
    [InlineData(0, 3, "3 days")]
    [InlineData(0, 7, "1 week")]
    [InlineData(0, 14, "2 weeks")]
    [InlineData(1, 0, "1 month")]
    [InlineData(3, 0, "3 months")]
    [InlineData(6, 0, "6 months")]
    [InlineData(12, 0, "1 year")]
    [InlineData(24, 0, "2 years")]
    [InlineData(36, 0, "3 years")]
    [InlineData(60, 0, "5 years")]
    [InlineData(0, 0, "All")]
    public void EachOfTheElevenIsWrittenAsItsPair_IntoTheAccountsSectionOnly(int months, int days, string label)
    {
        (FakeProfile profile, string hub, string other) = Profile(months: months == 1 ? 12 : 1, days: 0);

        AccountOfflinePeriodResult result = AccountOfflinePeriodWriter.Write(profile, Hub, months, days);

        Assert.Equal(AccountOfflinePeriodOutcome.Written, result.Outcome);
        Assert.Equal(label, AccountOfflinePeriodWriter.LabelOf(months, days));
        Assert.Equal(Long(months), profile.GetBytes(hub, "00036649"));
        Assert.Equal(Long(days), profile.GetBytes(hub, "0003665a"));
        Assert.Equal(new[] { hub + @"\00036649=" + Convert.ToHexString(Long(months)), hub + @"\0003665a=" + Convert.ToHexString(Long(days)) }, profile.Writes);
        Assert.Equal(Long(0), profile.GetBytes(other, "00036649"));
        Assert.True(result.RestartNeeded);
        Assert.Contains("Restart Outlook", result.Message, StringComparison.Ordinal);
        Assert.Equal(hub, result.SectionKey);
    }

    [Theory]
    [InlineData(12, 7)]
    [InlineData(18, 0)]
    [InlineData(0, 1)]
    [InlineData(120, 0)]
    [InlineData(-1, 0)]
    public void AnyOtherPair_IsRefused_BeforeTheProfileIsEvenRead(int months, int days)
    {
        (FakeProfile profile, string hub, _) = Profile();

        AccountOfflinePeriodResult result = AccountOfflinePeriodWriter.Write(profile, Hub, months, days);

        Assert.Equal(AccountOfflinePeriodOutcome.NotOffered, result.Outcome);
        Assert.Empty(profile.Writes);
        Assert.Equal(Long(12), profile.GetBytes(hub, "00036649"));
        Assert.False(AccountOfflinePeriodWriter.IsOffered(months, days));
        Assert.Null(AccountOfflinePeriodWriter.LabelOf(months, days));
    }

    [Fact]
    public void TheAddressMustBeAWholeValue_CaseAndSpacesAside()
    {
        (FakeProfile profile, string hub, _) = Profile();

        Assert.Equal(AccountOfflinePeriodOutcome.Written, AccountOfflinePeriodWriter.Write(profile, "  HUB-test@Example.INVALID ", 0, 0).Outcome);
        Assert.Equal(Long(0), profile.GetBytes(hub, "00036649"));

        var substring = new FakeProfile();
        substring.Add("s1", ("00036649", Long(12)), ("0003665a", Long(0)), ("001f3001", Unicode("x" + Hub)));
        Assert.Equal(AccountOfflinePeriodOutcome.NoSuchAccount, AccountOfflinePeriodWriter.Write(substring, Hub, 0, 0).Outcome);
        Assert.Empty(substring.Writes);
    }

    [Fact]
    public void ASectionWithTheAddressButNoPeriod_OrThePeriodButAnotherAddress_IsNeverTouched()
    {
        var profile = new FakeProfile();
        profile.Add("addressOnly", ("001f3001", Unicode(Hub)));
        profile.Add("periodOnly", ("00036649", Long(12)), ("0003665a", Long(0)), ("001f3001", Unicode(Other)));
        profile.Add("notUnicode", ("00036649", Long(12)), ("001e3001", Encoding.ASCII.GetBytes(Hub + "\0")));

        AccountOfflinePeriodResult result = AccountOfflinePeriodWriter.Write(profile, Hub, 0, 0);

        Assert.Equal(AccountOfflinePeriodOutcome.NoSuchAccount, result.Outcome);
        Assert.Contains("not an Exchange account of this profile", result.Message, StringComparison.Ordinal);
        Assert.Empty(profile.Writes);
    }

    [Fact]
    public void TwoSectionsHoldingBoth_AreARefusal_NeverAGuess()
    {
        (FakeProfile profile, _, _) = Profile();
        profile.Add("second", ("00036649", Long(1)), ("0003665a", Long(0)), ("001f6641", Unicode(Hub)));

        AccountOfflinePeriodResult result = AccountOfflinePeriodWriter.Write(profile, Hub, 0, 0);

        Assert.Equal(AccountOfflinePeriodOutcome.SeveralSections, result.Outcome);
        Assert.Empty(profile.Writes);
    }

    [Fact]
    public void ThePeriodItAlreadyHas_IsNotWrittenAgain()
    {
        (FakeProfile profile, _, _) = Profile(months: 0, days: 0);

        AccountOfflinePeriodResult result = AccountOfflinePeriodWriter.Write(profile, Hub, 0, 0);

        Assert.Equal(AccountOfflinePeriodOutcome.AlreadySet, result.Outcome);
        Assert.True(result.Succeeded);
        Assert.False(result.RestartNeeded);
        Assert.Empty(profile.Writes);
    }

    [Fact]
    public void AMissingDaysValue_IsWrittenBesideTheMonths_AsOutlookWritesThePair()
    {
        var profile = new FakeProfile();
        string section = profile.Add("s", ("00036649", Long(12)), ("001f3001", Unicode(Hub)));

        AccountOfflinePeriodResult result = AccountOfflinePeriodWriter.Write(profile, Hub, 0, 14);

        Assert.Equal(AccountOfflinePeriodOutcome.Written, result.Outcome);
        Assert.Null(result.DaysBefore);
        Assert.Equal(Long(14), profile.GetBytes(section, "0003665a"));
    }

    [Fact]
    public void AMonthsWriteWindowsRefuses_ChangesNothing()
    {
        (FakeProfile profile, string hub, _) = Profile();
        profile.Refused.Add("00036649");

        AccountOfflinePeriodResult result = AccountOfflinePeriodWriter.Write(profile, Hub, 0, 0);

        Assert.Equal(AccountOfflinePeriodOutcome.NotWritten, result.Outcome);
        Assert.Contains("Nothing was changed", result.Message, StringComparison.Ordinal);
        Assert.Equal(Long(12), profile.GetBytes(hub, "00036649"));
        Assert.Equal(Long(0), profile.GetBytes(hub, "0003665a"));
    }

    [Fact]
    public void ADaysWriteThatFails_PutsTheMonthsBack_SoNoHalfPairIsLeft()
    {
        (FakeProfile profile, string hub, _) = Profile(months: 12, days: 0);
        profile.Refused.Add("0003665a");

        AccountOfflinePeriodResult result = AccountOfflinePeriodWriter.Write(profile, Hub, 0, 7);

        Assert.Equal(AccountOfflinePeriodOutcome.NotWritten, result.Outcome);
        Assert.Contains("put back", result.Message, StringComparison.Ordinal);
        Assert.Equal(Long(12), profile.GetBytes(hub, "00036649"));
        Assert.Equal(Long(0), profile.GetBytes(hub, "0003665a"));
    }

    [Fact]
    public void AWriteThatDoesNotReadBack_IsSaidSo()
    {
        (FakeProfile profile, _, _) = Profile();
        profile.DropsWrites = true;

        AccountOfflinePeriodResult result = AccountOfflinePeriodWriter.Write(profile, Hub, 0, 0);

        Assert.Equal(AccountOfflinePeriodOutcome.ReadBackMismatch, result.Outcome);
        Assert.False(result.Succeeded);
        Assert.Contains("reads back as (12, 0) instead of (0, 0)", result.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ABlankAddress_IsRefused()
    {
        (FakeProfile profile, _, _) = Profile();

        Assert.Equal(AccountOfflinePeriodOutcome.NoSuchAccount, AccountOfflinePeriodWriter.Write(profile, "  ", 0, 0).Outcome);
        Assert.Empty(profile.Writes);
    }

    [Fact]
    public void TheValueShapes_FourLittleEndianBytes_AndNulTerminatedUtf16()
    {
        Assert.Equal(new byte[] { 0x3C, 0x00, 0x00, 0x00 }, AccountOfflinePeriodWriter.EncodeLong(60));
        Assert.Equal(36, AccountOfflinePeriodWriter.DecodeLong(new byte[] { 0x24, 0, 0, 0 }));
        Assert.Equal(12, AccountOfflinePeriodWriter.DecodeLong(12));
        Assert.Null(AccountOfflinePeriodWriter.DecodeLong(new byte[] { 1, 2 }));
        Assert.Null(AccountOfflinePeriodWriter.DecodeLong("12"));
        Assert.Equal(Hub, AccountOfflinePeriodWriter.DecodeUnicode(Unicode(Hub)));
        Assert.Equal(Hub, AccountOfflinePeriodWriter.DecodeUnicode(Hub + "\0"));
        Assert.Null(AccountOfflinePeriodWriter.DecodeUnicode(new byte[] { 1, 2, 3 }));
    }

    [Fact]
    public void TheLiveReaderLoadsOnlyWhatTheRuleNeeds()
    {
        Assert.True(AccountOfflinePeriodWriter.IsValueRead("00036649"));
        Assert.True(AccountOfflinePeriodWriter.IsValueRead("0003665A"));
        Assert.True(AccountOfflinePeriodWriter.IsValueRead("001f3001"));
        Assert.False(AccountOfflinePeriodWriter.IsValueRead("00036601"));
        Assert.False(AccountOfflinePeriodWriter.IsValueRead("01023d15"));
    }
}
