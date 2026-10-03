using System;
using System.Collections.Generic;
using System.Linq;

using OutlookAI.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The elevated helper's command line - <c>Services\PolicyWriterRequest.cs</c> and
/// <c>Services\CachedModePolicy.cs</c>, LINKED into this assembly - and EVERYTHING IT MUST REFUSE.
///
/// <para>
/// WHY SO MANY REFUSALS. <c>OutlookAI.PolicyWriter.exe</c> runs with an administrator's token once a
/// user approves its UAC prompt, and whoever controls its command line controls what that token
/// writes (Q128, decided 2026-10-03). So it accepts exactly what the add-in sends - one user SID,
/// one supported Office major, and some of five value names with the values offered for each - and
/// nothing else, refused WHOLE before anything is written. Each refusal below is a way a request
/// could reach a key or a value it should not: another hive, another key, another value, a value
/// Outlook may read differently, or a spelling Windows itself never produces.
/// </para>
/// </summary>
public sealed class PolicyWriterRequestTests
{
    private const string Sid = "S-1-5-21-1004336348-1177238915-682003330-1001";
    private const string AzureSid = "S-1-12-1-3473471489-1287520453-1916219034-1384431621";

    private static string[] Request(params string[] values) =>
        new[] { "--sid", Sid, "--office", "16.0" }.Concat(values).ToArray();

    // ===== What it accepts =====

    [Fact]
    public void TheRequestTheAddInBuilds_ParsesBackToTheSameRequest()
    {
        var values = CachedModePolicy.ValueNames.Select(n => new KeyValuePair<string, int>(n, CachedModePolicy.ShippedDefault(n))).ToList();

        string line = PolicyWriterCommandLine.Build(Sid, "16.0", values);
        PolicyWriteParse parsed = PolicyWriterCommandLine.Parse(line.Split(' '));

        Assert.True(parsed.Ok, parsed.Refusal);
        Assert.Equal(Sid, parsed.Request.Sid);
        Assert.Equal("16.0", parsed.Request.OfficeVersion);
        Assert.Equal(values, parsed.Request.Values);
        Assert.Equal(@"Software\Policies\Microsoft\Office\16.0\Outlook\Cached Mode", parsed.Request.KeyPath);
        Assert.Equal(
            "--sid " + Sid + " --office 16.0 SyncWindowSetting=0 SyncWindowSettingDays=0 DownloadSharedFolders=1 CacheOthersMail=1 DisableSyncSliderForSharedMailbox=1",
            line);
    }

    [Fact]
    public void ASubsetOfValues_AndTheSwitchesInAnyOrder_AreAccepted()
    {
        PolicyWriteParse parsed = PolicyWriterCommandLine.Parse(new[] { "CacheOthersMail=0", "--office", "15.0", "SyncWindowSetting=12", "--sid", AzureSid });

        Assert.True(parsed.Ok, parsed.Refusal);
        Assert.Equal(AzureSid, parsed.Request.Sid);
        Assert.Equal(@"Software\Policies\Microsoft\Office\15.0\Outlook\Cached Mode", parsed.Request.KeyPath);
        Assert.Equal(new[] { new KeyValuePair<string, int>("CacheOthersMail", 0), new KeyValuePair<string, int>("SyncWindowSetting", 12) }, parsed.Request.Values);
    }

    [Theory]
    [InlineData("SyncWindowSetting", new[] { 0, 1, 3, 6, 12, 24 })]
    [InlineData("SyncWindowSettingDays", new[] { 0, 3, 7, 14 })]
    [InlineData("DownloadSharedFolders", new[] { 0, 1 })]
    [InlineData("CacheOthersMail", new[] { 0, 1 })]
    [InlineData("DisableSyncSliderForSharedMailbox", new[] { 0, 1 })]
    public void EveryOfferedValue_IsAccepted_AndNothingBetweenThem(string name, int[] offered)
    {
        for (int value = -2; value <= 70; value++)
        {
            PolicyWriteParse parsed = PolicyWriterCommandLine.Parse(Request(name + "=" + value));
            Assert.True(parsed.Ok == offered.Contains(value), name + "=" + value + ": " + parsed.Refusal);
        }
    }

    // ===== What it refuses, whole =====

    public static TheoryData<string, string[]> Refused => new()
    {
        // No request at all.
        { "no arguments", Array.Empty<string>() },
        { "only the switches", new[] { "--sid", Sid, "--office", "16.0" } },
        { "no SID", new[] { "--office", "16.0", "SyncWindowSetting=0" } },
        { "no Office", new[] { "--sid", Sid, "SyncWindowSetting=0" } },
        { "a switch with nothing after it", new[] { "SyncWindowSetting=0", "--office", "16.0", "--sid" } },
        { "the SID twice", new[] { "--sid", Sid, "--sid", Sid, "--office", "16.0", "SyncWindowSetting=0" } },
        { "the Office twice", new[] { "--sid", Sid, "--office", "16.0", "--office", "16.0", "SyncWindowSetting=0" } },
        { "a value twice", Request("SyncWindowSetting=0", "SyncWindowSetting=0") },
        { "a value twice, two numbers", Request("CacheOthersMail=1", "CacheOthersMail=0") },

        // Not a user's own hive - or not spelled the way Windows spells one.
        { "SYSTEM", new[] { "--sid", "S-1-5-18", "--office", "16.0", "SyncWindowSetting=0" } },
        { "LOCAL SERVICE", new[] { "--sid", "S-1-5-19", "--office", "16.0", "SyncWindowSetting=0" } },
        { "the Administrators group", new[] { "--sid", "S-1-5-32-544", "--office", "16.0", "SyncWindowSetting=0" } },
        { "a service SID", new[] { "--sid", "S-1-5-80-1-2-3-4-5", "--office", "16.0", "SyncWindowSetting=0" } },
        { "the .DEFAULT hive", new[] { "--sid", ".DEFAULT", "--office", "16.0", "SyncWindowSetting=0" } },
        { "the _Classes hive", new[] { "--sid", Sid + "_Classes", "--office", "16.0", "SyncWindowSetting=0" } },
        { "a path instead of a SID", new[] { "--sid", Sid + @"\Software", "--office", "16.0", "SyncWindowSetting=0" } },
        { "HKCU", new[] { "--sid", "HKCU", "--office", "16.0", "SyncWindowSetting=0" } },
        { "three sub-authorities", new[] { "--sid", "S-1-5-21-1-2-3", "--office", "16.0", "SyncWindowSetting=0" } },
        { "five sub-authorities", new[] { "--sid", "S-1-5-21-1-2-3-4-5", "--office", "16.0", "SyncWindowSetting=0" } },
        { "a lowercase s", new[] { "--sid", "s-1-5-21-1-2-3-1001", "--office", "16.0", "SyncWindowSetting=0" } },
        { "a leading zero", new[] { "--sid", "S-1-5-21-01-2-3-1001", "--office", "16.0", "SyncWindowSetting=0" } },
        { "a plus sign", new[] { "--sid", "S-1-5-21-+1-2-3-1001", "--office", "16.0", "SyncWindowSetting=0" } },
        { "a sub-authority past 32 bits", new[] { "--sid", "S-1-5-21-1-2-3-4294967296", "--office", "16.0", "SyncWindowSetting=0" } },
        { "a leading space", new[] { "--sid", " " + Sid, "--office", "16.0", "SyncWindowSetting=0" } },
        { "an Arabic-Indic digit", new[] { "--sid", "S-1-5-21-1-2-3-٤", "--office", "16.0", "SyncWindowSetting=0" } },
        { "an empty SID", new[] { "--sid", "", "--office", "16.0", "SyncWindowSetting=0" } },

        // Not an Office this product knows.
        { "Office without its .0", new[] { "--sid", Sid, "--office", "16", "SyncWindowSetting=0" } },
        { "an Office that does not exist", new[] { "--sid", Sid, "--office", "18.0", "SyncWindowSetting=0" } },
        { "a path as the Office", new[] { "--sid", Sid, "--office", @"..\16.0", "SyncWindowSetting=0" } },
        { "an Office with a space", new[] { "--sid", Sid, "--office", "16.0 ", "SyncWindowSetting=0" } },

        // Not one of the five.
        { "a value name in other case", Request("syncwindowsetting=0") },
        { "an unknown value name", Request("NoOST=1") },
        { "a Run key", Request(@"..\..\..\..\Windows\CurrentVersion\Run\x=1") },
        { "a path as the name", Request(@"Software\Policies\Microsoft\Office\16.0\Outlook\Cached Mode\SyncWindowSetting=0") },
        { "a name with a space", Request("SyncWindowSetting =0") },
        { "no name", Request("=1") },
        { "no value", Request("SyncWindowSetting=") },
        { "no equals sign", Request("SyncWindowSetting") },
        { "two equals signs", Request("SyncWindowSetting==0") },
        { "two values in one", Request("SyncWindowSetting=0=1") },

        // Not a value OutlookAI offers - or not a plain number.
        { "a month count Outlook does not offer", Request("SyncWindowSetting=2") },
        { "five years", Request("SyncWindowSetting=60") },
        { "one day", Request("SyncWindowSettingDays=1") },
        { "a switch at 2", Request("DownloadSharedFolders=2") },
        { "a negative number", Request("SyncWindowSetting=-1") },
        { "a plus sign on a value", Request("SyncWindowSetting=+1") },
        { "a leading zero on a value", Request("SyncWindowSetting=01") },
        { "a space in a value", Request("SyncWindowSetting= 1") },
        { "hexadecimal", Request("SyncWindowSetting=0x1") },
        { "a decimal point", Request("CacheOthersMail=1.0") },
        { "past 32 bits", Request("SyncWindowSetting=4294967296") },
        { "a full-width digit", Request("SyncWindowSetting=１") },

        // Not a switch this helper has.
        { "a hive switch", new[] { "--hive", "HKLM", "--sid", Sid, "--office", "16.0", "SyncWindowSetting=0" } },
        { "a switch in other case", new[] { "--SID", Sid, "--office", "16.0", "SyncWindowSetting=0" } },
        { "a switch with an equals sign", new[] { "--sid=" + Sid, "--office", "16.0", "SyncWindowSetting=0" } },
        { "a slash switch", Request("/quiet") },
        { "a stray word", Request("SyncWindowSetting=0", "please") },
        { "an empty token", Request("SyncWindowSetting=0", "") },
    };

    [Theory]
    [MemberData(nameof(Refused))]
    public void Refuses(string why, string[] args)
    {
        PolicyWriteParse parsed = PolicyWriterCommandLine.Parse(args);

        Assert.False(parsed.Ok, why + " was accepted");
        Assert.False(string.IsNullOrWhiteSpace(parsed.Refusal), why + ": a refusal says why");
        // A refused parse carries no request that could be written by mistake.
        Assert.Empty(parsed.Request.Values);
        Assert.Equal(string.Empty, parsed.Request.Sid);
    }

    [Fact]
    public void ANullArgumentArray_IsRefusedNotThrown()
    {
        Assert.False(PolicyWriterCommandLine.Parse(null!).Ok);
    }

    [Fact]
    public void ARefusalNeverEchoesControlCharactersOrUnboundedInput()
    {
        PolicyWriteParse parsed = PolicyWriterCommandLine.Parse(Request("Bad\u0000Name\r\n" + new string('x', 5000) + "=1"));

        Assert.False(parsed.Ok);
        Assert.True(parsed.Refusal.Length < 400, "refusal is " + parsed.Refusal.Length + " characters");
        Assert.DoesNotContain('\0', parsed.Refusal);
        Assert.DoesNotContain('\n', parsed.Refusal);
    }

    // ===== The add-in side: it never builds what the helper would refuse =====

    [Fact]
    public void Build_RefusesWhatTheHelperWouldRefuse_BeforeAnyPromptIsShown()
    {
        var ok = new List<KeyValuePair<string, int>> { new("SyncWindowSetting", 0) };

        Assert.Throws<ArgumentException>(() => PolicyWriterCommandLine.Build("S-1-5-18", "16.0", ok));
        Assert.Throws<ArgumentException>(() => PolicyWriterCommandLine.Build(Sid, "18.0", ok));
        Assert.Throws<ArgumentException>(() => PolicyWriterCommandLine.Build(Sid, "16.0", new List<KeyValuePair<string, int>>()));
        Assert.Throws<ArgumentException>(() => PolicyWriterCommandLine.Build(Sid, "16.0", new List<KeyValuePair<string, int>> { new("SyncWindowSetting", 2) }));
        Assert.Throws<ArgumentException>(() => PolicyWriterCommandLine.Build(Sid, "16.0", new List<KeyValuePair<string, int>> { new("Run", 1) }));
        Assert.Throws<ArgumentException>(() => PolicyWriterCommandLine.Build(Sid, "16.0",
            new List<KeyValuePair<string, int>> { new("SyncWindowSetting", 0), new("SyncWindowSetting", 1) }));
    }

    // ===== The pieces =====

    [Theory]
    [InlineData(Sid, true)]
    [InlineData(AzureSid, true)]
    [InlineData("S-1-5-21-0-0-0-0", true)]
    [InlineData("S-1-5-21-4294967295-4294967295-4294967295-4294967295", true)]
    [InlineData("S-1-5-18", false)]
    [InlineData("S-1-5-21-1-2-3", false)]
    [InlineData("S-1-12-2-1-2-3-4", false)]
    [InlineData("S-1-5-22-1-2-3-4", false)]
    [InlineData("S-2-5-21-1-2-3-4", false)]
    [InlineData("S-1-5-21-1-2-3-4-", false)]
    [InlineData("-S-1-5-21-1-2-3-4", false)]
    [InlineData("S-1-5-21-1--3-4", false)]
    public void UserAccountSids(string text, bool expected)
    {
        Assert.Equal(expected, PolicyWriterCommandLine.IsUserAccountSid(text));
    }

    [Theory]
    [InlineData("0", true, 0)]
    [InlineData("14", true, 14)]
    [InlineData("2147483647", true, 2147483647)]
    [InlineData("2147483648", false, 0)]
    [InlineData("00", false, 0)]
    [InlineData("", false, 0)]
    [InlineData("1 ", false, 0)]
    [InlineData("١", false, 0)]
    public void PlainNumbers(string text, bool ok, int value)
    {
        Assert.Equal(ok, PolicyWriterCommandLine.TryParseCanonicalNumber(text, out int parsed));
        Assert.Equal(value, parsed);
    }

    [Fact]
    public void EveryExitCode_HasItsOwnSentence()
    {
        int[] codes =
        {
            PolicyWriterExit.Written, PolicyWriterExit.Failed, PolicyWriterExit.RefusedArguments, PolicyWriterExit.RefusedUser,
            PolicyWriterExit.RefusedHiveNotLoaded, PolicyWriterExit.AccessDenied, PolicyWriterExit.ReadBackMismatch,
        };

        Assert.Equal(codes.Length, codes.Distinct().Count());
        Assert.Equal(codes.Length, codes.Select(PolicyWriterExit.Describe).Distinct().Count());
        Assert.Contains("unknown", PolicyWriterExit.Describe(1223));
        Assert.Equal(0, PolicyWriterExit.Written);
    }
}
