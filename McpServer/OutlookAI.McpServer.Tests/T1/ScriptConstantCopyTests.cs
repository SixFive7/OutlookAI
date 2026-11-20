using System.Text.RegularExpressions;

using OutlookAI.Core.AddIn;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The guest scripts' own copies of values the C# holds, held equal to it - Q171's checked copies, (a), the way
/// these were settled on 2026-11-20:
/// the live tier's subject tag, the test signatures' name prefix, and the account classification's class
/// identifiers and service names. The scripts run on the test guests with no assembly of this repository
/// loaded, so each spells its value out; one that drifted would go on working, against the wrong value -
/// a mail-sink probe the zero-artifact sweep no longer matches, a measurement that no longer tells the test
/// signatures from the user's, a profile entry counted as the wrong kind. Each test reads one declaration
/// out of a script's text and compares it with the constant; no script is changed or run.
///
/// <para>
/// Not here: <c>Connect-ExchangeAccount.ps1</c>'s <c>AccountType -eq 0</c>, Outlook's value for an
/// Exchange account. No C# of this repository reads an account's type, so there is nothing to hold it
/// equal to; what the C# classifies is a profile's account entries, by class and service name, which
/// <c>ScriptHelpers.ps1</c>'s <c>Test-IsMailAccountEntry</c> and <c>Add-IdentityAccount.ps1</c> copy.
/// </para>
/// </summary>
public sealed class ScriptConstantCopyTests
{
    [Fact]
    public void TheMailSinksProbeTag_IsTheLiveTiersSubjectTag()
    {
        Assert.Equal(LiveOutlookTestMailer.SubjectTag, Declared("Install-MailSink.ps1", "ProbeSubjectTag"));
    }

    [Theory]
    [InlineData("Measure-SignatureDialog.ps1", "TestPrefix")]
    [InlineData("Measure-SignatureRoaming.ps1", "SignaturePrefix")]
    public void TheSignatureMeasurementsPrefix_IsTheTestSignaturesPrefix(string script, string variable)
    {
        Assert.Equal(SignatureCatalog.TestSignaturePrefix, Declared(script, variable));
    }

    [Fact]
    public void TheIdentityAccountsPop3Class_IsTheClassificationsPop3Class()
    {
        Assert.Equal(ProfileAccountEntries.Pop3AccountClsid, Guid.Parse(Declared("Add-IdentityAccount.ps1", "Pop3Clsid")));
    }

    /// <summary>
    /// <c>Test-IsMailAccountEntry</c> calls an entry no mail account when it wraps a MAPI service that is a
    /// data file or an address book: the class and the five service names have to be the ones the C#
    /// classifies that way, no more and no fewer. A name the C# learns and the script does not would make
    /// the script call a data file a mail account, and refuse a profile that has none.
    /// </summary>
    [Fact]
    public void TheScriptsMailAccountTest_LeavesOutExactlyTheDataFilesAndAddressBooksTheClassificationKnows()
    {
        string helpers = File.ReadAllText(GuestScript("ScriptHelpers.ps1"));
        Match function = Regex.Match(helpers, @"(?ms)^function Test-IsMailAccountEntry \{.*?^\}");
        Assert.True(function.Success, "ScriptHelpers.ps1 declares no function Test-IsMailAccountEntry at the start of a line.");
        MatchCollection rules = Regex.Matches(
            function.Value, @"\$Clsid -eq '(?<clsid>\{[0-9A-Fa-f-]+\})' -and @\((?<names>[^)]*)\) -contains \$ServiceName");
        Assert.True(rules.Count == 1, "Test-IsMailAccountEntry holds " + rules.Count + " class-and-service tests, not the one this reads.");

        Assert.Equal(ProfileAccountEntries.MapiServiceClsid, Guid.Parse(rules[0].Groups["clsid"].Value));
        string[] scriptNames = Regex.Matches(rules[0].Groups["names"].Value, "'([^']*)'").Select(m => m.Groups[1].Value).ToArray();
        string[] csharpNames = ProfileAccountEntries.DataFileServiceNames.Concat(ProfileAccountEntries.AddressBookServiceNames).ToArray();
        Assert.Equal(csharpNames.OrderBy(n => n, StringComparer.Ordinal), scriptNames.OrderBy(n => n, StringComparer.Ordinal));
        foreach (string name in scriptNames)
        {
            ProfileEntryKind kind = ProfileAccountEntries.ClassifyKind(ProfileAccountEntries.MapiServiceClsid.ToString("B"), name);
            Assert.True(kind == ProfileEntryKind.DataFile || kind == ProfileEntryKind.AddressBook, name + " is a " + kind + " to the C#.");
        }
    }

    private static string GuestScript(string name) => Path.Combine(RepoPaths.RepoRoot(), "Testbed", "guest", name);

    /// <summary>The one single-quoted value <paramref name="script"/> assigns to <c>$variable</c> at the start of a line.</summary>
    private static string Declared(string script, string variable)
    {
        MatchCollection declarations = Regex.Matches(
            File.ReadAllText(GuestScript(script)), @"(?m)^\$" + Regex.Escape(variable) + @"\s*=\s*'(?<value>[^']*)'");
        Assert.True(declarations.Count == 1, script + " assigns $" + variable + " a quoted value " + declarations.Count + " time(s) at the start of a line, not once.");
        return declarations[0].Groups["value"].Value;
    }
}
