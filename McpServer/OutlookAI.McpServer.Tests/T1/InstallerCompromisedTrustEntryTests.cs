using System.Security.Cryptography;
using System.Text.RegularExpressions;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The installer's removal of VSTO's trust entries for the COMPROMISED signing key - the key of the
/// certificate that signed v1.0.0.1 to v2.3.1.111, whose private half became public - read out of
/// <c>Installer.iss</c>, which is Inno Setup's Pascal and runs nowhere in this suite. The compiled
/// installer's own run against planted entries is <c>Testbed/guest/Measure-TrustRemoval.ps1 -Scenario
/// InstallerCleanup</c> on a test guest, recorded in Docs/live-tier-on-the-vm.md section 4.8.
///
/// <para>
/// Why it exists. Clicking Install on Office's "Publisher cannot be verified" prompt writes an entry
/// under HKCU\Software\Microsoft\VSTO\Security\Inclusion holding the .vsto's URL and the public key
/// that signed it, and from then on the VSTO runtime loads any add-in at that URL signed with that
/// key without asking: it compares the URL and the key's bytes and never reads the certificate
/// (its own IL, OfficeAddInDeploymentManager.VerifyAddInTrust and AddInSecurityEntry.Equals).
/// Users who installed v1 to v2.3.1 can still hold one for the leaked key. Decision 193 (b) removes
/// the leaked certificate from every trust store; these entries name the key, so they go by the key -
/// its public half only, which names nobody.
/// </para>
/// </summary>
public sealed class InstallerCompromisedTrustEntryTests
{
    /// <summary>
    /// The SHA-256 of the leaked key's modulus, read on 2026-11-17 from the public half of the
    /// certificate whose SHA-1 thumbprint is E205060633DD7062D4F90033130542948A69D068. That
    /// certificate is not in this repository and must not be (InstallerCompromisedCertificateTests);
    /// this ties the installer's define to it.
    /// </summary>
    private const string ModulusSha256 = "5C1D4C01B1E620EA2A41967C0272868B8EDA368F01E563778E8ABD42C9C724C3";

    /// <summary>The key the VSTO runtime's UserInclusionList reads and writes, under HKCU.</summary>
    private const string InclusionKey = @"Software\Microsoft\VSTO\Security\Inclusion";

    private static string Installer() => File.ReadAllText(Path.Combine(RepoPaths.RepoRoot(), "Installer.iss"));

    private static string Define(string iss, string name)
    {
        Match m = Regex.Match(iss, "^#define " + Regex.Escape(name) + " \"([^\"]*)\"\\r?$", RegexOptions.Multiline);
        Assert.True(m.Success, "Installer.iss defines no " + name);
        return m.Groups[1].Value;
    }

    /// <summary>A [Code] routine, from its heading to the "end;" that closes it at the start of a line.</summary>
    private static string Routine(string iss, string kind, string name)
    {
        Match m = Regex.Match(iss, "^" + kind + " " + Regex.Escape(name) + @"\b.*?^end;", RegexOptions.Multiline | RegexOptions.Singleline);
        Assert.True(m.Success, "Installer.iss has no " + kind + " " + name);
        return m.Value;
    }

    [Fact]
    public void TheKeyIsNamedByItsModulusAlone_WrittenOnceInTheInstaller()
    {
        string iss = Installer();
        string modulus = Define(iss, "CompromisedKeyModulus");
        byte[] bytes = Convert.FromBase64String(modulus);

        Assert.Equal(ModulusSha256, Convert.ToHexString(SHA256.HashData(bytes)));
        // 1024 bits with no leading zero byte, in canonical base64: the form .NET's ToXmlString writes,
        // which is how the VSTO runtime stores the key when somebody clicks Install.
        Assert.Equal(128, bytes.Length);
        Assert.True(bytes[0] >= 0x80, "the modulus starts with a zero byte, which the runtime never writes");
        Assert.Equal(modulus, Convert.ToBase64String(bytes));
        // Every use goes through the define, so the 172 characters are written exactly once.
        Assert.Single(Regex.Matches(iss, Regex.Escape(modulus)));
        Assert.Contains("Pos('{#CompromisedKeyModulus}', WithoutWhiteSpace(PublicKey)) > 0", Routine(iss, "procedure", "RemoveCompromisedTrustEntries"));
    }

    [Fact]
    public void EveryEntryHoldingTheKeyIsRemoved_MatchedOnThePublicKeyAlone()
    {
        string iss = Installer();
        string remove = Routine(iss, "procedure", "RemoveCompromisedTrustEntries");

        Assert.Matches(@"(?m)^\s*VstoInclusionKey = '" + Regex.Escape(InclusionKey) + "';\r?$", iss);
        Assert.Contains("RegGetSubkeyNames(HKCU, VstoInclusionKey, Names)", remove);
        Assert.Contains(@"RegQueryStringValue(HKCU, VstoInclusionKey + '\' + Names[I], 'PublicKey', PublicKey)", remove);
        Assert.Contains(@"RegDeleteKeyIncludingSubkeys(HKCU, VstoInclusionKey + '\' + Names[I])", remove);
        // Never the Url: it names a user's folders, and the key is all the runtime trusts.
        Assert.DoesNotContain("Url", remove, StringComparison.Ordinal);
        // One list read, one value read, one key deleted: nothing else in the registry is touched.
        Assert.Equal(new[] { "RegGetSubkeyNames", "RegQueryStringValue", "RegDeleteKeyIncludingSubkeys" },
                     Regex.Matches(remove, @"\b(Reg\w+)\(").Select(m => m.Groups[1].Value).ToArray());
    }

    [Fact]
    public void EveryInstallUpdateAndUninstallRunsIt_BeforeTheCertificateStores()
    {
        string iss = Installer();

        // In ssPostInstall after InstallCertificate, outside the interactive-only block: a silent
        // auto-update runs it too.
        Assert.Matches(@"InstallCertificate;\s+(?://[^\n]*\n\s*)*RemoveCompromisedTrustEntries;\s+RemoveCompromisedCertificate;", Routine(iss, "procedure", "CurStepChanged"));
        Assert.Matches(@"if CurUninstallStep = usUninstall then\s+begin\s+RemoveCompromisedTrustEntries;\s+RemoveCompromisedCertificate;\s+end;", Routine(iss, "procedure", "CurUninstallStepChanged"));
    }

    [Fact]
    public void WhiteSpaceIsWhatTheBase64DecoderSkips_AndNothingElse()
    {
        // .NET's base64 decoder, which FromXmlString uses, skips space, tab, carriage return and line
        // feed inside the text, so a value wrapped over lines is still the key to the runtime.
        Assert.Equal(new[] { '\t', '\n', '\r', ' ' }, DroppedCharacters(Installer()).OrderBy(c => c).ToArray());
    }

    /// <summary>
    /// The installer's rule applied to the value in each spelling the VSTO runtime reads as this key.
    /// Measured 2026-11-17 on .NET Framework 4.8, the runtime's CLR, with the runtime's own comparison
    /// - FromXmlString, then ExportCspBlob: all five equal to the canonical value. A modulus with a
    /// leading zero byte, or another exponent, was not equal to it, so the rule need not find those.
    /// </summary>
    [Theory]
    [InlineData("canonical")]
    [InlineData("wrapped")]
    [InlineData("namespaced")]
    [InlineData("prefixed")]
    [InlineData("swapped")]
    public void TheRuleFindsTheKeyInEverySpellingTheRuntimeReadsAsIt(string spelling)
    {
        string iss = Installer();

        Assert.True(InstallerRuleMatches(iss, Spell(Define(iss, "CompromisedKeyModulus"), spelling)));
    }

    [Fact]
    public void TheRuleLeavesEveryOtherKeysEntry()
    {
        string iss = Installer();
        using RSA other = RSA.Create(1024);
        string otherModulus = Convert.ToBase64String(other.ExportParameters(false).Modulus!);

        Assert.False(InstallerRuleMatches(iss, Spell(otherModulus, "canonical")));
        Assert.False(InstallerRuleMatches(iss, string.Empty));
    }

    /// <summary>The characters WithoutWhiteSpace drops, read out of its Pascal: each <c>S[I] &lt;&gt; ...</c>.</summary>
    private static HashSet<char> DroppedCharacters(string iss)
    {
        string function = Routine(iss, "function", "WithoutWhiteSpace");
        var dropped = new HashSet<char>();
        foreach (Match m in Regex.Matches(function, @"S\[I\] <> (?:'(.)'|#(\d+))"))
            dropped.Add(m.Groups[1].Success ? m.Groups[1].Value[0] : (char)int.Parse(m.Groups[2].Value, System.Globalization.CultureInfo.InvariantCulture));
        return dropped;
    }

    /// <summary>What RemoveCompromisedTrustEntries decides for one PublicKey value: white space out, then the modulus as a substring.</summary>
    private static bool InstallerRuleMatches(string iss, string publicKeyValue)
    {
        HashSet<char> dropped = DroppedCharacters(iss);
        string stripped = new(publicKeyValue.Where(c => !dropped.Contains(c)).ToArray());
        return stripped.Contains(Define(iss, "CompromisedKeyModulus"), StringComparison.Ordinal);
    }

    private static string Spell(string modulus, string spelling) => spelling switch
    {
        "canonical" => "<RSAKeyValue><Modulus>" + modulus + "</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>",
        "wrapped" => "<RSAKeyValue>\r\n  <Modulus>\r\n    " + string.Join("\r\n    ", modulus.Chunk(64).Select(c => new string(c))) + "\r\n  </Modulus>\r\n  <Exponent>AQAB</Exponent>\r\n</RSAKeyValue>",
        "namespaced" => "<RSAKeyValue xmlns=\"http://www.w3.org/2000/09/xmldsig#\"><Modulus>" + modulus + "</Modulus><Exponent>AQAB</Exponent></RSAKeyValue>",
        "prefixed" => "<ds:RSAKeyValue xmlns:ds=\"http://www.w3.org/2000/09/xmldsig#\"><ds:Modulus>" + modulus + "</ds:Modulus><ds:Exponent>AQAB</ds:Exponent></ds:RSAKeyValue>",
        "swapped" => "<RSAKeyValue><Exponent>AQAB</Exponent><Modulus>" + modulus + "</Modulus></RSAKeyValue>",
        _ => throw new ArgumentOutOfRangeException(nameof(spelling), spelling, "no such spelling"),
    };
}
