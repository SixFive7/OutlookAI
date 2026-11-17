using System.Security.Cryptography.X509Certificates;
using System.Text.RegularExpressions;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The installer's removal of the COMPROMISED signing certificate - the one that signed v1.0.0.1 to
/// v2.3.1.111, whose private key became public - read out of <c>Installer.iss</c>, which is Inno
/// Setup's Pascal and runs nowhere in this suite.
///
/// <para>
/// Decision 193 (b), the maintainer's, 2026-11-17: remove it by thumbprint from every certificate
/// trust store setup can reach - TrustedPublisher, as before, and Root and every other trust store,
/// in the scopes setup can write; ship no copy of it, because its subject names the maintainer's
/// computer and user, and let none enter the repository; never add it to Disallowed on a user's
/// computer. The thumbprint is the one identifier, here as in the installer.
/// </para>
/// </summary>
public sealed class InstallerCompromisedCertificateTests
{
    private const string Thumbprint = "E205060633DD7062D4F90033130542948A69D068";

    private static string Installer() => File.ReadAllText(Path.Combine(CachedModeSettingsTests.RepoRoot(), "Installer.iss"));

    private static string Define(string iss, string name)
    {
        Match m = Regex.Match(iss, "^#define " + Regex.Escape(name) + " \"([^\"]*)\"\\r?$", RegexOptions.Multiline);
        Assert.True(m.Success, "Installer.iss defines no " + name);
        return m.Groups[1].Value;
    }

    /// <summary>A [Code] procedure, from its heading to the "end;" that closes it at the start of a line.</summary>
    private static string Procedure(string iss, string name)
    {
        Match m = Regex.Match(iss, @"^procedure " + Regex.Escape(name) + @"\b.*?^end;", RegexOptions.Multiline | RegexOptions.Singleline);
        Assert.True(m.Success, "Installer.iss has no procedure " + name);
        return m.Value;
    }

    [Fact]
    public void TheCertificateIsNamedByItsThumbprintAlone_WrittenOnceInTheInstaller()
    {
        string iss = Installer();

        Assert.Equal(Thumbprint, Define(iss, "CompromisedCertThumbprint"));
        // Every use goes through the define, so the forty hex digits are written exactly once.
        Assert.Single(Regex.Matches(iss, Thumbprint, RegexOptions.IgnoreCase));
        string delete = Procedure(iss, "DeleteCompromisedCertificate");
        Assert.Contains("'-delstore ' + Store + ' {#CompromisedCertThumbprint}'", delete);
        Assert.Contains("'-store ' + Store + ' {#CompromisedCertThumbprint}'", delete);
    }

    [Fact]
    public void ItIsRemovedFromEveryTrustStore_TheUsersAlways_TheComputersWhenSetupCanWriteThem()
    {
        string iss = Installer();

        Assert.Equal(new[] { "TrustedPublisher", "TrustedPeople", "CA", "Root" }, Define(iss, "CompromisedCertUserStores").Split(' '));
        Assert.Equal(new[] { "TrustedPublisher", "TrustedPeople", "CA", "AuthRoot", "Root" }, Define(iss, "CompromisedCertMachineStores").Split(' '));

        string remove = Procedure(iss, "RemoveCompromisedCertificate");
        Assert.Contains("DeleteCompromisedCertificateFromStores('-user ', '{#CompromisedCertUserStores}');", remove);
        Assert.Matches(@"if IsAdmin then\s+DeleteCompromisedCertificateFromStores\('', '\{#CompromisedCertMachineStores\}'\);", remove);
    }

    [Fact]
    public void ARootStoreIsChangedOnlyWhenTheCertificateIsInIt_AndSetupNeverWaitsOnIt()
    {
        // Windows guards the user's Root store and can put up its own confirmation for a change to
        // it; a silent auto-update must not sit behind that.
        string delete = Procedure(Installer(), "DeleteCompromisedCertificate");

        Assert.Matches(@"if \(Store = 'Root'\) or \(Store = 'AuthRoot'\) then", delete);
        Assert.Matches(@"if ResultCode <> 0 then\s+exit;\s+Wait := ewNoWait;", delete);
        Assert.Contains("Wait := ewWaitUntilTerminated;", delete);
    }

    [Fact]
    public void EveryInstallUpdateAndUninstallRunsIt()
    {
        string iss = Installer();

        Assert.Contains("RemoveCompromisedCertificate;", Procedure(iss, "CurStepChanged"));
        Assert.Matches(@"if CurUninstallStep = usUninstall then\s+RemoveCompromisedCertificate;", Procedure(iss, "CurUninstallStepChanged"));
    }

    [Fact]
    public void NothingIsEverAddedToDisallowed_AndTheOneStoreAddedToIsTrustedPublisher()
    {
        string iss = Installer();

        Assert.DoesNotMatch(@"(?i)addstore\s+Disallowed", iss);
        Assert.Equal(new[] { "TrustedPublisher" },
                     Regex.Matches(iss, @"-addstore\s+(\w+)").Select(m => m.Groups[1].Value).Distinct().ToArray());
    }

    [Fact]
    public void NoReleaseIsSignedWithIt_AndTheUpdaterAcceptsNothingSignedWithIt()
    {
        string root = CachedModeSettingsTests.RepoRoot();

        Assert.DoesNotContain(Thumbprint, File.ReadAllText(Path.Combine(root, "OutlookAI.csproj")), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(Thumbprint, File.ReadAllText(Path.Combine(root, "Services", "UpdateService.cs")), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void TheRepositoryHoldsNoCopyOfIt()
    {
        string root = CachedModeSettingsTests.RepoRoot();
        var excluded = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { "bin", "obj", ".git", ".work", ".claude", ".vs", "node_modules", "TestResults" };
        var certificateFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".cer", ".crt", ".der" };
        var bundleOrKeyFiles = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { ".pfx", ".p12", ".p7b", ".p7c", ".sst", ".spc", ".pvk" };
        var problems = new List<string>();
        int pemBlocks = 0;

        foreach (string file in Files(root, excluded))
        {
            string relative = Path.GetRelativePath(root, file);
            string extension = Path.GetExtension(file);
            if (bundleOrKeyFiles.Contains(extension))
            {
                // A key or a bundle cannot be read here without its password, and has no place in a
                // public repository: the signing certificate left it on purpose.
                problems.Add(relative + ": a certificate bundle or private key file");
                continue;
            }

            var info = new FileInfo(file);
            if (info.Length > 4 * 1024 * 1024)
                continue;
            byte[] bytes = File.ReadAllBytes(file);
            if (certificateFiles.Contains(extension) && !LooksLikePem(bytes))
            {
                if (ThumbprintOf(bytes) is string thumbprint && string.Equals(thumbprint, Thumbprint, StringComparison.OrdinalIgnoreCase))
                    problems.Add(relative + ": IS the compromised certificate");
                continue;
            }

            // PEM anywhere - a .pem, a .cer in text form, or a certificate pasted into a script.
            string text = System.Text.Encoding.ASCII.GetString(bytes);
            foreach (Match block in Regex.Matches(text, "-----BEGIN CERTIFICATE-----(.*?)-----END CERTIFICATE-----", RegexOptions.Singleline))
            {
                pemBlocks++;
                byte[] der;
                try
                {
                    der = Convert.FromBase64String(Regex.Replace(block.Groups[1].Value, @"[^A-Za-z0-9+/=]", string.Empty));
                }
                catch (FormatException)
                {
                    continue;
                }
                if (ThumbprintOf(der) is string thumbprint && string.Equals(thumbprint, Thumbprint, StringComparison.OrdinalIgnoreCase))
                    problems.Add(relative + ": holds the compromised certificate as PEM");
            }
        }

        Assert.True(problems.Count == 0, string.Join(Environment.NewLine, problems));
    }

    private static IEnumerable<string> Files(string directory, HashSet<string> excluded)
    {
        foreach (string file in Directory.EnumerateFiles(directory))
            yield return file;
        foreach (string sub in Directory.EnumerateDirectories(directory))
        {
            if (excluded.Contains(Path.GetFileName(sub)))
                continue;
            foreach (string file in Files(sub, excluded))
                yield return file;
        }
    }

    private static bool LooksLikePem(byte[] bytes) =>
        System.Text.Encoding.ASCII.GetString(bytes, 0, Math.Min(bytes.Length, 64)).Contains("-----BEGIN", StringComparison.Ordinal);

    /// <summary>The SHA-1 thumbprint of a DER certificate, or null for bytes that are not one.</summary>
    private static string? ThumbprintOf(byte[] der)
    {
        try
        {
            using X509Certificate2 certificate = X509CertificateLoader.LoadCertificate(der);
            return certificate.Thumbprint;
        }
        catch (System.Security.Cryptography.CryptographicException)
        {
            return null;
        }
    }
}
