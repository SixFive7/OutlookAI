using System.Security.Cryptography.X509Certificates;

using OutlookAI.Core.AddIn;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The check an installer's signature has to pass before the add-in's updater runs it: Core's
/// <see cref="InstallerSignature"/>, which the add-in calls and <c>Tools/Publish-Release.ps1</c>
/// compiles with Add-Type, so that a release is held to the check every installed copy makes (Q171
/// (c)). The rule is pinned against WinVerifyTrust answers written out by hand, and the call against
/// real files: one with no signature, one that is not there, and this runtime's own core library,
/// which Microsoft signs.
///
/// <para>
/// WHAT WOULD BE LOST WITHOUT IT. Until this check was shared, nothing ran it outside Outlook: the
/// updater's copy was the add-in's, which no test can load, and the release script's was a copy of
/// its own. A rule that also accepted CERT_E_CHAINING, or a pin compared by culture, would still have
/// looked right, and only an installed copy refusing - or running - the wrong installer would have
/// shown it.
/// </para>
/// </summary>
public sealed class InstallerSignatureTests : IDisposable
{
    // A pin of the right shape. Which certificate it names does not matter to the rule; the updater's
    // own pin stays in Services/UpdateService.cs, where check-pinned-constants.ps1 #1 holds it.
    private const string Pin = "00112233445566778899AABBCCDDEEFF00112233";

    private const string WrongSigner = "Installer was not signed by the expected OutlookAI certificate.";

    private readonly string _dir = Path.Combine(Path.GetTempPath(), "OutlookAI-InstallerSignatureTests-" + Guid.NewGuid().ToString("N"));

    public InstallerSignatureTests()
    {
        Directory.CreateDirectory(_dir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_dir))
        {
            Directory.Delete(_dir, recursive: true);
        }
    }

    // ================================================================== the rule

    /// <summary>
    /// The two WinVerifyTrust answers the updater accepts: 0, and CERT_E_UNTRUSTEDROOT, which is what a
    /// valid signature by the self-signed release certificate gets on every machine.
    /// </summary>
    [Fact]
    public void ZeroAndTheUntrustedRoot_AreTheTwoAnswersTheUpdaterAccepts()
    {
        Assert.Equal("", InstallerSignature.TrustRefusal(0));
        Assert.Equal("", InstallerSignature.TrustRefusal(unchecked((int)0x800B0109)));
    }

    /// <summary>Every other answer is refused, and the refusal names it, in the words the update line shows.</summary>
    [Theory]
    [InlineData(0x80096010u, "80096010")] // TRUST_E_BAD_DIGEST: the file was changed after signing
    [InlineData(0x800B0100u, "800B0100")] // TRUST_E_NOSIGNATURE
    [InlineData(0x800B010Au, "800B010A")] // CERT_E_CHAINING: no chain to any root, trusted or not
    [InlineData(0x800B0101u, "800B0101")] // CERT_E_EXPIRED
    [InlineData(0x800B0003u, "800B0003")] // TRUST_E_SUBJECT_FORM_UNKNOWN: not a kind of file Windows signs
    [InlineData(0x80004005u, "80004005")] // E_FAIL: what the check answers when the call itself fails
    [InlineData(0x00000001u, "00000001")] // not an error code at all, and still not 0
    public void EveryOtherAnswer_IsRefused_AndNamed(uint answer, string hex)
    {
        Assert.Equal("Installer signature is invalid or missing (0x" + hex + ").", InstallerSignature.TrustRefusal(unchecked((int)answer)));
    }

    [Fact]
    public void ThePinnedSigner_IsAccepted_InEitherCase()
    {
        Assert.Equal("", InstallerSignature.SignerRefusal(Pin, Pin));
        Assert.Equal("", InstallerSignature.SignerRefusal(Pin.ToLowerInvariant(), Pin));
    }

    [Theory]
    [InlineData("")]
    [InlineData("00112233445566778899AABBCCDDEEFF00112234")]
    [InlineData("00112233445566778899AABBCCDDEEFF0011223")]
    [InlineData("00112233445566778899AABBCCDDEEFF00112233 ")]
    [InlineData("FFEEDDCCBBAA99887766554433221100FFEEDDCC")]
    public void AnyOtherSigner_IsRefused(string signer)
    {
        Assert.Equal(WrongSigner, InstallerSignature.SignerRefusal(signer, Pin));
    }

    // ================================================================== the call, on real files

    /// <summary>A file Windows can hold a signature in, holding none: refused, before any signer is read.</summary>
    [Fact]
    public void AFileWithNoSignature_IsRefused()
    {
        string unsigned = Path.Combine(_dir, "unsigned.ps1");
        File.WriteAllText(unsigned, "'not signed'\r\n");

        int answer = InstallerSignature.WinVerifyTrustFile(unsigned);

        Assert.NotEqual("", InstallerSignature.TrustRefusal(answer));
        Assert.StartsWith("Installer signature is invalid or missing (0x", InstallerSignature.Refusal(unsigned, Pin), StringComparison.Ordinal);
    }

    [Fact]
    public void AFileThatIsNotThere_IsRefused_WithoutThrowing()
    {
        string missing = Path.Combine(_dir, "missing.exe");

        Assert.StartsWith("Installer signature is invalid or missing (0x", InstallerSignature.Refusal(missing, Pin), StringComparison.Ordinal);
    }

    /// <summary>
    /// The call itself: a file Microsoft signed passes the first half of the rule - 0 where its root is
    /// in this machine's store, CERT_E_UNTRUSTEDROOT where it is not - and its signer is read and is
    /// not the pin. A call built wrong answers E_INVALIDARG or no signature here, and the first assert
    /// fails.
    /// </summary>
    [Fact]
    public void AFileMicrosoftSigned_PassesTheTrustHalf_AndIsRefusedForItsSigner()
    {
        string coreLibrary = typeof(object).Assembly.Location;

        int answer = InstallerSignature.WinVerifyTrustFile(coreLibrary);

        Assert.True(InstallerSignature.TrustRefusal(answer).Length == 0,
            "WinVerifyTrust answers 0x" + answer.ToString("X8", System.Globalization.CultureInfo.InvariantCulture) + " for " + coreLibrary
            + ", which Microsoft signs: the call is built wrong, or this runtime's core library carries no signature of its own.");
        Assert.Equal(WrongSigner, InstallerSignature.Refusal(coreLibrary, Pin));
    }

    /// <summary>The whole check accepts a signed file when the pin is its signer's, in either case.</summary>
    [Fact]
    public void AFileSignedByThePinnedCertificate_IsAccepted()
    {
        string coreLibrary = typeof(object).Assembly.Location;
        string signer;
        // The same reading the check makes; .NET 9 made it obsolete, for a loader that reads no signature.
#pragma warning disable SYSLIB0057
        using (X509Certificate2 cert = new X509Certificate2(X509Certificate.CreateFromSignedFile(coreLibrary)))
#pragma warning restore SYSLIB0057
        {
            signer = cert.Thumbprint;
        }

        Assert.Equal("", InstallerSignature.Refusal(coreLibrary, signer));
        Assert.Equal("", InstallerSignature.Refusal(coreLibrary, signer.ToLowerInvariant()));
    }
}
