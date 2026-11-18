using System.Text;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Soak fix D37 (signature steering): the pure filesystem + registry-parsing halves of
/// the signature catalog - name enumeration, excerpt extraction (BOM-aware), file
/// preference order, and the graceful-degradation contract of the per-account default
/// assignments (absent values = unknown, never guessed; non-mail rows filtered).
/// </summary>
public sealed class SignatureCatalogTests : IDisposable
{
    private readonly string _dir;

    public SignatureCatalogTests()
    {
        _dir = Path.Combine(Path.GetTempPath(), "OutlookAI-SigCatalogTests-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(_dir);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_dir, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    [Fact]
    public void ListSignatures_MissingDirectory_ReturnsEmpty()
    {
        Assert.Empty(SignatureCatalog.ListSignatures(Path.Combine(_dir, "does-not-exist")));
    }

    [Fact]
    public void ListSignatures_GroupsRenditions_SortsByName_AndPrefersHtml()
    {
        File.WriteAllText(Path.Combine(_dir, "Zeta.htm"), "<html><body><p>Zeta line</p></body></html>");
        File.WriteAllText(Path.Combine(_dir, "Zeta.txt"), "Zeta line");
        File.WriteAllText(Path.Combine(_dir, "Zeta.rtf"), @"{\rtf1 Zeta}");
        File.WriteAllText(Path.Combine(_dir, "Alpha.txt"), "Alpha only text");
        Directory.CreateDirectory(Path.Combine(_dir, "Zeta_files")); // resource dir must not become a signature

        IReadOnlyList<SignatureInfo> signatures = SignatureCatalog.ListSignatures(_dir);

        Assert.Equal(new[] { "Alpha", "Zeta" }, signatures.Select(s => s.Name).ToArray());
        SignatureInfo zeta = signatures[1];
        Assert.NotNull(zeta.HtmlPath);
        Assert.NotNull(zeta.RtfPath);
        Assert.NotNull(zeta.TextPath);
        Assert.Equal(zeta.HtmlPath, zeta.PreferredFilePath);
        SignatureInfo alpha = signatures[0];
        Assert.Null(alpha.HtmlPath);
        Assert.Equal(alpha.TextPath, alpha.PreferredFilePath);
    }

    [Fact]
    public void Excerpt_FromUtf16TextFile_FirstTwoNonEmptyLines()
    {
        // Outlook writes signature .txt files as UTF-16LE with BOM (observed live).
        string path = Path.Combine(_dir, "Dutch.txt");
        File.WriteAllText(path, "\r\nMet vriendelijke groet,\r\nAlex Example\r\n\r\nMore below", Encoding.Unicode);
        File.WriteAllText(Path.Combine(_dir, "Dutch.htm"), "<html><body>ignored - txt wins</body></html>");

        IReadOnlyList<SignatureInfo> signatures = SignatureCatalog.ListSignatures(_dir);

        Assert.Equal("Met vriendelijke groet, / Alex Example", Assert.Single(signatures).Excerpt);
    }

    [Fact]
    public void Excerpt_FallsBackToHtml_WhenNoTxt()
    {
        File.WriteAllText(Path.Combine(_dir, "HtmlOnly.htm"),
            "<html><head><style>p{}</style></head><body><p>Kind regards,</p><p>Team</p></body></html>");

        IReadOnlyList<SignatureInfo> signatures = SignatureCatalog.ListSignatures(_dir);

        string excerpt = Assert.Single(signatures).Excerpt!;
        Assert.Contains("Kind regards,", excerpt, StringComparison.Ordinal);
    }

    [Fact]
    public void Excerpt_IsCappedAtTheLimit()
    {
        File.WriteAllText(Path.Combine(_dir, "Long.txt"), new string('a', 500));

        string excerpt = Assert.Single(SignatureCatalog.ListSignatures(_dir)).Excerpt!;

        Assert.Equal(SignatureCatalog.ExcerptMaxChars, excerpt.Length);
    }

    [Fact]
    public void TryResolve_IsCaseInsensitive_AndNullForUnknown()
    {
        File.WriteAllText(Path.Combine(_dir, "My Sig.htm"), "<html><body>x</body></html>");

        Assert.NotNull(SignatureCatalog.TryResolve("my sig", _dir));
        Assert.NotNull(SignatureCatalog.TryResolve("  My Sig  ", _dir));
        Assert.Null(SignatureCatalog.TryResolve("Other", _dir));
        Assert.Null(SignatureCatalog.TryResolve("   ", _dir));
    }

    // ------------------------------------------------------------------ renditions, pictures, scope (Q147, Q159)

    [Fact]
    public void Signatures_CarryTheirFormats_PictureCount_AndTheAccountOfARoamingCopy()
    {
        File.WriteAllText(Path.Combine(_dir, "Work (jan@example.com).htm"), "<p>x</p>");
        File.WriteAllText(Path.Combine(_dir, "Logo.rtf"), @"{\rtf1 Logo}");
        File.WriteAllText(Path.Combine(_dir, "Logo.txt"), "Logo");
        Directory.CreateDirectory(Path.Combine(_dir, "Logo_files"));
        File.WriteAllBytes(Path.Combine(_dir, "Logo_files", "image001.png"), new byte[] { 1 });
        File.WriteAllBytes(Path.Combine(_dir, "Logo_files", "image002.JPG"), new byte[] { 2 });
        File.WriteAllText(Path.Combine(_dir, "Logo_files", "filelist.xml"), "<xml/>");

        IReadOnlyList<SignatureInfo> signatures = SignatureCatalog.ListSignatures(_dir);

        SignatureInfo logo = signatures.Single(s => s.Name == "Logo");
        Assert.Equal(new[] { "rtf", "text" }, logo.Formats);
        Assert.Equal(2, logo.PictureCount);
        Assert.Equal(Path.Combine(_dir, "Logo_files"), logo.ResourceDirectory);
        Assert.Null(logo.ScopeAccount);
        Assert.Equal("jan@example.com", signatures.Single(s => s.Name.StartsWith("Work", StringComparison.Ordinal)).ScopeAccount);
    }

    [Fact]
    public void Excerpt_FromAWordHtmlInWindows1252_DecodesItsAccents()
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        File.WriteAllBytes(Path.Combine(_dir, "Word.htm"), Encoding.GetEncoding(1252).GetBytes(
            "<html><head><meta http-equiv=Content-Type content=\"text/html; charset=windows-1252\"></head><body><p>Caf\u00e9 groet</p></body></html>"));

        Assert.Equal("Caf\u00e9 groet", Assert.Single(SignatureCatalog.ListSignatures(_dir)).Excerpt);
    }

    [Fact]
    public void Excerpt_FromTheRtf_WhenItIsTheOnlyRendition()
    {
        File.WriteAllText(Path.Combine(_dir, "RtfOnly.rtf"), @"{\rtf1\ansi{\fonttbl{\f0 Calibri;}}\f0 Caf\'e9\par Team}");

        Assert.Equal("Caf\u00e9 / Team", Assert.Single(SignatureCatalog.ListSignatures(_dir)).Excerpt);
    }

    [Fact]
    public void DecodeRegistryString_HandlesBothShapes()
    {
        Assert.Equal("plain", SignatureCatalog.DecodeRegistryString("plain\0"));
        Assert.Equal("bin", SignatureCatalog.DecodeRegistryString(Encoding.Unicode.GetBytes("bin\0")));
        Assert.Null(SignatureCatalog.DecodeRegistryString(null));
        Assert.Null(SignatureCatalog.DecodeRegistryString(7));
    }
}
