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
        File.WriteAllText(path, "\r\nMet vriendelijke groet,\r\nJori Huisman\r\n\r\nMore below", Encoding.Unicode);
        File.WriteAllText(Path.Combine(_dir, "Dutch.htm"), "<html><body>ignored - txt wins</body></html>");

        IReadOnlyList<SignatureInfo> signatures = SignatureCatalog.ListSignatures(_dir);

        Assert.Equal("Met vriendelijke groet, / Jori Huisman", Assert.Single(signatures).Excerpt);
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

    // ------------------------------------------------------------------ assignments

    [Fact]
    public void Assignments_ParseStringAndBinaryValues_AndFilterNonMailRows()
    {
        var rows = new List<IReadOnlyDictionary<string, object?>>
        {
            // Address book row - no '@', must be filtered.
            new Dictionary<string, object?> { ["Account Name"] = "Outlook Address Book" },
            // REG_SZ values.
            new Dictionary<string, object?>
            {
                ["Account Name"] = "a@example.com",
                ["New Signature"] = "Sig A",
                ["Reply-Forward Signature"] = "Sig B",
            },
            // REG_BINARY (UTF-16LE, NUL-terminated) values.
            new Dictionary<string, object?>
            {
                ["Account Name"] = Encoding.Unicode.GetBytes("b@example.com\0"),
                ["New Signature"] = Encoding.Unicode.GetBytes("Sig C\0"),
            },
            // Mail account without any signature values: unknown, reported as nulls.
            new Dictionary<string, object?> { ["Account Name"] = "c@example.com" },
        };

        IReadOnlyList<SignatureAssignment> assignments = SignatureCatalog.ReadAccountAssignments(() => rows);

        Assert.Equal(3, assignments.Count);
        Assert.Equal("a@example.com", assignments[0].Account);
        Assert.Equal("Sig A", assignments[0].NewMessageSignature);
        Assert.Equal("Sig B", assignments[0].ReplyForwardSignature);
        Assert.Equal("b@example.com", assignments[1].Account);
        Assert.Equal("Sig C", assignments[1].NewMessageSignature);
        Assert.Null(assignments[1].ReplyForwardSignature);
        Assert.Equal("c@example.com", assignments[2].Account);
        Assert.Null(assignments[2].NewMessageSignature);
        Assert.Null(assignments[2].ReplyForwardSignature);
    }

    [Fact]
    public void Assignments_ReaderThrow_DegradesToEmpty_NeverThrows()
    {
        IReadOnlyList<SignatureAssignment> assignments =
            SignatureCatalog.ReadAccountAssignments(() => throw new InvalidOperationException("registry unavailable"));

        Assert.Empty(assignments);
    }

    [Fact]
    public void Assignments_LiveRegistryRead_NeverThrows()
    {
        // Whatever this machine's profile registry contains, the read must not throw
        // (the Phase-4 degradation contract); shape-only assert, content-free (S4).
        IReadOnlyList<SignatureAssignment> assignments = SignatureCatalog.ReadAccountAssignments();

        Assert.All(assignments, a => Assert.Contains("@", a.Account, StringComparison.Ordinal));
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
