using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// D38: the signature-suite snapshot guard itself must be trustworthy - these
/// CI-safe tests prove (against temp directories) that it detects every kind of
/// non-test change (modified, removed, added - top-level and nested), that it
/// exempts EXACTLY the "OutlookAI-McpTest-" prefixed entries, and that a missing
/// directory is a valid empty snapshot while an unreadable one refuses the suite.
/// </summary>
public sealed class SignatureDirectorySnapshotTests : IDisposable
{
    private readonly string _dir;

    public SignatureDirectorySnapshotTests()
    {
        _dir = Path.Combine(Path.GetTempPath(), "OutlookAI-SigSnapshotTests-" + Guid.NewGuid().ToString("N"));
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
    public void UnchangedDirectory_VerifiesClean()
    {
        File.WriteAllText(Path.Combine(_dir, "Real Sig.htm"), "<p>real</p>");
        SignatureDirectorySnapshot before = SignatureDirectorySnapshot.Capture(_dir);

        before.VerifyRealSignaturesUntouched();
        Assert.Empty(before.DiffIgnoringTestEntries(SignatureDirectorySnapshot.Capture(_dir)));
    }

    [Fact]
    public void ChangedRemovedAndAddedRealEntries_AreAllDetected()
    {
        File.WriteAllText(Path.Combine(_dir, "Keep.htm"), "unchanged");
        File.WriteAllText(Path.Combine(_dir, "Change.htm"), "before");
        File.WriteAllText(Path.Combine(_dir, "Remove.htm"), "doomed");
        Directory.CreateDirectory(Path.Combine(_dir, "Nested_files"));
        File.WriteAllText(Path.Combine(_dir, "Nested_files", "img.png"), "img-before");
        SignatureDirectorySnapshot before = SignatureDirectorySnapshot.Capture(_dir);

        File.WriteAllText(Path.Combine(_dir, "Change.htm"), "AFTER");
        File.Delete(Path.Combine(_dir, "Remove.htm"));
        File.WriteAllText(Path.Combine(_dir, "Added.htm"), "new");
        File.WriteAllText(Path.Combine(_dir, "Nested_files", "img.png"), "img-AFTER");

        IReadOnlyList<string> diff = before.DiffIgnoringTestEntries(SignatureDirectorySnapshot.Capture(_dir));
        Assert.Contains("CHANGED: Change.htm", diff);
        Assert.Contains("REMOVED: Remove.htm", diff);
        Assert.Contains("ADDED: Added.htm", diff);
        Assert.Contains(diff, d => d.StartsWith("CHANGED: Nested_files", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain(diff, d => d.Contains("Keep.htm", StringComparison.OrdinalIgnoreCase));

        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(before.VerifyRealSignaturesUntouched);
        Assert.Contains("SIGNATURE GUARD VIOLATION", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void SameLengthContentChange_IsDetectedByHash()
    {
        File.WriteAllText(Path.Combine(_dir, "Subtle.htm"), "aaaa");
        SignatureDirectorySnapshot before = SignatureDirectorySnapshot.Capture(_dir);

        File.WriteAllText(Path.Combine(_dir, "Subtle.htm"), "aaab");

        Assert.Contains("CHANGED: Subtle.htm",
            before.DiffIgnoringTestEntries(SignatureDirectorySnapshot.Capture(_dir)));
    }

    [Fact]
    public void TestPrefixedEntries_MayComeAndGo_TopLevelAndNested()
    {
        File.WriteAllText(Path.Combine(_dir, "Real.htm"), "real");
        string preExistingTest = Path.Combine(_dir, SignatureCatalog.TestSignaturePrefix + "Old.htm");
        File.WriteAllText(preExistingTest, "old test leftover");
        SignatureDirectorySnapshot before = SignatureDirectorySnapshot.Capture(_dir);

        // Test entries may be created, changed AND deleted without tripping the guard.
        File.WriteAllText(Path.Combine(_dir, SignatureCatalog.TestSignaturePrefix + "New.htm"), "new test");
        string testResources = Path.Combine(_dir, SignatureCatalog.TestSignaturePrefix + "New_files");
        Directory.CreateDirectory(testResources);
        File.WriteAllText(Path.Combine(testResources, "img.png"), "test img");
        File.Delete(preExistingTest);

        Assert.Empty(before.DiffIgnoringTestEntries(SignatureDirectorySnapshot.Capture(_dir)));
        before.VerifyRealSignaturesUntouched();
    }

    [Fact]
    public void MissingDirectory_IsAValidEmptySnapshot()
    {
        string missing = Path.Combine(_dir, "does-not-exist");
        SignatureDirectorySnapshot snapshot = SignatureDirectorySnapshot.Capture(missing);
        Assert.Empty(snapshot.HashesByRelativePath);
        snapshot.VerifyRealSignaturesUntouched();
    }

    [Fact]
    public void TestEntryClassification_MatchesThePinnedPrefix()
    {
        Assert.True(SignatureDirectorySnapshot.IsTestEntry(SignatureCatalog.TestSignaturePrefix + "Sig.htm"));
        Assert.True(SignatureDirectorySnapshot.IsTestEntry(
            Path.Combine(SignatureCatalog.TestSignaturePrefix + "Sig_files", "img.png")));
        Assert.False(SignatureDirectorySnapshot.IsTestEntry("Some Person.htm"));
        Assert.False(SignatureDirectorySnapshot.IsTestEntry(Path.Combine("Company Sig_files", "logo.png")));
        Assert.False(SignatureDirectorySnapshot.IsTestEntry("Contains-OutlookAI-McpTest-Inside.htm"));
    }
}
