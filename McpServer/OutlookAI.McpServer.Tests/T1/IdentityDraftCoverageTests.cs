using System.Reflection;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the identity tests' refusal to report success having drafted in nothing.
/// <para>
/// <b>The defect, found 2026-08-24 by reading.</b> <c>LivePhase4Fixture.IdentityAccounts</c> is
/// "the configured primaries the write allowlist would let a draft into". Declaring the
/// measurement corpus and the tripwire's bystander store as BYSTANDERS - which is what keeps the
/// suite out of a corpus it would otherwise have drafted into and swept - empties that list on the
/// documented three-store VM layout. Its two consumers are <c>foreach</c> loops whose whole body is
/// the test, so both iterated nothing, asserted nothing and reported GREEN. That is the same shape
/// as the vacuous census this repository refused a day earlier: a check that cannot fail reads as
/// coverage in every report it appears in.
/// </para>
/// <para>
/// <b>Why the fix is pinned HERE.</b> Those two tests are <c>Category=Live</c>: CI has no Outlook,
/// no profile and no settings file, so it can never run them, and a cleverer live test would be a
/// promise nobody can check. So the decision - is there anything to test, and what does the run say
/// when there is not - lives in <see cref="IdentityDraftCoverage"/>, which is pure, and every branch
/// of it is exercised here. What CI cannot reach is one line per call site, and
/// <see cref="NeitherIdentityTestCanReachTheAccountListWithoutTheGuard"/> reads those two lines out
/// of the sources.
/// </para>
/// <para>
/// Synthetic store names, plus the COMMITTED example settings, which name placeholders only. No
/// machine-local settings file is read, nothing touches Outlook or a mailbox, and no real store
/// name is involved (S6).
/// </para>
/// </summary>
public sealed class IdentityDraftCoverageTests
{
    private const string Hub = "hub@example.test";
    private const string Business = "other@example.test";
    private const string SecondBusiness = "third@example.test";
    private const string Bystander = "OutlookAI Bystander";

    // ------------------------------------------------------------------ the VM layout

    [Fact]
    public void TheDocumentedVmLayoutGrantsNoIdentityAccountAtAll()
    {
        // Read off the committed example rather than retyped: this is the claim the whole file
        // rests on, and a hand-copied layout would keep passing after the real one changed.
        LiveTestSettings settings = Example();

        IdentityDraftCoverageReport coverage = IdentityDraftCoverage.Assess(settings);

        Assert.Empty(coverage.Accounts);
        Assert.True(coverage.ProvesNothing);
        Assert.False(coverage.Partial);
        Assert.Equal(2, coverage.NonHubStoreCount);
        Assert.Equal(settings.BystanderStoreDisplayNames.Count, coverage.Withheld.Count);
        Assert.Contains(settings.Corpus!.StoreDisplayName, coverage.Withheld, StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public void OnThatLayoutTheRunSaysProvedNothingInsteadOfPassingQuietly()
    {
        LiveTestSettings settings = Example();
        List<string> lines = new();

        IReadOnlyList<string> accounts = IdentityDraftCoverage.AccountsToDraftIn(
            settings, lines.Add, "the per-account identity draft");

        // Empty, and NOT a failure - this machine was never meant to run these tests.
        Assert.Empty(accounts);

        string provedNothing = Assert.Single(lines, l => l.StartsWith("PROVED NOTHING:", StringComparison.Ordinal));
        Assert.Contains("the per-account identity draft", provedNothing, StringComparison.Ordinal);
        Assert.Contains("declared BYSTANDER", provedNothing, StringComparison.Ordinal);
        Assert.Contains(settings.Corpus!.StoreDisplayName, provedNothing, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ the two profiles

    [Fact]
    public void AProductionProfileWithNoIdentityAccountRefusesTheRun()
    {
        // The other half of the same emptiness: on a real working profile an empty identity set
        // means the settings have drifted, and a test that shrugged would hide it. Same idiom as
        // LiveManageSignatureTests and LiveStaleIndexRowTests - throw here, say it there.
        LiveTestSettings settings = Settings(LiveMachineProfile.Production, new[] { Hub, Bystander }, Bystander);
        List<string> lines = new();

        InvalidOperationException refusal = Assert.Throws<InvalidOperationException>(
            () => IdentityDraftCoverage.AccountsToDraftIn(settings, lines.Add, "the per-account identity draft"));

        Assert.Contains(IdentityDraftCoverage.Population, refusal.Message, StringComparison.Ordinal);
        Assert.Contains("Production", refusal.Message, StringComparison.Ordinal);

        // It refused BEFORE claiming anything: the coverage line is out, the PROVED NOTHING line
        // is not, because nothing on a Production machine may read as an acceptable emptiness.
        Assert.DoesNotContain(lines, l => l.StartsWith("PROVED NOTHING:", StringComparison.Ordinal));
    }

    [Fact]
    public void APortableProfileWithNoIdentityAccountDoesNotRefuse()
    {
        LiveTestSettings settings = Settings(LiveMachineProfile.Portable, new[] { Hub, Bystander }, Bystander);
        List<string> lines = new();

        IReadOnlyList<string> accounts = IdentityDraftCoverage.AccountsToDraftIn(
            settings, lines.Add, "the A1 signature-placement matrix");

        Assert.Empty(accounts);
        Assert.Contains(lines, l => l.StartsWith("PROVED NOTHING:", StringComparison.Ordinal));
    }

    // ------------------------------------------------------------------ what a real run says

    [Fact]
    public void AMachineThatGrantsEveryPrimaryProceedsAndStillSaysHowMany()
    {
        // The coverage line is printed on EVERY run, not only the empty ones. A reader of a
        // passing identity test should not have to infer from the test's NAME how many accounts
        // it visited - that inference is exactly what was wrong before.
        LiveTestSettings settings = Settings(
            LiveMachineProfile.Production, new[] { Hub, Business, SecondBusiness }, bystander: null);
        List<string> lines = new();

        IReadOnlyList<string> accounts = IdentityDraftCoverage.AccountsToDraftIn(
            settings, lines.Add, "the per-account identity draft");

        Assert.Equal(new[] { Business, SecondBusiness }, accounts);
        string coverage = Assert.Single(lines);
        Assert.StartsWith("identity coverage: 2 of 2", coverage, StringComparison.Ordinal);
        Assert.DoesNotContain("PROVED NOTHING", coverage, StringComparison.Ordinal);
        Assert.DoesNotContain("PARTIAL", coverage, StringComparison.Ordinal);
    }

    [Fact]
    public void PartialCoverageIsAnnouncedAndNeverRefused()
    {
        // The judgement call. A machine that grants two of three IS exercising the identity path,
        // so calling that "proved nothing" would be false - and on a Production profile it would
        // refuse a run over a declaration somebody made on purpose. It gets a note, not a refusal.
        LiveTestSettings settings = Settings(
            LiveMachineProfile.Production, new[] { Hub, Business, SecondBusiness, Bystander }, Bystander);
        List<string> lines = new();

        IdentityDraftCoverageReport coverage = IdentityDraftCoverage.Assess(settings);
        IReadOnlyList<string> accounts = IdentityDraftCoverage.AccountsToDraftIn(
            settings, lines.Add, "the per-account identity draft");

        Assert.True(coverage.Partial);
        Assert.False(coverage.ProvesNothing);
        Assert.Equal(new[] { Business, SecondBusiness }, accounts);
        Assert.Equal(new[] { Bystander }, coverage.Withheld);

        string line = Assert.Single(lines);
        Assert.StartsWith("identity coverage: 2 of 3", line, StringComparison.Ordinal);
        Assert.Contains("PARTIAL", line, StringComparison.Ordinal);
        Assert.Contains("'" + Bystander + "' declared BYSTANDER", line, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ the arithmetic

    [Fact]
    public void TheAccountsAreExactlyWhatTheWriteAllowlistWouldPermit()
    {
        // Not a second opinion about which stores are writable: the granted half IS
        // IdentityAccountsAmong. Two derivations of that question is how a live test came to
        // throw at the guard halfway through a mailbox it was never entitled to touch.
        foreach (LiveTestSettings settings in new[]
                 {
                     Example(),
                     Settings(LiveMachineProfile.Production, new[] { Hub, Business, SecondBusiness }, null),
                     Settings(LiveMachineProfile.Portable, new[] { Hub, Business, Bystander }, Bystander),
                 })
        {
            Assert.Equal(
                LiveStoreWriteGuard.Build(settings).IdentityAccountsAmong(settings.ExpectedStoreDisplayNames),
                IdentityDraftCoverage.Assess(settings).Accounts);
        }
    }

    [Fact]
    public void AStoreTheAllowlistSimplyDoesNotGrantIsWithheldToo()
    {
        // Withholding is not only the bystander declaration: anything the allowlist would refuse a
        // draft is counted and named, so the coverage line cannot understate what was skipped.
        IdentityDraftCoverageReport coverage = IdentityDraftCoverage.Assess(
            new[] { Hub, Business }, new StoreWriteAllowlist(Hub));

        Assert.Empty(coverage.Accounts);
        Assert.Equal(new[] { Business }, coverage.Withheld);
        Assert.Contains("not granted a draft", coverage.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheHubIsNeitherGrantedNorWithheld()
    {
        // A hub-only machine has nothing to say about identity at all, and the line says so in
        // those words rather than reporting "0 of 0" and leaving the reader to work out why.
        IdentityDraftCoverageReport coverage = IdentityDraftCoverage.Assess(
            Settings(LiveMachineProfile.Portable, new[] { Hub }, bystander: null));

        Assert.Equal(0, coverage.NonHubStoreCount);
        Assert.Empty(coverage.Withheld);
        Assert.True(coverage.ProvesNothing);
        Assert.Contains(
            "declares none besides the hub", coverage.ProvedNothing("the identity draft"), StringComparison.Ordinal);
    }

    [Fact]
    public void ARepeatedStoreIsCountedOnce()
    {
        IdentityDraftCoverageReport coverage = IdentityDraftCoverage.Assess(
            new[] { Hub, Bystander, Bystander, Hub },
            new StoreWriteAllowlist(Hub, new[] { Bystander }, null, new[] { Bystander }));

        Assert.Equal(new[] { Bystander }, coverage.Withheld);
        Assert.Equal(1, coverage.NonHubStoreCount);
    }

    // ------------------------------------------------------------------ the two call sites

    [Fact]
    public void NeitherIdentityTestCanReachTheAccountListWithoutTheGuard()
    {
        // The one thing a pure function cannot pin: that the live tests still ASK it. Obtaining
        // the list now requires a sink to announce through and a name for what would not run, so
        // the old zero-argument property no longer compiles - but re-deriving the list from
        // IdentityAccountsAmong would, and would restore the silent green pass exactly.
        foreach (string file in new[] { "LiveDraftTests.cs", "LiveDraftOptionsTests.cs" })
        {
            string path = Path.Combine(TestProjectDir(), "T2", file);
            Assert.True(File.Exists(path), "identity test source is missing: " + path);
            string source = File.ReadAllText(path);

            Assert.Contains("_fixture.IdentityAccounts(", source, StringComparison.Ordinal);
            Assert.DoesNotContain("IdentityAccountsAmong", source, StringComparison.Ordinal);
        }
    }

    // ------------------------------------------------------------------ helpers

    private static LiveTestSettings Settings(
        LiveMachineProfile profile, IEnumerable<string> stores, string? bystander)
    {
        return new LiveTestSettings
        {
            MachineProfile = profile,
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = stores.ToList(),
            BystanderStoreDisplayNames = bystander == null ? new List<string>() : new List<string> { bystander },
        };
    }

    private static string TestProjectDir()
    {
        return typeof(LiveTestSettings).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
    }

    /// <summary>The committed example settings, parsed by the real loader.</summary>
    private static LiveTestSettings Example()
    {
        // <repo>/McpServer/OutlookAI.McpServer.Tests/ -> <repo>
        string path = Path.Combine(
            Path.GetFullPath(Path.Combine(TestProjectDir(), "..", "..")),
            "Testbed", "live-test-settings.example.json");
        Assert.True(File.Exists(path), "the committed example settings file is missing: " + path);
        return LiveTestSettings.Parse(File.ReadAllText(path));
    }
}
