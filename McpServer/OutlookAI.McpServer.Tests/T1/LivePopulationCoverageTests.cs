using System.Reflection;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the refusal of a live test to report success having iterated an empty set.
/// <para>
/// <b>The defect.</b> Reading every <c>foreach</c> in the live tier on 2026-08-25, after the
/// identity-draft pair was found doing exactly this, turned up three more places where a live test
/// walks a machine-dependent list with no non-empty guard. Two are fixed here, and they are
/// different shapes of one failure:
/// </para>
/// <list type="number">
/// <item><c>T2/LiveFolderScopeTests.DelegateFirstLevelFolders_StillResolve_AndTheWholeMailboxIsUnfiltered</c>
/// - a bare <c>foreach</c> over <c>expectedDelegateStoreDisplayNames</c>, which is <c>[]</c> on
/// every Portable machine including the VM. Its SIBLING two methods up has asserted that same list
/// is non-empty since the day it was written, so the omission was visibly an oversight.</item>
/// <item><c>T2/LiveSignatureTests.ListSignatures_SeesTestSignature_WithExcerpt_AndAccountRows</c> -
/// an <c>Assert.All</c> over the account rows <c>list_signatures</c> returned, which an EMPTY list
/// satisfies with no element ever examined.</item>
/// </list>
/// <para>
/// <b>Why it is pinned HERE.</b> Both callers are <c>Category=Live</c>: CI has no Outlook, no
/// profile and no settings file, so it can never run them, and a cleverer live test would be a
/// promise nobody can check. The decision - is there anything to test, and what does the run say
/// when there is not - lives in <see cref="LivePopulationCoverage"/>, which is pure, and every
/// branch of it is exercised here. What CI cannot reach is one line per call site, and
/// <see cref="BothCallSitesStillGoThroughTheGuard"/> reads those out of the sources.
/// </para>
/// <para>
/// Synthetic names throughout. Nothing touches Outlook, a mailbox or a machine-local settings file,
/// and no real store or account name appears (S6).
/// </para>
/// </summary>
public sealed class LivePopulationCoverageTests
{
    private const string Population = "a delegate or shared mailbox to resolve folders in";
    private const string WhatWouldNotRun = "the delegate first-level folder probe";
    private const string Remedy = "To exercise it, open one and name it in the settings.";

    // ------------------------------------------------------------------ the two profiles

    [Fact]
    public void APortableProfileWithNothingToIterateSaysSoInsteadOfPassingQuietly()
    {
        List<string> lines = new();

        IReadOnlyList<string> found = LivePopulationCoverage.Require(
            Settings(LiveMachineProfile.Portable), Array.Empty<string>(),
            Population, WhatWouldNotRun, Remedy, lines.Add);

        // Empty, and NOT a failure - this machine was never meant to run this test.
        Assert.Empty(found);

        string provedNothing = Assert.Single(lines, l => l.StartsWith("PROVED NOTHING:", StringComparison.Ordinal));
        Assert.Contains(WhatWouldNotRun, provedNothing, StringComparison.Ordinal);
        Assert.Contains(Population, provedNothing, StringComparison.Ordinal);
        Assert.Contains(Remedy, provedNothing, StringComparison.Ordinal);
    }

    [Fact]
    public void AProductionProfileWithNothingToIterateRefusesTheRun()
    {
        // The other half of the same emptiness: on the profile these tests were written for, an
        // empty population means the machine or the settings have drifted, and a test that
        // shrugged would hide it. Same idiom as IdentityDraftCoverage, LiveManageSignatureTests
        // and LiveStaleIndexRowTests - throw here, say it there.
        List<string> lines = new();

        InvalidOperationException refusal = Assert.Throws<InvalidOperationException>(
            () => LivePopulationCoverage.Require(
                Settings(LiveMachineProfile.Production), Array.Empty<string>(),
                Population, WhatWouldNotRun, Remedy, lines.Add));

        Assert.Contains(Population, refusal.Message, StringComparison.Ordinal);
        Assert.Contains("Production", refusal.Message, StringComparison.Ordinal);

        // It refused BEFORE announcing anything: the coverage line is out, the PROVED NOTHING line
        // is not, because nothing on a Production machine may read as an acceptable emptiness.
        Assert.Single(lines);
        Assert.DoesNotContain(lines, l => l.StartsWith("PROVED NOTHING:", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData(LiveMachineProfile.Portable)]
    [InlineData(LiveMachineProfile.Production)]
    public void APopulationThatEXISTSIsReturnedUnchangedOnEitherProfile(LiveMachineProfile profile)
    {
        string[] stores = { "delegate-one@example.test", "delegate-two@example.test" };
        List<string> lines = new();

        IReadOnlyList<string> found = LivePopulationCoverage.Require(
            Settings(profile), stores, Population, WhatWouldNotRun, Remedy, lines.Add);

        Assert.Equal(stores, found);

        // The coverage line is printed on EVERY run, not only the empty ones: a reader of a
        // passing test should not have to infer from its NAME how much it visited.
        string coverage = Assert.Single(lines);
        Assert.Contains("coverage: 2 " + Population, coverage, StringComparison.Ordinal);
        Assert.DoesNotContain("PROVED NOTHING", coverage, StringComparison.Ordinal);
    }

    [Fact]
    public void ANullPopulationIsTheSameAsAnEmptyOne()
    {
        // list_signatures returns a NULLABLE account list, so null and empty both reach this and
        // both mean "nothing was examined". Collapsing them here means the call site needs no
        // second opinion about which kind of nothing it has.
        List<string> lines = new();

        IReadOnlyList<string> found = LivePopulationCoverage.Require<string>(
            Settings(LiveMachineProfile.Portable), null,
            Population, WhatWouldNotRun, Remedy, lines.Add);

        Assert.Empty(found);
        Assert.Contains(lines, l => l.StartsWith("PROVED NOTHING:", StringComparison.Ordinal));
    }

    // ------------------------------------------------------------------ the lines themselves

    [Fact]
    public void TheCoverageLineNamesTheCountThePopulationAndTheTest()
    {
        LivePopulationReport report = LivePopulationCoverage.Assess(Population, WhatWouldNotRun, 3, Remedy);

        Assert.False(report.ProvesNothing);
        Assert.Equal(3, report.Found);
        Assert.Contains("coverage: 3 " + Population, report.Describe(), StringComparison.Ordinal);
        Assert.Contains(WhatWouldNotRun, report.Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheProvedNothingLineSaysWhatAGreenResultHereDoesAndDoesNotMean()
    {
        // The whole value of the line. "PROVED NOTHING" on its own is a label; what a reader needs
        // is that the assertions did not run, that green therefore means nothing, and what to do.
        string line = LivePopulationCoverage.Assess(Population, WhatWouldNotRun, 0, Remedy).ProvedNothing();

        Assert.StartsWith("PROVED NOTHING:", line, StringComparison.Ordinal);
        Assert.Contains("iterated nothing", line, StringComparison.Ordinal);
        Assert.Contains("a green result on this machine says only", line, StringComparison.Ordinal);
        Assert.Contains(Remedy, line, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("", WhatWouldNotRun, Remedy)]
    [InlineData("   ", WhatWouldNotRun, Remedy)]
    [InlineData(Population, "", Remedy)]
    [InlineData(Population, WhatWouldNotRun, "")]
    [InlineData(Population, WhatWouldNotRun, "  ")]
    public void ACoverageLineThatWouldNameNothingIsRefused(string population, string what, string remedy)
    {
        // A guard reads as coverage in every report it appears in, so it has to say WHICH
        // population and WHICH test, or the line it prints is worth less than nothing. An unnamed
        // population would also reach RequireProductionPopulation's refusal message as a blank.
        Assert.Throws<ArgumentException>(() => LivePopulationCoverage.Assess(population, what, 0, remedy));
    }

    [Fact]
    public void ANegativePopulationIsRefusedRatherThanReadAsEmpty()
    {
        Assert.Throws<ArgumentOutOfRangeException>(
            () => LivePopulationCoverage.Assess(Population, WhatWouldNotRun, -1, Remedy));
    }

    [Fact]
    public void ZeroIsTheOnlySizeThatProvesNothing()
    {
        Assert.True(LivePopulationCoverage.Assess(Population, WhatWouldNotRun, 0, Remedy).ProvesNothing);
        Assert.False(LivePopulationCoverage.Assess(Population, WhatWouldNotRun, 1, Remedy).ProvesNothing);
    }

    [Fact]
    public void TheSinkAndTheSettingsAreBothRequired()
    {
        Assert.Throws<ArgumentNullException>(() => LivePopulationCoverage.Require(
            null!, Array.Empty<string>(), Population, WhatWouldNotRun, Remedy, _ => { }));
        Assert.Throws<ArgumentNullException>(() => LivePopulationCoverage.Require(
            Settings(LiveMachineProfile.Portable), Array.Empty<string>(),
            Population, WhatWouldNotRun, Remedy, null!));
    }

    // ------------------------------------------------------------------ the call sites

    [Fact]
    public void BothCallSitesStillGoThroughTheGuard()
    {
        // The one thing a pure function cannot pin: that the live tests still ASK it. Unlike the
        // identity pair - where obtaining the list needs a sink, so the old property no longer
        // compiles - these two populations are a public settings property and a public tool
        // result, and nothing can stop a future edit reading them directly again. So this reads
        // the sources, which is the substitute this file's own precedent already uses.
        string folderScope = LiveSource("LiveFolderScopeTests.cs");

        // Exactly one read of the settings list in the whole file: the guard's own. Two would mean
        // the two delegate tests could answer the same question differently again, which is how
        // one of them came to have no guard at all.
        Assert.Equal(1, Occurrences(folderScope, "ExpectedDelegateStoreDisplayNames"));
        Assert.Equal(2, Occurrences(folderScope, "DelegateStores(\""));

        string signatures = LiveSource("LiveSignatureTests.cs");
        Assert.Contains("LivePopulationCoverage.Require(", signatures, StringComparison.Ordinal);

        // The exact line that was vacuous. Asserting over the raw nullable list is what an empty
        // list satisfied without examining anything.
        Assert.DoesNotContain("Assert.All(outcome.Accounts", signatures, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ helpers

    private static LiveTestSettings Settings(LiveMachineProfile profile)
    {
        return new LiveTestSettings
        {
            MachineProfile = profile,
            TestHubStoreDisplayName = "hub@example.test",
            ExpectedStoreDisplayNames = new List<string> { "hub@example.test" },
        };
    }

    private static int Occurrences(string haystack, string needle)
    {
        int count = 0;
        for (int i = haystack.IndexOf(needle, StringComparison.Ordinal);
             i >= 0;
             i = haystack.IndexOf(needle, i + needle.Length, StringComparison.Ordinal))
        {
            count++;
        }

        return count;
    }

    private static string LiveSource(string fileName)
    {
        string dir = typeof(LiveTestSettings).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
        string path = Path.Combine(dir, "T2", fileName);
        Assert.True(File.Exists(path), "live test source is missing: " + path);
        return File.ReadAllText(path);
    }
}
