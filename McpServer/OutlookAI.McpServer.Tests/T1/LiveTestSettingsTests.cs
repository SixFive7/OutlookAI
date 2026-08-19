using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins what a machine must declare before the live tier will run on it.
/// <para>
/// The settings file is gitignored and machine-specific, so nothing about it can be checked
/// by reading the repository. These tests pin the VALIDATION instead, which is the part that
/// decides whether a second machine can be configured at all: before the machine profile
/// existed, every machine had to supply an index probe term and a hand-curated population of
/// real mail with a term in the subject and not in the body. A dedicated test machine has
/// neither, and a requirement that cannot be met honestly gets met dishonestly.
/// </para>
/// </summary>
public sealed class LiveTestSettingsTests
{
    private const string Hub = "hub@example.test";

    private static LiveTestSettings Portable()
    {
        return new LiveTestSettings
        {
            MachineProfile = LiveMachineProfile.Portable,
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, "Corpus Data File" },
        };
    }

    private static SubjectOnlyProbeSettings CompleteProbe()
    {
        return new SubjectOnlyProbeSettings
        {
            StoreDisplayName = "Some Store",
            FolderPath = "Deleted Items",
            SubjectTerm = "term",
            SenderFragment = "term",
        };
    }

    [Fact]
    public void ProductionIsTheDefault_SoAnOlderSettingsFileKeepsItsStrictValidation()
    {
        Assert.Equal(LiveMachineProfile.Production, new LiveTestSettings().MachineProfile);
    }

    [Fact]
    public void APortableMachine_NeedsOnlyAHubAndTheStoresToWatch()
    {
        // No probeTerm, no subjectOnlyProbe: a machine with no search index and no real mail
        // cannot supply either, and inventing them is how a test starts failing far away from
        // the mistake.
        LiveTestSettings.Validate(Portable());
    }

    [Fact]
    public void AProductionMachine_StillNeedsTheProbeTermAndTheSubjectOnlyPopulation()
    {
        LiveTestSettings settings = Portable();
        settings.MachineProfile = LiveMachineProfile.Production;

        InvalidOperationException noTerm = Assert.Throws<InvalidOperationException>(
            () => LiveTestSettings.Validate(settings));
        Assert.Contains("probeTerm", noTerm.Message, StringComparison.Ordinal);

        settings.ProbeTerm = "factuur";
        InvalidOperationException noProbe = Assert.Throws<InvalidOperationException>(
            () => LiveTestSettings.Validate(settings));
        Assert.Contains("subjectOnlyProbe", noProbe.Message, StringComparison.Ordinal);

        settings.SubjectOnlyProbe = CompleteProbe();
        LiveTestSettings.Validate(settings);
    }

    [Fact]
    public void TheHubAndTheWatchedStores_AreRequiredOnEveryMachine()
    {
        LiveTestSettings noHub = Portable();
        noHub.TestHubStoreDisplayName = " ";
        Assert.Throws<InvalidOperationException>(() => LiveTestSettings.Validate(noHub));

        LiveTestSettings noStores = Portable();
        noStores.ExpectedStoreDisplayNames = new List<string>();
        Assert.Throws<InvalidOperationException>(() => LiveTestSettings.Validate(noStores));
    }

    [Fact]
    public void AHalfWrittenProbeBlock_IsRefusedEvenOnAPortableMachine()
    {
        // Absent means "this machine has no such population". Three fields out of four means
        // somebody stopped halfway, and it would read as configured while behaving as absent.
        LiveTestSettings settings = Portable();
        settings.SubjectOnlyProbe = CompleteProbe();
        settings.SubjectOnlyProbe.SenderFragment = string.Empty;

        InvalidOperationException error = Assert.Throws<InvalidOperationException>(
            () => LiveTestSettings.Validate(settings));
        Assert.Contains("partially filled", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AMissingPopulation_FailsOnProduction_AndIsToleratedOnAPortableMachine()
    {
        LiveTestSettings portable = Portable();
        portable.RequireProductionPopulation("the delegate nested-folder probe");

        LiveTestSettings production = Portable();
        production.MachineProfile = LiveMachineProfile.Production;
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(
            () => production.RequireProductionPopulation("the delegate nested-folder probe"));
        Assert.Contains("delegate nested-folder probe", error.Message, StringComparison.Ordinal);
        Assert.Contains("Portable", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Describe_NamesTheProfileAndWhetherTheOptionalBlocksArePresent()
    {
        // Printed at the top of a live run, so a run pointed at the wrong machine's settings
        // says so before it touches a mailbox rather than being inferred afterwards.
        string portable = Portable().Describe();
        Assert.Contains("machineProfile=Portable", portable, StringComparison.Ordinal);
        Assert.Contains("probeTerm=none", portable, StringComparison.Ordinal);
        Assert.Contains("subjectOnlyProbe=none", portable, StringComparison.Ordinal);

        LiveTestSettings full = Portable();
        full.ProbeTerm = "factuur";
        full.SubjectOnlyProbe = CompleteProbe();
        Assert.Contains("probeTerm=set", full.Describe(), StringComparison.Ordinal);
        Assert.Contains("subjectOnlyProbe=set", full.Describe(), StringComparison.Ordinal);
    }
}
