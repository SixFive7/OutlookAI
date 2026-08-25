using System.Reflection;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins <c>Testbed/live-test-settings.example.json</c> - the committed shape of the gitignored
/// live-test settings - through the real loader, and pins the one declaration in it that stands
/// between the suite and the measurement corpus.
/// <para>
/// <b>Why the corpus store is declared a bystander.</b> No live test writes to a corpus:
/// <see cref="LiveCorpusFreshness"/> reads the manifest and never the store, and rebuilding a
/// stale corpus is an operator action from the accountless profile. But a corpus store has to be
/// listed in <c>expectedStoreDisplayNames</c> to be censused at all, and every non-hub entry of
/// that list is inside the identity-draft grant. Two consequences followed, both of them writes
/// INTO the corpus:
/// <list type="bullet">
/// <item>the identity tests create one draft per granted store, so they would draft into the
/// measurement corpus the moment that machine gains the mail account it is already getting;</item>
/// <item>the post-run artifact sweep counts subjects carrying <c>[OutlookAI-McpTest]</c> and
/// deletes what it finds - and <c>CorpusPlan.BuildSubject</c> used to put that exact tag at the
/// front of every corpus subject, so the sweep would have found the whole corpus and tried to
/// remove it.</item>
/// </list>
/// The declaration turns both into a refusal at the write guard.
/// </para>
/// <para>
/// <b>The sweep half is closed at source since 2026-08-25</b> - corpus items carry
/// <c>CorpusPlan.SubjectTag</c>, which is no longer the artifact tag, and
/// <c>T1/CorpusTagSeparationTests</c> holds them apart. That makes the declaration a
/// second line of defence for the sweep rather than the only one; it is still the ONLY thing
/// standing between the identity tests and the corpus, so it stays mandatory.
/// </para>
/// <para>
/// This file reads a COMMITTED example that names placeholder stores only. No machine-local
/// settings file is read, nothing touches Outlook or a mailbox, and no real store name is
/// involved (S6).
/// </para>
/// </summary>
public sealed class BystanderCorpusDeclarationTests
{
    /// <summary>The committed example, located from the same assembly metadata the loader uses.</summary>
    private static string ExamplePath()
    {
        string testProjectDir =
            typeof(LiveTestSettings).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");

        // <repo>/McpServer/OutlookAI.McpServer.Tests/ -> <repo>
        string repoRoot = Path.GetFullPath(Path.Combine(testProjectDir, "..", ".."));
        return Path.Combine(repoRoot, "Testbed", "live-test-settings.example.json");
    }

    private static LiveTestSettings Example()
    {
        string path = ExamplePath();
        Assert.True(File.Exists(path), "the committed example settings file is missing: " + path);
        return LiveTestSettings.Parse(File.ReadAllText(path));
    }

    [Fact]
    public void TheCommittedExampleStillLoadsThroughTheRealLoader()
    {
        // The example is what a rebuilder copies. A key it spells differently from the loader,
        // or a block the validator rejects, is a refusal on a machine nobody is watching - and
        // nothing but this test reads the file.
        LiveTestSettings settings = Example();

        Assert.Equal(LiveMachineProfile.Portable, settings.MachineProfile);
        Assert.NotEmpty(settings.TestHubStoreDisplayName);
        Assert.Contains(settings.TestHubStoreDisplayName, settings.ExpectedStoreDisplayNames);
        Assert.NotNull(settings.Corpus);
        Assert.True(settings.Corpus!.IsComplete);
    }

    [Fact]
    public void TheStoreTheCorpusBlockNamesIsDeclaredABystander()
    {
        // Derived from the corpus block rather than hard-coded, so renaming the corpus store in
        // the example cannot quietly drop it out of the declaration.
        LiveTestSettings settings = Example();
        string corpusStore = settings.Corpus!.StoreDisplayName;

        Assert.Contains(
            corpusStore, settings.BystanderStoreDisplayNames, StringComparer.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData(StoreWriteKind.Send)]
    [InlineData(StoreWriteKind.Draft)]
    [InlineData(StoreWriteKind.Delete)]
    [InlineData(StoreWriteKind.Move)]
    [InlineData(StoreWriteKind.Folder)]
    public void NothingMayWriteToTheCorpusStore(StoreWriteKind kind)
    {
        LiveTestSettings settings = Example();
        StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(settings);
        string corpusStore = settings.Corpus!.StoreDisplayName;

        Assert.False(allowlist.IsAllowed(corpusStore, kind));
        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
            () => allowlist.Assert(corpusStore, kind, "unit"));
        Assert.Contains("BYSTANDER", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheIdentityTestsWouldNotDraftIntoTheCorpus()
    {
        // The defect this declaration closes. LivePhase4Fixture.IdentityAccounts is exactly this
        // call, and on the documented VM layout the corpus store was the only answer it had.
        LiveTestSettings settings = Example();

        IReadOnlyList<string> accounts = LiveStoreWriteGuard.Build(settings)
            .IdentityAccountsAmong(settings.ExpectedStoreDisplayNames);

        Assert.DoesNotContain(settings.Corpus!.StoreDisplayName, accounts, StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public void TheCensusStillWatchesEveryDeclaredBystanderIncludingTheCorpus()
    {
        LiveTestSettings settings = Example();
        IReadOnlyList<string> watched = LiveStoreCountTripwire.WatchedStores(settings);

        foreach (string bystander in settings.BystanderStoreDisplayNames)
        {
            Assert.Contains(bystander, watched, StringComparer.OrdinalIgnoreCase);
        }

        Assert.Contains(settings.Corpus!.StoreDisplayName, watched, StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public void EveryDeclaredBystanderIsAStoreTheExampleProfileActuallyMounts()
    {
        // The trap the second corpus store sets. A declared bystander is unioned into the
        // watched set, so a name the profile does not mount is censused, not found, and refuses
        // the tier - which is why Corpus B, in the OTHER Windows account's profile, must be
        // declared in THAT machine's settings file and not in this one. Nothing in the tier can
        // tell those two mistakes apart, so the example must not model one.
        LiveTestSettings settings = Example();

        foreach (string bystander in settings.BystanderStoreDisplayNames)
        {
            Assert.Contains(
                bystander,
                settings.ExpectedStoreDisplayNames.Concat(settings.ExpectedDelegateStoreDisplayNames),
                StringComparer.OrdinalIgnoreCase);
        }
    }

    [Fact]
    public void TheExampleIsAConfigurationTheTripwireWouldAccept()
    {
        // End to end on the committed file: sound, and with stores the census could actually
        // fail on. If the example ever stops being runnable, this is where it shows.
        LiveTestSettings settings = Example();

        TripwireWatchReport report = TripwireWatchSoundness.Require(
            LiveStoreCountTripwire.WatchedStores(settings),
            LiveStoreWriteGuard.Build(settings),
            settings.BystanderStoreDisplayNames);

        Assert.True(report.Usable);
        Assert.False(report.ProvesNothing);
        Assert.Empty(report.Writable);
        Assert.Contains(settings.Corpus!.StoreDisplayName, report.Policed, StringComparer.OrdinalIgnoreCase);
        Assert.Equal(settings.BystanderStoreDisplayNames.Count, report.Policed.Count);
    }
}
