using System.Reflection;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the post-run artifact sweep's refusal to point a delete at a store no test may write to.
///
/// <para>
/// <b>The defect.</b> <c>ArtifactSweep_AllThreeAccounts_ZeroTaggedRemain</c> - one copy in
/// <c>LiveDraftTests</c>, one in <c>LiveSendTests</c> - walked
/// <c>expectedStoreDisplayNames</c> and called <c>DeleteTaggedArtifactsUntilStableZero</c> on any
/// store whose tagged count came back above zero. That list also declares the measurement corpus
/// and the tripwire's bystander store: it must, or the census never visits them and
/// <c>list_accounts</c> exactness never counts them. Until the subject tags were split on
/// 2026-08-25 a corpus item's subject carried the very text the sweep matches, so the walk was one
/// non-zero count away from deleting twenty thousand real items. The split removed the match; it
/// did not remove the delete AIMED at those stores, and what stood between the two was a property
/// of the data.
/// </para>
///
/// <para>
/// <b>What is pinned here, and why here.</b> The decision - walk everything, delete only where the
/// allowlist permits it, and FAIL loudly on a count it may not act on - is
/// <see cref="ArtifactSweepPolicy"/>, which is pure. Its consumers are <c>Category=Live</c>: CI has
/// no Outlook, no profile and no settings file and can never run them, so a cleverer live test
/// would be a promise nobody checks. <see cref="NeverPurgesAStoreTheAllowlistRefuses"/> is the one
/// that matters: it drives the real loop with fakes and asserts the purge delegate was never
/// handed a store the allowlist refuses.
/// </para>
///
/// <para>
/// <b>Why the refusal is not left to the count tripwire.</b> That was the obvious alternative and
/// it does not work: the tripwire fires on a per-store item-count DECREASE, and an artifact
/// appearing in a bystander is an INCREASE. <see cref="TheRefusalSaysWhyNothingElseWouldCatchIt"/>
/// keeps that reasoning in the message a person actually reads.
/// </para>
///
/// <para>
/// Synthetic store names, plus the COMMITTED example settings, which name placeholders only. No
/// machine-local settings file is read, nothing touches Outlook or a mailbox, and no real store
/// name is involved (S6).
/// </para>
/// </summary>
public sealed class ArtifactSweepPlanTests
{
    private const string Hub = "hub@example.test";
    private const string Business = "other@example.test";
    private const string Bystander = "OutlookAI Bystander";
    private const string SecondBystander = "Corpus A";
    private const string Delegate = "Shared Mailbox";

    // ------------------------------------------------------------------ the VM layout

    [Fact]
    public void OnTheDocumentedVmLayoutOnlyTheHubIsSweptAndBothBystandersAreCounted()
    {
        // Read off the committed example rather than retyped: this is the layout the whole file
        // is about, and a hand-copied one would keep passing after the real one changed.
        LiveTestSettings settings = Example();
        ArtifactSweepPlan plan = ArtifactSweepPolicy.Assess(settings);

        Assert.Equal(LiveStoreCountTripwire.WatchedStores(settings).Count, plan.Steps.Count);
        Assert.Equal(new[] { settings.TestHubStoreDisplayName }, plan.Swept);
        Assert.Equal(settings.BystanderStoreDisplayNames.Count, plan.CountedOnly.Count);
        Assert.Contains(settings.Corpus!.StoreDisplayName, plan.CountedOnly, StringComparer.OrdinalIgnoreCase);
        Assert.All(plan.Steps.Where(s => !s.MayDelete), s => Assert.True(s.DeclaredBystander));
    }

    [Fact]
    public void TheCorpusStoreIsNeverHandedToADelete()
    {
        // The specific item this whole change exists for, stated as its own test so that a
        // regression names the corpus rather than "a bystander".
        LiveTestSettings settings = Example();
        string corpus = settings.Corpus!.StoreDisplayName;
        List<string> purged = new();

        // Every store reports artifacts, which is the worst case and the one that used to delete.
        Assert.Throws<InvalidOperationException>(
            () => ArtifactSweepPolicy.Run(
                ArtifactSweepPolicy.Assess(settings), _ => 20_000, purged.Add, _ => { }));

        Assert.DoesNotContain(corpus, purged, StringComparer.OrdinalIgnoreCase);
    }

    // ------------------------------------------------------------------ the safety property

    [Fact]
    public void NeverPurgesAStoreTheAllowlistRefuses()
    {
        // THE test. Everything else here is about what a run SAYS; this is about what it does.
        // The counts are non-zero everywhere, so the only thing keeping the bystander out of the
        // purge list is the plan - not the data, which is exactly the swap this change is.
        List<string> purged = new();
        List<string> counted = new();
        ArtifactSweepPlan plan = ArtifactSweepPolicy.Assess(
            new[] { Hub, Business, Bystander },
            new StoreWriteAllowlist(Hub, new[] { Business, Bystander }, null, new[] { Bystander }));

        Assert.ThrowsAny<InvalidOperationException>(
            () => ArtifactSweepPolicy.Run(
                plan,
                store =>
                {
                    counted.Add(store);
                    return 4;
                },
                purged.Add,
                _ => { }));

        // Counted everywhere - the bystander is WALKED, not skipped. Skipping it would give up
        // the detection, which was the option this decision rejected.
        Assert.Contains(Bystander, counted, StringComparer.OrdinalIgnoreCase);

        // Purged nowhere it was not entitled to.
        Assert.DoesNotContain(Bystander, purged, StringComparer.OrdinalIgnoreCase);
        Assert.Equal(new[] { Hub, Business }, purged.Distinct(StringComparer.OrdinalIgnoreCase));
    }

    [Fact]
    public void TheCountedStoresAreExactlyTheOnesTheAllowlistRefusesADelete()
    {
        // Not a second opinion about which stores are writable - the split IS the allowlist's
        // answer. Two derivations of "may we delete here" is how the aim survived the tag split.
        foreach (LiveTestSettings settings in EveryShape())
        {
            StoreWriteAllowlist allowlist = LiveStoreWriteGuard.Build(settings);
            ArtifactSweepPlan plan = ArtifactSweepPolicy.Assess(settings);
            IReadOnlyList<string> watched = LiveStoreCountTripwire.WatchedStores(settings);

            Assert.Equal(watched.Where(s => allowlist.IsAllowed(s, StoreWriteKind.Delete)), plan.Swept);
            Assert.Equal(watched.Where(s => !allowlist.IsAllowed(s, StoreWriteKind.Delete)), plan.CountedOnly);
        }
    }

    // ------------------------------------------------------------------ the walk set

    [Fact]
    public void TheSweepVisitsEveryStoreTheCountTripwireWatches()
    {
        // The second half of this decision, closed 2026-09-15. The sweep walked
        // expectedStoreDisplayNames, and that is NOT the watched set: the tripwire watches it
        // UNION the delegate/shared mailboxes UNION the declared bystanders. Every delegate store
        // was therefore censused for LOSS and never counted for ARRIVAL - and the argument is the
        // one this whole class rests on, unchanged: the tripwire fires on a DECREASE and an
        // artifact turning up is an INCREASE, so a store covered by only one of the two guards is
        // a store where one direction goes unreported.
        foreach (LiveTestSettings settings in EveryShape())
        {
            Assert.Equal(
                LiveStoreCountTripwire.WatchedStores(settings),
                ArtifactSweepPolicy.Assess(settings).Steps.Select(s => s.Store));
        }
    }

    [Fact]
    public void ADelegateMailboxIsCountedNeverSwept_AndTheRefusalCallsItOne()
    {
        // The store class this change actually adds, and the one it is most important to get
        // right: a delegate/shared mailbox is somebody else's mail. It must be visited (or an
        // artifact there is invisible), must never be deleted from, and must say WHICH kind of
        // off-limits it is - "the allowlist grants no delete" would read as a configuration gap
        // rather than as rule 3.
        LiveTestSettings settings = Settings(
            LiveMachineProfile.Production, new[] { Hub, Business }, bystander: null, delegates: new[] { Delegate });
        ArtifactSweepPlan plan = ArtifactSweepPolicy.Assess(settings);
        List<string> lines = new();

        Assert.Contains(Delegate, plan.CountedOnly, StringComparer.OrdinalIgnoreCase);
        Assert.DoesNotContain(Delegate, plan.Swept, StringComparer.OrdinalIgnoreCase);
        ArtifactSweepStep step = plan.Steps.Single(s => s.Store == Delegate);
        Assert.True(step.ReadOnlyDelegate);
        Assert.False(step.DeclaredBystander);

        string message = Assert.Throws<InvalidOperationException>(
            () => ArtifactSweepPolicy.Run(plan, store => store == Delegate ? 1 : 0, ShouldNotPurge, lines.Add))
            .Message;

        Assert.Contains(ArtifactSweepPolicy.Residue, message, StringComparison.Ordinal);
        Assert.Contains("DELEGATE/SHARED mailbox", message, StringComparison.Ordinal);
        Assert.Contains("mailbox-safety rule 3", message, StringComparison.Ordinal);
        Assert.DoesNotContain("declared BYSTANDER", message, StringComparison.Ordinal);

        // And it said so on the per-store line too, not only in the refusal.
        Assert.Contains("COUNTED, NOT SWEPT", Assert.Single(
            lines, l => l.StartsWith("sweep[" + Delegate + "]", StringComparison.Ordinal)), StringComparison.Ordinal);
    }

    [Fact]
    public void ABystanderDeclaredOnlyAmongTheDelegatesIsStillVisited()
    {
        // Permitted by BystanderCorpusDeclarationTests, which requires a declared bystander to be
        // in expectedStoreDisplayNames OR expectedDelegateStoreDisplayNames. Under the old walk
        // set the second spelling produced a bystander the sweep never looked at - the one store
        // in the whole configuration that exists to be looked at.
        LiveTestSettings settings = Settings(
            LiveMachineProfile.Production, new[] { Hub }, Delegate, delegates: new[] { Delegate });
        ArtifactSweepPlan plan = ArtifactSweepPolicy.Assess(settings);

        Assert.Equal(new[] { Hub, Delegate }, plan.Steps.Select(s => s.Store));
        Assert.Equal(new[] { Delegate }, plan.CountedOnly);

        // The bystander declaration outranks the delegate one: both are refusals, but one is a
        // decision somebody made about THIS store and the other is a whole tier's default.
        Assert.True(plan.Steps.Single(s => s.Store == Delegate).DeclaredBystander);
    }

    // ------------------------------------------------------------------ what a run says

    [Fact]
    public void ABystanderAtZeroPassesAndTheRunStatesThatItWasNotSwept()
    {
        // A run has to report what it did NOT do. A bare "taggedArtifacts=0" beside every other
        // store's "taggedArtifacts=0" claims a sweep that never happened.
        List<string> lines = new();

        ArtifactSweepPolicy.Run(Plan(), _ => 0, ShouldNotPurge, lines.Add);

        string bystanderLine = Assert.Single(lines, l => l.StartsWith("sweep[" + Bystander + "]", StringComparison.Ordinal));
        Assert.Contains("taggedArtifacts=0", bystanderLine, StringComparison.Ordinal);
        Assert.Contains("COUNTED, NOT SWEPT", bystanderLine, StringComparison.Ordinal);
        Assert.Contains("declared BYSTANDER", bystanderLine, StringComparison.Ordinal);

        // And the swept stores do not claim to have been left alone.
        Assert.DoesNotContain("COUNTED, NOT SWEPT", Assert.Single(
            lines, l => l.StartsWith("sweep[" + Hub + "]", StringComparison.Ordinal)), StringComparison.Ordinal);
    }

    [Fact]
    public void ThePlanLineSaysHowManyStoresWereLeftAloneAndWhichOnes()
    {
        List<string> lines = new();

        ArtifactSweepPolicy.Run(Plan(), _ => 0, ShouldNotPurge, lines.Add);

        string summary = lines[0];
        Assert.StartsWith("artifact sweep: 3 store(s) - 2 swept, 1 counted and left alone", summary, StringComparison.Ordinal);
        Assert.Contains("'" + Bystander + "'", summary, StringComparison.Ordinal);
    }

    [Fact]
    public void ASweptStoreIsPurgedOnceAndThenRecounted()
    {
        // The existing behaviour, unchanged: one purge pass for the documented sent-copy lag,
        // then the count is believed. A second purge would mean the loop had become a retry.
        Dictionary<string, int> counts = new(StringComparer.OrdinalIgnoreCase)
        {
            [Hub] = 3, [Business] = 0, [Bystander] = 0,
        };
        List<string> purged = new();
        List<string> lines = new();

        ArtifactSweepPolicy.Run(
            Plan(),
            store => counts[store],
            store =>
            {
                purged.Add(store);
                counts[store] = 0;
            },
            lines.Add);

        Assert.Equal(new[] { Hub }, purged);
        Assert.Contains(lines, l => l.Contains("late-materialized tagged artifact(s) found - purging", StringComparison.Ordinal));
        Assert.Contains(lines, l => l == "sweep[" + Hub + "]: taggedArtifacts=0");
    }

    [Fact]
    public void ACleanStoreIsNotPurgedAtAll()
    {
        ArtifactSweepPolicy.Run(Plan(), _ => 0, ShouldNotPurge, _ => { });
    }

    // ------------------------------------------------------------------ the refusals

    [Fact]
    public void ABystanderHoldingArtifactsFailsTheRunNamingTheStoreAndTheCount()
    {
        // The decision, in one assertion: counted, not deleted, and LOUD. Passing quietly was
        // never on the table, and neither was skipping - a skip gives up the detection entirely.
        Dictionary<string, int> counts = new(StringComparer.OrdinalIgnoreCase)
        {
            [Hub] = 0, [Business] = 0, [Bystander] = 7,
        };

        InvalidOperationException refusal = Assert.Throws<InvalidOperationException>(
            () => ArtifactSweepPolicy.Run(Plan(), store => counts[store], ShouldNotPurge, _ => { }));

        Assert.Contains(ArtifactSweepPolicy.Residue, refusal.Message, StringComparison.Ordinal);
        Assert.Contains("'" + Bystander + "'", refusal.Message, StringComparison.Ordinal);
        Assert.Contains("7 item(s)", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheRefusalSaysWhyNothingElseWouldCatchIt()
    {
        // The reasoning that made this a failure rather than a note. "The count tripwire would
        // notice" is the argument a future reader will reach for, and it is wrong: the tripwire
        // fires on a DECREASE. It is written where that reader will be standing.
        string message = Assert.Throws<InvalidOperationException>(
            () => ArtifactSweepPolicy.Run(Plan(), store => store == Bystander ? 1 : 0, ShouldNotPurge, _ => { }))
            .Message;

        Assert.Contains("DECREASE", message, StringComparison.Ordinal);
        Assert.Contains("increase", message, StringComparison.Ordinal);
        Assert.Contains("NOTHING HERE WILL DELETE THEM", message, StringComparison.Ordinal);
    }

    [Fact]
    public void ASweptStoreStillDirtyAfterItsPurgeFailsToo()
    {
        // The old Assert.Equal(0, count), kept. Moving the walk into a helper must not quietly
        // drop the assertion the helper was extracted from.
        InvalidOperationException refusal = Assert.Throws<InvalidOperationException>(
            () => ArtifactSweepPolicy.Run(Plan(), store => store == Hub ? 2 : 0, _ => { }, _ => { }));

        Assert.Contains("sweep[" + Hub + "]: 2 tagged artifact(s) still present", refusal.Message, StringComparison.Ordinal);
        Assert.DoesNotContain(ArtifactSweepPolicy.Residue, refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void EveryStoreIsVisitedEvenAfterOneHasAlreadyFailed()
    {
        // A live run costs minutes and a mailbox. Finding the second problem only after fixing
        // the first costs another one, so every reason goes in one message - the same rule
        // TripwireWatchReport.Refusal follows.
        List<string> counted = new();

        string message = Assert.Throws<InvalidOperationException>(
            () => ArtifactSweepPolicy.Run(
                ArtifactSweepPolicy.Assess(
                    new[] { Hub, Bystander, SecondBystander },
                    new StoreWriteAllowlist(Hub, null, null, new[] { Bystander, SecondBystander })),
                store =>
                {
                    counted.Add(store);
                    return store == Hub ? 0 : 1;
                },
                ShouldNotPurge,
                _ => { }))
            .Message;

        Assert.Equal(new[] { Hub, Bystander, SecondBystander }, counted);
        Assert.Contains("'" + Bystander + "'", message, StringComparison.Ordinal);
        Assert.Contains("'" + SecondBystander + "'", message, StringComparison.Ordinal);
    }

    [Fact]
    public void AStoreTheAllowlistSimplyDoesNotGrantIsCountedToo_AndSaysSoDifferently()
    {
        // Withholding is not only the bystander declaration. A store the allowlist would refuse
        // a delete is equally not something to point a delete at, and the two read differently
        // to whoever hits them: one is a decision somebody made, the other is a gap.
        ArtifactSweepPlan plan = ArtifactSweepPolicy.Assess(
            new[] { Hub, Business }, new StoreWriteAllowlist(Hub));
        List<string> lines = new();

        string message = Assert.Throws<InvalidOperationException>(
            () => ArtifactSweepPolicy.Run(plan, store => store == Business ? 1 : 0, ShouldNotPurge, lines.Add))
            .Message;

        Assert.Equal(new[] { Business }, plan.CountedOnly);
        Assert.False(plan.Steps.Single(s => s.Store == Business).DeclaredBystander);
        Assert.Contains("grants no delete on it", message, StringComparison.Ordinal);
        Assert.DoesNotContain("declared BYSTANDER", message, StringComparison.Ordinal);
    }

    [Fact]
    public void ASweepOverNoStoreAtAllIsRefusedRatherThanReportedClean()
    {
        // The vacuous shape this repository keeps finding: a check that could not have failed,
        // printing the output of a clean run. Same refusal as the vacuous census.
        ArtifactSweepPlan plan = ArtifactSweepPolicy.Assess(new[] { " ", string.Empty }, new StoreWriteAllowlist(Hub));

        Assert.True(plan.ProvesNothing);
        Assert.Contains(
            ArtifactSweepPolicy.NothingToVisit,
            Assert.Throws<InvalidOperationException>(
                () => ArtifactSweepPolicy.Run(plan, _ => 0, ShouldNotPurge, _ => { })).Message,
            StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ the arithmetic

    [Fact]
    public void ARepeatedStoreIsVisitedOnce()
    {
        // Counting a store twice would double a purge pass and, on a bystander, report the same
        // artifact as two findings.
        List<string> counted = new();

        ArtifactSweepPolicy.Run(
            ArtifactSweepPolicy.Assess(
                new[] { Hub, Bystander, Bystander, Hub },
                new StoreWriteAllowlist(Hub, null, null, new[] { Bystander })),
            store =>
            {
                counted.Add(store);
                return 0;
            },
            ShouldNotPurge,
            _ => { });

        Assert.Equal(new[] { Hub, Bystander }, counted);
    }

    [Fact]
    public void TheHubIsSweptEvenWhenSomebodyDeclaredItABystander()
    {
        // The allowlist answers for the hub first and cannot be told otherwise - that collision
        // is refused at the top of the run by TripwireWatchSoundness, which names it. This sweep
        // must not resolve it a second time and differently.
        ArtifactSweepPlan plan = ArtifactSweepPolicy.Assess(
            new[] { Hub }, new StoreWriteAllowlist(Hub, null, null, new[] { Hub }));

        Assert.Equal(new[] { Hub }, plan.Swept);
        Assert.Empty(plan.CountedOnly);
    }

    // ------------------------------------------------------------------ the two call sites

    [Fact]
    public void NeitherSweepStillWalksTheStoreListItself()
    {
        // The one thing a pure function cannot pin: that the live tests still ASK it. Re-deriving
        // the walk from ExpectedStoreDisplayNames compiles, and would restore the delete aimed at
        // a declared bystander exactly.
        foreach (string file in new[] { "LiveDraftTests.cs", "LiveSendTests.cs" })
        {
            string path = Path.Combine(TestProjectDir(), "T2", file);
            Assert.True(File.Exists(path), "artifact sweep source is missing: " + path);
            string source = File.ReadAllText(path);

            Assert.Contains("ArtifactSweepPolicy.Run(", source, StringComparison.Ordinal);
            Assert.DoesNotContain("in _fixture.Settings.ExpectedStoreDisplayNames", source, StringComparison.Ordinal);
        }
    }

    // ------------------------------------------------------------------ helpers

    /// <summary>A purge delegate that fails the test if anything reaches it.</summary>
    private static void ShouldNotPurge(string store)
    {
        Assert.Fail("the sweep purged '" + store + "', which nothing in this test entitles it to");
    }

    /// <summary>Hub + one granted business account + one declared bystander.</summary>
    private static ArtifactSweepPlan Plan()
    {
        return ArtifactSweepPolicy.Assess(
            new[] { Hub, Business, Bystander },
            new StoreWriteAllowlist(Hub, new[] { Business, Bystander }, null, new[] { Bystander }));
    }

    private static LiveTestSettings Settings(
        LiveMachineProfile profile,
        IEnumerable<string> stores,
        string? bystander,
        IEnumerable<string>? delegates = null)
    {
        return new LiveTestSettings
        {
            MachineProfile = profile,
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = stores.ToList(),
            ExpectedDelegateStoreDisplayNames = (delegates ?? []).ToList(),
            BystanderStoreDisplayNames = bystander == null ? new List<string>() : new List<string> { bystander },
        };
    }

    /// <summary>
    /// The configuration shapes every set-level assertion is made against: the committed example,
    /// a hub-plus-business machine, a machine with a declared bystander, and - the shapes the old
    /// walk set could not see - one with a delegate/shared mailbox and one whose bystander is
    /// declared only among the delegates.
    /// </summary>
    private static IEnumerable<LiveTestSettings> EveryShape()
    {
        yield return Example();
        yield return Settings(LiveMachineProfile.Production, new[] { Hub, Business }, bystander: null);
        yield return Settings(LiveMachineProfile.Portable, new[] { Hub, Business, Bystander }, Bystander);
        yield return Settings(
            LiveMachineProfile.Production, new[] { Hub, Business }, Bystander,
            delegates: new[] { Delegate, "Second Shared Mailbox" });
        yield return Settings(
            LiveMachineProfile.Production, new[] { Hub }, Delegate, delegates: new[] { Delegate });
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
