using System.Reflection;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the split of <c>expectedStoreDisplayNames</c> into a WATCHED list and an INDEXED list
/// (Q70, 2026-09-24), in both halves: what the loader accepts, and which consumer reads which list.
/// <para>
/// <b>What the split is for.</b> The one list meant two things. It was what the count tripwire
/// censuses, what the identity-draft grant is drawn from, what <c>list_accounts</c> exactness counts
/// and what <c>outlook_health</c> must reach - the WATCHED stores. And every index-tier test iterated
/// it and demanded each entry be discoverable in the Windows Search index with mail in it - the
/// INDEXED stores. A store can honestly be the first and not the second: the identity account's
/// store is written to and watched and holds almost nothing; every store on the unindexed guest is
/// watched and none is indexed. With one list, such a machine could only be described falsely.
/// </para>
/// <para>
/// <b>What it must not do</b> is loosen anything. Absent, the indexed list IS the watched list -
/// exactly what every index test read before - so a Production settings file written before the split
/// behaves as it did. Present, it may only name stores the census watches, never a delegate, and must
/// include the hub; empty is refused on Production and, on a Portable machine, turns every index test
/// into a refusal rather than an empty loop. The admission check (<see cref="TripwireWatchSoundness"/>)
/// does not read the indexed list at all, and one test below says so.
/// </para>
/// <para>
/// <b>The consumer half is pinned from the compiled IL</b>, the way this tier already pins other
/// wiring (<c>SweepSortWiringTests</c>, <c>TripwireReRunDriverTests</c>): a store-by-store index test
/// that went back to the watched list would pass every other test here and quietly start demanding an
/// index scope for the identity store.
/// </para>
/// </summary>
public sealed class StoreListSplitTests
{
    private const string Hub = "hub@split.invalid";
    private const string Corpus = "corpus@split.invalid";
    private const string Bystander = "bystander@split.invalid";
    private const string Identity = "identity@split.invalid";
    private const string Delegate = "Someone Else";

    private static LiveTestSettings Settings(List<string>? indexed, LiveMachineProfile profile = LiveMachineProfile.Portable)
    {
        return new LiveTestSettings
        {
            MachineProfile = profile,
            TestHubStoreDisplayName = Hub,
            ExpectedStoreDisplayNames = new List<string> { Hub, Bystander, Corpus, Identity },
            BystanderStoreDisplayNames = new List<string> { Bystander, Corpus },
            ExpectedDelegateStoreDisplayNames = new List<string> { Delegate },
            IndexedStoreDisplayNames = indexed,
            ProbeTerm = profile == LiveMachineProfile.Production ? "invoice" : string.Empty,
            SubjectOnlyProbe = profile == LiveMachineProfile.Production
                ? new SubjectOnlyProbeSettings { StoreDisplayName = Hub, FolderPath = "Inbox/x", SubjectTerm = "bulletin", SenderFragment = "noticebot" }
                : null,
        };
    }

    // ================================================================ the loader

    [Theory]
    [InlineData(LiveMachineProfile.Portable)]
    [InlineData(LiveMachineProfile.Production)]
    public void Absent_TheIndexedListIsTheWatchedList_ExactlyAsBeforeTheSplit(LiveMachineProfile profile)
    {
        LiveTestSettings settings = Settings(indexed: null, profile);
        LiveTestSettings.Validate(settings);
        Assert.Equal(settings.ExpectedStoreDisplayNames, settings.IndexedStores);
        Assert.Equal(settings.ExpectedStoreDisplayNames, settings.RequireIndexedStores());
    }

    [Fact]
    public void Present_ItIsReadAsWritten_InItsOrder()
    {
        LiveTestSettings settings = Settings(new List<string> { Hub, Bystander, Corpus });
        LiveTestSettings.Validate(settings);
        Assert.Equal(new[] { Hub, Bystander, Corpus }, settings.RequireIndexedStores());
        Assert.DoesNotContain(Identity, settings.IndexedStores);
        Assert.Contains(Identity, settings.ExpectedStoreDisplayNames);
    }

    [Fact]
    public void ABystanderNamedOnlyInTheBystanderList_MayBeIndexed_BecauseTheCensusWatchesIt()
    {
        LiveTestSettings settings = Settings(new List<string> { Hub, "Loose Bystander" });
        settings.BystanderStoreDisplayNames.Add("Loose Bystander");
        LiveTestSettings.Validate(settings);
        Assert.Contains("Loose Bystander", LiveStoreCountTripwire.WatchedStores(settings));
    }

    [Fact]
    public void AnIndexedStoreTheCensusDoesNotWatch_IsRefused()
    {
        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
            () => LiveTestSettings.Validate(Settings(new List<string> { Hub, "Nowhere Store" })));
        Assert.Contains("never censuses", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ADelegateMailboxInTheIndexedList_IsRefused()
    {
        LiveTestSettings settings = Settings(new List<string> { Hub, Delegate });
        settings.BystanderStoreDisplayNames.Add(Delegate);
        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() => LiveTestSettings.Validate(settings));
        Assert.Contains("delegate mailbox", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnIndexedListThatLeavesOutTheHub_IsRefused()
    {
        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
            () => LiveTestSettings.Validate(Settings(new List<string> { Bystander, Corpus })));
        Assert.Contains("not the hub", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ABlankOrRepeatedEntry_IsRefused()
    {
        Assert.Throws<InvalidOperationException>(() => LiveTestSettings.Validate(Settings(new List<string> { Hub, " " })));
        Assert.Throws<InvalidOperationException>(
            () => LiveTestSettings.Validate(Settings(new List<string> { Hub, Corpus, "CORPUS@split.invalid" })));
    }

    [Fact]
    public void Empty_IsTheTruthOnAPortableMachineWithNoIndex_AndEveryIndexTestThenRefusesRatherThanPasses()
    {
        LiveTestSettings settings = Settings(new List<string>());
        LiveTestSettings.Validate(settings);
        Assert.Empty(settings.IndexedStores);

        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(() => settings.RequireIndexedStores());
        Assert.Contains("Requires!=SearchIndex", ex.Message, StringComparison.Ordinal);
        Assert.Contains("never a reason to pass", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Empty_IsRefusedOnAProductionProfile()
    {
        InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
            () => LiveTestSettings.Validate(Settings(new List<string>(), LiveMachineProfile.Production)));
        Assert.Contains("EMPTY 'indexedStoreDisplayNames'", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheKeyIsReadFromJson_AndAbsentStaysAbsent()
    {
        const string withList = """
            { "machineProfile": "Portable", "testHubStoreDisplayName": "hub@split.invalid",
              "expectedStoreDisplayNames": [ "hub@split.invalid", "bystander@split.invalid" ],
              "bystanderStoreDisplayNames": [ "bystander@split.invalid" ],
              "indexedStoreDisplayNames": [ "hub@split.invalid", "bystander@split.invalid" ] }
            """;
        LiveTestSettings parsed = LiveTestSettings.Parse(withList);
        Assert.Equal(new[] { Hub, Bystander }, parsed.IndexedStoreDisplayNames);

        const string without = """
            { "machineProfile": "Portable", "testHubStoreDisplayName": "hub@split.invalid",
              "expectedStoreDisplayNames": [ "hub@split.invalid", "bystander@split.invalid" ],
              "bystanderStoreDisplayNames": [ "bystander@split.invalid" ] }
            """;
        LiveTestSettings bare = LiveTestSettings.Parse(without);
        Assert.Null(bare.IndexedStoreDisplayNames);
        Assert.Equal(bare.ExpectedStoreDisplayNames, bare.IndexedStores);
    }

    [Fact]
    public void Describe_SaysWhetherTheIndexedListIsDeclared()
    {
        Assert.Contains("indexed=as-watched(4)", Settings(null).Describe(), StringComparison.Ordinal);
        Assert.Contains("indexed=3", Settings(new List<string> { Hub, Bystander, Corpus }).Describe(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheAdmissionCheck_DoesNotReadTheIndexedList_AtAll()
    {
        // The census, the write allowlist and the bystander declaration decide admission; the split
        // may not move it in either direction. The same settings admit identically whatever the
        // indexed list says.
        TripwireWatchReport without = Admit(Settings(null));
        TripwireWatchReport with = Admit(Settings(new List<string> { Hub, Bystander, Corpus }));
        TripwireWatchReport empty = Admit(Settings(new List<string>()));
        foreach (TripwireWatchReport report in new[] { with, empty })
        {
            Assert.Equal(without.Policed, report.Policed);
            Assert.Equal(without.Writable, report.Writable);
            Assert.Equal(without.Usable, report.Usable);
        }

        Assert.Empty(CallsTo(typeof(TripwireWatchSoundness), nameof(LiveTestSettings.IndexedStores), nameof(LiveTestSettings.RequireIndexedStores)));
    }

    private static TripwireWatchReport Admit(LiveTestSettings settings)
        => TripwireWatchSoundness.Require(
            LiveStoreCountTripwire.WatchedStores(settings), LiveStoreWriteGuard.Build(settings), settings.BystanderStoreDisplayNames);

    // ================================================================ which consumer reads which list

    [Theory]
    [InlineData(typeof(LiveAttachmentKindRecallTests))]
    [InlineData(typeof(LiveDecodeVerifyTests))]
    [InlineData(typeof(LiveIndexSearchTests))]
    [InlineData(typeof(LiveOrderKeyCollationTests))]
    public void AStoreByStoreIndexTest_ReadsTheIndexedListAndNeverTheWatchedOne(Type testClass)
    {
        Assert.Empty(CallsTo(testClass, nameof(LiveTestSettings.ExpectedStoreDisplayNames)));
        Assert.NotEmpty(CallsTo(testClass, nameof(LiveTestSettings.RequireIndexedStores)));
    }

    [Fact]
    public void LiveMailServiceTests_ReadsTheWatchedListOnlyWhereItPinsTheAccountSet()
    {
        List<string> readers = CallsTo(typeof(LiveMailServiceTests), nameof(LiveTestSettings.ExpectedStoreDisplayNames))
            .Select(m => m.Name).Distinct().ToList();
        Assert.Equal(new[] { nameof(LiveMailServiceTests.ListAccounts_ExactAccountsDelegatesAndFlags) }, readers);
        Assert.NotEmpty(CallsTo(typeof(LiveMailServiceTests), nameof(LiveTestSettings.RequireIndexedStores)));
    }

    [Fact]
    public void TheExcludeSubfoldersMeasurement_PicksItsStoreFromTheIndexedList()
    {
        List<MethodBase> readers = CallsTo(typeof(LiveFolderScopeTests), nameof(LiveTestSettings.RequireIndexedStores));
        Assert.Contains(readers, m => m.Name == nameof(LiveFolderScopeTests.PrimaryStore_ExcludeSubfolders_NarrowsExactly_AndCostsNothing));
        Assert.Empty(CallsTo(typeof(LiveFolderScopeTests), nameof(LiveTestSettings.ExpectedStoreDisplayNames)));
    }

    [Fact]
    public void TheIndexFixtureDiscoversTheIndexedStores_NotTheWatchedOnes()
    {
        Assert.NotEmpty(CallsTo(typeof(LivePhase1Fixture), nameof(LiveTestSettings.IndexedStores)));
        Assert.Empty(CallsTo(typeof(LivePhase1Fixture), nameof(LiveTestSettings.ExpectedStoreDisplayNames)));
    }

    [Theory]
    [InlineData(typeof(LiveStoreCountTripwire))]
    [InlineData(typeof(LiveStoreWriteGuard))]
    [InlineData(typeof(IdentityDraftCoverage))]
    [InlineData(typeof(LiveHealthTests))]
    [InlineData(typeof(LiveMoveArchiveTests))]
    public void TheWatchedListConsumers_ReadTheWatchedListAndNeverTheIndexedOne(Type consumer)
    {
        // The census, the write allowlist, the identity grant, health's store reachability and the
        // archive resolution over every store: each is about the stores the profile HAS.
        Assert.NotEmpty(CallsTo(consumer, nameof(LiveTestSettings.ExpectedStoreDisplayNames)));
        Assert.Empty(CallsTo(consumer, nameof(LiveTestSettings.IndexedStores), nameof(LiveTestSettings.RequireIndexedStores)));
    }

    /// <summary>
    /// Every method of <paramref name="type"/> - its nested compiler-generated closures included -
    /// whose IL calls a <see cref="LiveTestSettings"/> member named in <paramref name="members"/>
    /// (a property's getter or a method). Tokens are RESOLVED, as in <c>SweepSortWiringTests</c>,
    /// so a stray operand byte cannot pass for a call.
    /// </summary>
    private static List<MethodBase> CallsTo(Type type, params string[] members)
    {
        var wanted = new HashSet<string>(members.SelectMany(m => new[] { m, "get_" + m }), StringComparer.Ordinal);
        var found = new List<MethodBase>();
        foreach (Type t in WithNested(type))
        {
            const BindingFlags all = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance
                | BindingFlags.Static | BindingFlags.DeclaredOnly;
            foreach (MethodBase method in t.GetMethods(all).Cast<MethodBase>().Concat(t.GetConstructors(all)))
            {
                byte[]? il = method.GetMethodBody()?.GetILAsByteArray();
                if (il == null)
                {
                    continue;
                }

                for (int i = 0; i + 4 < il.Length; i++)
                {
                    if (il[i] != 0x28 && il[i] != 0x6F)
                    {
                        continue;
                    }

                    int token = il[i + 1] | (il[i + 2] << 8) | (il[i + 3] << 16) | (il[i + 4] << 24);
                    try
                    {
                        MethodBase? resolved = method.Module.ResolveMethod(token);
                        if (resolved?.DeclaringType == typeof(LiveTestSettings) && wanted.Contains(resolved.Name))
                        {
                            found.Add(OwnerOf(method, type));
                        }
                    }
                    catch (ArgumentException)
                    {
                        // Not a real call - the bytes happened to look like one.
                    }
                }
            }
        }

        return found;
    }

    /// <summary>
    /// The declared method a closure belongs to, recovered from the compiler's naming
    /// (<c>&lt;Owner&gt;b__0_0</c>), so a read inside a lambda is charged to the test it is in.
    /// </summary>
    private static MethodBase OwnerOf(MethodBase method, Type declaring)
    {
        if (method.DeclaringType == declaring || !method.Name.StartsWith('<'))
        {
            return method;
        }

        string owner = method.Name[1..method.Name.IndexOf('>')];
        return declaring.GetMethod(owner, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static)
            ?? method;
    }

    private static IEnumerable<Type> WithNested(Type type)
    {
        yield return type;
        foreach (Type nested in type.GetNestedTypes(BindingFlags.Public | BindingFlags.NonPublic))
        {
            foreach (Type inner in WithNested(nested))
            {
                yield return inner;
            }
        }
    }
}
