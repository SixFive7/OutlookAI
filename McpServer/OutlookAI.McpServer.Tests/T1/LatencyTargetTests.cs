using System.Reflection;
using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tests.T2;
using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins two decisions the maintainer made on 2026-09-24 about what the index tier's live tests
/// MEASURE on a test guest, where CI can never run them:
/// <list type="number">
/// <item>(B, option c) every latency bound that used to be timed on the FIRST indexed store - on a guest
/// the few-dozen-item hub, where two seconds is met by construction - is timed on the LARGEST
/// (<see cref="LiveLatencyTarget"/>). Pinned twice: the chooser's semantics, and from the compiled IL,
/// that each of those tests takes its target from the chooser and has no other way to time a bound.</item>
/// <item>(D, option a) the hub is rebuilt before every run, and the frontier test FAILS on a stale one
/// rather than passing having proved nothing (<see cref="LiveHubPopulationFreshness"/>).</item>
/// </list>
/// Pure: no Outlook, no index, no settings file.
/// </summary>
public sealed class LatencyTargetTests
{
    // ================================================================ the chooser

    [Fact]
    public void TheLargestStoreIsChosen_WhereverTheIndexedListPutsIt()
    {
        Assert.Equal("corpus", LiveLatencyTarget.Largest(new[]
        {
            new LiveStoreSize("hub@vm.invalid", 68), new LiveStoreSize("bystander@vm.invalid", 342), new LiveStoreSize("corpus", 160_000),
        }));
        Assert.Equal("corpus", LiveLatencyTarget.Largest(new[]
        {
            new LiveStoreSize("corpus", 160_000), new LiveStoreSize("hub@vm.invalid", 68),
        }));
    }

    [Fact]
    public void TheFirstStore_IsNeverChosenWhenAnotherIsLarger_EvenByOne()
    {
        // The regression this exists for: the latency half going back to the first store - the hub.
        Assert.Equal("second", LiveLatencyTarget.Largest(new[] { new LiveStoreSize("first", 1_000), new LiveStoreSize("second", 1_001) }));
    }

    [Fact]
    public void ATie_GoesToTheStoreTheIndexedListNamesFirst_SoTheChoiceCannotFlap()
    {
        Assert.Equal("a", LiveLatencyTarget.Largest(new[] { new LiveStoreSize("a", 5), new LiveStoreSize("b", 5) }));
        Assert.Equal("b", LiveLatencyTarget.Largest(new[] { new LiveStoreSize("b", 5), new LiveStoreSize("a", 5) }));
    }

    [Fact]
    public void AStoreThatCouldNotBeCounted_OrWasOnlyPartlyCounted_RefusesRatherThanGuesses()
    {
        // A latency bound timed on a store that merely LOOKED largest is the silent weakening refused here.
        InvalidOperationException unknown = Assert.Throws<InvalidOperationException>(
            () => LiveLatencyTarget.Largest(new[] { new LiveStoreSize("hub", 68), new LiveStoreSize("corpus", null) }));
        Assert.Contains("'corpus' could not be counted", unknown.Message, StringComparison.Ordinal);

        InvalidOperationException partial = Assert.Throws<InvalidOperationException>(
            () => LiveLatencyTarget.Largest(new[] { new LiveStoreSize("hub", 68), new LiveStoreSize("corpus", 90_000, Complete: false) }));
        Assert.Contains("only partly counted", partial.Message, StringComparison.Ordinal);

        Assert.Throws<InvalidOperationException>(() => LiveLatencyTarget.Largest(Array.Empty<LiveStoreSize>()));
    }

    [Fact]
    public void AStoresSize_IsItsOwnFoldersItemCountsSummed_AndACutShortWalkIsMarked()
    {
        var tree = new ComFolderTree(new[]
        {
            new ComFolderInfo("corpus", "Inbox", "Inbox", 88_000, 10, 0),
            new ComFolderInfo("corpus", "Sent Items", "Sent Items", 40_000, 0, 0),
            new ComFolderInfo("corpus", "Search Root", "Search Root", null, null, 0),
            new ComFolderInfo("hub@vm.invalid", "Inbox", "Inbox", 24, 4, 2),
        });
        Assert.Equal(new LiveStoreSize("corpus", 128_000, true), LiveLatencyTarget.SizeOf("corpus", tree));
        Assert.Equal(new LiveStoreSize("hub@vm.invalid", 24, true), LiveLatencyTarget.SizeOf("HUB@vm.invalid", tree) with { Store = "hub@vm.invalid" });
        Assert.Null(LiveLatencyTarget.SizeOf("absent", tree).Items);
        Assert.False(LiveLatencyTarget.SizeOf("corpus", new ComFolderTree(tree.Folders, walkCapReached: true)).Complete);
        Assert.False(LiveLatencyTarget.SizeOf("corpus", new ComFolderTree(tree.Folders, depthLimitReached: true)).Complete);
    }

    // ================================================================ the wiring, from the IL

    [Fact]
    public void TheMeasurement_ChoosesWithTheChooser_FromEveryStoresFolderWalk()
    {
        MethodInfo measure = typeof(LiveLatencyTarget).GetMethod(nameof(LiveLatencyTarget.Measure))!;
        Assert.True(Calls(measure, typeof(LiveLatencyTarget), nameof(LiveLatencyTarget.Largest)));
        Assert.True(Calls(measure, typeof(LiveLatencyTarget), nameof(LiveLatencyTarget.SizeOf)));
        Assert.True(Calls(measure, typeof(LiveTestSettings), nameof(LiveTestSettings.RequireIndexedStores)));

        MethodInfo scope = typeof(LivePhase1Fixture).GetProperty(nameof(LivePhase1Fixture.LargestIndexedScope))!.GetGetMethod()!;
        Assert.True(Calls(scope, typeof(LiveLatencyTarget), nameof(LiveLatencyTarget.Measure)));
    }

    [Theory]
    [InlineData(nameof(LiveIndexSearchTests.FilterShapes_ReadAndAttachmentFlags_WorkUnder2s))]
    [InlineData(nameof(LiveIndexSearchTests.SenderFilter_PerColumnContains_IndexBackedUnder2s))]
    public void TheFilterShapeTests_TimeOnlyThroughTheLargestStoreGuard(string test)
    {
        // They time through the one helper that refuses any scope but the largest store's, they take
        // that scope from the chooser, and they have no Assert.InRange of their own - so putting a
        // latency bound back on the first indexed store cannot be done without failing here.
        MethodInfo method = typeof(LiveIndexSearchTests).GetMethod(test)!;
        Assert.True(Calls(method, typeof(LiveIndexSearchTests), "AssertTimedOnTheLargestStore"));
        Assert.True(Calls(method, typeof(LivePhase1Fixture), "get_" + nameof(LivePhase1Fixture.LargestIndexedScope)));
        Assert.False(Calls(method, typeof(Assert), nameof(Assert.InRange)), $"{test} asserts a latency bound of its own");

        MethodInfo guard = typeof(LiveIndexSearchTests).GetMethod("AssertTimedOnTheLargestStore", BindingFlags.NonPublic | BindingFlags.Instance)!;
        Assert.True(Calls(guard, typeof(LiveLatencyTarget), nameof(LiveLatencyTarget.Measure)));
        Assert.True(Calls(guard, typeof(Assert), nameof(Assert.InRange)));
    }

    [Fact]
    public void TheOrPairTest_TimesItsStoreScopedShapeOnTheLargestStore_NotOnTheProbesStore()
    {
        MethodInfo method = typeof(LiveSearchInTests).GetMethod(nameof(LiveSearchInTests.IndexTier_OrPairLatency_StaysAcceptableVersusSingleColumn))!;
        Assert.True(Calls(method, typeof(LiveLatencyTarget), nameof(LiveLatencyTarget.Measure)));
        Assert.False(
            Calls(method, typeof(SubjectOnlyProbeSettings), "get_" + nameof(SubjectOnlyProbeSettings.StoreDisplayName)),
            "the OR-pair test reads the probe's STORE again - its store-scoped shape is timed on the largest store");
    }

    [Fact]
    public void TheCachedReRead_IsTimedOnAHitFromTheLargestStore_NotOnTheFirstHit()
    {
        MethodInfo method = typeof(LiveMailServiceTests).GetMethod(nameof(LiveMailServiceTests.RoundTrip_SearchThenRead_TenHitsAcrossStores))!;
        Assert.True(Calls(method, typeof(LiveLatencyTarget), nameof(LiveLatencyTarget.Measure)));
        Assert.False(
            Calls(method, typeof(List<HitSummary>), "get_Item"),
            "the read test indexes its hit list again - hits[0] is the FIRST indexed store's, which is what it stopped timing");
    }

    [Fact]
    public void TheFrontierTest_ReadsTheHubPopulationAndFailsOnAStaleOne()
    {
        MethodInfo method = typeof(LiveIndexSearchTests).GetMethod(nameof(LiveIndexSearchTests.Staleness_SelfReportsPlausibleFrontier))!;
        Assert.True(Calls(method, typeof(LiveTestSettings), "get_" + nameof(LiveTestSettings.HubPopulationManifestPath)));
        Assert.True(Calls(method, typeof(LiveHubPopulationFreshness), nameof(LiveHubPopulationFreshness.Read)));
        Assert.True(Calls(method, typeof(LiveHubPopulationFreshness), nameof(LiveHubPopulationFreshness.Decide)));
    }

    // ================================================================ the hub's age

    private static readonly DateTime Anchor = new(2026, 9, 24, 9, 0, 0, DateTimeKind.Utc);

    private static string[] HubManifest(string store = "tier@vm.invalid", CorpusPopulationKind kind = CorpusPopulationKind.Hub, string? shapeKey = null)
    {
        var options = new CorpusPlanOptions("hub-indexed", 8181, Anchor)
        {
            Population = kind,
            Owner = CorpusMailboxOwner.ForStore(store),
        };
        var header = new CorpusManifestHeader(
            CorpusManifest.CurrentVersion, "hub-indexed", 8181, CorpusManifest.FormatUtc(Anchor),
            shapeKey ?? options.ShapeKey, store, null, "PropertyAccessorDates", "InPlaceWithSentFlag");
        return new[] { CorpusManifest.RenderLine(header) };
    }

    [Fact]
    public void AHubManifest_ReadsBackItsAnchor_AndItsNewestItem_OneMinuteBeforeIt()
    {
        HubPopulationFact fact = LiveHubPopulationFreshness.Read(HubManifest());
        Assert.Equal("hub-indexed", fact.CorpusId);
        Assert.Equal("tier@vm.invalid", fact.Store);
        Assert.Equal(Anchor, fact.AnchorUtc);
        Assert.Equal(Anchor.AddMinutes(-1), fact.NewestDatedUtc);
    }

    [Fact]
    public void AManifestThatIsNotAHubPopulationThisGeneratorWouldReproduce_IsRefused()
    {
        Assert.Throws<InvalidOperationException>(() => LiveHubPopulationFreshness.Read(HubManifest(kind: CorpusPopulationKind.Bystander)));
        Assert.Throws<InvalidOperationException>(() => LiveHubPopulationFreshness.Read(HubManifest(shapeKey: "v1|hub-indexed|8181|old")));

        // The measurement corpus's own manifest, pointed at by mistake.
        var corpus = new CorpusPlanOptions("vm-indexed", 7777, Anchor);
        string[] corpusManifest =
        {
            CorpusManifest.RenderLine(new CorpusManifestHeader(
                CorpusManifest.CurrentVersion, "vm-indexed", 7777, CorpusManifest.FormatUtc(Anchor), corpus.ShapeKey,
                "Corpus A", null, "PropertyAccessorDates")),
        };
        Assert.Throws<InvalidOperationException>(() => LiveHubPopulationFreshness.Read(corpusManifest));
    }

    [Fact]
    public void AFreshHub_Proceeds_AStaleOne_FailsWithTheRemedy()
    {
        HubPopulationFact fact = LiveHubPopulationFreshness.Read(HubManifest());
        TimeSpan winter = TimeSpan.FromHours(1);

        (bool fresh, string ok) = LiveHubPopulationFreshness.Decide(fact, Anchor.AddMinutes(20), winter);
        Assert.True(fresh, ok);

        // 55 minutes is the margin on a UTC+1 guest: the offset less the test's own five.
        Assert.True(LiveHubPopulationFreshness.Decide(fact, Anchor.AddMinutes(-1).AddMinutes(55), winter).Proceed);
        (bool stale, string why) = LiveHubPopulationFreshness.Decide(fact, Anchor.AddMinutes(-1).AddMinutes(56), winter);
        Assert.False(stale);
        Assert.Contains("STALE HUB", why, StringComparison.Ordinal);
        Assert.Contains("Reset-HubPopulation.ps1", why, StringComparison.Ordinal);

        // Summer time doubles the room.
        Assert.True(LiveHubPopulationFreshness.Decide(fact, Anchor.AddMinutes(100), TimeSpan.FromHours(2)).Proceed);
    }

    [Fact]
    public void AMachineOnUtc_OrAHubFromTheFuture_IsRefusedNotPassed()
    {
        HubPopulationFact fact = LiveHubPopulationFreshness.Read(HubManifest());
        (bool onUtc, string utcWhy) = LiveHubPopulationFreshness.Decide(fact, Anchor, TimeSpan.Zero);
        Assert.False(onUtc);
        Assert.Contains("cannot tell a local-time frontier", utcWhy, StringComparison.Ordinal);

        (bool future, string futureWhy) = LiveHubPopulationFreshness.Decide(fact, Anchor.AddHours(-2), TimeSpan.FromHours(1));
        Assert.False(future);
        Assert.Contains("FUTURE", futureWhy, StringComparison.Ordinal);

        Assert.Equal(TimeSpan.FromMinutes(55), LiveHubPopulationFreshness.DiscriminatingAge(TimeSpan.FromHours(-1)));
    }

    // ================================================================ the settings key

    [Fact]
    public void TheHubPopulationManifestKey_IsReadWhenPresent_AbsentWhenNot_AndNeverBlank()
    {
        const string Base = """
            {
              "machineProfile": "Portable",
              "testHubStoreDisplayName": "hub@vm.invalid",
              "expectedStoreDisplayNames": [ "hub@vm.invalid", "bystander@vm.invalid" ],
              "bystanderStoreDisplayNames": [ "bystander@vm.invalid" ]
            """;
        Assert.Null(LiveTestSettings.Parse(Base + "}").HubPopulationManifestPath);
        Assert.Equal(
            @"C:\OutlookAI-Q5\corpus-hub-indexed.jsonl",
            LiveTestSettings.Parse(Base + """, "hubPopulationManifestPath": "C:\\OutlookAI-Q5\\corpus-hub-indexed.jsonl" }""").HubPopulationManifestPath);
        InvalidOperationException blank = Assert.Throws<InvalidOperationException>(
            () => LiveTestSettings.Parse(Base + """, "hubPopulationManifestPath": "  " }"""));
        Assert.Contains("hubPopulationManifestPath", blank.Message, StringComparison.Ordinal);
    }

    // ================================================================ IL

    /// <summary>
    /// Whether <paramref name="method"/> - or a compiler-generated closure of its declaring type that
    /// belongs to it - calls a member of <paramref name="declaring"/> named <paramref name="name"/>.
    /// Tokens are RESOLVED, as in StoreListSplitTests, so a stray operand byte cannot pass for a call.
    /// </summary>
    private static bool Calls(MethodBase method, Type declaring, string name)
    {
        foreach (MethodBase body in WithClosures(method))
        {
            byte[]? il = body.GetMethodBody()?.GetILAsByteArray();
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
                    MethodBase? resolved = body.Module.ResolveMethod(
                        token, body.DeclaringType?.GetGenericArguments(), body.IsGenericMethod ? body.GetGenericArguments() : null);
                    Type? owner = resolved?.DeclaringType;
                    bool typeMatches = owner == declaring
                        || (owner != null && declaring.IsGenericType && owner.IsGenericType
                            && owner.GetGenericTypeDefinition() == declaring.GetGenericTypeDefinition()
                            && owner.GetGenericArguments().SequenceEqual(declaring.GetGenericArguments()));
                    if (typeMatches && resolved!.Name == name)
                    {
                        return true;
                    }
                }
                catch (ArgumentException)
                {
                    // Not a real call - the bytes happened to look like one.
                }
            }
        }

        return false;
    }

    private static IEnumerable<MethodBase> WithClosures(MethodBase method)
    {
        yield return method;
        Type? type = method.DeclaringType;
        if (type == null)
        {
            yield break;
        }

        const BindingFlags all = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance
            | BindingFlags.Static | BindingFlags.DeclaredOnly;
        string marker = "<" + method.Name + ">";
        foreach (Type nested in type.GetNestedTypes(BindingFlags.NonPublic))
        {
            foreach (MethodBase closure in nested.GetMethods(all).Where(m => m.Name.StartsWith(marker, StringComparison.Ordinal)))
            {
                yield return closure;
            }
        }

        foreach (MethodBase local in type.GetMethods(all).Where(m => m.Name.StartsWith(marker, StringComparison.Ordinal)))
        {
            yield return local;
        }
    }
}
