using System.Reflection;
using Microsoft.Win32;
using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tests.T2;
using OutlookAI.McpServer.Tests.T3;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Q74 layer 1, the static half: a live test carrying <c>Writes=None</c> cannot reach a write - in
/// its own body, in its class's or its fixtures' constructors and teardown, or through a tool it asks
/// the MCP server for - as far as that can be proven statically (<see cref="WritePathAnalyzer"/> says
/// where that stops). The workstation filter selects only such tests, so this is what stands behind
/// the filter at build time; layers 2 and 3 stand behind it at run time.
/// <para>
/// Three parts: every carrier is walked and must come back clean; every product member the walk is
/// allowed to stop at (<see cref="ReadOnlyProductApi"/>) is walked INSIDE the product and must reach no
/// write either; and a set of controls - classes no runner ever constructs - proves each kind of write
/// the walk claims to see is actually seen, so the check cannot pass by having stopped seeing.
/// </para>
/// </summary>
public sealed class ReadOnlyLiveTestTests
{
    private readonly ITestOutputHelper _output;

    public ReadOnlyLiveTestTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void EveryWritesNoneTest_ReachesNoWrite_InItsBodyOrItsFixtures()
    {
        List<MethodInfo> carriers = WritesNoneTests().ToList();
        List<string> problems = new();
        int walked = 0;
        foreach (MethodInfo test in carriers)
        {
            WritePathWalk walk = WritePathAnalyzer.Analyze(test);
            walked += walk.MethodsWalked;
            problems.AddRange(walk.Findings.Select(f => test.DeclaringType!.Name + "." + test.Name + ": " + f));
        }

        _output.WriteLine("Writes=None tests: " + carriers.Count + ", method bodies walked: " + walked);
        foreach (string problem in problems.Take(40))
        {
            _output.WriteLine("PROBLEM " + problem);
        }

        Assert.True(
            problems.Count == 0,
            problems.Count + " way(s) a Writes=None test can reach a write. Either the test writes - take the trait "
            + "off, it cannot run on the read-only workstation - or it calls a product member nobody has listed as a "
            + "read: list it in ReadOnlyProductApi with its reason, and the in-product walk will check the claim.\n"
            + string.Join("\n", problems.Take(20)));

        // The seven Exchange-only tests and the PST half of the short-id check carry it at the least; a
        // scan that found fewer is scanning the wrong thing, and the walk must have read real code.
        Assert.True(carriers.Count >= 8, "only " + carriers.Count + " Writes=None tests found");
        Assert.True(walked > carriers.Count * 50, "the walk read only " + walked + " method bodies - it has stopped following calls");
    }

    [Fact]
    public void EveryListedProductMember_ReachesNoWriteInsideTheProduct()
    {
        List<string> problems = new();
        int walked = 0;
        foreach ((string key, List<MethodBase> members) in ListedProductMembers())
        {
            if (members.Count == 0)
            {
                problems.Add(key + ": listed in ReadOnlyProductApi but no such product member exists - a stale entry");
                continue;
            }

            WritePathWalk walk = WritePathAnalyzer.WalkProduct(members.Select(m => (m, key)));
            walked += walk.MethodsWalked;
            problems.AddRange(walk.Findings.Select(f => key + ": " + f));
        }

        _output.WriteLine("product methods walked from the list: " + walked);
        Assert.True(
            problems.Count == 0,
            "A member listed as read-only reaches a write inside the product. Take it off ReadOnlyProductApi - and "
            + "the trait off any test that needs it - or, if the write is not one (a window the session itself owns, "
            + "say), record that as an InProductExemptions entry with its reason.\n" + string.Join("\n", problems.Take(20)));
        Assert.True(walked > 200, "the in-product walk read only " + walked + " method bodies - it has stopped following calls");
    }

    [Fact]
    public void EveryInProductExemption_StillPointsAtRealCode()
    {
        HashSet<string> sources = ProductTypes()
            .SelectMany(t => t.GetMethods(Everything).Cast<MethodBase>().Concat(t.GetConstructors(Everything)))
            .Select(ReadOnlyProductApi.SourceKeyOf)
            .ToHashSet(StringComparer.Ordinal);
        foreach (string exemption in ReadOnlyProductApi.InProductExemptions.Keys)
        {
            string source = exemption.Split('|')[0];
            Assert.True(sources.Contains(source), "the exemption '" + exemption + "' names a product member that no longer exists");
        }
    }

    // ------------------------------------------------------------------ controls: each kind of write is seen

    [Theory]
    [InlineData(nameof(Controls.AsksTheWriteGuard), "asks the write guard")]
    [InlineData(nameof(Controls.DraftsThroughTheProduct), "MailService.NewDraft")]
    [InlineData(nameof(Controls.SendsThroughTheServer), "write-capable MCP tool 'send'")]
    [InlineData(nameof(Controls.WritesInsideALambda), "asks the write guard")]
    [InlineData(nameof(Controls.DeletesThroughDynamicCom), "late-bound (COM) member 'Delete'")]
    [InlineData(nameof(Controls.SetsAComProperty), "SETS a property 'Subject'")]
    [InlineData(nameof(Controls.WritesTheRegistry), "writes the registry")]
    [InlineData(nameof(Controls.WritesThroughAnInterface), "asks the write guard")]
    public void Control_EachKindOfWrite_IsSeen(string control, string expected)
    {
        WritePathWalk walk = WritePathAnalyzer.Walk(new[] { ((MethodBase)typeof(Controls).GetMethod(control, Everything)!, control) });

        Assert.Contains(walk.Findings, f => f.Reason.Contains(expected, StringComparison.Ordinal));
    }

    [Fact]
    public void Control_AFixtureThatWrites_IsSeen_ThroughATestThatDoesNothing()
    {
        // The B1 shape: a test that only reads, in the MoveArchive collection, whose fixture deletes test
        // folders in the hub. The walk has to find the write in the FIXTURE, or a read-only test in a
        // writing collection would pass this check and throw at its fixture on the workstation.
        WritePathWalk walk = WritePathAnalyzer.Analyze(typeof(InAWritingCollection).GetMethod(nameof(InAWritingCollection.OnlyReads))!);

        Assert.Contains(walk.Findings, f => f.Chain.Any(link => link.Contains(nameof(LiveMoveArchiveFixture), StringComparison.Ordinal))
            && f.Reason.Contains("asks the write guard", StringComparison.Ordinal));
    }

    [Fact]
    public void Control_ATestThatOnlyReads_InACollectionThatOnlyReads_IsClean()
    {
        WritePathWalk walk = WritePathAnalyzer.Analyze(typeof(InAReadingCollection).GetMethod(nameof(InAReadingCollection.OnlyReads))!);

        Assert.Empty(walk.Findings);
        Assert.True(walk.MethodsWalked > 50, "the walk did not follow the fixture chain");
    }

    [Fact]
    public void Control_TheInProductWalk_SeesTheProductsOwnWrites()
    {
        // Every draft the product makes is audited, so the walk from new_draft's service method must
        // find the append - or the check above is green because it stopped looking.
        MethodBase newDraft = typeof(MailService).GetMethods(Everything).First(m => m.Name == nameof(MailService.NewDraft));
        WritePathWalk walk = WritePathAnalyzer.WalkProduct(new[] { (newDraft, "MailService.NewDraft") });

        Assert.Contains(walk.Findings, f => f.Reason.Contains("audit line", StringComparison.Ordinal));
    }

    // ------------------------------------------------------------------ plumbing

    private const BindingFlags Everything = BindingFlags.Public | BindingFlags.NonPublic
        | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly;

    private static IEnumerable<MethodInfo> WritesNoneTests()
    {
        return typeof(LiveCollections).Assembly.GetTypes()
            .Where(t => t.IsClass && !t.IsAbstract)
            .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly))
            .Where(m => m.GetCustomAttributes(typeof(FactAttribute), inherit: true).Length > 0)
            .Where(m => m.GetCustomAttributesData().Any(a => a.AttributeType == typeof(TraitAttribute)
                && a.ConstructorArguments.Count == 2
                && string.Equals(a.ConstructorArguments[0].Value as string, LiveRunFilters.WritesTrait, StringComparison.Ordinal)
                && string.Equals(a.ConstructorArguments[1].Value as string, LiveRunFilters.WritesNone, StringComparison.Ordinal)))
            .OrderBy(m => m.DeclaringType!.FullName + "." + m.Name, StringComparer.Ordinal);
    }

    /// <summary>Every member each list entry stands for: a type's every member, or a name's every overload.</summary>
    private static IEnumerable<(string Key, List<MethodBase> Members)> ListedProductMembers()
    {
        Type[] types = ProductTypes();
        foreach (string typeKey in ReadOnlyProductApi.Types.Keys)
        {
            Type? type = types.FirstOrDefault(t => ReadOnlyProductApi.TypeKeyOf(t) == typeKey);
            List<MethodBase> members = type == null
                ? new List<MethodBase>()
                : type.GetMethods(Everything).Cast<MethodBase>().Concat(type.GetConstructors(Everything)).ToList();
            yield return (typeKey, members);
        }

        foreach (string memberKey in ReadOnlyProductApi.Members.Keys)
        {
            List<MethodBase> members = types
                .SelectMany(t => t.GetMethods(Everything).Cast<MethodBase>())
                .Where(m => ReadOnlyProductApi.KeyOf(m) == memberKey)
                .ToList();
            yield return (memberKey, members);
        }
    }

    private static Type[] ProductTypes()
    {
        return new[]
            {
                typeof(MailService).Assembly,
                typeof(OutlookAI.McpServer.Tools.OutlookTools).Assembly,
                typeof(OutlookAI.RemediationTools.CorpusPlan).Assembly,
            }
            .SelectMany(a => a.GetTypes())
            .ToArray();
    }

    // ------------------------------------------------------------------ controls: never constructed, never run

    // Read by the walk, never executed: no [Fact], no runner builds them. Each one does exactly one kind
    // of write the walk claims to see.

    private static class Controls
    {
        internal static void AsksTheWriteGuard()
        {
            LiveStoreWriteGuard.Writable("hub@example.test", StoreWriteKind.Draft, "control");
        }

        internal static void DraftsThroughTheProduct(MailService service)
        {
            service.NewDraft("hub@example.test", "hub@example.test", cc: null, "control", "control", display: false);
        }

        internal static async Task SendsThroughTheServer()
        {
            await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync();
            await client.CallToolAsync("send", new { id = "control" });
        }

        internal static void WritesInsideALambda()
        {
            Action later = () => LiveStoreWriteGuard.Assert("hub@example.test", StoreWriteKind.Delete, "control");
            later();
        }

        internal static void DeletesThroughDynamicCom(dynamic item)
        {
            item.Delete();
        }

        internal static void SetsAComProperty(dynamic item)
        {
            item.Subject = "control";
        }

        internal static void WritesTheRegistry()
        {
            Registry.SetValue(@"HKEY_CURRENT_USER\Software\OutlookAI-Control", "Control", 1);
        }

        internal static void WritesThroughAnInterface(IControlWriter writer)
        {
            writer.Write();
        }
    }

    internal interface IControlWriter
    {
        void Write();
    }

    private sealed class ControlWriter : IControlWriter
    {
        public void Write()
        {
            LiveStoreWriteGuard.Writable("hub@example.test", StoreWriteKind.Move, "control");
        }
    }

    [Collection(LiveCollections.MoveArchive)]
    private sealed class InAWritingCollection
    {
        public void OnlyReads()
        {
            _ = LiveRunFilters.Guest.Length;
        }
    }

    [Collection(LiveCollections.Phase2)]
    private sealed class InAReadingCollection
    {
        private readonly LivePhase2Fixture _fixture;

        public InAReadingCollection(LivePhase2Fixture fixture)
        {
            _fixture = fixture;
        }

        public void OnlyReads()
        {
            _ = _fixture.Service.Search(new SearchRequest { Query = "control", IndexOnly = true, Top = 1 });
        }
    }
}
