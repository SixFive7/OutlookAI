using System.Reflection;
using OutlookAI.ComHost.Host;
using OutlookAI.McpServer.Tests.T2;
using OutlookAI.McpServer.Tests.T3;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the live tier's own inventory, by reflection over this assembly, so the tier cannot
/// be narrowed by accident.
/// <para>
/// <b>The problem these solve.</b> The live tier is excluded from CI, so nothing about it is
/// checked by running it. Two ways it had already gone wrong silently: three live classes
/// belonged to no collection at all and therefore passed through no census, no health
/// preflight and no store-count tripwire; and every live test was selectable only as one
/// undifferentiated <c>Category=Live</c> blob, so moving the runnable ones to a dedicated
/// test machine could only be done by naming them in a filter string that nothing would ever
/// check again.
/// </para>
/// <para>
/// <b>The mechanism.</b> Two traits. <c>LiveTier</c> is the SELECTOR - exactly one value per
/// live test, either <see cref="Portable"/> (honest on a machine with PST stores, no mail
/// accounts, no delegate mailboxes and nothing in the search index) or
/// <see cref="ProfileBound"/> (needs the maintainer's real profile). <c>Requires</c> is the
/// REASON, and it is not decoration: a ProfileBound test must name at least one capability a
/// test machine cannot have, and a Portable test must name none of them. That way "this test
/// cannot move to the VM" is a claim with evidence attached rather than an assertion, and a
/// test quietly reclassified to shrink the VM subset fails here.
/// </para>
/// <para>
/// <b>The stdio tier joined later, with the inverse problem.</b> The live tier's tests
/// announce themselves; the T3 stdio tier's did not. Sixteen files ran under
/// <c>Category!=Live</c> - several named <c>...CiToolShapeTests</c>, i.e. explicitly labelled
/// CI-safe - and eleven of their tests issued a <c>tools/call</c> that attaches to whatever
/// Outlook is on the machine. Because the tool decides that, not the signature, the
/// declaration is a literal handed to <c>McpStdioClient</c> and the pin reads it back out of
/// the compiled IL. See <see cref="EveryStdioTestReachingOutlook_DeclaresIt"/>.
/// </para>
/// <para>
/// These run in CI (they are not <c>Category=Live</c> themselves) and need no Outlook: they
/// read attributes and method bodies, not mailboxes.
/// </para>
/// </summary>
public sealed class LiveTierInventoryTests
{
    /// <summary>Runs on any configured machine, the dedicated test VM included.</summary>
    private const string Portable = "Portable";

    /// <summary>Needs the maintainer's own profile, and says which part of it under <c>Requires</c>.</summary>
    private const string ProfileBound = "ProfileBound";

    /// <summary>
    /// Capabilities a dedicated test machine cannot be given by configuration: they need real
    /// accounts, real delegate access, a populated search index, real transport, more than one
    /// account store, a hub small enough for a paging assertion, or a hand-curated population
    /// of real mail named in the gitignored settings.
    /// </summary>
    private static readonly string[] ProductionOnlyCapabilities =
    {
        "SearchIndex",
        "MailAccount",
        "Transport",
        "MultipleStores",
        "DelegateStore",
        "SmallHubStore",
        "ProbePopulation",
    };

    /// <summary>
    /// Capabilities a test machine CAN have, recorded because they still constrain how a run
    /// is launched: an interactive desktop session (Outlook windows and screenshots cannot be
    /// driven from session 0), the add-in's registry tuning state, and an Outlook instance to
    /// attach to at all.
    /// </summary>
    private static readonly string[] PortableCapabilities =
    {
        "InteractiveDesktop",
        "AddInRegistry",
        OutlookInstance,
    };

    /// <summary>
    /// An Outlook to attach to, and nothing more specific than that.
    /// <para>
    /// Added for the T3 stdio tier, whose problem was the opposite of the live tier's. The
    /// live tier's tests announce themselves; T3's did not - sixteen files sat under
    /// <c>Category!=Live</c> with names like <c>...CiToolShapeTests</c>, and eleven of their
    /// tests called <c>outlook_health</c>, <c>list_accounts</c> or <c>search</c>, which attach
    /// to whatever Outlook is on the machine. On the maintainer's box that is a production
    /// mailbox, read on every verification run for months, neither intended nor declared.
    /// </para>
    /// <para>
    /// It is a PORTABLE capability on purpose: any Outlook profile satisfies it, so a test
    /// declaring it stays runnable on the dedicated test VM. A test needing something OF that
    /// profile - real accounts, a populated index, a delegate store - names that as well, and
    /// is ProfileBound because of it.
    /// </para>
    /// </summary>
    private const string OutlookInstance = "OutlookInstance";

    /// <summary>The namespace whose classes speak real MCP over stdio to a real server process.</summary>
    private const string StdioTierNamespace = "OutlookAI.McpServer.Tests.T3";

    /// <summary>
    /// The one tool no COM-host fault can neutralise: its report is half COM store probe and half
    /// Windows Search index query, and the index half never goes near the host.
    /// </summary>
    private const string OutlookHealthTool = "outlook_health";

    private readonly ITestOutputHelper _output;

    public LiveTierInventoryTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void EveryLiveTest_DeclaresExactlyOneTier_FromTheKnownVocabulary()
    {
        List<string> problems = new();
        int portable = 0;
        int profileBound = 0;

        foreach (MethodInfo method in LiveTestMethods())
        {
            List<string> tiers = TraitValues(method, "LiveTier");
            if (tiers.Count != 1)
            {
                problems.Add(Name(method) + ": " + tiers.Count + " LiveTier trait(s), expected exactly 1");
                continue;
            }

            switch (tiers[0])
            {
                case Portable:
                    portable++;
                    break;
                case ProfileBound:
                    profileBound++;
                    break;
                default:
                    problems.Add(Name(method) + ": unknown LiveTier '" + tiers[0] + "'");
                    break;
            }
        }

        _output.WriteLine("live tests: " + (portable + profileBound) + " (Portable " + portable
            + ", ProfileBound " + profileBound + ")");
        Assert.Empty(problems);

        // A tier that has become all-or-nothing is a tier nobody can run anywhere but here,
        // or a classification that has stopped meaning anything. Both are worth a red test.
        Assert.True(portable > 0, "no live test is marked Portable - the VM subset would be empty");
        Assert.True(profileBound > 0, "no live test is marked ProfileBound - that would be surprising");
    }

    [Fact]
    public void ProfileBoundTests_NameACapabilityATestMachineCannotHave()
    {
        List<string> problems = new();
        foreach (MethodInfo method in LiveTestMethods())
        {
            List<string> tiers = TraitValues(method, "LiveTier");
            if (tiers.Count != 1 || tiers[0] != ProfileBound)
            {
                continue;
            }

            List<string> requires = TraitValues(method, "Requires");
            if (!requires.Any(ProductionOnlyCapabilities.Contains))
            {
                problems.Add(Name(method)
                    + ": ProfileBound with no production-only Requires trait, so nothing says WHY it "
                    + "cannot run on a test machine");
            }
        }

        Assert.Empty(problems);
    }

    [Fact]
    public void PortableTests_ClaimNoCapabilityATestMachineLacks()
    {
        List<string> problems = new();
        foreach (MethodInfo method in LiveTestMethods())
        {
            List<string> tiers = TraitValues(method, "LiveTier");
            if (tiers.Count != 1 || tiers[0] != Portable)
            {
                continue;
            }

            foreach (string requirement in TraitValues(method, "Requires").Where(ProductionOnlyCapabilities.Contains))
            {
                problems.Add(Name(method) + ": Portable but requires '" + requirement + "' - contradiction");
            }
        }

        Assert.Empty(problems);
    }

    [Fact]
    public void EveryRequiresValue_IsInTheVocabulary()
    {
        string[] known = ProductionOnlyCapabilities.Concat(PortableCapabilities).ToArray();
        List<string> problems = new();
        foreach (MethodInfo method in LiveTestMethods())
        {
            foreach (string requirement in TraitValues(method, "Requires").Where(r => !known.Contains(r)))
            {
                problems.Add(Name(method) + ": unknown Requires value '" + requirement + "'");
            }
        }

        // A free-text vocabulary drifts into synonyms, and two spellings of one capability
        // make the filter that excludes it silently incomplete.
        Assert.Empty(problems);
    }

    [Fact]
    public void EveryLiveTestClass_SitsInAGuardedCollection()
    {
        // This is the hole three T3 classes fell through: no [Collection] means xunit invents
        // one with no fixture, so the class runs against real mailboxes with no census, no
        // health preflight and no tripwire verification.
        List<string> problems = new();
        foreach (Type type in LiveTestClasses())
        {
            string? collection = SingleAttributeArgument(type, "CollectionAttribute");
            if (collection == null)
            {
                problems.Add(type.Name + ": no [Collection] - runs outside every live-tier guard");
            }
            else if (!LiveCollections.IsGuarded(collection))
            {
                problems.Add(type.Name + ": collection '" + collection + "' is not in LiveCollections.All");
            }
        }

        Assert.Empty(problems);
    }

    [Fact]
    public void LiveCollections_ListsEveryCollectionThisAssemblyDeclares()
    {
        // The run plan decides where a run gets VERIFIED by asking which collections in the
        // ordered list are guarded. A collection missing from LiveCollections.All is invisible
        // to that question, so a run ending in it would never be verified.
        List<string> declared = typeof(LiveCollections).Assembly.GetTypes()
            .Select(t => SingleAttributeArgument(t, "CollectionDefinitionAttribute"))
            .Where(name => name != null)
            .Select(name => name!)
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToList();

        Assert.Equal(
            LiveCollections.All.OrderBy(name => name, StringComparer.Ordinal).ToList(),
            declared);
    }

    [Fact]
    public void EveryTestClaimingAnOutlookInstance_IsAlsoInTheLiveTier()
    {
        // The reason without the selector is the worst of both: the test says it needs an
        // Outlook and a default run schedules it anyway.
        List<string> problems = new();
        foreach (Type type in TestClasses())
        {
            if (!TraitValues(type, "Requires").Contains(OutlookInstance))
            {
                continue;
            }

            if (!TraitValues(type, "Category").Contains("Live"))
            {
                problems.Add(type.Name + ": Requires=" + OutlookInstance + " without Category=Live, so a default "
                    + "run would still schedule it against whatever Outlook is on the machine");
            }
        }

        Assert.Empty(problems);
    }

    [Fact]
    public void EveryStdioTestReachingOutlook_DeclaresIt()
    {
        // The T3 tier talks to a real server process, so what it touches is decided by which
        // TOOL it calls, and no attribute can be derived from that by reflection over
        // signatures. It can be derived from the IL: a test that may call an Outlook-reaching
        // tool has to hand McpStdioClient a specific literal, so the literal's presence in a
        // class IS the declaration, and this reads it back out of the compiled method bodies
        // (async test methods included, whose real body is the state machine's MoveNext).
        //
        // Chosen over a hard-coded roster of class names because a roster only catches a new
        // CLASS. This catches a new METHOD in an existing class, which is how the tier drifted
        // in the first place.
        List<string> problems = new();
        int declaringClasses = 0;
        int healthCallers = 0;
        int scanned = 0;
        foreach (Type type in TestClasses().Where(t => t.Namespace == StdioTierNamespace))
        {
            scanned++;
            HashSet<string> literals = StringLiteralsOf(type);
            bool live = TraitValues(type, "Category").Contains("Live");
            bool declares = literals.Contains(McpStdioClient.OutlookReachingToolsAllowed);
            bool faultsTheHost = literals.Contains(ComHostFaultInjection.Variable);
            declaringClasses += declares ? 1 : 0;
            healthCallers += literals.Contains(OutlookHealthTool) ? 1 : 0;

            StdioDeclarationVerdict verdict = ClassifyStdioClass(
                declares,
                live,
                faultsTheHost,
                literals.Contains(OutlookHealthTool),
                TraitValues(type, "Requires").Contains(OutlookInstance));
            if (verdict != StdioDeclarationVerdict.Ok)
            {
                problems.Add(type.Name + ": " + DescribeStdioVerdict(verdict));
            }
        }

        _output.WriteLine("stdio classes scanned: " + scanned + ", declaring mailbox contact: " + declaringClasses
            + ", naming " + OutlookHealthTool + ": " + healthCallers);
        Assert.Empty(problems);

        // The detector must not be able to switch itself off. A scan that finds nothing reports
        // no problems, which reads exactly like a clean tier - so the two ways of finding nothing
        // are asserted against directly: no classes at all (the namespace moved) and no
        // declaration at all (the ldstr walk stopped resolving, or the live half was deleted).
        Assert.True(scanned > 0, "no classes found in " + StdioTierNamespace + " - this pin is scanning nothing");
        Assert.True(
            declaringClasses > 0,
            "no stdio class declares mailbox contact. Either the tier lost its live half, or the IL "
            + "walk has stopped resolving string literals and this pin is now green by accident.");

        // And the SPELLING of the one tool the fault exemption stops at. The exemption's logic is
        // pinned by TheStdioDeclarationMatrix_HoldsForEveryCombination, which passes booleans; a
        // typo in the constant would leave that green while making the check key on a tool that
        // does not exist. No class can be committed to demonstrate the failure - it would fail the
        // pin - so the name is proved by the fact that the tier names it at all.
        Assert.True(
            healthCallers > 0,
            "no stdio class names '" + OutlookHealthTool + "', which cannot be right for a tier that "
            + "asserts its description and calls it in the live half. The constant is probably misspelled, "
            + "and the fault exemption is keying on a tool that does not exist.");
    }

    [Theory]
    // Nothing to declare: a class that never names an Outlook-reaching tool is fine either way.
    [InlineData(false, false, false, false, false, StdioDeclarationVerdict.Ok)]
    [InlineData(false, true, false, false, true, StdioDeclarationVerdict.Ok)]
    // The failure this whole pin exists for.
    [InlineData(true, false, false, false, false, StdioDeclarationVerdict.UndeclaredMailboxContact)]
    // The fault exemption, and its edge. A faulted list_accounts never reaches a session; a
    // faulted outlook_health still queries the Windows Search index.
    [InlineData(true, false, true, false, false, StdioDeclarationVerdict.Ok)]
    [InlineData(true, false, true, true, false, StdioDeclarationVerdict.HealthSurvivesTheFault)]
    // Declared and live, but with no capability naming what it needs.
    [InlineData(true, true, false, false, false, StdioDeclarationVerdict.LiveWithoutItsReason)]
    [InlineData(true, true, false, true, true, StdioDeclarationVerdict.Ok)]
    public void TheStdioDeclarationMatrix_HoldsForEveryCombination(
        bool declares, bool live, bool faultsTheComHost, bool namesOutlookHealth, bool requiresOutlookInstance,
        StdioDeclarationVerdict expected)
    {
        // The matrix is pinned here rather than only by the classes that happen to exist. Two of
        // its four outcomes are unreachable from real classes by construction - a class that trips
        // them would fail the pin above, so it cannot be committed to demonstrate them - and a
        // decision line no test can reach is a decision line anybody may delete.
        Assert.Equal(
            expected,
            ClassifyStdioClass(declares, live, faultsTheComHost, namesOutlookHealth, requiresOutlookInstance));
    }

    /// <summary>What is wrong with one stdio class's declaration, if anything.</summary>
    public enum StdioDeclarationVerdict
    {
        /// <summary>Declared consistently, or nothing to declare.</summary>
        Ok = 0,

        /// <summary>Calls an Outlook-reaching tool from outside the live tier, and nothing stops it.</summary>
        UndeclaredMailboxContact = 1,

        /// <summary>Neutralises COM with an injected fault, but calls the one tool a fault cannot neutralise.</summary>
        HealthSurvivesTheFault = 2,

        /// <summary>In the live tier and reaching Outlook, but naming no capability as the reason.</summary>
        LiveWithoutItsReason = 3,
    }

    /// <summary>
    /// The whole decision, pure and in one place so it can be pinned by
    /// <see cref="TheStdioDeclarationMatrix_HoldsForEveryCombination"/>.
    /// </summary>
    /// <param name="declaresContact">The class hands <c>McpStdioClient</c> the contact declaration.</param>
    /// <param name="live">The class carries <c>Category=Live</c>.</param>
    /// <param name="faultsTheComHost">
    /// The class injects a COM-host fault. That fault is applied in <c>ComHostServer</c> ABOVE the
    /// routing proxy, so the faulted operation never reaches an Outlook session - which is the one
    /// legitimate way to name an Outlook-reaching tool outside the live tier.
    /// </param>
    /// <param name="namesOutlookHealth">
    /// The class names <c>outlook_health</c>, whose Windows Search index probe does not go through
    /// the COM host at all and therefore survives any fault.
    /// </param>
    /// <param name="requiresOutlookInstance">The class carries <c>Requires=OutlookInstance</c>.</param>
    internal static StdioDeclarationVerdict ClassifyStdioClass(
        bool declaresContact,
        bool live,
        bool faultsTheComHost,
        bool namesOutlookHealth,
        bool requiresOutlookInstance)
    {
        if (!declaresContact)
        {
            return StdioDeclarationVerdict.Ok;
        }

        if (live)
        {
            return requiresOutlookInstance
                ? StdioDeclarationVerdict.Ok
                : StdioDeclarationVerdict.LiveWithoutItsReason;
        }

        if (!faultsTheComHost)
        {
            return StdioDeclarationVerdict.UndeclaredMailboxContact;
        }

        return namesOutlookHealth
            ? StdioDeclarationVerdict.HealthSurvivesTheFault
            : StdioDeclarationVerdict.Ok;
    }

    private static string DescribeStdioVerdict(StdioDeclarationVerdict verdict)
    {
        switch (verdict)
        {
            case StdioDeclarationVerdict.UndeclaredMailboxContact:
                return "calls an Outlook-reaching tool but is not Category=Live - a default run would "
                    + "attach to the machine's own mailbox";
            case StdioDeclarationVerdict.HealthSurvivesTheFault:
                return "neutralises COM with an injected fault but still calls " + OutlookHealthTool
                    + ", whose index probe the fault cannot reach";
            case StdioDeclarationVerdict.LiveWithoutItsReason:
                return "calls an Outlook-reaching tool without Requires=" + OutlookInstance
                    + ", so nothing says WHAT it needs of the machine";
            default:
                return "ok";
        }
    }

    /// <summary>
    /// Every string a class's compiled code can push, its compiler-generated nested types
    /// included: async state machines, iterator state machines and lambda display classes,
    /// which is where the body of every test method in this assembly actually lives.
    /// </summary>
    private static HashSet<string> StringLiteralsOf(Type type)
    {
        HashSet<string> literals = new(StringComparer.Ordinal);
        CollectStringLiterals(type, literals);
        return literals;
    }

    private static void CollectStringLiterals(Type type, HashSet<string> literals)
    {
        const BindingFlags all = BindingFlags.Public | BindingFlags.NonPublic
            | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly;

        foreach (MethodBase method in type.GetMethods(all).Cast<MethodBase>().Concat(type.GetConstructors(all)))
        {
            byte[]? il = SafeIl(method);
            if (il == null)
            {
                continue;
            }

            // ldstr is 0x72 followed by a 4-byte metadata token whose high byte is 0x70
            // (the user-string heap). Requiring that byte is what keeps the scan from
            // resolving random operand bytes, and ResolveString covers the rest by throwing.
            for (int i = 0; i + 4 < il.Length; i++)
            {
                if (il[i] != 0x72 || il[i + 4] != 0x70)
                {
                    continue;
                }

                int token = il[i + 1] | (il[i + 2] << 8) | (il[i + 3] << 16) | (il[i + 4] << 24);
                try
                {
                    literals.Add(type.Module.ResolveString(token));
                }
                catch (ArgumentException)
                {
                    // Not a real ldstr - the bytes happened to look like one.
                }
            }
        }

        foreach (Type nested in type.GetNestedTypes(BindingFlags.Public | BindingFlags.NonPublic))
        {
            CollectStringLiterals(nested, literals);
        }
    }

    private static byte[]? SafeIl(MethodBase method)
    {
        try
        {
            return method.GetMethodBody()?.GetILAsByteArray();
        }
        catch (Exception ex) when (ex is InvalidOperationException || ex is NotSupportedException)
        {
            // Abstract, extern or otherwise bodiless.
            return null;
        }
    }

    /// <summary>Every non-abstract class in this assembly that declares at least one test.</summary>
    private static IEnumerable<Type> TestClasses()
    {
        return typeof(LiveCollections).Assembly.GetTypes()
            .Where(t => t.IsClass && !t.IsAbstract)
            .Where(t => t.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly)
                .Any(m => m.GetCustomAttributes(typeof(FactAttribute), inherit: true).Length > 0))
            .OrderBy(t => t.FullName, StringComparer.Ordinal);
    }

    /// <summary>Every test method carrying <c>Category=Live</c>, whether from its class or itself.</summary>
    private static IEnumerable<MethodInfo> LiveTestMethods()
    {
        return LiveTestClasses()
            .SelectMany(t => t.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly))
            .Where(m => m.GetCustomAttributes(typeof(FactAttribute), inherit: true).Length > 0);
    }

    private static IEnumerable<Type> LiveTestClasses()
    {
        return typeof(LiveCollections).Assembly.GetTypes()
            .Where(t => t.IsClass && !t.IsAbstract)
            .Where(t => TraitValues(t, "Category").Contains("Live"))
            .OrderBy(t => t.FullName, StringComparer.Ordinal);
    }

    /// <summary>
    /// All values of one trait name on a method, its declaring type included.
    /// <para>
    /// Read through <see cref="CustomAttributeData"/> rather than a typed attribute instance
    /// because xunit's <c>TraitAttribute</c> exposes no Name or Value property - the
    /// constructor arguments are the only place the pair exists.
    /// </para>
    /// </summary>
    private static List<string> TraitValues(MethodInfo method, string traitName)
    {
        List<string> values = TraitValues(method.DeclaringType!, traitName);
        values.AddRange(TraitValues(method.GetCustomAttributesData(), traitName));
        return values;
    }

    private static List<string> TraitValues(Type type, string traitName)
    {
        return TraitValues(type.GetCustomAttributesData(), traitName);
    }

    private static List<string> TraitValues(IEnumerable<CustomAttributeData> attributes, string traitName)
    {
        return attributes
            .Where(a => a.AttributeType == typeof(TraitAttribute) && a.ConstructorArguments.Count == 2)
            .Where(a => string.Equals(a.ConstructorArguments[0].Value as string, traitName, StringComparison.Ordinal))
            .Select(a => a.ConstructorArguments[1].Value as string)
            .Where(value => value != null)
            .Select(value => value!)
            .ToList();
    }

    /// <summary>The one string argument of an attribute named <paramref name="attributeName"/>, or null.</summary>
    private static string? SingleAttributeArgument(Type type, string attributeName)
    {
        return type.GetCustomAttributesData()
            .Where(a => string.Equals(a.AttributeType.Name, attributeName, StringComparison.Ordinal))
            .Where(a => a.ConstructorArguments.Count == 1)
            .Select(a => a.ConstructorArguments[0].Value as string)
            .FirstOrDefault(value => value != null);
    }

    private static string Name(MethodInfo method)
    {
        return method.DeclaringType!.Name + "." + method.Name;
    }
}
