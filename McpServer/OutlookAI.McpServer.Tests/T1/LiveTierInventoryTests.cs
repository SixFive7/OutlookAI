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
/// <b>The mechanism: TWO axes, and the second one carries everything.</b> <c>Category=Live</c>
/// says "this test needs a mailbox" - it is the CI gate, and it survives not because of the
/// VM but because CI runs on a GitHub Windows runner with no Outlook at all. <c>Requires</c>
/// says WHAT of a machine the test needs, from a closed vocabulary
/// (<see cref="AllCapabilities"/>), and it is declared <b>per method</b>. Everything else -
/// which bucket a test is in, whether it can leave the maintainer's box - is COMPUTED from
/// that list and never written down twice.
/// </para>
/// <para>
/// <b>Three buckets, computed, not declared.</b> <b>CI</b> = no <c>Category=Live</c>, needs no
/// mailbox. <b>VM</b> = live, and every capability it names is one the dedicated test VM can be
/// given. <b>Production-only</b> = live, and it names something the VM cannot be given - which
/// today is <c>DelegateStore</c> and nothing else: delegate mailboxes are indexed WITHOUT folder
/// nesting, a local PST cannot fake that, and faking it would manufacture confidence in the
/// one area this product has most often been surprised by.
/// </para>
/// <para>
/// <b>There used to be a third axis, <c>LiveTier</c>, and deleting it is the point.</b> It held
/// <c>Portable</c> or <c>ProfileBound</c> and had to agree with <c>Requires</c> by hand - a
/// computed value maintained manually, which is the exact drift this file exists to prevent.
/// Worse, it was paired with CLASS-level <c>Requires</c>, so a class read as the union of
/// everything any one of its methods needed: that is what turned a real floor of six impossible
/// tests into a reported 96. Both are now structurally impossible -
/// <see cref="NoTestDeclaresARetiredTrait"/> refuses the trait and
/// <see cref="EveryLiveTestMethod_NamesItsOwnCapabilities"/> refuses class-level attribution.
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
    /// <summary>
    /// The whole capability vocabulary, in one list, because there is only one axis now.
    /// <para>
    /// A free-text vocabulary drifts into synonyms, and two spellings of one capability make
    /// the filter that excludes it silently incomplete. Every value here is documented in
    /// <c>Docs/live-tier-on-the-vm.md</c>, and <c>.github/scripts/check-pinned-constants.ps1</c>
    /// fails the build if one of them stops appearing there - the runbook is what a human reads
    /// to decide whether a live test can move to a test machine.
    /// </para>
    /// <para>
    /// Six of these were "production-only" until the test VM's shape was settled: an indexed
    /// corpus store, a dummy mail account, a local SMTP sink that delivers back, three stores,
    /// a configurable small hub and a known population between them cover <c>SearchIndex</c>,
    /// <c>MailAccount</c>, <c>Transport</c>, <c>MultipleStores</c>, <c>SmallHubStore</c> and
    /// <c>ProbePopulation</c>. They are ordinary capabilities now.
    /// </para>
    /// </summary>
    private static readonly string[] AllCapabilities =
    {
        // An Outlook to attach to, and nothing more specific than that. The FLOOR: a live test
        // that needs nothing else still needs this, and says so rather than saying nothing.
        OutlookInstance,

        // An interactive desktop session - Outlook windows and screenshots cannot be driven
        // from session 0. Declared only by tests that PUT SOMETHING ON SCREEN; a test that
        // merely asserts no window appeared does not need one.
        "InteractiveDesktop",

        // The add-in's registry tuning state (the D24 groups), read or flipped.
        "AddInRegistry",

        // A populated Windows Search index.
        "SearchIndex",

        // A mail account, as opposed to a bare PST store.
        "MailAccount",

        // Mail that actually goes out and comes back.
        "Transport",

        // More than one store mounted at once.
        "MultipleStores",

        // A hub store small enough for a paging assertion to mean something.
        "SmallHubStore",

        // A hand-curated population named in the gitignored live-test settings.
        "ProbePopulation",

        // The one capability a dedicated test machine cannot be given.
        DelegateStore,
    };

    /// <summary>
    /// The capabilities no dedicated test machine can be given by configuration - the whole
    /// definition of the production-only bucket.
    /// <para>
    /// One entry, and it is not an oversight. A delegate/shared mailbox is indexed with its
    /// folder hierarchy FLATTENED, which is not a property a local PST can be made to have;
    /// simulating it would produce a green test about a shape the real thing does not have.
    /// Every other capability the VM can be built to provide, so a test naming one of those is
    /// a VM test even if it has only ever run on the maintainer's machine.
    /// </para>
    /// </summary>
    private static readonly string[] ProductionOnlyCapabilities =
    {
        DelegateStore,
    };

    /// <summary>An Outlook to attach to, and nothing more specific than that.</summary>
    /// <remarks>
    /// It is a VM capability on purpose: any Outlook profile satisfies it, so a test declaring
    /// it stays runnable on the dedicated test VM. A test needing something OF that profile -
    /// a delegate store - names that as well, and is production-only because of it.
    /// </remarks>
    private const string OutlookInstance = "OutlookInstance";

    /// <summary>A delegate/shared mailbox, whose index namespace drops every intermediate folder.</summary>
    private const string DelegateStore = "DelegateStore";

    /// <summary>The trait names this suite used to carry and must never carry again.</summary>
    private static readonly string[] RetiredTraits = { "LiveTier" };

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
    public void EveryLiveTestMethod_NamesItsOwnCapabilities()
    {
        // The say-why rule, and the whole reason the axis is worth having. Two halves:
        //
        //  (a) On the METHOD. A class-level Requires reads as the union of everything any one
        //      of its methods needs, so the impossible set inflates with every test added to a
        //      big class - measured: 6 methods genuinely need a delegate mailbox, class-level
        //      attribution reported 23, and the old LiveTier axis built on it reported 96
        //      tests that could not leave this machine.
        //
        //  (b) At least one. "Category=Live with no Requires" is a test that says it needs a
        //      mailbox and refuses to say what for; the floor is OutlookInstance, which costs
        //      one line and makes the claim checkable.
        List<string> problems = new();
        int methods = 0;

        foreach (Type type in LiveTestClasses())
        {
            foreach (string classLevel in TraitValues(type, "Requires"))
            {
                problems.Add(type.Name + ": class-level Requires='" + classLevel
                    + "'. Requires is declared per METHOD - on a class it reads as the union of "
                    + "everything any one test in it needs, which is how the impossible set inflates.");
            }

            foreach (MethodInfo method in TestMethodsOf(type))
            {
                methods++;
                if (MethodTraitValues(method, "Requires").Count == 0)
                {
                    problems.Add(Name(method) + ": Category=Live with no Requires on the method, so "
                        + "nothing says WHAT of a machine it needs. The floor is Requires="
                        + OutlookInstance + ".");
                }
            }
        }

        _output.WriteLine("live test methods: " + methods);
        Assert.Empty(problems);

        // A scan that found nothing reports no problems, which reads exactly like a clean tier.
        Assert.True(methods > 0, "no Category=Live test methods found - this pin is scanning nothing");
    }

    [Fact]
    public void NoTestDeclaresARetiredTrait()
    {
        // LiveTier held Portable/ProfileBound and had to agree with Requires by hand. It was a
        // COMPUTED value maintained manually, so it could disagree, and a test quietly moved
        // from ProfileBound to Portable shrank the impossible set with nothing to notice. The
        // bucket is derived from Requires now (see TheBuckets_AreDerivedFromRequiresAlone), and
        // this refuses the axis rather than trusting a comment to keep it out.
        List<string> problems = new();
        foreach (Type type in TestClasses())
        {
            foreach (string retired in RetiredTraits)
            {
                if (TraitValues(type, retired).Count > 0)
                {
                    problems.Add(type.Name + ": carries the retired trait '" + retired + "'");
                }

                foreach (MethodInfo method in TestMethodsOf(type))
                {
                    if (MethodTraitValues(method, retired).Count > 0)
                    {
                        problems.Add(Name(method) + ": carries the retired trait '" + retired + "'");
                    }
                }
            }
        }

        Assert.Empty(problems);
    }

    [Fact]
    public void EveryRequiresValue_IsInTheVocabulary()
    {
        List<string> problems = new();
        foreach (MethodInfo method in LiveTestMethods())
        {
            foreach (string requirement in TraitValues(method, "Requires").Where(r => !AllCapabilities.Contains(r)))
            {
                problems.Add(Name(method) + ": unknown Requires value '" + requirement + "'");
            }
        }

        Assert.Empty(problems);
    }

    [Fact]
    public void TheProductionOnlyList_IsAProperSubsetOfTheVocabulary()
    {
        // Both ends matter. An entry outside the vocabulary can never match a real trait, so the
        // production-only bucket would silently empty; a list that swallowed the whole vocabulary
        // would put every live test back on one machine. The runbook and
        // check-pinned-constants.ps1 read the vocabulary array, so a value missing from it is
        // also a value no human is told about.
        Assert.NotEmpty(ProductionOnlyCapabilities);
        Assert.All(ProductionOnlyCapabilities, c => Assert.Contains(c, AllCapabilities));
        Assert.True(
            ProductionOnlyCapabilities.Length < AllCapabilities.Length,
            "every capability is production-only, which would mean the VM can run nothing");
    }

    [Fact]
    public void TheBuckets_AreDerivedFromRequiresAlone()
    {
        // The bucket is a QUESTION asked of Requires, never a value stored beside it. This is
        // the whole replacement for the deleted LiveTier axis: the same classification, computed
        // where it cannot disagree with its own evidence.
        List<string> productionOnly = new();
        int vm = 0;

        foreach (MethodInfo method in LiveTestMethods())
        {
            List<string> requires = MethodTraitValues(method, "Requires");
            string[] blocking = requires.Where(ProductionOnlyCapabilities.Contains).ToArray();
            if (blocking.Length > 0)
            {
                productionOnly.Add(Name(method) + " (" + string.Join(", ", blocking) + ")");
            }
            else
            {
                vm++;
            }
        }

        _output.WriteLine("VM: " + vm + ", production-only: " + productionOnly.Count);
        foreach (string entry in productionOnly.OrderBy(e => e, StringComparer.Ordinal))
        {
            _output.WriteLine("  production-only: " + entry);
        }

        // A tier that has become all-or-nothing is a tier nobody can run anywhere but here, or a
        // classification that has stopped meaning anything. Both are worth a red test.
        Assert.True(vm > 0, "no live test can run on the VM - the VM subset would be empty");
        Assert.True(
            productionOnly.Count > 0,
            "no live test names a production-only capability. Either the delegate-store coverage "
            + "was deleted, or a capability was quietly reclassified as reproducible on the VM.");
    }

    [Fact]
    public void EveryTestClaimingACapability_IsAlsoInTheLiveTier()
    {
        // The reason without the gate is the worst of both: the test says it needs something of
        // the machine's Outlook and a default `Category!=Live` run schedules it anyway.
        List<string> problems = new();
        foreach (Type type in TestClasses())
        {
            bool live = TraitValues(type, "Category").Contains("Live");
            foreach (MethodInfo method in TestMethodsOf(type))
            {
                if (live || TraitValues(method, "Category").Contains("Live"))
                {
                    continue;
                }

                List<string> requires = TraitValues(method, "Requires");
                if (requires.Count > 0)
                {
                    problems.Add(Name(method) + ": Requires=" + string.Join("/", requires)
                        + " without Category=Live, so a default run would still schedule it against "
                        + "whatever Outlook is on the machine");
                }
            }
        }

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
    public void EveryStdioTestReachingOutlook_DeclaresIt()
    {
        // The T3 tier talks to a real server process, so what it touches is decided by which
        // TOOL it calls, and no attribute can be derived from that by reflection over
        // signatures. It can be derived from the IL: both the tool name and the client's
        // contact token are string literals, so their presence in a class IS the declaration,
        // and this reads them back out of the compiled method bodies (async test methods
        // included, whose real body is the state machine's MoveNext).
        //
        // Chosen over a hard-coded roster of class names because a roster only catches a new
        // CLASS. This catches a new METHOD in an existing class, which is how the tier drifted
        // in the first place.
        List<string> problems = new();
        int declaringClasses = 0;
        int healthCallers = 0;
        int guardedToolNamers = 0;
        int scanned = 0;
        foreach (Type type in TestClasses().Where(t => t.Namespace == StdioTierNamespace))
        {
            scanned++;
            HashSet<string> literals = StringLiteralsOf(type);
            bool live = TraitValues(type, "Category").Contains("Live");
            bool declares = literals.Contains(McpStdioClient.OutlookReachingToolsAllowed);
            bool faultsTheHost = literals.Contains(ComHostFaultInjection.Variable);
            bool namesGuardedTool = McpStdioClient.ToolsThatAlwaysReachOutlook.Any(literals.Contains);
            declaringClasses += declares ? 1 : 0;
            healthCallers += literals.Contains(OutlookHealthTool) ? 1 : 0;
            guardedToolNamers += namesGuardedTool ? 1 : 0;

            StdioDeclarationVerdict verdict = ClassifyStdioClass(
                declares,
                live,
                faultsTheHost,
                literals.Contains(OutlookHealthTool),
                namesGuardedTool);
            if (verdict != StdioDeclarationVerdict.Ok)
            {
                problems.Add(type.Name + ": " + DescribeStdioVerdict(verdict));
            }
        }

        _output.WriteLine("stdio classes scanned: " + scanned + ", declaring mailbox contact: " + declaringClasses
            + ", naming a guarded tool: " + guardedToolNamers
            + ", naming " + OutlookHealthTool + ": " + healthCallers);
        Assert.Empty(problems);

        // The detector must not be able to switch itself off. A scan that finds nothing reports
        // no problems, which reads exactly like a clean tier - so the ways of finding nothing
        // are asserted against directly: no classes at all (the namespace moved), no declaration
        // at all (the ldstr walk stopped resolving, or the live half was deleted), and no guarded
        // tool named anywhere (the roster in McpStdioClient emptied or was renamed).
        Assert.True(scanned > 0, "no classes found in " + StdioTierNamespace + " - this pin is scanning nothing");
        Assert.True(
            declaringClasses > 0,
            "no stdio class declares mailbox contact. Either the tier lost its live half, or the IL "
            + "walk has stopped resolving string literals and this pin is now green by accident.");
        Assert.True(
            guardedToolNamers > 0,
            "no stdio class names any tool from McpStdioClient.ToolsThatAlwaysReachOutlook, which "
            + "cannot be right for a tier whose live half exists to call them.");

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
    [InlineData(false, true, false, false, false, StdioDeclarationVerdict.Ok)]
    // The failure this whole pin exists for: contact declared from outside the live tier.
    [InlineData(true, false, false, false, true, StdioDeclarationVerdict.UndeclaredMailboxContact)]
    // The fault exemption, and its edge. A faulted list_accounts never reaches a session; a
    // faulted outlook_health still queries the Windows Search index.
    [InlineData(true, false, true, false, true, StdioDeclarationVerdict.Ok)]
    [InlineData(true, false, true, true, true, StdioDeclarationVerdict.HealthSurvivesTheFault)]
    // Live and reaching Outlook, but never handing the client the token - the call throws at the
    // first tools/call, so the test cannot have been run since the guard was added.
    [InlineData(false, true, false, false, true, StdioDeclarationVerdict.ReachesOutlookWithoutTheClientToken)]
    [InlineData(true, true, false, true, true, StdioDeclarationVerdict.Ok)]
    public void TheStdioDeclarationMatrix_HoldsForEveryCombination(
        bool declares, bool live, bool faultsTheComHost, bool namesOutlookHealth, bool namesAGuardedTool,
        StdioDeclarationVerdict expected)
    {
        // The matrix is pinned here rather than only by the classes that happen to exist. Two of
        // its four outcomes are unreachable from real classes by construction - a class that trips
        // them would fail the pin above, so it cannot be committed to demonstrate them - and a
        // decision line no test can reach is a decision line anybody may delete.
        Assert.Equal(
            expected,
            ClassifyStdioClass(declares, live, faultsTheComHost, namesOutlookHealth, namesAGuardedTool));
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

        /// <summary>In the live tier and naming a tool the client refuses without the contact token.</summary>
        ReachesOutlookWithoutTheClientToken = 3,
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
    /// <param name="namesAGuardedTool">
    /// The class names one of <c>McpStdioClient.ToolsThatAlwaysReachOutlook</c>, which the client
    /// refuses to send unless the test hands it the contact token.
    /// </param>
    internal static StdioDeclarationVerdict ClassifyStdioClass(
        bool declaresContact,
        bool live,
        bool faultsTheComHost,
        bool namesOutlookHealth,
        bool namesAGuardedTool)
    {
        if (live)
        {
            // Inside the live tier the classification is settled; what is NOT settled is whether
            // the test can run at all. McpStdioClient throws on a tools/call for a guarded tool
            // unless the token was passed, so a live class naming one without it is broken, and
            // broken in the way nothing notices: the live tier is excluded from every CI run.
            // Three classes were in exactly this state when the token guard landed.
            return namesAGuardedTool && !declaresContact
                ? StdioDeclarationVerdict.ReachesOutlookWithoutTheClientToken
                : StdioDeclarationVerdict.Ok;
        }

        if (!declaresContact)
        {
            // Outside the live tier, naming a guarded tool is not by itself a fault: every
            // tools/list roster assertion mentions all 21 tool names. The client's runtime
            // refusal is what stops a call, and it needs no declaration to fire.
            return StdioDeclarationVerdict.Ok;
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
                return "declares mailbox contact but is not Category=Live - a default run would "
                    + "attach to the machine's own mailbox";
            case StdioDeclarationVerdict.HealthSurvivesTheFault:
                return "neutralises COM with an injected fault but still calls " + OutlookHealthTool
                    + ", whose index probe the fault cannot reach";
            case StdioDeclarationVerdict.ReachesOutlookWithoutTheClientToken:
                return "is Category=Live and names a tool from McpStdioClient.ToolsThatAlwaysReachOutlook, "
                    + "but never passes outlookReachingTools: McpStdioClient.OutlookReachingToolsAllowed - "
                    + "the client refuses the call, so this test throws instead of testing";
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
            .Where(t => TestMethodsOf(t).Any())
            .OrderBy(t => t.FullName, StringComparer.Ordinal);
    }

    /// <summary>Every test method carrying <c>Category=Live</c>, whether from its class or itself.</summary>
    private static IEnumerable<MethodInfo> LiveTestMethods()
    {
        return LiveTestClasses().SelectMany(TestMethodsOf);
    }

    private static IEnumerable<Type> LiveTestClasses()
    {
        return typeof(LiveCollections).Assembly.GetTypes()
            .Where(t => t.IsClass && !t.IsAbstract)
            .Where(t => TraitValues(t, "Category").Contains("Live"))
            .OrderBy(t => t.FullName, StringComparer.Ordinal);
    }

    private static IEnumerable<MethodInfo> TestMethodsOf(Type type)
    {
        return type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.DeclaredOnly)
            .Where(m => m.GetCustomAttributes(typeof(FactAttribute), inherit: true).Length > 0);
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
        values.AddRange(MethodTraitValues(method, traitName));
        return values;
    }

    /// <summary>Values declared on the METHOD itself, with nothing inherited from its class.</summary>
    private static List<string> MethodTraitValues(MethodInfo method, string traitName)
    {
        return TraitValues(method.GetCustomAttributesData(), traitName);
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
