using System.Reflection;
using OutlookAI.McpServer.Tests.T2;
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
/// These run in CI (they are not <c>Category=Live</c> themselves) and need no Outlook: they
/// read attributes, not mailboxes.
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
    /// driven from session 0) and the add-in's registry tuning state.
    /// </summary>
    private static readonly string[] PortableCapabilities =
    {
        "InteractiveDesktop",
        "AddInRegistry",
    };

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
