using System.Reflection;
using System.Reflection.Emit;
using OutlookAI.McpServer.Tests.T2;
using Xunit;
using Xunit.Abstractions;
using Xunit.Sdk;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The live tier's per-run opt-in (<see cref="LiveRunOptIn"/>): what it decides, what it says,
/// and - the half that matters - that NO <c>Category=Live</c> test can start without passing it.
/// <para>
/// <b>WHY THE PROOF IS STRUCTURAL.</b> The live tier never runs in CI, so nothing about it is
/// checked by running it; and a gate that one live class can walk around is a gate in name only.
/// So this reads the compiled code. xunit builds a test collection's fixture before any test in
/// it, and when that constructor throws, every test in the collection fails with that exception
/// without its class being constructed or its method run - observed on 2026-09-24, when three
/// live tests stopped in 1 ms at <c>LiveMoveArchiveFixture</c>'s first line. So the funnel is:
/// every live class sits in a collection whose fixture's constructor reaches the gate before it
/// calls ANYTHING of ours or any framework method (<see cref="EveryLiveTestClass_StartsBehindTheGate"/>),
/// and the gate is the first thing <see cref="LiveTestSettings.Load"/> does
/// (<see cref="LoadAsksTheGateBeforeAnythingElse"/>). The stdio tier is inside that: its live
/// classes sit in the same guarded collections, so the server process they spawn is never started
/// by a run that did not opt in.
/// </para>
/// <para>
/// <b>THE CONTROL.</b> <see cref="Control_TheFunnelCheck_CatchesEveryWayAroundTheGate"/> runs the same
/// check over fixtures and classes built to get around the gate - one that never calls it, one that
/// calls our code first, one that reaches COM through the framework first, one with a static
/// initializer, one class with no collection, one naming a collection nobody defines, one with an
/// ungated class fixture - and each must be caught for the reason it was built to fail.
/// </para>
/// </summary>
public sealed class LiveRunOptInTests
{
    private const string ThisGuest = "OAI-INDEXED";

    private readonly ITestOutputHelper _output;

    public LiveRunOptInTests(ITestOutputHelper output)
    {
        _output = output;
    }

    // ------------------------------------------------------------------ the decision

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void NoOptIn_Refuses(string? value)
    {
        Assert.Equal(LiveRunOptIn.Verdict.Missing, LiveRunOptIn.Evaluate(value, null, null, ThisGuest));
    }

    [Theory]
    [InlineData("OAI-INDEXED")]
    [InlineData("oai-indexed")]
    [InlineData("  OAI-INDEXED ")]
    public void AnOptInNamingThisMachine_Opens(string value)
    {
        // NetBIOS names are case-insensitive, and a trailing space from a copied line is not a
        // different machine.
        Assert.Equal(LiveRunOptIn.Verdict.Open, LiveRunOptIn.Evaluate(value, null, null, ThisGuest));
    }

    [Theory]
    [InlineData("OAI-UNINDEXED")]
    [InlineData("1")]
    [InlineData("true")]
    [InlineData("yes")]
    public void AnOptInNamingAnythingElse_Refuses(string value)
    {
        // The whole point of a machine-name value: a boolean left in a profile, or a value carried
        // to another machine, opens nothing.
        Assert.Equal(LiveRunOptIn.Verdict.OtherMachine, LiveRunOptIn.Evaluate(value, null, null, ThisGuest));
    }

    [Theory]
    [InlineData("OAI-INDEXED", "OAI-INDEXED", null)]
    [InlineData("OAI-INDEXED", null, "OAI-INDEXED")]
    [InlineData(null, "OAI-INDEXED", null)]
    [InlineData(null, null, "anything")]
    public void AnOptInSavedInTheEnvironment_Refuses_EvenOnItsOwnMachine(string? process, string? user, string? machine)
    {
        // A saved value opens every later run on this machine - the accidental ones included -
        // which is exactly what a per-run opt-in exists to prevent.
        Assert.Equal(LiveRunOptIn.Verdict.Persisted, LiveRunOptIn.Evaluate(process, user, machine, ThisGuest));
    }

    // ------------------------------------------------------------------ what a refused run is told

    [Theory]
    [InlineData(LiveRunOptIn.Verdict.Missing)]
    [InlineData(LiveRunOptIn.Verdict.OtherMachine)]
    [InlineData(LiveRunOptIn.Verdict.Persisted)]
    public void EveryRefusal_SaysWhatTheOptInIsFor_HowAGuestGivesIt_AndThatTheWorkstationIsReadOnly(LiveRunOptIn.Verdict verdict)
    {
        string message = LiveRunOptIn.DescribeRefusal(verdict, ThisGuest, "OAI-UNINDEXED");

        Assert.StartsWith("LIVE TEST REFUSED", message, StringComparison.Ordinal);

        // What it is for.
        Assert.Contains("real Outlook profile", message, StringComparison.Ordinal);
        Assert.Contains("by accident", message, StringComparison.Ordinal);

        // How to set it for a deliberate run on a test guest.
        Assert.Contains("TEST GUEST", message, StringComparison.Ordinal);
        Assert.Contains("$env:" + LiveRunOptIn.Variable + " = ", message, StringComparison.Ordinal);
        Assert.Contains("$env:COMPUTERNAME", message, StringComparison.Ordinal);
        Assert.Contains("Category=Live", message, StringComparison.Ordinal);
        Assert.Contains("Register-InteractiveTask.ps1", message, StringComparison.Ordinal);
        Assert.Contains("Testbed/README.md, section 4c", message, StringComparison.Ordinal);
        Assert.Contains("setx", message, StringComparison.Ordinal);

        // That the maintainer's workstation is read-only for live tests.
        Assert.Contains("WORKSTATION IS READ-ONLY FOR LIVE TESTS", message, StringComparison.Ordinal);
        Assert.Contains("makes no test read-only", message, StringComparison.Ordinal);
    }

    [Fact]
    public void AWrongMachineRefusal_NamesBothMachines()
    {
        string message = LiveRunOptIn.DescribeRefusal(LiveRunOptIn.Verdict.OtherMachine, ThisGuest, " OAI-UNINDEXED ");

        Assert.Contains("'OAI-UNINDEXED'", message, StringComparison.Ordinal);
        Assert.Contains("'" + ThisGuest + "'", message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheSectionTheRefusalPointsAt_ExistsAndGivesTheOptIn()
    {
        // The message sends a refused run to Testbed/README.md section 4c. A pointer at a
        // heading that was renamed or never written is a refusal that explains nothing.
        string readme = File.ReadAllText(Path.Combine(RepoRoot(), "Testbed", "README.md"));

        Assert.Contains("## 4c.", readme, StringComparison.Ordinal);
        Assert.Contains("$env:" + LiveRunOptIn.Variable + " = ", readme, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ the gate, run for real

    [Fact]
    public void Require_Refuses_WhenThisRunDidNotOptIn()
    {
        WithoutTheOptIn(() =>
        {
            InvalidOperationException refused = Assert.Throws<InvalidOperationException>(LiveRunOptIn.Require);
            Assert.StartsWith("LIVE TEST REFUSED", refused.Message, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void Load_RefusesBeforeItReadsAnything()
    {
        // The door itself: without the opt-in, Load throws the gate's refusal - not "settings
        // not found", not a JSON error - so no settings are read and nothing after it runs.
        WithoutTheOptIn(() =>
        {
            InvalidOperationException refused = Assert.Throws<InvalidOperationException>(() => LiveTestSettings.Load());
            Assert.StartsWith("LIVE TEST REFUSED", refused.Message, StringComparison.Ordinal);
        });
    }

    // ------------------------------------------------------------------ the funnel

    [Fact]
    public void LoadAsksTheGateBeforeAnythingElse()
    {
        MethodInfo load = typeof(LiveTestSettings).GetMethod(nameof(LiveTestSettings.Load), BindingFlags.Public | BindingFlags.Static)
            ?? throw new InvalidOperationException("LiveTestSettings.Load is gone - this pin proves nothing.");

        (OpCode _, MethodBase? first) = CallsIn(load).FirstOrDefault();

        Assert.True(
            first != null && first.DeclaringType == typeof(LiveRunOptIn) && first.Name == nameof(LiveRunOptIn.Require),
            "LiveTestSettings.Load must call LiveRunOptIn.Require before anything else; its first call is "
            + Describe(first));
    }

    [Fact]
    public void EveryLiveTestClass_StartsBehindTheGate()
    {
        List<Type> liveClasses = LiveTestClasses().ToList();
        List<string> problems = liveClasses.SelectMany(FunnelProblems).ToList();
        HashSet<Type> fixtures = new(liveClasses.SelectMany(FixturesBehind));

        _output.WriteLine("live classes: " + liveClasses.Count + ", fixtures they start behind: " + fixtures.Count);
        foreach (Type fixture in fixtures.OrderBy(t => t.Name, StringComparer.Ordinal))
        {
            _output.WriteLine("  " + fixture.Name);
        }

        foreach (string problem in problems)
        {
            _output.WriteLine("PROBLEM " + problem);
        }

        Assert.Empty(problems);

        // A scan that found nothing reports no problems, which reads exactly like a gated tier.
        Assert.True(liveClasses.Count > 0, "no Category=Live class found - this pin is scanning nothing");
        Assert.True(fixtures.Count > 0, "no live fixture found - this pin is scanning nothing");

        // The stdio tier is in the scan: its live classes spawn the real server, so they must be
        // among the classes proven to start behind the gate, not assumed to.
        Assert.Contains(liveClasses, t => t.Namespace == typeof(T3.McpStdioClient).Namespace);
    }

    [Fact]
    public void Control_TheFunnelCheck_CatchesEveryWayAroundTheGate()
    {
        // The shape every real fixture has passes...
        Assert.Empty(ConstructorProblems(typeof(GatedControlFixture)));
        Assert.Empty(ConstructorProblems(typeof(RequireFirstControlFixture)));

        // ...and every way around it is caught, for the reason it was built to fail.
        Assert.Contains("never reaches the opt-in gate", Assert.Single(ConstructorProblems(typeof(UngatedControlFixture))),
            StringComparison.Ordinal);
        Assert.Contains(nameof(LiveCollections.IsGuarded), Assert.Single(ConstructorProblems(typeof(OursBeforeGateControlFixture))),
            StringComparison.Ordinal);
        Assert.Contains(nameof(Type.GetTypeFromProgID), Assert.Single(ConstructorProblems(typeof(ComBeforeGateControlFixture))),
            StringComparison.Ordinal);
        Assert.Contains("static initializer", Assert.Single(ConstructorProblems(typeof(StaticInitializerControlFixture))),
            StringComparison.Ordinal);

        Assert.Contains("no [Collection]", Assert.Single(FunnelProblems(typeof(NoCollectionControlClass))), StringComparison.Ordinal);
        Assert.Contains("no collection definition", Assert.Single(FunnelProblems(typeof(UnknownCollectionControlClass))),
            StringComparison.Ordinal);

        // Sits in a real guarded collection (declaring a collection of its own would fail the
        // inventory pin), so only the class-fixture problem is asserted, not a count - a broken
        // real fixture is the main test's to report.
        Assert.Contains(
            FunnelProblems(typeof(UngatedClassFixtureControlClass)),
            p => p.Contains("(class fixture): " + nameof(UngatedControlFixture), StringComparison.Ordinal));

        // And the whole-assembly assertion itself fails on them, so it is not decoration.
        Assert.ThrowsAny<XunitException>(() => Assert.Empty(FunnelProblems(typeof(UngatedClassFixtureControlClass))));
    }

    // ------------------------------------------------------------------ the check

    /// <summary>
    /// Everything that stops one live class from starting behind the gate: no collection, a
    /// collection nobody defines, a collection that builds no fixture, or a fixture - collection
    /// or class - whose construction can do anything before the gate.
    /// </summary>
    private static List<string> FunnelProblems(Type liveClass)
    {
        List<string> problems = new();
        string? collection = SingleAttributeArgument(liveClass, "CollectionAttribute");
        if (collection == null)
        {
            problems.Add(liveClass.Name + ": no [Collection], so xunit builds no fixture before its tests and nothing gates them");
        }
        else
        {
            Type? definition = CollectionDefinitionNamed(collection);
            if (definition == null)
            {
                problems.Add(liveClass.Name + ": no collection definition is named '" + collection + "', so no fixture runs first");
            }
            else
            {
                List<Type> fixtures = FixturesOf(definition, typeof(ICollectionFixture<>));
                if (fixtures.Count == 0)
                {
                    problems.Add(liveClass.Name + ": collection '" + collection + "' builds no fixture, so nothing runs before its tests");
                }

                foreach (Type fixture in fixtures)
                {
                    problems.AddRange(ConstructorProblems(fixture).Select(p => liveClass.Name + " via '" + collection + "': " + p));
                }
            }
        }

        // A class fixture is built by the class runner on its own account, so it is not covered by
        // the collection fixture having refused: it has to reach the gate itself.
        foreach (Type fixture in FixturesOf(liveClass, typeof(IClassFixture<>)))
        {
            problems.AddRange(ConstructorProblems(fixture).Select(p => liveClass.Name + " (class fixture): " + p));
        }

        return problems;
    }

    /// <summary>What a fixture's construction can do before the gate - nothing, or a reason.</summary>
    private static List<string> ConstructorProblems(Type fixture)
    {
        List<string> problems = new();

        // Static initializers run before any constructor, so before any gate can. A live fixture
        // may build framework objects there and nothing else.
        ConstructorInfo? typeInitializer = fixture.TypeInitializer;
        if (typeInitializer != null)
        {
            foreach ((OpCode op, MethodBase? target) in CallsIn(typeInitializer))
            {
                if (!IsInertBeforeTheGate(op, target))
                {
                    problems.Add(fixture.Name + ": its static initializer calls " + Describe(target)
                        + ", which runs before any gate can");
                }
            }
        }

        ConstructorInfo[] constructors = fixture.GetConstructors(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
        if (constructors.Length == 0)
        {
            problems.Add(fixture.Name + ": no constructor to put the gate in");
        }

        foreach (ConstructorInfo constructor in constructors)
        {
            string? problem = GateProblem(constructor);
            if (problem != null)
            {
                problems.Add(fixture.Name + ": " + problem);
            }
        }

        return problems;
    }

    /// <summary>
    /// Null when the first thing a constructor does - past its field initializers building
    /// framework objects, and its framework base constructor - is to pass the gate. Anything of
    /// ours before it, and any framework METHOD outside <see cref="InertBeforeTheGate"/>, is a way
    /// for a live run to reach a mailbox without having opted in: COM activation
    /// (<c>Type.GetTypeFromProgID</c>, <c>Activator</c>) and starting the server are framework calls.
    /// </summary>
    private static string? GateProblem(MethodBase constructor)
    {
        foreach ((OpCode op, MethodBase? target) in CallsIn(constructor))
        {
            if (target != null && IsGate(target))
            {
                return null;
            }

            if (!IsInertBeforeTheGate(op, target))
            {
                return "calls " + Describe(target) + " before the opt-in gate";
            }
        }

        return "never reaches the opt-in gate (" + nameof(LiveTestSettings) + "." + nameof(LiveTestSettings.Load)
            + " or " + nameof(LiveRunOptIn) + "." + nameof(LiveRunOptIn.Require) + ")";
    }

    private static bool IsGate(MethodBase method)
    {
        return (method.DeclaringType == typeof(LiveRunOptIn) && method.Name == nameof(LiveRunOptIn.Require))
            || (method.DeclaringType == typeof(LiveTestSettings) && method.Name == nameof(LiveTestSettings.Load));
    }

    /// <summary>
    /// The framework types whose methods a field initializer may call before the gate - the comparer
    /// family that dictionary and set initializers take, and nothing else. A REVIEWED list, not a
    /// guess about what is harmless: a new entry needs a reason, the way <c>LivePhase1Fixture</c>'s
    /// <c>StringComparer.OrdinalIgnoreCase</c> has one (it builds a dictionary, it reaches nothing).
    /// </summary>
    private static readonly Type[] InertBeforeTheGate =
    {
        typeof(StringComparer),
        typeof(EqualityComparer<>),
        typeof(Comparer<>),
    };

    /// <summary>
    /// What may run before the gate: building a framework object - a field initializer's
    /// collection, a Lazy that runs nothing yet, the base constructor - or asking one of
    /// <see cref="InertBeforeTheGate"/> for a comparer.
    /// </summary>
    private static bool IsInertBeforeTheGate(OpCode op, MethodBase? target)
    {
        if (target == null || !IsFramework(target.DeclaringType))
        {
            return false;
        }

        if (target is ConstructorInfo)
        {
            return op == OpCodes.Newobj || op == OpCodes.Call;
        }

        Type declaring = target.DeclaringType!;
        Type definition = declaring.IsGenericType ? declaring.GetGenericTypeDefinition() : declaring;
        return InertBeforeTheGate.Contains(definition);
    }

    private static bool IsFramework(Type? type)
    {
        string name = type?.Assembly.GetName().Name ?? string.Empty;
        return name == "System.Private.CoreLib"
            || name == "mscorlib"
            || name == "netstandard"
            || name.StartsWith("System.", StringComparison.Ordinal)
            || name.StartsWith("Microsoft.", StringComparison.Ordinal);
    }

    private static string Describe(MethodBase? method)
    {
        return method == null ? "a method that could not be resolved" : (method.DeclaringType?.Name ?? "?") + "." + method.Name;
    }

    // ------------------------------------------------------------------ reflection and IL

    /// <summary>Every class with a Category=Live test, whether the trait is on the class or on a method.</summary>
    private static IEnumerable<Type> LiveTestClasses()
    {
        return typeof(LiveCollections).Assembly.GetTypes()
            .Where(t => t.IsClass && !t.IsAbstract)
            .Where(t => IsLive(t.GetCustomAttributesData())
                || t.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly)
                    .Where(m => m.GetCustomAttributes(typeof(FactAttribute), inherit: true).Length > 0)
                    .Any(m => IsLive(m.GetCustomAttributesData())))
            .OrderBy(t => t.FullName, StringComparer.Ordinal);
    }

    private static bool IsLive(IEnumerable<CustomAttributeData> attributes)
    {
        return attributes.Any(a => a.AttributeType == typeof(TraitAttribute)
            && a.ConstructorArguments.Count == 2
            && string.Equals(a.ConstructorArguments[0].Value as string, "Category", StringComparison.Ordinal)
            && string.Equals(a.ConstructorArguments[1].Value as string, "Live", StringComparison.Ordinal));
    }

    private static IEnumerable<Type> FixturesBehind(Type liveClass)
    {
        string? collection = SingleAttributeArgument(liveClass, "CollectionAttribute");
        Type? definition = collection == null ? null : CollectionDefinitionNamed(collection);
        IEnumerable<Type> collectionFixtures = definition == null ? Array.Empty<Type>() : FixturesOf(definition, typeof(ICollectionFixture<>));
        return collectionFixtures.Concat(FixturesOf(liveClass, typeof(IClassFixture<>)));
    }

    private static Type? CollectionDefinitionNamed(string name)
    {
        return typeof(LiveCollections).Assembly.GetTypes()
            .FirstOrDefault(t => string.Equals(SingleAttributeArgument(t, "CollectionDefinitionAttribute"), name, StringComparison.Ordinal));
    }

    private static List<Type> FixturesOf(Type type, Type openFixtureInterface)
    {
        return type.GetInterfaces()
            .Where(i => i.IsGenericType && i.GetGenericTypeDefinition() == openFixtureInterface)
            .Select(i => i.GetGenericArguments()[0])
            .ToList();
    }

    private static string? SingleAttributeArgument(Type type, string attributeName)
    {
        return type.GetCustomAttributesData()
            .Where(a => string.Equals(a.AttributeType.Name, attributeName, StringComparison.Ordinal))
            .Where(a => a.ConstructorArguments.Count == 1)
            .Select(a => a.ConstructorArguments[0].Value as string)
            .FirstOrDefault(value => value != null);
    }

    /// <summary>Every opcode, by its value, so the walk below can step over operands of every size.</summary>
    private static readonly Dictionary<ushort, OpCode> OpCodesByValue = typeof(OpCodes)
        .GetFields(BindingFlags.Public | BindingFlags.Static)
        .Select(f => (OpCode)f.GetValue(null)!)
        .ToDictionary(op => unchecked((ushort)op.Value));

    /// <summary>
    /// The call, callvirt and newobj instructions of one method body, in order, with what each
    /// one calls. A real opcode walk rather than a byte search: a byte search reads operand bytes
    /// as opcodes, and a pin built on that can be green by accident.
    /// </summary>
    private static IEnumerable<(OpCode Op, MethodBase? Target)> CallsIn(MethodBase method)
    {
        byte[] il = method.GetMethodBody()?.GetILAsByteArray() ?? Array.Empty<byte>();
        Type[]? typeArguments = method.DeclaringType is { IsGenericType: true } declaring ? declaring.GetGenericArguments() : null;
        Type[]? methodArguments = method is MethodInfo { IsGenericMethod: true } generic ? generic.GetGenericArguments() : null;

        List<(OpCode, MethodBase?)> calls = new();
        int position = 0;
        while (position < il.Length)
        {
            ushort value = il[position++];
            if (value == 0xFE)
            {
                value = (ushort)(0xFE00 | il[position++]);
            }

            OpCode op = OpCodesByValue[value];
            int operand = position;
            position += OperandSize(op.OperandType, il, operand);
            if (op == OpCodes.Call || op == OpCodes.Callvirt || op == OpCodes.Newobj)
            {
                MethodBase? target;
                try
                {
                    target = method.Module.ResolveMethod(BitConverter.ToInt32(il, operand), typeArguments, methodArguments);
                }
                catch (ArgumentException)
                {
                    target = null;
                }

                calls.Add((op, target));
            }
        }

        return calls;
    }

    private static int OperandSize(OperandType type, byte[] il, int operand)
    {
        switch (type)
        {
            case OperandType.InlineNone:
                return 0;
            case OperandType.ShortInlineBrTarget:
            case OperandType.ShortInlineI:
            case OperandType.ShortInlineVar:
                return 1;
            case OperandType.InlineVar:
                return 2;
            case OperandType.InlineI8:
            case OperandType.InlineR:
                return 8;
            case OperandType.InlineSwitch:
                return 4 + (4 * BitConverter.ToInt32(il, operand));
            default:
                return 4;
        }
    }

    // ------------------------------------------------------------------ plumbing

    /// <summary>
    /// Runs <paramref name="body"/> with the opt-in cleared for this process, and puts it back.
    /// Safe beside other tests: collections run one at a time here (xunit.runner.json), and
    /// nothing in a non-live run reads the variable.
    /// </summary>
    private static void WithoutTheOptIn(Action body)
    {
        string? saved = Environment.GetEnvironmentVariable(LiveRunOptIn.Variable, EnvironmentVariableTarget.Process);
        Environment.SetEnvironmentVariable(LiveRunOptIn.Variable, null, EnvironmentVariableTarget.Process);
        try
        {
            body();
        }
        finally
        {
            Environment.SetEnvironmentVariable(LiveRunOptIn.Variable, saved, EnvironmentVariableTarget.Process);
        }
    }

    private static string RepoRoot()
    {
        string testProjectDir = typeof(LiveRunOptInTests).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");

        // <repo>/McpServer/OutlookAI.McpServer.Tests/ -> <repo>
        return Path.GetFullPath(Path.Combine(testProjectDir, "..", ".."));
    }

    // ------------------------------------------------------------------ controls: never constructed

    // None of these is a test class or a collection definition, so xunit never builds one. They
    // exist to be READ by the check above, which must pass the first two and catch the rest.

    private sealed class GatedControlFixture
    {
        // The shape of LivePhase1Fixture's field initializers: framework objects and a comparer.
        private readonly Dictionary<string, int> _counts = new(StringComparer.OrdinalIgnoreCase);

        public GatedControlFixture()
        {
            Settings = LiveTestSettings.Load();
            _counts[Guid.NewGuid().ToString("N")] = 1;
        }

        public LiveTestSettings Settings { get; }
    }

    private sealed class RequireFirstControlFixture
    {
        public RequireFirstControlFixture()
        {
            LiveRunOptIn.Require();
            Marker = LiveCollections.IsGuarded(LiveCollections.Phase1);
        }

        public bool Marker { get; }
    }

    private sealed class UngatedControlFixture
    {
        public UngatedControlFixture()
        {
            // Only framework construction, so the check walks the whole body and finds no gate.
            Notes = new List<string>();
        }

        public List<string> Notes { get; }
    }

    private sealed class OursBeforeGateControlFixture
    {
        public OursBeforeGateControlFixture()
        {
            Marker = LiveCollections.IsGuarded(LiveCollections.Phase1);
            Settings = LiveTestSettings.Load();
        }

        public bool Marker { get; }

        public LiveTestSettings Settings { get; }
    }

    private sealed class ComBeforeGateControlFixture
    {
        public ComBeforeGateControlFixture()
        {
            Outlook = Type.GetTypeFromProgID("Outlook.Application");
            Settings = LiveTestSettings.Load();
        }

        public Type? Outlook { get; }

        public LiveTestSettings Settings { get; }
    }

    private sealed class StaticInitializerControlFixture
    {
        private static readonly bool Guarded = LiveCollections.IsGuarded(LiveCollections.Phase1);

        public StaticInitializerControlFixture()
        {
            Settings = LiveTestSettings.Load();
            Marker = Guarded;
        }

        public bool Marker { get; }

        public LiveTestSettings Settings { get; }
    }

    private sealed class NoCollectionControlClass
    {
    }

    [Collection("NoSuchCollection-LiveRunOptInTests")]
    private sealed class UnknownCollectionControlClass
    {
    }

    [Collection(LiveCollections.Phase2)]
    private sealed class UngatedClassFixtureControlClass : IClassFixture<UngatedControlFixture>
    {
    }
}
