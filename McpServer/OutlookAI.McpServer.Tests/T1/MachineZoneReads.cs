using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.CompilerServices;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Which tests can read the machine's time zone (Q144, decided by the maintainer 2026-10-11). The
/// build VM runs the suite in UTC (Q126 (a)), where the machine's zone and UTC are the same zone, so
/// a test that builds a local value and compares it with a UTC one proves only the identity there.
/// Every test this finds must carry <c>[Trait("TimeZone", "Machine")]</c>, and the build-VM runner
/// runs those tests a second time with the VM in W. Europe Standard Time (Testbed/README.md section
/// 1c). <see cref="TimeZoneMarkerTests"/> is the check.
/// <para>
/// <b>How.</b> A walk over the compiled IL of this solution - every assembly named <c>OutlookAI.*</c>,
/// the product as well as the tests - from everything xunit runs for a test: the test method, its
/// class's constructors, static constructor and teardown, its fixtures, and the members its
/// <c>[MemberData]</c> and <c>[ClassData]</c> name. Every call, <c>newobj</c>, <c>ldftn</c> and
/// <c>ldvirtftn</c> into this solution is followed, so lambdas, local functions and async and
/// iterator state machines are walked too; so is the static constructor of every type of ours whose
/// member or static field a walked body touches; and an interface or virtual member of ours is
/// followed into every implementation in this test assembly, which is where the fakes a T1 test runs
/// against live. A test reads the machine's zone when the walk reaches one of the framework members
/// <see cref="ReadOf"/> lists. Run on 2026-11-17 it walked 11,667 methods in about a second.
/// </para>
/// <para>
/// <b>What it counts as a read</b> is a framework member that reads the machine's zone for a value of
/// the default kind, <see cref="DateTimeKind.Unspecified"/> - what <c>new DateTime(...)</c> makes - or
/// for any value: <see cref="TimeZoneInfo.Local"/>, <see cref="DateTime.Now"/>, <c>ToLocalTime</c>,
/// <c>ToUniversalTime</c>, a parse with no <see cref="DateTimeStyles"/>, local-time formatting with
/// <c>z</c> or <c>U</c>, and the rest of the list. A member that reads it only for a value already
/// marked <see cref="DateTimeKind.Local"/> - formatting with <c>o</c> or <c>K</c>,
/// <see cref="DateTime.ToBinary"/> - is not counted: the round-trip format <c>o</c> is written 33 times
/// in this solution's source, almost always for a UTC value.
/// </para>
/// <para>
/// <b>It over-counts, on purpose.</b> A walk cannot see a value's kind, so a test is counted when it
/// can reach a read, not only when it does: on 2026-11-17, 145 of the 323 tests it found reach the
/// zone only through a product conversion that reads it for a Local value alone -
/// <c>WsSqlBuilder.FormatUtc</c>, <c>DaslDateLiteral.FormatUtc</c>, <c>IndexRowMapper.AsUtcDateTime</c>,
/// <c>IndexSearchService.GetStaleness</c> and <c>ComDateValue.FromTableValue</c>. Running such a test
/// in a second zone costs milliseconds and proves nothing wrong; missing one is what the second pass
/// exists to prevent.
/// </para>
/// <para>
/// <b>What it cannot see, said once:</b>
/// </para>
/// <list type="bullet">
/// <item>a product interface or virtual member reached through the interface is followed only into
/// this test assembly's implementations, never into the product's own - following those as well would
/// have added 75 tests on 2026-11-17, every one only through <c>OutlookComSession</c>, the real COM
/// session, which no T1 test constructs;</item>
/// <item>a framework virtual member - <c>object.ToString</c>, <c>IComparer&lt;T&gt;.Compare</c> -
/// dispatched to an override of ours, and a delegate built in the framework and invoked here;</item>
/// <item>reflection, <c>dynamic</c>, and code in another process: a T3 test's server, a self-test's
/// script;</item>
/// <item>composite formatting - <c>string.Format("{0:zzz}", value)</c> - and a format string that is
/// not a literal, and a <see cref="DateTimeStyles"/> argument, which is taken as the caller's
/// decision even when it is <c>AssumeLocal</c>;</item>
/// <item>a serializer that writes or reads a <see cref="DateTime"/> (System.Text.Json, XmlConvert),
/// and a Windows function reached through P/Invoke - none in this solution reads the zone today;</item>
/// <item>a value already marked Local - <c>DateTime.SpecifyKind(x, DateTimeKind.Local)</c>,
/// <c>new DateTime(..., DateTimeKind.Local)</c> - handed to a member that reads the zone only for a
/// Local value.</item>
/// </list>
/// <para>
/// A test that reads the zone in one of these ways can be marked by hand: the check never fails a
/// test for carrying the marker.
/// </para>
/// </summary>
internal static class MachineZoneReads
{
    /// <summary>The trait name of the marker: <c>[Trait("TimeZone", "Machine")]</c>.</summary>
    internal const string TraitName = "TimeZone";

    /// <summary>The trait value of the marker.</summary>
    internal const string TraitValue = "Machine";

    /// <summary>The filter the build-VM runner's second pass selects the marked tests with.</summary>
    internal const string Filter = TraitName + "=" + TraitValue;

    private static readonly Assembly Tests = typeof(MachineZoneReads).Assembly;

    private static readonly Lazy<ZoneSurvey> Surveyed = new(SurveyAll);

    /// <summary>One test xunit runs, and the chain by which it reaches a read of the machine's zone.</summary>
    /// <param name="TestClass">The class xunit runs it in.</param>
    /// <param name="Test">The test method.</param>
    /// <param name="Chain">From the test - or what xunit runs around it - to the read, outermost first; null when it reaches none.</param>
    internal sealed record ZoneReadingTest(Type TestClass, MethodInfo Test, IReadOnlyList<string>? Chain)
    {
        /// <summary>The test as the TRX file names it, without its arguments.</summary>
        public string Name => TestClass.FullName + "." + Test.Name;

        /// <summary>Whether it carries <c>[Trait("TimeZone", "Machine")]</c>, on itself or its class.</summary>
        public bool IsMarked => TraitValues(TestClass, Test, TraitName).Contains(TraitValue);
    }

    /// <summary>What one survey found, and how much of the solution it read to find it.</summary>
    /// <param name="Tests">Every non-live test, with its chain or none.</param>
    /// <param name="MethodsWalked">How many method bodies the walk read.</param>
    internal sealed record ZoneSurvey(IReadOnlyList<ZoneReadingTest> Tests, int MethodsWalked);

    /// <summary>
    /// Every test xunit runs from this assembly that is not <c>Category=Live</c> - the build VM's
    /// suite - with the chain by which each reaches a read of the machine's zone. Computed once.
    /// </summary>
    internal static ZoneSurvey Survey() => Surveyed.Value;

    /// <summary>The chain from <paramref name="roots"/> to a read of the machine's zone, or null - what the controls walk.</summary>
    internal static IReadOnlyList<string>? ChainFrom(IEnumerable<(MethodBase Root, string Label)> roots)
    {
        List<(MethodBase Root, string Label)> list = roots.ToList();
        Graph graph = new();
        foreach ((MethodBase root, _) in list)
        {
            graph.Add(root);
        }

        graph.Run();
        return list.Select(r => graph.ChainFrom(r.Root, r.Label)).FirstOrDefault(c => c != null);
    }

    /// <summary>
    /// What xunit runs for one test - <see cref="WritePathAnalyzer.RootsOf"/>: the method, its class's
    /// constructors and teardown, its fixtures' - and, besides those, every static constructor of the
    /// class and its bases, and the members its data attributes name.
    /// </summary>
    internal static List<(MethodBase Root, string Label)> RootsOf(Type testClass, MethodInfo test)
    {
        List<(MethodBase Root, string Label)> roots = WritePathAnalyzer.RootsOf(test);
        for (Type? type = testClass; type != null && IsOurs(type); type = type.BaseType)
        {
            if (type.TypeInitializer != null)
            {
                roots.Add((type.TypeInitializer, "static constructor " + type.Name));
            }
        }

        foreach (CustomAttributeData data in test.GetCustomAttributesData())
        {
            if (data.AttributeType == typeof(MemberDataAttribute) && data.ConstructorArguments.Count >= 1
                && data.ConstructorArguments[0].Value is string memberName)
            {
                Type source = data.NamedArguments
                    .Where(a => a.MemberName == nameof(MemberDataAttribute.MemberType))
                    .Select(a => a.TypedValue.Value as Type)
                    .FirstOrDefault(t => t != null) ?? testClass;
                const BindingFlags Any = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static
                    | BindingFlags.Instance | BindingFlags.FlattenHierarchy;
                foreach (MemberInfo member in source.GetMember(memberName, Any))
                {
                    MethodBase? body = member switch
                    {
                        PropertyInfo property => property.GetMethod,
                        MethodInfo method => method,
                        FieldInfo field => field.DeclaringType?.TypeInitializer,
                        _ => null,
                    };
                    if (body != null)
                    {
                        roots.Add((body, "data " + source.Name + "." + memberName));
                    }
                }
            }
            else if (data.AttributeType == typeof(ClassDataAttribute) && data.ConstructorArguments.Count == 1
                && data.ConstructorArguments[0].Value is Type dataClass)
            {
                foreach (MethodBase body in dataClass.GetConstructors(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                    .Cast<MethodBase>()
                    .Concat(dataClass.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance)
                        .Where(m => m.Name.EndsWith("GetEnumerator", StringComparison.Ordinal))))
                {
                    roots.Add((body, "data " + dataClass.Name));
                }
            }
        }

        return roots;
    }

    /// <summary>
    /// Every test xunit runs from this assembly, live or not: each method carrying <see cref="FactAttribute"/>
    /// or an attribute derived from it (<see cref="TheoryAttribute"/>), in each class xunit discovers -
    /// public, and not abstract unless static - including methods it inherits.
    /// </summary>
    internal static IEnumerable<(Type TestClass, MethodInfo Test)> AllTests()
    {
        const BindingFlags Methods = BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static;
        return Tests.GetExportedTypes()
            .Where(t => t.IsClass && (!t.IsAbstract || t.IsSealed))
            .SelectMany(t => t.GetMethods(Methods)
                .Where(m => m.GetCustomAttributes(typeof(FactAttribute), inherit: true).Length > 0)
                .Select(m => (t, m)))
            .OrderBy(x => x.t.FullName, StringComparer.Ordinal)
            .ThenBy(x => x.m.Name, StringComparer.Ordinal);
    }

    /// <summary>
    /// Every value of the trait <paramref name="name"/> on a test, read where xunit reads it: the
    /// assembly, the method and the methods it overrides, and the class and its base classes.
    /// </summary>
    internal static IReadOnlyList<string> TraitValues(Type testClass, MethodInfo test, string name)
    {
        IEnumerable<CustomAttributeData> attributes = Tests.GetCustomAttributesData().Concat(test.GetCustomAttributesData());
        MethodInfo overridden = test.GetBaseDefinition();
        if (overridden != test)
        {
            attributes = attributes.Concat(overridden.GetCustomAttributesData());
        }

        for (Type? type = testClass; type != null; type = type.BaseType)
        {
            attributes = attributes.Concat(type.GetCustomAttributesData());
        }

        return attributes
            .Where(a => a.AttributeType == typeof(TraitAttribute) && a.ConstructorArguments.Count == 2)
            .Where(a => string.Equals(a.ConstructorArguments[0].Value as string, name, StringComparison.Ordinal))
            .Select(a => a.ConstructorArguments[1].Value as string)
            .Where(v => v != null)
            .Select(v => v!)
            .ToList();
    }

    /// <summary>
    /// The read of the machine's time zone that calling <paramref name="member"/> is, or null. For a
    /// member that formats a <see cref="DateTime"/>, <paramref name="format"/> is the string literal
    /// pushed just before the call, if there was one.
    /// </summary>
    internal static string? ReadOf(MethodBase member, string? format)
    {
        Type? type = member.DeclaringType;
        string name = member.Name;
        ParameterInfo[] parameters = member.GetParameters();
        bool takesDateTimeFirst = parameters.Length > 0 && parameters[0].ParameterType == typeof(DateTime);

        // Only an overload that takes a format can be given one: ToString(IFormatProvider) cannot,
        // whatever string happens to be pushed before it.
        if (!parameters.Any(p => p.Name == "format"))
        {
            format = null;
        }

        if (type == typeof(TimeZoneInfo))
        {
            return name switch
            {
                "get_Local" => "TimeZoneInfo.Local",
                nameof(TimeZoneInfo.ConvertTimeToUtc) when parameters.Length == 1
                    => "TimeZoneInfo.ConvertTimeToUtc(DateTime), which converts from the machine's zone",
                nameof(TimeZoneInfo.ConvertTime) when parameters.Length == 2 && takesDateTimeFirst
                    => "TimeZoneInfo.ConvertTime(DateTime, TimeZoneInfo), which converts from the machine's zone",
                nameof(TimeZoneInfo.ConvertTimeBySystemTimeZoneId) when parameters.Length == 2 && takesDateTimeFirst
                    => "TimeZoneInfo.ConvertTimeBySystemTimeZoneId(DateTime, string), which converts from the machine's zone",
                _ => null,
            };
        }

        if (type == typeof(DateTime))
        {
            switch (name)
            {
                case "get_Now":
                    return "DateTime.Now";
                case "get_Today":
                    return "DateTime.Today";
                case nameof(DateTime.ToLocalTime):
                case nameof(DateTime.ToUniversalTime):
                case nameof(DateTime.ToFileTime):
                case nameof(DateTime.FromFileTime):
                    return "DateTime." + name;
                case nameof(DateTime.ToString):
                case nameof(DateTime.TryFormat):
                    return FormatRead("DateTime." + name, format);
            }

            return ParseRead("DateTime", name, parameters);
        }

        if (type == typeof(DateTimeOffset))
        {
            return name switch
            {
                "get_Now" => "DateTimeOffset.Now",
                "get_LocalDateTime" => "DateTimeOffset.LocalDateTime",
                nameof(DateTimeOffset.ToLocalTime) => "DateTimeOffset.ToLocalTime",
                ".ctor" or "op_Implicit" when parameters.Length == 1 && takesDateTimeFirst
                    => "a DateTimeOffset made from a DateTime, which takes the machine's offset for a value that is not UTC",
                _ => ParseRead("DateTimeOffset", name, parameters),
            };
        }

        if (type == typeof(TimeProvider))
        {
            return name is "get_LocalTimeZone" or nameof(TimeProvider.GetLocalNow) ? "TimeProvider." + name : null;
        }

        if (type == typeof(File) || type == typeof(Directory))
        {
            return name is "GetCreationTime" or "GetLastAccessTime" or "GetLastWriteTime"
                or "SetCreationTime" or "SetLastAccessTime" or "SetLastWriteTime"
                ? type.Name + "." + name + ", in local time"
                : null;
        }

        if (type == typeof(FileSystemInfo))
        {
            return name is "get_CreationTime" or "get_LastAccessTime" or "get_LastWriteTime"
                or "set_CreationTime" or "set_LastAccessTime" or "set_LastWriteTime"
                ? "FileSystemInfo." + name + ", in local time"
                : null;
        }

        if (type == typeof(Process))
        {
            return name is "get_StartTime" or "get_ExitTime" ? "Process." + name + ", in local time" : null;
        }

        if (type == typeof(DefaultInterpolatedStringHandler) && name == nameof(DefaultInterpolatedStringHandler.AppendFormatted)
            && member is MethodInfo { IsGenericMethod: true } generic
            && generic.GetGenericArguments()[0] is Type value
            && (value == typeof(DateTime) || value == typeof(DateTime?)))
        {
            return FormatRead("an interpolated DateTime", format);
        }

        return null;

        static string? ParseRead(string typeName, string name, ParameterInfo[] parameters)
        {
            bool parse = name is "Parse" or "ParseExact" or "TryParse" or "TryParseExact";
            bool styles = parameters.Any(p => p.ParameterType == typeof(DateTimeStyles));
            return parse && !styles
                ? typeName + "." + name + " with no DateTimeStyles, which reads the machine's zone for a time that has, or lacks, an offset"
                : null;
        }

        static string? FormatRead(string what, string? format)
        {
            return format != null && FormatWritesTheMachinesOffset(format)
                ? what + " with the format '" + format + "', which writes or converts by the machine's offset"
                : null;
        }
    }

    /// <summary>
    /// Whether a <see cref="DateTime"/> format reads the machine's zone for an Unspecified or Local value:
    /// the standard format <c>U</c>, which converts to UTC, or a custom format with an unquoted
    /// <c>z</c>, which writes the offset. <c>o</c> and <c>K</c> read it only for a Local value and do
    /// not count (see the class remarks).
    /// </summary>
    internal static bool FormatWritesTheMachinesOffset(string format)
    {
        if (format.Length == 1)
        {
            return format == "U";
        }

        char quote = '\0';
        for (int i = 0; i < format.Length; i++)
        {
            char c = format[i];
            if (quote != '\0')
            {
                if (c == quote)
                {
                    quote = '\0';
                }

                continue;
            }

            if (c == '\\')
            {
                i++;
            }
            else if (c is '\'' or '"')
            {
                quote = c;
            }
            else if (c == 'z')
            {
                return true;
            }
        }

        return false;
    }

    private static ZoneSurvey SurveyAll()
    {
        List<(Type TestClass, MethodInfo Test, List<(MethodBase Root, string Label)> Roots)> tests = AllTests()
            .Where(t => !TraitValues(t.TestClass, t.Test, "Category").Contains("Live"))
            .Select(t => (t.TestClass, t.Test, RootsOf(t.TestClass, t.Test)))
            .ToList();

        Graph graph = new();
        foreach ((_, _, List<(MethodBase Root, string Label)> roots) in tests)
        {
            foreach ((MethodBase root, _) in roots)
            {
                graph.Add(root);
            }
        }

        graph.Run();
        List<ZoneReadingTest> found = tests
            .Select(t => new ZoneReadingTest(
                t.TestClass,
                t.Test,
                t.Roots.Select(r => graph.ChainFrom(r.Root, r.Label)).FirstOrDefault(c => c != null)))
            .ToList();
        return new ZoneSurvey(found, graph.Count);
    }

    private static bool IsOurs(Type? type)
    {
        return type != null && (type.Assembly.GetName().Name ?? string.Empty).StartsWith("OutlookAI.", StringComparison.Ordinal);
    }

    /// <summary>
    /// The call graph of this solution's code reachable from the roots added, and for each method the
    /// next step towards a read of the machine's zone. Built forwards, then answered backwards from the
    /// methods that read the zone themselves, so a cycle cannot hide a read and every chain is a
    /// shortest one.
    /// </summary>
    private sealed class Graph
    {
        private readonly Dictionary<MethodBase, List<MethodBase>> _calls = new();
        private readonly Dictionary<MethodBase, string> _reads = new();
        private readonly Queue<MethodBase> _queue = new();
        private readonly Dictionary<MethodBase, MethodBase[]> _implementations = new();
        private readonly Dictionary<MethodBase, MethodBase?> _next = new();

        /// <summary>How many methods the walk reached.</summary>
        public int Count => _calls.Count;

        public void Add(MethodBase method)
        {
            if (_calls.TryAdd(method, new List<MethodBase>()))
            {
                _queue.Enqueue(method);
            }
        }

        public void Run()
        {
            while (_queue.Count > 0)
            {
                Expand(_queue.Dequeue());
            }

            Dictionary<MethodBase, List<MethodBase>> callers = new();
            foreach ((MethodBase caller, List<MethodBase> callees) in _calls)
            {
                foreach (MethodBase callee in callees)
                {
                    if (!callers.TryGetValue(callee, out List<MethodBase>? list))
                    {
                        callers[callee] = list = new List<MethodBase>();
                    }

                    list.Add(caller);
                }
            }

            Queue<MethodBase> work = new();
            foreach (MethodBase reader in _reads.Keys)
            {
                _next[reader] = null;
                work.Enqueue(reader);
            }

            while (work.Count > 0)
            {
                MethodBase method = work.Dequeue();
                if (!callers.TryGetValue(method, out List<MethodBase>? list))
                {
                    continue;
                }

                foreach (MethodBase caller in list)
                {
                    if (_next.TryAdd(caller, method))
                    {
                        work.Enqueue(caller);
                    }
                }
            }
        }

        public IReadOnlyList<string>? ChainFrom(MethodBase root, string label)
        {
            if (!_next.ContainsKey(root))
            {
                return null;
            }

            List<string> chain = new() { label };
            MethodBase cursor = root;
            while (_next[cursor] is MethodBase next)
            {
                chain.Add(WritePathAnalyzer.Describe(next));
                cursor = next;
            }

            chain.Add("reads " + _reads[cursor]);
            return chain;
        }

        private void Expand(MethodBase method)
        {
            foreach (MethodBase moveNext in WritePathAnalyzer.StateMachineOf(method))
            {
                Link(method, moveNext);
            }

            IReadOnlyList<IlInstruction> body = IlReader.Read(method, withFields: true);
            for (int i = 0; i < body.Count; i++)
            {
                IlInstruction instruction = body[i];
                if (instruction.Field is FieldInfo field)
                {
                    if (field.IsStatic && IsOurs(field.DeclaringType) && field.DeclaringType!.TypeInitializer is ConstructorInfo initializer)
                    {
                        Link(method, initializer);
                    }

                    continue;
                }

                if (instruction.Method is not MethodBase target)
                {
                    continue;
                }

                string? read = ReadOf(target, FormatBefore(body, i));
                if (read != null)
                {
                    _reads.TryAdd(method, read);
                    continue;
                }

                if (!IsOurs(target.DeclaringType))
                {
                    continue;
                }

                Link(method, target);
                if (target.DeclaringType!.TypeInitializer is ConstructorInfo typeInitializer)
                {
                    Link(method, typeInitializer);
                }

                foreach (MethodBase implementation in ImplementationsOf(target))
                {
                    Link(method, implementation);
                }
            }
        }

        private void Link(MethodBase from, MethodBase to)
        {
            _calls[from].Add(to);
            Add(to);
        }

        private MethodBase[] ImplementationsOf(MethodBase target)
        {
            if (!_implementations.TryGetValue(target, out MethodBase[]? found))
            {
                try
                {
                    found = WritePathAnalyzer.ImplementationsOf(target).Distinct().ToArray();
                }
                catch (ArgumentException)
                {
                    found = Array.Empty<MethodBase>();
                }

                _implementations[target] = found;
            }

            return found;
        }

        /// <summary>
        /// The format a formatting call at <paramref name="at"/> is given: the string literal pushed
        /// just before it, or just before the one call that may stand between them - the culture
        /// getter of <c>ToString(format, CultureInfo.InvariantCulture)</c>, the span conversion of
        /// <c>TryFormat</c>.
        /// </summary>
        private static string? FormatBefore(IReadOnlyList<IlInstruction> body, int at)
        {
            for (int i = at - 1; i >= 0 && i >= at - 2; i--)
            {
                if (body[i].Text != null)
                {
                    return body[i].Text;
                }

                if (body[i].Method == null)
                {
                    return null;
                }
            }

            return null;
        }
    }
}
