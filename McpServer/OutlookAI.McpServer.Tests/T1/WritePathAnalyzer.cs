using System.Reflection;
using System.Runtime.CompilerServices;
using OutlookAI.McpServer.Tests.T2;
using OutlookAI.McpServer.Tests.T3;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>One reason a live test cannot be shown to write nothing, with the call chain that reaches it.</summary>
/// <param name="Chain">From the test (or its fixture) to the offending call, outermost first.</param>
/// <param name="Reason">What the last link does.</param>
internal sealed record WritePathFinding(IReadOnlyList<string> Chain, string Reason)
{
    public override string ToString() => string.Join(" -> ", Chain) + " :: " + Reason;
}

/// <summary>What one walk covered, so a pin can refuse to pass having walked nothing.</summary>
/// <param name="MethodsWalked">Bodies of this assembly the walk read.</param>
/// <param name="ProductCalls">Distinct product members the walk saw called, read-only or not.</param>
internal sealed record WritePathWalk(IReadOnlyList<WritePathFinding> Findings, int MethodsWalked, IReadOnlyCollection<string> ProductCalls);

/// <summary>
/// The static half of Q74 layer 1: can a live test - its body AND everything that runs for it, its
/// class's constructor and teardown and its collection and class fixtures' - reach a write?
/// <para>
/// <b>How.</b> A transitive walk over the compiled IL of THIS assembly, from the test method and every
/// fixture member xunit runs around it: every <c>call</c>, <c>callvirt</c>, <c>newobj</c>, <c>ldftn</c>
/// and <c>ldvirtftn</c> is followed - so lambdas, local functions, <c>Lazy</c> factories and async and
/// iterator state machines are walked too - and an interface or virtual method declared here is followed
/// into every implementation here. Each call that LEAVES this assembly is judged where it lands:
/// </para>
/// <list type="bullet">
/// <item><b>The write guard</b> (<see cref="LiveStoreWriteGuard"/>, <see cref="StoreWriteAllowlist.Assert"/>)
/// - every sanctioned mailbox write asks it first (mailbox-safety rules 1 and 3), so reaching it is
/// reaching a write.</item>
/// <item><b>A product method</b> (OutlookAI.Core, .McpServer, .ComHost, .RemediationTools) - allowed only
/// when it is on <see cref="ReadOnlyProductApi"/>, a reviewed list of members that change nothing.
/// Anything else - a write, or a member nobody has classified yet - is a finding. The walk deliberately
/// stops at this boundary: inside the product a read and a write share helpers behind runtime flags,
/// which no static walk can tell apart.</item>
/// <item><b>An MCP tool named in this assembly's code</b> - a string literal that is a tool classified
/// <see cref="McpToolEffect.Writes"/> in <see cref="McpToolWriteClassification"/>.</item>
/// <item><b>A COM mutation through <c>dynamic</c></b> - a property SET on a late-bound object, or a
/// late-bound call to a member that changes an item or the session (<see cref="MutatingComMembers"/>).</item>
/// <item><b>A registry write</b> through <c>Microsoft.Win32</c>.</item>
/// </list>
/// <para>
/// <b>What it cannot see, said once:</b> a static constructor's body (a field initializer runs before
/// any gate and is pinned elsewhere - <c>LiveRunOptInTests</c>); a delegate built in the product and
/// invoked here; a framework interface dispatched to an implementation of ours that the walk never saw
/// constructed (a type it DOES see constructed has its <c>Dispose</c> walked); reflection; and the bounded
/// re-run's child process, which is a separate, itself-guarded live run and since Q74 A1 is reached by no
/// declared machine profile. Layers 2 and 3 are the run-time gates behind every one of those.
/// </para>
/// </summary>
internal static class WritePathAnalyzer
{
    /// <summary>
    /// Late-bound member names that change an Outlook item, folder or session when CALLED. Reviewed
    /// against what the live tier's own read paths call: none of their late-bound calls is on this list.
    /// </summary>
    internal static readonly HashSet<string> MutatingComMembers = new(StringComparer.OrdinalIgnoreCase)
    {
        "Delete", "PermanentlyDelete", "Save", "SaveAs", "Send", "Move", "Copy", "Display", "Close", "Quit",
        "SendAndReceive", "Reply", "ReplyAll", "Forward", "CreateItem", "MarkAsTask", "ClearTaskFlag",
        "ShowCategoriesDialog", "AddStore", "AddStoreEx", "RemoveStore",

        // A FOLDER moved or copied (Folder.MoveTo / Folder.CopyTo) - added 2026-10-03 with the Q114
        // test-folder move helper, the first late-bound MoveTo in this assembly.
        "MoveTo", "CopyTo",
    };

    private static readonly Assembly Tests = typeof(LiveCollections).Assembly;

    private static readonly HashSet<string> ProductAssemblies = new(StringComparer.Ordinal)
    {
        "OutlookAI.Core", "OutlookAI.McpServer", "OutlookAI.ComHost", "OutlookAI.RemediationTools",
    };

    /// <summary>Walks one test and everything xunit runs around it.</summary>
    public static WritePathWalk Analyze(MethodInfo test)
    {
        ArgumentNullException.ThrowIfNull(test);
        List<(MethodBase Root, string Label)> roots = RootsOf(test);
        return Walk(roots);
    }

    /// <summary>Walks from arbitrary roots - the controls in the T1 pin use this.</summary>
    public static WritePathWalk Walk(IEnumerable<(MethodBase Root, string Label)> roots)
    {
        Dictionary<MethodBase, (MethodBase? Parent, string Label)> seen = new();
        Queue<MethodBase> queue = new();
        List<WritePathFinding> findings = new();
        HashSet<string> productCalls = new(StringComparer.Ordinal);

        foreach ((MethodBase root, string label) in roots)
        {
            if (seen.TryAdd(root, (null, label)))
            {
                queue.Enqueue(root);
            }
        }

        while (queue.Count > 0)
        {
            MethodBase method = queue.Dequeue();
            foreach (MethodBase follow in StateMachineOf(method))
            {
                Enqueue(follow, method);
            }

            IReadOnlyList<IlInstruction> body = IlReader.Read(method);
            for (int i = 0; i < body.Count; i++)
            {
                IlInstruction instruction = body[i];
                if (instruction.Text != null)
                {
                    if (McpToolWriteClassification.Classify(instruction.Text) == McpToolEffect.Writes
                        && !ReadOnlyProductApi.RefusalOnlyToolNames.ContainsKey(Describe(method)))
                    {
                        findings.Add(Finding(method, "names the write-capable MCP tool '" + instruction.Text + "'"));
                    }

                    continue;
                }

                MethodBase? target = instruction.Method;
                if (target == null)
                {
                    findings.Add(Finding(method, "calls a method the walk could not resolve, so it cannot be shown to write nothing"));
                    continue;
                }

                Judge(method, target, body, i);
            }
        }

        return new WritePathWalk(findings, seen.Count, productCalls);

        void Judge(MethodBase caller, MethodBase target, IReadOnlyList<IlInstruction> body, int at)
        {
            Type? declaring = target.DeclaringType;
            Assembly? assembly = declaring?.Assembly;
            if (assembly == Tests)
            {
                string? guard = WriteGuardCall(target);
                if (guard != null)
                {
                    findings.Add(Finding(caller, guard));
                    return;
                }

                Enqueue(target, caller);
                foreach (MethodBase implementation in ImplementationsOf(target))
                {
                    Enqueue(implementation, caller);
                }

                if (target is ConstructorInfo && declaring != null)
                {
                    foreach (MethodBase teardown in TeardownOf(declaring))
                    {
                        Enqueue(teardown, caller);
                    }
                }

                return;
            }

            string name = assembly?.GetName().Name ?? string.Empty;
            if (ProductAssemblies.Contains(name))
            {
                string member = ReadOnlyProductApi.KeyOf(target);
                productCalls.Add(member);
                if (!ReadOnlyProductApi.IsReadOnly(target))
                {
                    findings.Add(Finding(caller, "calls the product member " + member
                        + ", which is not on ReadOnlyProductApi - a write, or a member nobody has shown to be read-only"));
                }

                return;
            }

            string? framework = FrameworkWrite(target, body, at);
            if (framework != null)
            {
                findings.Add(Finding(caller, framework));
            }
        }

        void Enqueue(MethodBase next, MethodBase from)
        {
            if (next.DeclaringType?.Assembly != Tests)
            {
                return;
            }

            if (seen.TryAdd(next, (from, Describe(next))))
            {
                queue.Enqueue(next);
            }
        }

        WritePathFinding Finding(MethodBase at, string reason)
        {
            List<string> chain = new();
            MethodBase? cursor = at;
            int guardRail = 0;
            while (cursor != null && guardRail++ < 200)
            {
                (MethodBase? parent, string label) = seen[cursor];
                chain.Add(label);
                cursor = parent;
            }

            chain.Reverse();
            return new WritePathFinding(chain, reason);
        }
    }

    /// <summary>
    /// The CHECK ON THE LIST: walks INSIDE the product from <paramref name="starts"/> - direct calls,
    /// lambdas, state machines, and interface or virtual dispatch to every product implementation -
    /// and reports every write it can see there: an audit-log append (the product audits every write
    /// it makes, v3.MD S1), a late-bound COM mutation, a registry write, or a file written, moved or
    /// deleted. A member on <see cref="ReadOnlyProductApi"/> must reach none of them, so the list is a
    /// claim the build re-checks rather than one it trusts.
    /// </summary>
    public static WritePathWalk WalkProduct(IEnumerable<(MethodBase Root, string Label)> starts)
    {
        Dictionary<MethodBase, (MethodBase? Parent, string Label)> seen = new();
        Queue<MethodBase> queue = new();
        List<WritePathFinding> findings = new();
        foreach ((MethodBase root, string label) in starts)
        {
            if (seen.TryAdd(root, (null, label)))
            {
                queue.Enqueue(root);
            }
        }

        while (queue.Count > 0)
        {
            MethodBase method = queue.Dequeue();
            foreach (MethodBase follow in StateMachineOf(method))
            {
                Enqueue(follow, method);
            }

            IReadOnlyList<IlInstruction> body = IlReader.Read(method);
            for (int i = 0; i < body.Count; i++)
            {
                MethodBase? target = body[i].Method;
                if (target == null)
                {
                    continue;
                }

                if (target.DeclaringType?.FullName == "OutlookAI.Core.Audit.AuditLog" && target.Name == "Append")
                {
                    findings.Add(Finding(method, "appends an audit line - the product audits every write it makes"));
                    continue;
                }

                string name = target.DeclaringType?.Assembly.GetName().Name ?? string.Empty;
                if (ProductAssemblies.Contains(name))
                {
                    Enqueue(target, method);
                    foreach (MethodBase implementation in ProductImplementationsOf(target))
                    {
                        Enqueue(implementation, method);
                    }

                    continue;
                }

                string? framework = FrameworkWrite(target, body, i) ?? FileWrite(target);
                if (framework != null && !Exempt(method, target, body, i))
                {
                    findings.Add(Finding(method, framework));
                }
            }
        }

        return new WritePathWalk(findings, seen.Count, Array.Empty<string>());

        static bool Exempt(MethodBase method, MethodBase target, IReadOnlyList<IlInstruction> body, int at)
        {
            if (target.DeclaringType?.FullName != "Microsoft.CSharp.RuntimeBinder.Binder" || target.Name != "InvokeMember")
            {
                return false;
            }

            string key = ReadOnlyProductApi.SourceKeyOf(method) + "|" + (BinderMemberName(body, at) ?? "?");
            return ReadOnlyProductApi.InProductExemptions.ContainsKey(key);
        }

        void Enqueue(MethodBase next, MethodBase from)
        {
            string name = next.DeclaringType?.Assembly.GetName().Name ?? string.Empty;
            if (ProductAssemblies.Contains(name) && seen.TryAdd(next, (from, Describe(next))))
            {
                queue.Enqueue(next);
            }
        }

        WritePathFinding Finding(MethodBase at, string reason)
        {
            List<string> chain = new();
            MethodBase? cursor = at;
            int guardRail = 0;
            while (cursor != null && guardRail++ < 200)
            {
                (MethodBase? parent, string label) = seen[cursor];
                chain.Add(label);
                cursor = parent;
            }

            chain.Reverse();
            return new WritePathFinding(chain, reason);
        }
    }

    /// <summary>A file written, moved, copied or deleted through <c>System.IO</c>.</summary>
    private static string? FileWrite(MethodBase target)
    {
        string type = target.DeclaringType?.FullName ?? string.Empty;
        if ((type == "System.IO.File" && (target.Name.StartsWith("Write", StringComparison.Ordinal)
                || target.Name.StartsWith("Append", StringComparison.Ordinal)
                || target.Name is "Delete" or "Move" or "Copy" or "Replace" or "Create" or "CreateText"))
            || (type == "System.IO.Directory" && target.Name is "Delete" or "Move")
            || (type == "System.IO.FileInfo" && target.Name is "Delete" or "MoveTo" or "CopyTo" or "Create" or "CreateText"))
        {
            return "writes the file system (" + target.DeclaringType!.Name + "." + target.Name + ")";
        }

        return null;
    }

    /// <summary>Product implementations of a product interface member or overrides of a product virtual one.</summary>
    private static IEnumerable<MethodBase> ProductImplementationsOf(MethodBase target)
    {
        if (target is not MethodInfo method || !(method.IsVirtual || method.IsAbstract))
        {
            yield break;
        }

        Type declaring = method.DeclaringType!;
        if (declaring.ContainsGenericParameters)
        {
            yield break;
        }

        foreach (Type type in ProductTypes.Value.Where(t => t.IsClass && !t.IsAbstract && !t.ContainsGenericParameters))
        {
            if (declaring.IsInterface)
            {
                if (!declaring.IsAssignableFrom(type))
                {
                    continue;
                }

                InterfaceMapping map;
                try
                {
                    map = type.GetInterfaceMap(declaring);
                }
                catch (ArgumentException)
                {
                    continue;
                }

                int index = Array.IndexOf(map.InterfaceMethods, method);
                if (index >= 0)
                {
                    yield return map.TargetMethods[index];
                }
            }
            else if (type != declaring && declaring.IsAssignableFrom(type))
            {
                foreach (MethodInfo candidate in type.GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.DeclaredOnly))
                {
                    if (candidate.GetBaseDefinition() == method.GetBaseDefinition())
                    {
                        yield return candidate;
                    }
                }
            }
        }
    }

    private static readonly Lazy<Type[]> ProductTypes = new(() => AppDomain.CurrentDomain.GetAssemblies()
        .Where(a => ProductAssemblies.Contains(a.GetName().Name ?? string.Empty))
        .SelectMany(a =>
        {
            try
            {
                return a.GetTypes();
            }
            catch (ReflectionTypeLoadException e)
            {
                return e.Types.Where(t => t != null).Select(t => t!);
            }
        })
        .ToArray());

    /// <summary>
    /// What xunit runs for one test: the method, its class's constructors and teardown, and every
    /// collection and class fixture's constructors and teardown.
    /// </summary>
    public static List<(MethodBase Root, string Label)> RootsOf(MethodInfo test)
    {
        Type testClass = test.DeclaringType ?? throw new ArgumentException("a test method needs a class", nameof(test));
        List<(MethodBase, string)> roots = new() { (test, Describe(test)) };

        foreach (MethodBase member in ConstructorsOf(testClass).Concat(TeardownOf(testClass)))
        {
            roots.Add((member, Describe(member)));
        }

        foreach (Type fixture in FixturesOf(testClass))
        {
            foreach (MethodBase member in ConstructorsOf(fixture).Concat(TeardownOf(fixture)))
            {
                roots.Add((member, "fixture " + Describe(member)));
            }
        }

        return roots;
    }

    /// <summary>The collection fixtures of <paramref name="testClass"/>'s collection, and its class fixtures.</summary>
    public static IEnumerable<Type> FixturesOf(Type testClass)
    {
        string? collection = testClass.GetCustomAttributesData()
            .Where(a => a.AttributeType == typeof(CollectionAttribute) && a.ConstructorArguments.Count == 1)
            .Select(a => a.ConstructorArguments[0].Value as string)
            .FirstOrDefault(v => v != null);

        IEnumerable<Type> collectionFixtures = Array.Empty<Type>();
        if (collection != null)
        {
            Type? definition = Tests.GetTypes().FirstOrDefault(t => t.GetCustomAttributesData().Any(a =>
                a.AttributeType == typeof(CollectionDefinitionAttribute)
                && a.ConstructorArguments.Count >= 1
                && string.Equals(a.ConstructorArguments[0].Value as string, collection, StringComparison.Ordinal)));
            if (definition != null)
            {
                collectionFixtures = GenericArgumentsOf(definition, typeof(ICollectionFixture<>));
            }
        }

        return collectionFixtures.Concat(GenericArgumentsOf(testClass, typeof(IClassFixture<>))).Distinct();
    }

    private static IEnumerable<Type> GenericArgumentsOf(Type type, Type openInterface)
    {
        return type.GetInterfaces()
            .Where(i => i.IsGenericType && i.GetGenericTypeDefinition() == openInterface)
            .Select(i => i.GetGenericArguments()[0]);
    }

    private static IEnumerable<MethodBase> ConstructorsOf(Type type)
    {
        return type.GetConstructors(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
    }

    /// <summary>Dispose, DisposeAsync and IAsyncLifetime's two halves, wherever <paramref name="type"/> declares them.</summary>
    private static IEnumerable<MethodBase> TeardownOf(Type type)
    {
        return type.GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic)
            .Where(m => m.DeclaringType?.Assembly == Tests)
            .Where(m => m.GetParameters().Length == 0)
            .Where(m => m.Name is "Dispose" or "DisposeAsync" or "InitializeAsync"
                || m.Name.EndsWith(".Dispose", StringComparison.Ordinal)
                || m.Name.EndsWith(".DisposeAsync", StringComparison.Ordinal));
    }

    /// <summary>The compiler-generated MoveNext of an async or iterator method - where its body really is.</summary>
    private static IEnumerable<MethodBase> StateMachineOf(MethodBase method)
    {
        foreach (StateMachineAttribute attribute in method.GetCustomAttributes<StateMachineAttribute>())
        {
            MethodInfo? moveNext = attribute.StateMachineType.GetMethod(
                "MoveNext", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            if (moveNext != null)
            {
                yield return moveNext;
            }
        }
    }

    /// <summary>
    /// Every implementation in this assembly of an interface member, or override of a virtual one,
    /// declared in this assembly - so a call through <c>ITripwireRetrySource</c> is followed into the
    /// live source as well as into whatever else implements it.
    /// </summary>
    private static IEnumerable<MethodBase> ImplementationsOf(MethodBase target)
    {
        if (target is not MethodInfo method || !(method.IsVirtual || method.IsAbstract))
        {
            yield break;
        }

        Type declaring = method.DeclaringType!;
        foreach (Type type in Tests.GetTypes().Where(t => t.IsClass && !t.IsAbstract))
        {
            if (declaring.IsInterface)
            {
                if (!declaring.IsAssignableFrom(type) || declaring.ContainsGenericParameters)
                {
                    continue;
                }

                InterfaceMapping map = type.GetInterfaceMap(declaring);
                int index = Array.IndexOf(map.InterfaceMethods, method);
                if (index >= 0)
                {
                    yield return map.TargetMethods[index];
                }
            }
            else if (type != declaring && declaring.IsAssignableFrom(type))
            {
                foreach (MethodInfo candidate in type.GetMethods(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.DeclaredOnly))
                {
                    if (candidate.GetBaseDefinition() == method.GetBaseDefinition())
                    {
                        yield return candidate;
                    }
                }
            }
        }
    }

    /// <summary>The write guard's entry points: reaching one is reaching a write.</summary>
    private static string? WriteGuardCall(MethodBase target)
    {
        if (target.DeclaringType == typeof(LiveStoreWriteGuard)
            && target.Name is nameof(LiveStoreWriteGuard.Assert) or nameof(LiveStoreWriteGuard.Writable))
        {
            return "asks the write guard (" + target.Name + ") - every mailbox write does, and only a write does";
        }

        if (target.DeclaringType == typeof(StoreWriteAllowlist) && target.Name == nameof(StoreWriteAllowlist.Assert))
        {
            return "asks the write allowlist to permit a write (StoreWriteAllowlist.Assert)";
        }

        return null;
    }

    /// <summary>A write the framework makes for us: a late-bound COM mutation, or a registry write.</summary>
    private static string? FrameworkWrite(MethodBase target, IReadOnlyList<IlInstruction> body, int at)
    {
        string type = target.DeclaringType?.FullName ?? string.Empty;
        if (type == "Microsoft.CSharp.RuntimeBinder.Binder")
        {
            if (target.Name is "SetMember" or "SetIndex")
            {
                string member = target.Name == "SetMember" ? " '" + (BinderMemberName(body, at) ?? "?") + "'" : string.Empty;
                return "SETS a property" + member + " on a late-bound (COM) object";
            }

            if (target.Name == "InvokeMember")
            {
                string? member = BinderMemberName(body, at);
                if (member == null || MutatingComMembers.Contains(member))
                {
                    return "calls the late-bound (COM) member '" + (member ?? "?") + "', which changes an item, a folder or the session";
                }
            }

            return null;
        }

        if (type.StartsWith("Microsoft.Win32.Registry", StringComparison.Ordinal)
            && (target.Name is "SetValue" or "DeleteValue" or "CreateSubKey" or "DeleteSubKey" or "DeleteSubKeyTree"))
        {
            return "writes the registry (" + target.DeclaringType!.Name + "." + target.Name + ")";
        }

        return null;
    }

    /// <summary>
    /// The member name a <c>Binder.InvokeMember</c> or <c>SetMember</c> call site binds: the last string
    /// pushed before it that is not an argument NAME (those go straight into <c>CSharpArgumentInfo.Create</c>).
    /// </summary>
    private static string? BinderMemberName(IReadOnlyList<IlInstruction> body, int at)
    {
        for (int i = at - 1; i >= 0; i--)
        {
            IlInstruction candidate = body[i];
            if (candidate.Text == null)
            {
                if (candidate.Method?.DeclaringType?.FullName == "Microsoft.CSharp.RuntimeBinder.Binder")
                {
                    return null;
                }

                continue;
            }

            bool argumentName = i + 1 < body.Count
                && body[i + 1].Method?.DeclaringType?.FullName == "Microsoft.CSharp.RuntimeBinder.CSharpArgumentInfo";
            if (!argumentName)
            {
                return candidate.Text;
            }
        }

        return null;
    }

    private static string Describe(MethodBase method)
    {
        string type = method.DeclaringType?.Name ?? "?";

        // A compiler-generated host (state machine, display class) reads better as its source member.
        int open = type.IndexOf('<');
        int close = type.IndexOf('>');
        if (open >= 0 && close > open)
        {
            type = (method.DeclaringType?.DeclaringType?.Name ?? string.Empty) + "." + type.Substring(open + 1, close - open - 1);
        }

        return type + "." + method.Name;
    }
}
