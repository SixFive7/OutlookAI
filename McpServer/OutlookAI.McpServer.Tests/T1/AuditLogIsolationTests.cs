using System.Globalization;
using System.Reflection;
using OutlookAI.Core.Audit;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Q86 (decided 2026-10-03): proves that this test process writes its audit lines to a throwaway
/// directory of its own, and cannot write them to the real log - WITHOUT opening, reading or
/// diffing the real file.
/// <para>
/// <b>Why not just watch the real file.</b> The maintainer's own OutlookAI server may append to
/// <c>%LOCALAPPDATA%\OutlookAI\audit.log</c> at any moment, including during a test run, so "the
/// real file changed" proves nothing about the tests and would fail at random. Everything here is
/// therefore established inside this process: where it appends, that an append really lands
/// there, that anywhere else is refused before anything is created, and - read out of the
/// compiled IL - that no test names the real log at all except the few that are allowed to and
/// say why.
/// </para>
/// <para>
/// <b>Safe even when the isolation is broken.</b> Every test that writes asserts the redirect
/// FIRST, so on code that does not isolate it stops before writing; and the tripwire tests aim
/// at a directory under the build output, never at the real one. That is what made the control
/// run possible: with the module initializer removed, these tests fail without the real log
/// being touched.
/// </para>
/// </summary>
public sealed class AuditLogIsolationTests
{
    /// <summary>
    /// The test classes allowed to name the real audit log in compiled code, each with the reason
    /// it may. Everything else that does fails <see cref="OnlyTheAllowlist_NamesTheRealAuditLog_InCompiledTestCode"/>.
    /// </summary>
    private static readonly (string TypeName, string Reason)[] MayNameTheRealLog =
    {
        ("OutlookAI.McpServer.Tests.T1.AuditLogTests",
            "pins the product's default path against %LOCALAPPDATA% - reads the value, opens nothing"),
        ("OutlookAI.McpServer.Tests.T1.AuditLogIsolationTests",
            "this guard - compares the effective directory against the real one, opens nothing under it"),
        ("OutlookAI.McpServer.Tests.T3.Phase4LiveMcpToolShapeTests", LiveChildReader),
        ("OutlookAI.McpServer.Tests.T3.Phase5LiveMcpToolShapeTests", LiveChildReader),
        ("OutlookAI.McpServer.Tests.T3.LiveManageSignatureMcpToolTests", LiveChildReader),
    };

    private const string LiveChildReader =
        "Category=Live stdio class: reads the lines the SERVER CHILD process wrote on a test machine - a "
        + "process this redirect cannot reach, because nothing may redirect a shipped process";

    private const string StdioNamespace = "OutlookAI.McpServer.Tests.T3";

    private readonly ITestOutputHelper _output;

    public AuditLogIsolationTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public void ThisProcess_AuditsToItsOwnThrowawayDirectory_NotTheRealOne()
    {
        Assert.True(
            AuditLog.IsRedirected,
            "this test process was never redirected, so every write path the suite drives appends to the real "
            + "audit log - the module initializer in AuditIsolation did not run");

        string effective = AuditLog.EffectiveDirectory;
        Assert.True(
            SamePath(AuditIsolation.ThrowawayDirectory, effective),
            "the process appends to '" + effective + "', not to the directory its module initializer chose");
        Assert.False(
            IsSameOrUnder(effective, AuditLog.DefaultDirectory),
            "the redirect points at the real audit directory or inside it: '" + effective + "'");
        Assert.True(
            IsSameOrUnder(effective, Path.GetTempPath()),
            "the redirect is not under the temp directory, so it is not throwaway: '" + effective + "'");

        // Process-local: two test processes - two worktrees, two agents - never share a log.
        Assert.StartsWith(
            Environment.ProcessId.ToString(CultureInfo.InvariantCulture) + "-",
            Path.GetFileName(effective),
            StringComparison.Ordinal);

        Assert.Equal(Path.Combine(effective, AuditLog.LogFileName), AuditLog.EffectiveLogPath);
        Assert.False(
            SamePath(AuditLog.EffectiveLogPath, AuditLog.DefaultLogPath),
            "the effective log IS the real log");
    }

    [Fact]
    public void ASentinelAppendedThroughTheProductPath_LandsInTheThrowawayLog()
    {
        // FIRST, before anything is written: on code that does not isolate, this is where the
        // test stops - which is what lets the control run fail it without touching the real log.
        Assert.True(AuditLog.IsRedirected, "not redirected - refusing to append anything");
        Assert.False(
            IsSameOrUnder(AuditLog.EffectiveDirectory, AuditLog.DefaultDirectory),
            "redirected at the real directory - refusing to append anything");

        string sentinel = "q86-sentinel-" + Guid.NewGuid().ToString("N");

        // The exact call every product write path makes.
        AuditLog.Append("isolation_sentinel", ("sentinel", sentinel));

        string throwawayLog = Path.Combine(AuditIsolation.ThrowawayDirectory, AuditLog.LogFileName);
        string[] lines = ReadShared(throwawayLog);
        Assert.Single(lines, line =>
            line.Contains(" op=isolation_sentinel ", StringComparison.Ordinal)
            && line.Contains("sentinel=\"" + sentinel + "\"", StringComparison.Ordinal));
    }

    [Fact]
    public void WhileRedirected_AnAuditDirectoryOutsideTemp_IsRefusedBeforeAnythingIsCreated()
    {
        // Under the build output: outside temp, and certainly not the real directory. If the
        // tripwire is broken this test writes HERE, which is harmless - that is the point of
        // aiming it somewhere other than the real log.
        string outside = Path.Combine(AppContext.BaseDirectory, "audit-isolation-tripwire-" + Guid.NewGuid().ToString("N"));
        Assert.False(IsSameOrUnder(outside, Path.GetTempPath()), "precondition: the target has to be outside temp");
        Assert.False(IsSameOrUnder(outside, AuditLog.DefaultDirectory), "precondition: never aim this at the real directory");

        try
        {
            InvalidOperationException refused = Assert.Throws<InvalidOperationException>(
                () => AuditLog.AppendTo(outside, "isolation_tripwire", Array.Empty<(string, string?)>()));
            Assert.Contains("test process", refused.Message, StringComparison.Ordinal);
            Assert.False(Directory.Exists(outside), "the append was refused only AFTER it had created the directory");

            Assert.False(AuditLog.TryProbeWritable(outside, out string? error), "the health probe accepted a directory outside temp");
            Assert.Equal(AuditLog.RedirectRefusalError, error);
            Assert.False(Directory.Exists(outside), "the probe was refused only AFTER it had created the directory");
        }
        finally
        {
            if (Directory.Exists(outside))
            {
                Directory.Delete(outside, recursive: true);
            }
        }
    }

    [Theory]
    [InlineData("as written")]
    [InlineData("trailing separator")]
    [InlineData("upper case")]
    [InlineData("forward slashes")]
    [InlineData("device prefix")]
    [InlineData("dot-dot detour")]
    [InlineData("a folder inside it")]
    public void WhileRedirected_TheRealDirectory_IsRefusedInEverySpelling(string spelling)
    {
        // Pure: these two only resolve and compare paths, so the real directory is named here
        // without anything being created or opened under it.
        string real = AuditLog.DefaultDirectory;
        string candidate = spelling switch
        {
            "as written" => real,
            "trailing separator" => real + Path.DirectorySeparatorChar,
            "upper case" => real.ToUpperInvariant(),
            "forward slashes" => real.Replace(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
            "device prefix" => @"\\?\" + real,
            "dot-dot detour" => Path.Combine(real, "..", Path.GetFileName(real)),
            "a folder inside it" => Path.Combine(real, "sub"),
            _ => throw new ArgumentOutOfRangeException(nameof(spelling)),
        };

        Assert.NotNull(AuditLog.DescribeRefusalWhileRedirected(candidate));
        Assert.NotNull(AuditLog.DescribeRedirectRefusal(candidate));
    }

    [Fact]
    public void WhileRedirected_AnyDirectoryUnderTemp_IsAllowed()
    {
        // The other half of the allowlist, or the tripwire would refuse the suite's own temp
        // directories (AuditLogTests appends to one) along with the real log.
        Assert.Null(AuditLog.DescribeRefusalWhileRedirected(AuditIsolation.ThrowawayDirectory));
        Assert.Null(AuditLog.DescribeRefusalWhileRedirected(Path.Combine(Path.GetTempPath(), "OutlookAI-AuditLogTests-x")));
    }

    [Fact]
    public void ARedirectTarget_MustBeAbsolute_Throwaway_AndNotTheRealDirectory()
    {
        Assert.NotNull(AuditLog.DescribeRedirectRefusal(string.Empty));
        Assert.NotNull(AuditLog.DescribeRedirectRefusal("   "));
        Assert.NotNull(AuditLog.DescribeRedirectRefusal(Path.Combine("relative", "dir")));
        Assert.NotNull(AuditLog.DescribeRedirectRefusal(AppContext.BaseDirectory));
        Assert.NotNull(AuditLog.DescribeRedirectRefusal(AuditLog.DefaultDirectory));

        Assert.Null(AuditLog.DescribeRedirectRefusal(
            Path.Combine(Path.GetTempPath(), AuditIsolation.ParentFolderName, "any-run")));
    }

    [Fact]
    public void TheRedirect_IsSetOnce_AndAnotherTargetIsRefusedWithoutMovingIt()
    {
        string before = AuditLog.EffectiveDirectory;

        // The same directory again is a no-op.
        AuditLog.RedirectThisProcess(AuditIsolation.ThrowawayDirectory);
        Assert.Equal(before, AuditLog.EffectiveDirectory);

        // A second throwaway directory would split this run's lines across two logs.
        string other = Path.Combine(Path.GetTempPath(), AuditIsolation.ParentFolderName, "other-" + Guid.NewGuid().ToString("N"));
        Assert.Throws<InvalidOperationException>(() => AuditLog.RedirectThisProcess(other));
        Assert.Equal(before, AuditLog.EffectiveDirectory);

        // And the real directory is refused before anything else is looked at.
        Assert.Throws<ArgumentException>(() => AuditLog.RedirectThisProcess(AuditLog.DefaultDirectory));
        Assert.Equal(before, AuditLog.EffectiveDirectory);
    }

    [Fact]
    public void TheHealthReport_ProbesAndNamesTheLogThisProcessWritesTo()
    {
        // Asserted first so that, on code that does not isolate, the probe below never runs: it
        // opens the file it reports on.
        Assert.True(AuditLog.IsRedirected, "not redirected - refusing to probe anything");

        AuditHealthView audit = MailService.DescribeAuditLog(out string? problem);

        Assert.Equal(AuditLog.EffectiveLogPath, audit.Path);
        Assert.False(SamePath(audit.Path, AuditLog.DefaultLogPath), "outlook_health names the real log in a test process");
        Assert.True(audit.Writable, "the throwaway log should be writable; error: " + audit.Error);
        Assert.Null(audit.Error);
        Assert.Null(problem);
    }

    [Fact]
    public void OnlyTheAllowlist_NamesTheRealAuditLog_InCompiledTestCode()
    {
        // A ratchet over the compiled IL of every type in this assembly - nested and
        // compiler-generated ones included, which is where async and lambda bodies live. What it
        // looks for is any way of NAMING the real log: the two AuditLog properties, and the
        // %LOCALAPPDATA% lookup a test would use to build the path by hand. A new mention fails
        // until it is added to MayNameTheRealLog with its reason; an entry that no longer
        // mentions it fails too, so the list can only shrink on purpose.
        //
        // The runtime tripwire above already refuses every in-process WRITE outside temp, however
        // the path was built; this is the second layer, and the only one that also covers reads
        // and paths handed to another process.
        MethodBase defaultDirectory = typeof(AuditLog).GetProperty(nameof(AuditLog.DefaultDirectory))!.GetMethod!;
        MethodBase defaultLogPath = typeof(AuditLog).GetProperty(nameof(AuditLog.DefaultLogPath))!.GetMethod!;
        MethodBase getFolderPath = typeof(Environment).GetMethod(
            nameof(Environment.GetFolderPath), new[] { typeof(Environment.SpecialFolder) })!;

        Dictionary<string, SortedSet<string>> mentions = new(StringComparer.Ordinal);
        int methodsScanned = 0;
        Assembly tests = typeof(AuditLogIsolationTests).Assembly;
        foreach (Type type in tests.GetTypes())
        {
            foreach (MethodBase method in MethodsOf(type))
            {
                byte[]? il = SafeIl(method);
                if (il == null)
                {
                    continue;
                }

                methodsScanned++;
                foreach ((int offset, MethodBase target) in CallTargets(type, method, il))
                {
                    string? what = null;
                    if (SameMethod(target, defaultDirectory))
                    {
                        what = "AuditLog.DefaultDirectory";
                    }
                    else if (SameMethod(target, defaultLogPath))
                    {
                        what = "AuditLog.DefaultLogPath";
                    }
                    else if (SameMethod(target, getFolderPath) && PushesLocalApplicationData(il, offset))
                    {
                        what = "Environment.GetFolderPath(LocalApplicationData)";
                    }

                    if (what != null)
                    {
                        string owner = OutermostType(type).FullName!;
                        if (!mentions.TryGetValue(owner, out SortedSet<string>? set))
                        {
                            mentions[owner] = set = new SortedSet<string>(StringComparer.Ordinal);
                        }

                        set.Add(what);
                    }
                }
            }
        }

        foreach (KeyValuePair<string, SortedSet<string>> mention in mentions.OrderBy(m => m.Key, StringComparer.Ordinal))
        {
            _output.WriteLine(mention.Key + ": " + string.Join(", ", mention.Value));
        }

        HashSet<string> allowed = new(MayNameTheRealLog.Select(a => a.TypeName), StringComparer.Ordinal);
        List<string> problems = new();
        foreach (string owner in mentions.Keys.Where(o => !allowed.Contains(o)).OrderBy(o => o, StringComparer.Ordinal))
        {
            problems.Add(owner + " names the real audit log (" + string.Join(", ", mentions[owner]) + "). In a test "
                + "process use AuditLog.EffectiveLogPath / EffectiveDirectory, which is where this process's lines "
                + "really go; if it genuinely needs the real path, add it to MayNameTheRealLog with the reason.");
        }

        foreach ((string typeName, string reason) in MayNameTheRealLog)
        {
            if (!mentions.ContainsKey(typeName))
            {
                problems.Add(typeName + " is allowlisted (" + reason + ") but no longer names the real audit log - "
                    + "remove it from MayNameTheRealLog so the list cannot hide a later mention.");
            }

            if (ReferenceEquals(reason, LiveChildReader))
            {
                // The reason has to stay TRUE, not merely be written down: only an out-of-process
                // reader in the live tier has one.
                Type? type = tests.GetType(typeName, throwOnError: false);
                if (type == null || type.Namespace != StdioNamespace || !IsLiveClass(type))
                {
                    problems.Add(typeName + " is allowlisted as a live stdio reader but is not a Category=Live class in "
                        + StdioNamespace + ".");
                }
            }
        }

        Assert.Empty(problems);

        // The detector must not be able to switch itself off: a scan that resolves nothing reads
        // exactly like a clean assembly. AuditLogTests' default-path pin names all three, so all
        // three must be seen there.
        Assert.True(methodsScanned > 1000, "only " + methodsScanned + " method bodies were scanned");
        Assert.True(
            mentions.TryGetValue("OutlookAI.McpServer.Tests.T1.AuditLogTests", out SortedSet<string>? pin)
                && pin.Count == 3,
            "the IL walk no longer sees AuditLogTests' pin of the default path - it has stopped resolving call targets "
            + "and this ratchet is green by accident");
    }

    // ------------------------------------------------------------------ helpers

    private static IEnumerable<MethodBase> MethodsOf(Type type)
    {
        const BindingFlags all = BindingFlags.Public | BindingFlags.NonPublic
            | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly;
        return type.GetMethods(all).Cast<MethodBase>().Concat(type.GetConstructors(all));
    }

    /// <summary>
    /// Every method a body calls, loads as a delegate or calls virtually, with the offset of the
    /// opcode. A byte-level walk rather than a decoder, like LiveTierInventoryTests' ldstr walk:
    /// an operand that merely looks like a call either fails to resolve or resolves to something
    /// that is none of the three targets, so it cannot produce a false finding.
    /// </summary>
    private static IEnumerable<(int Offset, MethodBase Target)> CallTargets(Type type, MethodBase method, byte[] il)
    {
        Type[]? typeArguments = type.IsGenericType ? type.GetGenericArguments() : null;
        Type[]? methodArguments = method.IsGenericMethod ? method.GetGenericArguments() : null;

        for (int i = 0; i + 4 < il.Length; i++)
        {
            int operand;
            if (il[i] == 0x28 || il[i] == 0x6F)
            {
                operand = i + 1;
            }
            else if (il[i] == 0xFE && (il[i + 1] == 0x06 || il[i + 1] == 0x07) && i + 5 < il.Length)
            {
                operand = i + 2;
            }
            else
            {
                continue;
            }

            int token = il[operand] | (il[operand + 1] << 8) | (il[operand + 2] << 16) | (il[operand + 3] << 24);
            byte table = il[operand + 3];
            if (table != 0x06 && table != 0x0A && table != 0x2B)
            {
                continue;
            }

            MethodBase? target;
            try
            {
                target = type.Module.ResolveMethod(token, typeArguments, methodArguments);
            }
            catch (Exception ex) when (ex is ArgumentException || ex is BadImageFormatException
                || ex is TypeLoadException || ex is MissingMethodException)
            {
                continue;
            }

            if (target != null)
            {
                yield return (i, target);
            }
        }
    }

    /// <summary><c>ldc.i4.s 28</c> - <c>SpecialFolder.LocalApplicationData</c> - immediately before the call.</summary>
    private static bool PushesLocalApplicationData(byte[] il, int callOffset)
    {
        return callOffset >= 2
            && il[callOffset - 2] == 0x1F
            && il[callOffset - 1] == (byte)Environment.SpecialFolder.LocalApplicationData;
    }

    private static bool SameMethod(MethodBase candidate, MethodBase target)
    {
        return candidate.MetadataToken == target.MetadataToken && candidate.Module == target.Module;
    }

    private static Type OutermostType(Type type)
    {
        while (type.DeclaringType != null)
        {
            type = type.DeclaringType;
        }

        return type;
    }

    private static bool IsLiveClass(Type type)
    {
        return type.GetCustomAttributesData().Any(a =>
            a.AttributeType == typeof(TraitAttribute)
            && a.ConstructorArguments.Count == 2
            && string.Equals(a.ConstructorArguments[0].Value as string, "Category", StringComparison.Ordinal)
            && string.Equals(a.ConstructorArguments[1].Value as string, "Live", StringComparison.Ordinal));
    }

    private static byte[]? SafeIl(MethodBase method)
    {
        try
        {
            return method.GetMethodBody()?.GetILAsByteArray();
        }
        catch (Exception ex) when (ex is InvalidOperationException || ex is NotSupportedException)
        {
            return null;
        }
    }

    /// <summary>Reads a log the way the product's appends allow: shared, so a reader never blocks a writer.</summary>
    private static string[] ReadShared(string path)
    {
        using FileStream stream = new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using StreamReader reader = new(stream);
        return reader.ReadToEnd().Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
    }

    // Deliberately NOT AuditLog's own comparison: a guard that borrows the code under test agrees
    // with that code's mistakes.
    private static string Full(string path)
    {
        return Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar);
    }

    private static bool SamePath(string a, string b)
    {
        return string.Equals(Full(a), Full(b), StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsSameOrUnder(string path, string root)
    {
        string candidate = Full(path);
        string container = Full(root);
        return string.Equals(candidate, container, StringComparison.OrdinalIgnoreCase)
            || candidate.StartsWith(container + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
    }
}
