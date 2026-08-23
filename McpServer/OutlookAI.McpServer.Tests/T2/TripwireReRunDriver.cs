using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Text;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// What the child run's OWN store-count tripwire decided, handed back to the parent through a
/// file rather than inferred from an exit code.
/// <para>
/// The exit code alone cannot carry this. A child that exits non-zero may have reproduced the
/// delta, or it may have failed an assertion about paging, or it may have refused to start
/// because Outlook was wedged - and the sentence the parent prints for the first of those is
/// "THIS IS THE SUITE REMOVING MAIL. Stop and investigate.", which is the loudest claim in this
/// repository and must not be made about the other two.
/// </para>
/// </summary>
public enum TripwireReRunMarker
{
    /// <summary>
    /// The child never got as far as verifying itself: it could not start, it was killed, or
    /// its own baseline census refused the tier. Deliberately the ZERO value - a child that
    /// wrote nothing has answered nothing.
    /// </summary>
    Absent = 0,

    /// <summary>The child ran its own before/after census and found nothing to report.</summary>
    Clean = 1,

    /// <summary>The child's own tripwire fired. The delta came back.</summary>
    TripwireFailed = 2,
}

/// <summary>
/// The exact command that re-runs the implicated collections IN ANOTHER PROCESS, worked out
/// without starting anything - so CI can pin what the driver would do on a machine that has no
/// Outlook, no mailbox and no live tier.
/// </summary>
public sealed class TripwireReRunPlan
{
    internal TripwireReRunPlan(string fileName, IReadOnlyList<string> arguments, string markerPath)
    {
        FileName = fileName;
        Arguments = arguments;
        MarkerPath = markerPath;
    }

    /// <summary>The dotnet host that will run the child. Never this process.</summary>
    public string FileName { get; }

    /// <summary>The child's argument list, one element per argument (never a joined string).</summary>
    public IReadOnlyList<string> Arguments { get; }

    /// <summary>Where the child will write <see cref="TripwireReRunMarker"/>.</summary>
    public string MarkerPath { get; }

    /// <summary>One line naming the command, for the run's own output.</summary>
    public string Describe()
    {
        return FileName + " " + string.Join(" ", Arguments.Select(Quote));
    }

    private static string Quote(string argument)
    {
        return argument.Contains(' ', StringComparison.Ordinal) ? "\"" + argument + "\"" : argument;
    }
}

/// <summary>
/// The third rung of <see cref="TripwireRetryLadder"/>, made real: a bounded re-run of the
/// plausibly-implicated collections IN A SEPARATE PROCESS.
/// <para>
/// <b>Out of process is what makes the rung safe at all, and it is the whole design.</b>
/// Starting a second xunit run inside the first one's teardown would re-enter fixtures that are
/// currently disposing, re-take a baseline over a profile mid-teardown, and write to a mailbox
/// at the one moment when nothing is left to sweep the artifacts away. A child process shares
/// none of that: it is an ordinary, fully guarded live run - its own preflight, its own
/// baseline census, its own write allowlist, its own zero-artifact sweep, its own verification -
/// which is exactly the run a maintainer would perform by hand, and the run this rung used to
/// print instructions for instead of performing.
/// </para>
/// <para>
/// <b>The child's own tripwire is the oracle.</b> The parent does not compare censuses across
/// processes and does not try to; it asks the child what its own guard decided. That is why the
/// answer arrives as a <see cref="TripwireReRunMarker"/> file rather than as an exit code: an
/// exit code cannot tell "the delta came back" from "an unrelated test failed", and only the
/// first of those may be reported as the suite removing mail.
/// </para>
/// <para>
/// <b>Recursion is impossible by construction.</b> The child is started with
/// <see cref="MarkerVariable"/> set, and a process that sees that variable takes
/// <see cref="TripwireRetryPolicy.WithoutReRuns"/> - so the third rung exists in exactly one
/// process, the one at the top.
/// </para>
/// <para>
/// <b>Every way this can fail is <see cref="TripwireReRunOutcome.Inconclusive"/>, which FAILS
/// the run.</b> No dotnet host, no project file, a process that will not start, a budget that
/// expires, a marker that is not there: an experiment nobody could carry out exonerates
/// nothing, and the ladder is written so that only <see cref="TripwireReRunOutcome.NotReproduced"/>
/// can clear anything - and even that reports
/// <see cref="TripwireRunOutcome.PassedWithASurvivedDelta"/>, which still exits non-zero.
/// </para>
/// </summary>
public static class TripwireReRunDriver
{
    /// <summary>
    /// Set on the CHILD, and on nothing else. Its presence says two things at once, which is
    /// why it is one variable rather than two: this process is a bounded re-run (so it may not
    /// start one of its own), and this is where it writes what its tripwire decided. Two
    /// variables could disagree; one cannot.
    /// </summary>
    public const string MarkerVariable = "OUTLOOKAI_TRIPWIRE_RERUN_MARKER";

    /// <summary>
    /// How long the parent waits for the child before giving up on it.
    /// <para>
    /// DERIVED, and generous on purpose. The re-run covers the guarded collections this run
    /// executed, which on a whole-tier run is the whole tier; the only measured tier run is
    /// 26.8 minutes (2026-08-18), and the child additionally pays its own baseline and post-run
    /// censuses (16.9 s each) and its own re-census bounds. Sixty minutes is a little over
    /// twice the measured run - a CEILING, not a target, and the risk is asymmetric: too low
    /// turns a working experiment into a failure, too high costs nothing at all when the child
    /// finishes early.
    /// </para>
    /// <para>
    /// <b>Expiry does not kill the child.</b> Killing a live run mid-flight leaves tagged
    /// artifacts in a real mailbox with nothing left to sweep them - the very hazard that kept
    /// this rung unimplemented. An abandoned child finishes its own teardown and sweeps itself;
    /// the parent simply stops waiting, reports Inconclusive (which fails), and prints the PID
    /// and the log so the maintainer can watch it out.
    /// </para>
    /// </summary>
    public const int ReRunBudgetMinutes = 60;

    /// <summary>The budget as a <see cref="TimeSpan"/>.</summary>
    public static TimeSpan ReRunBudget => TimeSpan.FromMinutes(ReRunBudgetMinutes);

    /// <summary>True when THIS process is the bounded re-run started by another one.</summary>
    public static bool IsReRunChild => MarkerPathForThisProcess() != null;

    /// <summary>Where this process must record its own tripwire verdict, or null when it is not a child.</summary>
    public static string? MarkerPathForThisProcess()
    {
        string? path = Environment.GetEnvironmentVariable(MarkerVariable);
        return string.IsNullOrWhiteSpace(path) ? null : path;
    }

    /// <summary>
    /// Records what this process's own tripwire decided, for the parent that started it. A
    /// no-op in an ordinary run. Never throws: a marker that cannot be written leaves the
    /// parent reading <see cref="TripwireReRunMarker.Absent"/>, which fails, so the failure
    /// direction is already the safe one and a throw here would only replace a clear verdict
    /// with a confusing teardown error.
    /// </summary>
    public static void RecordOwnVerdict(bool tripwireFailed)
    {
        string? path = MarkerPathForThisProcess();
        if (path == null)
        {
            return;
        }

        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path)!);
            File.WriteAllText(
                path,
                (tripwireFailed ? TripwireReRunMarker.TripwireFailed : TripwireReRunMarker.Clean).ToString(),
                new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or ArgumentException)
        {
            Console.WriteLine(
                "[tripwire] could not write the re-run marker to '" + path + "' (" + ex.GetType().Name
                + "). The parent will read this as no answer, which fails.");
        }
    }

    /// <summary>
    /// Reads a marker file. Anything that is not one of the two written spellings reads as
    /// <see cref="TripwireReRunMarker.Absent"/> - a half-written or corrupt marker is not an
    /// answer, and the fail-closed reading is the one that does not exonerate.
    /// </summary>
    public static TripwireReRunMarker ReadMarker(string? text)
    {
        string trimmed = (text ?? string.Empty).Trim();
        if (string.Equals(trimmed, nameof(TripwireReRunMarker.TripwireFailed), StringComparison.Ordinal))
        {
            return TripwireReRunMarker.TripwireFailed;
        }

        return string.Equals(trimmed, nameof(TripwireReRunMarker.Clean), StringComparison.Ordinal)
            ? TripwireReRunMarker.Clean
            : TripwireReRunMarker.Absent;
    }

    /// <summary>
    /// The whole verdict rule, pure, so CI drives every combination without starting a process.
    /// </summary>
    /// <param name="exitCode">
    /// The child's exit code, or null when it never started or never finished inside the budget.
    /// </param>
    /// <param name="marker">What the child's own tripwire recorded.</param>
    public static TripwireReRunOutcome Classify(int? exitCode, TripwireReRunMarker marker)
    {
        if (exitCode == null)
        {
            // Never started, or still running past the budget. No experiment, no answer.
            return TripwireReRunOutcome.Inconclusive;
        }

        if (marker == TripwireReRunMarker.TripwireFailed)
        {
            // The child's own before/after census fired. That is the delta coming back, and it
            // is the ONLY reading that earns the sentence about the suite removing mail.
            return TripwireReRunOutcome.Reproduced;
        }

        if (marker == TripwireReRunMarker.Absent)
        {
            // The child ran and never reached its own verification - a refused preflight, a
            // crash, a collection that threw before the census. It proves nothing either way.
            return TripwireReRunOutcome.Inconclusive;
        }

        // The child's census was clean. A zero exit means the whole child run agreed; a
        // non-zero one means something ELSE in it failed, which is not an exoneration - the
        // experiment did not complete under the conditions it was supposed to.
        return exitCode == 0 ? TripwireReRunOutcome.NotReproduced : TripwireReRunOutcome.Inconclusive;
    }

    /// <summary>
    /// Every environment variable the child must be started with.
    /// <para>
    /// Pure, and separate from the launch, because what is IN it is the recursion guard: the
    /// child recognises itself by <see cref="MarkerVariable"/> being present and therefore
    /// refuses to start a re-run of its own. Spelling that key wrong - or writing it to a
    /// different name - silently re-arms recursion AND leaves the parent with no marker to
    /// read, and the launch itself is behind a process start no CI test can perform. So the
    /// content is decided here, where it is pinned by value, and the launch only applies it.
    /// </para>
    /// </summary>
    public static IReadOnlyDictionary<string, string> ChildEnvironment(TripwireReRunPlan plan)
    {
        ArgumentNullException.ThrowIfNull(plan);
        return new Dictionary<string, string>(StringComparer.Ordinal)
        {
            [MarkerVariable] = plan.MarkerPath,
        };
    }

    /// <summary>
    /// The VSTest filter selecting exactly the given test classes, or null when there are none.
    /// <para>
    /// Classes rather than collections, because a collection is an xunit concept the VSTest
    /// filter language cannot express. The translation is by reflection over this same
    /// assembly, so it cannot drift from the <c>[Collection]</c> attributes it reads.
    /// </para>
    /// </summary>
    public static string? FilterFor(IReadOnlyList<string> testClassFullNames)
    {
        ArgumentNullException.ThrowIfNull(testClassFullNames);
        List<string> clauses = testClassFullNames
            .Where(name => !string.IsNullOrWhiteSpace(name))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(name => name, StringComparer.Ordinal)
            .Select(name => "FullyQualifiedName~" + name + ".")
            .ToList();
        return clauses.Count == 0 ? null : string.Join("|", clauses);
    }

    /// <summary>
    /// Every test class in <paramref name="assembly"/> belonging to one of
    /// <paramref name="collectionNames"/>, by its <c>[Collection]</c> attribute.
    /// </summary>
    public static IReadOnlyList<string> ClassesIn(Assembly assembly, IReadOnlyList<string> collectionNames)
    {
        ArgumentNullException.ThrowIfNull(assembly);
        ArgumentNullException.ThrowIfNull(collectionNames);
        HashSet<string> wanted = new(collectionNames.Where(n => n != null), StringComparer.Ordinal);
        return assembly.GetTypes()
            .Where(t => t.IsClass && !t.IsAbstract && t.FullName != null)
            .Where(t => wanted.Contains(CollectionOf(t) ?? string.Empty))
            .Select(t => t.FullName!)
            .OrderBy(name => name, StringComparer.Ordinal)
            .ToList();
    }

    /// <summary>
    /// Builds the command, or throws naming what is missing. Split from
    /// <see cref="Run"/> so that everything decidable without a process is decidable in CI.
    /// </summary>
    /// <param name="dotnetHostPath">
    /// The dotnet muxer, normally from <c>DOTNET_HOST_PATH</c> (the SDK sets it for anything it
    /// launches). Null or blank falls back to <c>dotnet</c> on PATH.
    /// </param>
    /// <param name="projectPath">The test project to re-run.</param>
    /// <param name="filter">The VSTest filter selecting the implicated classes.</param>
    /// <param name="configuration">Build configuration the parent was built in.</param>
    /// <param name="markerPath">Where the child writes its verdict.</param>
    public static TripwireReRunPlan PlanFor(
        string? dotnetHostPath, string projectPath, string filter, string configuration, string markerPath)
    {
        if (string.IsNullOrWhiteSpace(projectPath))
        {
            throw new ArgumentException("The bounded re-run needs a test project to run.", nameof(projectPath));
        }

        if (string.IsNullOrWhiteSpace(filter))
        {
            throw new ArgumentException(
                "The bounded re-run needs a filter. Re-running the WHOLE assembly would schedule live "
                + "collections this run never executed, which is a different experiment from the one the "
                + "ladder asked for.",
                nameof(filter));
        }

        if (string.IsNullOrWhiteSpace(markerPath))
        {
            throw new ArgumentException(
                "The bounded re-run needs a marker path - without it the child has no way to say what its own "
                + "tripwire decided, and every outcome would read as Inconclusive.",
                nameof(markerPath));
        }

        // --no-build on purpose: the parent's own assembly is what is being re-run, a rebuild
        // in the middle of a teardown is a second way for the experiment to fail, and rebuilding
        // could quietly test different code from the code that produced the delta.
        List<string> arguments = new()
        {
            "test",
            projectPath,
            "--no-build",
            "--filter",
            filter,
        };

        if (!string.IsNullOrWhiteSpace(configuration))
        {
            arguments.Add("-c");
            arguments.Add(configuration);
        }

        return new TripwireReRunPlan(
            string.IsNullOrWhiteSpace(dotnetHostPath) ? "dotnet" : dotnetHostPath!,
            arguments,
            markerPath);
    }

    /// <summary>
    /// Performs the re-run: builds the plan from this machine, starts the child, waits for it
    /// inside <see cref="ReRunBudget"/>, and classifies what came back.
    /// <para>
    /// The only method here that touches a process. Every decision it makes is delegated to a
    /// pure one above, so what is left is the launch itself.
    /// </para>
    /// </summary>
    public static TripwireReRunOutcome Run(IReadOnlyList<string> implicatedCollections, int attempt)
    {
        ArgumentNullException.ThrowIfNull(implicatedCollections);
        string scratch = Path.Combine(
            Path.GetTempPath(),
            "outlookai-tripwire-rerun-" + Environment.ProcessId.ToString(CultureInfo.InvariantCulture)
                + "-" + attempt.ToString(CultureInfo.InvariantCulture));
        TripwireReRunPlan plan;
        try
        {
            plan = PlanFor(
                Environment.GetEnvironmentVariable("DOTNET_HOST_PATH"),
                LocateTestProject(),
                FilterFor(ClassesIn(typeof(TripwireReRunDriver).Assembly, implicatedCollections))
                    ?? throw new InvalidOperationException(
                        "no test class belongs to the implicated collection(s): "
                        + string.Join(", ", implicatedCollections)),
                BuildConfiguration(),
                Path.Combine(scratch, "verdict.txt"));
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            Console.WriteLine(
                "[tripwire] the bounded re-run could not be prepared (" + ex.GetType().Name + ": " + ex.Message
                + "). An experiment nobody could carry out exonerates nothing, so this FAILS the run.");
            return TripwireReRunOutcome.Inconclusive;
        }

        string logPath = Path.Combine(scratch, "rerun.log");
        Console.WriteLine(
            "[tripwire] bounded re-run " + attempt.ToString(CultureInfo.InvariantCulture)
            + ", OUT OF PROCESS so it cannot re-enter the fixtures disposing right now: " + plan.Describe());
        Console.WriteLine("[tripwire] re-run log: " + logPath + "; budget " + ReRunBudgetMinutes + " min.");

        int? exitCode = Launch(plan, logPath);
        TripwireReRunMarker marker = TripwireReRunMarker.Absent;
        try
        {
            if (File.Exists(plan.MarkerPath))
            {
                marker = ReadMarker(File.ReadAllText(plan.MarkerPath));
            }
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Console.WriteLine("[tripwire] the re-run marker could not be read (" + ex.GetType().Name + ").");
        }

        Console.WriteLine(
            "[tripwire] re-run child exit=" + (exitCode?.ToString(CultureInfo.InvariantCulture) ?? "none")
            + ", its own tripwire=" + marker + ".");
        return Classify(exitCode, marker);
    }

    /// <summary>
    /// Starts the child and waits for it. Returns its exit code, or null when it could not be
    /// started or outlived the budget.
    /// </summary>
    private static int? Launch(TripwireReRunPlan plan, string logPath)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(logPath)!);
        ProcessStartInfo psi = new(plan.FileName)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,

            // Never a window and never the foreground: this runs unattended on a machine
            // somebody may be using, and a console appearing mid-teardown steals focus.
            CreateNoWindow = true,
        };
        foreach (string argument in plan.Arguments)
        {
            psi.ArgumentList.Add(argument);
        }

        foreach (KeyValuePair<string, string> entry in ChildEnvironment(plan))
        {
            psi.Environment[entry.Key] = entry.Value;
        }

        Process? child = null;
        try
        {
            using StreamWriter log = new(logPath, append: false, new UTF8Encoding(false));
            log.AutoFlush = true;
            child = Process.Start(psi);
            if (child == null)
            {
                Console.WriteLine("[tripwire] the re-run child did not start.");
                return null;
            }

            // Pumped through events rather than read to the end here: a child that fills a
            // redirected pipe while the parent blocks in WaitForExit deadlocks, which would
            // wedge the teardown with no timeout able to fire.
            object gate = new();
            child.OutputDataReceived += (_, e) => Append(log, gate, e.Data);
            child.ErrorDataReceived += (_, e) => Append(log, gate, e.Data);
            child.BeginOutputReadLine();
            child.BeginErrorReadLine();

            if (child.WaitForExit((int)ReRunBudget.TotalMilliseconds))
            {
                // The overload with no arguments, AFTER the timed one returned true: it is what
                // waits for the redirected streams to be drained, so the log is complete.
                child.WaitForExit();
                return child.ExitCode;
            }

            // NOT killed. See ReRunBudgetMinutes: an abandoned live run sweeps its own
            // artifacts, a killed one leaves them in a real mailbox.
            Console.WriteLine(
                "[tripwire] the re-run child (pid " + child.Id.ToString(CultureInfo.InvariantCulture)
                + ") is still running after " + ReRunBudgetMinutes
                + " min. It is deliberately NOT killed - killing a live run mid-flight leaves tagged "
                + "artifacts with nothing to sweep them. Watch it out at " + logPath
                + "; this attempt counts as no answer, which FAILS the run.");
            return null;
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            Console.WriteLine(
                "[tripwire] the re-run child could not be run (" + ex.GetType().Name + ": " + ex.Message + ").");
            return null;
        }
        finally
        {
            child?.Dispose();
        }
    }

    private static void Append(StreamWriter log, object gate, string? line)
    {
        if (line == null)
        {
            return;
        }

        lock (gate)
        {
            log.WriteLine(line);
        }
    }

    /// <summary>
    /// The test project to re-run, found beside this assembly's source rather than named. It
    /// must be exactly one: a directory with two project files cannot say which one produced
    /// this assembly, and guessing is how the wrong tier gets re-run.
    /// </summary>
    private static string LocateTestProject()
    {
        string directory = typeof(TripwireReRunDriver).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");

        string[] projects = Directory.GetFiles(directory, "*.csproj");
        return projects.Length == 1
            ? projects[0]
            : throw new InvalidOperationException(
                "expected exactly one .csproj in '" + directory + "', found "
                + projects.Length.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// The configuration this assembly was built in, so the child re-runs the same binaries
    /// under <c>--no-build</c> rather than whichever configuration is the default.
    /// </summary>
    private static string BuildConfiguration()
    {
        return typeof(TripwireReRunDriver).Assembly
            .GetCustomAttribute<AssemblyConfigurationAttribute>()?.Configuration ?? string.Empty;
    }

    private static string? CollectionOf(Type type)
    {
        return type.GetCustomAttributesData()
            .Where(a => string.Equals(a.AttributeType.Name, "CollectionAttribute", StringComparison.Ordinal))
            .Where(a => a.ConstructorArguments.Count == 1)
            .Select(a => a.ConstructorArguments[0].Value as string)
            .FirstOrDefault(value => value != null);
    }
}
