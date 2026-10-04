using System.Diagnostics;
using System.Globalization;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// The shipped server, started where Claude Code starts it - in a project folder - spends no CPU
/// on file activity in that folder.
/// <para>
/// Up to 3.1.0.325 it did. Its generic host made the working directory its content root and
/// watched it, recursively, for a settings file the server never reads, so every file created,
/// changed or deleted in a project cost each server started there a check, and every renamed
/// folder a walk of the whole renamed tree (<see cref="ServerHost"/>). On the build VM on
/// 2026-10-04 this test's activity cost the unfixed server 2.281 CPU-seconds - churn 0.813, tree
/// 0.375, renames 1.094 - and the fixed one 0.000 in every phase, which is nothing Windows can
/// measure. <c>T1.ServerHostWatchesNothingTests</c> pins the two settings that fixed it; this holds
/// the exe that ships to the effect, from the outside.
/// </para>
/// <para>
/// CI-safe: nothing but <c>initialize</c> goes over the wire, so nothing here reaches Outlook.
/// </para>
/// </summary>
public sealed class ServerWorkingDirectoryCiTests
{
    private const int Files = 1000;
    private const int Renames = 30;

    /// <summary>
    /// A ninth of what the unfixed server spent on the build VM, and sixteen of the 15.6 ms steps in
    /// which Windows counts a process's CPU time above what the fixed one spent.
    /// </summary>
    private static readonly TimeSpan Budget = TimeSpan.FromSeconds(0.25);

    /// <summary>The longest a settle may take; the unfixed server caught up within 2 s.</summary>
    private static readonly TimeSpan SettleCap = TimeSpan.FromSeconds(60);

    private readonly ITestOutputHelper _output;

    public ServerWorkingDirectoryCiTests(ITestOutputHelper output)
    {
        _output = output;
    }

    [Fact]
    public async Task FileActivityInTheWorkingDirectory_CostsTheServerNoCpu()
    {
        string project = Path.Combine(Path.GetTempPath(), "OutlookAI-ServerWorkingDirectory-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(project);
        try
        {
            await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(workingDirectory: project);
            Process server = client.ServerProcess;
            await SettleAsync(server);

            // What builds and checkouts do to a project, one kind at a time, each caught up with
            // before the next: files created, changed and deleted; a tree written; and the tree
            // renamed back and forth - for a watcher, a walk of the whole tree every time.
            string churn = Path.Combine(project, "churn");
            string tree = Path.Combine(project, "tree");
            string renamed = Path.Combine(project, "tree-renamed");
            (string Name, TimeSpan Cpu)[] phases =
            [
                ("churn", await CpuSpentOnAsync(server, () => Churn(churn))),
                ("tree", await CpuSpentOnAsync(server, () => WriteTree(tree))),
                ("renames", await CpuSpentOnAsync(server, () => RenameBackAndForthAsync(tree, renamed))),
            ];

            TimeSpan spent = phases.Aggregate(TimeSpan.Zero, (sum, phase) => sum + phase.Cpu);
            string detail = string.Join(", ", phases.Select(phase => $"{phase.Name} {Seconds(phase.Cpu)}"));
            _output.WriteLine($"CPU-seconds spent: {detail}; {Seconds(spent)} in all, budget {Seconds(Budget)}");

            Assert.True(
                spent < Budget,
                $"The server spent {Seconds(spent)} CPU-seconds ({detail}) on file activity in its working directory - "
                + $"{Files} files created, changed and deleted, a {Files}-file tree written and renamed {Renames} times - "
                + $"against a budget of {Seconds(Budget)}. Something in it watches the folder it was started in: see ServerHost.");
        }
        finally
        {
            await DeleteAsync(project);
        }
    }

    private static Task Churn(string folder)
    {
        Directory.CreateDirectory(folder);
        for (int i = 0; i < Files; i++)
        {
            File.WriteAllText(Path.Combine(folder, $"f{i}.txt"), "a");
        }

        for (int i = 0; i < Files; i++)
        {
            File.AppendAllText(Path.Combine(folder, $"f{i}.txt"), "b");
        }

        for (int i = 0; i < Files; i++)
        {
            File.Delete(Path.Combine(folder, $"f{i}.txt"));
        }

        return Task.CompletedTask;
    }

    private static Task WriteTree(string root)
    {
        for (int d = 0; d < 10; d++)
        {
            string folder = Directory.CreateDirectory(Path.Combine(root, $"d{d}")).FullName;
            for (int i = 0; i < Files / 10; i++)
            {
                File.WriteAllText(Path.Combine(folder, $"f{i}.txt"), "x");
            }
        }

        return Task.CompletedTask;
    }

    private static async Task RenameBackAndForthAsync(string tree, string renamed)
    {
        for (int r = 0; r < Renames; r++)
        {
            await MoveAsync(r % 2 == 0 ? tree : renamed, r % 2 == 0 ? renamed : tree);
        }
    }

    private static async Task<TimeSpan> CpuSpentOnAsync(Process server, Func<Task> activity)
    {
        TimeSpan before = CpuOf(server);
        await activity();
        await SettleAsync(server);
        return CpuOf(server) - before;
    }

    private static string Seconds(TimeSpan span) => span.TotalSeconds.ToString("0.000", CultureInfo.InvariantCulture);

    private static TimeSpan CpuOf(Process server)
    {
        server.Refresh();
        return server.TotalProcessorTime;
    }

    /// <summary>
    /// Returns once the server has gone a whole second without a measurable step of CPU time: a
    /// watcher still working through a burst is never that quiet.
    /// </summary>
    private static async Task SettleAsync(Process server)
    {
        var samples = new Queue<TimeSpan>();
        var clock = Stopwatch.StartNew();
        samples.Enqueue(CpuOf(server));
        while (clock.Elapsed < SettleCap)
        {
            await Task.Delay(250);
            samples.Enqueue(CpuOf(server));
            if (samples.Count > 5)
            {
                samples.Dequeue();
            }

            if (samples.Count == 5 && samples.Last() - samples.Peek() < TimeSpan.FromMilliseconds(10))
            {
                return;
            }
        }
    }

    /// <summary>
    /// Removes the scratch folder once the server is gone. An antivirus scan can still hold a file
    /// in it for a moment, and a folder left in the temp directory must not fail the test.
    /// </summary>
    private static async Task DeleteAsync(string folder)
    {
        for (int attempt = 0; attempt < 50 && Directory.Exists(folder); attempt++)
        {
            try
            {
                Directory.Delete(folder, recursive: true);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                await Task.Delay(100);
            }
        }
    }

    /// <summary>
    /// A folder cannot be renamed while anything holds a handle inside it - an antivirus scan, or a
    /// watcher walking the tree it was just told about - so a rename is retried for up to 10 s.
    /// </summary>
    private static async Task MoveAsync(string from, string to)
    {
        var clock = Stopwatch.StartNew();
        while (true)
        {
            try
            {
                Directory.Move(from, to);
                return;
            }
            catch (Exception ex) when ((ex is IOException || ex is UnauthorizedAccessException) && clock.Elapsed < TimeSpan.FromSeconds(10))
            {
                await Task.Delay(20);
            }
        }
    }
}
