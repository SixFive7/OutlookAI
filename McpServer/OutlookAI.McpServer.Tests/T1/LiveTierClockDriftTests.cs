using System.Runtime.CompilerServices;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Keeps the live tier's waits on a clock that only goes forward.
/// <para>
/// Thirteen live-tier loops measured their own timeout as
/// <c>DateTime deadline = DateTime.UtcNow + timeout; while (DateTime.UtcNow &lt; deadline)</c>.
/// That is a DURATION on the wall clock, and the wall clock jumps - an NTP correction, a
/// person setting the time, a VM resuming. A backwards jump lengthens every wait by the size
/// of the jump, which on this suite is indistinguishable from the hang that already cost a
/// night; a forwards jump ends a wait for real mail before it could have arrived. They now
/// use <see cref="LiveWaitBudget"/>, and this test is what stops the fourteenth copy.
/// </para>
/// <para>
/// Deliberately scoped to the live fixtures (T2/T3). The server and add-in halves were swept
/// separately and their remaining wall-clock reads are ABSOLUTE INSTANTS on purpose - a DASL
/// window base, a screenshot filename, an audit stamp - so a scan there would be mostly
/// exceptions. It is also text-based on purpose: no compiler can see this rule.
/// </para>
/// </summary>
public sealed class LiveTierClockDriftTests
{
    /// <summary>
    /// Directories scanned. A live-tier directory that stops existing is a failure, not a
    /// pass: a guard that quietly finds nothing to check has switched itself off.
    /// </summary>
    private static readonly string[] ScannedDirectories = { "T2", "T3" };

    [Fact]
    public void NoLiveTierWaitMeasuresItsOwnTimeoutOnTheWallClock()
    {
        List<string> offenders = new();
        int scannedFiles = 0;

        foreach (string relative in ScannedDirectories)
        {
            string directory = Path.Combine(TestsProjectDirectory(), relative);
            Assert.True(
                Directory.Exists(directory),
                $"The clock-drift guard cannot find '{relative}'. If the live tier moved, point this test at it - "
                + "a scan with nothing to scan proves nothing.");

            foreach (string file in Directory.GetFiles(directory, "*.cs", SearchOption.AllDirectories))
            {
                scannedFiles++;
                string[] lines = File.ReadAllLines(file);
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i];
                    if (IsComment(line) || !MeasuresElapsedOnTheWallClock(line))
                    {
                        continue;
                    }

                    offenders.Add($"{Path.GetFileName(file)}:{i + 1}: {line.Trim()}");
                }
            }
        }

        Assert.True(scannedFiles > 20, $"Only {scannedFiles} live-tier source files were scanned; that is too few to trust.");
        Assert.True(
            offenders.Count == 0,
            "These live-tier waits measure elapsed time on the wall clock, which jumps. Use LiveWaitBudget:"
            + Environment.NewLine + string.Join(Environment.NewLine, offenders));
    }

    [Fact]
    public void TheDetectorRecognisesTheShapeItWasWrittenFor()
    {
        // A scan that passes because its pattern no longer matches anything is a scan that
        // has switched itself off. These are the exact lines this guard replaced.
        Assert.True(MeasuresElapsedOnTheWallClock("        DateTime deadline = DateTime.UtcNow + timeout;"));
        Assert.True(MeasuresElapsedOnTheWallClock("        DateTime deadline = DateTime.UtcNow.AddSeconds(180);"));
        Assert.True(MeasuresElapsedOnTheWallClock("        while (DateTime.UtcNow < deadline)"));
        Assert.True(MeasuresElapsedOnTheWallClock("            while (hitId == null && DateTime.UtcNow < deadline)"));

        // ...and leaves the absolute instants alone, which is why it is a shape and not a ban.
        Assert.False(MeasuresElapsedOnTheWallClock("        DateTime sentUtc = DateTime.UtcNow;"));
        Assert.False(MeasuresElapsedOnTheWallClock("            $\"phase3-search-{DateTime.Now:yyyyMMdd-HHmmss}.png\");"));
        Assert.False(MeasuresElapsedOnTheWallClock("            ReceivedOnOrAfterUtc = DateTime.UtcNow.AddDays(-30),"));
    }

    [Fact]
    public void ALiveWaitBudget_CountsForwardsAndRunsOut()
    {
        LiveWaitBudget spent = LiveWaitBudget.OfSeconds(0);
        Assert.False(spent.HasTimeLeft);

        LiveWaitBudget wait = LiveWaitBudget.Of(TimeSpan.FromMinutes(5));
        Assert.True(wait.HasTimeLeft);
        Assert.Equal(TimeSpan.FromMinutes(5), wait.Budget);

        TimeSpan first = wait.Elapsed;
        Thread.Sleep(2);
        Assert.True(wait.Elapsed >= first, "elapsed time must never go backwards");
        Assert.True(first >= TimeSpan.Zero, "elapsed time must never be negative");
    }

    /// <summary>
    /// A wall-clock read that is part of measuring how long something has been going: a
    /// deadline built from "now", or a loop that keeps re-reading "now" to decide whether to
    /// stop. Absolute instants - a send time, a screenshot filename, a sweep window base -
    /// are left alone, because for those the wall clock is the correct clock.
    /// </summary>
    private static bool MeasuresElapsedOnTheWallClock(string line)
    {
        if (!line.Contains("DateTime.UtcNow", StringComparison.Ordinal)
            && !line.Contains("DateTime.Now", StringComparison.Ordinal))
        {
            return false;
        }

        return line.Contains("deadline", StringComparison.OrdinalIgnoreCase)
            || line.Contains("while (DateTime.", StringComparison.Ordinal)
            || line.Contains("&& DateTime.", StringComparison.Ordinal);
    }

    private static bool IsComment(string line)
    {
        string trimmed = line.TrimStart();
        return trimmed.StartsWith("//", StringComparison.Ordinal)
            || trimmed.StartsWith("*", StringComparison.Ordinal);
    }

    /// <summary>
    /// The test project's own directory, taken from this file's compile-time path. The same
    /// checkout builds and runs the tests (locally and in the McpServer workflow), so the
    /// path is the real one rather than a guess relative to the output folder.
    /// </summary>
    private static string TestsProjectDirectory([CallerFilePath] string thisFile = "")
    {
        return Path.GetFullPath(Path.Combine(Path.GetDirectoryName(thisFile)!, ".."));
    }
}
