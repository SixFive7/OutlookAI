using Xunit;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// The lines of an audit log a live test reads to see that an operation was audited (Q154): how many
/// there were before it, and those added since. The caller names the log - this process's own,
/// <c>AuditLog.EffectiveLogPath</c>, where the service runs in-process, and the real one only where the
/// server CHILD wrote the lines - so a class that reads the real log goes on naming it itself, as
/// <c>T1/AuditLogIsolationTests</c>' list of the classes allowed to requires.
/// </summary>
internal static class AuditLines
{
    /// <summary>How many lines the log at <paramref name="path"/> holds; none when it is not there yet.</summary>
    public static int Count(string path)
    {
        return File.Exists(path) ? File.ReadAllLines(path).Length : 0;
    }

    /// <summary>The lines added to the log at <paramref name="path"/> since it held <paramref name="skip"/>.</summary>
    public static IReadOnlyList<string> After(string path, int skip)
    {
        Assert.True(File.Exists(path), "the audit log must exist after the operations under test");
        return File.ReadAllLines(path).Skip(skip).ToList();
    }
}
