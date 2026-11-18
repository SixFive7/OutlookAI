using System.Reflection;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// Where this test project and the repository around it are on disk, for a test that reads a file
/// of the repository. The test project's directory is baked into the assembly at build time
/// (<c>AssemblyMetadata("TestProjectDir")</c> in the csproj), so the answer is the checkout that was
/// built - the expanded archive on the build VM, a worktree here.
/// <para>
/// The one copy (Q154): a test that needs either path calls this, and reads no assembly metadata
/// of its own.
/// </para>
/// </summary>
internal static class RepoPaths
{
    /// <summary>The test project's directory, as it was when the assembly was built.</summary>
    public static string TestProjectDir()
    {
        return typeof(RepoPaths).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
    }

    /// <summary>The repository root: <c>&lt;repo&gt;/McpServer/OutlookAI.McpServer.Tests/</c> is two levels below it.</summary>
    public static string RepoRoot()
    {
        return Path.GetFullPath(Path.Combine(TestProjectDir(), "..", ".."));
    }
}
