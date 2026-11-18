using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// <c>Testbed/live-test-settings.example.json</c>, the committed shape of the gitignored live-test
/// settings - the file a rebuilder copies. The one copy of where it is and how a test reads it (Q154).
/// </summary>
internal static class ExampleLiveSettings
{
    /// <summary>Where the committed example is, in the checkout that was built.</summary>
    public static string FilePath => Path.Combine(RepoPaths.RepoRoot(), "Testbed", "live-test-settings.example.json");

    /// <summary>The committed example settings, parsed by the real loader.</summary>
    public static LiveTestSettings Load()
    {
        string path = FilePath;
        Assert.True(File.Exists(path), "the committed example settings file is missing: " + path);
        return LiveTestSettings.Parse(File.ReadAllText(path));
    }
}
