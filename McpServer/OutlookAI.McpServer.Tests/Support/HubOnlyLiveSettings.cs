using OutlookAI.McpServer.Tests.T2;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// Live-test settings that name one store, the hub, and nothing else: the least a rule of the live
/// tier can be asked about. The one copy (Q154).
/// </summary>
internal static class HubOnlyLiveSettings
{
    /// <summary>The hub alone, on a machine of <paramref name="profile"/>.</summary>
    public static LiveTestSettings For(LiveMachineProfile profile)
    {
        return new LiveTestSettings
        {
            MachineProfile = profile,
            TestHubStoreDisplayName = "hub@example.test",
            ExpectedStoreDisplayNames = new List<string> { "hub@example.test" },
        };
    }
}
