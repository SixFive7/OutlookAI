using System.Reflection;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The product never reads <c>PR_CONVERSATION_INDEX_TRACKING</c> (PidTagConversationIndexTracking,
/// 0x3016). On 2026-10-03 a diagnostic read of it, made on every draft snapshot through the item's
/// PropertyAccessor (commit <c>67c4afb</c>), was the one change whose removal stopped OUTLOOK.EXE
/// (Office LTSC 2024, 16.0.17932) crashing on the test guest: with it, four of four full live runs of
/// the branch merged with master <c>21bbdfd</c> crashed Outlook in the drafting collection
/// (ntdll.dll and OLMAPI32.DLL access violations); with only that read removed, the same run did
/// not. OutlookAI must never crash a user's Outlook, so the read stays out - this pins it.
/// <para>
/// Read out of the sources, because the failure needs a real Outlook. Nothing here touches one.
/// </para>
/// </summary>
public sealed class ConversationTrackingReadTests
{
    [Fact]
    public void NoProductSource_ReadsConversationIndexTracking()
    {
        string coreDir = Path.GetFullPath(Path.Combine(TestProjectDir(), "..", "OutlookAI.Core"));
        List<string> offenders = new List<string>();
        int scanned = 0;
        foreach (string file in Directory.EnumerateFiles(coreDir, "*.cs", SearchOption.AllDirectories))
        {
            if (file.Contains(Path.DirectorySeparatorChar + "obj" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                || file.Contains(Path.DirectorySeparatorChar + "bin" + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            scanned++;
            string text = File.ReadAllText(file);
            if (text.Contains("0x3016", StringComparison.OrdinalIgnoreCase)
                || text.Contains("ConversationIndexTracking", StringComparison.Ordinal))
            {
                offenders.Add(Path.GetFileName(file));
            }
        }

        Assert.True(scanned > 50, "only " + scanned + " Core source file(s) found - this test has stopped proving anything");
        Assert.Empty(offenders);
    }

    private static string TestProjectDir()
    {
        return typeof(ConversationTrackingReadTests).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
    }
}
