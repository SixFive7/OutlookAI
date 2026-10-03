using System.Reflection;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The re-locate after discard_draft, on a store that KEEPS an item's EntryID across a soft
/// delete. On the first live run on a test guest (2026-10-03, a POP3 PST) the discarded draft's old
/// id still opened after a discard that reported Deleted Items as its destination. The re-locate
/// scan excludes the old id - right on Exchange, which mints a new one on the move - so on that
/// store it could find nothing, or an older discarded draft of the same subject. A non-Exchange
/// store is now asked for the kept id first; Exchange, and a store whose kind cannot be read, keep
/// the scan exactly as before.
/// <para>
/// Read out of the source, because the path needs a real Outlook; the live half is
/// <c>T2/LiveUpdateDiscardTests</c>.
/// </para>
/// </summary>
public sealed class DiscardReLocateTests
{
    [Fact]
    public void ANonExchangeStore_IsAskedForTheKeptIdBeforeAnyScan()
    {
        string body = CoreMemberBody("private string? TryFindDiscardedCopy(");

        int guard = body.IndexOf("if (!exchangeStore)", StringComparison.Ordinal);
        int kept = body.IndexOf("TryKeptEntryIdInFolder(oldEntryId, deletedItemsEntryId)", StringComparison.Ordinal);
        int scan = body.IndexOf(".GetFolderFromID(deletedItemsEntryId)", StringComparison.Ordinal);

        Assert.True(scan >= 0, "the re-locate scan is no longer where this test reads it - it has stopped proving anything");
        Assert.True(guard >= 0 && kept > guard, "a non-Exchange store is no longer asked for the id it kept");
        Assert.True(kept < scan, "the kept id must be asked for before the scan, which excludes it");
    }

    [Fact]
    public void TheKeptIdCounts_OnlyWhenItOpensInDeletedItems()
    {
        // An id that still opens somewhere ELSE - in Drafts, say, because the discard did not move it
        // at all - is not a re-located copy, and must never be reported as one.
        string body = CoreMemberBody("private string? TryKeptEntryIdInFolder(");

        Assert.Contains("GetItemFromID(entryId)", body, StringComparison.Ordinal);
        Assert.Contains("string.Equals(parentId, folderEntryId, StringComparison.OrdinalIgnoreCase)", body, StringComparison.Ordinal);
    }

    [Fact]
    public void AStoreOfUnknownKind_ReLocatesExactlyAsBefore()
    {
        // The discard reads the store's kind; failing to read it must leave the Exchange behaviour
        // (scan only) in place rather than ask a question whose answer Exchange may get wrong.
        string text = CoreSource();

        Assert.Contains("bool exchangeStore = true;", text, StringComparison.Ordinal);
        Assert.Contains(
            "TryFindDiscardedCopy(deletedItemsEntryId, info.Subject, info.EntryId, exchangeStore)",
            text,
            StringComparison.Ordinal);
    }

    private static string CoreSource()
    {
        string testProjectDir = typeof(DiscardReLocateTests).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");

        // <repo>/McpServer/OutlookAI.McpServer.Tests/ -> <repo>/McpServer/OutlookAI.Core/Com/OutlookComSession.cs
        return File.ReadAllText(Path.GetFullPath(
            Path.Combine(testProjectDir, "..", "OutlookAI.Core", "Com", "OutlookComSession.cs")));
    }

    private static string CoreMemberBody(string declarationStart)
    {
        string[] lines = CoreSource().Split('\n').Select(l => l.TrimEnd('\r')).ToArray();
        int declaration = Array.FindIndex(lines, l => l.TrimStart().StartsWith(declarationStart, StringComparison.Ordinal));
        Assert.True(declaration >= 0, declarationStart + " was not found - this test has stopped proving anything");
        int open = Array.FindIndex(lines, declaration, l => string.Equals(l, "        {", StringComparison.Ordinal));
        int close = open < 0 ? -1 : Array.FindIndex(lines, open + 1, l => string.Equals(l, "        }", StringComparison.Ordinal));
        Assert.True(open > declaration && close > open, "could not find the body of " + declarationStart);
        return string.Join("\n", lines[(open + 1)..close]);
    }
}
