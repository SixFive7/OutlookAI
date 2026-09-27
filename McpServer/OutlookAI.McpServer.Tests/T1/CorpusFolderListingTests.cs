using System.Globalization;
using OutlookAI.Core.Com;
using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins <see cref="CorpusFolderListing"/> - the pure half of <c>corpus-folders</c>, the read-only verb
/// added 2026-09-27 so the first guest run of population v2 could answer, with the project's own tool
/// and not an improvised COM script, whether a PST attached with <c>AddStoreEx</c> still lacks its
/// default folders after a build, and whether a teardown leaves emptied folders in Deleted Items.
/// </summary>
public sealed class CorpusFolderListingTests
{
    private static ListedDefaultFolder Absent(int id) => new(id, OutlookComSession.DefaultFolderResolution.Absent, false, null);

    [Fact]
    public void TheDefaultFolders_AreListedInAFixedOrder_WithTheirEnglishLabels()
    {
        Assert.Equal(new[] { 6, 5, 3, 4, 16, 23, 9, 10, 13 }, CorpusFolderListing.DefaultFolderIds);
        Assert.Equal(
            new[] { "Inbox", "Sent Items", "Deleted Items", "Outbox", "Drafts", "Junk Email", "Calendar", "Contacts", "Tasks" },
            CorpusFolderListing.DefaultFolderIds.Select(CorpusFolderListing.Label));
        Assert.Equal("39", CorpusFolderListing.Label(39));
    }

    [Fact]
    public void EachResolution_SaysWhatItIs_AndTheHiddenRootIsNamedAsSuch()
    {
        // The bystander's measured shape (OAI-UNINDEXED, 2026-09-24): Deleted Items and nothing else -
        // and, asked the old way, an "Inbox" that was the PST's hidden root.
        var listing = new FolderListing(
            new[]
            {
                new ListedDefaultFolder(6, OutlookComSession.DefaultFolderResolution.Resolved, true, null),
                Absent(5),
                new ListedDefaultFolder(3, OutlookComSession.DefaultFolderResolution.Resolved, false, "Deleted Items"),
                new ListedDefaultFolder(4, OutlookComSession.DefaultFolderResolution.Unreadable, false, null),
            },
            Array.Empty<ListedFolder>(),
            false);

        IReadOnlyList<string> lines = CorpusFolderListing.Render("bystander@vm.invalid", listing);
        Assert.Contains(lines, l => l.StartsWith("  Inbox ", StringComparison.Ordinal) && l.Contains("NOT in the visible tree", StringComparison.Ordinal));
        Assert.Contains(lines, l => l.StartsWith("  Sent Items ", StringComparison.Ordinal) && l.EndsWith("ABSENT", StringComparison.Ordinal));
        Assert.Contains(lines, l => l.StartsWith("  Deleted Items ", StringComparison.Ordinal) && l.EndsWith("'Deleted Items'", StringComparison.Ordinal));
        Assert.Contains(lines, l => l.StartsWith("  Outbox ", StringComparison.Ordinal) && l.EndsWith("UNREADABLE", StringComparison.Ordinal));
        Assert.Contains("read-only", lines[0], StringComparison.Ordinal);
        Assert.Contains("  (no folder)", lines);
    }

    [Fact]
    public void TheTree_IsIndentedByDepth_CountsItems_AndMarksWhatTheCorpusToolMade()
    {
        // A hub after a v2 rebuild, and the question section 3b item 7 asks: are the emptied
        // subfolders of the last population sitting in Deleted Items?
        var listing = new FolderListing(
            CorpusFolderListing.DefaultFolderIds.Select(Absent).ToList(),
            new[]
            {
                new ListedFolder("Inbox", 0, 32, 0),
                new ListedFolder("Inbox/OutlookAI-Corpus-Folder-Projects", 1, 6, 0),
                new ListedFolder("Deleted Items", 0, 0, 0),
                new ListedFolder("Deleted Items/OutlookAI-Corpus-Folder-Notices", 1, 0, 0),
                new ListedFolder("OutlookAI-Corpus-Folder-9", 0, 4, 1),
                new ListedFolder("Odd", 0, null, null),
            },
            false);

        IReadOnlyList<string> lines = CorpusFolderListing.Render("tier@vm.invalid", listing);
        Assert.Contains("  Inbox  items=32  mail", lines);
        Assert.Contains("    OutlookAI-Corpus-Folder-Projects  items=6  mail  [made by the corpus tool]", lines);
        Assert.Contains("    OutlookAI-Corpus-Folder-Notices  items=0  mail  [made by the corpus tool]", lines);
        Assert.Contains("  OutlookAI-Corpus-Folder-9  items=4  calendar  [made by the corpus tool]", lines);
        Assert.Contains("  Odd  items=?  kind?", lines);
        Assert.DoesNotContain(lines, l => l.Contains("TRUNCATED", StringComparison.Ordinal));
    }

    [Fact]
    public void ATruncatedWalk_SaysSo_WithItsCaps()
    {
        var listing = new FolderListing(Array.Empty<ListedDefaultFolder>(), new[] { new ListedFolder("A", 0, 1, 0) }, true);
        string last = CorpusFolderListing.Render("s", listing)[^1];
        Assert.Contains("TRUNCATED", last, StringComparison.Ordinal);
        Assert.Contains(CorpusFolderListing.MaxFolders.ToString(CultureInfo.InvariantCulture), last, StringComparison.Ordinal);
        Assert.Contains(CorpusFolderListing.MaxDepth.ToString(CultureInfo.InvariantCulture), last, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0, "mail")]
    [InlineData(1, "calendar")]
    [InlineData(2, "contacts")]
    [InlineData(3, "tasks")]
    [InlineData(5, "notes")]
    [InlineData(7, "type 7")]
    public void AFolderKind_IsNamedFromItsDefaultItemType(int type, string kind)
    {
        Assert.Equal(kind, CorpusFolderListing.KindOf(type));
        Assert.Equal("kind?", CorpusFolderListing.KindOf(null));
    }

    [Fact]
    public void TheVerb_NeedsAStore_BeforeItTouchesAnything()
    {
        // corpus-folders runs behind the same guard as every other verb; without --store it refuses
        // before reading a single store fact.
        CorpusOptions none = CorpusOptions.Parse(Array.Empty<string>());
        using var output = new StringWriter(CultureInfo.InvariantCulture);
        ArgumentException refused = Assert.Throws<ArgumentException>(() => CorpusCommands.RunFolders(none, output));
        Assert.Contains("--store", refused.Message, StringComparison.Ordinal);
        Assert.Equal(string.Empty, output.ToString());
    }
}
