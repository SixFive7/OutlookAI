using OutlookAI.Core.Com;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the index-hit locator's path fallback over folders that share a path (decided 2026-11-18, option (a)
/// of the TODO item stream A's finding opened: "try every folder at that path, not the first"). Two folders of
/// one name can sit side by side - the Exchange guest has two top-level <c>Conversation Action Settings</c> -
/// and the index keeps one path for both, so a hit in the second could not be opened while
/// <c>Folders[name]</c> answered with the first. <see cref="FolderPathCandidates"/> is the walk that finds
/// every folder at a path; <see cref="OutlookComSession.MayTryNamesakes"/> says when the locator looks for
/// them; <see cref="HitLocator.MoreInformative"/> which failure it keeps. A tree of plain objects stands in
/// for Outlook's folders, and counts what is released: a walk over COM must hand back every folder it opened.
/// </summary>
public sealed class FolderPathCandidatesTests
{
    private sealed class Node
    {
        internal Node(string? name, params Node[] children)
        {
            Name = name;
            Children = children;
        }

        internal string? Name { get; }

        internal Node[] Children { get; }

        internal int Released { get; set; }

        internal int Listed { get; set; }
    }

    private static IReadOnlyList<(string? Name, Node Folder)> ChildrenOf(Node folder)
    {
        folder.Listed++;
        return folder.Children.Select(c => (c.Name, c)).ToList();
    }

    private static IReadOnlyList<Node> Walk(Node root, params string[] path)
        => FolderPathCandidates.FindAll(root, path, ChildrenOf, n => n.Released++);

    [Fact]
    public void TwoFoldersOfOneNameSideBySide_AreBothFound_InTheOrderTheirParentListsThem()
    {
        // The measured shape: two top-level folders of one name.
        Node first = new Node("Conversation Action Settings");
        Node second = new Node("Conversation Action Settings");
        Node root = new Node("Top", new Node("Inbox"), first, second, new Node("Sent Items"));

        IReadOnlyList<Node> found = Walk(root, "Conversation Action Settings");

        Assert.Equal(new[] { first, second }, found);

        // Control: a path only one folder has gives that one alone.
        Assert.Single(Walk(root, "Inbox"));
    }

    [Fact]
    public void NamesakesAboveThePath_AreWalkedDownBoth_SoTheFolderUnderEitherIsFound()
    {
        Node underFirst = new Node("Projects");
        Node underSecond = new Node("Projects");
        Node root = new Node(
            "Top",
            new Node("Archive", new Node("2025"), underFirst),
            new Node("Archive", underSecond));

        Assert.Equal(new[] { underFirst, underSecond }, Walk(root, "Archive", "Projects"));
    }

    [Fact]
    public void NamesAreComparedAsOutlookRefusesADuplicate_CaseInsensitive_NothingTrimmed()
    {
        Node exact = new Node("inbox");
        Node spaced = new Node("Inbox ");
        Node root = new Node("Top", exact, spaced, new Node(null));

        Assert.Equal(new[] { exact }, Walk(root, "INBOX"));
        Assert.True(FolderPathCandidates.NameMatches("Ünïcode", "ÜNÏCODE") || !FolderPathCandidates.NameMatches("x", "y"));
        Assert.False(FolderPathCandidates.NameMatches(null, "Inbox"));
        Assert.False(FolderPathCandidates.NameMatches("Inbox", null));
    }

    [Fact]
    public void AFolderNameHoldingASlash_IsMatchedAsOneName_NeverSplit()
    {
        // Folders[name] does not answer for a name holding '/' (Q99); reading each child's own name does.
        Node slashed = new Node("a/b");
        Node root = new Node("Top", new Node("a", new Node("b")), slashed);

        Assert.Equal(new[] { slashed }, Walk(root, "a/b"));
    }

    [Fact]
    public void AMissingName_FindsNothing_AndTheEmptyPathHasNoNamesake()
    {
        Node root = new Node("Top", new Node("Inbox", new Node("Sub")));

        Assert.Empty(Walk(root, "Inbox", "Missing"));
        Assert.Empty(Walk(root, "Missing", "Sub"));
        Assert.Empty(Walk(root));
        Assert.Equal(0, root.Released);
    }

    [Fact]
    public void EveryFolderTheWalkOpened_IsReleasedOnce_ButTheRootAndWhatItReturns()
    {
        Node keptA = new Node("Projects", new Node("x"));
        Node keptB = new Node("Projects");
        Node otherUnderA = new Node("Other");
        Node archiveA = new Node("Archive", keptA, otherUnderA);
        Node archiveB = new Node("Archive", keptB);
        Node sent = new Node("Sent Items");
        Node root = new Node("Top", archiveA, sent, archiveB);

        IReadOnlyList<Node> found = Walk(root, "Archive", "Projects");

        Assert.Equal(new[] { keptA, keptB }, found);
        Assert.Equal(0, root.Released);
        Assert.Equal(0, keptA.Released);
        Assert.Equal(0, keptB.Released);
        Assert.Equal(1, sent.Released);
        Assert.Equal(1, archiveA.Released);
        Assert.Equal(1, archiveB.Released);
        Assert.Equal(1, otherUnderA.Released);
        Assert.Equal(0, keptA.Listed);
    }

    [Fact]
    public void TheCapBoundsEveryLevel_AndWhatItLeavesOutIsReleased()
    {
        Node[] many = Enumerable.Range(0, 5).Select(_ => new Node("Same")).ToArray();
        Node root = new Node("Top", many);

        IReadOnlyList<Node> found = FolderPathCandidates.FindAll(root, new[] { "Same" }, ChildrenOf, n => n.Released++, cap: 2);

        Assert.Equal(new[] { many[0], many[1] }, found);
        Assert.All(many.Skip(2), n => Assert.Equal(1, n.Released));
        Assert.Throws<ArgumentOutOfRangeException>(() => FolderPathCandidates.FindAll(root, new[] { "Same" }, ChildrenOf, n => n.Released++, cap: 0));
    }

    [Fact]
    public void AListingThatFails_ReleasesWhatTheWalkHeld_AndThrows()
    {
        Node failing = new Node("Archive");
        Node keptAlready = new Node("Archive", new Node("Projects"));
        Node notListedYet = new Node("Archive", new Node("Projects"));
        Node root = new Node("Top", keptAlready, failing, notListedYet);

        IReadOnlyList<(string? Name, Node Folder)> Children(Node n)
            => ReferenceEquals(n, failing) ? throw new InvalidOperationException("the listing failed") : ChildrenOf(n);

        List<Node> released = new List<Node>();
        Assert.Throws<InvalidOperationException>(() =>
            FolderPathCandidates.FindAll(root, new[] { "Archive", "Projects" }, Children, n => { n.Released++; released.Add(n); }));

        Assert.Equal(1, failing.Released);
        Assert.Equal(1, keptAlready.Released);
        Assert.Equal(1, notListedYet.Released);
        Assert.Equal(1, keptAlready.Children[0].Released);
        Assert.Equal(0, root.Released);
    }

    // ============================================================ when the locator looks, and what it keeps

    [Theory]
    [InlineData("FolderNotFound", true)]
    [InlineData("NoSubjectTimeMatch", true)]
    [InlineData("NoTimeOnlyMatch", true)]
    [InlineData("FolderTooLargeForTimeOnlyProbe", true)]
    [InlineData("COMException 0x80040107", true)]
    [InlineData("StoreNotFound", false)]
    [InlineData("RootFolderUnavailable", false)]
    [InlineData("EmptySubjectAndNoReceivedTime", false)]
    public void TheNamesakesAreLookedFor_WhenTheFirstFolderCouldNotGiveTheItem(string reason, bool looks)
    {
        Assert.Equal(looks, OutlookComSession.MayTryNamesakes(new[] { "Inbox" }, reason));
    }

    [Fact]
    public void TheStoresTopFolder_AndASuccess_LookForNoNamesake()
    {
        Assert.False(OutlookComSession.MayTryNamesakes(Array.Empty<string>(), "NoSubjectTimeMatch"));
        Assert.False(OutlookComSession.MayTryNamesakes(new[] { "Inbox" }, null));
    }

    [Fact]
    public void TheFailureKept_IsTheFolderThatExistsWithoutTheItem_OverNoFolderOfThatName()
    {
        Assert.Equal("NoSubjectTimeMatch", HitLocator.MoreInformative("FolderNotFound", "NoSubjectTimeMatch"));
        Assert.Equal("NoSubjectTimeMatch", HitLocator.MoreInformative("NoSubjectTimeMatch", "FolderNotFound"));
        Assert.Equal("NoSubjectTimeMatch", HitLocator.MoreInformative("NoSubjectTimeMatch", null));
        Assert.Equal("COMException 0x80004005", HitLocator.MoreInformative("NoSubjectTimeMatch", "COMException 0x80004005"));
        Assert.Null(HitLocator.MoreInformative(null, "FolderNotFound"));
    }
}
