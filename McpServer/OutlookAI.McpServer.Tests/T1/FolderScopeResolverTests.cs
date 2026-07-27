using OutlookAI.Core.Mapi;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Folder-scope resolution (soak fix 15) - the predicate shapes per store type and
/// include_subfolders value, pinned as pure logic.
/// <para>
/// The defect these pins exist for: DELEGATE stores are indexed FLAT. A delegate folder
/// lives at <c>&lt;host&gt;/1/&lt;delegate&gt;/&lt;LEAF&gt;</c> with every intermediate
/// COM folder dropped, so the old nested construction
/// (<c>delegateScope + "/" + folderPath</c>) addressed a URL that does not exist and
/// returned 0 rows, SILENTLY, for every delegate subfolder - about 3,871 items across 15
/// subfolders on the discovery machine.
/// </para>
/// <para>
/// Fixtures are synthetic (S6: no live-mailbox identifiers in the repo) but structurally
/// identical to the measured URLs.
/// </para>
/// </summary>
public sealed class FolderScopeResolverTests
{
    private const string Sid = "{S-1-5-21-1111111111-2222222222-3333333333-1001}";
    private const string PrimaryPrefix = "mapi16://" + Sid + "/alice@example.com($deadbeef)";
    private const string DelegatePrefix = "mapi16://" + Sid + "/alice@example.com($deadbeef)/1/Sam Delegate";

    // ------------------------------------------------------- display-path derivation

    [Theory]
    // Primary store, nesting preserved in both URL and display path.
    [InlineData(PrimaryPrefix + "/0/Inbox", "/alice@example.com/Inbox")]
    [InlineData(PrimaryPrefix + "/0/Inbox/Fun/Immich", "/alice@example.com/Inbox/Fun/Immich")]
    [InlineData(PrimaryPrefix + "/0/Auto ongeluk", "/alice@example.com/Auto ongeluk")]
    // Store root (no store-type segment).
    [InlineData(PrimaryPrefix, "/alice@example.com")]
    // Delegate: the leading name is the HOST account, never the delegate's own store
    // display name - building it from Store.DisplayName returned 0 rows on 15/15 folders.
    [InlineData(DelegatePrefix, "/alice@example.com/Sam Delegate")]
    [InlineData(DelegatePrefix + "/20251015", "/alice@example.com/Sam Delegate/20251015")]
    // Trailing slash tolerated.
    [InlineData(PrimaryPrefix + "/0/Inbox/", "/alice@example.com/Inbox")]
    public void FolderPathDisplay_IsAPureStringTransformOfTheUrl(string url, string expected)
    {
        Assert.True(MapiItemUrl.TryBuildFolderPathDisplay(url, out string? path));
        Assert.Equal(expected, path);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("file:///c:/temp")]
    [InlineData("csc://x/y")]
    public void FolderPathDisplay_RejectsNonMapiInput(string? url)
    {
        Assert.False(MapiItemUrl.TryBuildFolderPathDisplay(url, out string? path));
        Assert.Null(path);
    }

    // --------------------------------------------------------------- primary stores

    [Fact]
    public void PrimaryStore_WithoutFolder_IsTheBareStoreScope()
    {
        FolderScopeResolution r = FolderScopeResolver.ForPrimaryStore(PrimaryPrefix, folder: null, includeSubfolders: true);

        Assert.Equal(FolderScopeKind.WholeStore, r.Kind);
        Assert.Equal(PrimaryPrefix, r.Scope);
        Assert.Null(r.FolderPaths);
        Assert.False(r.IsDelegateStore);
        Assert.Null(r.RequestedFolder);
    }

    [Fact]
    public void PrimaryStore_IncludeSubfolders_KeepsTodaysRecursiveScope_Unchanged()
    {
        FolderScopeResolution r = FolderScopeResolver.ForPrimaryStore(PrimaryPrefix, "Inbox/Fun", includeSubfolders: true);

        Assert.Equal(FolderScopeKind.PrimaryRecursive, r.Kind);
        Assert.Equal(PrimaryPrefix + "/0/Inbox/Fun", r.Scope);
        // No extra predicate: SCOPE= is already recursive, so existing callers see the
        // byte-identical query they saw before the flag existed.
        Assert.Null(r.FolderPaths);
        Assert.Equal(PrimaryPrefix, r.StoreScope);
    }

    [Fact]
    public void PrimaryStore_ExcludeSubfolders_AddsOneFolderPathEquality()
    {
        FolderScopeResolution r = FolderScopeResolver.ForPrimaryStore(PrimaryPrefix, "Inbox/Fun", includeSubfolders: false);

        Assert.Equal(FolderScopeKind.PrimaryNonRecursive, r.Kind);
        Assert.Equal(PrimaryPrefix + "/0/Inbox/Fun", r.Scope);
        Assert.NotNull(r.FolderPaths);
        Assert.Equal(new[] { "/alice@example.com/Inbox/Fun" }, r.FolderPaths!);
    }

    [Theory]
    [InlineData("/Inbox")]
    [InlineData("Inbox/")]
    [InlineData("  Inbox  ")]
    public void PrimaryStore_NormalizesTheFolderPath(string folder)
    {
        FolderScopeResolution r = FolderScopeResolver.ForPrimaryStore(PrimaryPrefix, folder, includeSubfolders: false);

        Assert.Equal(PrimaryPrefix + "/0/Inbox", r.Scope);
        Assert.Equal(new[] { "/alice@example.com/Inbox" }, r.FolderPaths!);
        Assert.Equal("Inbox", r.RequestedFolder);
    }

    [Fact]
    public void PrimaryStore_ApostropheInAFolderName_ResolvesInsteadOfThrowing()
    {
        FolderScopeResolution r = FolderScopeResolver.ForPrimaryStore(PrimaryPrefix, "Clients/O'Brien", includeSubfolders: false);

        Assert.Equal(PrimaryPrefix + "/0/Clients/O'Brien", r.Scope);
        Assert.Equal(new[] { "/alice@example.com/Clients/O'Brien" }, r.FolderPaths!);
    }

    // -------------------------------------------------------------- delegate stores

    [Fact]
    public void DelegateStore_WithoutFolder_ScopesTheDelegateRoot_WithNoFilter()
    {
        // The delegate root scope is recursive over the FLAT namespace, so it already
        // covers every folder of that mailbox - a filter would only narrow it.
        foreach (bool includeSubfolders in new[] { true, false })
        {
            FolderScopeResolution r = FolderScopeResolver.ForDelegateStore(
                DelegatePrefix, folder: null, includeSubfolders, comFolderPaths: new[] { "Inbox", "Inbox/Old" });

            Assert.Equal(FolderScopeKind.WholeStore, r.Kind);
            Assert.Equal(DelegatePrefix, r.Scope);
            Assert.Null(r.FolderPaths);
            Assert.True(r.IsDelegateStore);
        }
    }

    [Fact]
    public void DelegateStore_NeverBuildsANestedUrl_TheDefectThisFixExists_For()
    {
        FolderScopeResolution r = FolderScopeResolver.ForDelegateStore(
            DelegatePrefix, "Inbox/20251015", includeSubfolders: false, comFolderPaths: new[] { "Inbox", "Inbox/20251015" });

        // The old (broken) shape was DelegatePrefix + "/Inbox/20251015" - a URL that does
        // not exist in the index, returning zero rows with no error.
        Assert.Equal(DelegatePrefix, r.Scope);
        Assert.DoesNotContain("/Inbox/20251015", r.Scope, StringComparison.Ordinal);

        // The shipped shape: delegate STORE ROOT scope + the FLAT leaf path.
        Assert.Equal(FolderScopeKind.DelegateFlat, r.Kind);
        Assert.Equal(new[] { "/alice@example.com/Sam Delegate/20251015" }, r.FolderPaths!);
    }

    [Fact]
    public void DelegateStore_IncludeSubfolders_OrsTheSubtreesLeafNames()
    {
        string[] tree =
        {
            "Inbox", "Inbox/20251015", "Inbox/tutanota deleted items", "Inbox/20251015/deeper",
            "Archive", "Sent Items",
        };

        FolderScopeResolution r = FolderScopeResolver.ForDelegateStore(
            DelegatePrefix, "Inbox", includeSubfolders: true, comFolderPaths: tree);

        Assert.Equal(FolderScopeKind.DelegateFlat, r.Kind);
        Assert.Equal(DelegatePrefix, r.Scope);
        Assert.Equal(
            new[]
            {
                "/alice@example.com/Sam Delegate/Inbox",
                "/alice@example.com/Sam Delegate/20251015",
                "/alice@example.com/Sam Delegate/tutanota deleted items",
                "/alice@example.com/Sam Delegate/deeper",
            },
            r.FolderPaths!);

        // Folders outside the requested subtree stay out.
        Assert.DoesNotContain("/alice@example.com/Sam Delegate/Archive", r.FolderPaths!);
        Assert.False(r.Widened);
    }

    [Fact]
    public void DelegateStore_ExcludeSubfolders_MatchesTheLeafAlone()
    {
        string[] tree = { "Inbox", "Inbox/20251015", "Inbox/tutanota deleted items" };

        FolderScopeResolution r = FolderScopeResolver.ForDelegateStore(
            DelegatePrefix, "Inbox", includeSubfolders: false, comFolderPaths: tree);

        Assert.Equal(new[] { "/alice@example.com/Sam Delegate/Inbox" }, r.FolderPaths!);
    }

    [Fact]
    public void DelegateStore_WithoutAFolderTree_WidensRatherThanNarrowingOnAGuess()
    {
        // Outlook unreachable: the subtree's leaf names are unknowable. Narrowing to the
        // single leaf would silently drop the subfolders that were explicitly asked for,
        // so the query widens to the whole delegate mailbox - a SUPERSET - and says so.
        FolderScopeResolution r = FolderScopeResolver.ForDelegateStore(
            DelegatePrefix, "Inbox", includeSubfolders: true, comFolderPaths: null);

        Assert.Equal(FolderScopeKind.DelegateWidened, r.Kind);
        Assert.True(r.Widened);
        Assert.Null(r.FolderPaths);
        Assert.Equal(DelegatePrefix, r.Scope);
        Assert.True(r.FolderTreeUnavailable);

        string advice = FolderScopeResolver.DescribeWidening(r);
        Assert.Contains("WIDENED", advice, StringComparison.Ordinal);
        Assert.Contains("Outlook could not be reached", advice, StringComparison.Ordinal);
        Assert.Contains("include_subfolders:false", advice, StringComparison.Ordinal);
    }

    [Fact]
    public void DelegateStore_WithoutAFolderTree_StillNarrowsWhenSubfoldersAreExcluded()
    {
        // A single-folder request needs no tree: the leaf path is derivable from the
        // request itself, so this case must NOT widen.
        FolderScopeResolution r = FolderScopeResolver.ForDelegateStore(
            DelegatePrefix, "Inbox/20251015", includeSubfolders: false, comFolderPaths: null);

        Assert.Equal(FolderScopeKind.DelegateFlat, r.Kind);
        Assert.False(r.Widened);
        Assert.Equal(new[] { "/alice@example.com/Sam Delegate/20251015" }, r.FolderPaths!);
        Assert.True(r.FolderTreeUnavailable);
    }

    [Fact]
    public void DelegateStore_OverlongSubtree_WidensInsteadOfEmittingADoomedOrSet()
    {
        // MEASURED: the provider FAILS OUTRIGHT near 100 OR literals, so an uncapped
        // OR-set is a crash rather than a slowdown.
        Assert.Equal(40, FolderScopeResolver.DelegateFolderOrSetCap);

        List<string> tree = new() { "Big" };
        for (int i = 0; i < FolderScopeResolver.DelegateFolderOrSetCap + 5; i++)
        {
            tree.Add("Big/child" + i);
        }

        FolderScopeResolution r = FolderScopeResolver.ForDelegateStore(
            DelegatePrefix, "Big", includeSubfolders: true, comFolderPaths: tree);

        Assert.Equal(FolderScopeKind.DelegateWidened, r.Kind);
        Assert.Null(r.FolderPaths);
        Assert.False(r.FolderTreeUnavailable);

        string advice = FolderScopeResolver.DescribeWidening(r);
        Assert.Contains("more than 40 folders", advice, StringComparison.Ordinal);
    }

    [Fact]
    public void DelegateStore_ExactlyAtTheCap_StillNarrows()
    {
        List<string> tree = new() { "Big" };
        for (int i = 0; i < FolderScopeResolver.DelegateFolderOrSetCap - 1; i++)
        {
            tree.Add("Big/child" + i);
        }

        FolderScopeResolution r = FolderScopeResolver.ForDelegateStore(
            DelegatePrefix, "Big", includeSubfolders: true, comFolderPaths: tree);

        Assert.Equal(FolderScopeKind.DelegateFlat, r.Kind);
        Assert.Equal(FolderScopeResolver.DelegateFolderOrSetCap, r.FolderPaths!.Count);
    }

    [Fact]
    public void DelegateStore_LeafNameCollision_IsDetectedAndReported()
    {
        // Measured on the discovery machine: two distinct COM folders
        // (Synchronisatieproblemen/Conflicten and Overig/Synchronisatieproblemen/Conflicten)
        // both land on ONE flat index folder, and no index column can separate them - so
        // the answer over-returns. That must never be silent (constraint C3).
        string[] tree =
        {
            "Sync", "Sync/Conflicts", "Other", "Other/Sync", "Other/Sync/Conflicts",
        };

        FolderScopeResolution r = FolderScopeResolver.ForDelegateStore(
            DelegatePrefix, "Sync", includeSubfolders: true, comFolderPaths: tree);

        // BOTH selected names collide here: the requested folder's own leaf ("Sync" also
        // exists under Other) and its child's ("Conflicts" likewise) - so the query
        // returns four COM folders' mail while asking for two.
        Assert.NotNull(r.CollidingLeafNames);
        Assert.Equal(new[] { "Sync", "Conflicts" }, r.CollidingLeafNames!);

        string advice = FolderScopeResolver.DescribeCollision(r);
        Assert.Contains("'Sync'", advice, StringComparison.Ordinal);
        Assert.Contains("'Conflicts'", advice, StringComparison.Ordinal);
        Assert.Contains("over-return", advice, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("exhaustive:true", advice, StringComparison.Ordinal);
    }

    [Fact]
    public void DelegateStore_CollisionOnTheRequestedFolderItself_IsReportedEvenWithoutSubfolders()
    {
        string[] tree = { "Sync", "Other", "Other/Sync" };

        FolderScopeResolution r = FolderScopeResolver.ForDelegateStore(
            DelegatePrefix, "Sync", includeSubfolders: false, comFolderPaths: tree);

        Assert.Equal(new[] { "Sync" }, r.CollidingLeafNames!);
    }

    [Fact]
    public void DelegateStore_NoCollision_ReportsNothing()
    {
        string[] tree = { "Inbox", "Inbox/2025", "Archive", "Archive/2024" };

        FolderScopeResolution r = FolderScopeResolver.ForDelegateStore(
            DelegatePrefix, "Inbox", includeSubfolders: true, comFolderPaths: tree);

        Assert.Null(r.CollidingLeafNames);
    }

    [Fact]
    public void DelegateStore_DuplicateLeavesInsideOneSubtree_CollapseToOneEquality()
    {
        // The flat namespace merges them anyway; emitting the literal twice would only
        // pad the OR-set toward the provider's hard limit.
        string[] tree = { "Root", "Root/a/Conflicts", "Root/b/Conflicts", "Root/a", "Root/b" };

        FolderScopeResolution r = FolderScopeResolver.ForDelegateStore(
            DelegatePrefix, "Root", includeSubfolders: true, comFolderPaths: tree);

        Assert.Single(r.FolderPaths!, p => p.EndsWith("/Conflicts", StringComparison.Ordinal));
        Assert.Equal(new[] { "Conflicts" }, r.CollidingLeafNames!);
    }

    [Fact]
    public void UnresolvedFolderAdvice_NamesTheResolutionProblem_NotAnEmptyFolder()
    {
        string advice = FolderScopeResolver.DescribeUnresolvedFolder("Inbox/Typo", "Sam Delegate");

        Assert.Contains("matched NOTHING in the index", advice, StringComparison.Ordinal);
        Assert.Contains("not an empty folder", advice, StringComparison.Ordinal);
        Assert.Contains("list_folders", advice, StringComparison.Ordinal);
        Assert.Contains("exhaustive:true", advice, StringComparison.Ordinal);
    }
}
