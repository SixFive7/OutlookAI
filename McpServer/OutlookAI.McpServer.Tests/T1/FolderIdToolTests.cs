using System.Text.Json;
using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The pure halves of folders and stores addressed by ID (Q131-Q136): the index scope a folder's own
/// row gives (Q132 (i)), find_folder's candidates (Q136 (ii)), create_folder's rules and refusals
/// (Q134 (i)), the store an ID names, and the refusal of every argument this release retired
/// (Q135 (i)). Every id here is synthetic.
/// </summary>
public sealed class FolderIdToolTests
{
    private const string Prefix = "mapi16://{S-1-5-21-1}/tier@vm.invalid($1234abcd)";

    // ------------------------------------------------------------------ the scope from a folder's own index row

    [Fact]
    public void AFolderRow_IsTheRecursiveScope_AndItsNamesAreDecoded()
    {
        FolderScopeResolution scope = FolderScopeResolver.ForIndexedFolder(Prefix, Prefix + "/0/Projects/50%25 off", includeSubfolders: true);

        Assert.Equal(FolderScopeKind.PrimaryRecursive, scope.Kind);
        Assert.Equal(Prefix + "/0/Projects/50%25 off", scope.Scope);
        Assert.Equal("Projects/50% off", scope.RequestedFolder);
        Assert.Equal(Prefix, scope.StoreScope);
    }

    [Fact]
    public void AFolderRow_NarrowedToTheFolderAlone_UsesItsDisplayPath()
    {
        FolderScopeResolution scope = FolderScopeResolver.ForIndexedFolder(Prefix, Prefix + "/0/Projects/a%2Fb", includeSubfolders: false);

        Assert.Equal(FolderScopeKind.PrimaryNonRecursive, scope.Kind);
        Assert.Equal(Prefix + "/0/Projects/a%2Fb", scope.Scope);
        Assert.Equal("/tier@vm.invalid/Projects/a/b", Assert.Single(scope.FolderPaths!));
    }

    [Fact]
    public void AFolderWhoseNameEndsWithASpace_KeepsItsSeparator_OrIsBoundByItsParent()
    {
        FolderScopeResolution recursive = FolderScopeResolver.ForIndexedFolder(Prefix, Prefix + "/0/Inbox/Fun ", includeSubfolders: true);
        Assert.Equal(Prefix + "/0/Inbox/Fun /", recursive.Scope);

        FolderScopeResolution only = FolderScopeResolver.ForIndexedFolder(Prefix, Prefix + "/0/Inbox/Fun ", includeSubfolders: false);
        Assert.Equal(Prefix + "/0/Inbox", only.Scope);
        Assert.Equal("/tier@vm.invalid/Inbox/Fun ", Assert.Single(only.FolderPaths!));
    }

    [Fact]
    public void ARowOfAnotherStore_IsRefused()
    {
        Assert.Throws<ArgumentException>(() => FolderScopeResolver.ForIndexedFolder(
            Prefix, "mapi16://{S-1-5-21-1}/other($1)/0/Inbox", includeSubfolders: true));
        Assert.Throws<ArgumentException>(() => FolderScopeResolver.ForIndexedFolder(Prefix, Prefix, includeSubfolders: true));
    }

    [Fact]
    public void FolderNamesFromOutlook_AreEncodedOneNameAtATime_SoASlashStaysInItsName()
    {
        FolderScopeResolution scope = FolderScopeResolver.ForPrimaryStoreSegments(Prefix, new[] { "Projects", "Q3/Q4" }, includeSubfolders: true);

        Assert.Equal(Prefix + "/0/Projects/Q3%2FQ4", scope.Scope);
        Assert.Equal("Projects/Q3/Q4", scope.RequestedFolder);
        Assert.Equal(FolderScopeKind.WholeStore, FolderScopeResolver.ForPrimaryStoreSegments(Prefix, null, true).Kind);
    }

    // ------------------------------------------------------------------ find_folder

    private static ComFolderTree Tree(params (string Store, string Path)[] folders)
    {
        return new ComFolderTree(
            folders.Select(f => new ComFolderInfo(
                f.Store, f.Path, f.Path.Split('/')[^1], 3, 1, 0, TestFolders.FolderEntryId(f.Store, f.Path))).ToList(),
            stores: folders.Select(f => f.Store).Distinct()
                .Select(s => new ComStoreRoot(s, "store-" + s, TestFolders.RootEntryId(s))).ToList());
    }

    [Fact]
    public void FindFolder_OffersExactThenPathThenPartialMatches_EachWithItsIdAndStore()
    {
        ComFolderTree tree = Tree(
            ("Work", "Projects"),
            ("Work", "Projects/2026"),
            ("Work", "Archive/Projects old"),
            ("Home", "projects"));

        FindFolderOutcome found = MailService.MatchFolders(tree, " Projects ", null);

        Assert.Equal("Projects", found.Name);
        Assert.Equal(new[] { "exact", "exact", "partial" }, found.Candidates.Select(c => c.Match).ToArray());
        Assert.Equal(new[] { "Projects", "projects", "Archive/Projects old" }, found.Candidates.Select(c => c.Folder).ToArray());
        FolderCandidateView first = found.Candidates[0];
        Assert.Equal(TestFolders.FolderId("Work", "Projects"), first.FolderId);
        Assert.Equal(TestFolders.StoreId("Work"), first.StoreId);
        Assert.Equal("Work", first.Store);
        Assert.Equal(TestFolders.StoreId("Home"), found.Candidates[1].StoreId);
        Assert.Contains(found.Advice!, a => a.Contains("ask the user", StringComparison.OrdinalIgnoreCase));

        FindFolderOutcome byPath = MailService.MatchFolders(tree, "Projects/2026", null);
        FolderCandidateView pathMatch = Assert.Single(byPath.Candidates);
        Assert.Equal("path", pathMatch.Match);
        Assert.Equal(TestFolders.FolderId("Work", "Projects/2026"), pathMatch.FolderId);
        Assert.Null(byPath.Advice);
    }

    [Fact]
    public void FindFolder_SaysWhenNothingMatches_AndCapsALongAnswerKeepingTheTotal()
    {
        FindFolderOutcome none = MailService.MatchFolders(Tree(("Work", "Inbox")), "Receipts", null);
        Assert.Empty(none.Candidates);
        Assert.Equal(0, none.Total);
        Assert.Contains(none.Advice!, a => a.Contains("create_folder", StringComparison.Ordinal));

        ComFolderTree many = Tree(Enumerable.Range(0, 12).Select(i => ("Work", "Box " + i)).ToArray());
        FindFolderOutcome capped = MailService.MatchFolders(many, "Box", null, cap: 5);
        Assert.Equal(5, capped.Candidates.Count);
        Assert.Equal(12, capped.Total);
        Assert.True(capped.Truncated);
        Assert.Throws<ArgumentException>(() => MailService.MatchFolders(many, "  ", null));
    }

    [Fact]
    public void AStoreTheProfileListsWithoutItsRoot_IsShownTheWayListAccountsShowsIt()
    {
        ComFolderTree tree = new ComFolderTree(
            new[] { new ComFolderInfo("Archive", "Old", "Old", 1, 0, 0, TestFolders.FolderEntryId("Archive", "Old")) },
            stores: new[] { new ComStoreRoot("Archive", "0000000038A1BB1005E5101AA1BB08002B2A56C2AABBCCDD", TestFolders.RootEntryId("Archive")) });
        ComStoreDetail uncached = new ComStoreDetail("Archive", "0000000038A1BB1005E5101AA1BB08002B2A56C2AABBCCDD", 0, false);

        FindFolderOutcome found = MailService.MatchFolders(tree, "Old", new[] { uncached });

        string shown = Assert.Single(found.Candidates).StoreId!;
        Assert.Equal(MailService.StoreIdOf(uncached), shown);
        Assert.Equal('T', shown[0]);
    }

    [Fact]
    public void TwoFoldersOfOneNameSideBySide_AreBothOffered_EachByItsOwnId_AndNamedAsSharingAPath()
    {
        // Measured on the Exchange guest: two top-level "Conversation Action Settings" folders, one index
        // row for the path. find_folder must offer both, never merge them by path.
        const string Name = "Conversation Action Settings";
        ComFolderTree tree = new ComFolderTree(
            new[]
            {
                new ComFolderInfo("Work", Name, Name, 0, 0, 0, TestFolders.FolderEntryId("Work", Name + "#1")),
                new ComFolderInfo("Work", Name, Name, 0, 0, 0, TestFolders.FolderEntryId("Work", Name + "#2")),
            },
            stores: new[] { new ComStoreRoot("Work", "store-work", TestFolders.RootEntryId("Work")) });

        FindFolderOutcome found = MailService.MatchFolders(tree, Name, null);

        Assert.Equal(2, found.Candidates.Count);
        Assert.NotEqual(found.Candidates[0].FolderId, found.Candidates[1].FolderId);
        Assert.Contains(found.Advice!, a => a.Contains("side by side", StringComparison.Ordinal)
            && a.Contains("'" + Name + "' in 'Work'", StringComparison.Ordinal));
        Assert.Contains("exhaustive:true", MailService.DescribeSharedFolderPath(Name), StringComparison.Ordinal);
    }

    [Fact]
    public void SeveralIndexStoresOfOneName_AreNamedInTheRefusal_WithTheOstRebuildCause()
    {
        // Measured on the Exchange guest: after an .ost rebuild the index keeps the old copy as a second root.
        StoreScopeInfo stale = StoreScopeInfo.FromStorePrefix("mapi16://{S-1-5-21-1}/someone@example.com($9b7b502)")!;
        StoreScopeInfo live = StoreScopeInfo.FromStorePrefix("mapi16://{S-1-5-21-1}/someone@example.com($b37770e)")!;

        string refusal = MailService.DescribeSeveralIndexStoresNamed("someone@example.com", new[] { stale, live });

        Assert.Contains("2 stores named 'someone@example.com'", refusal, StringComparison.Ordinal);
        Assert.Contains("($9b7b502)", refusal, StringComparison.Ordinal);
        Assert.Contains("($b37770e)", refusal, StringComparison.Ordinal);
        Assert.Contains(".ost", refusal, StringComparison.Ordinal);
        Assert.Contains("exhaustive:true", refusal, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ the store an ID names

    [Fact]
    public void PickStore_FindsTheOneStore_RefusesNoneAndSeveral()
    {
        ComStoreDetail work = TestFolders.WithRoot(new ComStoreDetail("Work", "store-work", 0, true));
        ComStoreDetail home = TestFolders.WithRoot(new ComStoreDetail("Home", "store-home", 3, null));
        Assert.True(OutlookIds.TryDecodeStoreId(TestFolders.StoreId("Work"), "store_id", out StoreIdParts? parts, out _));

        Assert.Same(work, MailService.PickStore(new[] { work, home }, parts!, "store_id"));

        ArgumentException none = Assert.Throws<ArgumentException>(() => MailService.PickStore(new[] { home }, parts!, "store_id"));
        Assert.Contains("No store in this Outlook profile has the ID", none.Message, StringComparison.Ordinal);
        Assert.Contains("'Home' (" + TestFolders.StoreId("Home") + ")", none.Message, StringComparison.Ordinal);

        ComStoreDetail twin = new ComStoreDetail("Work (shared)", "store-twin", 1, true, rootFolderEntryId: work.RootFolderEntryId);
        ArgumentException several = Assert.Throws<ArgumentException>(() => MailService.PickStore(new[] { work, twin }, parts!, "store_id"));
        Assert.Contains("2 stores", several.Message, StringComparison.Ordinal);
        Assert.Contains("folderId", several.Message, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ create_folder's rules

    [Theory]
    [InlineData(null)]
    [InlineData("   ")]
    [InlineData("tab\tname")]
    public void ANameThatCannotBeAFolderName_IsRefused(string? name)
    {
        Assert.Throws<ArgumentException>(() => MailService.ValidateFolderName(name));
    }

    [Fact]
    public void ASlashIsPartOfAName_AndTheLongestNameIsAllowed()
    {
        MailService.ValidateFolderName("Q3/Q4");
        MailService.ValidateFolderName(new string('x', MailService.FolderNameCharsCap));
        Assert.Throws<ArgumentException>(() => MailService.ValidateFolderName(new string('x', MailService.FolderNameCharsCap + 1)));
    }

    [Theory]
    [InlineData("TargetIsDeletedItems", "unchanged", "Deleted Items")]
    [InlineData("TargetIsOutbox", "unchanged", "Outbox")]
    [InlineData("TargetNotAMailFolder", "unchanged", "not a mail folder")]
    [InlineData("ParentFolderNotFound", "unchanged", "find_folder")]
    [InlineData("CreateFailed:COMException 0x80004005", "unchanged", "no folder of that name is there")]
    [InlineData("COMException 0x80004005", "unknown", "find_folder")]
    public void EveryCreateFolderFailure_StatesItsOutcome(string error, string outcome, string fragment)
    {
        OperationOutcomeException failure = MailService.DescribeCreateFolderFailure(error, TestFolders.FolderId("Work", "Inbox"), "New");
        Assert.Equal(outcome, failure.Outcome);
        Assert.Contains(fragment, failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AFolderThatExists_IsRefused_NamingItsId_AndOneMadeButUnreadable_IsApplied()
    {
        string existing = TestFolders.FolderEntryId("Work", "Projects");
        OperationOutcomeException exists = MailService.DescribeCreateFolderFailure(
            "FolderExists:" + existing, TestFolders.FolderId("Work", "Inbox"), "Projects");
        Assert.Equal(MutationOutcome.Unchanged, exists.Outcome);
        Assert.Contains(OutlookIds.FolderIdFromEntryId(existing)!, exists.Message, StringComparison.Ordinal);

        OperationOutcomeException made = MailService.DescribeCreateFolderFailure(
            "CreatedUnreadable:" + existing, TestFolders.FolderId("Work", "Inbox"), "Projects");
        Assert.Equal(MutationOutcome.Applied, made.Outcome);
        Assert.Contains("Do not create it again", made.Message, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ the retired arguments (Q135 (i))

    private static IDictionary<string, JsonElement> Arguments(object values)
    {
        return JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(JsonSerializer.Serialize(values))!;
    }

    [Theory]
    [InlineData("search", "store")]
    [InlineData("search", "folder")]
    [InlineData("thread", "store")]
    [InlineData("move_mail", "folder")]
    [InlineData("move_mail", "create_folder")]
    [InlineData("move_mail", "store")]
    [InlineData("list_folders", "store")]
    [InlineData("goto_folder", "store")]
    [InlineData("goto_folder", "folder")]
    [InlineData("show_search_results", "store")]
    [InlineData("show_search_results", "folder")]
    public void EveryRetiredArgument_IsRefused_NamingWhatReplacedIt(string tool, string argument)
    {
        string? refusal = RetiredArguments.Refusal(tool, Arguments(new Dictionary<string, object> { [argument] = "Projects/2026" }));

        Assert.NotNull(refusal);
        Assert.StartsWith(tool + " no longer takes '" + argument + "'", refusal, StringComparison.Ordinal);
        Assert.Contains(argument == "store" ? "store_id" : argument == "folder" ? "folder_id" : "create_folder(parent_id, name)", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void AFolderPassedTheOldWay_IsToldWhichNameToAskFindFolderFor()
    {
        string refusal = RetiredArguments.Refusal("search", Arguments(new { query = "x", folder = "Clients/Acme" }))!;
        Assert.Contains("find_folder", refusal, StringComparison.Ordinal);
        Assert.Contains("\"Acme\"", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void TheNewArguments_AndOtherToolsAndNulls_PassUntouched()
    {
        Assert.Null(RetiredArguments.Refusal("search", Arguments(new { store_id = "S1", folder_id = "P1" })));
        Assert.Null(RetiredArguments.Refusal("read", Arguments(new { store = "x", folder = "y" })));
        Assert.Null(RetiredArguments.Refusal("search", Arguments(new Dictionary<string, object?> { ["store"] = null })));
        Assert.Null(RetiredArguments.Refusal(null, null));
    }
}
