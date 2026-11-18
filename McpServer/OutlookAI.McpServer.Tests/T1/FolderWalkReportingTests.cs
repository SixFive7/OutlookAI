using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The store list and the folder walk are not allowed to lose things quietly - gaps G2, G3
/// and G4 of <c>Docs/completeness-gaps.md</c>, which are one subject: everything in this
/// server is keyed by a store DISPLAY NAME and bounded by one folder walk, so a store that
/// cannot be named and a walk that stopped early both went missing without a field to say so.
/// <para>
/// WHAT THE THREE LOOKED LIKE BEFORE. <b>G2</b>: a store whose <c>DisplayName</c> read threw
/// was dropped from <c>list_folders</c> outright - its whole folder tree simply absent - and
/// in the freshness sweep its four default folders were added to the SKIPPED total and to no
/// per-store bucket, so an unscoped search lost that store's fresh mail with nothing
/// attributing the loss. <b>G3</b>: <c>CollectFolders</c> / <c>CollectFolderPaths</c> stopped
/// at the 10 000-folder walk cap and at depth 64 with no flag, and <c>FoldersOutcome.Truncated</c>
/// was computed as "is this page short of the LIST" - a list the walk had already truncated -
/// so the only truncation it could not see was the one that lost folders rather than deferring
/// them. <b>G4</b>: the delegate folder-NAME list comes from that same walk, and the delegate
/// index namespace is FLAT, so a short list is not a short listing but folders no tier looks
/// in; <c>scope.folderNamesMatched</c> is a count and reads identically either way.
/// </para>
/// <para>
/// Everything below drives the REAL <see cref="MailService"/> - its paging, its scope
/// resolver, its advice and its sweep attribution - through a stand-in Outlook session and a
/// stand-in index client. No mailbox and no Windows Search index are touched. The one half
/// that cannot be reached from here is the COM read that FAILS in the first place: producing
/// a store whose <c>DisplayName</c> throws, a profile of 10 000 folders or a cyclic folder
/// tree needs a live profile. So the COM layer's job is to hand up the label and the bounds,
/// and these tests own everything after that - which is where the whole defect lived, since
/// the drops were all decisions taken above the COM call.
/// </para>
/// </summary>
public sealed class FolderWalkReportingTests
{
    private const string Sid = "{S-1-5-21-1111111111-2222222222-3333333333-1001}";
    private const string OwnerStore = "alice@example.com";
    private const string OwnerPrefix = "mapi16://" + Sid + "/" + OwnerStore + "($deadbeef)";
    private const string DelegateStore = "Shared Mailbox";
    private const string NamedStore = "Store A";

    private static readonly DateTime Frontier = new(2026, 8, 18, 9, 30, 0, DateTimeKind.Utc);

    /// <summary>The label the COM layer gives the second store of a profile when it will not name it.</summary>
    private static readonly string Unnamed = StoreNaming.LabelForUnnamedStore(2);

    // ============================================================ G2: naming the unnameable

    [Fact]
    public void AnUnnameableStore_IsListed_UnderALabelThatSaysWhatItIs()
    {
        // The defect: this store's entire folder tree was absent from list_folders, and the
        // payload said nothing at all - not a count, not a flag. It is listed now, and the
        // label travels with the one fact that keeps it honest.
        using MailService service = Service(Tree(
            new[] { Folder(NamedStore, "Inbox"), Folder(Unnamed, "Inbox"), Folder(Unnamed, "Archive") },
            storesUnnamed: 1));

        FoldersOutcome outcome = service.ListFolders();

        StoreFoldersView labelled = outcome.Stores.Single(s => s.Store == Unnamed);
        Assert.True(labelled.NameUnreadable);
        Assert.Equal(2, labelled.Folders.Count);
        Assert.Null(outcome.Stores.Single(s => s.Store == NamedStore).NameUnreadable);
        Assert.Equal(1, outcome.StoresUnnamed);
        Assert.Equal(3, outcome.FolderTotal);
    }

    [Fact]
    public void TheLabel_IsAdvertisedAsOnlyForReading_AndTheStoreIdAsTheAddress()
    {
        // Putting the label in the payload creates exactly one wrong turn - an agent reads it as a
        // store name and passes it back. Since Q133 a store is addressed by its storeId, which a
        // store without a name has like any other, so the answer that prints the label says the
        // label is only for reading and the storeId is what to pass.
        using MailService service = Service(Tree(new[] { Folder(Unnamed, "Inbox") }, storesUnnamed: 1));

        FoldersOutcome outcome = service.ListFolders();

        Assert.Contains(
            outcome.Advice!,
            a => a.Contains("would not report a display name", StringComparison.Ordinal)
                && a.Contains("storeId addresses it", StringComparison.Ordinal)
                && a.Contains("only something to read", StringComparison.Ordinal));
    }

    [Fact]
    public void AStoreIdNoStoreHas_IsRefused_NamingEveryStoreWithItsId_EvenOneWithoutAName()
    {
        // Gaps G1 and G2 were about names: a typo answered as an empty tree, and a store that would
        // not say its name could be neither matched nor ruled out. Stores are addressed by ID now
        // (Q133), so an unknown ID is simply refused - and the unnamed store is listed with its own ID.
        using MailService service = Service(Tree(new[] { Folder(NamedStore, "Inbox") }));

        ArgumentException refused = Assert.Throws<ArgumentException>(() => service.ListFolders(TestFolders.StoreId("typo")));

        Assert.Contains("No store in this Outlook profile has the ID", refused.Message, StringComparison.Ordinal);
        Assert.Contains("'" + NamedStore + "' (" + TestFolders.StoreId(NamedStore) + ")", refused.Message, StringComparison.Ordinal);
        Assert.Contains("(" + TestFolders.StoreId(Unnamed) + ")", refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AStoreScopedListing_AddressesAStoreWithoutAName_ByItsId()
    {
        // A scope by NAME had to leave such a store out - its name could not be read to match. Its ID can.
        using MailService service = Service(Tree(new[] { Folder(Unnamed, "Inbox") }, storesUnnamed: 1));

        FoldersOutcome outcome = service.ListFolders(TestFolders.StoreId(Unnamed));

        StoreFoldersView store = Assert.Single(outcome.Stores);
        Assert.True(store.NameUnreadable);
        Assert.Contains(outcome.Advice!, a => a.Contains("storeId addresses it", StringComparison.Ordinal));
    }

    [Fact]
    public void TheSweep_AttributesAnUnnameableStore_AndSaysSoWithoutCryingWolf()
    {
        // The sweep half of G2. Those four folders now land in a bucket and the store is
        // SWEPT, so there is no coverage hole to raise - the loss is that the caller cannot
        // scope a follow-up to it, which is a fact about the next call, not this one.
        SweepInfo info = new SweepInfo { Performed = true };
        MailService.ApplySweepCounters(info, SweepWithUnnamedStore(), store: null);

        Assert.Equal(1, info.StoresUnnamed);
        Assert.Equal(4, info.FoldersSwept);
        Assert.Equal(0, info.FoldersSkipped);
        Assert.Null(FreshMerge.DescribeCoverageGaps(info));
    }

    [Fact]
    public void AStoreScopedSearch_DoesNotInheritAnotherStoresNamingFailure()
    {
        // The sweep cache can serve an all-stores sweep to a store-scoped request, which is
        // how a counter leaks across accounts. A SCOPED sweep never reaches an unnameable
        // store at all, so a non-zero count can only have come from the wider one.
        SweepInfo info = new SweepInfo { Performed = true };
        MailService.ApplySweepCounters(info, SweepWithUnnamedStore(), store: NamedStore);

        Assert.Null(info.StoresUnnamed);
    }

    [Fact]
    public void TheUnnamedStoreSentence_NamesTheLabelAndTheConsequence()
    {
        Assert.Null(MailService.DescribeUnnamedStores(null));
        Assert.Null(MailService.DescribeUnnamedStores(0));

        string sentence = MailService.DescribeUnnamedStores(2)!;
        Assert.Contains("2 store(s)", sentence, StringComparison.Ordinal);
        Assert.Contains(StoreNaming.UnnamedStorePrefix, sentence, StringComparison.Ordinal);

        // The consequence since Q133: the label is not an address, and list_accounts gives the one that is.
        Assert.Contains("list_accounts gives each its storeId", sentence, StringComparison.Ordinal);
    }

    [Fact]
    public void TheLabel_IsStablePerProfilePosition_AndRecognisable()
    {
        Assert.Equal("(unnamed store 1)", StoreNaming.LabelForUnnamedStore(1));
        Assert.Equal("(unnamed store 17)", StoreNaming.LabelForUnnamedStore(17));
        Assert.True(StoreNaming.IsUnnamedStoreLabel(StoreNaming.LabelForUnnamedStore(3)));
        Assert.False(StoreNaming.IsUnnamedStoreLabel("Outlook Data File"));
        Assert.False(StoreNaming.IsUnnamedStoreLabel(null));

        // 1-based, because Namespace.Stores is - an off-by-one here would label two different
        // stores the same on two different code paths.
        Assert.Throws<ArgumentOutOfRangeException>(() => StoreNaming.LabelForUnnamedStore(0));
    }

    // ====================================================== G3: a walk that stopped early

    [Fact]
    public void AWalkStoppedByItsCap_ReportsTruncated_WhereItUsedToReportComplete()
    {
        // THE DEFECT, exactly: the tree is short, the page is not, and `truncated` was
        // computed against the already-truncated list - so it answered false and the listing
        // read as the whole profile.
        using MailService service = Service(Tree(new[] { Folder(NamedStore, "Inbox") }, walkCapReached: true));

        FoldersOutcome outcome = service.ListFolders();

        Assert.True(outcome.Truncated);
        Assert.True(outcome.WalkCapReached);
        Assert.Null(outcome.DepthLimitReached);
    }

    [Fact]
    public void AWalkStoppedByItsCap_OffersNoContinuation_BecauseNoneExists()
    {
        // nextOffset would be an instruction that cannot work: the next call re-walks the
        // same tree and stops in the same place. The remedy is a narrower walk, and it is in
        // the advice rather than in a field that promises paging.
        using MailService service = Service(Tree(new[] { Folder(NamedStore, "Inbox") }, walkCapReached: true));

        FoldersOutcome outcome = service.ListFolders();

        Assert.Null(outcome.NextOffset);
        Assert.Contains(
            outcome.Advice!,
            a => a.Contains("INCOMPLETE LISTING", StringComparison.Ordinal)
                && a.Contains("Paging cannot reach them", StringComparison.Ordinal)
                && a.Contains(MailService.FolderWalkAbsoluteCap.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal));
    }

    [Fact]
    public void ADepthGuardHit_IsReportedApartFromTheCap()
    {
        // Two different holes with two different meanings: a cap means the profile is huge,
        // a depth guard hit means the tree is pathological or cyclic.
        using MailService service = Service(Tree(new[] { Folder(NamedStore, "Inbox") }, depthLimitReached: true));

        FoldersOutcome outcome = service.ListFolders();

        Assert.True(outcome.Truncated);
        Assert.True(outcome.DepthLimitReached);
        Assert.Null(outcome.WalkCapReached);
        Assert.Contains(outcome.Advice!, a => a.Contains("deeper than 64 levels", StringComparison.Ordinal));
    }

    [Fact]
    public void ACompleteWalk_RaisesNothing()
    {
        // The other half of every flag: a listing that lost nothing must stay silent, or the
        // flag becomes wallpaper before the day it matters.
        using MailService service = Service(Tree(new[] { Folder(NamedStore, "Inbox") }));

        FoldersOutcome outcome = service.ListFolders();

        Assert.False(outcome.Truncated);
        Assert.Null(outcome.WalkCapReached);
        Assert.Null(outcome.DepthLimitReached);
        Assert.Null(outcome.StoresUnnamed);
        Assert.Null(outcome.Advice);
    }

    [Fact]
    public void APageableRemainderStillPages_EvenWhenTheWalkWasAlsoCut()
    {
        // The two causes are independent: folders past this page are reachable by offset,
        // folders past the cap are not, and an answer can be short of both at once.
        List<ComFolderInfo> many = new List<ComFolderInfo>();
        for (int i = 0; i < MailService.FoldersPerCallCap + 3; i++)
        {
            many.Add(Folder(NamedStore, "F" + i.ToString("D5", CultureInfo.InvariantCulture)));
        }

        FoldersOutcome outcome = MailService.PageFolders(new ComFolderTree(many, walkCapReached: true), offset: 0);

        Assert.True(outcome.Truncated);
        Assert.Equal(MailService.FoldersPerCallCap, outcome.NextOffset);
        Assert.True(outcome.WalkCapReached);
    }

    [Fact]
    public void TheListingAdvice_IsPureAndSaysOneThingPerHole()
    {
        Assert.Null(MailService.DescribeFolderListingCoverage(false, false, 0));

        IReadOnlyList<string> all = MailService.DescribeFolderListingCoverage(true, true, 2)!;
        Assert.Equal(3, all.Count);
    }

    // ================================================ G4: a delegate scope built from a stub

    [Fact]
    [Trait("TimeZone", "Machine")]
    public void ADelegateFolderScope_SaysWhenItsNameListCameFromATruncatedWalk()
    {
        // The delegate index namespace is FLAT, so this scope is an OR of folder NAMES read
        // out of the COM tree. A name the walk never reached is a folder searched by no tier
        // at all - and folderNamesMatched, being a count, reads the same either way.
        using MailService service = DelegateService(folderTreeTruncated: true);

        SearchOutcome outcome = service.Search(new SearchRequest
        {
            Query = "test",
            StoreId = TestFolders.StoreId(DelegateStore),
            FolderId = TestFolders.FolderId(DelegateStore, "Archive"),
            Top = 25,
            SnippetChars = 0,
        });

        Assert.True(outcome.Scope!.FolderNamesTruncated);
        Assert.Equal("delegate_folders", outcome.Scope.Shape);
        Assert.True(outcome.Degraded);
        Assert.Contains(
            outcome.Advice!,
            a => a.Contains("INCOMPLETE SCOPE", StringComparison.Ordinal)
                && a.Contains("matched by folder NAME", StringComparison.Ordinal));
    }

    [Fact]
    [Trait("TimeZone", "Machine")]
    public void TheSameScopeFromACompleteWalk_StaysSilent()
    {
        using MailService service = DelegateService(folderTreeTruncated: false);

        SearchOutcome outcome = service.Search(new SearchRequest
        {
            Query = "test",
            StoreId = TestFolders.StoreId(DelegateStore),
            FolderId = TestFolders.FolderId(DelegateStore, "Archive"),
            Top = 25,
            SnippetChars = 0,
        });

        Assert.Null(outcome.Scope!.FolderNamesTruncated);
        Assert.DoesNotContain(outcome.Advice ?? Array.Empty<string>(), a => a.Contains("INCOMPLETE SCOPE", StringComparison.Ordinal));
    }

    [Fact]
    public void TheTruncationFlag_ChangesTheReport_NeverTheScope()
    {
        // Narrowing further would drop mail; widening would answer with the whole delegate
        // mailbox under a folder scope the caller chose. So the predicates are identical and
        // only the report differs - which is what makes this fix free of behaviour risk.
        string[] tree = { "Archive", "Archive/2024", "Archive/2025" };

        FolderScopeResolution complete = FolderScopeResolver.ForDelegateStore(
            OwnerPrefix + "/1/" + DelegateStore, "Archive", includeSubfolders: true, tree, false);
        FolderScopeResolution cut = FolderScopeResolver.ForDelegateStore(
            OwnerPrefix + "/1/" + DelegateStore, "Archive", includeSubfolders: true, tree, true);

        Assert.Equal(complete.Kind, cut.Kind);
        Assert.Equal(complete.Scope, cut.Scope);
        Assert.Equal(complete.FolderPaths, cut.FolderPaths);
        Assert.False(complete.FolderTreeTruncated);
        Assert.True(cut.FolderTreeTruncated);
    }

    [Fact]
    public void ATruncatedWalkIsNotAnUnavailableOne()
    {
        // Unavailable means there was no tree to read and the scope WIDENS - over-return, the
        // safe direction. Truncated means the tree was read and is short, so the scope
        // narrows. Reporting them as one flag would send a caller the wrong remedy.
        FolderScopeResolution cut = FolderScopeResolver.ForDelegateStore(
            OwnerPrefix + "/1/" + DelegateStore, "Archive", includeSubfolders: true,
            new[] { "Archive" }, true);

        Assert.True(cut.FolderTreeTruncated);
        Assert.False(cut.FolderTreeUnavailable);
        Assert.False(cut.Widened);
    }

    [Fact]
    public void TheScopeSentence_IsPure()
    {
        Assert.Null(MailService.DescribeTruncatedFolderNames(null));
        Assert.Null(MailService.DescribeTruncatedFolderNames(
            FolderScopeResolver.ForDelegateStore(
                OwnerPrefix + "/1/" + DelegateStore, "Archive", true, new[] { "Archive" }, false)));

        string sentence = MailService.DescribeTruncatedFolderNames(
            FolderScopeResolver.ForDelegateStore(
                OwnerPrefix + "/1/" + DelegateStore, "Archive", true, new[] { "Archive" }, true))!;
        Assert.Contains("exhaustive:true", sentence, StringComparison.Ordinal);
    }

    // ================================================================== fixtures

    private static ComFolderInfo Folder(string store, string path)
    {
        return new ComFolderInfo(store, path, path, 1, 0, 0);
    }

    private static ComFolderTree Tree(
        IReadOnlyList<ComFolderInfo> folders,
        bool walkCapReached = false,
        bool depthLimitReached = false,
        int storesUnnamed = 0)
    {
        return new ComFolderTree(folders, walkCapReached, depthLimitReached, storesUnnamed);
    }

    /// <summary>An all-stores sweep that covered an unnameable store under its label.</summary>
    private static ComSweepResult SweepWithUnnamedStore()
    {
        return new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 4,
            foldersSkipped: 0,
            sweptFolders: new[] { Unnamed + "/Inbox" },
            perStore: new[]
            {
                new ComStoreSweepCounters(Unnamed, foldersSwept: 4, foldersSkipped: 0, foldersFailed: 0, foldersAbsent: 0),
            },
            storesUnnamed: 1);
    }

    private static MailService Service(ComFolderTree tree)
    {
        return new MailService(
            new DirectGateway(FolderSession.Create(tree, new ComFolderPathList(Array.Empty<string>()))),
            null,
            new DelegateIndexClient(Array.Empty<string>(), delegateScope: null));
    }

    /// <summary>
    /// A profile whose only indexed store is an owner mailbox carrying a delegate subtree, so
    /// <c>search</c> resolves <see cref="DelegateStore"/> through the delegate branch and the
    /// COM folder walk is what supplies the folder names.
    /// </summary>
    private static MailService DelegateService(bool folderTreeTruncated)
    {
        ComFolderPathList paths = new ComFolderPathList(
            new[] { "Archive", "Archive/2024", "Archive/2025" }, folderTreeTruncated, false);
        return new MailService(
            new DirectGateway(FolderSession.Create(new ComFolderTree(Array.Empty<ComFolderInfo>()), paths)),
            null,
            new DelegateIndexClient(new[] { OwnerPrefix }, OwnerPrefix + "/1/" + DelegateStore));
    }

    /// <summary>
    /// The index of the owner's stores, which also holds - the branch this fixture exists for - the
    /// delegate's subtree, and whose reachability probe finds an item in whatever scope it names.
    /// </summary>
    private sealed class DelegateIndexClient : StubIndexClient
    {
        private readonly string? _delegateScope;

        internal DelegateIndexClient(IReadOnlyList<string> knownPrefixes, string? delegateScope)
            : base(knownPrefixes, FolderWalkReportingTests.Frontier)
        {
            _delegateScope = delegateScope;
        }

        protected override bool Knows(string scope)
        {
            if (_delegateScope != null
                && (string.Equals(scope, _delegateScope, StringComparison.OrdinalIgnoreCase)
                    || scope.StartsWith(_delegateScope + "/", StringComparison.OrdinalIgnoreCase)))
            {
                return true;
            }

            return base.Knows(scope);
        }

        protected override string ProbedItemUrl(string sql)
        {
            return (IndexRows.ScopeOf(sql) ?? string.Empty) + "/probed-item";
        }
    }

    /// <summary>
    /// A session that answers the two folder walks, the profile store list and a sweep that
    /// found nothing. A <see cref="DispatchProxy"/> rather than a stub per contract method, so
    /// a method added to <c>IOutlookSession</c> needs no change here. Not sealed:
    /// DispatchProxy derives from its TProxy at runtime and refuses a sealed one.
    /// </summary>
    private class FolderSession : DispatchProxy
    {
        private ComFolderTree _tree = new ComFolderTree(Array.Empty<ComFolderInfo>());
        private ComFolderPathList _paths = new ComFolderPathList(Array.Empty<string>());

        internal static IOutlookSession Create(ComFolderTree tree, ComFolderPathList paths)
        {
            object proxy = Create<IOutlookSession, FolderSession>()
                ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
            ((FolderSession)proxy)._tree = tree;
            ((FolderSession)proxy)._paths = paths;
            return (IOutlookSession)proxy;
        }

        /// <inheritdoc />
        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            return targetMethod?.Name switch
            {
                nameof(IOutlookSession.GetProfileName) => "T1 stand-in profile",
                nameof(IOutlookSession.GetStoreDetails) => TestFolders.WithRoots(new[]
                {
                    new ComStoreDetail(NamedStore, "store-a", 0, true),
                    new ComStoreDetail(Unnamed, "store-unnamed", 3, null, nameUnreadable: true),
                    new ComStoreDetail(DelegateStore, "store-delegate", 1, true),
                }),
                nameof(IOutlookSession.TryDescribeFolder) => TestFolders.Describe(args),
                nameof(IOutlookSession.ListFolders) => _tree,
                nameof(IOutlookSession.ListFolderPaths) => _paths,
                nameof(IOutlookSession.SweepFoldersNewerThan) => new ComSweepResult(
                    Array.Empty<ComMailBrief>(),
                    foldersSwept: 4,
                    foldersSkipped: 0,
                    sweptFolders: new[] { "swept/Inbox" },
                    perStore: new[]
                    {
                        new ComStoreSweepCounters(
                            (args?[3] as string) ?? NamedStore,
                            foldersSwept: 4, foldersSkipped: 0, foldersFailed: 0, foldersAbsent: 0),
                    }),
                _ => throw new NotSupportedException(
                    "The stand-in session was asked for " + (targetMethod?.Name ?? "an unnamed member")
                    + ", which this test does not model."),
            };
        }
    }
}
