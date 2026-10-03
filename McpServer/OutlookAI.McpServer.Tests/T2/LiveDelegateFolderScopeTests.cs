using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// T2 live acceptance for soak fix 15, the delegate half - READ-ONLY. Delegate mailboxes are
/// indexed FLAT, so the nested delegate URL the product used to build addressed a folder that
/// does not exist and every delegate SUBFOLDER search returned zero rows, silently. Proven here
/// against COM ground truth: the old shape still returns 0, the shipped shape returns the folder's
/// real population, for every delegate mailbox.
/// <para>
/// <b>Why it is its own class, in <see cref="LiveCollections.Phase2"/> (Q74 B1, 2026-10-03).</b>
/// These two tests lived in <see cref="LiveFolderScopeTests"/>, in the
/// <see cref="LiveCollections.MoveArchive"/> collection - whose fixture creates and deletes test
/// folders in the hub (<c>LiveOutlookTestMailer.DeleteTestFolders</c>, in its constructor and its
/// teardown). They need a delegate mailbox, so they can only ever run on the maintainer's
/// workstation, and that machine is read-only: under its profile the MoveArchive fixture throws at
/// its first folder write, before either test starts. They use nothing of that fixture but the
/// settings and the <see cref="MailService"/>, which <see cref="LivePhase2Fixture"/> holds too, and
/// that collection's fixture writes nothing.
/// </para>
/// <para>
/// SAFETY: the delegate mailboxes are READ-ONLY here - counts, booleans and folder paths only,
/// never a subject, sender or body (S4). Nothing is written anywhere.
/// </para>
/// </summary>
[Collection(LiveCollections.Phase2)]
[Trait("Category", "Live")]
public sealed class LiveDelegateFolderScopeTests
{
    /// <summary>
    /// The population BOTH delegate tests iterate, named as the Production refusal will wrap it.
    /// </summary>
    private const string DelegatePopulation = "a delegate or shared mailbox to resolve folders in";

    /// <summary>What a reader of a PROVED NOTHING line here is to do about it.</summary>
    private const string DelegateRemedy =
        "To exercise it, run on a profile that opens a delegate or shared mailbox and name that "
        + "mailbox in 'expectedDelegateStoreDisplayNames'. A local PST cannot stand in for one: the "
        + "shape these two tests exist for is Windows Search publishing a delegate mailbox's folders "
        + "FLAT while Outlook nests them, which no PST produces.";

    private readonly LivePhase2Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveDelegateFolderScopeTests(LivePhase2Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    /// <summary>
    /// The delegate mailboxes this machine has - and the ONLY way this file obtains them.
    /// <para>
    /// <b>Why it is a method with a sink rather than a read of the settings.</b>
    /// <see cref="DelegateFirstLevelFolders_StillResolve_AndTheWholeMailboxIsUnfiltered"/> was a
    /// bare <c>foreach</c> over that list with no non-empty guard, so on any machine whose settings
    /// name no delegate mailbox it iterated nothing, asserted nothing and reported GREEN - while
    /// its sibling had asserted the list was non-empty since the day it was written.
    /// The omission was an oversight rather than a decision (<c>Requires=DelegateStore</c> should
    /// keep both unselected on such a machine, but selection is a filter string somebody types, not
    /// a guarantee). Routing both through one call means the two can no longer answer the same
    /// question differently. See <see cref="LivePopulationCoverage"/>.
    /// </para>
    /// </summary>
    private IReadOnlyList<string> DelegateStores(string whatWouldNotRun)
    {
        return LivePopulationCoverage.Require(
            _fixture.Settings,
            _fixture.Settings.ExpectedDelegateStoreDisplayNames,
            DelegatePopulation,
            whatWouldNotRun,
            DelegateRemedy,
            _output.WriteLine);
    }

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "DelegateStore")]
    [Trait("Writes", "Nothing")]
    public void DelegateSubfolders_AreReachableAgain_AndTheOldNestedShapeStillReturnsZero()
    {
        // This one has always asserted the list is non-empty, and it keeps doing so: the guard
        // announces and (on Production) refuses, and the assertion below still fails a Portable
        // machine that selected this test anyway. Deliberately NOT weakened to an announcement -
        // the defect being fixed here is a test that passes proving nothing, and this one never did.
        IReadOnlyList<string> delegates = DelegateStores("the delegate nested-subfolder probe");
        Assert.True(delegates.Count > 0, "the live settings must name at least one delegate store");

        IndexSearchService index = IndexSearchService.CreateDefault(out _);
        int provenSubfolders = 0;

        foreach (string delegateStore in delegates)
        {
            // COM ground truth: a NESTED folder (depth >= 2) of this delegate mailbox
            // that actually holds mail. Discovered at runtime - no mailbox identifier
            // is ever committed to this public repo (S6).
            string delegateRoot = LiveFolderProbe.ResolveDelegateRootScope(index, delegateStore);
            Assert.True(
                MapiItemUrl.TryBuildFolderPathDisplay(delegateRoot, out string? rootPath) && rootPath != null,
                "the delegate root scope must yield a folder display path");

            // Candidates are NESTED folders holding mail. Item counts alone are not
            // enough: a delegate mailbox's biggest subfolder can be a CALENDAR subtree,
            // whose items are not email rows at all - so a candidate is accepted only
            // once the shipped shape actually returns email for it.
            FolderView? nested = null;
            long comCount = 0;
            int newRows = 0;
            foreach (FolderView candidate in LiveFolderProbe.FolderTree(Service, delegateStore)
                .Where(f => f.Path.Contains('/') && (f.Items ?? 0) >= 5)
                .OrderByDescending(f => f.Items ?? 0)
                .Take(12))
            {
                string candidateLeaf = candidate.Path[(candidate.Path.LastIndexOf('/') + 1)..];
                int rows = LiveFolderProbe.DrainCount(index, new IndexQuery
                {
                    Scope = delegateRoot,
                    FolderPathsAnyOf = new[] { rootPath + "/" + candidateLeaf },
                    Kinds = KindFilter.MailKindOnly,
                    Top = 5000,
                });

                if (rows > 0)
                {
                    nested = candidate;
                    comCount = candidate.Items ?? 0;
                    newRows = rows;
                    break;
                }
            }

            Assert.True(nested != null, $"no reachable nested MAIL folder found in delegate store '{delegateStore}'");
            _output.WriteLine($"[{delegateStore}] nested folder '{nested!.Path}' COM items={comCount}");

            // --- (a) THE DEFECT, still reproducible: the pre-fix nested delegate URL.
            string oldShape = delegateRoot + "/" + nested.Path;
            int oldRows = LiveFolderProbe.DrainCount(index, new IndexQuery
            {
                Scope = oldShape,
                Kinds = KindFilter.MessagesAndAttachments,
                Top = 5000,
            });
            Assert.Equal(0, oldRows);
            _output.WriteLine($"[{delegateStore}] pre-fix nested URL -> {oldRows} rows (the silent zero)");

            // --- (b) THE FIX: delegate store root + flat folder-name equality.
            _output.WriteLine($"[{delegateStore}] shipped shape -> {newRows} email rows (COM {comCount})");
            Assert.True(newRows > 0, $"the delegate subfolder is still unreachable ({delegateStore}/{nested.Path})");

            // The index census and COM disagree only by index lag (and, on a colliding
            // leaf name, by the merged folder's extra rows); a predicate error would be
            // an order-of-magnitude miss, not a few percent.
            double ratio = comCount == 0 ? 1 : newRows / (double)comCount;
            Assert.InRange(ratio, 0.70, 2.00);

            // --- (c) end to end through the product, with the flag both ways.
            foreach (bool includeSubfolders in new[] { false, true })
            {
                SearchOutcome outcome = Service.Search(new SearchRequest
                {
                    Store = delegateStore,
                    Folder = nested.Path,
                    IncludeSubfolders = includeSubfolders,
                    IndexOnly = true,
                    Top = 5,
                    SnippetChars = 0,
                });

                Assert.NotEmpty(outcome.Hits);
                Assert.All(outcome.Hits, h => Assert.Equal(delegateStore, h.Store));
                Assert.NotNull(outcome.Scope);
                Assert.Equal(nested.Path, outcome.Scope!.Folder);
                Assert.Equal(includeSubfolders, outcome.Scope.IncludeSubfolders);
                Assert.StartsWith("delegate_", outcome.Scope.Shape, StringComparison.Ordinal);

                _output.WriteLine(
                    $"[{delegateStore}] search include_subfolders={includeSubfolders}: hits={outcome.Hits.Count} "
                    + $"shape={outcome.Scope.Shape} widened={outcome.Scope.Widened} "
                    + $"folderNames={outcome.Scope.FolderNamesMatched}");
            }

            provenSubfolders++;
        }

        Assert.Equal(delegates.Count, provenSubfolders);
    }

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "DelegateStore")]
    [Trait("Writes", "Nothing")]
    public void DelegateFirstLevelFolders_StillResolve_AndTheWholeMailboxIsUnfiltered()
    {
        foreach (string delegateStore in DelegateStores("the delegate first-level folder probe"))
        {
            IReadOnlyList<FolderView> tree = LiveFolderProbe.FolderTree(Service, delegateStore);
            FolderView? topLevel = tree
                .Where(f => !f.Path.Contains('/') && (f.Items ?? 0) >= 1)
                .OrderByDescending(f => f.Items ?? 0)
                .FirstOrDefault();
            Assert.True(topLevel != null, $"no populated first-level folder in '{delegateStore}'");

            SearchOutcome folderScoped = Service.Search(new SearchRequest
            {
                Store = delegateStore,
                Folder = topLevel!.Path,
                IncludeSubfolders = false,
                IndexOnly = true,
                Top = 3,
                SnippetChars = 0,
            });
            Assert.NotEmpty(folderScoped.Hits);

            // The whole delegate mailbox needs no folder filter at all - its root scope
            // already covers every flat folder.
            SearchOutcome wholeStore = Service.Search(new SearchRequest
            {
                Store = delegateStore,
                IndexOnly = true,
                Top = 3,
                SnippetChars = 0,
            });
            Assert.NotEmpty(wholeStore.Hits);
            Assert.Null(wholeStore.Scope);

            _output.WriteLine(
                $"[{delegateStore}] first-level '{topLevel.Path}' hits={folderScoped.Hits.Count}; "
                + $"whole-mailbox hits={wholeStore.Hits.Count}");
        }
    }
}
