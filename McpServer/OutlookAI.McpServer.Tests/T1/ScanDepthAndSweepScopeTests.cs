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
/// Two bounds of the search surface that had nothing to say for themselves.
/// <para>
/// GAP F4 - the exhaustive scan's folder walk had no depth bound at all, alone among the
/// three walks over Outlook's folder graph. A cyclic or pathologically deep tree is therefore
/// a <c>StackOverflowException</c>, which .NET lets nobody catch: it ends the COM host
/// process, the parent reports Outlook as having gone away, and two of those open the breaker
/// for 30 s. The guard is the comparison the sweep walk already makes against the same
/// constant - but stopping quietly would replace a crash with a silent truncation, in the one
/// mode a caller picks BECAUSE completeness matters, so the stop latches and reaches the
/// payload as a flag plus the <c>depth_limit</c> code the sweep already uses.
/// </para>
/// <para>
/// GAP E2 - the sweep's default folder set is four arrival-path folders per store and it does
/// not descend, so mail a server-side rule files into a subfolder before the indexer reaches
/// it is in neither tier. That was stated only as an English sentence in <c>sweep.scope</c>,
/// which an agent can only branch on by comparing against prose. It is now a token beside it,
/// rendered from ONE classifier so the two cannot describe different breadths - and
/// deliberately NOT a coverage code, because this shape holds for nearly every search and a
/// flag that fires always devalues the ones that fire rarely (the B2 rule).
/// </para>
/// <para>
/// What T1 CANNOT reach and needs a live profile: that the recursion really is bounded, which
/// takes a folder tree deeper than 64 levels. This tier pins that once the walk says it
/// stopped, every rendering of the answer says so too.
/// </para>
/// </summary>
public sealed class ScanDepthAndSweepScopeTests
{
    private const string Sid = "{S-1-5-21-1111111111-2222222222-3333333333-1001}";

    private const string Store = "alice@example.com";

    private const string StorePrefix = "mapi16://" + Sid + "/" + Store + "($deadbeef)";

    private static readonly DateTime Frontier = new(2026, 8, 18, 9, 30, 0, DateTimeKind.Utc);

    private static readonly ComStoreDetail[] ProfileStores = { new(Store, "store-alice", 0, true) };

    // ==================================================================== F4: the classifiers

    [Fact]
    public void AScanStoppedByTheDepthGuard_IsPartial_AndRaisesTheSharedCode()
    {
        ExhaustiveInfo scan = new ExhaustiveInfo { FoldersScanned = 9, DepthLimitReached = true };

        scan.CoverageGaps = FreshMerge.DescribeExhaustiveCoverageGaps(scan);

        Assert.Contains(FreshMerge.ScanGapDepthLimit, scan.CoverageGaps!);
        Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyExhaustiveFreshness(scan));
    }

    [Fact]
    public void TheScanCodeIsTheSweepsOwnToken_NotASecondNameForOneBound()
    {
        // One guard, one number, one word for it. The sweep's walk and the listing walk both
        // stop at FolderWalkDepthGuard and both say depth_limit; a third spelling would make
        // an agent learn a vocabulary per tier for a fact that is identical in all three.
        Assert.Equal(FreshMerge.GapDepthLimit, FreshMerge.ScanGapDepthLimit);
    }

    [Fact]
    public void AScanThatCoveredItsScope_RaisesNothing()
    {
        ExhaustiveInfo scan = new ExhaustiveInfo { FoldersScanned = 9 };

        Assert.Null(FreshMerge.DescribeExhaustiveCoverageGaps(scan));
        Assert.Equal(FreshMerge.FreshnessLive, FreshMerge.ClassifyExhaustiveFreshness(scan));
    }

    [Fact]
    public void TheDepthSentence_QuotesTheGuard_AndSaysWhereToLook()
    {
        ExhaustiveInfo scan = new ExhaustiveInfo { FoldersScanned = 9, DepthLimitReached = true };
        scan.CoverageGaps = new[] { FreshMerge.ScanGapDepthLimit };

        string line = Assert.Single(MailService.DescribeExhaustiveCoverage(scan, top: 25));

        // Quoted from the constant, never restated: a bound named in prose beside a number
        // nothing compares it with is exactly how the pair drifts (gap G3's lesson).
        Assert.Contains(
            OutlookComSession.FolderWalkDepthGuard.ToString(CultureInfo.InvariantCulture),
            line,
            StringComparison.Ordinal);
        Assert.Contains("NOT covered", line, StringComparison.Ordinal);
        Assert.Contains("list_folders", line, StringComparison.Ordinal);
    }

    // ================================================================ F4: through the tool

    [Fact]
    public void AnExhaustiveSearchWhoseWalkHitTheGuard_ReportsItOnEveryField()
    {
        using MailService service = Service(scan: DeepScan);

        SearchOutcome outcome = service.Search(ExhaustiveRequest());

        Assert.True(outcome.Exhaustive!.DepthLimitReached);
        Assert.Contains(FreshMerge.ScanGapDepthLimit, outcome.Exhaustive.CoverageGaps!);
        Assert.Equal(FreshMerge.FreshnessPartial, outcome.Freshness);
        Assert.True(outcome.Degraded);
        Assert.Contains(
            outcome.Advice!,
            a => a.Contains("deeper than", StringComparison.Ordinal));
    }

    [Fact]
    public void AnOrdinaryExhaustiveSearch_SaysNothingAboutDepth()
    {
        using MailService service = Service(scan: PlainScan);

        SearchOutcome outcome = service.Search(ExhaustiveRequest());

        Assert.False(outcome.Exhaustive!.DepthLimitReached);
        Assert.DoesNotContain(
            FreshMerge.ScanGapDepthLimit,
            outcome.Exhaustive.CoverageGaps ?? Array.Empty<string>());
        Assert.Equal(FreshMerge.FreshnessLive, outcome.Freshness);
    }

    // ==================================================================== E2: the classifier

    [Theory]
    [InlineData(false, true, MailService.SweepScopeDefaultFolders)]
    [InlineData(false, false, MailService.SweepScopeDefaultFolders)]
    [InlineData(true, true, MailService.SweepScopeFolder)]
    [InlineData(true, false, MailService.SweepScopeFolderOnly)]
    public void TheScopeShape_IsDecidedByTheFolderScopeAndTheSubfolderFlagAlone(
        bool folderScoped, bool includeSubfolders, string expected)
    {
        Assert.Equal(expected, MailService.ClassifySweepScope(folderScoped, includeSubfolders));
    }

    [Fact]
    public void WithoutAFolder_TheSubfolderFlagCannotWidenTheDefaultSet()
    {
        // The default set is shallow by construction (SweepFolder, not SweepFolderTree), so
        // include_subfolders means nothing there - and a shape that implied otherwise would
        // be the E2 hole restated in a field.
        Assert.Equal(
            MailService.ClassifySweepScope(folderScoped: false, includeSubfolders: false),
            MailService.ClassifySweepScope(folderScoped: false, includeSubfolders: true));
    }

    [Fact]
    public void EveryDeclaredShape_HasItsOwnSentence_AndTheSentencesAreTheOnesCallersRead()
    {
        // The prose is rendered FROM the token, so a token nobody wrote a sentence for would
        // put an unexplained breadth in the payload. Read off the type so a new shape cannot
        // ship untested.
        foreach (string shape in AllSweepScopeShapes())
        {
            string sentence = MailService.DescribeSweepScope(shape);

            Assert.DoesNotContain("unknown sweep scope", sentence, StringComparison.Ordinal);
        }

        // Byte-for-byte what sweep.scope has always carried: this is a payload contract that
        // shipped, and the token was added beside it rather than in place of it.
        Assert.Equal(
            MailService.DefaultSweepScopeDescription,
            MailService.DescribeSweepScope(MailService.SweepScopeDefaultFolders));
        Assert.Equal("folder", MailService.DescribeSweepScope(MailService.SweepScopeFolder));
        Assert.Equal("folder (no subfolders)", MailService.DescribeSweepScope(MailService.SweepScopeFolderOnly));
    }

    // ================================================================ E2: through the tool

    [Fact]
    public void AnUnscopedSearch_SaysItSweptTheShallowDefaultSet()
    {
        using MailService service = Service();

        SearchOutcome outcome = service.Search(new SearchRequest { Query = "test", Top = 25, SnippetChars = 0 });

        Assert.Equal(MailService.SweepScopeDefaultFolders, outcome.Sweep!.ScopeShape);
        Assert.Equal(MailService.DefaultSweepScopeDescription, outcome.Sweep.Scope);
    }

    [Fact]
    public void TheShallowDefaultSet_IsNotACoverageHole_AndDoesNotDegradeTheSearch()
    {
        // The decision this row turned on. It is the breadth of the tier rather than a hole
        // in what the sweep was asked to do, and ClassifyFreshness derives 'partial' from the
        // code list - so a code here would make almost every search this product answers
        // permanently degraded and blunt the flags that mean something.
        using MailService service = Service();

        SearchOutcome outcome = service.Search(new SearchRequest { Query = "test", Top = 25, SnippetChars = 0 });

        Assert.Null(outcome.Sweep!.CoverageGaps);
        Assert.Equal(FreshMerge.FreshnessLive, outcome.Freshness);
        Assert.Null(outcome.Degraded);
    }

    [Theory]
    [InlineData(true, MailService.SweepScopeFolder)]
    [InlineData(false, MailService.SweepScopeFolderOnly)]
    public void AFolderScopedSearch_ReportsWhichOfTheTwoFolderBreadthsItUsed(
        bool includeSubfolders, string expected)
    {
        using MailService service = Service();

        SearchOutcome outcome = service.Search(new SearchRequest
        {
            Query = "test",
            Store = Store,
            Folder = "Projects/2026",
            IncludeSubfolders = includeSubfolders,
            Top = 25,
            SnippetChars = 0,
        });

        Assert.Equal(expected, outcome.Sweep!.ScopeShape);
        Assert.Equal(MailService.DescribeSweepScope(expected), outcome.Sweep.Scope);
    }

    [Fact]
    public void ASweepThatCouldNotRun_ClaimsNoBreadthAtAll()
    {
        // Paired with sweep.scope, which has always been absent here: "it covered the default
        // folders" is a claim about coverage, and a refused sweep covered nothing. That is
        // the one way this differs from attachmentTextCovered, which survives every early
        // return because it states a capability instead.
        using MailService service = Service();

        SearchOutcome outcome = service.Search(new SearchRequest
        {
            Query = "test",
            AttachmentHitsOnly = true,
            Top = 25,
            SnippetChars = 0,
        });

        Assert.False(outcome.Sweep!.Performed);
        Assert.Null(outcome.Sweep.Scope);
        Assert.Null(outcome.Sweep.ScopeShape);
    }

    // ===================================================================== fixtures

    private static IReadOnlyList<string> AllSweepScopeShapes()
    {
        return typeof(MailService)
            .GetFields(BindingFlags.Public | BindingFlags.Static)
            .Where(f => f.IsLiteral && f.FieldType == typeof(string)
                && f.Name.StartsWith("SweepScope", StringComparison.Ordinal))
            .Select(f => (string)f.GetRawConstantValue()!)
            .ToList();
    }

    private static SearchRequest ExhaustiveRequest()
    {
        return new SearchRequest
        {
            Query = "test",
            Store = Store,

            // An exhaustive scan refuses to run unbounded, so the request carries the bound
            // its own validation demands.
            Folder = "Inbox",
            Exhaustive = true,
            Top = 25,
            SnippetChars = 0,
        };
    }

    private static MailService Service(Func<string?, ComExhaustiveResult>? scan = null)
    {
        return new MailService(
            new DirectGateway(StandInSession.Create(ProfileStores, Sweep, scan)), null, new StubIndexClient());
    }

    /// <summary>A scan whose walk refused a subtree for being too deep.</summary>
    private static ComExhaustiveResult DeepScan(string? store) => Scan(store, depthLimitReached: true);

    private static ComExhaustiveResult PlainScan(string? store) => Scan(store, depthLimitReached: false);

    private static ComExhaustiveResult Scan(string? store, bool depthLimitReached)
    {
        return new ComExhaustiveResult(
            new[] { Mail("EX1") },
            foldersScanned: 9,
            foldersSkipped: 0,
            engine: "ci_phrasematch",
            instantSearchEnabled: true,
            truncated: false,
            timedOut: false,
            rowsDropped: 0,
            rowsUnreadable: 0,
            depthLimitReached: depthLimitReached);
    }

    private static ComSweepResult Sweep(string? onlyStore)
    {
        return new ComSweepResult(
            new[] { Mail("AA1") },
            foldersSwept: 4,
            foldersSkipped: 0,
            sweptFolders: new[] { Store + "/Inbox" },
            perStore: new[] { new ComStoreSweepCounters(Store, 4, 0, 0, 0) });
    }

    private static ComMailBrief Mail(string entryId)
    {
        return new ComMailBrief(
            entryId: entryId,
            storeDisplayName: Store,
            storeId: "store-alice",
            folderName: "Inbox",
            folderKind: "inbox",
            subject: "a test mail",
            senderName: "Bob",
            senderAddress: "bob@example.com",
            receivedTime: Frontier,
            isRead: true,
            hasAttachments: false,
            sizeBytes: 2048,
            body: "test body");
    }

    /// <summary>An index that knows this one store and holds no search rows.</summary>
    private sealed class StubIndexClient : IIndexClient
    {
        private const string DiscoveryTail = " System.ItemUrl FROM SystemIndex WHERE System.Kind='email'";

        public IndexProviderKind Provider => IndexProviderKind.OleDb;

        public IReadOnlyList<IReadOnlyDictionary<string, object?>> ExecuteRows(
            string sql, int maxRows, int? commandTimeoutSeconds = null)
        {
            if (sql.EndsWith(DiscoveryTail, StringComparison.Ordinal))
            {
                return Rows(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                {
                    ["System.ItemUrl"] = StorePrefix + "/0/Inbox/sampled-item",
                });
            }

            if (sql.Contains("System.Message.DateReceived FROM SystemIndex", StringComparison.Ordinal))
            {
                return Rows(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                {
                    ["System.Message.DateReceived"] = Frontier,
                });
            }

            if (sql.StartsWith("SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE", StringComparison.Ordinal))
            {
                return Rows(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                {
                    ["System.ItemUrl"] = StorePrefix + "/0/Inbox/probed-item",
                });
            }

            return Array.Empty<IReadOnlyDictionary<string, object?>>();
        }

        private static IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows(
            IReadOnlyDictionary<string, object?> row)
        {
            return new[] { row };
        }
    }

    private sealed class DirectGateway : IComGateway
    {
        private readonly IOutlookSession _session;

        internal DirectGateway(IOutlookSession session)
        {
            _session = session;
        }

        public event Action? OutlookGone
        {
            add { }
            remove { }
        }

        public bool IsConnected => true;

        public bool? QuitSinkActive => null;

        public bool ProbeConnected() => true;

        public T Run<T>(Func<IOutlookSession, T> operation) => operation(_session);

        public T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery) => operation(_session);

        public T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false)
            => operation(_session);

        public ComHostDiagnostics GetDiagnostics() => new ComHostDiagnostics("in-process", "ready");

        public void Dispose()
        {
        }
    }

    /// <summary>Answers the store list, the sweep and the scan; refuses everything else.</summary>
    private class StandInSession : DispatchProxy
    {
        private IReadOnlyList<ComStoreDetail> _stores = Array.Empty<ComStoreDetail>();
        private Func<string?, ComSweepResult> _sweep = _ => throw new NotSupportedException();
        private Func<string?, ComExhaustiveResult>? _scan;

        internal static IOutlookSession Create(
            IReadOnlyList<ComStoreDetail> stores,
            Func<string?, ComSweepResult> sweep,
            Func<string?, ComExhaustiveResult>? scan)
        {
            object proxy = Create<IOutlookSession, StandInSession>()
                ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
            ((StandInSession)proxy)._stores = stores;
            ((StandInSession)proxy)._sweep = sweep;
            ((StandInSession)proxy)._scan = scan;
            return (IOutlookSession)proxy;
        }

        /// <inheritdoc />
        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            switch (targetMethod?.Name)
            {
                case nameof(IOutlookSession.GetProfileName):
                    return "T1 stand-in profile";

                case nameof(IOutlookSession.GetStoreDetails):
                    return _stores;

                // Argument 3 is onlyStoreDisplayName; see IOutlookSession.SweepFoldersNewerThan.
                case nameof(IOutlookSession.SweepFoldersNewerThan):
                    return _sweep(args?[3] as string);

                // Argument 0 is storeDisplayName; see IOutlookSession.ExhaustiveScan.
                case nameof(IOutlookSession.ExhaustiveScan):
                    return _scan != null
                        ? _scan(args?[0] as string)
                        : throw new NotSupportedException("This fixture has no exhaustive scan.");

                default:
                    throw new NotSupportedException(
                        "The stand-in session does not implement " + (targetMethod?.Name ?? "?") + ".");
            }
        }
    }
}
