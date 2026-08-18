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
/// A store scope must be resolved against the profile OUTLOOK has, with the index used only
/// to decide what the index can contribute.
/// <para>
/// THE DEFECT, measured 2026-08-18 against a real Outlook on a clean machine: one store, a
/// PST, not in the Windows Search index, Outlook connected and responsive. <c>list_folders</c>
/// on that store answered. An UNSCOPED search answered and degraded honestly
/// (<c>degraded:true</c>, <c>freshness:"partial"</c>, <c>coverageGaps:["no_index_frontier"]</c>,
/// four folders swept). An exhaustive search of the same store answered <c>freshness:"live"</c>.
/// Only the ordinary store-SCOPED search failed, every time, with
/// <c>InvalidArgument: Store 'Outlook Data File' was not found in the local index. Known
/// stores: . Use list_accounts for the full store list.</c> - three defects in one sentence:
/// </para>
/// <list type="number">
/// <item><description>
/// The scope was resolved against the LOCAL INDEX rather than against Outlook, so on any
/// profile whose stores are not indexed - a PST, an archive-only store, a fresh install, a
/// machine where indexing is off or excluded by policy - every non-exhaustive store-scoped
/// search failed outright. Worse than the unscoped case, which degrades and says so.
/// </description></item>
/// <item><description>
/// A REAL store and a TYPO produced the identical message, so an agent could not choose
/// between retrying with <c>exhaustive:true</c> and correcting the name.
/// </description></item>
/// <item><description>
/// <c>Known stores: .</c> - an empty enumeration, because the list came from the index's own
/// (empty) catalog, and the remedy it offered was a loop: it said to use <c>list_accounts</c>,
/// which returns the exact name that had just failed.
/// </description></item>
/// </list>
/// <para>
/// WHY THIS SHIPPED, and why the fixtures below are shaped as they are: every test until now
/// ran against an index that knew about every store. The profile shape that breaks it - a
/// store catalog that is EMPTY, or that is missing one store the profile has - had no
/// coverage at any tier. Both are driven here through the REAL <see cref="MailService"/>
/// search path, with a stand-in index client and a stand-in Outlook session; no mailbox and
/// no Windows Search index are touched.
/// </para>
/// <para>
/// The widening question, decided rather than inherited: <c>thread</c>'s <c>store</c> widens
/// to the whole profile when it does not resolve, because there it is only a speed hint - the
/// conversation is pinned by id either way, so a wider lookup returns the same conversation.
/// <c>search</c>'s <c>store</c> is a FILTER on the result set, so widening it would answer
/// with another account's mail under a scope the caller chose. It therefore proceeds
/// unindexed for a store the profile has, and refuses a store the profile does not have -
/// never widens. <see cref="TheIndexTierIsSkippedRatherThanWidened"/> is that guarantee.
/// </para>
/// </summary>
public sealed class UnindexedStoreScopeTests
{
    private const string Sid = "{S-1-5-21-1111111111-2222222222-3333333333-1001}";

    /// <summary>An ordinary Exchange store, present in both the profile and the index.</summary>
    private const string IndexedStore = "alice@example.com";

    private const string IndexedPrefix = "mapi16://" + Sid + "/" + IndexedStore + "($deadbeef)";

    /// <summary>The measured case: a local data file the profile has and the index has never seen.</summary>
    private const string PstStore = "Outlook Data File";

    private static readonly DateTime Frontier = new(2026, 8, 18, 9, 30, 0, DateTimeKind.Utc);

    private static readonly ComStoreDetail[] ProfileStores =
    {
        new ComStoreDetail(IndexedStore, "store-alice", 0, true),
        new ComStoreDetail(PstStore, "store-pst", 3, null),
    };

    // ==================================================== the store the profile has

    [Fact]
    public void AStoreTheProfileHasAndTheIndexDoesNot_Searches_RatherThanFailing()
    {
        // Probe 14 of the measured run, which used to be an outright failure. The index tier
        // contributes nothing and the freshness sweep is the whole answer - which is a state
        // this payload already knew how to describe, on the unscoped path, for this very store.
        using MailService service = Service(NothingIndexed());

        SearchOutcome outcome = service.Search(Request(PstStore));

        Assert.Single(outcome.Hits);
        Assert.Equal("live", outcome.Hits[0].Source);
        Assert.Equal(PstStore, outcome.Hits[0].Store);
    }

    [Fact]
    public void ThatSearch_ReportsTheHoleInFields_NotOnlyInProse()
    {
        // Every one of these already existed and already fired on the unscoped path. The fix
        // is not to invent a report, it is to stop failing before reaching this one.
        using MailService service = Service(NothingIndexed());

        SearchOutcome outcome = service.Search(Request(PstStore));

        Assert.True(outcome.Degraded);
        Assert.Equal(FreshMerge.FreshnessPartial, outcome.Freshness);
        Assert.Contains(FreshMerge.GapNoIndexFrontier, outcome.Sweep!.CoverageGaps!);
        Assert.True(outcome.Sweep.IndexFrontierMissing);
        Assert.Equal(new[] { PstStore }, outcome.Sweep.StoresWithoutIndex);

        // The one fact none of the above can state: the index tier did not RUN. rowsScanned 0
        // alone reads as "a statement ran and matched nothing", which points an agent at "no
        // such mail" instead of at exhaustive:true.
        Assert.True(outcome.Index!.StoreNotIndexed);
        Assert.Equal(0, outcome.Index.RowsScanned);

        // And the scope block says which shape resolved, so a folder-scoped call cannot claim
        // a narrowing the index never performed.
        Assert.Equal("store_not_indexed", outcome.Scope!.Shape);
    }

    [Fact]
    public void TheIndexTierIsSkippedRatherThanWidened()
    {
        // THE WIDENING DECISION, pinned. The stand-in index answers EVERY search statement
        // with a mail that lives in the OTHER store, so a scope quietly dropped to null would
        // show up here as a hit from alice@example.com under store="Outlook Data File" - a
        // wrong answer wearing the shape of a right one.
        using MailService service = Service(NothingIndexed(searchRows: OneIndexedMailInAlicesStore()));

        SearchOutcome outcome = service.Search(Request(PstStore));

        Assert.DoesNotContain(outcome.Hits, h => h.Store == IndexedStore);
        Assert.All(outcome.Hits, h => Assert.Equal("live", h.Source));
        Assert.Equal(0, outcome.Index!.RowsScanned);
        Assert.Null(outcome.Scope!.Widened);
    }

    [Fact]
    public void AFolderInsideThatStore_IsStillSearchable_AndTheSweepIsWhatBoundsIt()
    {
        // Probe 16, which failed with the identical message. There is no index folder scope to
        // build, so the folder bound lives entirely in the sweep - and the shape says so
        // rather than reporting "folder", which would describe a narrowing that never happened.
        using MailService service = Service(NothingIndexed());
        SearchRequest request = Request(PstStore);
        request.Folder = "Inbox";

        SearchOutcome outcome = service.Search(request);

        Assert.Single(outcome.Hits);
        Assert.Equal("Inbox", outcome.Scope!.Folder);
        Assert.Equal("store_not_indexed", outcome.Scope.Shape);
        Assert.True(outcome.Index!.StoreNotIndexed);
    }

    [Fact]
    public void AStoreMissingFromAnOtherwisePopulatedCatalog_BehavesTheSameWay()
    {
        // The second fixture the defect needed and nothing had: an index that knows about one
        // store and not the other. It is the ordinary mixed profile - an indexed Exchange
        // account plus a PST - and it is also the shape an indexed-but-small store takes when
        // the unordered 2000-row discovery sample misses it.
        using MailService service = Service(OnlyAliceIndexed());

        SearchOutcome outcome = service.Search(Request(PstStore));

        Assert.Single(outcome.Hits);
        Assert.True(outcome.Index!.StoreNotIndexed);
        Assert.Equal(new[] { PstStore }, outcome.Sweep!.StoresWithoutIndex);
    }

    [Fact]
    public void TheStoreTheIndexDoesKnow_IsUnaffected()
    {
        // The other half: this must not buy the unindexed store an answer by turning every
        // search into a sweep. A catalogued store still queries the index tier and still
        // reports itself live.
        using MailService service = Service(OnlyAliceIndexed(searchRows: OneIndexedMailInAlicesStore()));

        SearchOutcome outcome = service.Search(Request(IndexedStore));

        Assert.Contains(outcome.Hits, h => h.Source == "index" && h.Store == IndexedStore);
        Assert.Null(outcome.Index!.StoreNotIndexed);
        Assert.Equal(1, outcome.Index.RowsScanned);
        Assert.Null(outcome.Scope);
        Assert.Null(outcome.Sweep!.StoresWithoutIndex);
    }

    // ================================================ the store that is nowhere at all

    [Fact]
    public void AStoreTheProfileDoesNotHave_IsStillRefused()
    {
        // Probe 15. It must stay an error - the caller asked for something that does not
        // exist - and it must now be TELLABLE APART from probe 14, which no longer errors at
        // all. That is the distinction, in the strongest form the payload has.
        using MailService service = Service(NothingIndexed());

        ArgumentException error =
            Assert.Throws<ArgumentException>(() => service.Search(Request("no-such-store-xyz")));

        Assert.Contains("was not found in Outlook", error.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("was not found in the local index", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ThatRefusal_NamesTheStoresOutlookHas_NotTheIndexsEmptyCatalog()
    {
        // "Known stores: ." was the third defect: the list came from the index catalog, which
        // on this profile is empty, so the message named nothing and sent the caller to
        // list_accounts - which returns exactly the names that were missing here.
        using MailService service = Service(NothingIndexed());

        ArgumentException error =
            Assert.Throws<ArgumentException>(() => service.Search(Request("no-such-store-xyz")));

        Assert.Contains(IndexedStore, error.Message, StringComparison.Ordinal);
        Assert.Contains(PstStore, error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheRefusal_IsTheSameOneListFoldersUses()
    {
        // One shape for one mistake, across the two tools that can make it (gap G1's
        // precedent): the pure classifier is shared rather than reimplemented, so the two
        // messages cannot drift into two vocabularies for one error.
        string expected = MailService.DescribeUnresolvedFolderStore(
            "no-such-store-xyz", new[] { IndexedStore, PstStore })!;
        using MailService service = Service(NothingIndexed());

        ArgumentException error =
            Assert.Throws<ArgumentException>(() => service.Search(Request("no-such-store-xyz")));

        Assert.StartsWith(expected, error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void WhenOutlookCannotBeReached_TheRefusalSaysTheTwoCasesCannotBeToldApart()
    {
        // The honest third answer. With no index scope AND no profile list there is no
        // evidence either way, and picking one would let a wedged Outlook produce a confident
        // "that store does not exist" about a store that is sitting right there.
        using MailService service = new MailService(
            new ThrowingGateway(), null, NothingIndexed());

        ArgumentException error = Assert.Throws<ArgumentException>(() => service.Search(Request(PstStore)));

        Assert.Contains("Outlook could not be reached", error.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("was not found in Outlook", error.Message, StringComparison.Ordinal);
    }

    // ==================================================================== thread

    [Fact]
    public void Thread_StillReportsScopeWidened_ForAStoreTheIndexCannotAddress()
    {
        // thread resolves its store through the same resolver, and that resolver no longer
        // THROWS for this store - it answers "no scope". The widening flag used to be set in
        // the catch alone, so the fix would have made half of C3 silent again.
        using MailService service = Service(NothingIndexed());

        ThreadOutcome outcome = service.Thread("conv-1", id: null, store: PstStore);

        Assert.True(outcome.ScopeWidened);
        Assert.Contains(
            outcome.Advice!,
            a => a.Contains("did not resolve to an index scope", StringComparison.Ordinal)
                && a.Contains("Windows Search does not index", StringComparison.Ordinal));
    }

    // =================================================================== fixtures

    private static SearchRequest Request(string store)
    {
        return new SearchRequest { Query = "test", Store = store, Top = 25, SnippetChars = 0 };
    }

    private static MailService Service(StubIndexClient index)
    {
        return new MailService(new DirectGateway(ProfileSession.Create(ProfileStores, Sweep)), null, index);
    }

    /// <summary>The measured machine: one profile, nothing of it in the index.</summary>
    private static StubIndexClient NothingIndexed(
        IReadOnlyList<IReadOnlyDictionary<string, object?>>? searchRows = null)
    {
        return new StubIndexClient(Array.Empty<string>(), searchRows);
    }

    /// <summary>The mixed profile: the Exchange store indexed, the data file not.</summary>
    private static StubIndexClient OnlyAliceIndexed(
        IReadOnlyList<IReadOnlyDictionary<string, object?>>? searchRows = null)
    {
        return new StubIndexClient(new[] { IndexedPrefix }, searchRows);
    }

    /// <summary>
    /// One indexed mail, in the store the caller did NOT ask for. Returned for every search
    /// statement, so a scope that was dropped instead of honoured shows up as a hit.
    /// </summary>
    private static IReadOnlyList<IReadOnlyDictionary<string, object?>> OneIndexedMailInAlicesStore()
    {
        return new[]
        {
            (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
            {
                ["System.ItemUrl"] = IndexedPrefix + "/0/Inbox/item-1",
                ["System.Kind"] = new[] { "email" },
                ["System.Message.DateReceived"] = Frontier,
                ["System.Subject"] = "a test mail in the OTHER store",
                ["System.Size"] = 1000L,
            },
        };
    }

    /// <summary>
    /// The sweep the COM tier would perform: the four arrival-path folders of the requested
    /// store, one matching mail in the data file. Shaped exactly as
    /// <c>OutlookComSession.SweepFoldersNewerThan</c> builds it, per-store counters included -
    /// without those a store-scoped search reads its coverage as zero.
    /// </summary>
    private static ComSweepResult Sweep(string? onlyStore)
    {
        string store = onlyStore ?? PstStore;
        return new ComSweepResult(
            new[]
            {
                new ComMailBrief(
                    entryId: "AA" + store.Length.ToString(CultureInfo.InvariantCulture),
                    storeDisplayName: store,
                    storeId: "store-pst",
                    folderName: "Inbox",
                    folderKind: "inbox",
                    subject: "a test mail swept from " + store,
                    senderName: "Bob",
                    senderAddress: "bob@example.com",
                    receivedTime: Frontier,
                    isRead: true,
                    hasAttachments: false,
                    sizeBytes: 2048,
                    body: "test body"),
            },
            foldersSwept: 4,
            foldersSkipped: 0,
            sweptFolders: new[] { store + "/Inbox", store + "/Sent Items", store + "/Deleted Items", store + "/Junk Email" },
            perStore: new[] { new ComStoreSweepCounters(store, foldersSwept: 4, foldersSkipped: 0, foldersFailed: 0, foldersAbsent: 0) });
    }

    /// <summary>
    /// A Windows Search stand-in that knows about a chosen SET of store prefixes and nothing
    /// else - the fixture the whole defect turned on. It answers the three probe statements by
    /// shape (store-discovery sample, newest-received frontier, scope existence) and hands
    /// every remaining statement the caller's scripted search rows.
    /// </summary>
    private sealed class StubIndexClient : IIndexClient
    {
        private const string DiscoveryTail = " System.ItemUrl FROM SystemIndex WHERE System.Kind='email'";

        private readonly IReadOnlyList<string> _knownPrefixes;
        private readonly IReadOnlyList<IReadOnlyDictionary<string, object?>> _searchRows;

        internal StubIndexClient(
            IReadOnlyList<string> knownPrefixes,
            IReadOnlyList<IReadOnlyDictionary<string, object?>>? searchRows)
        {
            _knownPrefixes = knownPrefixes;
            _searchRows = searchRows ?? Array.Empty<IReadOnlyDictionary<string, object?>>();
        }

        public IndexProviderKind Provider => IndexProviderKind.OleDb;

        public IReadOnlyList<IReadOnlyDictionary<string, object?>> ExecuteRows(
            string sql, int maxRows, int? commandTimeoutSeconds = null)
        {
            if (sql.EndsWith(DiscoveryTail, StringComparison.Ordinal))
            {
                return _knownPrefixes
                    .Select(p => (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["System.ItemUrl"] = p + "/0/Inbox/sampled-item",
                    })
                    .ToList();
            }

            if (sql.Contains("System.Message.DateReceived FROM SystemIndex", StringComparison.Ordinal))
            {
                return Known(sql)
                    ? Rows(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["System.Message.DateReceived"] = Frontier,
                    })
                    : Array.Empty<IReadOnlyDictionary<string, object?>>();
            }

            if (sql.StartsWith("SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE", StringComparison.Ordinal))
            {
                return Known(sql)
                    ? Rows(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["System.ItemUrl"] = _knownPrefixes[0] + "/0/Inbox/probed-item",
                    })
                    : Array.Empty<IReadOnlyDictionary<string, object?>>();
            }

            return _searchRows.Count <= maxRows ? _searchRows : _searchRows.Take(maxRows).ToList();
        }

        private static IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows(IReadOnlyDictionary<string, object?> row)
        {
            return new[] { row };
        }

        /// <summary>
        /// Whether a statement's SCOPE names a store this index knows. An UNSCOPED statement
        /// asks about the whole catalog, so it is known exactly when the catalog is non-empty:
        /// that is what makes an empty index report no profile-wide frontier either, which is
        /// the real machine's behaviour and the reason the unscoped search there degraded.
        /// </summary>
        private bool Known(string sql)
        {
            int start = sql.IndexOf("SCOPE='", StringComparison.Ordinal);
            if (start < 0)
            {
                return _knownPrefixes.Count > 0;
            }

            start += "SCOPE='".Length;
            int end = sql.IndexOf('\'', start);
            string scope = end < 0 ? sql.Substring(start) : sql.Substring(start, end - start);

            // Exact store root, or a folder beneath it. Deliberately NOT a bare StartsWith:
            // the delegate probes ask about "<prefix>/1..." subtrees this profile does not
            // have, and answering yes to those would invent a delegate store.
            return _knownPrefixes.Any(p =>
                string.Equals(scope, p, StringComparison.OrdinalIgnoreCase)
                || scope.StartsWith(p + "/0/", StringComparison.OrdinalIgnoreCase));
        }
    }

    /// <summary>Runs operations straight against the stand-in session (no COM host, no pipe).</summary>
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

    /// <summary>An Outlook that cannot be reached at all - the third answer's fixture.</summary>
    private sealed class ThrowingGateway : IComGateway
    {
        public event Action? OutlookGone
        {
            add { }
            remove { }
        }

        public bool IsConnected => false;

        public bool? QuitSinkActive => null;

        public bool ProbeConnected() => false;

        public T Run<T>(Func<IOutlookSession, T> operation) => throw Unavailable();

        public T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery) => throw Unavailable();

        public T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false)
            => throw Unavailable();

        public ComHostDiagnostics GetDiagnostics() => new ComHostDiagnostics("in-process", "down");

        public void Dispose()
        {
        }

        private static OutlookUnavailableException Unavailable()
        {
            return new OutlookUnavailableException("Outlook is not running.");
        }
    }

    /// <summary>
    /// A session that answers the profile's store list and the freshness sweep, and refuses
    /// everything else. A <see cref="DispatchProxy"/> rather than 26 stubs, so a method added
    /// to the contract needs no change here. Not sealed: DispatchProxy derives from its
    /// TProxy at runtime and refuses a sealed one.
    /// </summary>
    private class ProfileSession : DispatchProxy
    {
        private IReadOnlyList<ComStoreDetail> _stores = Array.Empty<ComStoreDetail>();
        private Func<string?, ComSweepResult> _sweep = _ => throw new NotSupportedException();

        internal static IOutlookSession Create(
            IReadOnlyList<ComStoreDetail> stores, Func<string?, ComSweepResult> sweep)
        {
            object proxy = Create<IOutlookSession, ProfileSession>()
                ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
            ((ProfileSession)proxy)._stores = stores;
            ((ProfileSession)proxy)._sweep = sweep;
            return (IOutlookSession)proxy;
        }

        /// <inheritdoc />
        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            return targetMethod?.Name switch
            {
                nameof(IOutlookSession.GetProfileName) => "T1 stand-in profile",
                nameof(IOutlookSession.GetStoreDetails) => _stores,
                // Argument 3 is onlyStoreDisplayName; see IOutlookSession.SweepFoldersNewerThan.
                nameof(IOutlookSession.SweepFoldersNewerThan) => _sweep(args?[3] as string),
                _ => throw new NotSupportedException(
                    "The stand-in session was asked for " + (targetMethod?.Name ?? "an unnamed member")
                    + ", which this test does not model."),
            };
        }
    }
}
