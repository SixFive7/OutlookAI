using System.Globalization;
using System.Reflection;
using System.Text;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Drives the REAL <see cref="MailService"/> search path over a stand-in index and a stand-in
/// Outlook, to pin what the store map changes end to end (Q92/Q99) under the maintainer's Q113 (a)
/// decision - deterministic matching, never a guess - and what it does not.
/// <list type="bullet">
/// <item>A store Outlook reports as not Exchange - a PST, an IMAP or Outlook.com .ost - is scoped to
/// the index store carrying its own NAME AND its HASH, whatever the profile calls it; one no index
/// store carries both for is not indexed, whatever store of its name, or of its hash under another
/// name, the index holds. It never falls back to its name alone - not when the index's store list
/// cannot be read either.</item>
/// <item>Two stores the index cannot tell apart - sharing both name and hash - are refused: a search
/// scoped to one says so, and an unscoped search leaves their index store's rows out and reports
/// them as unmatched. So is a store name two stores of the profile share.</item>
/// <item>THE ONE OPEN EXCEPTION: an Exchange store keeps the name rule exactly as before (Q113 (b)).</item>
/// <item>Index hits come back under the name Outlook gives their store, the one every tool takes.</item>
/// </list>
/// <para>
/// The stand-in index answers the store listing under THIS user's real MAPI root
/// (<see cref="IndexSearchService.CurrentUserMapiRoot"/>), because that is the root the service
/// lists; nothing touches a real index. The PST cases carry a DECOY: an index store whose NAME is
/// the one the caller passes, but whose hash is another store's. The name rule chose the decoy
/// every time.
/// </para>
/// </summary>
public sealed class StoreIndexHashScopeTests
{
    private static readonly string UserRoot = IndexSearchService.CurrentUserMapiRoot();
    private static readonly DateTime Frontier = new(2026, 10, 3, 1, 0, 0, DateTimeKind.Utc);
    private const string IdentityPath = @"C:\OutlookAI-Tier\identity.pst";

    private static string PstStoreId(string path)
    {
        byte[] header = Convert.FromHexString(
            "0000000038A1BB1005E5101AA1BB08002B2A56C200006D737073742E646C6C00000000004E495441F9BFB80100AA0037D96E0000" + "0000");
        return Convert.ToHexString(header.Concat(Encoding.Unicode.GetBytes(path + "\0")).ToArray());
    }

    /// <summary>The index root a store at <paramref name="path"/> is filed under when its own name is <paramref name="ownName"/>.</summary>
    private static string RootOf(string ownName, string path)
        => UserRoot + ownName + "($" + StoreHash.Compute(Convert.FromHexString(PstStoreId(path))).ToString("x", CultureInfo.InvariantCulture) + ")";

    /// <summary>The measured identity store: profile name 'identity@vm.invalid', own name 'Outlook Data File'.</summary>
    private static readonly string IdentityOwnRoot = UserRoot + "Outlook Data File($b25ac20a)";

    /// <summary>The decoy: another store whose own name IS 'identity@vm.invalid' (another profile's).</summary>
    private static readonly string DecoyRoot = UserRoot + "identity@vm.invalid($11111111)";

    private static readonly ComStoreDetail Identity =
        new("identity@vm.invalid", PstStoreId(IdentityPath), 3, null, filePath: IdentityPath, ownName: "Outlook Data File");

    [Fact]
    public void AStoreScopedSearch_IsScopedToTheRootOfItsOwnNameAndHash_NotTheOneItsProfileNameNames()
    {
        var index = new StubIndexClient(new[] { DecoyRoot, IdentityOwnRoot }, IdentityOwnRoot, DecoyRoot);
        using MailService service = Service(index, Identity);

        SearchOutcome outcome = service.Search(Request("identity@vm.invalid"));

        string search = Assert.Single(index.Statements, IsTheSearch);
        Assert.Contains("SCOPE='" + IdentityOwnRoot, search, StringComparison.Ordinal);
        Assert.DoesNotContain(index.Statements, s => s.Contains("SCOPE='" + DecoyRoot, StringComparison.Ordinal));

        // The index hit comes back under Outlook's name for the store, not the index's.
        HitSummary indexed = Assert.Single(outcome.Hits, h => h.Source != "live");
        Assert.Equal("identity@vm.invalid", indexed.Store);
    }

    [Fact]
    public void Control_WithTheHashAbsent_ThePstIsNotIndexed_AndTheDecoyIsNeverSearched()
    {
        // The same profile, but the index holds only the decoy - its name, another hash - with
        // mail, so the name rule finds it. The name rule scoped the search to it and answered with
        // the other store's mail; by name and hash the store is not indexed.
        var index = new StubIndexClient(new[] { DecoyRoot }, DecoyRoot, DecoyRoot);
        using MailService service = Service(index, Identity);

        SearchOutcome outcome = service.Search(Request("identity@vm.invalid"));

        Assert.DoesNotContain(index.Statements, s => s.Contains("SCOPE='" + DecoyRoot, StringComparison.Ordinal));
        Assert.True(outcome.Index!.StoreNotIndexed);
        Assert.Equal(0, outcome.Index.RowsScanned);
        Assert.All(outcome.Hits, h => Assert.Equal("live", h.Source));
    }

    [Fact]
    public void ARenamesLeftoverRoot_TheHashUnderAnotherName_IsNeverSearchedAsTheStore()
    {
        // The store's own hash, but filed under a name that is no longer its own - what a rename
        // leaves until the index files the store again. The hash alone tied the store to it; by
        // name and hash the store is not (yet) indexed, and the leftover is never searched.
        string leftover = UserRoot + "old name($b25ac20a)";
        var index = new StubIndexClient(new[] { leftover }, leftover, leftover);
        using MailService service = Service(index, Identity);

        SearchOutcome outcome = service.Search(Request("identity@vm.invalid"));

        Assert.DoesNotContain(index.Statements, s => s.Contains("SCOPE='" + leftover, StringComparison.Ordinal));
        Assert.True(outcome.Index!.StoreNotIndexed);

        // Control: the store under its own name beside the leftover - searched there.
        index = new StubIndexClient(new[] { leftover, IdentityOwnRoot }, IdentityOwnRoot, leftover);
        using MailService both = Service(index, Identity);
        both.Search(Request("identity@vm.invalid"));
        Assert.Contains("SCOPE='" + IdentityOwnRoot, Assert.Single(index.Statements, IsTheSearch), StringComparison.Ordinal);
    }

    [Fact]
    public void AnUnscopedSearch_ReportsIndexHitsUnderOutlooksStoreName()
    {
        var index = new StubIndexClient(new[] { IdentityOwnRoot }, IdentityOwnRoot, IdentityOwnRoot);
        using MailService service = Service(index, Identity);

        SearchOutcome outcome = service.Search(Unscoped());

        HitSummary indexed = Assert.Single(outcome.Hits, h => h.Source != "live");
        Assert.Equal("identity@vm.invalid", indexed.Store);
        Assert.Null(outcome.Index!.StoresUnmatched);
    }

    [Fact]
    public void AnUnscopedSearch_NamesAPstTheIndexDoesNotHold_EvenWhenAStoreOfItsNameIsIndexed()
    {
        // The decoy carries mail under the PST's name. By name the PST looked indexed - so it was
        // never reported, and its sweep window started at the DECOY's clock. By name and hash it is
        // named as a store the index holds nothing for, and swept from the widest window.
        var index = new StubIndexClient(new[] { DecoyRoot }, DecoyRoot, DecoyRoot);
        using MailService service = Service(index, Identity);

        SearchOutcome outcome = service.Search(Unscoped());

        Assert.Contains("identity@vm.invalid", outcome.Sweep!.StoresWithoutIndex ?? Array.Empty<string>());

        // Control: its own root indexed - not named.
        index = new StubIndexClient(new[] { DecoyRoot, IdentityOwnRoot }, IdentityOwnRoot, DecoyRoot);
        using MailService indexed = Service(index, Identity);
        outcome = indexed.Search(Unscoped());
        Assert.DoesNotContain("identity@vm.invalid", outcome.Sweep!.StoresWithoutIndex ?? Array.Empty<string>());
    }

    [Fact]
    public void AnImapOst_IsHeldToNameAndHash_AndNeverToItsNameAlone()
    {
        // An IMAP store: not Exchange, in an .ost. The index holds a store of its name with mail,
        // under another hash. D55 handed such a store to the name rule; Q113 (a) does not - it is
        // not indexed as far as anything here can establish, and the other store is never searched.
        string ostPath = @"C:\Users\a\AppData\Local\Microsoft\Outlook\carol@example.com.ost";
        var carol = new ComStoreDetail("carol@example.com", PstStoreId(ostPath), 3, null, filePath: ostPath, ownName: "carol@example.com");
        string byNameOnly = UserRoot + "carol@example.com($22222222)";
        var index = new StubIndexClient(new[] { byNameOnly }, byNameOnly, byNameOnly);
        using MailService service = Service(index, carol);

        SearchOutcome outcome = service.Search(Request("carol@example.com"));

        Assert.DoesNotContain(index.Statements, s => s.Contains("SCOPE='" + byNameOnly, StringComparison.Ordinal));
        Assert.True(outcome.Index!.StoreNotIndexed);

        // Control: its own name with its own hash - searched there.
        string own = RootOf("carol@example.com", ostPath);
        index = new StubIndexClient(new[] { byNameOnly, own }, own, byNameOnly);
        using MailService tied = Service(index, carol);
        outcome = tied.Search(Request("carol@example.com"));
        Assert.Contains("SCOPE='" + own, Assert.Single(index.Statements, IsTheSearch), StringComparison.Ordinal);
        Assert.Single(outcome.Hits, h => h.Source != "live");
    }

    [Fact]
    public void TwoStoresSharingNameAndHash_AScopedSearchIsRefused_AndSaysWhy()
    {
        (ComStoreDetail a, ComStoreDetail b, string shared) = TwoStoresTheIndexCannotTellApart();
        var index = new StubIndexClient(new[] { shared, IdentityOwnRoot }, shared, shared);
        using MailService service = Service(index, Identity, a, b);

        ArgumentException refused = Assert.Throws<ArgumentException>(() => service.Search(Request("Archive A")));

        Assert.Contains("'Archive A' cannot be searched through the local index", refused.Message, StringComparison.Ordinal);
        Assert.Contains("'Archive B'", refused.Message, StringComparison.Ordinal);
        Assert.Contains("never by a guess", refused.Message, StringComparison.Ordinal);
        Assert.DoesNotContain(index.Statements, IsTheSearch);

        // Control: the third store of the same profile is searched as usual.
        service.Search(Request("identity@vm.invalid"));
        Assert.Contains("SCOPE='" + IdentityOwnRoot, Assert.Single(index.Statements, IsTheSearch), StringComparison.Ordinal);
    }

    [Fact]
    public void TwoStoresSharingNameAndHash_AnUnscopedSearchLeavesTheirRowsOut_AndReportsThemUnmatched()
    {
        (ComStoreDetail a, ComStoreDetail b, string shared) = TwoStoresTheIndexCannotTellApart();
        var index = new StubIndexClient(new[] { shared, IdentityOwnRoot }, IdentityOwnRoot, IdentityOwnRoot)
        {
            MoreMailRoots = new[] { shared },
        };
        using MailService service = Service(index, Identity, a, b);

        SearchOutcome outcome = service.Search(Unscoped());

        // The index offered a row under each root; the shared one's was dropped, not mixed in.
        Assert.Equal(2, outcome.Index!.RowsScanned);
        Assert.Equal(1, outcome.Index.RowsDropped);
        HitSummary indexed = Assert.Single(outcome.Hits, h => h.Source != "live");
        Assert.Equal("identity@vm.invalid", indexed.Store);
        Assert.Equal(new[] { "Archive A", "Archive B" }, outcome.Index.StoresUnmatched);
        Assert.True(outcome.Degraded);
        Assert.Contains(outcome.Advice ?? Array.Empty<string>(),
            a => a.StartsWith("INCOMPLETE RESULTS - The local index cannot tell 2 store(s)", StringComparison.Ordinal));

        // Control: the same index with only one of the two stores in the profile - tied, its row
        // kept under its name, nothing unmatched.
        index = new StubIndexClient(new[] { shared, IdentityOwnRoot }, IdentityOwnRoot, IdentityOwnRoot)
        {
            MoreMailRoots = new[] { shared },
        };
        using MailService single = Service(index, Identity, a);
        outcome = single.Search(Unscoped());
        Assert.Equal(0, outcome.Index!.RowsDropped);
        Assert.Contains(outcome.Hits, h => h.Source != "live" && h.Store == "Archive A");
        Assert.Null(outcome.Index.StoresUnmatched);
    }

    [Fact]
    public void AStoreNameTwoStoresShare_IsRefused_RatherThanAnsweredFromOneOfThem()
    {
        // Two data files both called 'Shared' in the profile, each indexed under its own name and
        // hash. The index tells them apart; the caller's name cannot.
        string pathA = @"C:\OutlookAI-Tier\shared-a.pst";
        string pathB = @"C:\OutlookAI-Tier\shared-b.pst";
        var a = new ComStoreDetail("Shared", PstStoreId(pathA), 3, null, filePath: pathA, ownName: "Shared");
        var b = new ComStoreDetail("Shared", PstStoreId(pathB), 3, null, filePath: pathB, ownName: "Shared");
        string rootA = RootOf("Shared", pathA);
        string rootB = RootOf("Shared", pathB);
        var index = new StubIndexClient(new[] { rootA, rootB }, rootA, rootA);
        using MailService service = Service(index, a, b);

        ArgumentException refused = Assert.Throws<ArgumentException>(() => service.Search(Request("Shared")));

        Assert.Contains("2 stores in this Outlook profile are named 'Shared'", refused.Message, StringComparison.Ordinal);
        Assert.DoesNotContain(index.Statements, IsTheSearch);
    }

    [Fact]
    public void AnExchangeStoreTheHashDoesNotDecide_IsScopedByTheNameRule_TheOneOpenException()
    {
        // An Exchange store whose documented hash input matches nothing in the index: the miss
        // proves nothing (the input is unmeasured until Q113 (b)), so the store keeps the rule it
        // always had - the index store of its name - rather than being declared unindexed.
        string aliceRoot = UserRoot + "alice@example.com($2468ace0)";
        var alice = new ComStoreDetail("alice@example.com", "00112233", 0, true, false,
            filePath: @"C:\Users\a\alice.ost", mappingSignatureHex: "0A0B0C0D");
        var index = new StubIndexClient(new[] { aliceRoot }, aliceRoot, aliceRoot);
        using MailService service = Service(index, alice);

        SearchOutcome outcome = service.Search(Request("alice@example.com"));

        string search = Assert.Single(index.Statements, IsTheSearch);
        Assert.Contains("SCOPE='" + aliceRoot, search, StringComparison.Ordinal);
        Assert.NotEqual(true, outcome.Index!.StoreNotIndexed);
        Assert.Single(outcome.Hits, h => h.Source != "live");
    }

    [Fact]
    public void WithTheIndexStoreListUnreadable_APstIsRefused_NeverLookedUpByName()
    {
        // The listing fails, so no map is built. The name rule would scope the PST to the decoy -
        // the case Q113 (a) forbids ("for PSTs, never fall back to name-only matching"): refused,
        // saying why, and the decoy is never searched.
        var index = new StubIndexClient(new[] { DecoyRoot, IdentityOwnRoot }, DecoyRoot, DecoyRoot) { FailListing = true };
        using MailService service = Service(index, Identity);

        ArgumentException refused = Assert.Throws<ArgumentException>(() => service.Search(Request("identity@vm.invalid")));

        Assert.Contains("list of stores could not be read", refused.Message, StringComparison.Ordinal);
        Assert.Contains("never by its name alone", refused.Message, StringComparison.Ordinal);
        Assert.DoesNotContain(index.Statements, s => s.Contains("SCOPE='" + DecoyRoot, StringComparison.Ordinal));
    }

    [Fact]
    public void WithTheIndexStoreListUnreadable_AnExchangeStoreKeepsTheNameRule()
    {
        // The open exception, without a map: exactly as before.
        string aliceRoot = UserRoot + "alice@example.com($2468ace0)";
        var alice = new ComStoreDetail("alice@example.com", "00112233", 0, true, false, filePath: @"C:\Users\a\alice.ost");
        var index = new StubIndexClient(new[] { aliceRoot }, aliceRoot, aliceRoot) { FailListing = true };
        using MailService service = Service(index, alice);

        service.Search(Request("alice@example.com"));

        Assert.Contains("SCOPE='" + aliceRoot, Assert.Single(index.Statements, IsTheSearch), StringComparison.Ordinal);
    }

    [Fact]
    public void WhenOutlookStopsAnswering_TheLastStoreListStillMatchesByNameAndHash()
    {
        // Index-only searches are what a closed or wedged Outlook leaves, and the store list Outlook
        // gave last is still the right input: a store's id and own name do not change while it is
        // attached. Without it the map could not be rebuilt, and the PST would fall to the name rule.
        var index = new StubIndexClient(new[] { DecoyRoot, IdentityOwnRoot }, IdentityOwnRoot, DecoyRoot);
        var gateway = new DirectGateway(ProfileSession.Create(new[] { Identity }, Sweep));
        using MailService service = new MailService(gateway, null, index);
        service.Search(Request("identity@vm.invalid"));

        gateway.Down = true;
        ExpireStoreIndexMap(service);
        index.Statements.Clear();
        service.Search(Request("identity@vm.invalid"));

        Assert.Contains("SCOPE='" + IdentityOwnRoot, Assert.Single(index.Statements, IsTheSearch), StringComparison.Ordinal);
        Assert.DoesNotContain(index.Statements, s => s.Contains("SCOPE='" + DecoyRoot, StringComparison.Ordinal));
    }

    // =================================================================== fixtures

    /// <summary>
    /// Two data files that share both their own name ('Archive') and their hash - the same entry
    /// ID, as a stand-in for a 32-bit collision under one name - and the one index store both claim.
    /// </summary>
    private static (ComStoreDetail A, ComStoreDetail B, string SharedRoot) TwoStoresTheIndexCannotTellApart()
    {
        string path = @"C:\OutlookAI-Tier\archive.pst";
        return (
            new ComStoreDetail("Archive A", PstStoreId(path), 3, null, filePath: path, ownName: "Archive"),
            new ComStoreDetail("Archive B", PstStoreId(path), 3, null, filePath: @"C:\OutlookAI-Tier\elsewhere\archive.pst", ownName: "Archive"),
            RootOf("Archive", path));
    }

    /// <summary>Makes the next store-map read rebuild it, as five minutes passing would.</summary>
    private static void ExpireStoreIndexMap(MailService service)
    {
        FieldInfo built = typeof(MailService).GetField("_storeIndexMapBuiltUtc", BindingFlags.NonPublic | BindingFlags.Instance)
            ?? throw new InvalidOperationException("MailService no longer has _storeIndexMapBuiltUtc; update this fixture.");
        built.SetValue(service, DateTime.MinValue);
    }

    private static bool IsTheSearch(string sql)
        => sql.Contains("ORDER BY System.Message.DateReceived DESC", StringComparison.Ordinal)
        && sql.Contains("CONTAINS", StringComparison.Ordinal);

    private static SearchRequest Request(string store)
    {
        return new SearchRequest { Query = "test", Store = store, Top = 25, SnippetChars = 0 };
    }

    private static SearchRequest Unscoped()
    {
        return new SearchRequest { Query = "test", Top = 25, SnippetChars = 0 };
    }

    private static MailService Service(StubIndexClient index, params ComStoreDetail[] stores)
    {
        return new MailService(new DirectGateway(ProfileSession.Create(stores, Sweep)), null, index);
    }

    private static ComSweepResult Sweep(string? onlyStore)
    {
        string store = onlyStore ?? "identity@vm.invalid";
        return new ComSweepResult(
            new[]
            {
                new ComMailBrief(
                    entryId: "AA" + store.Length.ToString(CultureInfo.InvariantCulture),
                    storeDisplayName: store,
                    storeId: "live-store",
                    folderName: "Inbox",
                    folderKind: "inbox",
                    subject: "a test mail swept from " + store,
                    senderName: "Bob",
                    senderAddress: "bob@example.com",
                    receivedTime: Frontier.AddMinutes(-90),
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
    /// A Windows Search stand-in holding a chosen set of store roots, with indexed mail under
    /// <c>mailRoot</c> (and <see cref="MoreMailRoots"/>), and a discovery sample - the old name
    /// rule's catalog - drawn from <c>sampleRoot</c>. It answers the listing, the sample, the
    /// frontier and existence probes by shape, and the search itself only for a statement scoped to
    /// a root mail is in, or for an unscoped one - one row per mail root - so a search scoped
    /// anywhere else finds nothing, as a real index would.
    /// </summary>
    private sealed class StubIndexClient : IIndexClient
    {
        private readonly IReadOnlyList<string> _roots;
        private readonly string _mailRoot;
        private readonly string _sampleRoot;

        internal StubIndexClient(IReadOnlyList<string> roots, string mailRoot, string sampleRoot)
        {
            _roots = roots;
            _mailRoot = mailRoot;
            _sampleRoot = sampleRoot;
        }

        public bool FailListing { get; init; }

        public IReadOnlyList<string> MoreMailRoots { get; init; } = Array.Empty<string>();

        public List<string> Statements { get; } = new();

        public IndexProviderKind Provider => IndexProviderKind.OleDb;

        public IReadOnlyList<IReadOnlyDictionary<string, object?>> ExecuteRows(string sql, int maxRows, int? commandTimeoutSeconds = null)
        {
            Statements.Add(sql);
            if (sql.Contains("DIRECTORY='", StringComparison.Ordinal))
            {
                if (FailListing)
                {
                    throw new InvalidOperationException("listing refused");
                }

                return _roots.Select(r => Row(("System.ItemUrl", r))).ToList();
            }

            if (sql.StartsWith("SELECT TOP " + maxRows.ToString(CultureInfo.InvariantCulture) + " System.ItemUrl FROM SystemIndex WHERE System.Kind='email'", StringComparison.Ordinal))
            {
                // The discovery sample behind the name rule's catalog.
                return new[] { Row(("System.ItemUrl", _sampleRoot + "/0/Inbox/sampled-item")) };
            }

            List<string> mailRoots = new[] { _mailRoot }.Concat(MoreMailRoots).ToList();
            string? scopedTo = mailRoots.FirstOrDefault(r =>
                sql.Contains("SCOPE='" + r + "'", StringComparison.Ordinal)
                || sql.Contains("SCOPE='" + r + "/", StringComparison.Ordinal));
            bool unscoped = !sql.Contains("SCOPE='", StringComparison.Ordinal);
            if (sql.Contains("System.Message.DateReceived FROM SystemIndex", StringComparison.Ordinal))
            {
                return scopedTo != null || unscoped
                    ? new[] { Row(("System.Message.DateReceived", Frontier)) }
                    : Array.Empty<IReadOnlyDictionary<string, object?>>();
            }

            if (sql.StartsWith("SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE", StringComparison.Ordinal))
            {
                return scopedTo != null && !sql.Contains(scopedTo + "/1'", StringComparison.Ordinal)
                    ? new[] { Row(("System.ItemUrl", scopedTo + "/0/Inbox/probed-item")) }
                    : Array.Empty<IReadOnlyDictionary<string, object?>>();
            }

            if (sql.Contains("CONTAINS", StringComparison.Ordinal) && (scopedTo != null || unscoped))
            {
                return (unscoped ? mailRoots : new List<string> { scopedTo! })
                    .Select((root, i) => Row(
                        ("System.ItemUrl", root + "/0/Inbox/item-" + (i + 1).ToString(CultureInfo.InvariantCulture)),
                        ("System.Kind", new[] { "email" }),
                        ("System.Message.DateReceived", Frontier.AddMinutes(-10 - i)),
                        ("System.Subject", "an indexed test mail"),
                        ("System.Size", 1000L)))
                    .ToList();
            }

            return Array.Empty<IReadOnlyDictionary<string, object?>>();
        }

        private static IReadOnlyDictionary<string, object?> Row(params (string Key, object? Value)[] cells)
        {
            var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            foreach ((string key, object? value) in cells)
            {
                row[key] = value;
            }

            return row;
        }
    }

    /// <summary>Runs operations straight against the stand-in session (no COM host, no pipe) - or, when <see cref="Down"/>, refuses as an unreachable Outlook does.</summary>
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

        /// <summary>When true, every operation fails as it does when Outlook is closed or wedged.</summary>
        public bool Down { get; set; }

        public bool IsConnected => !Down;

        public bool? QuitSinkActive => null;

        public bool ProbeConnected() => !Down;

        public T Run<T>(Func<IOutlookSession, T> operation) => Invoke(operation);

        public T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery) => Invoke(operation);

        public T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false)
            => Invoke(operation);

        public ComHostDiagnostics GetDiagnostics() => new ComHostDiagnostics("in-process", "ready");

        public void Dispose()
        {
        }

        private T Invoke<T>(Func<IOutlookSession, T> operation)
        {
            if (Down)
            {
                throw new OutlookUnavailableException("Outlook is not answering (T1 stand-in).");
            }

            return operation(_session);
        }
    }

    /// <summary>A session answering the store list and the sweep, refusing everything else. Not sealed: DispatchProxy.</summary>
    private class ProfileSession : DispatchProxy
    {
        private IReadOnlyList<ComStoreDetail> _stores = Array.Empty<ComStoreDetail>();
        private Func<string?, ComSweepResult> _sweep = _ => throw new NotSupportedException();

        internal static IOutlookSession Create(IReadOnlyList<ComStoreDetail> stores, Func<string?, ComSweepResult> sweep)
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
                nameof(IOutlookSession.SweepFoldersNewerThan) => _sweep(args?[3] as string),
                _ => throw new NotSupportedException(
                    "The stand-in session was asked for " + (targetMethod?.Name ?? "an unnamed member")
                    + ", which this test does not model."),
            };
        }
    }
}
