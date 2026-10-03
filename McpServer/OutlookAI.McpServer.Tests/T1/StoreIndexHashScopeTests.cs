using System.Globalization;
using System.Reflection;
using System.Text;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Drives the REAL <see cref="MailService"/> search path over a stand-in index and a stand-in
/// Outlook, to pin what the store map changes end to end (Q92/Q99) - and what it does not.
/// <list type="bullet">
/// <item>A PST - the measured family - is scoped to the index store whose <c>($hash)</c> is its
/// own, whatever either side calls it; a PST no index store carries the hash of is not indexed,
/// whatever store of its name the index holds.</item>
/// <item>A store the hash does not decide - here an Exchange store, whose hash input is
/// unmeasured - is resolved by the name rule exactly as before, and so is every store when no
/// map can be built. Never worse than before is the contract for those.</item>
/// <item>Index hits come back under the name Outlook gives their store, the one every tool
/// takes.</item>
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

    /// <summary>The measured identity store: profile name 'identity@vm.invalid', own name 'Outlook Data File'.</summary>
    private static readonly string IdentityOwnRoot = UserRoot + "Outlook Data File($b25ac20a)";

    /// <summary>The decoy: another store whose own name IS 'identity@vm.invalid' (another profile's).</summary>
    private static readonly string DecoyRoot = UserRoot + "identity@vm.invalid($11111111)";

    private static readonly ComStoreDetail Identity =
        new("identity@vm.invalid", PstStoreId(IdentityPath), 3, null, filePath: IdentityPath);

    [Fact]
    public void AStoreScopedSearch_IsScopedToTheRootItsHashNames_NotTheOneItsNameNames()
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
        // The same profile, but the index holds only the decoy - with mail, so the name rule
        // finds it. The name rule scoped the search to it and answered with the other store's
        // mail; by hash the store is not indexed.
        var index = new StubIndexClient(new[] { DecoyRoot }, DecoyRoot, DecoyRoot);
        using MailService service = Service(index, Identity);

        SearchOutcome outcome = service.Search(Request("identity@vm.invalid"));

        Assert.DoesNotContain(index.Statements, s => s.Contains("SCOPE='" + DecoyRoot, StringComparison.Ordinal));
        Assert.True(outcome.Index!.StoreNotIndexed);
        Assert.Equal(0, outcome.Index.RowsScanned);
        Assert.All(outcome.Hits, h => Assert.Equal("live", h.Source));
    }

    [Fact]
    public void AnUnscopedSearch_ReportsIndexHitsUnderOutlooksStoreName()
    {
        var index = new StubIndexClient(new[] { IdentityOwnRoot }, IdentityOwnRoot, IdentityOwnRoot);
        using MailService service = Service(index, Identity);

        SearchOutcome outcome = service.Search(new SearchRequest { Query = "test", Top = 25, SnippetChars = 0 });

        HitSummary indexed = Assert.Single(outcome.Hits, h => h.Source != "live");
        Assert.Equal("identity@vm.invalid", indexed.Store);
    }

    [Fact]
    public void AnUnscopedSearch_NamesAPstTheIndexDoesNotHold_EvenWhenAStoreOfItsNameIsIndexed()
    {
        // The decoy carries mail under the PST's name. By name the PST looked indexed - so it was
        // never reported, and its sweep window started at the DECOY's clock. By hash it is named
        // as a store the index holds nothing for, and swept from the widest window.
        var index = new StubIndexClient(new[] { DecoyRoot }, DecoyRoot, DecoyRoot);
        using MailService service = Service(index, Identity);

        SearchOutcome outcome = service.Search(new SearchRequest { Query = "test", Top = 25, SnippetChars = 0 });

        Assert.Contains("identity@vm.invalid", outcome.Sweep!.StoresWithoutIndex ?? Array.Empty<string>());

        // Control: its own root indexed - not named.
        index = new StubIndexClient(new[] { DecoyRoot, IdentityOwnRoot }, IdentityOwnRoot, DecoyRoot);
        using MailService indexed = Service(index, Identity);
        outcome = indexed.Search(new SearchRequest { Query = "test", Top = 25, SnippetChars = 0 });
        Assert.DoesNotContain("identity@vm.invalid", outcome.Sweep!.StoresWithoutIndex ?? Array.Empty<string>());
    }

    [Fact]
    public void AStoreTheHashDoesNotDecide_IsScopedByTheNameRule_ExactlyAsBefore()
    {
        // An Exchange store whose documented hash input matches nothing in the index: the miss
        // proves nothing (the input is unmeasured), so the store keeps the rule it always had -
        // the index store of its name - rather than being declared unindexed.
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
    public void WithNoStoreMap_EveryStoreKeepsTheNameRule()
    {
        // The listing fails, so no map is built - and nothing changes from before it existed:
        // the PST is resolved by name, to the decoy, as the name rule always did. This is the
        // fallback's contract (never worse than before), and the case the map exists to fix.
        var index = new StubIndexClient(new[] { DecoyRoot, IdentityOwnRoot }, DecoyRoot, DecoyRoot) { FailListing = true };
        using MailService service = Service(index, Identity);

        service.Search(Request("identity@vm.invalid"));

        string search = Assert.Single(index.Statements, IsTheSearch);
        Assert.Contains("SCOPE='" + DecoyRoot, search, StringComparison.Ordinal);
    }

    // =================================================================== fixtures

    private static bool IsTheSearch(string sql)
        => sql.Contains("ORDER BY System.Message.DateReceived DESC", StringComparison.Ordinal)
        && sql.Contains("CONTAINS", StringComparison.Ordinal);

    private static SearchRequest Request(string store)
    {
        return new SearchRequest { Query = "test", Store = store, Top = 25, SnippetChars = 0 };
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
    /// <c>mailRoot</c>, and a discovery sample - the old name rule's catalog - drawn from
    /// <c>sampleRoot</c>. It answers the listing, the sample, the frontier and existence probes by
    /// shape, and the search itself only for a statement scoped to the root the mail is in - so a
    /// search scoped anywhere else finds nothing, as a real index would.
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

            bool scopedToMail = sql.Contains("SCOPE='" + _mailRoot + "'", StringComparison.Ordinal)
                || sql.Contains("SCOPE='" + _mailRoot + "/", StringComparison.Ordinal);
            bool unscoped = !sql.Contains("SCOPE='", StringComparison.Ordinal);
            if (sql.Contains("System.Message.DateReceived FROM SystemIndex", StringComparison.Ordinal))
            {
                return scopedToMail || unscoped
                    ? new[] { Row(("System.Message.DateReceived", Frontier)) }
                    : Array.Empty<IReadOnlyDictionary<string, object?>>();
            }

            if (sql.StartsWith("SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE", StringComparison.Ordinal))
            {
                return scopedToMail && !sql.Contains(_mailRoot + "/1'", StringComparison.Ordinal)
                    ? new[] { Row(("System.ItemUrl", _mailRoot + "/0/Inbox/probed-item")) }
                    : Array.Empty<IReadOnlyDictionary<string, object?>>();
            }

            if (sql.Contains("CONTAINS", StringComparison.Ordinal) && (scopedToMail || unscoped))
            {
                return new[]
                {
                    Row(
                        ("System.ItemUrl", _mailRoot + "/0/Inbox/item-1"),
                        ("System.Kind", new[] { "email" }),
                        ("System.Message.DateReceived", Frontier.AddMinutes(-10)),
                        ("System.Subject", "an indexed test mail"),
                        ("System.Size", 1000L)),
                };
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
