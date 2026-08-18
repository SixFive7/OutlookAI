using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Gap C4's remaining half: <c>thread</c>'s live walk covers ONE store, and until now the
/// only thing that could say so was <c>unwalked_store</c> - which is raised from the stores
/// the conversation's INDEX ROWS name.
/// <para>
/// SO IT WAS BLIND EXACTLY WHERE THE INDEX IS. Point it at the profile shape this product
/// keeps meeting - one indexed mailbox plus a data file Windows Search has never opened - and
/// it says nothing at all: the data file contributes no rows to compare against, the walk
/// covers the anchor's store, and half a conversation can be missing under
/// <c>freshness: "live"</c> with <c>degraded</c> absent. That is the unindexed-PST profile
/// the whole A-group of the audit is about, so the case the code could not see is the case
/// it was most needed for.
/// </para>
/// <para>
/// The fix asks OUTLOOK which stores exist rather than asking the index about itself, and the
/// verdict per store is the same pure <see cref="MailService.StoresMissingFromIndex"/> that
/// <c>2d28957</c> and <c>3bd512f</c> built for the sweep - one rule in this server for "the
/// index holds nothing for this store", not two.
/// </para>
/// <para>
/// Driven through the real <see cref="MailService"/> against a stand-in Outlook session and a
/// stand-in index client whose catalog holds one of the profile's two stores. No mailbox and
/// no Windows Search index are touched. What T1 CANNOT reach, and needs a live profile: that
/// Outlook's Conversation object really does stop at the anchor's store.
/// </para>
/// </summary>
public sealed class ThreadUnindexedStoreTests
{
    private const string Sid = "{S-1-5-21-1111111111-2222222222-3333333333-1001}";

    private const string IndexedStore = "alice@example.com";

    private const string IndexedPrefix = "mapi16://" + Sid + "/" + IndexedStore + "($deadbeef)";

    private const string PstStore = "Archive 2019.pst";

    private const string PstPrefix = "mapi16://" + Sid + "/" + PstStore + "($feedface)";

    /// <summary>A raw EntryID hex: long enough to be accepted without any hit cache or COM locate.</summary>
    private const string AnchorId = "00000000AABBCCDDEEFF00112233445566778899AABBCCDD";

    private static readonly DateTime Frontier = new(2026, 8, 18, 9, 30, 0, DateTimeKind.Utc);

    // ======================================================= the pure rule (the exclusion)

    [Fact]
    public void AnUnindexedStoreOtherThanTheAnchors_IsAHoleTheWalkLeft()
    {
        IReadOnlyList<string> unwalked = FreshMerge.UnwalkedUnindexedStores(
            Walked(IndexedStore), new[] { PstStore });

        Assert.Equal(new[] { PstStore }, unwalked);
    }

    [Fact]
    public void TheAnchorsOwnStore_IsNeverAHole_EvenWhenTheIndexHasNothingForIt()
    {
        // The single-PST profile, and the cry-wolf case that decides whether this flag is
        // worth anything: Outlook enumerated the conversation THERE, member by member, so
        // its coverage is complete whatever the index holds.
        Assert.Empty(FreshMerge.UnwalkedUnindexedStores(Walked(PstStore), new[] { PstStore }));
    }

    [Fact]
    public void TheAnchorIsMatchedWithoutRegardToCase()
    {
        Assert.Empty(FreshMerge.UnwalkedUnindexedStores(
            Walked("ARCHIVE 2019.PST"), new[] { PstStore }));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void AWalkThatCouldNotNameItsStore_ClaimsNothing(string? anchor)
    {
        // Without the anchor the exclusion above cannot be applied, and a list built without
        // it would name the very store that WAS walked - worse than saying nothing.
        Assert.Empty(FreshMerge.UnwalkedUnindexedStores(Walked(anchor), new[] { PstStore }));
    }

    [Fact]
    public void AWalkThatFoundNoMembers_ClaimsNothing()
    {
        ThreadLiveInfo live = Walked(IndexedStore);
        live.MembersWalked = 0;

        Assert.Empty(FreshMerge.UnwalkedUnindexedStores(live, new[] { PstStore }));
    }

    [Fact]
    public void AWalkThatNeverRan_ClaimsNothing()
    {
        // "Did not run" is index-only, a state with its own remedy - the same split the
        // sweep makes, and folding it in here would blur two different answers.
        ThreadLiveInfo live = new ThreadLiveInfo { Performed = false, Error = "NoAnchorItem" };

        Assert.Empty(FreshMerge.UnwalkedUnindexedStores(live, new[] { PstStore }));
    }

    [Fact]
    public void NoUnindexedStoreAnywhere_IsNoHole()
    {
        Assert.Empty(FreshMerge.UnwalkedUnindexedStores(Walked(IndexedStore), Array.Empty<string>()));
        Assert.Empty(FreshMerge.UnwalkedUnindexedStores(Walked(IndexedStore), null));
    }

    // ============================================== the code, the verdict and the sentence

    [Fact]
    public void TheNamedStores_RaiseTheCode_AndMakeTheLookupPartial()
    {
        ThreadLiveInfo live = Walked(IndexedStore);
        live.StoresWithoutIndex = new[] { PstStore };

        Assert.Contains(
            FreshMerge.ThreadGapUnindexedStore,
            FreshMerge.DescribeThreadCoverageGaps(live, new[] { IndexedStore })!);
        Assert.Equal(
            FreshMerge.FreshnessPartial,
            FreshMerge.ClassifyThreadFreshness(live, new[] { IndexedStore }));
    }

    [Fact]
    public void TheSentence_NamesTheStore_TheWalkedStore_AndWithdrawsTheStrongerClaim()
    {
        ThreadLiveInfo live = Walked(IndexedStore);
        live.StoresWithoutIndex = new[] { PstStore };
        live.CoverageGaps = new[] { FreshMerge.ThreadGapUnindexedStore };

        string line = Assert.Single(MailService.DescribeThreadCoverage(
            live, FreshMerge.FreshnessPartial, store: null, scopeWidened: false, top: 50)!);

        Assert.Contains(PstStore, line, StringComparison.Ordinal);
        Assert.Contains(IndexedStore, line, StringComparison.Ordinal);
        Assert.Contains("TELL THE USER", line, StringComparison.Ordinal);

        // The weaker claim is the honest one: this is "we could not ask", not "we asked and
        // there was nothing there". An agent that reports the second has invented a fact.
        Assert.Contains("cannot be established", line, StringComparison.Ordinal);
    }

    [Fact]
    public void TheCappedList_SaysHowManyMoreThereWere()
    {
        // The Q7b rule, applied to the second block that carries this list: naming 12 of 30
        // stores and stopping reads as "these twelve".
        ThreadLiveInfo live = Walked(IndexedStore);
        live.StoresWithoutIndex = Enumerable.Range(1, MailService.UnindexedStoreListCap)
            .Select(i => "Archive " + i + ".pst").ToList();
        live.StoresWithoutIndexTruncated = true;
        live.StoresWithoutIndexTotal = 30;
        live.CoverageGaps = new[] { FreshMerge.ThreadGapUnindexedStore };

        string line = Assert.Single(MailService.DescribeThreadCoverage(
            live, FreshMerge.FreshnessPartial, store: null, scopeWidened: false, top: 50)!);

        Assert.Contains("and 18 more", line, StringComparison.Ordinal);
        Assert.Contains("live.storesWithoutIndexTotal", line, StringComparison.Ordinal);
    }

    // ============================================================ through the whole tool

    [Fact]
    public void AConversationWalkedOnAMixedProfile_NamesTheStoreNeitherTierCovers()
    {
        // The shipped defect, end to end: one indexed mailbox, one data file the index has
        // never opened, a conversation walked in the mailbox. The answer used to be
        // freshness "live" with degraded absent and no field naming the data file.
        using MailService service = Service(OnlyAliceIndexed());

        ThreadOutcome outcome = service.Thread(conversationId: null, id: AnchorId, store: null);

        Assert.True(outcome.Live!.Performed);
        Assert.Equal(new[] { PstStore }, outcome.Live.StoresWithoutIndex);
        Assert.Contains(FreshMerge.ThreadGapUnindexedStore, outcome.Live.CoverageGaps!);
        Assert.Equal(FreshMerge.FreshnessPartial, outcome.Freshness);
        Assert.True(outcome.Degraded);
        Assert.Contains(
            outcome.Advice!,
            a => a.Contains(PstStore, StringComparison.Ordinal));
    }

    [Fact]
    public void AFullyIndexedProfile_RaisesNothing()
    {
        // The flag that matters is the one that does not cry wolf: every store here is in
        // the index, so nothing is outside both tiers and the walk answers live.
        using MailService service = Service(EverythingIndexed());

        ThreadOutcome outcome = service.Thread(conversationId: null, id: AnchorId, store: null);

        Assert.Null(outcome.Live!.StoresWithoutIndex);
        Assert.Null(outcome.Live.CoverageGaps);
        Assert.Equal(FreshMerge.FreshnessLive, outcome.Freshness);
        Assert.Null(outcome.Degraded);
    }

    [Fact]
    public void AProfileOfNothingButTheWalkedStore_RaisesNothing()
    {
        // An all-PST profile of ONE store: the index holds nothing for it, and the walk
        // covered it completely. Reporting it would degrade every thread on such a machine.
        using MailService service = Service(
            NothingIndexed(), new[] { new ComStoreDetail(PstStore, "store-pst", 3, null) }, walkedStore: PstStore);

        ThreadOutcome outcome = service.Thread(conversationId: null, id: AnchorId, store: null);

        Assert.Null(outcome.Live!.StoresWithoutIndex);
        Assert.Equal(FreshMerge.FreshnessLive, outcome.Freshness);
    }

    [Fact]
    public void AStoreListWithNothingInIt_MakesNoClaimEitherWay()
    {
        // Silence means "not established", never "indexed" - the same rule the sweep's
        // naming pass follows. The store list is the only evidence, and an empty one is no
        // evidence; a null one (Outlook unreachable) takes the same branch of
        // StoresMissingFromIndex, which is why one fixture covers both.
        using MailService service = Service(OnlyAliceIndexed(), profileStores: Array.Empty<ComStoreDetail>());

        ThreadOutcome outcome = service.Thread(conversationId: null, id: AnchorId, store: null);

        Assert.Null(outcome.Live!.StoresWithoutIndex);
        Assert.Equal(FreshMerge.FreshnessLive, outcome.Freshness);
    }

    [Fact]
    public void ManyUnindexedStores_AreCappedAndTheCutIsReported()
    {
        List<ComStoreDetail> stores = new List<ComStoreDetail> { new(IndexedStore, "store-alice", 0, true) };
        for (int i = 1; i <= MailService.UnindexedStoreListCap + 3; i++)
        {
            stores.Add(new ComStoreDetail("Archive " + i + ".pst", "store-" + i, 3, null));
        }

        using MailService service = Service(OnlyAliceIndexed(), stores);

        ThreadOutcome outcome = service.Thread(conversationId: null, id: AnchorId, store: null);

        Assert.Equal(MailService.UnindexedStoreListCap, outcome.Live!.StoresWithoutIndex!.Count);
        Assert.True(outcome.Live.StoresWithoutIndexTruncated);
        Assert.Equal(MailService.UnindexedStoreListCap + 3, outcome.Live.StoresWithoutIndexTotal);
    }

    // ================================================================== fixtures

    private static ThreadLiveInfo Walked(string? anchorStore)
    {
        return new ThreadLiveInfo
        {
            Performed = true,
            MembersWalked = 3,
            MembersAdded = 3,
            AnchorStore = anchorStore,
        };
    }

    private static readonly ComStoreDetail[] MixedProfile =
    {
        new(IndexedStore, "store-alice", 0, true),
        new(PstStore, "store-pst", 3, null),
    };

    private static MailService Service(
        StubIndexClient index,
        IReadOnlyList<ComStoreDetail>? profileStores = null,
        string walkedStore = IndexedStore)
    {
        return new MailService(
            new DirectGateway(WalkSession.Create(profileStores ?? MixedProfile, walkedStore)), null, index);
    }

    private static StubIndexClient OnlyAliceIndexed() => new(new[] { IndexedPrefix });

    private static StubIndexClient EverythingIndexed() => new(new[] { IndexedPrefix, PstPrefix });

    private static StubIndexClient NothingIndexed() => new(Array.Empty<string>());

    /// <summary>
    /// A Windows Search stand-in that knows a chosen SET of store prefixes and nothing else,
    /// answering the probe statements by shape. Copied in spirit from
    /// <c>UnindexedStoreReportingTests</c>: this file is about what the answer SAYS.
    /// </summary>
    private sealed class StubIndexClient : IIndexClient
    {
        private const string DiscoveryTail = " System.ItemUrl FROM SystemIndex WHERE System.Kind='email'";

        private readonly IReadOnlyList<string> _knownPrefixes;

        internal StubIndexClient(IReadOnlyList<string> knownPrefixes)
        {
            _knownPrefixes = knownPrefixes;
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

            return Array.Empty<IReadOnlyDictionary<string, object?>>();
        }

        private static IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows(IReadOnlyDictionary<string, object?> row)
        {
            return new[] { row };
        }

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

    /// <summary>
    /// A session that answers the profile's store list and one conversation walk, and refuses
    /// everything else. A <see cref="DispatchProxy"/> so a member added to the contract needs
    /// no change here. Not sealed: DispatchProxy derives from its TProxy at runtime.
    /// </summary>
    private class WalkSession : DispatchProxy
    {
        private IReadOnlyList<ComStoreDetail> _stores = Array.Empty<ComStoreDetail>();
        private string _walkedStore = IndexedStore;

        internal static IOutlookSession Create(IReadOnlyList<ComStoreDetail> stores, string walkedStore)
        {
            object proxy = Create<IOutlookSession, WalkSession>()
                ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
            ((WalkSession)proxy)._stores = stores;
            ((WalkSession)proxy)._walkedStore = walkedStore;
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

                case nameof(IOutlookSession.TryGetConversationItems):
                    return Members();

                default:
                    throw new NotSupportedException(
                        "The stand-in session was asked for " + (targetMethod?.Name ?? "an unnamed member")
                        + ", which this test does not model.");
            }
        }

        private IReadOnlyList<ComMailBrief> Members()
        {
            return Enumerable.Range(1, 3)
                .Select(i => new ComMailBrief(
                    entryId: "MEMBER" + i,
                    storeDisplayName: _walkedStore,
                    storeId: "store-walked",
                    folderName: "Inbox",
                    folderKind: "inbox",
                    subject: "a thread member",
                    senderName: "Bob",
                    senderAddress: "bob@example.com",
                    receivedTime: Frontier.AddMinutes(i),
                    isRead: true,
                    hasAttachments: false,
                    sizeBytes: 2048,
                    body: null))
                .ToList();
        }
    }
}
