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
/// Gap C5: a <c>store</c> on <c>thread</c> scopes the INDEX query, so it narrows the very
/// evidence <c>unwalked_store</c> is computed from - the parameter silences the code that
/// would have reported what the parameter cost.
/// <para>
/// THE SHAPE, and why it is worse than it sounds. <c>unwalked_store</c> fires when the
/// conversation's index rows name a store the live walk did not cover. Scope that query to one
/// store and its rows can only ever name that store, so the code cannot fire at all: two
/// INDEXED accounts, a conversation reaching into both, and the answer comes back
/// <c>freshness: "live"</c> with the second account's members absent and unmentioned. C4's fix
/// covers the half where the other store is UNINDEXED, because that pass reads Outlook's store
/// list rather than the index rows; this is the half where it is indexed and simply scoped out.
/// </para>
/// <para>
/// AND NOBODY ASKED FOR IT. <c>thread</c> DERIVES the store from the referenced hit whenever
/// <c>id</c> is passed without <c>conversation_id</c> - the shape an agent reaches for straight
/// out of a search - so the narrowing, and the silence, both arrive unrequested. That is why
/// the remedy the sentence names depends on how the store arrived: a scope the caller chose is
/// dropped, a derived one is cleared by passing <c>conversation_id</c> beside <c>id</c>.
/// </para>
/// <para>
/// The fix asks OUTLOOK which stores the profile has, exactly as C4's does and for exactly the
/// same reason: the thing a scope suppressed cannot be used to detect what the scope
/// suppressed. Driven through the real <see cref="MailService"/> against a stand-in session and
/// a stand-in index client. No mailbox and no Windows Search index are touched.
/// </para>
/// </summary>
public sealed class ThreadScopedStoreTests
{
    private const string Sid = "{S-1-5-21-1111111111-2222222222-3333333333-1001}";

    private const string AliceStore = "alice@example.com";

    private const string AlicePrefix = "mapi16://" + Sid + "/" + AliceStore + "($deadbeef)";

    /// <summary>A second mailbox the index knows all about - the case C4's fix cannot see.</summary>
    private const string BobStore = "bob@example.com";

    private const string BobPrefix = "mapi16://" + Sid + "/" + BobStore + "($c0ffee00)";

    private const string PstStore = "Archive 2019.pst";

    /// <summary>A raw EntryID hex: long enough to be accepted without any hit cache or COM locate.</summary>
    private const string AnchorId = "00000000AABBCCDDEEFF00112233445566778899AABBCCDD";

    private static readonly DateTime Frontier = new(2026, 8, 18, 9, 30, 0, DateTimeKind.Utc);

    private static readonly ComStoreDetail[] TwoMailboxes =
    {
        new ComStoreDetail(AliceStore, "store-alice", 0, true),
        new ComStoreDetail(BobStore, "store-bob", 0, true),
    };

    // ================================================================= the pure rule

    [Fact]
    public void AProfileStoreTheScopedQueryCouldNotAskAbout_IsAHole()
    {
        Assert.Equal(
            new[] { BobStore },
            FreshMerge.StoresScopedOutOfThreadLookup(
                Walked(AliceStore), AliceStore, new[] { AliceStore, BobStore }, null));
    }

    [Fact]
    public void TheStoreTheQueryWasScopedTo_IsNeverAHole()
    {
        // It is the one store that WAS asked about, which is the whole point of the scope.
        Assert.DoesNotContain(
            AliceStore,
            FreshMerge.StoresScopedOutOfThreadLookup(
                Walked(AliceStore), AliceStore, new[] { AliceStore, BobStore }, null));
    }

    [Fact]
    public void TheStoreTheWalkCovered_IsNeverAHole()
    {
        // Outlook enumerated the conversation there member by member, so its coverage is
        // complete whatever the index was or was not asked. Scope and anchor can differ: a
        // caller may name one store and pass an id from another.
        Assert.Equal(
            Array.Empty<string>(),
            FreshMerge.StoresScopedOutOfThreadLookup(
                Walked(BobStore), AliceStore, new[] { AliceStore, BobStore }, null));
    }

    [Fact]
    public void BothExclusionsMatchWithoutRegardToCase()
    {
        Assert.Empty(FreshMerge.StoresScopedOutOfThreadLookup(
            Walked("ALICE@EXAMPLE.COM"), "Bob@Example.Com", new[] { AliceStore, BobStore }, null));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void AnUnscopedQuery_ClaimsNothing(string? scope)
    {
        // Nothing was suppressed, so the STRONGER code speaks for itself: unwalked_store says
        // the conversation demonstrably has members elsewhere, and this one only ever says the
        // question could not be asked. Preferring the weaker reading where the stronger one is
        // available would be a downgrade dressed as a fix.
        Assert.Empty(FreshMerge.StoresScopedOutOfThreadLookup(
            Walked(AliceStore), scope, new[] { AliceStore, BobStore }, null));
    }

    [Fact]
    public void AProfileListWithNothingInIt_MakesNoClaimEitherWay()
    {
        // Absence of a list is not evidence of a store - the rule every store pass in this
        // server is built on. A null list is Outlook unreachable; both say nothing.
        Assert.Empty(FreshMerge.StoresScopedOutOfThreadLookup(
            Walked(AliceStore), AliceStore, Array.Empty<string>(), null));
        Assert.Empty(FreshMerge.StoresScopedOutOfThreadLookup(
            Walked(AliceStore), AliceStore, null, null));
    }

    [Fact]
    public void AWalkThatNeverRan_OrFoundNothing_OrCouldNotNameItsStore_ClaimsNothing()
    {
        ThreadLiveInfo notRun = new ThreadLiveInfo { Performed = false, Error = "NoAnchorItem" };
        ThreadLiveInfo empty = Walked(AliceStore);
        empty.MembersWalked = 0;

        foreach (ThreadLiveInfo live in new[] { notRun, empty, Walked(null), Walked(string.Empty) })
        {
            Assert.Empty(FreshMerge.StoresScopedOutOfThreadLookup(
                live, AliceStore, new[] { AliceStore, BobStore }, null));
        }
    }

    [Fact]
    public void AStoreTheIndexHoldsNothingFor_IsLeftToTheCodeWhoseRemedyWorks()
    {
        // unindexed_store already names it, and its remedy (walk that store live) works.
        // This code's remedy is to drop the scope, which cannot help where there is no index
        // tier to widen to - so naming it here would hand the caller a remedy that does
        // nothing, twice over.
        Assert.Empty(FreshMerge.StoresScopedOutOfThreadLookup(
            Walked(AliceStore), AliceStore, new[] { AliceStore, PstStore }, new[] { PstStore }));
    }

    [Fact]
    public void ADuplicatedProfileEntry_IsNamedOnce()
    {
        Assert.Equal(
            new[] { BobStore },
            FreshMerge.StoresScopedOutOfThreadLookup(
                Walked(AliceStore), AliceStore, new[] { AliceStore, BobStore, "BOB@EXAMPLE.COM" }, null));
    }

    // ========================================== the code, the verdict and the sentence

    [Fact]
    public void TheNamedStores_RaiseTheCode_AndMakeTheLookupPartial()
    {
        ThreadLiveInfo live = Walked(AliceStore);
        live.StoresNotQueried = new[] { BobStore };

        Assert.Contains(
            FreshMerge.ThreadGapUnqueriedStore,
            FreshMerge.DescribeThreadCoverageGaps(live, new[] { AliceStore })!);
        Assert.Equal(
            FreshMerge.FreshnessPartial,
            FreshMerge.ClassifyThreadFreshness(live, new[] { AliceStore }));
    }

    [Fact]
    public void TheSentence_NamesTheScope_TheWalkedStore_AndWithdrawsTheStrongerClaim()
    {
        string line = Assert.Single(MailService.DescribeThreadCoverage(
            Unqueried(), FreshMerge.FreshnessPartial, AliceStore, scopeWidened: false, top: 50)!);

        Assert.Contains(AliceStore, line, StringComparison.Ordinal);
        Assert.Contains(BobStore, line, StringComparison.Ordinal);
        Assert.Contains("TELL THE USER", line, StringComparison.Ordinal);

        // "We could not ask", never "we asked and there was nothing there".
        Assert.Contains("cannot be established", line, StringComparison.Ordinal);
    }

    [Fact]
    public void TheSentence_TellsACallerWhoCHOSETheScope_ToDropIt()
    {
        string line = Assert.Single(MailService.DescribeThreadCoverage(
            Unqueried(), FreshMerge.FreshnessPartial, AliceStore, scopeWidened: false, top: 50)!);

        Assert.Contains("without store", line, StringComparison.Ordinal);
        Assert.DoesNotContain("DERIVED", line, StringComparison.Ordinal);
    }

    [Fact]
    public void TheSentence_TellsACallerWhoNEVERChoseOne_HowToStopItBeingDerived()
    {
        // The remedy has to differ, and this is the case that decides whether the sentence is
        // worth anything: telling a caller who passed no store to drop one reads as advice
        // they had already followed, and leaves them with no way to clear the flag at all.
        string line = Assert.Single(MailService.DescribeThreadCoverage(
            Unqueried(), FreshMerge.FreshnessPartial, AliceStore, scopeWidened: false, top: 50,
            scopeStoreDerived: true)!);

        Assert.Contains("DERIVED", line, StringComparison.Ordinal);
        Assert.Contains("conversation_id", line, StringComparison.Ordinal);
        Assert.DoesNotContain("without store", line, StringComparison.Ordinal);
    }

    [Fact]
    public void TheCappedList_SaysHowManyMoreThereWere()
    {
        ThreadLiveInfo live = Walked(AliceStore);
        live.StoresNotQueried = Enumerable.Range(1, MailService.UnindexedStoreListCap)
            .Select(i => "mailbox" + i + "@example.com").ToList();
        live.StoresNotQueriedTruncated = true;
        live.StoresNotQueriedTotal = 30;
        live.CoverageGaps = new[] { FreshMerge.ThreadGapUnqueriedStore };

        string line = Assert.Single(MailService.DescribeThreadCoverage(
            live, FreshMerge.FreshnessPartial, AliceStore, scopeWidened: false, top: 50)!);

        Assert.Contains("and 18 more", line, StringComparison.Ordinal);
        Assert.Contains("live.storesNotQueriedTotal", line, StringComparison.Ordinal);
    }

    // ============================================================ through the whole tool

    [Fact]
    public void AScopedLookupOnATwoMailboxProfile_NamesTheAccountItNeverAsked()
    {
        // The shipped defect, end to end. Both accounts are indexed, so C4's pass finds
        // nothing to report and the index rows - narrowed to one store - cannot name the
        // other. The answer used to be freshness "live" with degraded absent.
        using MailService service = Service(BothIndexed());

        ThreadOutcome outcome = service.Thread(conversationId: "conv-1", id: AnchorId, store: AliceStore);

        Assert.Equal(new[] { BobStore }, outcome.Live!.StoresNotQueried);
        Assert.Contains(FreshMerge.ThreadGapUnqueriedStore, outcome.Live.CoverageGaps!);
        Assert.Equal(FreshMerge.FreshnessPartial, outcome.Freshness);
        Assert.True(outcome.Degraded);
        Assert.Equal(AliceStore, outcome.ScopeStore);
        Assert.Null(outcome.ScopeStoreDerived);
        Assert.Contains(outcome.Advice!, a => a.Contains(BobStore, StringComparison.Ordinal));
    }

    [Fact]
    public void TheSameProfileWithNoScope_RaisesNothing()
    {
        // Nothing was narrowed, so nothing is claimed here and unwalked_store is free to make
        // the stronger claim off evidence it can now actually see.
        using MailService service = Service(BothIndexed());

        ThreadOutcome outcome = service.Thread(conversationId: "conv-1", id: AnchorId, store: null);

        Assert.Null(outcome.Live!.StoresNotQueried);
        Assert.Null(outcome.ScopeStore);
        Assert.Equal(FreshMerge.FreshnessLive, outcome.Freshness);
    }

    [Fact]
    public void AScopedLookupOnASingleStoreProfile_RaisesNothing()
    {
        // The cry-wolf case that decides whether the flag is worth anything: there is no other
        // store, so the scope cost this lookup nothing.
        using MailService service = Service(
            BothIndexed(), new[] { new ComStoreDetail(AliceStore, "store-alice", 0, true) });

        ThreadOutcome outcome = service.Thread(conversationId: "conv-1", id: AnchorId, store: AliceStore);

        Assert.Null(outcome.Live!.StoresNotQueried);
        Assert.Null(outcome.Live.CoverageGaps);
        Assert.Equal(FreshMerge.FreshnessLive, outcome.Freshness);
        Assert.Equal(AliceStore, outcome.ScopeStore);
    }

    [Fact]
    public void AnUnindexedSecondStore_IsReportedOnceByTheOtherCode()
    {
        // Both codes are computed from Outlook's store list, so without the subtraction this
        // store would be named twice - once with a remedy that works and once with one that
        // does not.
        using MailService service = Service(
            OnlyAliceIndexed(),
            new[]
            {
                new ComStoreDetail(AliceStore, "store-alice", 0, true),
                new ComStoreDetail(PstStore, "store-pst", 3, null),
            });

        ThreadOutcome outcome = service.Thread(conversationId: "conv-1", id: AnchorId, store: AliceStore);

        Assert.Equal(new[] { PstStore }, outcome.Live!.StoresWithoutIndex);
        Assert.Null(outcome.Live.StoresNotQueried);
        Assert.DoesNotContain(FreshMerge.ThreadGapUnqueriedStore, outcome.Live.CoverageGaps!);
    }

    [Fact]
    public void ADerivedStore_NarrowsTheLookup_AndSaysSo()
    {
        // The shape an agent actually reaches for: search, then thread on a hit id. No store
        // was passed, a store was applied anyway, and until now nothing in the payload said
        // either thing had happened.
        using MailService service = Service(BothIndexed());

        SearchOutcome search = service.Search(
            new SearchRequest { Query = "test", Top = 25, SnippetChars = 0 });
        string hitId = Assert.Single(search.Hits).Id;

        ThreadOutcome outcome = service.Thread(conversationId: null, id: hitId, store: null);

        Assert.Equal(AliceStore, outcome.ScopeStore);
        Assert.True(outcome.ScopeStoreDerived);
        Assert.Equal(new[] { BobStore }, outcome.Live!.StoresNotQueried);
        Assert.Equal(FreshMerge.FreshnessPartial, outcome.Freshness);
        Assert.Contains(
            outcome.Advice!,
            a => a.Contains("conversation_id", StringComparison.Ordinal)
                && a.Contains(BobStore, StringComparison.Ordinal));
    }

    [Fact]
    public void AStoreTheCallerNAMED_IsNeverReportedAsDerived()
    {
        // The case the derived/chosen split actually turns on, and the one the first pass of
        // this file missed: a hit id that COULD have supplied a store, beside a store the
        // caller passed anyway. The derivation must not claim it, because the remedy printed
        // beside it differs - telling this caller to pass conversation_id would send them the
        // long way round an argument they can simply drop.
        using MailService service = Service(BothIndexed());

        SearchOutcome search = service.Search(
            new SearchRequest { Query = "test", Top = 25, SnippetChars = 0 });
        string hitId = Assert.Single(search.Hits).Id;

        ThreadOutcome outcome = service.Thread(conversationId: null, id: hitId, store: AliceStore);

        Assert.Equal(AliceStore, outcome.ScopeStore);
        Assert.Null(outcome.ScopeStoreDerived);
        Assert.Contains(outcome.Advice!, a => a.Contains("without store", StringComparison.Ordinal));
    }

    [Fact]
    public void AScopeThatWIDENED_NarrowedNothing_AndIsNotReportedAsAScope()
    {
        // C3's state, read through C5's fields. A store the index cannot address resolves to
        // NO scope and the lookup runs profile-wide, so nothing was suppressed: reporting a
        // scopeStore here would name a narrowing that did not happen, and raising
        // unqueried_store would claim the profile went unasked when it was asked in full.
        using MailService service = Service(
            BothIndexed(),
            new[]
            {
                new ComStoreDetail(AliceStore, "store-alice", 0, true),
                new ComStoreDetail(BobStore, "store-bob", 0, true),
                new ComStoreDetail(PstStore, "store-pst", 3, null),
            });

        ThreadOutcome outcome = service.Thread(conversationId: "conv-1", id: AnchorId, store: PstStore);

        Assert.True(outcome.ScopeWidened);
        Assert.Null(outcome.ScopeStore);
        Assert.Null(outcome.Live!.StoresNotQueried);
        Assert.DoesNotContain(FreshMerge.ThreadGapUnqueriedStore, outcome.Live.CoverageGaps!);
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

    /// <summary>A walk of Alice's store with Bob's named as never asked about.</summary>
    private static ThreadLiveInfo Unqueried()
    {
        ThreadLiveInfo live = Walked(AliceStore);
        live.StoresNotQueried = new[] { BobStore };
        live.CoverageGaps = new[] { FreshMerge.ThreadGapUnqueriedStore };
        return live;
    }

    private static MailService Service(
        StubIndexClient index,
        IReadOnlyList<ComStoreDetail>? profileStores = null)
    {
        return new MailService(
            new DirectGateway(ScopedSession.Create(profileStores ?? TwoMailboxes)), null, index);
    }

    private static StubIndexClient BothIndexed() => new(new[] { AlicePrefix, BobPrefix });

    private static StubIndexClient OnlyAliceIndexed() => new(new[] { AlicePrefix });

    /// <summary>
    /// A Windows Search stand-in that knows a chosen SET of store prefixes and nothing else,
    /// answering the probe statements by shape and returning NO conversation rows - which is
    /// the state this file is about: a scoped query cannot produce a row from a store it was
    /// scoped away from, so the evidence unwalked_store needs is gone by construction.
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
    /// A session that answers the store list, one freshness sweep (so a search can register a
    /// hit for the derived-store case) and one conversation walk, both anchored in Alice's
    /// store. A <see cref="DispatchProxy"/> so a member added to the contract needs no change
    /// here. Not sealed: DispatchProxy derives from its TProxy at runtime.
    /// </summary>
    private class ScopedSession : DispatchProxy
    {
        private IReadOnlyList<ComStoreDetail> _stores = Array.Empty<ComStoreDetail>();

        internal static IOutlookSession Create(IReadOnlyList<ComStoreDetail> stores)
        {
            object proxy = Create<IOutlookSession, ScopedSession>()
                ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
            ((ScopedSession)proxy)._stores = stores;
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

                case nameof(IOutlookSession.SweepFoldersNewerThan):
                    return new ComSweepResult(
                        new[] { Member(1) },
                        foldersSwept: 4,
                        foldersSkipped: 0,
                        sweptFolders: new[] { AliceStore + "/Inbox" },
                        perStore: new[]
                        {
                            new ComStoreSweepCounters(
                                AliceStore, foldersSwept: 4, foldersSkipped: 0, foldersFailed: 0, foldersAbsent: 0),
                        });

                case nameof(IOutlookSession.TryGetConversationItems):
                    return Enumerable.Range(1, 3).Select(Member).ToList();

                default:
                    throw new NotSupportedException(
                        "The stand-in session was asked for " + (targetMethod?.Name ?? "an unnamed member")
                        + ", which this test does not model.");
            }
        }

        private static ComMailBrief Member(int i)
        {
            return new ComMailBrief(
                entryId: "MEMBER" + i,
                storeDisplayName: AliceStore,
                storeId: "store-alice",
                folderName: "Inbox",
                folderKind: "inbox",
                subject: "a test thread member",
                senderName: "Bob",
                senderAddress: "bob@example.com",
                receivedTime: Frontier.AddMinutes(i),
                isRead: true,
                hasAttachments: false,
                sizeBytes: 2048,
                body: "test body");
        }
    }
}
