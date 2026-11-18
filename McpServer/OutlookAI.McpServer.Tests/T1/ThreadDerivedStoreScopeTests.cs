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
/// C5's behavioural half: a store <c>thread</c> DERIVED from the referenced hit no longer
/// narrows the lookup. A store the caller NAMED still does.
/// <para>
/// THE DEFECT. Passing <c>id</c> alone is the shape an agent reaches for straight out of a
/// search, and <c>thread</c> has to open that hit anyway to recover its conversation id - so
/// it took the hit's STORE while it was there and scoped the index query to it. A
/// conversation reaching into a second indexed account then came back missing those members,
/// on a tool whose description promises the whole conversation. C5 made that visible
/// (<c>scopeStore</c>, <c>scopeStoreDerived</c>, <c>live.storesNotQueried</c> and the
/// <c>unqueried_store</c> code); naming a cost is not the same as not paying it, and nobody
/// had asked for the narrowing in the first place.
/// </para>
/// <para>
/// WHAT THE FIX BUYS BEYOND THE MEMBERS. The conversation's index rows are the evidence
/// <c>unwalked_store</c> is computed from, and a scope narrows that evidence - the parameter
/// silenced the code that would have reported what the parameter cost. Unscoped, the rows can
/// finally name the other store, so the STRONGER claim ("this conversation demonstrably has
/// members the walk did not cover") replaces the weaker one ("the question could not be
/// asked"), and <c>unqueried_store</c> goes quiet unless a caller really chose a scope.
/// </para>
/// <para>
/// THE COST, pinned as deliberate rather than accidental: one UNSCOPED ConversationID query
/// per <c>thread</c> call. That is the accepted trade.
/// </para>
/// <para>
/// Driven through the real <see cref="MailService"/> against a stand-in session and a
/// stand-in index client that records every statement it is given. No mailbox and no Windows
/// Search index are touched. What T1 cannot reach, and needs a live profile: that Outlook's
/// Conversation object really does stop at the anchor's store.
/// </para>
/// </summary>
[Trait("TimeZone", "Machine")]
public sealed class ThreadDerivedStoreScopeTests
{
    private const string Sid = "{S-1-5-21-1111111111-2222222222-3333333333-1001}";

    private const string AliceStore = "alice@example.com";

    private const string AlicePrefix = "mapi16://" + Sid + "/" + AliceStore + "($deadbeef)";

    private const string BobStore = "bob@example.com";

    private const string BobPrefix = "mapi16://" + Sid + "/" + BobStore + "($c0ffee00)";

    private const string ConversationId = "conv-1";

    private static readonly DateTime AliceFrontier = new(2026, 8, 18, 9, 30, 0, DateTimeKind.Utc);

    /// <summary>Bob's index is AHEAD, so a profile-wide staleness probe is distinguishable from Alice's.</summary>
    private static readonly DateTime BobFrontier = AliceFrontier.AddHours(4);

    private static readonly ComStoreDetail[] TwoMailboxes =
    {
        new ComStoreDetail(AliceStore, "store-alice", 0, true),
        new ComStoreDetail(BobStore, "store-bob", 0, true),
    };

    // ================================================= the derived store no longer scopes

    [Fact]
    public void AThreadOnAHitId_AsksTheINDEXAboutTheWholeProfile()
    {
        StandIn world = new StandIn();
        using MailService service = world.Service();

        ThreadOutcome outcome = service.Thread(conversationId: null, id: world.SearchForAHitId(service), storeId: null);

        // The statement itself, which is the only place the narrowing was ever real.
        Assert.DoesNotContain("SCOPE=", world.ConversationStatements.Single(), StringComparison.Ordinal);

        // And the payload says a scope was not applied, rather than saying which one was -
        // while scopeStoreDerived reports that a store WAS available to narrow to and was
        // passed over, which is the field's meaning since it stopped being able to carry the
        // old one.
        Assert.Null(outcome.ScopeStore);
        Assert.True(outcome.ScopeStoreDerived);
        Assert.Null(outcome.ScopeWidened);
    }

    [Fact]
    public void TheDerivedStore_IsReportedAsPassedOver_WithNoRemedyBesideIt()
    {
        // scopeStoreDerived's meaning since the narrowing went away, and the reason it carries
        // no advice: a store was there to narrow to and was deliberately not used, so no member
        // is missing on its account and there is nothing for the caller to do about it. The
        // advice list exists to name partial coverage; a sentence here would be a remedy for a
        // decision that cost nothing, which is the kind that never clears.
        StandIn world = new StandIn();
        using MailService service = world.Service();

        ThreadOutcome outcome = service.Thread(conversationId: null, id: world.SearchForAHitId(service), storeId: null);

        Assert.True(outcome.ScopeStoreDerived);
        Assert.Null(outcome.ScopeStore);
        Assert.DoesNotContain(FreshMerge.ThreadGapUnqueriedStore, outcome.Live!.CoverageGaps!);
        Assert.DoesNotContain(
            outcome.Advice ?? Array.Empty<string>(),
            a => a.Contains("DERIVED", StringComparison.Ordinal)
                || a.Contains("narrowed to one store", StringComparison.Ordinal));
    }

    [Fact]
    public void TheSecondAccountsMember_IsInTheAnswer()
    {
        // The whole point. Before this the member existed, was indexed, and was absent.
        StandIn world = new StandIn();
        using MailService service = world.Service();

        ThreadOutcome outcome = service.Thread(conversationId: null, id: world.SearchForAHitId(service), storeId: null);

        Assert.Contains(outcome.Hits, h => string.Equals(h.Store, BobStore, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void TheUnwalkedStoreCode_CanNowMakeItsStrongerClaim()
    {
        // unwalked_store is computed from the stores the conversation's index rows name, so a
        // scope used to make it structurally unable to fire. Unscoped, it fires - and it says
        // the members are demonstrably elsewhere, not merely that nobody asked.
        StandIn world = new StandIn();
        using MailService service = world.Service();

        ThreadOutcome outcome = service.Thread(conversationId: null, id: world.SearchForAHitId(service), storeId: null);

        Assert.Contains(FreshMerge.ThreadGapUnwalkedStore, outcome.Live!.CoverageGaps!);
        Assert.DoesNotContain(FreshMerge.ThreadGapUnqueriedStore, outcome.Live.CoverageGaps!);
        Assert.Null(outcome.Live.StoresNotQueried);
        Assert.Equal(FreshMerge.FreshnessPartial, outcome.Freshness);
        Assert.Contains(
            outcome.Advice!,
            a => a.Contains("The live check covered", StringComparison.Ordinal)
                && a.Contains("id from that account", StringComparison.Ordinal));
    }

    [Fact]
    public void TheStalenessBlock_IsTheWHOLEProfilesToo()
    {
        // Staleness describes the scope the lookup ran under, so it has to widen with it -
        // reporting one account's frontier over a profile-wide answer would be the same
        // mismatch in the freshness half.
        StandIn world = new StandIn();
        using MailService service = world.Service();

        ThreadOutcome outcome = service.Thread(conversationId: null, id: world.SearchForAHitId(service), storeId: null);

        Assert.Equal(BobFrontier, outcome.Staleness!.NewestIndexedUtc);
    }

    // ==================================================== a store the CALLER named still does

    [Fact]
    public void AStoreTheCallerNamed_StillScopesTheLookup_AndStillSaysWhatItCost()
    {
        // The hint stays a hint the caller can drop, and the reporting C5 added is exactly
        // what still covers it.
        StandIn world = new StandIn();
        using MailService service = world.Service();

        ThreadOutcome outcome = service.Thread(
            conversationId: null, id: world.SearchForAHitId(service), storeId: TestFolders.StoreId(AliceStore));

        Assert.Contains("SCOPE='" + AlicePrefix + "'", world.ConversationStatements.Single(), StringComparison.Ordinal);
        Assert.Equal(AliceStore, outcome.ScopeStore);
        Assert.Null(outcome.ScopeStoreDerived);
        Assert.Equal(new[] { BobStore }, outcome.Live!.StoresNotQueried);
        Assert.Contains(FreshMerge.ThreadGapUnqueriedStore, outcome.Live.CoverageGaps!);
        Assert.Contains(
            outcome.Advice!,
            a => a.Contains("narrowed to one store", StringComparison.Ordinal)
                && a.Contains("without store", StringComparison.Ordinal));
    }

    [Fact]
    public void AScopedLookup_StillReportsThatOneAccountsFreshness()
    {
        StandIn world = new StandIn();
        using MailService service = world.Service();

        ThreadOutcome outcome = service.Thread(
            conversationId: null, id: world.SearchForAHitId(service), storeId: TestFolders.StoreId(AliceStore));

        Assert.Equal(AliceFrontier, outcome.Staleness!.NewestIndexedUtc);
    }

    // ============================================================== the unchanged shapes

    [Fact]
    public void AConversationIdWithNoStore_WasAlreadyUnscoped_AndStaysThatWay()
    {
        // Nothing is derived when the caller supplies the conversation id, so this shape never
        // had the defect - re-pinned because the branch that used to scope it is the branch
        // that moved.
        StandIn world = new StandIn();
        using MailService service = world.Service();

        ThreadOutcome outcome = service.Thread(conversationId: ConversationId, id: StandIn.AnchorEntryId, storeId: null);

        Assert.DoesNotContain("SCOPE=", world.ConversationStatements.Single(), StringComparison.Ordinal);
        Assert.Null(outcome.ScopeStore);

        // And nothing was DERIVED either: the conversation id came from the caller, so the hit
        // was never consulted for a store. The unscoped lookup here is the caller's own doing.
        Assert.Null(outcome.ScopeStoreDerived);
        Assert.Contains(outcome.Hits, h => string.Equals(h.Store, BobStore, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void AStoreTheIndexCannotAddress_StillWidens_AndStillSaysSo()
    {
        // C3's state is a CALLER-chosen store that resolved to no scope. The derived-store
        // change must not swallow it: the caller asked for something that was not applied.
        StandIn world = new StandIn();
        using MailService service = world.Service();

        ThreadOutcome outcome = service.Thread(
            conversationId: ConversationId, id: StandIn.AnchorEntryId, storeId: TestFolders.StoreId("Archive Old.pst"));

        Assert.True(outcome.ScopeWidened);
        Assert.Null(outcome.ScopeStore);
        Assert.Contains(
            outcome.Advice!,
            a => a.Contains("did not resolve to an index scope", StringComparison.Ordinal));
    }

    // ====================================================================== the fixture

    /// <summary>
    /// One profile, two indexed mailboxes, one conversation with a member in each - plus the
    /// statements the index was actually given, which is where the narrowing was ever real.
    /// </summary>
    private sealed class StandIn
    {
        /// <summary>A raw EntryID hex: long enough to be accepted without any hit cache or COM locate.</summary>
        internal const string AnchorEntryId = "00000000AABBCCDDEEFF00112233445566778899AABBCCDD";

        private readonly ConversationIndexClient _index;

        internal StandIn()
        {
            _index = new ConversationIndexClient();
        }

        /// <summary>Every ConversationID statement the index was given, in order.</summary>
        internal IReadOnlyList<string> ConversationStatements => _index.ConversationStatements;

        internal MailService Service()
        {
            return new MailService(new DirectGateway(Session.Create()), null, _index);
        }

        /// <summary>
        /// The shape an agent reaches for: search first, then thread on what came back. The
        /// hit is an INDEX hit in Alice's store, which is what makes a store derivable at all.
        /// </summary>
        internal string SearchForAHitId(MailService service)
        {
            SearchOutcome search = service.Search(
                new SearchRequest { Query = "test", Top = 25, SnippetChars = 0 });
            return search.Hits.First(h => string.Equals(h.Store, AliceStore, StringComparison.OrdinalIgnoreCase)).Id;
        }

        private static IReadOnlyDictionary<string, object?> Row(string prefix, string leaf, DateTime received)
        {
            return new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
            {
                ["System.ItemUrl"] = prefix + "/0/Inbox/" + leaf,
                ["System.Kind"] = new[] { "email" },
                ["System.Message.DateReceived"] = received,
                ["System.Subject"] = "a test thread member",
                ["System.Size"] = 2048L,
                ["System.Message.ConversationID"] = ConversationId,
            };
        }

        private static ComMailBrief Member(int i)
        {
            return new ComMailBrief(
                entryId: "MEMBER" + i.ToString(CultureInfo.InvariantCulture),
                storeDisplayName: AliceStore,
                storeId: "store-alice",
                folderName: "Inbox",
                folderKind: "inbox",
                subject: "a test thread member",
                senderName: "Bob",
                senderAddress: "bob@example.com",
                receivedTime: AliceFrontier.AddMinutes(i),
                isRead: true,
                hasAttachments: false,
                sizeBytes: 2048,
                body: "test body");
        }

        /// <summary>
        /// A Windows Search stand-in that knows both mailboxes and - the point of the fixture -
        /// returns the conversation's rows FILTERED BY THE STATEMENT'S OWN SCOPE, exactly as the
        /// real index would. A scope that is applied therefore costs Bob's member here too. The
        /// reachability probe is the shared stand-in's: a store's root or its own /0/ subtree, and
        /// nothing else - a looser rule answers YES to the DELEGATE probe '<owner>/1/<name>' as
        /// well, so every unknown store name would resolve to a delegate scope that exists only in
        /// the fixture.
        /// </summary>
        private sealed class ConversationIndexClient : StubIndexClient
        {
            private readonly List<string> _conversationStatements = new List<string>();

            internal ConversationIndexClient()
                : base(new[] { AlicePrefix, BobPrefix }, AliceFrontier)
            {
            }

            internal IReadOnlyList<string> ConversationStatements => _conversationStatements;

            public override IReadOnlyList<IReadOnlyDictionary<string, object?>> ExecuteRows(
                string sql, int maxRows, int? commandTimeoutSeconds = null)
            {
                if (!IsDiscovery(sql) && IsFreshness(sql))
                {
                    string? scope = IndexRows.ScopeOf(sql);
                    DateTime frontier = scope == null
                        ? BobFrontier // Profile-wide: the newest instant ANY store ingested.
                        : scope.StartsWith(BobPrefix, StringComparison.OrdinalIgnoreCase) ? BobFrontier : AliceFrontier;
                    return IndexRows.One("System.Message.DateReceived", frontier);
                }

                if (IsDiscovery(sql) || IsProbe(sql))
                {
                    return base.ExecuteRows(sql, maxRows, commandTimeoutSeconds);
                }

                if (sql.Contains("System.Message.ConversationID=", StringComparison.Ordinal))
                {
                    _conversationStatements.Add(sql);
                    string? scope = IndexRows.ScopeOf(sql);
                    List<IReadOnlyDictionary<string, object?>> rows = new List<IReadOnlyDictionary<string, object?>>();
                    if (scope == null || scope.StartsWith(AlicePrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        rows.Add(Row(AlicePrefix, "member-alice", AliceFrontier));
                    }

                    if (scope == null || scope.StartsWith(BobPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        rows.Add(Row(BobPrefix, "member-bob", BobFrontier));
                    }

                    return rows;
                }

                // Everything else is the SEARCH statement: one indexed mail in Alice's store,
                // carrying the conversation id, which is what makes a store derivable.
                return new[] { Row(AlicePrefix, "member-alice", AliceFrontier) };
            }
        }

        /// <summary>
        /// A session that answers the store list, one freshness sweep, the hit locate an index
        /// hit needs before it can anchor a walk, and one conversation walk - all of it inside
        /// Alice's store, which is Outlook's own constraint and the reason Bob's member can
        /// only ever come from the index. Scripted on the shared <see cref="StandInSession"/>.
        /// </summary>
        private static class Session
        {
            internal static IOutlookSession Create() =>
                StandInSession.Create(
                    TwoMailboxes,
                    _ => new ComSweepResult(
                        Array.Empty<ComMailBrief>(),
                        foldersSwept: 8,
                        foldersSkipped: 0,
                        sweptFolders: new[] { AliceStore + "/Inbox", BobStore + "/Inbox" },
                        perStore: new[]
                        {
                            new ComStoreSweepCounters(AliceStore, 4, 0, 0, 0),
                            new ComStoreSweepCounters(BobStore, 4, 0, 0, 0),
                        }),
                    standIn => standIn
                        // An index hit carries no EntryID, so the walk's anchor is located by
                        // folder path first (HitLocator tier 1).
                        .Answering(
                            nameof(IOutlookSession.TryResolveByPath),
                            _ => new ComOpenResult("MEMBER1", "a test thread member", AliceFrontier, 43))
                        .Answering(
                            nameof(IOutlookSession.TryGetConversationItems),
                            _ => Enumerable.Range(1, 3).Select(Member).ToList()));
        }
    }
}
