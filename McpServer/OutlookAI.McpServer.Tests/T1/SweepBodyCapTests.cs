using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text.Json;

using OutlookAI.ComHost.Protocol;
using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The freshness sweep's mail bodies are bounded at the COM layer, so an answer too big to
/// send cannot be built in the first place - and every cut is reported.
/// <para>
/// THE DEFECT. One <c>SweepFoldersNewerThan</c> answer is ONE frame with a hard 64 MB
/// ceiling, <c>MailService</c> asks for it with <c>includeBodies: true</c>, and
/// <c>OutlookComSession.SnapshotBrief</c> took <c>item.Body</c> WHOLE. The caps that bound
/// body size - <c>BodyCharsDefault</c> and <c>BodyCharsCap</c> - live in <c>MailService</c>,
/// on the far side of that pipe, so they bounded what an agent sees and not what crosses.
/// The reachable path is a fast LOCAL store missing from the index, whose window falls back
/// to seven days: 4 folders x 200 items per store, full bodies, gathered fast enough to fit
/// the sweep's time budget.
/// </para>
/// <para>
/// WHY IT IS NOT AN ORDINARY TRUNCATION. These bodies are shown to nobody - they exist so
/// <c>FreshMerge.MatchesTerms</c> can run over mail newer than the index frontier - so
/// cutting one is not like <c>read</c>'s windowing, which loses nothing because it pages. A
/// cut here can make a search MISS a real match. That is why the cut is measured per item,
/// why the payload reports the intersection that could actually have cost a hit, and why
/// nothing in this file lets a cut body pass as a whole one.
/// </para>
/// <para>
/// Driven through the real <see cref="MailService"/> against a stand-in Outlook session and a
/// stand-in index client, on the model of <c>SearchCoverageClaimTests</c>. No mailbox and no
/// Windows Search index are touched.
/// </para>
/// </summary>
public sealed class SweepBodyCapTests
{
    private const string Sid = "{S-1-5-21-1111111111-2222222222-3333333333-1001}";

    private const string Store = "alice@example.com";

    private const string OtherStore = "Archive 2019.pst";

    private const string StorePrefix = "mapi16://" + Sid + "/" + Store + "($deadbeef)";

    private static readonly DateTime Frontier = new(2026, 8, 18, 9, 30, 0, DateTimeKind.Utc);

    private static readonly ComStoreDetail[] ProfileStores =
    {
        new ComStoreDetail(Store, "store-alice", 0, true),
        new ComStoreDetail(OtherStore, "store-archive", null, false),
    };

    // ============================================================ the bound itself

    [Fact]
    public void ABodyInsideBothBounds_CrossesWhole_AndIsNotMarked()
    {
        string body = new string('a', 1000);

        string? kept = OutlookComSession.CapSweepBody(
            body, OutlookComSession.SweepBodyBytesBudget, out long spent, out bool truncated);

        Assert.Same(body, kept);
        Assert.False(truncated);
        Assert.Equal(1000, spent);
    }

    [Fact]
    public void ABodyPastThePerItemCeiling_IsCutAtIt_AndSaysSo()
    {
        // The ceiling, on its own, with the whole budget available - so nothing else can be
        // what cut it.
        string body = new string('a', OutlookComSession.SweepBodyCharsCap + 1);

        string? kept = OutlookComSession.CapSweepBody(
            body, OutlookComSession.SweepBodyBytesBudget, out long spent, out bool truncated);

        Assert.Equal(OutlookComSession.SweepBodyCharsCap, kept!.Length);
        Assert.True(truncated);
        Assert.Equal(OutlookComSession.SweepBodyCharsCap, spent);
        Assert.False(OutlookComSession.BodyCutByBudget(kept));
    }

    [Fact]
    public void ABodyTheRemainingBudgetCannotPayFor_IsCutThere_AndIsAttributedToTheBudget()
    {
        string body = new string('a', 5000);

        string? kept = OutlookComSession.CapSweepBody(body, 120, out long spent, out bool truncated);

        Assert.Equal(120, kept!.Length);
        Assert.True(truncated);
        Assert.Equal(120, spent);
        Assert.True(OutlookComSession.BodyCutByBudget(kept));
    }

    [Fact]
    public void AnExhaustedBudget_LeavesTheBodyEmptyRatherThanWhole()
    {
        // The failure this replaces: an item swept after the budget ran out must not arrive
        // carrying a body that reads as complete. Empty plus the flag is the honest shape.
        string? kept = OutlookComSession.CapSweepBody("anything", 0, out long spent, out bool truncated);

        Assert.Equal(string.Empty, kept);
        Assert.True(truncated);
        Assert.Equal(0, spent);
        Assert.True(OutlookComSession.BodyCutByBudget(kept));
    }

    [Fact]
    public void ANullOrEmptyBody_SpendsNothing_AndIsNotACut()
    {
        Assert.Null(OutlookComSession.CapSweepBody(null, 1000, out long spentNull, out bool cutNull));
        Assert.Equal(0, spentNull);
        Assert.False(cutNull);

        Assert.Equal(string.Empty, OutlookComSession.CapSweepBody(string.Empty, 1000, out long spentEmpty, out bool cutEmpty));
        Assert.Equal(0, spentEmpty);
        Assert.False(cutEmpty);
    }

    [Fact]
    public void NonLatinText_CostsMorePerCharacter_WhichIsTheWholeReasonTheBudgetIsInBytes()
    {
        // A character budget would have to be set for this case and would then bite far too
        // early on ordinary Latin mail. In bytes, the same budget carries ~6x more Latin text
        // than CJK - which is right, because those frames really are that much bigger.
        OutlookComSession.CapSweepBody(new string('a', 600), 600, out long latin, out bool latinCut);
        string? cjk = OutlookComSession.CapSweepBody(new string('漢', 600), 600, out long wide, out bool wideCut);

        Assert.False(latinCut);
        Assert.Equal(600, latin);
        Assert.True(wideCut);
        Assert.Equal(100, cjk!.Length);
        Assert.Equal(600, wide);
    }

    // ============================================ the frame guarantee this all rests on

    [Fact]
    public void TheByteCeiling_IsNeverBelowWhatTheRealSerializerEmits()
    {
        // The whole safety argument is that the estimate cannot UNDER-count: an estimate that
        // is sometimes low is an estimate that lets an unsendable frame be built. Checked
        // against the actual encoder the pipe uses rather than against a belief about it.
        foreach (string sample in new[]
                 {
                     "plain ascii text 12345",
                     "quotes \" and backslash \\ and angle <brackets> & ampersand",
                     "control\r\n\tcharacters",
                     "accented eeee: éèêë",
                     "漢字 русский",
                     "astral 😀 emoji",
                     "'+`~!@#$%^*()-_=[]{};:,./?|",
                 })
        {
            byte[] encoded = JsonSerializer.SerializeToUtf8Bytes(sample, ComHostProtocol.Json);

            // Two bytes of the encoding are the opening and closing quotes, which are framing
            // rather than content - the ceiling measures content only.
            Assert.True(
                OutlookComSession.EncodedBodyByteCeiling(sample) >= encoded.Length - 2,
                $"ceiling {OutlookComSession.EncodedBodyByteCeiling(sample)} under-counted '{sample}' "
                + $"which encodes to {encoded.Length - 2} content bytes");
        }
    }

    [Fact]
    public void TheSweepBodyBudget_LeavesHalfTheFrameForEverythingElse()
    {
        // The cross-assembly invariant that makes the cap mean anything: Core cannot see the
        // pipe's constant, so the relationship is pinned by the one compilation that sees
        // both. Bodies take at most half the frame; the other half carries per-item EntryIDs,
        // StoreIDs, subjects, senders and folder names (~1-2 KB each, ~16 MB at the 8000 items
        // a folder-scoped sweep can reach), and what is left over is margin.
        Assert.Equal(ComHostProtocol.MaxFrameBytes, OutlookComSession.SweepBodyBytesBudget * 2);
    }

    // ================================================== what the payload says when it bites

    [Fact]
    public void ACutBodyOnAnItemThatDidNotMatch_RaisesTheCode_AndDegradesTheAnswer()
    {
        // THE CASE THAT MATTERS. The item was swept, its body was cut, and it then failed the
        // term match - so a term living past the cut would have been missed and this may be a
        // hit the answer does not contain.
        using MailService service = Service(_ => SweepOf(Mail("AA1", body: "nothing relevant here", cut: true)));

        SearchOutcome outcome = service.Search(Request("needle"));

        Assert.Equal(1, outcome.Sweep!.ItemsBodyCapped);
        Assert.Equal(1, outcome.Sweep.ItemsBodyCappedUnmatched);
        Assert.Contains(FreshMerge.GapBodyCap, outcome.Sweep.CoverageGaps!);
        Assert.Equal(FreshMerge.FreshnessPartial, outcome.Freshness);
        Assert.True(outcome.Degraded);
        Assert.Contains(outcome.Advice!, a => a.Contains("MAY be hits", StringComparison.Ordinal));
    }

    [Fact]
    public void TheAdvice_QuotesTheCapAndBothCounts_RatherThanRestatingThem()
    {
        using MailService service = Service(_ => SweepOf(
            Mail("AA1", body: "nothing relevant here", cut: true),
            Mail("AA2", body: "needle is right here", cut: true)));

        SearchOutcome outcome = service.Search(Request("needle"));

        string line = Assert.Single(outcome.Advice!, a => a.Contains("MAY be hits", StringComparison.Ordinal));
        Assert.Contains(
            OutlookComSession.SweepBodyCharsCap.ToString(System.Globalization.CultureInfo.InvariantCulture),
            line,
            StringComparison.Ordinal);

        // Two cut, one of them unmatched - and the sentence says both, because "some bodies
        // were cut" and "one of them might have been a hit" are different sizes of problem.
        Assert.Equal(2, outcome.Sweep!.ItemsBodyCapped);
        Assert.Equal(1, outcome.Sweep.ItemsBodyCappedUnmatched);
        Assert.Contains("of 2 just-arrived item(s)", line, StringComparison.Ordinal);
        Assert.Contains("1 of them did not match", line, StringComparison.Ordinal);
    }

    [Fact]
    public void ACutBodyOnAnItemThatMatchedAnyway_IsReportedAsAFact_AndRaisesNothing()
    {
        // The cry-wolf guard. These bodies are never displayed, so a cut on an item that
        // matched cost the answer nothing at all - and a code that fired here would blunt the
        // one that fires when something really may be missing.
        using MailService service = Service(_ => SweepOf(Mail("AA1", body: "the needle is here", cut: true)));

        SearchOutcome outcome = service.Search(Request("needle"));

        Assert.Equal(1, outcome.Sweep!.ItemsBodyCapped);
        Assert.Null(outcome.Sweep.ItemsBodyCappedUnmatched);
        Assert.Null(outcome.Sweep.CoverageGaps);
        Assert.Equal(FreshMerge.FreshnessLive, outcome.Freshness);
        Assert.Null(outcome.Degraded);
        Assert.DoesNotContain(outcome.Advice ?? Array.Empty<string>(), a => a.Contains("MAY be hits", StringComparison.Ordinal));
    }

    [Fact]
    public void ASubjectOnlySearch_NeverRaisesIt_BecauseTheBodyIsNeverConsulted()
    {
        using MailService service = Service(_ => SweepOf(Mail("AA1", body: "nothing relevant here", cut: true)));

        SearchOutcome outcome = service.Search(Request("needle", searchIn: SearchIn.SubjectOnly));

        Assert.Equal(1, outcome.Sweep!.ItemsBodyCapped);
        Assert.Null(outcome.Sweep.ItemsBodyCappedUnmatched);
        Assert.Null(outcome.Sweep.CoverageGaps);
        Assert.Equal(FreshMerge.FreshnessLive, outcome.Freshness);
    }

    [Fact]
    public void ATermLessSearch_NeverRaisesIt_BecauseEveryItemMatches()
    {
        using MailService service = Service(_ => SweepOf(Mail("AA1", body: "anything", cut: true)));

        SearchOutcome outcome = service.Search(Request(query: null, from: "bob@example.com"));

        Assert.Equal(1, outcome.Sweep!.ItemsBodyCapped);
        Assert.Null(outcome.Sweep.ItemsBodyCappedUnmatched);
        Assert.Null(outcome.Sweep.CoverageGaps);
    }

    [Fact]
    public void AnUncutSweep_CarriesNeitherFieldAtAll()
    {
        using MailService service = Service(_ => SweepOf(Mail("AA1", body: "the needle is here")));

        SearchOutcome outcome = service.Search(Request("needle"));

        Assert.Null(outcome.Sweep!.ItemsBodyCapped);
        Assert.Null(outcome.Sweep.ItemsBodyCappedUnmatched);
        Assert.Null(outcome.Sweep.BodyBudgetExhausted);
    }

    // ========================================================= which bound, and whose store

    [Fact]
    public void AnExhaustedBudget_ChangesTheRemedyTheAdviceGives()
    {
        // A per-item cut points at ONE enormous mail and read pages the whole of it; an
        // exhausted budget points at the sweep's own breadth and is answered by narrowing it.
        using MailService budgetService = Service(
            _ => SweepOf(new[] { Mail("AA1", body: "nothing relevant", cut: true) }, budgetExhausted: true));
        using MailService itemService = Service(
            _ => SweepOf(new[] { Mail("AA1", body: "nothing relevant", cut: true) }, budgetExhausted: false));

        SearchOutcome budget = budgetService.Search(Request("needle"));
        SearchOutcome perItem = itemService.Search(Request("needle"));

        Assert.True(budget.Sweep!.BodyBudgetExhausted);
        Assert.Contains(budget.Advice!, a => a.Contains("narrow the search with store, folder or 'after'", StringComparison.Ordinal));

        Assert.Null(perItem.Sweep!.BodyBudgetExhausted);
        Assert.Contains(perItem.Advice!, a => a.Contains("body_offset", StringComparison.Ordinal));
        Assert.DoesNotContain(perItem.Advice!, a => a.Contains("narrow the search with store, folder or 'after'", StringComparison.Ordinal));
    }

    [Fact]
    public void AnExhaustedBudget_IsNotReportedToAScopeThatLostNothingByIt()
    {
        // The budget belongs to the FRAME, which spans every store the sweep visited, so a
        // store-scoped answer that lost nothing must not import another account's condition -
        // the cross-store leak the per-store counters exist to prevent.
        using MailService service = Service(
            _ => SweepOf(
                new[] { Mail("AA1", body: "the needle is here"), Mail("ZZ1", store: OtherStore, body: "x", cut: true) },
                budgetExhausted: true));

        SearchOutcome outcome = service.Search(Request("needle", store: Store));

        Assert.Null(outcome.Sweep!.ItemsBodyCapped);
        Assert.Null(outcome.Sweep.BodyBudgetExhausted);
        Assert.Null(outcome.Sweep.CoverageGaps);
    }

    [Fact]
    public void AnotherStoresCutBody_IsNeverCountedInThisStoresAnswer()
    {
        // Counted inside the store filter, so a cached all-stores sweep serving a
        // store-scoped request reports this store's cuts and not the sweep's total.
        using MailService service = Service(
            _ => SweepOf(
                Mail("AA1", body: "the needle is here"),
                Mail("ZZ1", store: OtherStore, body: "nothing relevant", cut: true)));

        SearchOutcome scoped = service.Search(Request("needle", store: Store));
        Assert.Null(scoped.Sweep!.ItemsBodyCapped);
        Assert.Null(scoped.Sweep.CoverageGaps);

        // Unscoped, the same sweep does report it - the loss is real, it just belongs to the
        // other store.
        using MailService unscoped = Service(
            _ => SweepOf(
                Mail("AA1", body: "the needle is here"),
                Mail("ZZ1", store: OtherStore, body: "nothing relevant", cut: true)));
        SearchOutcome all = unscoped.Search(Request("needle"));
        Assert.Equal(1, all.Sweep!.ItemsBodyCapped);
        Assert.Equal(1, all.Sweep.ItemsBodyCappedUnmatched);
        Assert.Contains(FreshMerge.GapBodyCap, all.Sweep.CoverageGaps!);
    }

    // ============================================================== the body cache

    [Fact]
    public void ACutSweepBody_IsNeverServedAsAWholeBody_ByReadOrItsCache()
    {
        // THE BUG THIS FORBIDS, which would be worse than the one being fixed: a truncated
        // body cached as if it were whole. It is structurally absent - the cache is fed only
        // from ComItemDetail, on read's own path, and a ComMailBrief never reaches it - and
        // this pins the behaviour rather than the structure, because that is what a future
        // refactor would break.
        const string Whole = "the needle is in the part that was cut";
        using MailService service = Service(
            _ => SweepOf(Mail("AA1", body: "the needle is in the p", cut: true)),
            entryId => Detail(entryId, Whole));

        SearchOutcome outcome = service.Search(Request("needle"));
        HitSummary hit = Assert.Single(outcome.Hits);

        // THE CONTINUATION READ FIRST, and the order is the whole test. An offset-0 read
        // always goes to Outlook and refreshes the cache, so it would paper over a truncated
        // entry the sweep had left there; only body_offset > 0 is SERVED from the cache. A
        // sweep that cached its own cut body would answer this from 21 characters and hand
        // back an empty window over a body that is 38 characters long.
        ReadOutcome tail = service.Read(hit.Id!, maxBodyChars: 10, bodyOffset: 25);
        Assert.Equal(Whole.Substring(25, 10), tail.Body);
        Assert.Equal(Whole.Length, tail.BodyTotalChars);

        ReadOutcome read = service.Read(hit.Id!);

        Assert.Equal(Whole, read.Body);
        Assert.Equal(Whole.Length, read.BodyTotalChars);
        Assert.False(read.BodyTruncated);
    }

    // ===================================================================== fixtures

    private static SearchRequest Request(
        string? query = "needle",
        string? store = null,
        string? from = null,
        SearchIn searchIn = SearchInValues.Default)
    {
        return new SearchRequest
        {
            Query = query,
            SearchIn = searchIn,
            Store = store,
            From = from,
            Top = 25,
            SnippetChars = 0,
        };
    }

    private static ComSweepResult SweepOf(params ComMailBrief[] items)
    {
        return SweepOf(items, budgetExhausted: false);
    }

    private static ComSweepResult SweepOf(ComMailBrief[] items, bool budgetExhausted)
    {
        int cut = 0;
        foreach (ComMailBrief item in items)
        {
            if (item.BodyTruncated == true)
            {
                cut++;
            }
        }

        return new ComSweepResult(
            items,
            foldersSwept: 8,
            foldersSkipped: 0,
            sweptFolders: new[] { Store + "/Inbox", OtherStore + "/Inbox" },
            perStore: new[]
            {
                new ComStoreSweepCounters(Store, 4, 0, 0, 0),
                new ComStoreSweepCounters(OtherStore, 4, 0, 0, 0),
            },
            bodiesTruncated: cut,
            bodyBudgetExhausted: budgetExhausted);
    }

    private static ComMailBrief Mail(string entryId, string? body = null, bool cut = false, string store = Store)
    {
        return new ComMailBrief(
            entryId: entryId,
            storeDisplayName: store,
            storeId: store == Store ? "store-alice" : "store-archive",
            folderName: "Inbox",
            folderKind: "inbox",
            subject: "a test mail",
            senderName: "Bob",
            senderAddress: "bob@example.com",
            receivedTime: Frontier,
            isRead: true,
            hasAttachments: false,
            sizeBytes: 2048,
            body: body,
            messageClass: null,
            bodyTruncated: cut ? true : (bool?)null);
    }

    private static ComItemDetail Detail(string entryId, string body)
    {
        return new ComItemDetail(
            entryId: entryId,
            storeDisplayName: Store,
            folderPath: "\\\\" + Store + "\\Inbox",
            itemClass: 43,
            subject: "a test mail",
            senderName: "Bob",
            senderAddress: "bob@example.com",
            receivedTime: Frontier,
            sentTime: Frontier,
            recipients: Array.Empty<ComRecipientInfo>(),
            body: body,
            bodyTotalChars: body.Length,
            bodyOrigin: "text",
            attachments: Array.Empty<ComAttachmentInfo>(),
            sizeBytes: 2048,
            isRead: true,
            conversationId: null,
            internetMessageId: null,
            headers: null);
    }

    private static MailService Service(
        Func<string?, ComSweepResult> sweep,
        Func<string, ComItemDetail>? read = null)
    {
        return new MailService(
            new DirectGateway(StandInSession.Create(ProfileStores, sweep, read)), null, new StubIndexClient());
    }

    /// <summary>
    /// A Windows Search stand-in that knows one store and answers the probe statements by
    /// shape, returning no SEARCH rows at all - so the index tier contributes nothing and the
    /// question is what the freshness tier then reports about itself.
    /// </summary>
    private sealed class StubIndexClient : IIndexClient
    {
        private const string DiscoveryTail = " System.ItemUrl FROM SystemIndex WHERE System.Kind='email'";

        public IndexProviderKind Provider => IndexProviderKind.OleDb;

        public IReadOnlyList<IReadOnlyDictionary<string, object?>> ExecuteRows(
            string sql, int maxRows, int? commandTimeoutSeconds = null)
        {
            if (sql.EndsWith(DiscoveryTail, StringComparison.Ordinal))
            {
                return Rows(StorePrefix + "/0/Inbox/sampled-item");
            }

            if (sql.Contains("System.Message.DateReceived FROM SystemIndex", StringComparison.Ordinal))
            {
                return new[]
                {
                    (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                    {
                        ["System.Message.DateReceived"] = Frontier,
                    },
                };
            }

            if (sql.StartsWith("SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE", StringComparison.Ordinal))
            {
                return Rows(StorePrefix + "/0/Inbox/probed-item");
            }

            return Array.Empty<IReadOnlyDictionary<string, object?>>();
        }

        private static IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows(string itemUrl)
        {
            return new[]
            {
                (IReadOnlyDictionary<string, object?>)new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                {
                    ["System.ItemUrl"] = itemUrl,
                },
            };
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
    /// A session that answers the profile's store list, the freshness sweep and - for the body
    /// cache case - one item read, and refuses everything else. A <see cref="DispatchProxy"/>
    /// rather than a stub per member, so a method added to the contract needs no change here.
    /// Not sealed: DispatchProxy derives from its TProxy at runtime and refuses a sealed one.
    /// </summary>
    private class StandInSession : DispatchProxy
    {
        private IReadOnlyList<ComStoreDetail> _stores = Array.Empty<ComStoreDetail>();
        private Func<string?, ComSweepResult> _sweep = _ => throw new NotSupportedException();
        private Func<string, ComItemDetail>? _read;

        internal static IOutlookSession Create(
            IReadOnlyList<ComStoreDetail> stores,
            Func<string?, ComSweepResult> sweep,
            Func<string, ComItemDetail>? read)
        {
            object proxy = Create<IOutlookSession, StandInSession>()
                ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
            ((StandInSession)proxy)._stores = stores;
            ((StandInSession)proxy)._sweep = sweep;
            ((StandInSession)proxy)._read = read;
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

                // Argument 0 is entryIdHex and argument 4 is the out error; see
                // IOutlookSession.TryReadItem. The read always succeeds here - what this
                // fixture exists to show is that read goes to Outlook at all rather than
                // being served the sweep's truncated copy.
                case nameof(IOutlookSession.TryReadItem):
                    if (_read == null)
                    {
                        throw new NotSupportedException("This fixture has no item read.");
                    }

                    if (args != null && args.Length > 4)
                    {
                        args[4] = null;
                    }

                    return _read((string)args![0]!);

                default:
                    throw new NotSupportedException(
                        "The stand-in session does not implement " + (targetMethod?.Name ?? "?") + ".");
            }
        }
    }
}
