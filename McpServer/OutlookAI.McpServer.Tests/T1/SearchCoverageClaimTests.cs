using System;
using System.Collections.Generic;
using System.Reflection;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Four claims <c>search</c> makes about its own coverage that were false, or that went
/// missing exactly when they mattered (completeness-gap rows H2, G5, B2, F3). Each one is a
/// sentence an agent relays to a human as fact, which is the species the A2 fix established:
/// a health report saying "Index is current" over an index holding zero mail is worse than
/// silence, because nothing downstream can tell.
/// <list type="bullet">
/// <item><description><b>H2</b> - the sweep's <c>Table.Sort</c> failure was swallowed
/// ("unsorted sweep still works; the cap just cuts arbitrarily") while the <c>item_cap</c>
/// advice kept asserting the sweep "reads newest-first, so the OLDEST is not covered". When
/// the sort failed that sentence is simply wrong: the cut was arbitrary, so the hole is
/// arbitrary, and an agent told the oldest is missing reasons about the wrong mail rather
/// than merely about less of it.</description></item>
/// <item><description><b>G5</b> - the unresolved-folder guard returned early on any merged
/// hit, so one item the freshness sweep returned silenced it entirely. The swept item comes
/// from COM, where the folder resolved fine; it says nothing about whether the INDEX can
/// address the folder, which is the guard's whole question.</description></item>
/// <item><description><b>B2</b> - attachment text is index-only. The attachment-ONLY refusal
/// was reported; the DEFAULT case said nothing, so a term inside an attachment of mail that
/// arrived after the index frontier was invisible under
/// <c>freshness: "live"</c>.</description></item>
/// <item><description><b>F3</b> - <c>from</c> / <c>unread_only</c> / <c>has_attachments</c>
/// are applied to what the exhaustive scan's result cap already kept, so a scan can return 2
/// rows with <c>truncated: true</c> while thousands match, and nothing said the filter ran
/// after the cap.</description></item>
/// </list>
/// <para>
/// Driven through the real <see cref="MailService"/> against a stand-in Outlook session and a
/// stand-in index client, on the model of <c>UnindexedStoreReportingTests</c>. No mailbox and
/// no Windows Search index are touched.
/// </para>
/// </summary>
public sealed class SearchCoverageClaimTests
{
    private const string Sid = "{S-1-5-21-1111111111-2222222222-3333333333-1001}";

    private const string Store = "alice@example.com";

    private const string StorePrefix = "mapi16://" + Sid + "/" + Store + "($deadbeef)";

    /// <summary>A folder Outlook has and the index cannot address (renamed, or localized).</summary>
    private const string UnindexedFolder = "Archief";

    private static readonly DateTime Frontier = new(2026, 8, 18, 9, 30, 0, DateTimeKind.Utc);

    private static readonly ComStoreDetail[] ProfileStores =
    {
        new ComStoreDetail(Store, "store-alice", 0, true),
    };

    // ================================================================== H2

    [Fact]
    public void ACapThatCutAnUnsortedFolder_IsItsOwnCode_NotTheNewestFirstOne()
    {
        // The fields first: one cap, two meanings, so one code cannot carry both.
        SweepInfo sweep = CappedSweep(unsorted: Store + "/Inbox");

        IReadOnlyList<string> gaps = FreshMerge.DescribeCoverageGaps(sweep)!;

        Assert.Contains(FreshMerge.GapItemCapUnsorted, gaps);
        Assert.DoesNotContain(FreshMerge.GapItemCap, gaps);
        Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyFreshness(sweep));
    }

    [Fact]
    public void TheAdviceForAnUnsortedCap_NeverClaimsTheOldestIsWhatIsMissing()
    {
        // THE FALSE SENTENCE. This is the whole of H2: the cap fired over a table Outlook
        // would not sort, so "it reads newest-first, so the OLDEST ... is not covered" is not
        // an approximation, it points at mail that may be perfectly well covered while the
        // mail that is actually gone could be anything.
        IReadOnlyList<string> advice = MailService.DescribeSweepCoverage(
            Classified(CappedSweep(unsorted: Store + "/Inbox")), "12 minutes", folderScoped: false);

        Assert.DoesNotContain(advice, a => a.Contains("newest-first", StringComparison.Ordinal));
        Assert.DoesNotContain(advice, a => a.Contains("OLDEST", StringComparison.Ordinal));
        Assert.Contains(advice, a => a.Contains("ARBITRARY", StringComparison.Ordinal)
            && a.Contains(Store + "/Inbox", StringComparison.Ordinal));
    }

    [Fact]
    public void AMixedSweep_TellsEachSentenceOnlyAboutTheFoldersItIsTrueOf()
    {
        // One store, two folders, one of which sorted. Both sentences are emitted and
        // neither may name the other's folder - which is what makes the split worth having
        // rather than degrading every capped sweep to "we do not know".
        SweepInfo sweep = new SweepInfo
        {
            Performed = true,
            FoldersSwept = 4,
            ItemCappedFolders = new[] { Store + "/Inbox", Store + "/Sent Items" },
            ItemCappedFoldersUnsorted = new[] { Store + "/Sent Items" },
        };

        IReadOnlyList<string> gaps = FreshMerge.DescribeCoverageGaps(sweep)!;
        Assert.Contains(FreshMerge.GapItemCap, gaps);
        Assert.Contains(FreshMerge.GapItemCapUnsorted, gaps);

        IReadOnlyList<string> advice = MailService.DescribeSweepCoverage(
            Classified(sweep), "12 minutes", folderScoped: false);
        string newestFirst = Assert.Single(advice, a => a.Contains("newest-first", StringComparison.Ordinal));
        string arbitrary = Assert.Single(advice, a => a.Contains("ARBITRARY", StringComparison.Ordinal));

        Assert.Contains(Store + "/Inbox", newestFirst, StringComparison.Ordinal);
        Assert.DoesNotContain(Store + "/Sent Items", newestFirst, StringComparison.Ordinal);
        Assert.Contains(Store + "/Sent Items", arbitrary, StringComparison.Ordinal);
        Assert.DoesNotContain(Store + "/Inbox", arbitrary, StringComparison.Ordinal);
    }

    [Fact]
    public void AnOrdinaryCappedSweep_KeepsTheNewestFirstSentence()
    {
        // The regression guard on the claim that IS true: nothing about H2 may quietly
        // retire the one sentence that tells a caller which mail to go looking for.
        SweepInfo sweep = CappedSweep(unsorted: null);

        Assert.Contains(FreshMerge.GapItemCap, FreshMerge.DescribeCoverageGaps(sweep)!);
        Assert.DoesNotContain(FreshMerge.GapItemCapUnsorted, FreshMerge.DescribeCoverageGaps(sweep)!);
        Assert.Contains(
            MailService.DescribeSweepCoverage(Classified(sweep), "12 minutes", folderScoped: false),
            a => a.Contains("newest-first", StringComparison.Ordinal));
    }

    [Fact]
    public void AnUnsortedCap_ReachesTheSearchPayload_AndDegradesIt()
    {
        // End to end, because the COM layer is where the fact is measured and the payload is
        // where an agent reads it - the two used to be joined by an assumption.
        using MailService service = Service(Index(), _ => CappedComSweep(unsorted: true));

        SearchOutcome outcome = service.Search(Request());

        Assert.Equal(new[] { Store + "/Inbox" }, outcome.Sweep!.ItemCappedFoldersUnsorted);
        Assert.Contains(FreshMerge.GapItemCapUnsorted, outcome.Sweep.CoverageGaps!);
        Assert.Equal(FreshMerge.FreshnessPartial, outcome.Freshness);
        Assert.True(outcome.Degraded);
        Assert.Contains(outcome.Advice!, a => a.Contains("ARBITRARY", StringComparison.Ordinal));
        Assert.DoesNotContain(outcome.Advice!, a => a.Contains("newest-first", StringComparison.Ordinal));
    }

    [Fact]
    public void ASortedCap_ReachesTheSamePayloadWithTheOtherSentence()
    {
        using MailService service = Service(Index(), _ => CappedComSweep(unsorted: false));

        SearchOutcome outcome = service.Search(Request());

        Assert.Null(outcome.Sweep!.ItemCappedFoldersUnsorted);
        Assert.Contains(FreshMerge.GapItemCap, outcome.Sweep.CoverageGaps!);
        Assert.Contains(outcome.Advice!, a => a.Contains("newest-first", StringComparison.Ordinal));
    }

    [Fact]
    public void TheUnsortedSubset_IsFilteredByTheSameStoreRuleAsTheSetItBelongsTo()
    {
        // If the subset were filtered differently it would stop being a subset, and the
        // difference the two sentences are built from would name a folder neither list has.
        SweepInfo info = new SweepInfo();
        MailService.ApplySweepCounters(
            info,
            new ComSweepResult(
                Array.Empty<ComMailBrief>(),
                foldersSwept: 8,
                foldersSkipped: 0,
                sweptFolders: new[] { Store + "/Inbox", "Other Store/Inbox" },
                itemCappedFolders: new[] { Store + "/Inbox", "Other Store/Inbox" },
                perStore: new[]
                {
                    new ComStoreSweepCounters(Store, 4, 0, 0, 0),
                    new ComStoreSweepCounters("Other Store", 4, 0, 0, 0),
                },
                itemCappedFoldersUnsorted: new[] { Store + "/Inbox", "Other Store/Inbox" }),
            Store);

        Assert.Equal(new[] { Store + "/Inbox" }, info.ItemCappedFolders);
        Assert.Equal(new[] { Store + "/Inbox" }, info.ItemCappedFoldersUnsorted);
        Assert.Empty(FreshMerge.SortedItemCappedFolders(info));
    }

    // ================================================================== G5

    [Fact]
    public void AFolderTheIndexCannotAddress_IsReported_EvenWhenTheSweepFoundSomething()
    {
        // THE DEFECT. The sweep returns one item, so the merged answer is not empty, and the
        // old guard returned before it asked a single question. The index tier meanwhile
        // contributed nothing at all for this folder, which means everything in it older
        // than the sweep window is in neither tier.
        using MailService service = Service(Index(folderResolves: false), Sweep);

        SearchOutcome outcome = service.Search(FolderRequest());

        Assert.NotEmpty(outcome.Hits);
        Assert.True(outcome.Index!.FolderNotIndexed);
        Assert.Contains(outcome.Advice!, a => a.Contains(UnindexedFolder, StringComparison.Ordinal));
    }

    [Fact]
    public void AFolderTheIndexCanAddress_ReportsNothing()
    {
        // The probe pair still has to stay quiet over a folder that merely holds no match,
        // or the flag means nothing.
        using MailService service = Service(Index(folderResolves: true), Sweep);

        SearchOutcome outcome = service.Search(FolderRequest());

        Assert.Null(outcome.Index!.FolderNotIndexed);
        Assert.DoesNotContain(
            outcome.Advice ?? Array.Empty<string>(),
            a => a.Contains("did not resolve", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void AStoreWithNoIndexedRowsAtAll_DoesNotGetBlamedOnItsFolderPath()
    {
        // The store probe is the second half of the guard for a reason: with no rows for the
        // store either, "the folder did not resolve" is a guess, and that state has its own
        // reporting (storeNotIndexed / storesWithoutIndex / no_index_frontier).
        using MailService service = Service(Index(folderResolves: false, storeResolves: false), Sweep);

        SearchOutcome outcome = service.Search(FolderRequest());

        Assert.Null(outcome.Index!.FolderNotIndexed);
    }

    [Fact]
    public void AnEmptyAnswerOverAnUnaddressableFolder_StillReportsIt()
    {
        // The case that always worked, kept working: widening the trigger must not narrow it.
        using MailService service = Service(Index(folderResolves: false), EmptySweep);

        SearchOutcome outcome = service.Search(FolderRequest());

        Assert.Empty(outcome.Hits);
        Assert.True(outcome.Index!.FolderNotIndexed);
    }

    // ================================================================== B2

    [Fact]
    public void ADefaultSearch_SaysTheSweepCannotSeeInsideAttachments()
    {
        // freshness reads "live" and is not wrong about what the sweep checked - it is
        // wrong about what a caller takes "live" to mean, which is "nothing is missing".
        using MailService service = Service(Index(), Sweep);

        SearchOutcome outcome = service.Search(Request());

        Assert.False(outcome.Sweep!.AttachmentTextCovered);
        Assert.Contains(
            outcome.Advice!,
            a => a.Contains("Attachment CONTENT", StringComparison.Ordinal)
                && a.Contains("index tier alone", StringComparison.Ordinal));
    }

    [Fact]
    public void ASubjectOnlySearch_SaysNothingAboutAttachments()
    {
        // Nothing was asked of the body scope, so no tier could have matched an attachment
        // and there is no asymmetry to report. A flag that fires here would fire always.
        using MailService service = Service(Index(), Sweep);

        SearchOutcome outcome = service.Search(new SearchRequest
        {
            Query = "test",
            SearchIn = SearchIn.SubjectOnly,
            Top = 25,
            SnippetChars = 0,
        });

        Assert.Null(outcome.Sweep!.AttachmentTextCovered);
        Assert.DoesNotContain(
            outcome.Advice ?? Array.Empty<string>(),
            a => a.Contains("Attachment CONTENT", StringComparison.Ordinal));
    }

    [Fact]
    public void AnEmptyFreshnessWindow_CarriesTheFieldAndSkipsTheSentence()
    {
        // The field is a fact about the TIER, so it holds whatever arrived; the sentence is
        // about a hole that could exist, and an empty window cannot hide a mail. This is not
        // the G5 shape in disguise: it suppresses on provably nothing to report, not on the
        // answer happening to look full.
        using MailService service = Service(Index(), EmptySweep);

        SearchOutcome outcome = service.Search(Request());

        Assert.False(outcome.Sweep!.AttachmentTextCovered);
        Assert.Equal(0, outcome.Sweep.ItemsSeen);
        Assert.DoesNotContain(
            outcome.Advice ?? Array.Empty<string>(),
            a => a.Contains("Attachment CONTENT", StringComparison.Ordinal));
    }

    [Fact]
    public void TheAttachmentOnlyRefusal_IsNotSaidTwice()
    {
        // That case already had a sentence, and it names the same asymmetry. A second one
        // beside it would be noise on the one path that was never silent.
        using MailService service = Service(Index(), Sweep);

        SearchOutcome outcome = service.Search(new SearchRequest
        {
            Query = "test",
            AttachmentHitsOnly = true,
            Top = 25,
            SnippetChars = 0,
        });

        Assert.Equal(FreshMerge.AttachmentContentNotSweepable, outcome.Sweep!.Error);
        Assert.False(outcome.Sweep.AttachmentTextCovered);
        Assert.DoesNotContain(
            outcome.Advice!,
            a => a.Contains("Attachment CONTENT", StringComparison.Ordinal));
    }

    [Fact]
    public void TheAttachmentTextSentence_IsGeneratedFromTheFieldAlone()
    {
        // Pure, so the payload and the prose are one decision. A sentence with no field
        // behind it is what gap B2 is, one level up.
        Assert.Null(MailService.DescribeAttachmentTextGap(
            new SweepInfo { Performed = true, ItemsSeen = 9, AttachmentTextCovered = null }));
        Assert.Null(MailService.DescribeAttachmentTextGap(
            new SweepInfo { Performed = false, ItemsSeen = 9, AttachmentTextCovered = false }));
        Assert.Null(MailService.DescribeAttachmentTextGap(
            new SweepInfo { Performed = true, ItemsSeen = 0, AttachmentTextCovered = false }));
        Assert.NotNull(MailService.DescribeAttachmentTextGap(
            new SweepInfo { Performed = true, ItemsSeen = 9, AttachmentTextCovered = false }));
    }

    // ================================================================== F3

    [Fact]
    public void AnExhaustiveScanCappedBeforeItsFilterRan_SaysSo()
    {
        // 25 candidates matched the subject/body filter, the cap closed the walk, and 'from'
        // then discarded 24 of them. `truncated: true` beside ONE result reads as "a couple
        // more exist"; what it means is "the scan stopped after 25 candidates and most were
        // thrown away afterwards", and far more may lie further into the tree.
        using MailService service = Service(Index(), Sweep, CappedScan);

        SearchOutcome outcome = service.Search(ExhaustiveRequest(from: "bob@example.com"));

        Assert.Single(outcome.Hits);
        Assert.True(outcome.Truncated);
        Assert.Equal(new[] { "from" }, outcome.Exhaustive!.PostCapFilters);
        Assert.Equal(24, outcome.Exhaustive.ItemsFilteredOut);
        Assert.Contains(FreshMerge.ScanGapPostCapFilter, outcome.Exhaustive.CoverageGaps!);
        Assert.Contains(
            outcome.Advice!,
            a => a.Contains("counted CANDIDATES", StringComparison.Ordinal)
                && a.Contains("from", StringComparison.Ordinal));
    }

    [Fact]
    public void AnUncappedScan_RaisesNothing_EvenWithAPostFilter()
    {
        // With no cap the filter saw the whole matched set and removed exactly what the
        // caller asked to remove. Reporting that would be crying wolf on the mode working.
        using MailService service = Service(Index(), Sweep, s => Scan(s, truncated: false));

        SearchOutcome outcome = service.Search(ExhaustiveRequest(from: "bob@example.com"));

        Assert.NotNull(outcome.Exhaustive!.PostCapFilters);
        Assert.DoesNotContain(
            FreshMerge.ScanGapPostCapFilter,
            outcome.Exhaustive.CoverageGaps ?? Array.Empty<string>());
    }

    [Fact]
    public void ACappedScanWithNoPostFilter_RaisesOnlyTheResultCap()
    {
        // Nothing thinned the list after the cap, so the cap truncated results rather than
        // candidates and the existing sentence is the whole story.
        using MailService service = Service(Index(), Sweep, CappedScan);

        SearchOutcome outcome = service.Search(ExhaustiveRequest(from: null));

        Assert.Null(outcome.Exhaustive!.PostCapFilters);
        Assert.Equal(0, outcome.Exhaustive.ItemsFilteredOut);
        Assert.Contains(FreshMerge.ScanGapResultCap, outcome.Exhaustive.CoverageGaps!);
        Assert.DoesNotContain(FreshMerge.ScanGapPostCapFilter, outcome.Exhaustive.CoverageGaps!);
    }

    [Fact]
    public void ThePostCapFilterNames_AreTheCallersOwnParameters_InRequestOrder()
    {
        // The names are the remedy, so they have to be what the caller typed. before/after
        // are absent by design: they go into the DASL filter and bound the scan itself.
        Assert.Null(FreshMerge.PostCapFilters(false, false, false));
        Assert.Equal(
            new[] { "from", "unread_only", "has_attachments" },
            FreshMerge.PostCapFilters(true, true, true));
        Assert.Equal(new[] { "unread_only" }, FreshMerge.PostCapFilters(false, true, false));
    }

    // =============================================================== fixtures

    private static SearchRequest Request()
    {
        return new SearchRequest { Query = "test", Top = 25, SnippetChars = 0 };
    }

    private static SearchRequest FolderRequest()
    {
        return new SearchRequest
        {
            Query = "test",
            Store = Store,
            Folder = UnindexedFolder,
            Top = 25,
            SnippetChars = 0,
        };
    }

    private static SearchRequest ExhaustiveRequest(string? from)
    {
        return new SearchRequest
        {
            Query = "test",
            Store = Store,
            Folder = "Inbox",
            Exhaustive = true,
            From = from,
            Top = 25,
            SnippetChars = 0,
        };
    }

    /// <summary>
    /// The sweep with its coverage codes filled in, exactly as <c>Search</c> does before it
    /// asks for prose. The codes ARE the input to the sentences - that pairing is the whole
    /// shape - so a fixture that skipped this step would be testing a state the product
    /// never produces.
    /// </summary>
    private static SweepInfo Classified(SweepInfo sweep)
    {
        sweep.CoverageGaps = FreshMerge.DescribeCoverageGaps(sweep);
        return sweep;
    }

    private static SweepInfo CappedSweep(string? unsorted)
    {
        return new SweepInfo
        {
            Performed = true,
            FoldersSwept = 4,
            ItemCappedFolders = new[] { Store + "/Inbox" },
            ItemCappedFoldersUnsorted = unsorted == null ? null : new[] { unsorted },
        };
    }

    private static MailService Service(
        StubIndexClient index,
        Func<string?, ComSweepResult> sweep,
        Func<string?, ComExhaustiveResult>? scan = null)
    {
        return new MailService(
            new DirectGateway(StandInSession.Create(ProfileStores, sweep, scan)), null, index);
    }

    private static StubIndexClient Index(bool folderResolves = true, bool storeResolves = true)
    {
        return new StubIndexClient(folderResolves, storeResolves);
    }

    /// <summary>A sweep that reaches the store and finds one mail.</summary>
    private static ComSweepResult Sweep(string? onlyStore)
    {
        return new ComSweepResult(
            new[] { Mail("AA1") },
            foldersSwept: 4,
            foldersSkipped: 0,
            sweptFolders: new[] { Store + "/Inbox" },
            perStore: new[] { new ComStoreSweepCounters(Store, 4, 0, 0, 0) });
    }

    /// <summary>A sweep that ran, covered its scope and found no arrival at all.</summary>
    private static ComSweepResult EmptySweep(string? onlyStore)
    {
        return new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 4,
            foldersSkipped: 0,
            sweptFolders: new[] { Store + "/Inbox" },
            perStore: new[] { new ComStoreSweepCounters(Store, 4, 0, 0, 0) });
    }

    /// <summary>A sweep whose per-folder cap fired, with or without a usable sort.</summary>
    private static ComSweepResult CappedComSweep(bool unsorted)
    {
        return new ComSweepResult(
            new[] { Mail("AA1") },
            foldersSwept: 4,
            foldersSkipped: 0,
            sweptFolders: new[] { Store + "/Inbox" },
            itemCappedFolders: new[] { Store + "/Inbox" },
            perStore: new[] { new ComStoreSweepCounters(Store, 4, 0, 0, 0) },
            itemCappedFoldersUnsorted: unsorted ? new[] { Store + "/Inbox" } : null);
    }

    /// <summary>A scan the result cap stopped, carrying one match plus 24 other senders.</summary>
    private static ComExhaustiveResult CappedScan(string? store) => Scan(store, truncated: true);

    private static ComExhaustiveResult Scan(string? store, bool truncated)
    {
        List<ComMailBrief> items = new List<ComMailBrief> { Mail("EX0", "bob@example.com") };
        for (int i = 1; i < 25; i++)
        {
            items.Add(Mail("EX" + i.ToString(System.Globalization.CultureInfo.InvariantCulture), "carol@example.com"));
        }

        return new ComExhaustiveResult(
            items, foldersScanned: 3, foldersSkipped: 0, engine: "ci_phrasematch",
            instantSearchEnabled: true, truncated: truncated, timedOut: false);
    }

    private static ComMailBrief Mail(string entryId, string sender = "bob@example.com")
    {
        return new ComMailBrief(
            entryId: entryId,
            storeDisplayName: Store,
            storeId: "store-alice",
            folderName: "Inbox",
            folderKind: "inbox",
            subject: "a test mail",
            senderName: "Bob",
            senderAddress: sender,
            receivedTime: Frontier,
            isRead: true,
            hasAttachments: false,
            sizeBytes: 2048,
            body: "test body");
    }

    /// <summary>
    /// A Windows Search stand-in that knows one store and answers the probe statements by
    /// shape. It returns no SEARCH rows at all, which is the state every case here needs:
    /// the index tier contributes nothing and the question is what the answer then SAYS.
    /// <para>
    /// <c>folderResolves</c> / <c>storeResolves</c> drive the two TOP-1 existence probes
    /// independently, which is the pair of fixtures gap G5 needs and nothing had.
    /// </para>
    /// </summary>
    private sealed class StubIndexClient : IIndexClient
    {
        private const string DiscoveryTail = " System.ItemUrl FROM SystemIndex WHERE System.Kind='email'";

        private readonly bool _folderResolves;
        private readonly bool _storeResolves;

        internal StubIndexClient(bool folderResolves, bool storeResolves)
        {
            _folderResolves = folderResolves;
            _storeResolves = storeResolves;
        }

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
                // The folder probe scopes BELOW the store prefix; the store probe is the
                // prefix itself. Telling them apart is the whole point of the fixture.
                bool folderProbe = sql.Contains(StorePrefix + "/0/", StringComparison.OrdinalIgnoreCase)
                    || sql.Contains("ItemFolderPathDisplay", StringComparison.Ordinal);
                bool answers = folderProbe ? _folderResolves : _storeResolves;
                return answers ? Rows(StorePrefix + "/0/Inbox/probed-item") : Array.Empty<IReadOnlyDictionary<string, object?>>();
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
    /// A session that answers the profile's store list, the freshness sweep and the
    /// exhaustive scan, and refuses everything else. A <see cref="DispatchProxy"/> rather
    /// than a stub per member, so a method added to the contract needs no change here. Not
    /// sealed: DispatchProxy derives from its TProxy at runtime and refuses a sealed one.
    /// </summary>
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
