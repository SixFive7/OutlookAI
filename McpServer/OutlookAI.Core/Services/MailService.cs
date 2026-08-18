using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// Host-neutral orchestrator behind the MCP L1/L2 tools (v3.MD section 0.5): index
    /// search + always-on COM gap-sweep merge (D19/D34 - the sweep is cached ~10 s and
    /// degrades gracefully to index-only results when COM is unavailable), lazy hit
    /// location with caching (Phase-1 guidance: locate cost avg ~2 s - never locate
    /// eagerly, always cache), EntryID-based reads, attachment saving, thread lookup,
    /// account/folder listing and staleness self-reporting. No MCP types, no console
    /// assumptions; per-process hit cache only (a server process lives for one agent
    /// session).
    /// </summary>
    public sealed class MailService : IDisposable
    {
        // Live-verified (Phase 1: 25/25 within 5 s; Phase-2 run 1): a wide tolerance
        // lets a same-subject sibling mail within the window win the folder probe, so
        // email hits use a tight 5 s. Attachment (document) rows keep a wide window -
        // their DateReceived equals the parent's only approximately.
        private const int EmailLocateToleranceSeconds = 5;
        private const int AttachmentLocateToleranceSeconds = 120;
        private const int DedupeToleranceSeconds = 15;

        /// <summary>
        /// Soft budget handed to <c>ExhaustiveScan</c>, and the number the "results are
        /// partial" advice quotes.
        /// <para>
        /// DERIVED, never a literal: it used to be its own <c>120_000</c> and so equalled
        /// the COM host's hard operation deadline exactly. An inner budget equal to its
        /// outer one can never degrade gracefully - the scan stops only once elapsed has
        /// PASSED the budget and then still has to serialize its result set back over the
        /// pipe, while the watchdog fires at <c>&gt;=</c>. So the documented partial-results
        /// outcome was unreachable whenever the scan actually ran long: the caller got a
        /// Timeout, the COM host was killed, and two of those open the breaker for 30 s.
        /// </para>
        /// </summary>
        public const int ExhaustiveTimeBudgetMs = ComOperationBudgets.ChildWorkBudgetMs;

        /// <summary>
        /// What a non-folder-scoped sweep covers, echoed in the sweep block so an agent
        /// can see its freshness coverage (soak fix 13). Kept in sync with
        /// <see cref="OutlookComSession.DefaultSweepFolderKinds"/>.
        /// </summary>
        public const string DefaultSweepScopeDescription =
            "default folders (Inbox, Sent Items, Deleted Items, Junk Email)";

        /// <summary>
        /// Above this many swept folders the sweep block reports the count only - the
        /// list exists to make a narrow scope legible, not to bloat every payload
        /// (section-12 compact-payload discipline).
        /// </summary>
        public const int SweptFolderListCap = 12;

        /// <summary>
        /// Items ONE folder may contribute to a freshness sweep.
        /// <para>
        /// The sweep reads newest-first, so hitting this cap drops the OLDEST
        /// not-yet-indexed mail in that folder - which is exactly why it is not a silent
        /// cap. The affected folders are named in the response sweep block
        /// (<see cref="SweepInfo.ItemCappedFolders"/>, carried up from
        /// <see cref="ComSweepResult.ItemCappedFolders"/> by <c>ApplySweepCounters</c>) and
        /// in an advice line that quotes this number (<c>AddSweepCoverageAdvice</c>) - the
        /// same shape the folder cap, the time budget and the top clamp already use.
        /// </para>
        /// <para>
        /// Public so T1 pins it with every other payload cap: it was the one cap in this
        /// file declared privately and pinned by nothing, so creep here was the single cap
        /// change no test would have caught. The VALUE has no recorded measurement behind
        /// it and is unchanged.
        /// </para>
        /// </summary>
        public const int SweepPerFolderCap = 200;

        /// <summary>
        /// Age of the newest indexed mail above which the answer says so and points at
        /// exhaustive:true (12 h). Public so T1 pins the threshold together with the two
        /// wordings it selects between (<see cref="DescribeStaleIndex"/>).
        /// </summary>
        public const double VeryStaleAdviceMinutes = 720;

        /// <summary>
        /// Age of the newest indexed mail above which <c>outlook_health</c> mentions the lag
        /// at all (30 min). The NOTICE threshold, not the alarm one: it is far below
        /// <see cref="VeryStaleAdviceMinutes"/> because all it does is name a normal condition
        /// and say the sweep already covers it, whereas that one tells the agent to change how
        /// it searches. Sized above the ordinary quiet-mailbox gap - on the dev profile the
        /// median index frontier age is ~6 min and p90 ~30 min (2026-08-18, 177 probes), so it
        /// is the p90 of "nothing has arrived lately" and speaks up just past it. Public so T1
        /// pins it next to the threshold it must stay under.
        /// </summary>
        public const double StaleIndexNoticeMinutes = 30;

        /// <summary>
        /// Longest query <c>show_search_results</c> will put in Outlook's search box. Not our
        /// limit - it is the box's - and the tool's own error message quotes this constant
        /// rather than restating the number, which is the shape that went stale for the
        /// subject cap (<see cref="SubjectCharsCap"/>). The tool description deliberately does
        /// not repeat it: a query that long is an agent bug, not a thing to design around.
        /// </summary>
        public const int ShowSearchQueryCharsCap = 256;
        private static readonly TimeSpan SweepSafetyMargin = TimeSpan.FromMinutes(10);

        /// <summary>
        /// How far back the freshness sweep looks when there is no index frontier to open a
        /// window from - the scope has no indexed mail at all, so there is nothing to be
        /// "caught up to". Public so T1 pins it next to the advice sentence that quotes it:
        /// this span IS the reachable history of an unindexed store, and a change to it is a
        /// change to what such a search can find.
        /// </summary>
        public static readonly TimeSpan EmptyIndexSweepWindow = TimeSpan.FromDays(7);

        /// <summary>
        /// Budget for the post-sweep "is this store in the index at all" probes. Small on
        /// purpose: the probes only refine what the answer SAYS, never what it contains, and
        /// the stores that need them are the ones the catalog missed, i.e. usually none. A
        /// store left unprobed is left unreported rather than guessed at.
        /// <para>
        /// Measured 2026-08-18 on this machine (OleDb, 8 passes): a delegate-subtree miss is
        /// 9-10 ms across a two-store catalog and a targeted '@' discovery miss 27-30 ms, so
        /// the worst case for one unknown store is ~40 ms - and it is paid once per store per
        /// <see cref="StoreDetailsCacheTtl"/>, not once per search.
        /// </para>
        /// </summary>
        public const int StoreIndexProbeBudgetMs = 1_500;

        /// <summary>
        /// How long Outlook's store list is reused before it is re-read over COM. Both
        /// per-store caches share this one value - see <see cref="FolderPathCacheTtl"/>.
        /// </summary>
        private static readonly TimeSpan StoreDetailsCacheTtl = TimeSpan.FromMinutes(5);

        /// <summary>
        /// How long a store's COM folder-path list is reused. Only DELEGATE folder scopes
        /// need it (the flat index namespace has to be mapped back onto the real tree), so
        /// this keeps a per-search COM walk off the hot path.
        /// <para>
        /// It IS the store-details TTL, not a second copy of five minutes that a comment
        /// claimed matched it. The two caches describe the same profile from two angles and
        /// a search that mixes a fresh store list with a five-minute-old folder list (or the
        /// reverse) sees a profile that never existed.
        /// </para>
        /// </summary>
        private static readonly TimeSpan FolderPathCacheTtl = StoreDetailsCacheTtl;

        // ------------------------------------------------------------ payload caps
        // Section 12 compact-payload discipline, reviewed in Phase 7: every list a tool
        // can return is capped, every cap has a has-more/truncated indicator, and the
        // values are public constants so T1 tests pin them against accidental creep.

        /// <summary>Hard cap on search hits per call (tool default <see cref="SearchTopDefault"/>).</summary>
        public const int SearchTopCap = 100;

        /// <summary>Default search hit count - small on purpose: iterate instead.</summary>
        public const int SearchTopDefault = 25;

        /// <summary>Hard cap on per-hit snippet length.</summary>
        public const int SnippetCharsCap = 1000;

        /// <summary>Default per-hit snippet length.</summary>
        public const int SnippetCharsDefault = 200;

        /// <summary>Hard cap on thread members per call.</summary>
        public const int ThreadTopCap = 200;

        /// <summary>Default thread member count.</summary>
        public const int ThreadTopDefault = 50;

        /// <summary>Hard cap on read body characters.</summary>
        public const int BodyCharsCap = 500_000;

        /// <summary>Default read body cap.</summary>
        public const int BodyCharsDefault = 20_000;

        /// <summary>
        /// Default budget for the opt-in raw HTML (read include_html). It has its OWN
        /// default rather than sharing the text body's: Outlook's compose boilerplate is
        /// ~40 KB of stylesheet BEFORE any message content, so a 20 000-character window
        /// would hand the agent nothing but CSS and defeat the point of the option
        /// (measured live on this build, batch B). Still bounded and always reported via
        /// bodyHtmlTotalChars / bodyHtmlTruncated; the hard ceiling stays BodyCharsCap.
        /// </summary>
        public const int HtmlCharsDefault = 100_000;

        /// <summary>Minimum header cap (headers are opt-in).</summary>
        public const int HeaderCharsMin = 256;

        /// <summary>Hard cap on returned transport headers.</summary>
        public const int HeaderCharsCap = 65_536;

        /// <summary>Default header cap (8 KB).</summary>
        public const int HeaderCharsDefault = 8_192;

        /// <summary>Cap on recipients listed in read/draft/send payloads (flagged; operations always use ALL recipients).</summary>
        public const int RecipientsCap = 100;

        /// <summary>Cap on attachments listed in read payloads (flagged; higher indexes stay saveable).</summary>
        public const int AttachmentsCap = 100;

        /// <summary>Cap on unresolvable addresses echoed back by the draft tools (batch A, A2).</summary>
        public const int UnresolvedRecipientsCap = 20;

        /// <summary>
        /// Longest subject the draft tools accept, in characters.
        /// <para>
        /// The VALUE is inherited and unchanged: it is what the draft paths have always
        /// refused above, and no measurement or rationale was ever recorded for it. What is
        /// fixed here is the DUPLICATION - it was three bare <c>255</c> literals (new,
        /// derived, update) plus the number a FOURTH time as prose in the derived-subject
        /// tool hint, which is the shape that goes stale unnoticed. The check now lives in
        /// one place (<see cref="RequireSubjectWithinCap"/>) and T1
        /// <c>BudgetCompositionTests</c> derives the hint's phrase from this constant, so a
        /// changed cap fails the build instead of leaving the tool surface lying.
        /// </para>
        /// </summary>
        public const int SubjectCharsCap = 255;

        /// <summary>
        /// Shortest string accepted as a raw EntryID hex id, in hex characters. Two per byte
        /// of <see cref="Mapi.EntryIdCodec.MessageEntryIdLength"/> - the shortest entry id
        /// that can carry the flags, store UID and node id a MAPI entry id is made of.
        /// </summary>
        public const int MinRawEntryIdHexChars = Mapi.EntryIdCodec.MessageEntryIdLength * 2;

        /// <summary>
        /// Folders per list_folders page (section 12 discipline bound; real profiles fit
        /// in one page - offset paging exists for the pathological rest). Raised
        /// 500 -> 1000 (soak fix D38, user-ordered).
        /// </summary>
        public const int FoldersPerCallCap = 1000;

        /// <summary>Absolute guard on the underlying COM folder walk (pathological stores).</summary>
        public const int FolderWalkAbsoluteCap = 10_000;

        /// <summary>
        /// Advice appended to show_search_results whenever the EFFECTIVE registry state
        /// says Outlook's UI search is server-assisted (DisableServerAssistedSearch
        /// absent/0 - D22 coupling made self-documenting, D35). Deliberately strong: it
        /// must convey BOTH the divergence and the recommendation, because the agent's
        /// own search stays local/uncapped while the user-visible list silently differs.
        /// </summary>
        public const string ServerAssistedUiSearchAdvice =
            "Outlook UI search is currently server-assisted: the displayed results may differ from agent search results "
            + "(server-capped and differently ranked). Disabling server-assisted search is RECOMMENDED for consistent, "
            + "uncapped, fully local search - enable the Search tuning group in OutlookAI Settings.";

        private readonly Lazy<IndexSearchService> _index;
        private readonly IComGateway _gateway;
        private readonly SendConfirmationTokens _sendTokens;
        private readonly ServerDraftRegistry _draftRegistry = new ServerDraftRegistry();
        private readonly SweepCache _sweepCache = new SweepCache();
        private readonly BodyCache _bodies = new BodyCache();
        private readonly ConcurrentDictionary<string, CachedHit> _hits =
            new ConcurrentDictionary<string, CachedHit>(StringComparer.Ordinal);
        private readonly object _catalogLock = new object();
        private readonly Dictionary<string, (IReadOnlyList<string> Paths, DateTime FetchedUtc)> _folderPaths =
            new Dictionary<string, (IReadOnlyList<string>, DateTime)>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Whether the index holds anything for a store, by display name, with the instant it
        /// was established. Shares <see cref="StoreDetailsCacheTtl"/> and
        /// <see cref="_catalogLock"/> with the other two per-profile caches for the reason
        /// stated there: these are three angles on ONE profile and an answer mixing a fresh
        /// one with a five-minute-old one describes a profile that never existed.
        /// </summary>
        private readonly Dictionary<string, (bool Present, DateTime AtUtc)> _storeIndexPresence =
            new Dictionary<string, (bool, DateTime)>(StringComparer.OrdinalIgnoreCase);
        private string? _providerReport;
        private IReadOnlyList<StoreScopeInfo>? _catalog;
        private IReadOnlyList<ComStoreDetail>? _storeDetails;
        private DateTime _storeDetailsFetchedUtc;
        private int _nextHitId;

        /// <summary>Creates the service; both the index client and the COM session attach lazily.</summary>
        public MailService(IComGateway gateway)
            : this(gateway, null)
        {
        }

        /// <summary>
        /// Creates the service with an explicit send-confirmation token store (tests
        /// inject short-TTL stores; production uses the 120 s default).
        /// </summary>
        public MailService(IComGateway gateway, SendConfirmationTokens? sendTokens)
        {
            _gateway = gateway ?? throw new ArgumentNullException(nameof(gateway));
            _sendTokens = sendTokens ?? new SendConfirmationTokens();
            _index = new Lazy<IndexSearchService>(
                () => IndexSearchService.CreateDefault(out _providerReport),
                System.Threading.LazyThreadSafetyMode.ExecutionAndPublication);
        }

        /// <summary>Creates the default production instance (autostart per D17 enabled).</summary>
        public static MailService CreateDefault()
        {
            return new MailService(new ComGateway(allowStartingOutlook: true));
        }

        /// <summary>
        /// Time budget for the freshness sweep.
        /// <para>
        /// Much shorter than an ordinary operation, because the sweep is an ENHANCEMENT:
        /// search already has its indexed answer in hand before the sweep runs, and the
        /// tool's own contract calls search "sub-second and cheap". Measured healthy on
        /// this machine at 0.5-6 s; measured against a wedged Outlook it spent the full
        /// 120 s operation budget before degrading, which made every search feel broken
        /// even though the answer was already computed and waiting.
        /// </para>
        /// <para>
        /// It is a budget for the SWEEP, and the sweep call passes allowConnectFloor so the
        /// COM host may still add its cold-start connect allowance on a fresh host. Without
        /// that the very first search had to fit the COM attach AND the whole sweep into
        /// 30 s - on a machine where attaching to a large OST takes longer than that (the
        /// reason ConnectDeadlineMilliseconds is 90 s at all) the sweep could never succeed:
        /// every attempt timed out, killed the host, bumped the restart count and blamed the
        /// sweep.
        /// </para>
        /// </summary>
        public const int SweepBudgetMs = 30_000;

        /// <summary>
        /// Time budget for health's COM probe. Short by design: outlook_health exists to
        /// report an unresponsive Outlook, so it must never wait the ordinary operation
        /// budget to discover one. Shared with the supervisor's own health-probe deadline
        /// rather than declared as a second 5 000.
        /// </summary>
        public const int HealthProbeBudgetMs = ComOperationBudgets.HealthProbeDeadlineMs;

        /// <summary>Per-query index timeout used by health only.</summary>
        public const int HealthIndexTimeoutSeconds = 4;

        /// <summary>
        /// Overall budget for health's WHOLE index block: the global freshness probe, the
        /// COM-assisted catalog enrichment and the per-store rows.
        /// <para>
        /// It used to cover only the per-store loop, so the two steps in front of it -
        /// catalog discovery and the per-address enrichment searches - ran outside the
        /// budget they were meant to be inside, on the 30 s default per query. A profile
        /// with many stores multiplies even a short timeout. Reporting fewer rows is a fine
        /// outcome; taking minutes to report all of them is not.
        /// </para>
        /// </summary>
        public const int HealthPerStoreIndexBudgetMs = 8_000;

        /// <summary>
        /// Per-query index timeout on the SEARCH path.
        /// <para>
        /// Search's own description promises "sub-second and cheap", and one search is
        /// several index statements plus the freshness sweep. On
        /// <see cref="OleDbIndexClient.DefaultCommandTimeoutSeconds"/> each of those
        /// statements could take 30 s on its own, so the composed worst case had no relation
        /// to what the tool advertises. Measured healthy on this machine at 60-550 ms, so
        /// this is roughly 27x headroom: exceeding it means the indexer is saturated, and
        /// search says so and degrades rather than waiting twice as long to say the same
        /// thing.
        /// </para>
        /// </summary>
        public const int SearchIndexTimeoutSeconds = 15;

        /// <summary>
        /// The tool-level wall-clock shape of one indexed search, stated as a relationship
        /// rather than as an unrelated literal: index statement plus freshness sweep. Pinned
        /// against the COM host's operation deadline so the two cannot drift into a search
        /// that outlives the budget its own sweep runs under.
        /// </summary>
        public const int SearchBudgetMs = (SearchIndexTimeoutSeconds * 1000) + SweepBudgetMs;

        /// <summary>
        /// Aggregate budget for one move_mail / archive_mail batch.
        /// <para>
        /// Each item is 2-3 gateway calls, each bounded on its own, and up to
        /// <see cref="MoveIdsCap"/> items ran with no bound across the batch and no
        /// cancellation checkpoint between them - a theoretical 150 round trips at a full
        /// operation deadline each. Items still attempted after this elapses are reported as
        /// not attempted, exactly like the audit-log short circuit beside it, so a partial
        /// batch stays legible and every EntryID that did move is still returned.
        /// </para>
        /// </summary>
        public const int MoveBatchBudgetMs = ComOperationBudgets.OperationDeadlineMs;

        /// <summary>Default directory attachments are saved to when the caller names none.</summary>
        public static string DefaultAttachmentDirectory =>
            Path.Combine(SharedStateDirectory, "scratch", "attachments");

        /// <summary>Shared OutlookAI state root (v3.MD section 0.5.2: %LOCALAPPDATA%\OutlookAI).</summary>
        public static string SharedStateDirectory =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookAI");

        /// <summary>Provider report captured when the index client was created (diagnostics).</summary>
        public string? ProviderReport
        {
            get
            {
                _ = _index.Value;
                return _providerReport;
            }
        }

        // ------------------------------------------------------------------ search

        /// <summary>
        /// Runs one search (v3.MD section 8 L1, D34): index query + freshness gap-sweep
        /// merged and deduped - the sweep is always on, served from a ~10 s cache for
        /// rapid-fire iteration, and degrades to index-only results (with advice) when
        /// it cannot run. exhaustive:true switches to the bounded index-bypassing COM scan.
        /// </summary>
        public SearchOutcome Search(SearchRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            int top = Clamp(request.Top, 1, SearchTopCap);
            int snippetChars = Clamp(request.SnippetChars, 0, SnippetCharsCap);
            IReadOnlyList<string> terms = SplitTerms(request.Query);
            if (request.Folder != null && request.Store == null)
            {
                throw new ArgumentException("folder requires store.", nameof(request));
            }

            if (request.Exhaustive)
            {
                return RunExhaustive(request, terms, top);
            }

            FolderScopeResolution? folderScope = null;
            string? scope = null;
            if (request.Store != null)
            {
                folderScope = ResolveFolderScope(request.Store, request.Folder, request.IncludeSubfolders);
                scope = folderScope.Scope;
            }

            IndexQuery query = new IndexQuery
            {
                Scope = scope,
                FolderPathsAnyOf = folderScope?.FolderPaths,
                Terms = terms.Count > 0 ? terms : null,
                SearchIn = request.SearchIn,
                Kinds = request.AttachmentHitsOnly
                    ? KindFilter.DocumentsOnly
                    : request.IncludeAttachmentHits ? KindFilter.EmailAndDocuments : KindFilter.EmailOnly,
                FromAddressContains = request.From,
                RecipientContains = request.To,
                ReceivedOnOrAfterUtc = request.AfterUtc,
                ReceivedBeforeUtc = request.BeforeUtc,
                IsRead = request.UnreadOnly == true ? false : (bool?)null,
                HasAttachments = request.HasAttachments,
                OrderBy = request.OrderBySizeDescending ? IndexOrder.SizeDescending : IndexOrder.DateReceivedDescending,
                // Over-fetch by ONE row: a full page then means "more matches EXIST",
                // making the truncated flag definite (section 12 has-more discipline).
                Top = top + 1,
            };

            IndexSearchResult indexResult = _index.Value.Search(query, SearchIndexTimeoutSeconds);

            // The frontier is measured over the SCOPE being searched, not over the profile
            // (StalenessScopeFor): it sets this search's sweep window, and a busy store's
            // frontier would otherwise pin a quiet store's window to the last few minutes
            // while that store's own index lagged by hours.
            IndexStalenessReport staleness =
                _index.Value.GetStaleness(StalenessScopeFor(folderScope), SearchIndexTimeoutSeconds);

            bool truncated = indexResult.Hits.Count > top;
            List<HitSummary> summaries = new List<HitSummary>(Math.Min(indexResult.Hits.Count, top));
            foreach (IndexHit hit in indexResult.Hits)
            {
                if (summaries.Count >= top)
                {
                    break; // The over-fetched row is evidence, not a result.
                }

                summaries.Add(RegisterIndexHit(hit, snippetChars));
            }

            SweepInfo? sweep = null;
            List<string> advice = new List<string>();

            // The folder bound is reported BEFORE anything else: an agent that gets a
            // widened or name-matched delegate scope must see that first (constraints
            // C2/C3). The zero-row guard (C7) runs AFTER the sweep instead - the freshness
            // sweep can still supply hits for a folder the index has never seen, and
            // "the folder did not resolve" must not be said over a non-empty answer.
            AddFolderScopeAdvice(advice, folderScope);

            if (!request.IndexOnly)
            {
                sweep = RunGapSweep(
                    request, terms, staleness, indexResult.Hits, summaries, snippetChars,
                    out DateTime? widestFrontierUtc);

                // The exposure these sentences quote is the OLDEST frontier in scope, not the
                // profile-wide one. Now that the sweep opens a window per store, an unscoped
                // search whose sweep fails is index-only against every store's OWN lag, and
                // the profile figure - the newest instant ANY store ingested - understates
                // that by exactly the spread the per-store windows exist to cover (measured
                // 11 min 19 s between two stores on this machine, 45.4 h across three the day
                // before). Same number for a store-scoped search, which has one frontier.
                string indexAge = DescribeAge(staleness, widestFrontierUtc);

                // The conclusion the counters add up to, computed ONCE and carried in the
                // payload: it decides the advice below, the top-level freshness value and
                // the degraded flag, so those three can never disagree with each other.
                sweep.CoverageGaps = FreshMerge.DescribeCoverageGaps(sweep);
                if (sweep.Error == FreshMerge.RecipientFilterNotSweepable)
                {
                    advice.Add("Freshness sweep skipped: recipient ('to') filters cannot be matched by the sweep, so results are "
                        + "index-only and may lag the last " + indexAge + " of mail.");
                }
                else if (sweep.Error == FreshMerge.AttachmentContentNotSweepable)
                {
                    advice.Add("Freshness sweep skipped: attachment content is matched by the index only, so an attachment-only "
                        + "search is index-only by construction and does not cover the last " + indexAge
                        + " of mail. Search without the attachment-only filter to get freshness coverage of subject and body.");
                }
                else if (sweep.Error != null)
                {
                    advice.Add("INCOMPLETE RESULTS - TELL THE USER: these are indexed results only and may be missing mail "
                        + "from roughly the last " + indexAge + ". The live check against Outlook could not "
                        + "run (" + sweep.Error + "). Everything already indexed is here and correct; only very recent "
                        + "mail may be absent. " + (ComGateway.IsInstallerMutexHeld()
                            ? "An add-in update is in progress - retry shortly (D17)."
                            : "Retry shortly, check outlook_health, or search again with exhaustive:true plus store + "
                              + "folder/after bounds for an index-free COM search."));
                }
                else if (sweep.Performed && sweep.FoldersSwept == 0 && request.Folder != null)
                {
                    // The folder-scoped sweep could not resolve the folder through COM
                    // (renamed, gone, or not a mail folder): index results still stand,
                    // but this query has no freshness coverage - say so rather than
                    // implying the last few minutes were checked.
                    advice.Add("Freshness sweep covered no folder: '" + request.Folder
                        + "' could not be opened in Outlook, so results for it are index-only and may lag the last "
                        + indexAge + " of mail. Check the path with list_folders.");
                }

                advice.AddRange(DescribeSweepCoverage(sweep, indexAge, request.Folder != null));
            }

            AddUnresolvedFolderAdvice(advice, folderScope, request, summaries.Count);

            // Snapshot AFTER the sweep: the sweep may have just autostarted Outlook
            // (D17) and the staleness block must reflect that reality, not the
            // pre-autostart state (D34 self-consistency fix).
            //
            // A sweep that was NOT NEEDED is excluded: "recent mail may be missing" names a
            // risk this query cannot run - its window ends before the index frontier - and
            // saying it in prose would just relocate the false alarm the notNeeded state
            // exists to remove, into the field an agent is told to relay to the user.
            bool outlookRunning = ComGateway.IsOutlookRunning();
            if (!outlookRunning
                && (sweep == null || (!sweep.Performed && sweep.Error == null && sweep.NotNeeded != true)))
            {
                advice.Add("Outlook is not running, so the index is frozen; recent mail may be missing until Outlook runs again.");
            }

            string? staleAdvice = DescribeStaleIndex(staleness.Age?.TotalMinutes, request.Store != null);
            if (staleAdvice != null)
            {
                advice.Add(staleAdvice);
            }

            summaries.Sort((a, b) => DateTime.Compare(b.ReceivedUtc ?? DateTime.MinValue, a.ReceivedUtc ?? DateTime.MinValue));
            if (summaries.Count > top)
            {
                summaries.RemoveRange(top, summaries.Count - top);
                truncated = true;
            }

            if (truncated)
            {
                advice.Add("Result list capped at " + top.ToString(CultureInfo.InvariantCulture)
                    + " (top); more matches exist - raise top (max " + SearchTopCap.ToString(CultureInfo.InvariantCulture)
                    + ") or narrow with store/folder/from/after filters.");
            }

            AddTopClampAdvice(advice, request.Top, top);

            if (indexResult.CandidatesExhausted)
            {
                // The index tier admits rows in code (attachment rows of every kind, mail
                // messages only) over an over-fetched candidate list. Running out of
                // candidates is the one way that could hide matches - say so (D42).
                advice.Add("The index tier ran out of candidate rows while filtering non-mail entries, "
                    + "so this list may be short of matches. Narrow with store/folder/after, or lower top.");
            }

            // Say the live check's outcome in a FIELD, not only in prose: a result that
            // looks complete but silently lags recent mail is the one failure here that can
            // mislead rather than merely inconvenience. Three states, because a sweep that
            // ran and covered part of its scope is neither of the other two - it did run, so
            // it is not "index-only", and it left mail unchecked, so it is not "live".
            string freshness = FreshMerge.ClassifyFreshness(sweep);

            return new SearchOutcome
            {
                Hits = summaries,
                Truncated = truncated,
                Degraded = freshness == FreshMerge.FreshnessLive ? (bool?)null : true,
                Freshness = freshness,
                IndexElapsedMs = indexResult.ElapsedMilliseconds,
                Sweep = sweep,
                Scope = DescribeSearchScope(folderScope, request),
                Staleness = new StalenessInfo
                {
                    NewestIndexedUtc = staleness.NewestIndexedReceivedUtc,
                    AgeMinutes = staleness.Age?.TotalMinutes,
                    OutlookRunning = outlookRunning,
                },
                Advice = advice.Count > 0 ? advice : null,
            };
        }

        /// <summary>
        /// The index SCOPE the freshness frontier is measured over: the STORE the search is
        /// scoped to, or null (the whole profile) when it names no store.
        /// <para>
        /// MEASURED 2026-08-18: the probe ran unscoped for every search, so five
        /// store-scoped searches all reported the same frontier while outlook_health's
        /// per-store probe reported three values spanning 45.4 hours. That frontier sets the
        /// sweep window, so a busy store held a quiet store's window down to minutes and
        /// recent mail in the quiet store fell outside it - and the result reported itself
        /// fresh.
        /// </para>
        /// <para>
        /// STORE level, not folder level, even for a folder-scoped search. This number is an
        /// INGESTION frontier, and it only approximates one where mail ARRIVES: a store
        /// indexes as its own catalog subtree, so its newest indexed item tracks how far its
        /// ingestion has got. A quiet Archive folder's newest item says nothing about
        /// ingestion - it is old because nothing arrives there - so scoping the probe to a
        /// folder would widen that search's sweep window to years and make it read hundreds
        /// of already-indexed items per folder for nothing.
        /// </para>
        /// <para>
        /// Pure, and public so T1 pins the choice: the alternatives (profile-wide, or the
        /// folder URL) are both one field away, and neither is distinguishable from this one
        /// at any call site.
        /// </para>
        /// </summary>
        public static string? StalenessScopeFor(FolderScopeResolution? folderScope)
        {
            return folderScope?.StoreScope;
        }

        /// <summary>
        /// The "the index may be far behind" advice, or null when it is not far enough
        /// behind to be worth saying (<see cref="VeryStaleAdviceMinutes"/>).
        /// <para>
        /// It has TWO wordings because the number now means two things. Unscoped it is the
        /// whole profile's frontier and a large age really does mean a lagging index. Scoped
        /// to one store it is that store's, where a large age has a second, likelier cause:
        /// the store is quiet. Saying "the index is very stale" over a low-traffic account
        /// would state as fact something the number cannot distinguish - and scope-aware
        /// staleness makes that the common case rather than the rare one.
        /// </para>
        /// <para>
        /// The remedy is the same in both, because it does not depend on which cause it is.
        /// </para>
        /// </summary>
        public static string? DescribeStaleIndex(double? ageMinutes, bool storeScoped)
        {
            if (!ageMinutes.HasValue || ageMinutes.Value <= VeryStaleAdviceMinutes)
            {
                return null;
            }

            string hours = (ageMinutes.Value / 60).ToString("F0", CultureInfo.InvariantCulture);
            return (storeScoped
                    ? "The newest indexed mail in this store is " + hours
                        + " h old - either the store is quiet or its index is behind. "
                    : "The index is very stale (" + hours + " h behind). ")
                + "For correctness-critical queries search again with exhaustive:true (bounded COM scan, store + "
                + "folder/after required) - it bypasses the index entirely.";
        }

        /// <summary>
        /// A caller <c>top</c> above <see cref="SearchTopCap"/> is clamped, not rejected -
        /// but a silent clamp makes a 500-hit request look like a complete 100-hit answer.
        /// Section-12 discipline: every cap is reported (soak fix 15).
        /// </summary>
        private static void AddTopClampAdvice(List<string> advice, int requestedTop, int effectiveTop)
        {
            if (requestedTop <= effectiveTop)
            {
                return;
            }

            advice.Add("top=" + requestedTop.ToString(CultureInfo.InvariantCulture) + " was reduced to "
                + effectiveTop.ToString(CultureInfo.InvariantCulture) + " (the hard cap): search returns at most "
                + SearchTopCap.ToString(CultureInfo.InvariantCulture)
                + " hits per call. Narrow with store/folder/from/after, or page by moving the 'before' bound.");
        }

        /// <summary>
        /// Reports every way the folder bound differs from what was asked (v3.MD
        /// constraints C2/C3). Nothing here changes the result set - it exists so the
        /// result set can never be misread. The C7 zero-row guard is separate
        /// (<see cref="AddUnresolvedFolderAdvice"/>): it can only judge the MERGED answer.
        /// </summary>
        private static void AddFolderScopeAdvice(List<string> advice, FolderScopeResolution? folderScope)
        {
            if (folderScope == null || folderScope.RequestedFolder == null)
            {
                return;
            }

            if (folderScope.Widened)
            {
                advice.Add(FolderScopeResolver.DescribeWidening(folderScope));
            }

            if (folderScope.CollidingLeafNames != null && folderScope.CollidingLeafNames.Count > 0)
            {
                advice.Add(FolderScopeResolver.DescribeCollision(folderScope));
            }
            else if (folderScope.IsDelegateStore && folderScope.FolderTreeUnavailable && !folderScope.Widened)
            {
                advice.Add("Outlook could not be reached, so this delegate folder scope was not checked against the "
                    + "mailbox's folder tree: delegate mailboxes are indexed by folder NAME only, so a same-named "
                    + "folder elsewhere in that mailbox would also be included.");
            }
        }

        /// <summary>
        /// The non-silent zero-row guard (v3.MD constraint C7). Two TOP-1 probes, only
        /// ever on a FULLY empty answer (index plus freshness sweep): first "does this
        /// folder bound match ANY row" - so a folder that merely holds no match stays
        /// quiet, which the one-probe form would not - then "does the store". Rows in the
        /// store but none for the folder means the PATH did not resolve, which is the
        /// failure mode that hid the delegate defect.
        /// </summary>
        private void AddUnresolvedFolderAdvice(
            List<string> advice, FolderScopeResolution? folderScope, SearchRequest request, int hitCount)
        {
            if (hitCount > 0 || folderScope == null || folderScope.RequestedFolder == null)
            {
                return;
            }

            try
            {
                if (_index.Value.FolderScopeHasAnyItem(folderScope.Scope, folderScope.FolderPaths, SearchIndexTimeoutSeconds))
                {
                    return;
                }

                if (_index.Value.FolderScopeHasAnyItem(folderScope.StoreScope, null, SearchIndexTimeoutSeconds))
                {
                    advice.Add(FolderScopeResolver.DescribeUnresolvedFolder(request.Folder, request.Store!));
                }
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                // The guard is a diagnostic; a failing probe must never fail the search.
            }
        }

        /// <summary>
        /// Narrates every freshness-coverage hole the sweep just reported. Before soak
        /// fix 15 all of these landed in the payload as bare integers with no advice
        /// branch anywhere - or, for a failed folder, were not reported at all.
        /// <para>
        /// Driven by <see cref="SweepInfo.CoverageGaps"/> rather than by its own copy of
        /// the conditions: the codes and these sentences are then two renderings of ONE
        /// decision (<see cref="FreshMerge.DescribeCoverageGaps"/>) instead of two lists
        /// that have to be kept in step, and every code is guaranteed a sentence. That
        /// pairing is what T1 pins - a gap the payload flags but the prose never explains
        /// (or the reverse) is exactly the drift this shape removes.
        /// </para>
        /// <para>
        /// Public and pure so T1 can exercise every hole without a mailbox: the states
        /// below need a folder tree that fails, truncates or runs long, which no CI runner
        /// has. <paramref name="indexAge"/> is the already-formatted staleness span, and
        /// <paramref name="folderScoped"/> suppresses the total-miss sentence for a
        /// folder-scoped search, whose caller emits a better one naming the folder.
        /// </para>
        /// </summary>
        public static IReadOnlyList<string> DescribeSweepCoverage(SweepInfo sweep, string indexAge, bool folderScoped)
        {
            if (sweep == null)
            {
                throw new ArgumentNullException(nameof(sweep));
            }

            List<string> advice = new List<string>();
            foreach (string gap in sweep.CoverageGaps ?? Array.Empty<string>())
            {
                switch (gap)
                {
                    case FreshMerge.GapNoIndexFrontier:
                        advice.Add("Freshness sweep is the ONLY tier covering "
                            + (sweep.StoresWithoutIndex != null && sweep.StoresWithoutIndex.Count > 0
                                ? "store(s) " + string.Join(", ", sweep.StoresWithoutIndex)
                                : "this profile")
                            + ": the local index holds no mail there, so there was no frontier to sweep up to and the "
                            + "window fell back to the last "
                            + EmptyIndexSweepWindow.TotalDays.ToString("F0", CultureInfo.InvariantCulture)
                            + " days of the arrival-path folders. Mail older than that, or filed anywhere else there, "
                            + "is in NEITHER tier - it is not findable by search at all. Confirm with outlook_health "
                            + "(index.perStore) or list_accounts (locallySearchable), and use exhaustive:true with "
                            + "store plus folder/after to read that store without the index.");
                        break;

                    case FreshMerge.GapNothingSwept:
                        if (!folderScoped)
                        {
                            advice.Add("Freshness sweep covered NO folder at all, so nothing in this answer was checked "
                                + "against live Outlook and it may lag the last " + indexAge + " of mail. Check the store "
                                + "name with list_accounts; index results are unaffected.");
                        }

                        break;

                    case FreshMerge.GapFoldersFailed:
                        advice.Add("Freshness sweep FAILED on " + sweep.FoldersFailed.ToString(CultureInfo.InvariantCulture)
                            + " folder(s) - Outlook would not enumerate them, so mail that arrived there in the last "
                            + indexAge + " is missing from these results. Retry, or use exhaustive:true for that folder.");
                        break;

                    case FreshMerge.GapFolderCap:
                        advice.Add("Freshness sweep stopped at its folder cap ("
                            + OutlookComSession.MaxScopedSweepFolders.ToString(CultureInfo.InvariantCulture)
                            + " folders visited), so deeper subfolders were never visited - index results still cover them, but "
                            + "brand-new mail there may be missing. Scope the search to a narrower folder for full freshness coverage.");
                        break;

                    case FreshMerge.GapTimeBudget:
                        advice.Add("Freshness sweep stopped at its "
                            + (OutlookComSession.ScopedSweepTimeBudgetMs / 1000).ToString(CultureInfo.InvariantCulture)
                            + " s time budget after " + sweep.FoldersSwept.ToString(CultureInfo.InvariantCulture)
                            + " folder(s), so the rest of the subtree has no freshness coverage - index results still cover it, "
                            + "but brand-new mail there may be missing. Scope the search to a narrower folder, or pass "
                            + "include_subfolders:false to sweep just the named folder.");
                        break;

                    case FreshMerge.GapDepthLimit:
                        advice.Add("Freshness sweep refused to descend past its depth guard, so the deepest folders in this "
                            + "subtree were never swept - a folder tree that deep is unusual enough to be worth checking with "
                            + "list_folders. Index results still cover them.");
                        break;

                    case FreshMerge.GapFoldersSkipped:
                        advice.Add("Freshness sweep skipped " + sweep.FoldersSkipped.ToString(CultureInfo.InvariantCulture)
                            + " folder(s) it could not resolve or enumerate, so mail that arrived there in the last "
                            + indexAge + " may be missing. Check paths with list_folders.");
                        break;

                    case FreshMerge.GapItemCap:
                        advice.Add("Freshness sweep hit its per-folder cap of "
                            + SweepPerFolderCap.ToString(CultureInfo.InvariantCulture) + " items in: "
                            + string.Join(", ", sweep.ItemCappedFolders ?? Array.Empty<string>())
                            + ". It reads newest-first, so the OLDEST not-yet-indexed mail in those folders is not covered - "
                            + "narrow the window with 'after' or search those folders directly.");
                        break;

                    default:
                        // A code with no sentence would be a silent partial result, which is
                        // the whole defect this reporting exists to remove. T1 pins that
                        // every code is handled, so this can only be reached by a code added
                        // without its advice - say so rather than dropping it.
                        advice.Add("Freshness sweep reported partial coverage (" + gap
                            + ") with no further detail available; treat these results as incomplete.");
                        break;
                }
            }

            // Not a coverage hole: the sweep covered these folders, the payload just does
            // not list them all. Reported for the same no-silent-caps reason, outside the
            // gap set so it never marks a complete answer partial.
            if (sweep.FolderListOmitted == true)
            {
                advice.Add("The swept-folder list is omitted above " + SweptFolderListCap.ToString(CultureInfo.InvariantCulture)
                    + " folders (payload discipline); sweep.foldersSwept is the true count.");
            }

            return advice;
        }

        /// <summary>Compact scope block; present only for folder-scoped searches.</summary>
        private static SearchScopeInfo? DescribeSearchScope(FolderScopeResolution? folderScope, SearchRequest request)
        {
            if (folderScope == null || folderScope.RequestedFolder == null)
            {
                return null;
            }

            string shape = folderScope.Kind switch
            {
                FolderScopeKind.PrimaryRecursive => "folder",
                FolderScopeKind.PrimaryNonRecursive => "folder_only",
                FolderScopeKind.DelegateFlat => "delegate_folders",
                FolderScopeKind.DelegateWidened => "delegate_store_widened",
                _ => "store",
            };

            return new SearchScopeInfo
            {
                Folder = request.Folder,
                IncludeSubfolders = request.IncludeSubfolders,
                Shape = shape,
                Widened = folderScope.Widened ? true : (bool?)null,
                FolderNamesMatched = folderScope.IsDelegateStore && folderScope.FolderPaths != null
                    ? folderScope.FolderPaths.Count
                    : (int?)null,
            };
        }

        /// <summary>
        /// Drops the cached freshness sweep so the next search sweeps live. Test and
        /// diagnostic use (arrival-latency measurements); agents never need it - the
        /// cache self-invalidates on frontier advance or after its ~10 s TTL (D34).
        /// </summary>
        public void ClearSweepCache()
        {
            _sweepCache.Clear();
        }

        /// <summary>
        /// The sweep window(s) for one search: one start per store in scope, plus the fallback
        /// a store with no index frontier gets, plus what could not be established at all.
        /// </summary>
        private sealed class SweepWindowPlan
        {
            /// <summary>Window start for a store with no frontier of its own (unclamped).</summary>
            internal DateTime FallbackBaseUtc { get; set; }

            /// <summary>Per-store window starts (unclamped), empty for a store-scoped search.</summary>
            internal Dictionary<string, DateTime>? PerStoreBaseUtc { get; set; }

            /// <summary>A scope in this search had no index frontier at all.</summary>
            internal bool FrontierMissing { get; set; }

            /// <summary>The stores behind <see cref="FrontierMissing"/>, where nameable.</summary>
            internal List<string>? StoresWithoutIndex { get; set; }

            internal void AddStoreWithoutIndex(string store)
            {
                FrontierMissing = true;
                StoresWithoutIndex ??= new List<string>();
                if (!StoresWithoutIndex.Contains(store, StringComparer.OrdinalIgnoreCase))
                {
                    StoresWithoutIndex.Add(store);
                }
            }
        }

        /// <summary>
        /// One sweep window PER STORE, each opened from that store's own index frontier.
        /// <para>
        /// A store-scoped search already had this: its frontier probe is scoped to the store
        /// (<see cref="StalenessScopeFor"/>), so <paramref name="staleness"/> IS that store's
        /// window base and nothing more is probed. An UNSCOPED search did not - it opened one
        /// window from the profile-wide frontier, which is the newest instant any store
        /// ingested, so a store lagging by hours was swept back only as far as the busiest
        /// store's clock and the rest of its gap sat in neither tier: not indexed yet, and
        /// before the window. Half of the per-store fix landed and read like all of it.
        /// </para>
        /// <para>
        /// COST, measured rather than assumed (2026-08-18, this machine, OleDb over the live
        /// SystemIndex): a scoped frontier probe is 14-21 ms, so a two-store catalog costs
        /// 33-39 ms median per unscoped search - against a sweep budget of 30 s and a
        /// per-query index timeout of 15 s. Catalog discovery itself (115-141 ms) is once per
        /// process and every store-scoped search already paid it. The same runs measured the
        /// reason it is worth paying: the two stores' frontiers sat 11 min 19 s apart, so the
        /// single profile-wide window was eleven minutes short for one of them, silently.
        /// That is the cost of the completeness, stated so it is known, not so it can be
        /// traded against - a store swept from another store's clock returns less than it
        /// should and says nothing.
        /// </para>
        /// <para>
        /// THE FALLBACK IS THE WIDEST WINDOW, NOT THE PROFILE FRONTIER. A store the index
        /// catalog does not know is exactly the store whose frontier we could not measure, so
        /// giving it the profile-wide value would hand the one store that needs a wide window
        /// the narrowest one on the profile. It gets <see cref="EmptyIndexSweepWindow"/>
        /// instead: over-covering costs sweep time and is bounded by the per-folder cap;
        /// under-covering loses mail silently.
        /// </para>
        /// </summary>
        private SweepWindowPlan ResolveSweepWindows(SearchRequest request, IndexStalenessReport staleness)
        {
            DateTime emptyIndexBase = DateTime.UtcNow - EmptyIndexSweepWindow;
            SweepWindowPlan plan = new SweepWindowPlan();

            if (request.Store != null)
            {
                // Scoped: the search's own probe was already scoped to this store.
                plan.FallbackBaseUtc = staleness.NewestIndexedReceivedUtc ?? emptyIndexBase;
                if (!staleness.NewestIndexedReceivedUtc.HasValue)
                {
                    plan.AddStoreWithoutIndex(request.Store);
                }

                return plan;
            }

            plan.FallbackBaseUtc = emptyIndexBase;
            if (!staleness.NewestIndexedReceivedUtc.HasValue)
            {
                // The profile-wide probe found no mail ANYWHERE: an unindexed profile, which
                // is the shape a local-PST-only Outlook takes when Windows Search never
                // indexed it. No store names to give - there is no catalog to name them from.
                plan.FrontierMissing = true;
                return plan;
            }

            try
            {
                foreach (StoreScopeInfo scopeInfo in GetCatalog(SearchIndexTimeoutSeconds))
                {
                    IndexStalenessReport scoped =
                        _index.Value.GetStaleness(scopeInfo.StorePrefix, SearchIndexTimeoutSeconds);
                    if (!scoped.NewestIndexedReceivedUtc.HasValue)
                    {
                        // In the catalog because it has indexed ITEMS, but none of them mail.
                        plan.AddStoreWithoutIndex(scopeInfo.StoreDisplayName);
                        continue;
                    }

                    plan.PerStoreBaseUtc ??= new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
                    if (!plan.PerStoreBaseUtc.TryGetValue(scopeInfo.StoreDisplayName, out DateTime existing)
                        || scoped.NewestIndexedReceivedUtc.Value < existing)
                    {
                        // Two prefixes parsing to one display name cannot both be this
                        // store's frontier; the earlier one is the window that cannot hide
                        // mail, so it wins.
                        plan.PerStoreBaseUtc[scopeInfo.StoreDisplayName] = scoped.NewestIndexedReceivedUtc.Value;
                    }
                }
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                // A frontier probe is an optimisation of the WINDOW, never a precondition of
                // the search: whatever was not probed keeps the widest window, which
                // over-covers. Failing the search here would turn a slow indexer into no
                // answer at all.
            }

            return plan;
        }

        private SweepInfo RunGapSweep(
            SearchRequest request,
            IReadOnlyList<string> terms,
            IndexStalenessReport staleness,
            IReadOnlyList<IndexHit> indexHits,
            List<HitSummary> summaries,
            int snippetChars,
            out DateTime? widestFrontierUtc)
        {
            SweepInfo info = new SweepInfo();
            SweepWindowPlan windows = ResolveSweepWindows(request, staleness);
            widestFrontierUtc = windows.PerStoreBaseUtc == null || windows.PerStoreBaseUtc.Count == 0
                ? null
                : WidestWindow(DateTime.MaxValue, windows.PerStoreBaseUtc);
            info.IndexFrontierMissing = windows.FrontierMissing ? true : (bool?)null;
            info.StoresWithoutIndex = windows.StoresWithoutIndex;

            DateTime baseGapStart = windows.FallbackBaseUtc - SweepSafetyMargin;
            Dictionary<string, DateTime>? perStoreBase = null;
            if (windows.PerStoreBaseUtc != null)
            {
                perStoreBase = new Dictionary<string, DateTime>(windows.PerStoreBaseUtc.Count, StringComparer.OrdinalIgnoreCase);
                foreach (KeyValuePair<string, DateTime> entry in windows.PerStoreBaseUtc)
                {
                    perStoreBase[entry.Key] = entry.Value - SweepSafetyMargin;
                }
            }

            DateTime gapStart = baseGapStart;
            if (request.AfterUtc.HasValue && request.AfterUtc.Value > gapStart)
            {
                gapStart = request.AfterUtc.Value;
            }

            // The caller's own lower bound clamps EVERY window, not just the fallback: an
            // unclamped per-store window would sweep a store back past what the request asked
            // for, which costs time and returns rows the item filter then drops.
            Dictionary<string, DateTime>? perStoreGapStart = perStoreBase;
            bool windowsClamped = gapStart != baseGapStart;
            if (perStoreBase != null && request.AfterUtc.HasValue)
            {
                perStoreGapStart = new Dictionary<string, DateTime>(perStoreBase.Count, StringComparer.OrdinalIgnoreCase);
                foreach (KeyValuePair<string, DateTime> entry in perStoreBase)
                {
                    DateTime clamped = request.AfterUtc.Value > entry.Value ? request.AfterUtc.Value : entry.Value;
                    windowsClamped |= clamped != entry.Value;
                    perStoreGapStart[entry.Key] = clamped;
                }
            }

            // The widest window any store in scope could get, which is what "is a sweep
            // needed at all" has to be judged against: a 'before' bound that ends before the
            // NARROWEST window would still leave the widest one something to find.
            //
            // It COSTS a sweep that used to be skipped. An unscoped search bounded to mail
            // older than the index frontier but newer than the fallback window used to decide
            // "not needed" off the profile frontier alone and do no COM work at all; it now
            // sweeps, because a store with no frontier could hold matching mail inside that
            // band and nothing else would ever find it. The sweep it runs is the ORDINARY
            // narrow one - the per-store windows still apply, so a catalogued store is read
            // back minutes, not days - and only a store the index does not know reads the
            // full fallback span. That is the trade this whole change makes, in the one place
            // it is invisible from the payload.
            DateTime plannedGapStart = WidestWindow(gapStart, perStoreGapStart);
            info.GapStartUtc = plannedGapStart;
            if (FreshMerge.DecideSweepWindow(plannedGapStart, request.BeforeUtc) == FreshMerge.SweepWindowVerdict.NotNeeded)
            {
                // Nothing to sweep, and that is a COMPLETE answer, not a degraded one: the
                // requested window ends before the sweep would start, so the index already
                // covers all of it (FreshMerge.DecideSweepWindow).
                info.Performed = false;
                info.NotNeeded = true;
                info.Error = null;
                return info;
            }

            // What the sweep structurally cannot answer, decided in one pure place
            // (FreshMerge.SweepRefusalReason, T1-pinned). D47 added the attachment-only
            // case: the sweep never opens an attachment, so under an attachment-ONLY
            // filter every row it could contribute is one the filter excludes - merging
            // them was incoherent, and refusing matches the exhaustive tier, which has
            // always refused such a search outright for the same reason.
            string? refusal = FreshMerge.SweepRefusalReason(request.To != null, request.AttachmentHitsOnly);
            if (refusal != null)
            {
                info.Performed = false;
                info.Error = refusal;
                return info;
            }

            // The sweep follows the SEARCH scope (soak fix 13): a folder-scoped search
            // sweeps exactly that folder subtree (the index tier's SCOPE= is recursive,
            // so the sweep is too), anything else sweeps the arrival-path default
            // folders of the store(s) in scope. Before the fix every search swept
            // Inbox + Sent Items only, so mail a server-side rule filed elsewhere on
            // arrival was invisible until the index caught up.
            IReadOnlyList<string>? sweepFolderPath = ParseFolderSegments(request.Folder);
            string? folderKey = sweepFolderPath == null ? null : string.Join("/", sweepFolderPath);

            // The default folder set is shallow by construction (SweepFolder, not
            // SweepFolderTree), so only a folder-scoped sweep can honor the flag.
            bool sweepRecursive = folderKey != null && request.IncludeSubfolders;
            info.Scope = folderKey == null
                ? DefaultSweepScopeDescription
                : sweepRecursive ? "folder" : "folder (no subfolders)";

            // D34 sweep cache: rapid-fire iterative searches reuse one sweep for up to
            // ~10 s (keyed on the frontier-derived window base + store scope + folder
            // scope), so repeat calls run at index speed. Bodies are always fetched by
            // cacheable sweeps, so a term-less sweep can serve later termed searches
            // too. A cached all-stores sweep serves store-scoped requests via the
            // client-side store filter below (same folder set in every store); a
            // folder-scoped sweep only ever serves that same folder scope. Item-level
            // After/Before filters also apply below, so a wider cached window never
            // over-returns.
            //
            // MonotonicClock, and deliberately NOT the same clock as baseGapStart above. That
            // one is a real calendar instant because it becomes a DASL date compared against
            // mail Outlook received; this one is only ever subtracted from itself - the cache
            // TTL and the reported cache age - so a wall-clock jump would either hold a stale
            // sweep past its 10 s window or throw away a fresh one. Read once and used for
            // both the TryGet below and the Store further down, so an entry is always aged on
            // the clock it was stamped with.
            DateTime nowUtc = MonotonicClock.UtcNow;
            ComSweepResult effectiveResult;
            IReadOnlyList<ComMailBrief> sweptItems;
            if (_sweepCache.TryGet(
                    baseGapStart, request.Store, folderKey, sweepRecursive, nowUtc,
                    out SweepCache.CachedSweep? cachedSweep, perStoreBase)
                && cachedSweep != null)
            {
                info.Performed = true;
                info.Cached = true;
                info.CacheAgeSeconds = Math.Round((nowUtc - cachedSweep.FetchedAtUtc).TotalSeconds, 1);
                info.ElapsedMs = 0;
                ApplySweepCounters(info, cachedSweep.Result, request.Store);
                effectiveResult = cachedSweep.Result;
                sweptItems = cachedSweep.Result.Items;
            }
            else
            {
                Stopwatch stopwatch = Stopwatch.StartNew();
                ComSweepResult sweepResult;
                try
                {
                    sweepResult = _gateway.Run(
                    s => s.SweepFoldersNewerThan(
                        gapStart, SweepPerFolderCap, includeBodies: true, request.Store, sweepFolderPath, sweepRecursive,
                        perStoreGapStart),
                    SweepBudgetMs,
                    allowConnectFloor: true);
                }
                catch (OutlookUnavailableException ex)
                {
                    info.Performed = false;
                    info.Error = ex.Message;
                    return info;
                }
                catch (TimeoutException ex)
                {
                    // A bounded COM failure: the operation exceeded its budget and the COM
                    // host was restarted. Its message already says what timed out and what
                    // was done about it, so surface that rather than a bare type name -
                    // this text reaches the agent, and through it the user.
                    info.Performed = false;
                    info.Error = ex.Message;
                    return info;
                }
                catch (Exception ex) when (ex is not OutOfMemoryException)
                {
                    // The sweep is an enhancement over index results - any failure degrades
                    // to index-only with a content-free error (S4) instead of failing the
                    // whole search. Late-bound COM maps some HRESULTs to plain .NET
                    // exception types (e.g. E_INVALIDARG -> ArgumentException).
                    info.Performed = false;
                    info.Error = ex is System.Runtime.InteropServices.COMException com
                        ? string.Format(CultureInfo.InvariantCulture, "COMException 0x{0:X8}", com.HResult)
                        : ex.GetType().Name;
                    return info;
                }

                stopwatch.Stop();
                info.Performed = true;
                info.ElapsedMs = stopwatch.ElapsedMilliseconds;
                ApplySweepCounters(info, sweepResult, request.Store);
                effectiveResult = sweepResult;
                sweptItems = sweepResult.Items;

                // Only unclamped windows are cacheable: an After-narrowed sweep must
                // not poison wider follow-up searches. That is now a statement about EVERY
                // window - a request whose 'after' clamped one store's window took a
                // narrower sweep of that store than an unclamped request would.
                if (!windowsClamped)
                {
                    _sweepCache.Store(
                        baseGapStart, request.Store, folderKey, sweepRecursive, sweepResult, info.ElapsedMs, nowUtc,
                        perStoreBase);
                }
            }

            // What the sweep ACTUALLY looked back to, now that the stores it visited are
            // known: the planned value has to assume the fallback window applies to someone,
            // and on a fully catalogued profile it applies to no one. Reporting the plan
            // would say "swept back 7 days" over a sweep that swept back eleven minutes.
            info.GapStartUtc = WindowActuallyUsed(effectiveResult, request.Store, gapStart, perStoreGapStart)
                ?? plannedGapStart;

            // A store the sweep visited and the index catalog has never heard of. Its window
            // was the fallback, which is right, but whether that is COMPLETE depends on
            // something the catalog cannot answer: an indexed store missing from the
            // discovery sample is covered by the index tier regardless (an unscoped index
            // query has no SCOPE predicate), while a store that is genuinely not indexed is
            // reachable only as far back as the fallback window goes. So it is probed - the
            // same probe list_accounts reports as inLocalIndex - and only a NO is reported.
            NoteStoresWithoutIndex(info, effectiveResult, request.Store, perStoreBase);

            List<ComMailBrief> filtered = new List<ComMailBrief>();
            foreach (ComMailBrief item in sweptItems)
            {
                if (request.Store != null
                    && !string.Equals(item.StoreDisplayName, request.Store, StringComparison.OrdinalIgnoreCase))
                {
                    continue; // Cached all-stores sweep serving a store-scoped request.
                }

                info.ItemsSeen++;
                if (!FreshMerge.MatchesTerms(item, terms, request.SearchIn))
                {
                    continue;
                }

                if (request.From != null
                    && !(Contains(item.SenderAddress, request.From) || Contains(item.SenderName, request.From)))
                {
                    continue;
                }

                DateTime? receivedUtc = ToUtc(item.ReceivedTime);
                if (request.BeforeUtc.HasValue && (receivedUtc == null || receivedUtc.Value >= request.BeforeUtc.Value))
                {
                    continue;
                }

                if (request.AfterUtc.HasValue && (receivedUtc == null || receivedUtc.Value < request.AfterUtc.Value))
                {
                    continue;
                }

                if (request.UnreadOnly == true && item.IsRead != false)
                {
                    continue;
                }

                if (request.HasAttachments.HasValue && item.HasAttachments != request.HasAttachments.Value)
                {
                    continue;
                }

                filtered.Add(item);
            }

            IReadOnlyList<ComMailBrief> freshOnly = FreshMerge.SelectFreshOnly(
                filtered, indexHits, DedupeToleranceSeconds, out int duplicates);
            info.Duplicates = duplicates;
            foreach (ComMailBrief item in freshOnly)
            {
                summaries.Add(RegisterLiveHit(item, snippetChars));
            }

            return info;
        }

        /// <summary>
        /// The EARLIEST of a set of per-store window starts and the fallback every unnamed
        /// store gets - i.e. the widest window the sweep opens. Pure, and public so T1 pins
        /// the direction: taking the latest instead would report a sweep as covering less
        /// than it did, which is the safe-looking error that hides a real one.
        /// </summary>
        public static DateTime WidestWindow(DateTime fallbackUtc, IReadOnlyDictionary<string, DateTime>? perStoreUtc)
        {
            DateTime widest = fallbackUtc;
            if (perStoreUtc != null)
            {
                foreach (KeyValuePair<string, DateTime> entry in perStoreUtc)
                {
                    if (entry.Value < widest)
                    {
                        widest = entry.Value;
                    }
                }
            }

            return widest;
        }

        /// <summary>
        /// The widest window over the stores the sweep ACTUALLY visited, or null when it
        /// visited none it can name (a folder-scoped sweep of an unnameable store, or a sweep
        /// that reached nothing). Pure over the COM result, so T1 pins it without a mailbox.
        /// </summary>
        public static DateTime? WindowActuallyUsed(
            ComSweepResult result,
            string? requestedStore,
            DateTime fallbackUtc,
            IReadOnlyDictionary<string, DateTime>? perStoreUtc)
        {
            if (result == null)
            {
                throw new ArgumentNullException(nameof(result));
            }

            DateTime? widest = null;
            foreach (ComStoreSweepCounters counters in result.PerStore)
            {
                if (requestedStore != null
                    && !string.Equals(counters.StoreDisplayName, requestedStore, StringComparison.OrdinalIgnoreCase))
                {
                    continue; // A cached all-stores sweep serving a store-scoped request.
                }

                DateTime window = perStoreUtc != null
                    && perStoreUtc.TryGetValue(counters.StoreDisplayName, out DateTime since)
                    ? since
                    : fallbackUtc;
                if (widest == null || window < widest.Value)
                {
                    widest = window;
                }
            }

            return widest;
        }

        /// <summary>
        /// Names the stores this sweep covered that the index holds nothing for, so an answer
        /// resting on the sweep alone says which store it is resting on. Only ever ADDS to
        /// what the frontier probes already found.
        /// <para>
        /// Bounded twice over. The verdict per store is cached for
        /// <see cref="StoreDetailsCacheTtl"/>, so the probes run once per store per five
        /// minutes rather than once per search; and the whole pass runs under
        /// <see cref="StoreIndexProbeBudgetMs"/>, after which unprobed stores are left
        /// unreported rather than guessed at. Silence here means "not established", never
        /// "indexed" - a store that could not be probed keeps the wide fallback window it was
        /// already swept with.
        /// </para>
        /// </summary>
        private void NoteStoresWithoutIndex(
            SweepInfo info,
            ComSweepResult result,
            string? requestedStore,
            IReadOnlyDictionary<string, DateTime>? perStoreBase)
        {
            if (requestedStore != null)
            {
                return; // Scoped: the search's own frontier probe already settled it.
            }

            Stopwatch clock = Stopwatch.StartNew();
            List<string>? missing = null;
            foreach (ComStoreSweepCounters counters in result.PerStore)
            {
                if (perStoreBase != null && perStoreBase.ContainsKey(counters.StoreDisplayName))
                {
                    continue; // Has a frontier of its own, so the index knows it.
                }

                if (clock.ElapsedMilliseconds > StoreIndexProbeBudgetMs)
                {
                    break;
                }

                if (StoreHasIndexRows(counters.StoreDisplayName) == false)
                {
                    (missing ??= new List<string>()).Add(counters.StoreDisplayName);
                }
            }

            if (missing == null)
            {
                return;
            }

            info.IndexFrontierMissing = true;
            if (info.StoresWithoutIndex == null)
            {
                info.StoresWithoutIndex = missing;
                return;
            }

            List<string> combined = new List<string>(info.StoresWithoutIndex);
            foreach (string store in missing)
            {
                if (!combined.Contains(store, StringComparer.OrdinalIgnoreCase))
                {
                    combined.Add(store);
                }
            }

            info.StoresWithoutIndex = combined;
        }

        /// <summary>
        /// Whether the local index holds anything for a store, by display name, cached for
        /// <see cref="StoreDetailsCacheTtl"/>. Null when it could not be established.
        /// <para>
        /// Both shapes of the probe, in cost order: the non-delegate one (catalog by name,
        /// then targeted discovery for an '@'-named store the discovery sample missed), then
        /// the delegate one (a delegate mailbox is indexed under its OWNER's <c>/1/name</c>
        /// subtree, so it never appears in the catalog under its own name and a
        /// name-comparison alone would report every shared mailbox on the profile as
        /// unindexed). A YES from either is a yes.
        /// </para>
        /// </summary>
        private bool? StoreHasIndexRows(string displayName, int? commandTimeoutSeconds = null)
        {
            lock (_catalogLock)
            {
                if (_storeIndexPresence.TryGetValue(displayName, out (bool Present, DateTime AtUtc) known)
                    && MonotonicClock.UtcNow - known.AtUtc <= StoreDetailsCacheTtl)
                {
                    return known.Present;
                }
            }

            bool present;
            try
            {
                present = ProbeStoreInIndex(displayName, isDelegate: false, commandTimeoutSeconds)
                    || ProbeStoreInIndex(displayName, isDelegate: true, commandTimeoutSeconds);
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                return null; // Index unreachable - never answered from a probe that did not run.
            }

            lock (_catalogLock)
            {
                _storeIndexPresence[displayName] = (present, MonotonicClock.UtcNow);
            }

            return present;
        }

        // ------------------------------------------------------------------ exhaustive (Phase 3, D19)

        /// <summary>
        /// exhaustive:true - folder/date-bounded COM scan that bypasses the index
        /// entirely (ci_phrasematch DASL when Store.IsInstantSearchEnabled, LIKE
        /// fallback). Bounding rules: store is required, plus a folder or an 'after'
        /// date - an unbounded scan of a multi-GB store would be the multi-minute
        /// anti-pattern this project exists to avoid (v3.MD section 0.6 Phase 3).
        /// </summary>
        private SearchOutcome RunExhaustive(SearchRequest request, IReadOnlyList<string> terms, int top)
        {
            if (string.IsNullOrWhiteSpace(request.Store))
            {
                throw new ArgumentException(
                    "An exhaustive search requires 'store' (a display name from list_accounts) - it scans Outlook folders directly instead of the index.",
                    nameof(request));
            }

            if (request.Folder == null && !request.AfterUtc.HasValue)
            {
                throw new ArgumentException(
                    "An exhaustive search requires a bound: pass 'folder' (scan one folder) and/or 'after' (date-bounded store scan). Unbounded store scans take minutes - use a normal (indexed) search for those.",
                    nameof(request));
            }

            if (request.To != null)
            {
                throw new ArgumentException(
                    "'to' filtering is not supported in an exhaustive search (scanned items carry no recipient list). Use a normal (indexed) search or filter after read.",
                    nameof(request));
            }

            if (request.AttachmentHitsOnly)
            {
                throw new ArgumentException(
                    "Attachment-content matching requires the index; an exhaustive search scans mail subject/body only.",
                    nameof(request));
            }

            IReadOnlyList<string>? folderSegments = ParseFolderSegments(request.Folder);

            Stopwatch stopwatch = Stopwatch.StartNew();
            ComExhaustiveResult scan = _gateway.Run(s => s.ExhaustiveScan(
                request.Store!,
                folderSegments,
                terms,
                request.AfterUtc,
                request.BeforeUtc,
                maxItems: top,
                timeBudgetMs: ExhaustiveTimeBudgetMs,
                searchIn: request.SearchIn,
                includeSubfolders: request.IncludeSubfolders));
            stopwatch.Stop();

            List<HitSummary> summaries = new List<HitSummary>();
            foreach (ComMailBrief item in scan.Items)
            {
                if (request.From != null
                    && !(Contains(item.SenderAddress, request.From) || Contains(item.SenderName, request.From)))
                {
                    continue;
                }

                if (request.UnreadOnly == true && item.IsRead != false)
                {
                    continue;
                }

                if (request.HasAttachments.HasValue && item.HasAttachments != request.HasAttachments.Value)
                {
                    continue;
                }

                summaries.Add(RegisterLiveHit(item, snippetChars: 0, source: "exhaustive"));
            }

            summaries.Sort((a, b) => DateTime.Compare(b.ReceivedUtc ?? DateTime.MinValue, a.ReceivedUtc ?? DateTime.MinValue));
            if (summaries.Count > top)
            {
                summaries.RemoveRange(top, summaries.Count - top);
            }

            List<string> advice = new List<string>();
            if (!scan.InstantSearchEnabled || scan.Engine.IndexOf("like", StringComparison.Ordinal) >= 0)
            {
                advice.Add("Term matching used LIKE (substring semantics" + (scan.InstantSearchEnabled
                    ? "; ci_phrasematch was rejected here" : "; Instant Search is disabled for this store") + ") - slower and broader than index word matching.");
            }

            if (scan.Truncated)
            {
                advice.Add("Result cap (" + top.ToString(CultureInfo.InvariantCulture)
                    + ") stopped the scan - results may be incomplete. Narrow the folder/date bounds or raise top.");
            }

            if (scan.TimedOut)
            {
                advice.Add("The " + (ExhaustiveTimeBudgetMs / 1000).ToString(CultureInfo.InvariantCulture)
                    + " s time budget stopped the scan after " + scan.FoldersScanned.ToString(CultureInfo.InvariantCulture)
                    + " folder(s) - results are partial. Narrow the folder/date bounds, or pass include_subfolders:false "
                    + "to scan just the named folder.");
            }

            if (scan.FoldersSkipped > 0)
            {
                advice.Add("The scan SKIPPED " + scan.FoldersSkipped.ToString(CultureInfo.InvariantCulture)
                    + " folder(s) Outlook would not filter or enumerate (of "
                    + (scan.FoldersScanned + scan.FoldersSkipped).ToString(CultureInfo.InvariantCulture)
                    + " reached) - mail in them is NOT covered by these results.");
            }

            AddTopClampAdvice(advice, request.Top, top);

            // Staleness is best-effort context here: exhaustive works even when the
            // SystemIndex is unreachable (that is one of its jobs).
            DateTime? newestIndexed = null;
            double? ageMinutes = null;
            try
            {
                IndexStalenessReport staleness = _index.Value.GetStaleness(commandTimeoutSeconds: SearchIndexTimeoutSeconds);
                newestIndexed = staleness.NewestIndexedReceivedUtc;
                ageMinutes = staleness.Age?.TotalMinutes;
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                advice.Add("SystemIndex is unreachable (" + ex.GetType().Name + ") - exhaustive results are unaffected (COM-only path).");
            }

            return new SearchOutcome
            {
                Hits = summaries,
                Truncated = scan.Truncated,
                IndexElapsedMs = 0,
                Sweep = null,
                Exhaustive = new ExhaustiveInfo
                {
                    Engine = scan.Engine,
                    InstantSearchEnabled = scan.InstantSearchEnabled,
                    FoldersScanned = scan.FoldersScanned,
                    FoldersSkipped = scan.FoldersSkipped,
                    Truncated = scan.Truncated,
                    TimedOut = scan.TimedOut,
                    ElapsedMs = stopwatch.ElapsedMilliseconds,
                },
                Scope = request.Folder == null ? null : new SearchScopeInfo
                {
                    Folder = request.Folder,
                    IncludeSubfolders = request.IncludeSubfolders,
                    Shape = request.IncludeSubfolders ? "folder" : "folder_only",
                },
                Staleness = new StalenessInfo
                {
                    NewestIndexedUtc = newestIndexed,
                    AgeMinutes = ageMinutes,
                    OutlookRunning = ComGateway.IsOutlookRunning(),
                },
                Advice = advice.Count > 0 ? advice : null,
            };
        }

        // ------------------------------------------------------------------ read

        /// <summary>
        /// Reads one item by hit id (from search/thread) or by a REAL EntryID hex
        /// string. Index hits are located lazily (HitLocator) and the located EntryID is
        /// cached for the rest of the process lifetime.
        /// </summary>
        public ReadOutcome Read(
            string id,
            int maxBodyChars = BodyCharsDefault,
            bool includeHeaders = false,
            int maxHeaderChars = HeaderCharsDefault,
            int bodyOffset = 0,
            bool includeHtml = false,
            int maxHtmlChars = HtmlCharsDefault)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                throw new ArgumentException("id is required.", nameof(id));
            }

            maxBodyChars = Clamp(maxBodyChars, 0, BodyCharsCap);
            maxHtmlChars = Clamp(maxHtmlChars, 0, BodyCharsCap);
            maxHeaderChars = Clamp(maxHeaderChars, HeaderCharsMin, HeaderCharsCap);
            if (bodyOffset < 0)
            {
                bodyOffset = 0;
            }

            (string entryId, string? storeId, string? locatedVia, long locateMs, string? hitId) = ResolveToEntryId(id);

            // True body paging (D37): a continuation read (body_offset > 0) is served
            // from the cached one-time extraction - the body is NOT re-transferred over
            // COM. An offset-0 read always extracts fresh and refreshes the cache, so
            // plain re-reads keep their always-fresh semantics.
            string cachedBody = string.Empty;
            string cachedOrigin = "none";
            bool haveCachedBody = bodyOffset > 0 && _bodies.TryGet(entryId, out cachedBody, out cachedOrigin);

            ComItemDetail detail = _gateway.Run(s =>
            {
                ComItemDetail? d = s.TryReadItem(entryId, storeId, includeHeaders, includeBody: !haveCachedBody, out string? error, includeHtml);
                if (d == null && storeId == null)
                {
                    // Direct EntryID without a known store: retry across stores.
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        d = s.TryReadItem(entryId, store.StoreId, includeHeaders, includeBody: !haveCachedBody, out error, includeHtml);
                        if (d != null)
                        {
                            break;
                        }
                    }
                }

                return d ?? throw new InvalidOperationException("Item could not be opened (" + (error ?? "unknown") + ").");
            });

            string fullBody;
            string bodyOrigin;
            if (haveCachedBody)
            {
                fullBody = cachedBody;
                bodyOrigin = cachedOrigin;
            }
            else
            {
                fullBody = detail.Body;
                bodyOrigin = detail.BodyOrigin;
                _bodies.Put(detail.EntryId, fullBody, bodyOrigin);
                if (!string.Equals(detail.EntryId, entryId, StringComparison.OrdinalIgnoreCase))
                {
                    _bodies.Put(entryId, fullBody, bodyOrigin);
                }
            }

            (int windowStart, string window, bool moreBeyondWindow) = ComputeBodyWindow(fullBody, bodyOffset, maxBodyChars);

            string? headers = detail.Headers;
            bool? headersTruncated = null;
            if (headers != null)
            {
                headersTruncated = headers.Length > maxHeaderChars;
                if (headersTruncated.Value)
                {
                    headers = headers.Substring(0, maxHeaderChars);
                }
            }

            IReadOnlyList<RecipientView> recipients = CapRecipients(detail.Recipients, out int recipientTotal, out bool recipientsTruncated);
            IReadOnlyList<AttachmentView> attachments = CapAttachments(detail.Attachments, out int attachmentTotal, out bool attachmentsTruncated);

            // The raw HTML is bulky, so it obeys the payload discipline of section 12: one
            // window from the START (no paging - the point is to inspect structure), with
            // the true total size and a truncation flag always reported. Its budget is its
            // own (HtmlCharsDefault), see that constant for why.
            string? bodyHtml = null;
            long? bodyHtmlTotalChars = null;
            bool? bodyHtmlTruncated = null;
            if (includeHtml)
            {
                string html = detail.HtmlBody ?? string.Empty;
                bodyHtmlTotalChars = html.Length;
                bodyHtmlTruncated = html.Length > maxHtmlChars;
                bodyHtml = html.Length > maxHtmlChars ? html.Substring(0, maxHtmlChars) : html;
            }

            return new ReadOutcome
            {
                Id = hitId,
                EntryId = detail.EntryId,
                Store = detail.StoreDisplayName,
                Folder = detail.FolderPath,
                Subject = detail.Subject,
                FromName = detail.SenderName,
                FromAddress = detail.SenderAddress,
                ReceivedUtc = ToUtc(detail.ReceivedTime),
                SentUtc = ToUtc(detail.SentTime),
                Recipients = recipients,
                RecipientsTruncated = recipientsTruncated ? true : (bool?)null,
                RecipientsTotal = recipientsTruncated ? recipientTotal : (int?)null,
                Body = window,
                BodyOffset = windowStart > 0 || bodyOffset > 0 ? Math.Max(windowStart, 0) : (int?)null,
                BodyTotalChars = fullBody.Length,
                BodyTruncated = moreBeyondWindow,
                BodyOrigin = bodyOrigin,
                BodyHtml = bodyHtml,
                BodyHtmlTotalChars = bodyHtmlTotalChars,
                BodyHtmlTruncated = bodyHtmlTruncated,
                SizeBytes = detail.SizeBytes,
                IsRead = detail.IsRead,
                ConversationId = detail.ConversationId,
                InternetMessageId = detail.InternetMessageId,
                Headers = headers,
                HeadersTruncated = headersTruncated,
                Attachments = attachments,
                AttachmentsTruncated = attachmentsTruncated ? true : (bool?)null,
                AttachmentsTotal = attachmentsTruncated ? attachmentTotal : (int?)null,
                LocatedVia = locatedVia,
                LocateMs = locateMs > 0 ? locateMs : (long?)null,
            };
        }

        // ------------------------------------------------------------------ save_attachment

        /// <summary>Saves one attachment of a hit/EntryID to disk and returns the absolute path.</summary>
        public SaveAttachmentOutcome SaveAttachment(string id, int attachmentIndex, string? targetDirectory = null)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                throw new ArgumentException("id is required.", nameof(id));
            }

            if (attachmentIndex < 1)
            {
                throw new ArgumentException("attachment_index is 1-based.", nameof(attachmentIndex));
            }

            string directory = targetDirectory ?? DefaultAttachmentDirectory;
            if (!Path.IsPathRooted(directory))
            {
                throw new ArgumentException("target_dir must be an absolute path.", nameof(targetDirectory));
            }

            (string entryId, string? storeId, string? _, long _, string? hitId) = ResolveToEntryId(id);
            (string path, long size) = _gateway.Run(s =>
            {
                string? saved = s.TrySaveAttachment(entryId, storeId, attachmentIndex, directory, out long sizeBytes, out string? error);
                if (saved == null && storeId == null)
                {
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        saved = s.TrySaveAttachment(entryId, store.StoreId, attachmentIndex, directory, out sizeBytes, out error);
                        if (saved != null)
                        {
                            break;
                        }
                    }
                }

                if (saved == null)
                {
                    throw new InvalidOperationException("Attachment could not be saved (" + (error ?? "unknown") + ").");
                }

                return (saved, sizeBytes);
            });

            // Write-op audit is load-bearing from Phase 4: a failure surfaces (with the
            // saved path preserved in the message) instead of being swallowed.
            try
            {
                Audit.AuditLog.Append(
                    "save_attachment",
                    ("entryId", entryId),
                    ("path", path),
                    ("bytes", size.ToString(CultureInfo.InvariantCulture)));
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException(
                    "Attachment was saved to '" + path + "' but the audit line could not be written: " + ex.Message, ex);
            }

            return new SaveAttachmentOutcome
            {
                Id = hitId,
                EntryId = entryId,
                AttachmentIndex = attachmentIndex,
                FileName = Path.GetFileName(path),
                SavedPath = path,
                SizeBytes = size,
            };
        }

        // ------------------------------------------------------------------ thread

        /// <summary>
        /// Resolves a conversation: index ConversationID query first (scoped when the
        /// store is known), COM Conversation walk as fallback (v3.MD section 0.6 Phase 2).
        /// </summary>
        public ThreadOutcome Thread(string? conversationId, string? id, string? store, int top = ThreadTopDefault)
        {
            top = Clamp(top, 1, ThreadTopCap);
            if (conversationId == null && id == null)
            {
                throw new ArgumentException("Provide conversation_id (from a hit) or id (a hit id / EntryID).");
            }

            Stopwatch stopwatch = Stopwatch.StartNew();

            // Derive the conversation id from the referenced hit when only id was given.
            string? effectiveStore = store;
            if (conversationId == null && id != null && _hits.TryGetValue(id, out CachedHit? referenced))
            {
                conversationId = referenced.IndexHit?.ConversationId;
                effectiveStore ??= referenced.IndexHit != null
                    ? FreshMerge.ResolveHitStore(referenced.IndexHit)
                    : referenced.Live?.StoreDisplayName;
            }

            if (conversationId != null)
            {
                string? scope = null;
                if (effectiveStore != null)
                {
                    try
                    {
                        scope = ResolveScope(effectiveStore, null);
                    }
                    catch (ArgumentException)
                    {
                        scope = null;
                    }
                }

                IndexSearchResult result = _index.Value.Search(
                    new IndexQuery
                    {
                        Scope = scope,
                        Kinds = KindFilter.EmailOnly,
                        ConversationIdEquals = conversationId,
                        Top = top + 1, // Over-fetch by one: definite has-more flag.
                    },
                    SearchIndexTimeoutSeconds);
                if (result.Hits.Count > 0)
                {
                    bool indexTruncated = result.Hits.Count > top;
                    List<HitSummary> hits = result.Hits
                        .Take(top)
                        .Select(h => RegisterIndexHit(h, snippetChars: SnippetCharsDefault))
                        .OrderBy(h => h.ReceivedUtc ?? DateTime.MinValue)
                        .ToList();
                    stopwatch.Stop();
                    return new ThreadOutcome
                    {
                        ConversationId = conversationId,
                        Source = "index",
                        Hits = hits,
                        Truncated = indexTruncated,
                        ElapsedMs = stopwatch.ElapsedMilliseconds,
                    };
                }
            }

            if (id == null)
            {
                stopwatch.Stop();
                return new ThreadOutcome
                {
                    ConversationId = conversationId,
                    Source = "index",
                    Hits = Array.Empty<HitSummary>(),
                    ElapsedMs = stopwatch.ElapsedMilliseconds,
                };
            }

            // COM fallback: walk the Outlook Conversation of the referenced item
            // (over-fetch by one, same has-more contract as the index path).
            (string entryId, string? storeId, string? _, long _, string? _) = ResolveToEntryId(id);
            IReadOnlyList<ComMailBrief> briefs = _gateway.Run(s =>
            {
                IReadOnlyList<ComMailBrief>? items = s.TryGetConversationItems(entryId, storeId, top + 1, out string? error);
                return items ?? throw new InvalidOperationException("Conversation walk failed (" + (error ?? "unknown") + ").");
            });

            bool comTruncated = briefs.Count > top;
            List<HitSummary> comHits = briefs
                .Take(top)
                .Select(b => RegisterLiveHit(b, snippetChars: 0, source: "com"))
                .ToList();
            stopwatch.Stop();
            return new ThreadOutcome
            {
                ConversationId = conversationId,
                Source = "com",
                Hits = comHits,
                Truncated = comTruncated,
                ElapsedMs = stopwatch.ElapsedMilliseconds,
            };
        }

        // ------------------------------------------------------------------ show-me (Phase 3, v3.MD L3)

        /// <summary>
        /// Opens a mail in a visible Outlook Inspector window (MailItem.Display) so the
        /// user can see it. Accepts a hit id or a raw EntryID like read does.
        /// </summary>
        public OpenInOutlookOutcome OpenInOutlook(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                throw new ArgumentException("id is required.", nameof(id));
            }

            (string entryId, string? storeId, string? _, long _, string? hitId) = ResolveToEntryId(id);
            ComOpenResult displayed = _gateway.Run(s =>
            {
                ComOpenResult? d = s.TryDisplayItem(entryId, storeId, out string? error);
                if (d == null && storeId == null)
                {
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        d = s.TryDisplayItem(entryId, store.StoreId, out error);
                        if (d != null)
                        {
                            break;
                        }
                    }
                }

                return d ?? throw new InvalidOperationException("Item could not be displayed (" + (error ?? "unknown") + ").");
            });

            // open_in_outlook is a UI action, not a data write - audit stays best-effort.
            try
            {
                Audit.AuditLog.Append("open_in_outlook", ("entryId", displayed.EntryId));
            }
            catch (InvalidOperationException)
            {
            }

            return new OpenInOutlookOutcome
            {
                Id = hitId,
                EntryId = displayed.EntryId,
                Subject = displayed.Subject,
                Displayed = true,
            };
        }

        /// <summary>
        /// Navigates the Outlook window to a folder (ActiveExplorer().CurrentFolder).
        /// Omitting the folder goes to the store's Inbox (root when it has none). Creates
        /// and shows an Explorer when Outlook runs headless.
        /// </summary>
        public GotoFolderOutcome GotoFolder(string store, string? folder = null)
        {
            if (string.IsNullOrWhiteSpace(store))
            {
                throw new ArgumentException("store is required (a display name from list_accounts).", nameof(store));
            }

            IReadOnlyList<string>? segments = ParseFolderSegments(folder);
            ComExplorerState state = _gateway.Run(s =>
            {
                ComExplorerState? result = s.TryGotoFolder(store, segments, out string? error);
                return result ?? throw new InvalidOperationException(BuildNavigationError(error, store, folder));
            });

            return new GotoFolderOutcome
            {
                Store = store,
                Folder = folder,
                ExplorerFolderPath = state.CurrentFolderPath,
                ExplorerCaption = state.Caption,
                Displayed = true,
            };
        }

        /// <summary>
        /// Drives Outlook's real search UI (Explorer.Search) so the user sees the result
        /// list. Optional store/folder navigate the window there first, which is what
        /// the current_folder/subfolders scopes apply to. When the effective registry
        /// state says the UI search backend is server-assisted (D22/D35), the outcome
        /// carries advice that the displayed list may diverge from agent search.
        /// </summary>
        public ShowSearchResultsOutcome ShowSearchResults(string query, string scope = "current_folder", string? store = null, string? folder = null)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                throw new ArgumentException("query is required.", nameof(query));
            }

            if (query.Length > ShowSearchQueryCharsCap)
            {
                throw new ArgumentException(
                    "query is too long for the Outlook search box (max "
                    + ShowSearchQueryCharsCap.ToString(CultureInfo.InvariantCulture) + " chars).",
                    nameof(query));
            }

            foreach (char c in query)
            {
                if (char.IsControl(c))
                {
                    throw new ArgumentException("query must not contain control characters.", nameof(query));
                }
            }

            if (folder != null && store == null)
            {
                throw new ArgumentException("folder requires store.", nameof(folder));
            }

            int olScope = MapSearchScope(scope);
            IReadOnlyList<string>? segments = ParseFolderSegments(folder);
            ComExplorerState state = _gateway.Run(s =>
            {
                ComExplorerState? result = s.TryShowSearchResults(query, olScope, store, segments, out string? error);
                return result ?? throw new InvalidOperationException(BuildNavigationError(error, store, folder));
            });

            // Effective registry state, re-read per call (policy hive authoritative):
            // with server-assisted UI search active, what the user now SEES is capped +
            // ranked by Exchange and can silently diverge from agent search (D22/D35).
            List<string>? advice = null;
            if (HealthReporting.ReadUiSearchBackendFromRegistry() == HealthReporting.UiSearchBackendServerAssisted)
            {
                advice = new List<string> { ServerAssistedUiSearchAdvice };
            }

            return new ShowSearchResultsOutcome
            {
                Query = query,
                Scope = NormalizeScopeName(scope),
                ExplorerFolderPath = state.CurrentFolderPath,
                ExplorerCaption = state.Caption,
                Displayed = true,
                Advice = advice,
            };
        }

        /// <summary>
        /// Maps the tool-facing scope name to the OlSearchScope enum value
        /// (feature-tested live in Phase 3: all four values accepted on this Outlook
        /// build - see v3.MD section 0.8 Phase-3 facts).
        /// </summary>
        public static int MapSearchScope(string scope)
        {
            switch (NormalizeScopeName(scope))
            {
                case "current_folder":
                    return 0; // olSearchScopeCurrentFolder
                case "all_folders":
                    return 1; // olSearchScopeAllFolders (current store's mail folders)
                case "all_outlook":
                    return 2; // olSearchScopeAllOutlookItems (every store)
                case "subfolders":
                    return 3; // olSearchScopeSubfolders (current folder + children)
                default:
                    throw new ArgumentException(
                        "scope must be one of current_folder | subfolders | all_folders | all_outlook.", nameof(scope));
            }
        }

        private static string NormalizeScopeName(string scope)
        {
            return (scope ?? string.Empty).Trim().ToLowerInvariant();
        }

        private static string BuildNavigationError(string? error, string? store, string? folder)
        {
            if (error == "StoreNotFound")
            {
                return "Store '" + store + "' was not found in Outlook. Use list_accounts for store display names.";
            }

            if (error == "FolderNotFound")
            {
                return "Folder '" + folder + "' was not found in store '" + store + "'. Use list_folders for store-relative paths.";
            }

            return "Outlook could not show the requested view (" + (error ?? "unknown") + ").";
        }

        /// <summary>
        /// The folders the sweep covered IN SCOPE: all of them, or only the requested
        /// store's when a cached all-stores sweep is serving a store-scoped request.
        /// Never null - the list cap is applied by the caller, so "too long to carry" and
        /// "nothing in scope" stay distinguishable.
        /// </summary>
        private static IReadOnlyList<string> SweptFoldersInScope(ComSweepResult result, string? store)
        {
            if (store == null)
            {
                return result.SweptFolders;
            }

            List<string> folders = new List<string>(result.SweptFolders.Count);
            foreach (string entry in result.SweptFolders)
            {
                int separator = entry.IndexOf('/');
                if (separator >= 0
                    && string.Equals(entry.Substring(0, separator), store, StringComparison.OrdinalIgnoreCase))
                {
                    folders.Add(entry);
                }
            }

            return folders;
        }

        /// <summary>
        /// Copies the sweep's coverage counters onto the response block, including the
        /// ones added by soak fix 15 (failed folders, per-folder item truncation, folder
        /// cap, and whether the folder LIST was dropped by its own cap). Every one of
        /// these was previously either absent or indistinguishable from success.
        /// <para>
        /// EVERYTHING HERE IS SCOPED TO <paramref name="store"/>. The sweep result may be
        /// broader than the request - a cached all-stores sweep serves store-scoped
        /// searches (SweepCache), which is safe for the DATA because a superset contains
        /// what was asked - but the counters describe COVERAGE, and coverage of another
        /// store is no answer about this one. The lists were narrowed here from the start;
        /// the counters were not, and once they began driving the coverage codes a search
        /// scoped to store A could report <c>degraded: true</c> because store B had an
        /// unreadable folder. They are attributed per store in the COM layer
        /// (<see cref="ComSweepResult.PerStore"/>) and picked out here.
        /// </para>
        /// <para>
        /// <see cref="SweepInfo.FoldersAbsent"/> travels with them but is the one counter
        /// that is NOT a shortfall: it explains why a store contributed three folders
        /// instead of four without claiming anything was lost, so it is carried only when
        /// non-zero. It has to be attributed as carefully as the rest, because
        /// <see cref="FreshMerge.DescribeCoverageGaps"/> reads it to tell "swept nothing
        /// because there was nothing to sweep" (a store with none of the four default
        /// folders - a PST, an archive-only store) apart from "swept nothing because it all
        /// failed". Lend one store's absence to another and the wrong one goes quiet.
        /// </para>
        /// <para>
        /// Public and pure so T1 can exercise the store attribution over a hand-built
        /// sweep result: reaching these states for real needs a multi-store profile with a
        /// folder that will not open, which no CI runner has.
        /// </para>
        /// </summary>
        public static void ApplySweepCounters(SweepInfo info, ComSweepResult result, string? store)
        {
            if (info == null)
            {
                throw new ArgumentNullException(nameof(info));
            }

            if (result == null)
            {
                throw new ArgumentNullException(nameof(result));
            }

            // A store-scoped request reads that store's entry - and a MISSING entry means
            // the sweep never reached the store, which is zero coverage, not the whole
            // sweep's. Falling back to the totals there would resurrect the defect in the
            // one case where it matters most.
            ComStoreSweepCounters? scoped = store == null ? null : FindStoreCounters(result, store);
            bool perStore = store != null && result.PerStore.Count > 0;

            info.FoldersSwept = perStore ? scoped?.FoldersSwept ?? 0 : result.FoldersSwept;
            info.FoldersSkipped = perStore ? scoped?.FoldersSkipped ?? 0 : result.FoldersSkipped;
            info.FoldersFailed = perStore ? scoped?.FoldersFailed ?? 0 : result.FoldersFailed;
            int absent = perStore ? scoped?.FoldersAbsent ?? 0 : result.FoldersAbsent;
            info.FoldersAbsent = absent > 0 ? absent : (int?)null;

            IReadOnlyList<string> inScope = SweptFoldersInScope(result, store);
            info.Folders = inScope.Count == 0 || inScope.Count > SweptFolderListCap ? null : inScope;
            info.FolderListOmitted = inScope.Count > SweptFolderListCap ? true : (bool?)null;

            // Bounds of a subtree walk, which only a folder-scoped sweep performs - and a
            // folder-scoped sweep covers exactly one store, so these need no attribution.
            info.FolderCapReached = result.FolderCapReached ? true : (bool?)null;
            info.DepthLimitReached = result.DepthLimitReached ? true : (bool?)null;
            info.TimeBudgetExceeded = result.TimeBudgetExceeded ? true : (bool?)null;

            List<string> capped = new List<string>();
            foreach (string entry in result.ItemCappedFolders)
            {
                if (store != null)
                {
                    int separator = entry.IndexOf('/');
                    if (separator < 0
                        || !string.Equals(entry.Substring(0, separator), store, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                }

                capped.Add(entry);
            }

            info.ItemCappedFolders = capped.Count == 0 ? null : capped;
        }

        /// <summary>This store's counters, or null when the sweep never reached it.</summary>
        private static ComStoreSweepCounters? FindStoreCounters(ComSweepResult result, string store)
        {
            foreach (ComStoreSweepCounters entry in result.PerStore)
            {
                if (string.Equals(entry.StoreDisplayName, store, StringComparison.OrdinalIgnoreCase))
                {
                    return entry;
                }
            }

            return null;
        }

        private static IReadOnlyList<string>? ParseFolderSegments(string? folder)
        {
            if (folder == null)
            {
                return null;
            }

            string trimmed = folder.Trim().Trim('/');
            if (trimmed.Length == 0)
            {
                return null;
            }

            return trimmed.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        }

        // ------------------------------------------------------------------ drafts (Phase 4, v3.MD L4/D4)

        /// <summary>
        /// The subject-length gate for every draft path, in ONE place - and the only place
        /// the number reaches an agent's eyes, so the message quotes
        /// <see cref="SubjectCharsCap"/> rather than restating it. The cap is inclusive: a
        /// subject of exactly <see cref="SubjectCharsCap"/> characters is accepted.
        /// <para>
        /// This limit is taught by the ERROR, not by the tool schema: an over-long subject
        /// is rare, and spending description budget on it would cost every call to warn
        /// about a mistake almost none of them make. That is only a fair trade while the
        /// error is one a model can self-correct from without a second question, which is
        /// what <see cref="BuildOverlongSubjectMessage"/> owes it - the same contract as
        /// the writing-rules gate, send's confirm-token refusal and the fail-closed
        /// attachment validation, and it is pinned in T1 and over the wire in T3.
        /// </para>
        /// </summary>
        private static void RequireSubjectWithinCap(string? subject, string parameterName)
        {
            if (subject != null && subject.Length > SubjectCharsCap)
            {
                throw new ArgumentException(BuildOverlongSubjectMessage(subject.Length), parameterName);
            }
        }

        /// <summary>
        /// What an agent is told when a subject exceeds <see cref="SubjectCharsCap"/>: the
        /// limit as a number, what was actually supplied (its LENGTH, never the subject
        /// itself), the fact that nothing was written, and the retry that works.
        /// <para>
        /// Public because T1 pins those four properties against a future edit that shortens
        /// this back to "invalid subject". Every draft path shares it, so no two of them can
        /// answer the same mistake differently.
        /// </para>
        /// </summary>
        public static string BuildOverlongSubjectMessage(int suppliedLength)
        {
            return "subject is too long: " + suppliedLength.ToString(CultureInfo.InvariantCulture)
                + " characters supplied, max " + SubjectCharsCap.ToString(CultureInfo.InvariantCulture)
                + " characters. Nothing was created or changed - call again with a subject of "
                + SubjectCharsCap.ToString(CultureInfo.InvariantCulture) + " characters or fewer.";
        }

        /// <summary>
        /// Creates a new draft in <paramref name="account"/>'s Drafts folder with that
        /// account's identity and signature (v3.MD section 3 mechanics), optionally
        /// displayed for the user (D4 default). Never sends. Audit-logged (load-bearing).
        /// </summary>
        public DraftOutcome NewDraft(
            string account,
            string? to,
            string? cc,
            string? subject,
            string? body,
            bool display = true,
            string? signature = null,
            string? bcc = null,
            string? importance = null,
            bool? requestReadReceipt = null,
            string? bodyHtml = null,
            IReadOnlyList<string>? attachments = null)
        {
            if (string.IsNullOrWhiteSpace(account))
            {
                throw new ArgumentException("account is required (a sending account SMTP address from list_accounts).", nameof(account));
            }

            IReadOnlyList<string> toList = Text.HtmlBodyComposer.SplitRecipients(to);
            IReadOnlyList<string> ccList = Text.HtmlBodyComposer.SplitRecipients(cc);
            IReadOnlyList<string> bccList = Text.HtmlBodyComposer.SplitRecipients(bcc);
            if (toList.Count == 0)
            {
                throw new ArgumentException("to is required: one or more recipient addresses separated by ';' or ','.", nameof(to));
            }

            if (string.IsNullOrWhiteSpace(subject))
            {
                throw new ArgumentException("subject is required.", nameof(subject));
            }

            RequireSubjectWithinCap(subject, nameof(subject));

            ComDraftBody draftBody = ResolveDraftBody(body, bodyHtml, "the signature", out IReadOnlyList<string> htmlAdjustments);
            IReadOnlyList<DraftAttachmentFile> files = DraftAttachments.Validate(attachments);

            ComSignatureOverride? signatureOverride = ResolveSignatureOverride(signature);
            ComDraftOptions options = new(
                ccList, bccList, subjectOverride: null, ParseImportance(importance), requestReadReceipt,
                files.Select(f => f.Path).ToList());
            ComDraftCreateResult created = _gateway.Run(s =>
            {
                ComDraftCreateResult? r = s.TryCreateNewDraft(account, toList, subject!, draftBody, display, signatureOverride, options, out string? error);
                return r ?? throw new InvalidOperationException(BuildDraftError(error, account));
            });

            _draftRegistry.Register(created.Draft.EntryId);
            AuditDraft("new_draft", created, requestedAccount: account, sourceEntryId: null, draftBody, files);
            return ToDraftOutcome("new", created, hitId: null, sourceEntryId: null, draftBody, htmlAdjustments, files);
        }

        /// <summary>
        /// Re-reads a just-saved draft's attachments in a SEPARATE COM call and returns the
        /// better of the two snapshots.
        /// <para>
        /// Soak fix 21, and it fixes a defect that was reported to us as data loss: the
        /// snapshot taken INSIDE a compose call reports an attachment Outlook materialized
        /// during that composition - a signature's inline logo - as ZERO bytes, and in the
        /// HTMLBody-fallback shape does not see it at all. Nothing is wrong with the item;
        /// the size is simply not committed while the composing call still holds the item
        /// open, which is why <c>read</c> always reported it correctly and the draft tools
        /// did not. A second, plain call answers the truth, so every draft tool now echoes
        /// the same bytes the recipient will get. Failure here can never fail a draft: the
        /// original snapshot is kept.
        /// </para>
        /// </summary>
        private IReadOnlyList<Com.ComAttachmentInfo> VerifiedAttachments(
            Com.ComDraftInfo draft,
            IReadOnlyList<Com.ComAttachmentInfo> inHand)
        {
            if (!Com.AttachmentSnapshotMerge.HasUnsizedAttachment(inHand) && inHand.Count > 0)
            {
                return inHand;
            }

            try
            {
                IReadOnlyList<Com.ComAttachmentInfo> fresh =
                    _gateway.Run(s => s.SnapshotAttachmentsById(draft.EntryId, draft.StoreId));
                return Com.AttachmentSnapshotMerge.IsBetter(fresh, inHand) ? fresh : inHand;
            }
            catch (Exception)
            {
                // Reporting must never cost a caller their draft.
                return inHand;
            }
        }

        /// <summary>
        /// Picks the ONE body form a draft call supplied and prepares it for the COM layer
        /// (soak fix batch B - B1). Exactly one of <paramref name="body"/> (plain text, the
        /// unchanged default) and <paramref name="bodyHtml"/> (HTML) must be present; the
        /// HTML is put through <see cref="Text.HtmlFragmentNormalizer"/> here, PRE-COM, so a
        /// hostile or malformed fragment is repaired (or rejected) before it can touch a
        /// mailbox, and everything that changed is reported back to the caller.
        /// </summary>
        private static ComDraftBody ResolveDraftBody(string? body, string? bodyHtml, string placementHint, out IReadOnlyList<string> htmlAdjustments)
        {
            htmlAdjustments = Array.Empty<string>();
            bool hasText = !string.IsNullOrWhiteSpace(body);
            bool hasHtml = !string.IsNullOrWhiteSpace(bodyHtml);

            if (hasText && hasHtml)
            {
                throw new ArgumentException(
                    "body and body_html are mutually exclusive - supply exactly one. Use body for plain text, body_html for formatted HTML.",
                    nameof(bodyHtml));
            }

            if (!hasText && !hasHtml)
            {
                throw new ArgumentException(
                    "A body is required: supply either body (plain text) or body_html (formatted HTML). Either one is placed above "
                    + placementHint + ".",
                    nameof(body));
            }

            if (!hasHtml)
            {
                return ComDraftBody.FromText(body!);
            }

            Text.HtmlNormalizationResult normalized = Text.HtmlFragmentNormalizer.Normalize(bodyHtml);
            if (!normalized.HasVisibleContent)
            {
                throw new ArgumentException(
                    "body_html contained no usable content after normalization (only unsupported or removed markup). "
                    + "Send visible HTML, or use body for plain text.",
                    nameof(bodyHtml));
            }

            htmlAdjustments = normalized.Adjustments;
            return ComDraftBody.FromHtml(normalized.Html);
        }

        /// <summary>
        /// Creates a reply (or reply-all) draft for a hit id / EntryID via COM
        /// <c>Reply()</c>/<c>ReplyAll()</c> - threading and quoted history preserved,
        /// agent text above the quote, saved to the source store's Drafts (D4). Never sends.
        /// </summary>
        public DraftOutcome ReplyDraft(
            string id,
            string? body,
            bool replyAll = false,
            bool display = true,
            string? signature = null,
            string? cc = null,
            string? bcc = null,
            string? subject = null,
            string? importance = null,
            bool? requestReadReceipt = null,
            string? bodyHtml = null,
            IReadOnlyList<string>? attachments = null)
        {
            (string? hitId, string sourceEntryId, ComDraftCreateResult created, ComDraftBody draftBody, IReadOnlyList<string> htmlAdjustments, IReadOnlyList<DraftAttachmentFile> files) = CreateDerived(
                id,
                replyAll ? ComDerivedDraftKind.ReplyAll : ComDerivedDraftKind.Reply,
                to: null,
                body,
                display,
                signature,
                cc,
                bcc,
                subject,
                importance,
                requestReadReceipt,
                bodyHtml,
                attachments);
            string op = replyAll ? "replyall_draft" : "reply_draft";
            _draftRegistry.Register(created.Draft.EntryId);
            AuditDraft(op, created, requestedAccount: null, sourceEntryId, draftBody, files);
            return ToDraftOutcome(replyAll ? "replyall" : "reply", created, hitId, sourceEntryId, draftBody, htmlAdjustments, files);
        }

        /// <summary>
        /// Creates a forward draft for a hit id / EntryID via COM <c>Forward()</c> -
        /// quoted content and attachments preserved, agent text above the quote, saved to
        /// the source store's Drafts (D4). Never sends.
        /// </summary>
        public DraftOutcome ForwardDraft(
            string id,
            string? body,
            string? to,
            bool display = true,
            string? signature = null,
            string? cc = null,
            string? bcc = null,
            string? subject = null,
            string? importance = null,
            bool? requestReadReceipt = null,
            string? bodyHtml = null,
            IReadOnlyList<string>? attachments = null)
        {
            IReadOnlyList<string> toList = Text.HtmlBodyComposer.SplitRecipients(to);
            if (toList.Count == 0)
            {
                throw new ArgumentException("to is required for forward_draft: one or more recipient addresses separated by ';' or ','.", nameof(to));
            }

            (string? hitId, string sourceEntryId, ComDraftCreateResult created, ComDraftBody draftBody, IReadOnlyList<string> htmlAdjustments, IReadOnlyList<DraftAttachmentFile> files) = CreateDerived(
                id, ComDerivedDraftKind.Forward, toList, body, display, signature, cc, bcc, subject, importance, requestReadReceipt, bodyHtml, attachments);
            _draftRegistry.Register(created.Draft.EntryId);
            AuditDraft("forward_draft", created, requestedAccount: null, sourceEntryId, draftBody, files);
            return ToDraftOutcome("forward", created, hitId, sourceEntryId, draftBody, htmlAdjustments, files);
        }

        private (string? HitId, string SourceEntryId, ComDraftCreateResult Created, ComDraftBody Body, IReadOnlyList<string> HtmlAdjustments, IReadOnlyList<DraftAttachmentFile> Files) CreateDerived(
            string id,
            ComDerivedDraftKind kind,
            IReadOnlyList<string>? to,
            string? body,
            bool display,
            string? signature,
            string? cc,
            string? bcc,
            string? subject,
            string? importance,
            bool? requestReadReceipt,
            string? bodyHtml,
            IReadOnlyList<string>? attachments)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                throw new ArgumentException("id is required (a hit id from search/thread or a full EntryID).", nameof(id));
            }

            ComDraftBody draftBody = ResolveDraftBody(body, bodyHtml, "the quoted mail", out IReadOnlyList<string> htmlAdjustments);
            IReadOnlyList<DraftAttachmentFile> files = DraftAttachments.Validate(attachments);

            RequireSubjectWithinCap(subject, nameof(subject));

            ComSignatureOverride? signatureOverride = ResolveSignatureOverride(signature);
            ComDraftOptions options = new(
                Text.HtmlBodyComposer.SplitRecipients(cc),
                Text.HtmlBodyComposer.SplitRecipients(bcc),
                string.IsNullOrWhiteSpace(subject) ? null : subject!.Trim(),
                ParseImportance(importance),
                requestReadReceipt,
                files.Select(f => f.Path).ToList());
            (string entryId, string? storeId, string? _, long _, string? hitId) = ResolveToEntryId(id);
            IReadOnlyList<string> toList = to ?? Array.Empty<string>();
            ComDraftCreateResult created = _gateway.Run(s =>
            {
                ComDraftCreateResult? r = s.TryCreateDerivedDraft(entryId, storeId, kind, toList, draftBody, display, signatureOverride, options, out string? error);
                if (r == null && storeId == null)
                {
                    // Direct EntryID without a known store: retry across stores (same
                    // pattern as read/open_in_outlook).
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        r = s.TryCreateDerivedDraft(entryId, store.StoreId, kind, toList, draftBody, display, signatureOverride, options, out error);
                        if (r != null)
                        {
                            break;
                        }
                    }
                }

                return r ?? throw new InvalidOperationException(
                    "The source mail could not be opened or the draft could not be created (" + (error ?? "unknown")
                    + "). Re-run search - the item may have moved.");
            });

            return (hitId, entryId, created, draftBody, htmlAdjustments, files);
        }

        /// <summary>
        /// Validates and resolves the optional draft-tool signature name BEFORE any COM
        /// work: null stays null (account default), an unknown name is rejected with
        /// the available names listed (agent self-correction), a known one becomes the
        /// COM override request carrying the preferred file (.htm - Word converts and
        /// embeds natively).
        /// </summary>
        private static ComSignatureOverride? ResolveSignatureOverride(string? signature)
        {
            if (string.IsNullOrWhiteSpace(signature))
            {
                return null;
            }

            SignatureInfo? resolved = SignatureCatalog.TryResolve(signature!);
            if (resolved?.PreferredFilePath == null)
            {
                IReadOnlyList<SignatureInfo> available = SignatureCatalog.ListSignatures();
                throw new ArgumentException(
                    "signature '" + signature!.Trim() + "' was not found"
                    + (available.Count > 0
                        ? ". Available signatures: " + string.Join(", ", available.Select(s => s.Name)) + ". Use list_signatures for details."
                        : " - no signatures are installed (see list_signatures)."),
                    nameof(signature));
            }

            return new ComSignatureOverride(resolved.Name, resolved.PreferredFilePath!);
        }

        /// <summary>
        /// Validates the optional draft-tool importance BEFORE any COM work (batch A,
        /// A4): null/blank = leave Outlook's default; "low"/"normal"/"high" map to
        /// OlImportance 0/1/2; anything else is rejected with the allowed values so an
        /// agent can self-correct.
        /// </summary>
        public static int? ParseImportance(string? importance)
        {
            if (string.IsNullOrWhiteSpace(importance))
            {
                return null;
            }

            switch (importance!.Trim().ToLowerInvariant())
            {
                case "low":
                    return 0;
                case "normal":
                    return 1;
                case "high":
                    return 2;
                default:
                    throw new ArgumentException(
                        "importance must be 'low', 'normal' or 'high' (got '" + importance.Trim() + "').",
                        nameof(importance));
            }
        }

        /// <summary>OlImportance value back to the wire vocabulary.</summary>
        private static string? ImportanceName(int? importance)
        {
            return importance switch
            {
                0 => "low",
                1 => "normal",
                2 => "high",
                _ => null,
            };
        }

        private static string BuildDraftError(string? error, string account)
        {
            if (error == "AccountNotFound")
            {
                return "Account '" + account + "' was not found in the Outlook profile. Use list_accounts for the exact account SMTP addresses.";
            }

            if (error == "AccountHasNoDeliveryStore")
            {
                return "Account '" + account + "' has no delivery store; a draft cannot be filed for it.";
            }

            return "The draft could not be created (" + (error ?? "unknown") + ").";
        }

        /// <summary>
        /// Write-op audit (LIVE and load-bearing from Phase 4): the structured line is
        /// appended for every created draft; a failure surfaces with the draft's EntryID
        /// preserved in the message instead of being swallowed.
        /// </summary>
        private static void AuditDraft(
            string operation,
            ComDraftCreateResult created,
            string? requestedAccount,
            string? sourceEntryId,
            ComDraftBody body,
            IReadOnlyList<DraftAttachmentFile> attachments)
        {
            try
            {
                Audit.AuditLog.Append(
                    operation,
                    ("attachments", attachments.Count > 0 ? attachments.Count.ToString(CultureInfo.InvariantCulture) : null),
                    ("attachmentBytes", attachments.Count > 0
                        ? DraftAttachments.TotalBytes(attachments).ToString(CultureInfo.InvariantCulture)
                        : null),
                    ("entryId", created.Draft.EntryId),
                    ("store", created.Draft.StoreDisplayName),
                    ("account", created.Draft.SendUsingAccountSmtp ?? requestedAccount),
                    ("accountResolved", created.AccountResolved ? "true" : "false"),
                    ("signatureInjected", created.SignatureInjected ? "true" : "false"),
                    ("signature", created.SignatureOverrideName),
                    ("signatureApplied", created.SignatureOverrideName != null ? (created.SignatureOverrideApplied ? "true" : "false") : null),
                    ("bodyPlacement", created.BodyPlacedViaWordEditor ? "wordEditor" : "html"),
                    // D49: the compose surface is audited on EVERY draft - "promoted" is the
                    // headless-but-fully-capable case, and a degraded composition names its
                    // reason here as well as on the wire, so a lesser draft is never silent.
                    ("composeSurface", created.BodyPlacedViaWordEditor
                        ? (created.ComposeSurfacePromoted ? "wordEditorPromoted" : "wordEditor")
                        : "htmlFallback"),
                    ("composeSurfaceError", created.BodyPlacedViaWordEditor ? null : created.ComposeSurfaceError),
                    ("bodyFormat", body.FormatName),
                    ("displayed", created.Displayed ? "true" : "false"),
                    ("recipients", created.Draft.Recipients.Count.ToString(CultureInfo.InvariantCulture)),
                    ("unresolvedRecipients", created.UnresolvedRecipients.Count > 0
                        ? created.UnresolvedRecipients.Count.ToString(CultureInfo.InvariantCulture)
                        : null),
                    ("conversationTopicPreserved", created.ConversationTopicPreserved?.ToString().ToLowerInvariant()),
                    ("movedToDrafts", created.MovedToDrafts ? "true" : "false"),
                    ("initialFolder", created.InitialSaveFolderName),
                    ("sourceEntryId", sourceEntryId));
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException(
                    "The draft was created (EntryID " + created.Draft.EntryId
                    + ") but the audit line could not be written: " + ex.Message, ex);
            }
        }

        private DraftOutcome ToDraftOutcome(
            string kind,
            ComDraftCreateResult created,
            string? hitId,
            string? sourceEntryId,
            ComDraftBody body,
            IReadOnlyList<string> htmlAdjustments,
            IReadOnlyList<DraftAttachmentFile> attachments)
        {
            IReadOnlyList<RecipientView> recipients = CapRecipients(created.Draft.Recipients, out int total, out bool truncated);

            // Attachments come from the SAVED item, never from the request (the A4
            // round-trip discipline): what the agent is told is what Outlook holds.
            IReadOnlyList<AttachmentView> attachmentViews = CapAttachments(
                VerifiedAttachments(created.Draft, created.Attachments), out int _, out bool _);
            return new DraftOutcome
            {
                Attachments = attachmentViews.Count > 0 ? attachmentViews : null,
                AttachmentsTotalBytes = attachmentViews.Count > 0
                    ? attachmentViews.Sum(a => a.SizeBytes ?? 0)
                    : (long?)null,
                AttachmentsRequested = attachments.Count > 0 ? attachments.Count : (int?)null,
                Kind = kind,
                Id = hitId,
                SourceEntryId = sourceEntryId,
                EntryId = created.Draft.EntryId,
                Store = created.Draft.StoreDisplayName,
                Folder = created.Draft.ParentFolderName,
                Account = created.Draft.SendUsingAccountSmtp,
                AccountResolved = created.AccountResolved,
                Subject = created.Draft.Subject,
                SignatureInjected = created.SignatureInjected,
                Signature = created.SignatureOverrideName,
                SignatureApplied = created.SignatureOverrideName != null ? created.SignatureOverrideApplied : (bool?)null,
                SignatureError = created.SignatureOverrideName != null ? created.SignatureOverrideError : null,
                Displayed = created.Displayed,
                ConversationId = created.Draft.ConversationId,
                Recipients = recipients,
                RecipientsTruncated = truncated ? true : (bool?)null,
                RecipientsTotal = truncated ? total : (int?)null,
                UnresolvedRecipients = created.UnresolvedRecipients.Count > 0
                    ? created.UnresolvedRecipients.Take(UnresolvedRecipientsCap).ToList()
                    : null,
                ConversationTopicPreserved = created.ConversationTopicPreserved,
                Importance = ImportanceName(created.Draft.Importance) is string name && name != "normal" ? name : null,
                ReadReceiptRequested = created.Draft.ReadReceiptRequested ? true : (bool?)null,
                BodyFormat = body.FormatName,
                BodyPlacement = created.BodyPlacedViaWordEditor ? "wordEditor" : "html",
                ComposeSurfacePromoted = created.ComposeSurfacePromoted ? true : (bool?)null,
                ComposeSurfaceError = created.BodyPlacedViaWordEditor ? null : created.ComposeSurfaceError,
                ComposeSurfaceAdvice = created.BodyPlacedViaWordEditor ? null : ComposeSurfaceDegradedAdvice,
                HtmlAdjustments = htmlAdjustments.Count > 0 ? htmlAdjustments : null,
            };
        }

        /// <summary>
        /// D49: what a caller is told when a draft could only be composed through the
        /// HTMLBody fallback. It names every capability the fallback does NOT deliver,
        /// because the D48 defect was precisely that this degradation was invisible.
        /// </summary>
        internal const string ComposeSurfaceDegradedAdvice =
            "Outlook's Word compose surface could not be obtained, so this draft was composed by the HTMLBody fallback: "
            + "the body sits outside Outlook's own WordSection1 container (it does not inherit the message style), "
            + "an explicit signature override could not be applied, and a signature's linked images were not embedded. "
            + "The draft was still created and nothing was lost. Retry once - if it keeps happening, Outlook is in a state "
            + "that refuses an editor; opening any Outlook window (goto_folder) makes the full compose path available again.";

        // ------------------------------------------------------------------ update / discard drafts (D46, soak fix 19)

        /// <summary>
        /// Test/diagnostic view of the per-process registry of drafts THIS server created
        /// or last updated - the allowlist <c>discard_draft</c> is gated on (D46/C2).
        /// </summary>
        public ServerDraftRegistry DraftRegistry => _draftRegistry;

        /// <summary>
        /// update_draft (v3.MD D46/C1): revises an existing UNSENT draft in place. Only
        /// the parts supplied are touched; everything omitted is left exactly as it is.
        /// <para>
        /// RECIPIENTS ARE REPLACE, NOT APPEND - deliberately the opposite of the draft
        /// creators' cc/bcc append (batch A, A2). On creation "append" is the only sane
        /// reading because Outlook has just filled the list itself; on a revision the
        /// caller is stating the final list, and there would otherwise be no way to REMOVE
        /// a recipient. Passing the full list is therefore how you add one.
        /// </para>
        /// <para>
        /// ATTACHMENTS ARE ADD, plus an explicit <paramref name="removeAttachments"/> name
        /// list - the simplest surface that can express add, remove and replace without a
        /// mode flag: remove+add of the same name in one call IS a replace, and the
        /// removals run first so that works.
        /// </para>
        /// </summary>
        public UpdateDraftOutcome UpdateDraft(
            string id,
            string? body = null,
            string? bodyHtml = null,
            string? subject = null,
            string? to = null,
            string? cc = null,
            string? bcc = null,
            string? importance = null,
            bool? requestReadReceipt = null,
            string? signature = null,
            IReadOnlyList<string>? attachments = null,
            IReadOnlyList<string>? removeAttachments = null,
            bool display = true)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                throw new ArgumentException(
                    "id is required (the entryId a draft tool returned, or a hit id of a saved unsent draft).", nameof(id));
            }

            // A body is OPTIONAL here (unlike the creators), but the two forms stay
            // mutually exclusive.
            ComDraftBody? draftBody = null;
            IReadOnlyList<string> htmlAdjustments = Array.Empty<string>();
            if (!string.IsNullOrWhiteSpace(body) || !string.IsNullOrWhiteSpace(bodyHtml))
            {
                draftBody = ResolveDraftBody(body, bodyHtml, "the signature", out htmlAdjustments);
            }

            if (subject != null && subject.Trim().Length == 0)
            {
                throw new ArgumentException(
                    "subject must not be blank - omit it to keep the draft's current subject.", nameof(subject));
            }

            RequireSubjectWithinCap(subject, nameof(subject));

            IReadOnlyList<string>? toList = to == null ? null : Text.HtmlBodyComposer.SplitRecipients(to);
            IReadOnlyList<string>? ccList = cc == null ? null : Text.HtmlBodyComposer.SplitRecipients(cc);
            IReadOnlyList<string>? bccList = bcc == null ? null : Text.HtmlBodyComposer.SplitRecipients(bcc);
            if (toList != null && toList.Count == 0)
            {
                throw new ArgumentException(
                    "to was supplied but holds no usable address. update_draft REPLACES the To list, so an empty value would "
                    + "leave the draft with no recipient - omit to to keep the current list.",
                    nameof(to));
            }

            int? parsedImportance = ParseImportance(importance);
            ComSignatureOverride? signatureOverride = ResolveSignatureOverride(signature);
            IReadOnlyList<DraftAttachmentFile> files = DraftAttachments.Validate(attachments);
            IReadOnlyList<string> removeNames = DraftAttachments.ValidateRemoveNames(removeAttachments);

            if (draftBody == null && subject == null && toList == null && ccList == null && bccList == null
                && parsedImportance == null && requestReadReceipt == null && signatureOverride == null
                && files.Count == 0 && removeNames.Count == 0)
            {
                throw new ArgumentException(
                    "Nothing to update: supply at least one of body / body_html / subject / to / cc / bcc / importance / "
                    + "request_read_receipt / signature / attachments / remove_attachments.",
                    nameof(id));
            }

            (string entryId, string? storeId, string? _, long _, string? hitId) = ResolveToEntryId(id);

            ComDraftUpdateResult updated = _gateway.Run(s =>
            {
                string? error = null;
                ComDraftUpdateResult? r = s.TryUpdateDraft(
                    entryId, storeId, draftBody, subject?.Trim(), toList, ccList, bccList,
                    parsedImportance, requestReadReceipt, signatureOverride,
                    files.Select(f => f.Path).ToList(), removeNames, display, out error);

                if (r == null && storeId == null && error == "ItemNotFound")
                {
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        r = s.TryUpdateDraft(
                            entryId, store.StoreId, draftBody, subject?.Trim(), toList, ccList, bccList,
                            parsedImportance, requestReadReceipt, signatureOverride,
                            files.Select(f => f.Path).ToList(), removeNames, display, out error);
                        if (r != null)
                        {
                            break;
                        }
                    }
                }

                return r ?? throw BuildDraftRefusal("update_draft", error, entryId);
            });

            // EntryIDs are not stable - re-key the registry so a following discard_draft
            // still recognises the draft this call just rewrote (D46/C2).
            _draftRegistry.Replace(entryId, updated.Draft.EntryId);
            AuditUpdate(updated, draftBody, hitId);
            return ToUpdateOutcome(updated, hitId, draftBody, htmlAdjustments, removeNames);
        }

        /// <summary>
        /// discard_draft (v3.MD D46/C2, the S1 v3 amendment): SOFT-deletes a draft THIS
        /// server created or last updated - <c>Delete()</c>, which moves it to Deleted
        /// Items exactly like pressing Delete in Outlook. It is the only mail-deleting
        /// tool in the product and it is narrow by construction:
        /// <list type="bullet">
        /// <item>the item must be in the session registry (this server authored it),</item>
        /// <item>it must live in a Drafts folder,</item>
        /// <item>it must be UNSENT.</item>
        /// </list>
        /// Never <c>PermanentlyDelete</c>, never empties anything, never touches Deleted
        /// Items' existing contents. Every refusal is explicit and audited - there is no
        /// silent no-op path.
        /// </summary>
        public DiscardDraftOutcome DiscardDraft(string id)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                throw new ArgumentException(
                    "id is required (the entryId a draft tool returned for a draft created in THIS session).", nameof(id));
            }

            (string entryId, string? storeId, string? _, long _, string? hitId) = ResolveToEntryId(id);

            // THE GUARDRAIL, checked before any COM work: only drafts this server process
            // produced are reachable. Both the raw argument and the resolved EntryID count,
            // so a hit id pointing at our own draft works too.
            if (!_draftRegistry.Contains(entryId) && !_draftRegistry.Contains(id))
            {
                throw RefuseDraft(
                    "not_created_by_this_server",
                    "discard_draft",
                    entryId,
                    "This draft was not created or last updated by this server session, so it cannot be discarded. "
                    + "discard_draft exists only to clean up drafts the assistant itself just made: it can reach a draft "
                    + "returned by new_draft / reply_draft / replyall_draft / forward_draft / update_draft in THIS session, "
                    + "and nothing else - not mail you received, not anything you wrote yourself, not a sent item, and not a "
                    + "draft from an earlier session (a server restart clears the list). Delete it in Outlook instead.");
            }

            ComDraftDiscardResult discarded = _gateway.Run(s =>
            {
                string? error = null;
                ComDraftDiscardResult? r = s.TryDiscardDraft(entryId, storeId, out error);
                if (r == null && storeId == null && error == "ItemNotFound")
                {
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        r = s.TryDiscardDraft(entryId, store.StoreId, out error);
                        if (r != null)
                        {
                            break;
                        }
                    }
                }

                return r ?? throw BuildDraftRefusal("discard_draft", error, entryId);
            });

            AuditDiscard(discarded, hitId);
            _draftRegistry.Forget(entryId);
            _draftRegistry.Register(discarded.NewEntryId);
            return new DiscardDraftOutcome
            {
                Status = "discarded",
                Discarded = true,
                Id = hitId,
                EntryId = discarded.OldEntryId,
                NewEntryId = discarded.NewEntryId,
                Store = discarded.StoreDisplayName,
                FromFolder = discarded.FromFolder,
                ToFolder = discarded.ToFolder,
                Subject = discarded.Subject,
                Advice = discarded.NewEntryId != null && discarded.FromFolder != null
                    ? "Soft delete only - the draft is in " + (discarded.ToFolder ?? "Deleted Items")
                        + " and can be restored with move_mail using newEntryId and folder='" + discarded.FromFolder + "'."
                    : "Soft delete only - the draft was moved to " + (discarded.ToFolder ?? "Deleted Items")
                        + " and can be restored from there in Outlook.",
            };
        }

        /// <summary>
        /// Maps a COM-side refusal code from update/discard to the user-facing refusal.
        /// Every one of these means NOTHING was changed or deleted.
        /// </summary>
        private static Exception BuildDraftRefusal(string operation, string? comError, string entryId)
        {
            switch (comError)
            {
                case "NotAMailItem":
                    return RefuseDraft("not_a_mail_item", operation, entryId,
                        "That id is not a mail item (it may be an appointment, contact or task). "
                        + (operation == "discard_draft" ? "discard_draft" : "update_draft") + " works on mail drafts only.");
                case "AlreadySent":
                    return RefuseDraft("not_an_unsent_draft", operation, entryId,
                        "That item has already been sent, so it is no longer a draft and cannot be "
                        + (operation == "discard_draft" ? "discarded" : "revised")
                        + ". Only saved, UNSENT drafts can be. A sent mail can never be changed or deleted by this server.");
                case "NotInDraftsFolder":
                    return RefuseDraft("not_in_drafts_folder", operation, entryId,
                        "That item does not live in a Drafts folder, so it is not a draft this server may touch. "
                        + "Move it to Drafts in Outlook first if it really is an unfinished message.");
                case "NoInspector":
                case "NoWordEditor":
                    return RefuseDraft("compose_surface_unavailable", operation, entryId,
                        "Outlook would not open the draft's editor (" + comError + "), so the body could not be replaced. "
                        + "The draft is unchanged. This usually means Outlook is still starting up or is busy - retry in a "
                        + "moment, and close any compose window that already has this draft open.");
                case "SignatureFileMissing":
                    return RefuseDraft("signature_file_missing", operation, entryId,
                        "The requested signature's file is missing on disk. The draft is unchanged; see list_signatures.");
                case "ItemNotFound":
                    return new InvalidOperationException(
                        "The draft could not be opened - it may have been deleted, moved or already sent. "
                        + "Re-check with read, or re-run search for a fresh id.");
                default:
                    return RefuseDraft("com_failure", operation, entryId,
                        "The draft could not be " + (operation == "discard_draft" ? "discarded" : "updated")
                        + " (" + (comError ?? "unknown") + "). Nothing was changed. Check outlook_health and retry.");
            }
        }

        /// <summary>Audit-logs the refusal and builds the exception (nothing was changed).</summary>
        private static DraftRefusedException RefuseDraft(string reason, string operation, string? entryId, string message)
        {
            try
            {
                Audit.AuditLog.Append(
                    operation + "_refused",
                    ("entryId", entryId),
                    ("reason", reason));
            }
            catch (InvalidOperationException)
            {
                // A refusal changed nothing; an unwritable audit log must not convert it
                // into a different, more confusing error.
            }

            return new DraftRefusedException(reason, message);
        }

        private static void AuditUpdate(ComDraftUpdateResult updated, ComDraftBody? body, string? hitId)
        {
            try
            {
                Audit.AuditLog.Append(
                    "update_draft",
                    ("entryId", updated.Draft.EntryId),
                    ("hitId", hitId),
                    ("store", updated.Draft.StoreDisplayName),
                    ("folder", updated.Draft.ParentFolderName),
                    ("account", updated.Draft.SendUsingAccountSmtp),
                    ("changed", updated.ChangedFields.Count > 0 ? string.Join("+", updated.ChangedFields) : "none"),
                    ("bodyFormat", body?.FormatName),
                    ("bodyPlacement", updated.BodyReplaced ? (updated.BodyPlacedViaWordEditor ? "wordEditor" : "html") : null),
                    ("signature", updated.SignatureOverrideName),
                    ("attachmentsAdded", updated.AttachmentsAdded.Count > 0
                        ? updated.AttachmentsAdded.Count.ToString(CultureInfo.InvariantCulture)
                        : null),
                    ("attachmentsRemoved", updated.AttachmentsRemoved.Count > 0
                        ? updated.AttachmentsRemoved.Count.ToString(CultureInfo.InvariantCulture)
                        : null),
                    ("attachmentsFailed", updated.AttachmentsFailed.Count > 0
                        ? updated.AttachmentsFailed.Count.ToString(CultureInfo.InvariantCulture)
                        : null),
                    ("attachmentsTotal", updated.Attachments.Count.ToString(CultureInfo.InvariantCulture)),
                    ("recipients", updated.Draft.Recipients.Count.ToString(CultureInfo.InvariantCulture)),
                    ("unresolvedRecipients", updated.UnresolvedRecipients.Count > 0
                        ? updated.UnresolvedRecipients.Count.ToString(CultureInfo.InvariantCulture)
                        : null),
                    ("conversationTopicPreserved", updated.ConversationTopicPreserved?.ToString().ToLowerInvariant()),
                    ("inlineImagesDropped", updated.InlineImagesDropped > 0
                        ? updated.InlineImagesDropped.ToString(CultureInfo.InvariantCulture)
                        : null),
                    ("displayed", updated.Displayed ? "true" : "false"));
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException(
                    "The draft was updated (EntryID " + updated.Draft.EntryId
                    + ") but the audit line could not be written: " + ex.Message, ex);
            }
        }

        private static void AuditDiscard(ComDraftDiscardResult discarded, string? hitId)
        {
            try
            {
                Audit.AuditLog.Append(
                    "discard_draft",
                    ("entryId", discarded.OldEntryId),
                    ("newEntryId", discarded.NewEntryId),
                    ("hitId", hitId),
                    ("store", discarded.StoreDisplayName),
                    ("fromFolder", discarded.FromFolder),
                    ("toFolder", discarded.ToFolder),
                    ("mode", "soft"));
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException(
                    "The draft was discarded (EntryID " + discarded.OldEntryId
                    + ", now in " + (discarded.ToFolder ?? "Deleted Items")
                    + ") but the audit line could not be written: " + ex.Message, ex);
            }
        }

        private UpdateDraftOutcome ToUpdateOutcome(
            ComDraftUpdateResult updated,
            string? hitId,
            ComDraftBody? body,
            IReadOnlyList<string> htmlAdjustments,
            IReadOnlyList<string> requestedRemovals)
        {
            IReadOnlyList<RecipientView> recipients = CapRecipients(updated.Draft.Recipients, out int total, out bool truncated);
            IReadOnlyList<AttachmentView> attachmentViews = CapAttachments(
                VerifiedAttachments(updated.Draft, updated.Attachments), out int _, out bool _);
            IReadOnlyList<string>? notRemoved = requestedRemovals
                .Where(n => !updated.AttachmentsRemoved.Any(r => string.Equals(r, n, StringComparison.OrdinalIgnoreCase)))
                .ToList();

            return new UpdateDraftOutcome
            {
                Status = "updated",
                Id = hitId,
                EntryId = updated.Draft.EntryId,
                Store = updated.Draft.StoreDisplayName,
                Folder = updated.Draft.ParentFolderName,
                Account = updated.Draft.SendUsingAccountSmtp,
                Subject = updated.Draft.Subject,
                Changed = updated.ChangedFields.Count > 0 ? updated.ChangedFields : null,
                Displayed = updated.Displayed,
                ConversationId = updated.Draft.ConversationId,
                Recipients = recipients,
                RecipientsTruncated = truncated ? true : (bool?)null,
                RecipientsTotal = truncated ? total : (int?)null,
                UnresolvedRecipients = updated.UnresolvedRecipients.Count > 0
                    ? updated.UnresolvedRecipients.Take(UnresolvedRecipientsCap).ToList()
                    : null,
                ConversationTopicPreserved = updated.ConversationTopicPreserved,
                Importance = ImportanceName(updated.Draft.Importance) is string name && name != "normal" ? name : null,
                ReadReceiptRequested = updated.Draft.ReadReceiptRequested ? true : (bool?)null,
                Signature = updated.SignatureOverrideName,
                SignatureApplied = updated.SignatureOverrideName != null ? updated.SignatureOverrideApplied : (bool?)null,
                BodyFormat = body?.FormatName,
                BodyPlacement = updated.BodyReplaced ? (updated.BodyPlacedViaWordEditor ? "wordEditor" : "html") : null,
                HtmlAdjustments = htmlAdjustments.Count > 0 ? htmlAdjustments : null,
                Attachments = attachmentViews.Count > 0 ? attachmentViews : null,
                AttachmentsTotalBytes = attachmentViews.Count > 0 ? attachmentViews.Sum(a => a.SizeBytes ?? 0) : (long?)null,
                AttachmentsAdded = updated.AttachmentsAdded.Count > 0 ? updated.AttachmentsAdded : null,
                AttachmentsRemoved = updated.AttachmentsRemoved.Count > 0 ? updated.AttachmentsRemoved : null,
                AttachmentsNotFound = notRemoved.Count > 0 ? notRemoved : null,
                AttachmentsFailed = updated.AttachmentsFailed.Count > 0 ? updated.AttachmentsFailed : null,
                InlineImagesDropped = updated.InlineImagesDropped > 0 ? updated.InlineImagesDropped : (int?)null,
                InlineImagesAdvice = updated.InlineImagesDropped > 0 ? InlineImagesDroppedAdvice : null,
            };
        }

        /// <summary>
        /// The remedy for a revision that lost an inline image (D47) - live-proven, not
        /// guessed: re-supplying <c>signature</c> makes the update re-insert the signature
        /// file, and the picture is then embedded as it goes in.
        /// </summary>
        internal const string InlineImagesDroppedAdvice =
            "Inline image(s) the draft carried were lost by this revision: they were still linked to a file "
            + "on disk rather than embedded, and re-rendering the document cannot preserve such a link. "
            + "Only drafts composed before this was fixed are affected. Call update_draft again with the "
            + "signature argument set to restore the signature (and its images) in embedded form.";

        // ------------------------------------------------------------------ send (Phase 5, v3.MD L5/D4)

        /// <summary>
        /// High-friction two-step send (D4). WITHOUT a valid <paramref name="confirmToken"/>
        /// nothing is sent: the call returns a warning plus a one-time token bound to the
        /// draft's EntryID and current content hash. WITH the token (single-use, short
        /// TTL, invalidated by any draft change) the send executes: identity is resolved
        /// from the draft's own store, pinned via the Phase-4 putref path and getter-
        /// verified in-session immediately before <c>Send()</c> - a mismatch aborts.
        /// Every step (token issued / send / refusal) writes an audit line; refusals
        /// throw <see cref="SendRefusedException"/>.
        /// </summary>
        public SendOutcome Send(string id, string? confirmToken = null, string? sentOnBehalfOf = null)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                throw new ArgumentException(
                    "id is required (the draft EntryID returned by a draft tool, or a hit id of a saved unsent draft).", nameof(id));
            }

            sentOnBehalfOf = string.IsNullOrWhiteSpace(sentOnBehalfOf) ? null : sentOnBehalfOf!.Trim();

            (string entryId, string? storeId, string? _, long _, string? hitId) = ResolveToEntryId(id);

            // Snapshot the draft's sendable state (across-store retry for direct
            // EntryIDs, same pattern as read/reply).
            ComSendableDraftState state = _gateway.Run(s =>
            {
                string? error = null;
                ComSendableDraftState? st = s.TryGetSendableDraftState(entryId, storeId, out error);
                if (st == null && storeId == null)
                {
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        st = s.TryGetSendableDraftState(entryId, store.StoreId, out error);
                        if (st != null)
                        {
                            break;
                        }
                    }
                }

                return st ?? throw new InvalidOperationException(
                    "The draft could not be opened (" + (error ?? "unknown")
                    + "). It may have been deleted, moved, or already sent - re-check with read or re-run search.");
            });

            if (state.IsSent)
            {
                throw RefuseSend("not_an_unsent_draft", state.EntryId, state.StoreDisplayName, state.ResolvedAccountSmtp,
                    "This item has already been sent (or is not a saved draft). Only saved, unsent drafts can be sent.");
            }

            if (state.ResolvedAccountSmtp == null)
            {
                throw RefuseSend("no_sending_account", state.EntryId, state.StoreDisplayName, null,
                    "No profile account delivers into the store holding this draft ('" + (state.StoreDisplayName ?? "unknown")
                    + "'), so a verified send identity cannot be established. Move the draft creation to one of the accounts from list_accounts.");
            }

            string contentHash = SendContentHash.Compute(
                state.Subject, state.Recipients, state.BodyText, sentOnBehalfOf, state.Attachments, state.BodyHtmlDigest);

            if (string.IsNullOrWhiteSpace(confirmToken))
            {
                return IssueSendToken(state, contentHash, hitId, sentOnBehalfOf);
            }

            SendTokenDecision decision = _sendTokens.Consume(confirmToken!.Trim(), state.EntryId, contentHash);
            if (decision != SendTokenDecision.Valid)
            {
                throw RefuseSend(DescribeTokenDecision(decision), state.EntryId, state.StoreDisplayName, state.ResolvedAccountSmtp,
                    BuildTokenRefusalMessage(decision));
            }

            // Confirmed: execute as ONE STA operation (re-verify content INSIDE, pin +
            // hard-verify identity, then Send) - v3.MD section 12 Phase-4/5 rules.
            string? sendError = null;
            ComSendResult? sent = _gateway.Run(s => s.TrySendDraft(state.EntryId, state.StoreId, contentHash, sentOnBehalfOf, out sendError));
            if (sent == null)
            {
                throw MapSendFailure(sendError, state);
            }

            AuditSend(sent, hitId);
            IReadOnlyList<RecipientView> sentRecipients = CapRecipients(sent.Recipients, out int sentTotal, out bool sentTruncated);
            return new SendOutcome
            {
                Status = "sent",
                Sent = true,
                Id = hitId,
                EntryId = sent.EntryIdAtSend,
                Store = sent.StoreDisplayName,
                Account = sent.AccountSmtp,
                AccountVerified = true,
                SentOnBehalfOf = sent.SentOnBehalfOfName,
                Subject = sent.Subject,
                Recipients = sentRecipients,
                RecipientsTruncated = sentTruncated ? true : (bool?)null,
                RecipientsTotal = sentTruncated ? sentTotal : (int?)null,
            };
        }

        private SendOutcome IssueSendToken(ComSendableDraftState state, string contentHash, string? hitId, string? sentOnBehalfOf)
        {
            string token = _sendTokens.Issue(state.EntryId, contentHash);
            double ttlSeconds = _sendTokens.TimeToLive.TotalSeconds;
            try
            {
                Audit.AuditLog.Append(
                    "send_token_issued",
                    ("entryId", state.EntryId),
                    ("store", state.StoreDisplayName),
                    ("account", state.ResolvedAccountSmtp),
                    ("recipients", state.Recipients.Count.ToString(CultureInfo.InvariantCulture)),
                    ("expiresInSeconds", ttlSeconds.ToString("F0", CultureInfo.InvariantCulture)),
                    ("onBehalfOf", sentOnBehalfOf),
                    ("token", token));
            }
            catch (InvalidOperationException ex)
            {
                // No token without its audit line (D4 discipline).
                _sendTokens.Invalidate(token);
                throw new InvalidOperationException(
                    "The send confirmation token could not be audit-logged and was NOT issued: " + ex.Message, ex);
            }

            return new SendOutcome
            {
                Status = "confirmation_required",
                Sent = false,
                Warning = "NOT SENT (step 1 of 2). Automatic sending is a high-friction opt-in action; the default OutlookAI "
                    + "workflow is drafting and letting the user press Send themselves. Re-confirm with the user that THIS draft "
                    + "(check subject and recipients below) should be sent automatically. Only if that is explicitly wanted, call "
                    + "send again with confirm_token within " + ttlSeconds.ToString("F0", CultureInfo.InvariantCulture)
                    + " seconds. The token works exactly once, is bound to this draft and its current content, and becomes invalid "
                    + "if the draft changes.",
                ConfirmToken = token,
                TokenExpiresInSeconds = ttlSeconds,
                Id = hitId,
                EntryId = state.EntryId,
                Store = state.StoreDisplayName,
                Folder = state.ParentFolderName,
                Account = state.ResolvedAccountSmtp,
                SentOnBehalfOf = sentOnBehalfOf,
                Subject = state.Subject,
                Recipients = CapRecipients(state.Recipients, out int total, out bool truncated),
                RecipientsTruncated = truncated ? true : (bool?)null,
                RecipientsTotal = truncated ? total : (int?)null,
            };
        }

        /// <summary>Audit-logs a refusal and builds the exception (nothing was sent).</summary>
        private static SendRefusedException RefuseSend(string reason, string? entryId, string? store, string? account, string message)
        {
            Audit.AuditLog.Append(
                "send_refused",
                ("entryId", entryId),
                ("store", store),
                ("account", account),
                ("reason", reason));
            return new SendRefusedException(reason, message);
        }

        private static string DescribeTokenDecision(SendTokenDecision decision)
        {
            return decision switch
            {
                SendTokenDecision.Expired => "token_expired",
                SendTokenDecision.DraftMismatch => "token_draft_mismatch",
                SendTokenDecision.ContentChanged => "draft_changed",
                _ => "unknown_or_used_token",
            };
        }

        private static string BuildTokenRefusalMessage(SendTokenDecision decision)
        {
            return decision switch
            {
                SendTokenDecision.Expired =>
                    "The confirm_token has expired (tokens are short-lived by design). Nothing was sent.",
                SendTokenDecision.DraftMismatch =>
                    "The confirm_token was issued for a DIFFERENT draft and has now been invalidated. Nothing was sent.",
                SendTokenDecision.ContentChanged =>
                    "The draft changed after the confirm_token was issued, so the token is no longer valid. Nothing was sent - review the current draft first.",
                _ =>
                    "The confirm_token is unknown, already used, or from a previous server session (tokens work exactly once). Nothing was sent.",
            };
        }

        private static Exception MapSendFailure(string? sendError, ComSendableDraftState state)
        {
            string entryId = state.EntryId;
            string? store = state.StoreDisplayName;
            string? account = state.ResolvedAccountSmtp;
            if (sendError == "ContentChangedSinceToken")
            {
                return RefuseSend("draft_changed", entryId, store, account,
                    "The draft changed between token validation and the send, so the send was aborted. Nothing was sent.");
            }

            if (sendError == "AlreadySent" || sendError == "NotAMailItem")
            {
                return RefuseSend("not_an_unsent_draft", entryId, store, account,
                    "The item is no longer a saved, unsent draft. Nothing was sent.");
            }

            if (sendError == "NoSendingAccountForStore")
            {
                return RefuseSend("no_sending_account", entryId, store, null,
                    "No profile account delivers into the draft's store, so a verified send identity cannot be established. Nothing was sent.");
            }

            if (sendError == "SendIdentityVerificationFailed")
            {
                return RefuseSend("identity_verification_failed", entryId, store, account,
                    "The sending identity could not be verified on the draft (SendUsingAccount readback mismatch) - the send was "
                    + "aborted to avoid sending from the wrong account. Nothing was sent.");
            }

            if (sendError != null && sendError.StartsWith("SendCallFailed:", StringComparison.Ordinal))
            {
                return new InvalidOperationException(
                    "Outlook's Send call failed (" + sendError.Substring("SendCallFailed:".Length)
                    + "). The mail MAY be sitting in the Outbox - verify before retrying.");
            }

            return new InvalidOperationException(
                "The draft could not be re-opened for sending (" + (sendError ?? "unknown") + "). Nothing was sent.");
        }

        /// <summary>Send audit (load-bearing, D4): a failure surfaces with the send already executed.</summary>
        private static void AuditSend(ComSendResult sent, string? hitId)
        {
            try
            {
                Audit.AuditLog.Append(
                    "send",
                    ("entryId", sent.EntryIdAtSend),
                    ("store", sent.StoreDisplayName),
                    ("account", sent.AccountSmtp),
                    ("accountVerified", "true"),
                    ("recipients", sent.Recipients.Count.ToString(CultureInfo.InvariantCulture)),
                    ("onBehalfOf", sent.SentOnBehalfOfName),
                    ("hitId", hitId));
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException(
                    "The mail WAS SENT (draft EntryID " + sent.EntryIdAtSend
                    + ") but the audit line could not be written: " + ex.Message, ex);
            }
        }

        // ------------------------------------------------------------------ move + archive (D39, S1 v2)

        /// <summary>
        /// Per-item reason when the batch ran out of <see cref="MoveBatchBudgetMs"/>. Same
        /// shape as the audit-log short circuit: nothing is silently dropped, the items that
        /// did move still carry their new EntryIDs, and the remedy is a smaller batch.
        /// </summary>
        internal const string BatchBudgetExhaustedMessage =
            "Not attempted: the batch ran out of its time budget. The items above were processed - re-issue the rest "
            + "as a smaller batch.";

        /// <summary>Maximum ids per move_mail/archive_mail call (T1-pinned).</summary>
        public const int MoveIdsCap = 50;

        /// <summary>
        /// Standing advice attached whenever items moved (T1-pinned): the EntryID-change
        /// and undo semantics agents must know.
        /// </summary>
        public const string MoveEntryIdAdvice =
            "Moving changed each item's EntryID: use newEntryId from now on (old ids and existing index rows for these "
            + "items are stale until the index catches up - re-run search for fresh ids). Undo any move by calling "
            + "move_mail with newEntryId and folder = its fromFolder.";

        /// <summary>
        /// Moves 1-50 items (hit ids or EntryIDs) to a store-relative folder path,
        /// SAME-STORE only (D39 v1): each item moves within its own store; when
        /// <paramref name="store"/> is given, items living elsewhere fail per-item with
        /// a cross-store error. Content-preserving and reversible - per-item results
        /// carry old/new EntryIDs and fromFolder as the undo address; every move is
        /// audit-logged (load-bearing: an unwritable audit log stops the batch).
        /// Deleted Items (and subtree) and Outbox are refused as targets (S1 v2: no
        /// delete surface).
        /// </summary>
        public MoveMailOutcome MoveMail(IReadOnlyList<string>? ids, string? folder, bool createFolder = false, string? store = null)
        {
            IReadOnlyList<string> requestIds = ValidateMoveIds(ids);
            IReadOnlyList<string> segments = ParseFolderSegments(folder)
                ?? throw new ArgumentException(
                    "folder is required: the store-relative target path, e.g. 'Archive/2026' (see list_folders).", nameof(folder));
            string? requiredStore = string.IsNullOrWhiteSpace(store) ? null : store!.Trim();
            string targetFolderEcho = string.Join("/", segments);

            List<MoveItemView> items = new List<MoveItemView>(requestIds.Count);
            List<string> createdFolders = new List<string>();
            bool auditBroken = false;
            Stopwatch batchClock = Stopwatch.StartNew();
            foreach (string id in requestIds)
            {
                if (auditBroken)
                {
                    items.Add(FailedItem(id, "Not attempted: the audit log is unavailable (every move must be audited)."));
                    continue;
                }

                if (batchClock.ElapsedMilliseconds > MoveBatchBudgetMs)
                {
                    items.Add(FailedItem(id, BatchBudgetExhaustedMessage));
                    continue;
                }

                MoveItemView item = MoveOne(
                    id, segments, createFolder, requiredStore, targetFolderEcho, createdFolders, out bool auditFailed);
                auditBroken |= auditFailed;
                items.Add(item);
            }

            int movedCount = items.Count(i => i.Ok);
            return new MoveMailOutcome
            {
                Requested = requestIds.Count,
                Moved = movedCount,
                Failed = items.Count - movedCount,
                TargetFolder = targetFolderEcho,
                CreatedFolders = createdFolders.Count > 0 ? createdFolders : null,
                Items = items,
                Advice = movedCount > 0 ? new[] { MoveEntryIdAdvice } : null,
            };
        }

        /// <summary>
        /// Archives 1-50 items (hit ids or EntryIDs): moves each to ITS OWN store's
        /// DESIGNATED Archive folder - the folder Outlook's own Archive action
        /// (Backspace/mobile swipe/OWA) uses, resolved per store
        /// (localization-proof, never guessed by name; see
        /// <see cref="ArchiveFolderResolution"/>). A store without a designated
        /// archive folder fails per-item; nothing is ever created for it. Same result,
        /// undo and audit semantics as <see cref="MoveMail"/>.
        /// </summary>
        public ArchiveMailOutcome ArchiveMail(IReadOnlyList<string>? ids)
        {
            IReadOnlyList<string> requestIds = ValidateMoveIds(ids);

            Dictionary<string, (ComArchiveFolderInfo? Info, string? Error)> archiveByStore =
                new Dictionary<string, (ComArchiveFolderInfo?, string?)>(StringComparer.OrdinalIgnoreCase);
            List<MoveItemView> items = new List<MoveItemView>(requestIds.Count);
            bool auditBroken = false;
            Stopwatch batchClock = Stopwatch.StartNew();
            foreach (string id in requestIds)
            {
                if (auditBroken)
                {
                    items.Add(FailedItem(id, "Not attempted: the audit log is unavailable (every move must be audited)."));
                    continue;
                }

                if (batchClock.ElapsedMilliseconds > MoveBatchBudgetMs)
                {
                    items.Add(FailedItem(id, BatchBudgetExhaustedMessage));
                    continue;
                }

                MoveItemView item = ArchiveOne(id, archiveByStore, out bool auditFailed);
                auditBroken |= auditFailed;
                items.Add(item);
            }

            int archivedCount = items.Count(i => i.Ok);
            List<ArchiveFolderView> resolved = archiveByStore.Values
                .Where(v => v.Info != null)
                .Select(v => new ArchiveFolderView
                {
                    Store = v.Info!.StoreDisplayName,
                    Folder = v.Info!.StoreRelativePath,
                    Via = v.Info!.Via,
                })
                .OrderBy(v => v.Store, StringComparer.OrdinalIgnoreCase)
                .ToList();

            return new ArchiveMailOutcome
            {
                Requested = requestIds.Count,
                Archived = archivedCount,
                Failed = items.Count - archivedCount,
                ArchiveFolders = resolved.Count > 0 ? resolved : null,
                Items = items,
                Advice = archivedCount > 0 ? new[] { MoveEntryIdAdvice } : null,
            };
        }

        private MoveItemView ArchiveOne(
            string id,
            Dictionary<string, (ComArchiveFolderInfo? Info, string? Error)> archiveByStore,
            out bool auditFailed)
        {
            auditFailed = false;
            string entryId;
            string? storeId;
            string? hitId;
            try
            {
                (entryId, storeId, _, _, hitId) = ResolveToEntryId(id);
            }
            catch (Exception ex) when (ex is ArgumentException || ex is InvalidOperationException)
            {
                return FailedItem(id, ex.Message);
            }

            try
            {
                // Learn the item's own store first (cross-store retry for bare EntryIDs),
                // then resolve THAT store's designated archive folder (memoized per call).
                ComDraftInfo? info = _gateway.Run(s =>
                {
                    ComDraftInfo? r = s.TryGetMailInfo(entryId, storeId, out string? infoError);
                    if (r == null && storeId == null)
                    {
                        foreach (ComStoreDetail candidate in GetStoreDetails(s))
                        {
                            r = s.TryGetMailInfo(entryId, candidate.StoreId, out infoError);
                            if (r != null)
                            {
                                break;
                            }
                        }
                    }

                    return r;
                });
                if (info?.StoreDisplayName == null)
                {
                    return FailedItem(id, "The item could not be opened. Re-run search - it may have moved (EntryIDs change on moves).");
                }

                if (!archiveByStore.TryGetValue(info.StoreDisplayName, out (ComArchiveFolderInfo? Info, string? Error) archive))
                {
                    archive = _gateway.Run(s =>
                    {
                        ComArchiveFolderInfo? resolvedInfo = s.TryResolveArchiveFolder(info.StoreDisplayName, out string? resolveError);
                        return (resolvedInfo, resolveError);
                    });
                    archiveByStore[info.StoreDisplayName] = archive;
                }

                if (archive.Info == null)
                {
                    return FailedItem(id, DescribeArchiveResolutionFailure(info.StoreDisplayName, archive.Error));
                }

                ComArchiveFolderInfo target = archive.Info;
                (ComMoveItemResult? moved, string? moveError) = _gateway.Run(s =>
                {
                    ComMoveItemResult? r = s.TryMoveItemToFolderId(entryId, info.StoreId ?? storeId, target.EntryId, target.StoreId, out string? e);
                    return (r, e);
                });
                if (moved == null)
                {
                    return FailedItem(id, DescribeMoveFailure(moveError, target.StoreRelativePath, requestedStore: null, createFolder: false));
                }

                return CompleteMove("archive_mail", id, hitId, moved, out auditFailed);
            }
            catch (OutlookUnavailableException)
            {
                throw;
            }
            catch (Exception ex) when (ex is InvalidOperationException || OutlookComSession.IsComCallFailure(ex))
            {
                return FailedItem(id, ex.Message);
            }
        }

        private MoveItemView MoveOne(
            string id,
            IReadOnlyList<string> segments,
            bool createFolder,
            string? requestedStore,
            string targetFolderEcho,
            List<string> createdFolders,
            out bool auditFailed)
        {
            auditFailed = false;
            string entryId;
            string? storeId;
            string? hitId;
            try
            {
                (entryId, storeId, _, _, hitId) = ResolveToEntryId(id);
            }
            catch (Exception ex) when (ex is ArgumentException || ex is InvalidOperationException)
            {
                return FailedItem(id, ex.Message);
            }

            try
            {
                (ComMoveItemResult? moved, string? comError) = _gateway.Run(s =>
                {
                    ComMoveItemResult? r = s.TryMoveItemToPath(entryId, storeId, segments, createFolder, requestedStore, out string? e);
                    if (r == null && storeId == null && e == "ItemNotFound")
                    {
                        // Direct EntryID without a known store: retry across stores
                        // (same pattern as read/draft ops).
                        foreach (ComStoreDetail candidate in GetStoreDetails(s))
                        {
                            r = s.TryMoveItemToPath(entryId, candidate.StoreId, segments, createFolder, requestedStore, out e);
                            if (r != null || e != "ItemNotFound")
                            {
                                break;
                            }
                        }
                    }

                    return (r, e);
                });

                if (moved == null)
                {
                    return FailedItem(id, DescribeMoveFailure(comError, targetFolderEcho, requestedStore, createFolder));
                }

                foreach (string created in moved.CreatedFolderPaths)
                {
                    if (!createdFolders.Contains(created, StringComparer.OrdinalIgnoreCase))
                    {
                        createdFolders.Add(created);
                    }
                }

                return CompleteMove("move_mail", id, hitId, moved, out auditFailed);
            }
            catch (OutlookUnavailableException)
            {
                throw;
            }
            catch (Exception ex) when (ex is InvalidOperationException || OutlookComSession.IsComCallFailure(ex))
            {
                return FailedItem(id, ex.Message);
            }
        }

        /// <summary>
        /// Post-move bookkeeping shared by move_mail/archive_mail: writes the
        /// load-bearing audit line (op=move_mail/archive_mail with from-&gt;to), and
        /// refreshes the hit cache so a moved hit id keeps resolving (to the NEW
        /// EntryID) within this session. When the audit write fails the move HAS
        /// happened - the item result says so and carries the new EntryID, and the
        /// caller stops the batch.
        /// </summary>
        private MoveItemView CompleteMove(string operation, string id, string? hitId, ComMoveItemResult moved, out bool auditFailed)
        {
            auditFailed = false;
            try
            {
                Audit.AuditLog.Append(
                    operation,
                    ("entryId", moved.OldEntryId),
                    ("newEntryId", moved.NewEntryId),
                    ("store", moved.StoreDisplayName),
                    ("fromFolder", moved.FromFolderPath),
                    ("toFolder", moved.ToFolderPath),
                    ("hitId", hitId));
            }
            catch (InvalidOperationException ex)
            {
                auditFailed = true;
                return FailedItem(
                    id,
                    "The item WAS moved (newEntryId " + moved.NewEntryId + ", fromFolder '" + moved.FromFolderPath
                    + "') but the audit line could not be written: " + ex.Message);
            }

            if (hitId != null && _hits.TryGetValue(hitId, out CachedHit? cached))
            {
                cached.LocatedEntryId = moved.NewEntryId;
                cached.LocatedVia = "cached";
            }

            return new MoveItemView
            {
                Id = id,
                Ok = true,
                Store = moved.StoreDisplayName,
                FromFolder = moved.FromFolderPath,
                ToFolder = moved.ToFolderPath,
                OldEntryId = moved.OldEntryId,
                NewEntryId = moved.NewEntryId,
            };
        }

        private static MoveItemView FailedItem(string id, string error)
        {
            return new MoveItemView { Id = id, Ok = false, Error = error };
        }

        /// <summary>
        /// Validates the ids array of move_mail/archive_mail BEFORE any COM work
        /// (1-<see cref="MoveIdsCap"/> non-blank unique entries). Pure; throws
        /// ArgumentException with agent-actionable text.
        /// </summary>
        public static IReadOnlyList<string> ValidateMoveIds(IReadOnlyList<string>? ids)
        {
            if (ids == null || ids.Count == 0)
            {
                throw new ArgumentException(
                    "ids is required: 1-" + MoveIdsCap + " hit ids (from search/thread) or EntryID hex strings.", nameof(ids));
            }

            if (ids.Count > MoveIdsCap)
            {
                throw new ArgumentException(
                    "Too many ids (" + ids.Count + "): at most " + MoveIdsCap + " items per call - split the batch.", nameof(ids));
            }

            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<string> cleaned = new List<string>(ids.Count);
            foreach (string? id in ids)
            {
                if (string.IsNullOrWhiteSpace(id))
                {
                    throw new ArgumentException("ids must not contain blank entries.", nameof(ids));
                }

                string trimmed = id!.Trim();
                if (!seen.Add(trimmed))
                {
                    throw new ArgumentException(
                        "Duplicate id '" + trimmed + "' - each item can be moved once per call.", nameof(ids));
                }

                cleaned.Add(trimmed);
            }

            return cleaned;
        }

        /// <summary>
        /// Maps content-free COM move errors to agent-actionable per-item error text
        /// (pure, public for T1 pinning).
        /// </summary>
        public static string DescribeMoveFailure(string? comError, string targetFolder, string? requestedStore, bool createFolder)
        {
            if (comError != null && comError.StartsWith("CrossStoreTarget:", StringComparison.Ordinal))
            {
                string itemStore = comError.Substring("CrossStoreTarget:".Length);
                return "Cross-store move refused: the item lives in store '" + itemStore + "'"
                    + (requestedStore != null ? " but the requested store is '" + requestedStore + "'" : string.Empty)
                    + ". v1 moves are same-store only (archive semantics are same-store and EntryIDs are store-scoped) - "
                    + "call move_mail per store, or omit 'store' to move each item within its own store.";
            }

            switch (comError)
            {
                case "ItemNotFound":
                    return "The item could not be opened. Re-run search - it may have moved (EntryIDs change on moves).";
                case "NotAMailItem":
                    return "Only mail items can be moved by this tool.";
                case "TargetFolderNotFound":
                    return "Target folder '" + targetFolder + "' does not exist in the item's store"
                        + (createFolder ? "." : " - pass create_folder=true to create it, or check list_folders.");
                case "TargetFolderCreateFailed":
                    return "Target folder '" + targetFolder + "' could not be created in the item's store.";
                case "TargetNotAMailFolder":
                    return "Target folder '" + targetFolder + "' is not a mail folder.";
                case "TargetIsDeletedItems":
                    return "Refused: moving to Deleted Items (or a subfolder of it) is deletion semantics - this server has no "
                        + "delete surface. Ask the user to delete mail in Outlook themselves.";
                case "TargetIsOutbox":
                    return "Refused: the Outbox is not a valid move target.";
                case "AlreadyInTargetFolder":
                    return "The item is already in the target folder - nothing to move.";
                case "RootFolderUnavailable":
                    return "The item's store root could not be opened; retry when Outlook is responsive (see outlook_health).";
                default:
                    return "The move failed (" + (comError ?? "unknown") + "). Check outlook_health and retry.";
            }
        }

        /// <summary>
        /// Maps archive-resolution failures to agent-actionable per-item error text
        /// (pure, public for T1 pinning). A store without a designated archive folder
        /// is an ERROR - nothing is created silently.
        /// </summary>
        public static string DescribeArchiveResolutionFailure(string store, string? resolveError)
        {
            if (resolveError == "NoDesignatedArchiveFolder")
            {
                return "Store '" + store + "' has no designated Archive folder. Nothing was created - set one up via "
                    + "Outlook/OWA first (the folder the Archive button uses), then retry.";
            }

            if (resolveError != null && resolveError.StartsWith("ArchiveFolderVerificationFailed", StringComparison.Ordinal))
            {
                return "Store '" + store + "': the resolved archive-folder candidate failed verification ("
                    + resolveError + ") - refusing to move anything there.";
            }

            return "Store '" + store + "': the designated Archive folder could not be resolved ("
                + (resolveError ?? "unknown") + ").";
        }

        /// <summary>
        /// Maps recipients into the payload view capped at <see cref="RecipientsCap"/>
        /// (section 12: caps with has-more indicators). Pure and public for T1 pinning.
        /// The CAP is presentation-only - operations (send hash, identity checks,
        /// transport) always use the full COM-side list.
        /// </summary>
        public static IReadOnlyList<RecipientView> CapRecipients(
            IReadOnlyList<ComRecipientInfo> recipients, out int total, out bool truncated)
        {
            if (recipients == null)
            {
                throw new ArgumentNullException(nameof(recipients));
            }

            total = recipients.Count;
            truncated = total > RecipientsCap;
            return recipients
                .Take(truncated ? RecipientsCap : total)
                .Select(r => new RecipientView { Kind = r.Kind, Name = r.Name, Address = r.Address })
                .ToList();
        }

        /// <summary>
        /// Maps attachments into the payload view capped at <see cref="AttachmentsCap"/>.
        /// Original 1-based indexes are preserved, so attachments beyond the cap remain
        /// saveable via save_attachment even though they are not listed.
        /// </summary>
        /// <summary>
        /// Pure body-window math for read's body_offset paging (public for T1): returns
        /// the effective window start (the requested offset clamped to the body length),
        /// the window text of at most <paramref name="maxChars"/> characters, and
        /// whether more body exists BEYOND the window (the bodyTruncated contract - text
        /// before the window was skipped on request and does not count as truncation).
        /// </summary>
        public static (int Start, string Window, bool MoreBeyondWindow) ComputeBodyWindow(string fullBody, int offset, int maxChars)
        {
            if (fullBody == null)
            {
                throw new ArgumentNullException(nameof(fullBody));
            }

            if (offset < 0)
            {
                offset = 0;
            }

            if (maxChars < 0)
            {
                maxChars = 0;
            }

            int start = Math.Min(offset, fullBody.Length);
            int length = Math.Min(maxChars, fullBody.Length - start);
            string window = length > 0 ? fullBody.Substring(start, length) : string.Empty;
            return (start, window, start + length < fullBody.Length);
        }

        public static IReadOnlyList<AttachmentView> CapAttachments(
            IReadOnlyList<ComAttachmentInfo> attachments, out int total, out bool truncated)
        {
            if (attachments == null)
            {
                throw new ArgumentNullException(nameof(attachments));
            }

            total = attachments.Count;
            truncated = total > AttachmentsCap;
            return attachments
                .Take(truncated ? AttachmentsCap : total)
                .Select(a => new AttachmentView { Index = a.Index, FileName = a.FileName, SizeBytes = a.SizeBytes })
                .ToList();
        }

        // ------------------------------------------------------------------ outlook_health (Phase 7; merged with index_status in soak fix D37)

        /// <summary>
        /// Compact server + environment self-check behind the <c>outlook_health</c> tool
        /// (Phase 7, merged with the former index_status in soak fix D37): Outlook
        /// process/version/headless, probed COM liveness, store reachability, index
        /// freshness GLOBAL AND PER STORE with actionable freshness advice, WSearch
        /// service state, audit-log writability, tuning state (registry read - decoupled
        /// from the add-in) and the OutlookAISetup installer-mutex state. Read-only: COM
        /// is touched only while Outlook is ALREADY running (attach), so this never
        /// starts Outlook. Always returns a report; problems degrade the status instead
        /// of throwing.
        /// </summary>
        public HealthOutcome Health()
        {
            bool outlookRunning = ComGateway.IsOutlookRunning();
            bool mutexHeld = ComGateway.IsInstallerMutexHeld();
            List<string> problems = new List<string>();

            // Windows' own verdict on Outlook, obtained without COM. Costs microseconds,
            // cannot block, and stays truthful when every COM-shaped thing is stuck - so
            // health can state plainly what is wrong instead of inferring it from its own
            // failures.
            OutlookLivenessState liveness = OutlookLiveness.Probe(out string livenessDetail);
            if (liveness == OutlookLivenessState.Hung)
            {
                problems.Add("Outlook is running but is NOT RESPONDING (" + livenessDetail + "). Windows reports its "
                    + "windows as hung, so anything needing Outlook is refused immediately rather than left waiting. "
                    + "search still returns indexed mail. Restarting Outlook clears this.");
            }
            else if (liveness == OutlookLivenessState.Starting)
            {
                problems.Add("Outlook is still starting up. Requests needing it return retry guidance instead of "
                    + "waiting; search returns indexed mail meanwhile.");
            }

            // Which Office hive every registry-backed answer below is coming out of. Reported
            // FIRST among the registry problems because it is the one that explains the others:
            // an unsupported major makes accounts, signature defaults and the search settings all
            // read empty at once, which is otherwise indistinguishable from a broken install.
            string? officeVersion = HealthReporting.DetectedOfficeVersion();
            if (officeVersion == null)
            {
                problems.Add(HealthReporting.NoOfficeVersionProblem);
            }


            // Outlook + stores (attach-only while running; never a cold start).
            bool comProbeFailed = false;
            int? storesReachable = null;
            List<string>? storeNames = null;
            if (outlookRunning)
            {
                try
                {
                    // Bounded: health is asked precisely when Outlook may be unresponsive,
                    // so it must report that quickly rather than join it. Exceeding the
                    // budget degrades this block; it never fails the report.
                    IReadOnlyList<ComStoreDetail> stores = _gateway.Run(GetStoreDetails, HealthProbeBudgetMs);
                    storesReachable = stores.Count;
                    storeNames = stores.Select(s => s.DisplayName).ToList();
                    if (stores.Count == 0)
                    {
                        problems.Add("Outlook is running but reports no stores.");
                    }
                }
                catch (Exception ex) when (ex is not OutOfMemoryException)
                {
                    comProbeFailed = true;

                    // Only claim a wait when there actually was one. When the liveness
                    // check above already found Outlook hung or starting, this call was
                    // refused in microseconds - saying it "did not answer within 5s" would
                    // be a second, contradictory account of the same fact.
                    if (liveness == OutlookLivenessState.Responsive || liveness == OutlookLivenessState.NotRunning)
                    {
                        problems.Add("Outlook is running but did not answer within "
                            + (HealthProbeBudgetMs / 1000).ToString(CultureInfo.InvariantCulture)
                            + "s (" + ex.GetType().Name + "). It may be busy, showing a dialog, or not responding; "
                            + "search still returns indexed results meanwhile.");
                    }
                }
            }

            // Supervision state, read AFTER the probe above: the probe may itself have
            // replaced a faulted child, and reading first would report state=faulted in the
            // same breath as a store list the new child just fetched.
            ComHostDiagnostics comHost = _gateway.GetDiagnostics();
            if (comHost.RestartCount > 0)
            {
                problems.Add("The Outlook COM host has been restarted "
                    + comHost.RestartCount.ToString(CultureInfo.InvariantCulture)
                    + " time(s) this session after Outlook stopped answering.");
            }

            if (!string.IsNullOrEmpty(comHost.LastFailure))
            {
                problems.Add("Last COM host failure: " + comHost.LastFailure);
            }

            if (!string.IsNullOrEmpty(comHost.InjectedFault))
            {
                problems.Add("A TEST FAULT is injected into the COM host (" + comHost.InjectedFault
                    + "). This is not a real failure; unset OUTLOOKAI_COMHOST_FAULT to clear it.");
            }

            if (mutexHeld)
            {
                problems.Add("The add-in installer mutex (OutlookAISetup) is held - COM tools return retry-later until the update finishes.");
            }

            // Index freshness (global + per store) + WSearch service state + advice
            // (the former index_status content, merged in soak fix D37).
            List<string> advice = new List<string>();
            string provider;
            DateTime? newestIndexed = null;
            double? ageMinutes = null;
            List<StoreStaleness>? perStore = null;
            try
            {
                IndexSearchService index = _index.Value;
                provider = index.Provider.ToString();

                // ONE clock over the WHOLE index block, started before the first statement.
                // It used to start after catalog discovery and the COM-assisted enrichment,
                // so the two steps most able to run long were the two outside the budget
                // that was supposed to bound them.
                System.Diagnostics.Stopwatch indexClock = System.Diagnostics.Stopwatch.StartNew();

                // Short per-query timeout, for the same reason the COM probe has one: a
                // saturated Windows Search indexer is precisely when health gets asked,
                // and the default 30 s per query across a global probe plus one per store
                // would let this block alone run into minutes.
                IndexStalenessReport staleness = index.GetStaleness(commandTimeoutSeconds: HealthIndexTimeoutSeconds);
                newestIndexed = staleness.NewestIndexedReceivedUtc;
                ageMinutes = staleness.Age?.TotalMinutes;

                // The unordered discovery sample misses tiny idle stores (Phase-1 fact
                // 5); when Outlook is already running, its store list closes the gap via
                // targeted per-address discovery. Never STARTS Outlook here.
                // Skipped when the probe above already timed out: Outlook is demonstrably
                // not answering right now, and this step is an ENRICHMENT whose result is
                // optional. Asking again only costs the caller another full budget to
                // learn what we just learned.
                if (outlookRunning && !comProbeFailed)
                {
                    try
                    {
                        EnsureCatalogCoverageFromCom(indexClock);
                    }
                    catch (Exception ex) when (ex is not OutOfMemoryException)
                    {
                        // Best-effort enrichment only.
                    }
                }

                perStore = new List<StoreStaleness>();
                bool perStoreComplete = true;
                foreach (StoreScopeInfo scopeInfo in GetCatalog(HealthIndexTimeoutSeconds))
                {
                    // Overall budget as well as a per-query one: a profile with many
                    // stores multiplies even a short timeout. Reporting fewer rows is a
                    // fine outcome; taking minutes to report all of them is not.
                    if (indexClock.ElapsedMilliseconds > HealthPerStoreIndexBudgetMs)
                    {
                        perStoreComplete = false;
                        advice.Add("Per-store index freshness is incomplete: the index did not answer quickly enough for "
                            + "every store. The global figure above is still accurate.");
                        break;
                    }

                    IndexStalenessReport scoped = index.GetStaleness(
                        scopeInfo.StorePrefix, commandTimeoutSeconds: HealthIndexTimeoutSeconds);
                    perStore.Add(new StoreStaleness
                    {
                        Store = scopeInfo.StoreDisplayName,
                        NewestIndexedUtc = scoped.NewestIndexedReceivedUtc,

                        // In the catalog means the index has rows under this store's own
                        // prefix. A missing timestamp beside it says those rows hold no mail.
                        InLocalIndex = true,
                    });
                }

                // The two lists, compared. Until this ran, index.perStore was built from the
                // index-derived catalog and outlook.stores from COM, and a store present in
                // one and absent from the other was reported by neither - which is precisely
                // the store whose searches fall back to a fixed window, i.e. the condition an
                // operator opens this tool to find. Both lists were already in hand.
                AddStoresMissingFromIndex(perStore, storeNames, perStoreComplete, indexClock, problems);
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                provider = "unavailable: " + ex.GetType().Name;
                problems.Add("The SystemIndex is unreachable - index-backed search cannot run.");
            }

            IndexCurrency currency = ClassifyIndexCurrency(provider, newestIndexed);
            if (currency == IndexCurrency.NoMailAtAll)
            {
                // Nothing indexed ANYWHERE. Not a lag, not a quiet mailbox: every search on
                // this machine is then served by the freshness sweep's fallback window alone,
                // and mail older than it cannot be found. A PROBLEM, so status says degraded -
                // this is the tool an operator opens to discover exactly this, and it used to
                // answer "Index is current; searches run at index speed" over it, because a
                // null age was read as "no lag" rather than as "no index".
                problems.Add("The local index holds NO mail at all, so searches cannot use it: they fall back to a live "
                    + "sweep of the last " + EmptyIndexSweepWindow.TotalDays.ToString("F0", CultureInfo.InvariantCulture)
                    + " days in Inbox, Sent Items, Deleted Items and Junk Email. Older mail is not findable through "
                    + "search - check that Windows Search indexes Outlook (WSearch running, the Outlook data files "
                    + "ticked in Indexing Options) and use exhaustive:true meanwhile.");
            }

            if (!outlookRunning)
            {
                advice.Add("Outlook is not running: the index stops advancing and search's freshness sweep will start Outlook"
                    + (mutexHeld
                        ? " - but an add-in update is in progress, so the sweep degrades to index-only results until it finishes."
                        : "."));
            }
            else if (ageMinutes.HasValue && ageMinutes.Value > StaleIndexNoticeMinutes)
            {
                advice.Add("Newest indexed mail is " + ageMinutes.Value.ToString("F0", CultureInfo.InvariantCulture)
                    + " minutes old. search covers the gap automatically with its freshness sweep.");
            }

            // "Current" is a claim about a MEASURED frontier, so it needs one. It used to be
            // said on a null age too - i.e. over an index with no mail in it at all - which
            // told the operator that searches run at index speed over an index that cannot
            // serve one. Same decision as the problem above, read the other way round.
            if (advice.Count == 0 && currency == IndexCurrency.Measured && ageMinutes.HasValue)
            {
                advice.Add("Index is current (newest mail "
                    + ageMinutes.Value.ToString("F1", CultureInfo.InvariantCulture)
                    + " min ago); searches run at index speed.");
            }

            int? wsearchStart = HealthReporting.TryReadWSearchStartValue();
            bool? indexerRunning = HealthReporting.TryIsProcessRunning("SearchIndexer");
            if (wsearchStart == 4)
            {
                problems.Add("The Windows Search service (WSearch) is disabled - the index cannot advance.");
            }
            else if (indexerRunning == false)
            {
                problems.Add("SearchIndexer.exe is not running - index updates are paused until the WSearch service runs.");
            }

            // Audit log writability (write tools fail-closed without it).
            bool auditWritable = Audit.AuditLog.TryProbeWritable(Audit.AuditLog.DefaultDirectory, out string? auditError);
            if (!auditWritable)
            {
                problems.Add("The audit log is not writable (" + (auditError ?? "unknown") + ") - draft/save/send operations will fail.");
            }

            return new HealthOutcome
            {
                Status = problems.Count == 0 ? "ok" : "degraded",
                Problems = problems.Count > 0 ? problems : null,
                Outlook = new OutlookHealthView
                {
                    Running = outlookRunning,
                    // Headless probe AFTER the store attach above: window presence is
                    // re-read here so a just-promoted Outlook reports false (SF-3).
                    Headless = outlookRunning ? HealthReporting.TryGetOutlookHeadless() : null,
                    Version = HealthReporting.TryGetOutlookVersion(),
                    // Null exactly when no supported Office major is registered - the state
                    // NoOfficeVersionProblem above spells out.
                    OfficeVersion = officeVersion,
                    InstallerMutexHeld = mutexHeld,
                    // PROBED liveness (SF-1 fix): never report a dead held session as
                    // connected; the probe also releases a dead session's refs.
                    ComConnected = _gateway.ProbeConnected(),
                    Responding = liveness == OutlookLivenessState.NotRunning
                        ? (bool?)null
                        : liveness == OutlookLivenessState.Responsive,
                    State = OutlookLiveness.Describe(liveness),
                    StoresReachable = storesReachable,
                    Stores = storeNames,
                    ComHost = comHost,
                },
                Index = new IndexHealthView
                {
                    Provider = provider,
                    NewestIndexedUtc = newestIndexed,
                    AgeMinutes = ageMinutes,
                    PerStore = perStore,
                    WSearchStartMode = HealthReporting.DescribeServiceStartMode(wsearchStart),
                    IndexerProcessRunning = indexerRunning,
                },
                Advice = advice.Count > 0 ? advice : null,
                Audit = new AuditHealthView
                {
                    Path = Audit.AuditLog.DefaultLogPath,
                    Writable = auditWritable,
                    Error = auditError,
                },
                Tuning = HealthReporting.ReadTuningStateFromRegistry(),
                // The apphost that was actually launched, which is precisely what a
                // registration has to name in order to spawn this server.
                Registration = HealthReporting.ReadMcpRegistration(HealthReporting.CurrentProcessPath()),
            };
        }

        // ------------------------------------------------------------------ list_accounts / list_folders

        /// <summary>Accounts + all stores with delegate and local-searchability flags (D22/D25).</summary>
        public AccountsOutcome ListAccounts()
        {
            (IReadOnlyList<ComAccountInfo> accounts, IReadOnlyList<ComStoreDetail> stores) = _gateway.Run(s =>
                (s.GetAccounts(), s.GetStoreDetails()));

            HashSet<string> deliveryStores = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (ComAccountInfo account in accounts)
            {
                if (account.DeliveryStoreDisplayName != null)
                {
                    deliveryStores.Add(account.DeliveryStoreDisplayName);
                }
            }

            List<StoreView> storeViews = new List<StoreView>(stores.Count);
            foreach (ComStoreDetail store in stores)
            {
                // Live-verified on this machine (Phase 2): delegate caches report
                // OlExchangeStoreType 1 (olExchangeDelegateMailbox) and, despite being
                // locally cached AND indexed, IsCachedExchange=false - so index
                // presence, not the cached flag, is the searchability ground truth
                // (D22/D25). Non-default account mailboxes report type 4.
                bool isDelegate = store.ExchangeStoreType == 1 && !deliveryStores.Contains(store.DisplayName);

                bool? inLocalIndex = null;
                try
                {
                    inLocalIndex = ProbeStoreInIndex(store.DisplayName, isDelegate);
                }
                catch (Exception ex) when (ex is not OutOfMemoryException)
                {
                    // Index unavailable - flag unknown.
                }

                bool onlineOnly = inLocalIndex.HasValue
                    ? !inLocalIndex.Value
                    : store.IsCachedExchange == false && store.ExchangeStoreType != 1 && store.ExchangeStoreType != 3;

                storeViews.Add(new StoreView
                {
                    DisplayName = store.DisplayName,
                    IsDelegate = isDelegate,
                    IsCachedExchange = store.IsCachedExchange,
                    ExchangeStoreType = store.ExchangeStoreType,
                    OnlineOnly = onlineOnly,
                    LocallySearchable = !onlineOnly,
                    InLocalIndex = inLocalIndex,
                });
            }

            return new AccountsOutcome
            {
                Accounts = accounts.Select(a => new AccountView
                {
                    SmtpAddress = a.SmtpAddress,
                    DisplayName = a.DisplayName,
                    DeliveryStore = a.DeliveryStoreDisplayName,
                }).ToList(),
                Stores = storeViews,
            };
        }

        /// <summary>
        /// Signature landscape (list_signatures - soak fix D37, R5 steering): installed
        /// signatures with plain-text excerpts (language/purpose detection) plus the
        /// registry-determined per-account default assignments. Pure filesystem +
        /// registry - no COM, never starts Outlook. Assignments degrade to unknown
        /// (absent fields + note) when the registry does not carry them - never guessed.
        /// </summary>
        public SignaturesOutcome ListSignatures()
        {
            IReadOnlyList<SignatureInfo> signatures = SignatureCatalog.ListSignatures();
            IReadOnlyList<SignatureAssignment> assignments = SignatureCatalog.ReadAccountAssignments();

            List<SignatureView> signatureViews = signatures
                .Select(s => new SignatureView { Name = s.Name, Excerpt = s.Excerpt })
                .ToList();

            List<SignatureAccountView>? accountViews = assignments.Count > 0
                ? assignments.Select(a => new SignatureAccountView
                {
                    Account = a.Account,
                    NewMessage = a.NewMessageSignature,
                    ReplyForward = a.ReplyForwardSignature,
                }).ToList()
                : null;

            string? note = null;
            if (assignments.Count == 0)
            {
                note = "Per-account default assignments could not be read from the profile registry - defaults are unknown "
                    + "(they may be roaming-managed). Omitting the signature parameter still applies whatever default Outlook uses.";
            }
            else if (assignments.Any(a => a.NewMessageSignature == null && a.ReplyForwardSignature == null))
            {
                note = "Accounts without listed assignments have no registry-recorded default - unknown (possibly no signature, "
                    + "possibly roaming-managed). Omitting the signature parameter applies whatever default Outlook uses.";
            }

            return new SignaturesOutcome
            {
                Signatures = signatureViews,
                Accounts = accountViews,
                Note = note,
            };
        }

        /// <summary>
        /// Signature management (manage_signature - soak fix D38): create/update/delete
        /// the signature file set under %APPDATA%\Microsoft\Signatures with the
        /// ALWAYS-ON pre-modification backup and optional per-account default
        /// assignment. Pure filesystem + registry - no COM, never starts Outlook. The
        /// audit line is load-bearing: a write that cannot be audited surfaces the
        /// failure (with the outcome preserved in the message) instead of hiding it.
        /// </summary>
        public ManageSignatureOutcome ManageSignature(ManageSignatureRequest request)
        {
            ManageSignatureOutcome outcome = SignatureManager.Manage(request);
            try
            {
                Audit.AuditLog.Append(
                    "manage_signature",
                    ("action", outcome.Action),
                    ("name", outcome.Name),
                    ("filesWritten", outcome.FilesWritten?.Count.ToString(CultureInfo.InvariantCulture)),
                    ("filesDeleted", outcome.FilesDeleted?.Count.ToString(CultureInfo.InvariantCulture)),
                    ("backupPath", outcome.BackupPath),
                    ("defaultAccount", outcome.DefaultSetForAccount),
                    ("defaultScope", outcome.DefaultSetScope),
                    ("defaultsCleared", outcome.DefaultsClearedForAccounts != null
                        ? string.Join(";", outcome.DefaultsClearedForAccounts)
                        : null));
            }
            catch (InvalidOperationException ex)
            {
                throw new InvalidOperationException(
                    "The signature operation succeeded (" + outcome.Action + " '" + outcome.Name + "'"
                    + (outcome.BackupPath != null ? ", backup at " + outcome.BackupPath : string.Empty)
                    + ") but the audit line could not be written: " + ex.Message, ex);
            }

            return outcome;
        }

        /// <summary>
        /// FULL folder trees (list_folders) in the stable traversal order (stores by
        /// display name, then depth-first with siblings by name), paged by
        /// <paramref name="offset"/> into windows of <see cref="FoldersPerCallCap"/>.
        /// Real profiles fit in one page; the bound is section 12 discipline.
        /// </summary>
        public FoldersOutcome ListFolders(string? store = null, int offset = 0)
        {
            if (offset < 0)
            {
                offset = 0;
            }

            IReadOnlyList<ComFolderInfo> folders = _gateway.Run(s => s.ListFolders(store, FolderWalkAbsoluteCap));
            return PageFolders(folders, offset);
        }

        /// <summary>
        /// Pure paging step of list_folders (public for T1): slices the stable-order
        /// flattened walk at [offset, offset + <see cref="FoldersPerCallCap"/>) and
        /// derives the has-more contract (truncated + nextOffset + total).
        /// </summary>
        public static FoldersOutcome PageFolders(IReadOnlyList<ComFolderInfo> folders, int offset)
        {
            if (folders == null)
            {
                throw new ArgumentNullException(nameof(folders));
            }

            if (offset < 0)
            {
                offset = 0;
            }

            int end = (int)Math.Min((long)offset + FoldersPerCallCap, folders.Count);
            List<ComFolderInfo> page = new List<ComFolderInfo>(Math.Max(0, end - offset));
            for (int i = offset; i < end; i++)
            {
                page.Add(folders[i]);
            }

            List<StoreFoldersView> byStore = page
                .GroupBy(f => f.StoreDisplayName, StringComparer.OrdinalIgnoreCase)
                .Select(g => new StoreFoldersView
                {
                    Store = g.Key,
                    Folders = g.Select(f => new FolderView
                    {
                        Path = f.Path,
                        Items = f.ItemCount,
                        Unread = f.UnreadCount,
                    }).ToList(),
                })
                .ToList();

            bool truncated = end < folders.Count;
            return new FoldersOutcome
            {
                Stores = byStore,
                FolderTotal = folders.Count,
                Offset = offset > 0 ? offset : (int?)null,
                Truncated = truncated,
                NextOffset = truncated ? end : (int?)null,
            };
        }

        // ------------------------------------------------------------------ hit cache + location

        private HitSummary RegisterIndexHit(IndexHit hit, int snippetChars)
        {
            string hitId = NextHitId();
            _hits[hitId] = new CachedHit { IndexHit = hit };
            string? snippet = hit.AutoSummary;
            if (snippet != null && snippetChars == 0)
            {
                snippet = null;
            }
            else if (snippet != null && snippet.Length > snippetChars)
            {
                snippet = snippet.Substring(0, snippetChars);
            }

            return new HitSummary
            {
                Id = hitId,
                Source = "index",
                Subject = hit.Subject,
                FromName = hit.FromName,
                FromAddress = hit.FromAddress,
                ReceivedUtc = hit.DateReceivedUtc,
                Store = FreshMerge.ResolveHitStore(hit),
                Folder = DescribeHitFolder(hit),
                Snippet = snippet,
                SizeBytes = hit.SizeBytes,
                IsRead = hit.IsRead,
                HasAttachments = hit.HasAttachments,
                IsAttachmentHit = hit.IsAttachmentHit,
                AttachmentFileName = hit.AttachmentFileName,
                ConversationId = hit.ConversationId,
            };
        }

        private HitSummary RegisterLiveHit(ComMailBrief item, int snippetChars, string source = "live")
        {
            string hitId = NextHitId();
            _hits[hitId] = new CachedHit
            {
                Live = item,
                LocatedEntryId = item.EntryId,
                LocatedStoreId = item.StoreId,
                LocatedVia = source == "live" ? "sweep" : source == "exhaustive" ? "exhaustive" : "conversation",
            };

            string? snippet = null;
            if (snippetChars > 0 && item.Body != null)
            {
                string collapsed = item.Body.Replace("\r", " ").Replace("\n", " ").Trim();
                snippet = collapsed.Length > snippetChars ? collapsed.Substring(0, snippetChars) : collapsed;
            }

            return new HitSummary
            {
                Id = hitId,
                Source = source,
                Subject = item.Subject,
                FromName = item.SenderName,
                FromAddress = item.SenderAddress,
                ReceivedUtc = ToUtc(item.ReceivedTime),
                Store = item.StoreDisplayName,
                Folder = item.FolderName,
                FolderKind = item.FolderKind,
                Snippet = snippet,
                SizeBytes = item.SizeBytes,
                IsRead = item.IsRead,
                HasAttachments = item.HasAttachments,
                IsAttachmentHit = false,
                ConversationId = null,
            };
        }

        private (string EntryId, string? StoreId, string? LocatedVia, long LocateMs, string? HitId) ResolveToEntryId(string id)
        {
            if (_hits.TryGetValue(id, out CachedHit? cached))
            {
                if (cached.LocatedEntryId != null)
                {
                    // Live hits and previously located hits resolve without any COM
                    // probing; report how THIS call resolved (LocatedVia keeps the
                    // original tier internally).
                    return (cached.LocatedEntryId, cached.LocatedStoreId, "cached", 0, id);
                }

                // Lazy locate (Phase-1: avg ~2 s per hit - cache the result).
                IndexHit hit = cached.IndexHit
                    ?? throw new InvalidOperationException("Hit cache entry is unlocatable.");
                int tolerance = hit.IsAttachmentHit ? AttachmentLocateToleranceSeconds : EmailLocateToleranceSeconds;
                Stopwatch stopwatch = Stopwatch.StartNew();
                HitLocationResult location = _gateway.Run(s => HitLocator.Locate(s, hit, tolerance));
                stopwatch.Stop();
                if (location.Tier == HitLocationTier.Failed || location.Located == null)
                {
                    // The remedy depends on WHY: a stale row for a folder that no longer
                    // exists cannot be fixed by re-running the search (it returns the same
                    // orphan row), so say which case this is (block (q), soak fix 16).
                    throw new InvalidOperationException(LocateFailureAdvice.Describe(location.Error));
                }

                string? storeId = null;
                if (location.StoreDisplayName != null)
                {
                    storeId = _gateway.Run(s => GetStoreDetails(s)
                        .FirstOrDefault(d => string.Equals(d.DisplayName, location.StoreDisplayName, StringComparison.OrdinalIgnoreCase))?.StoreId);
                }

                cached.LocatedEntryId = location.Located.EntryId;
                cached.LocatedStoreId = storeId;
                cached.LocatedVia = location.Tier switch
                {
                    HitLocationTier.UrlSegments => "urlSegments",
                    HitLocationTier.DelegateLeafName => "delegateLeafName",
                    _ => "itemPathDisplay",
                };
                return (cached.LocatedEntryId, storeId, cached.LocatedVia, stopwatch.ElapsedMilliseconds, id);
            }

            // Raw EntryID hex. The floor is DERIVED from the shortest structurally valid
            // MAPI entry id rather than picked: EntryIdCodec.MessageEntryIdLength bytes
            // (4 flag bytes + a 16-byte store UID + a 4-byte node id), two hex chars each.
            // A real message EntryID is typically far longer - 70+ bytes, 140+ hex chars -
            // but this is an ACCEPTANCE test, not a validation one: anything plausibly hex
            // and longer than a hit id ("h7") is handed to Outlook, which is the authority
            // on whether it opens. The comment used to claim 140 while the code said 48,
            // three times looser than its own stated truth and with nothing pinning either.
            if (id.Length >= MinRawEntryIdHexChars && id.Length % 2 == 0 && IsHex(id))
            {
                return (id, null, "directEntryId", 0, null);
            }

            throw new ArgumentException(
                "Unknown id '" + id + "'. Pass a hit id from a previous search/thread call in this session, or a full EntryID hex string.");
        }

        /// <summary>
        /// Enriches the index-derived store catalog from Outlook's own store list, to catch
        /// tiny idle stores the unordered discovery sample misses.
        /// <para>
        /// Called ONLY from <see cref="Health"/>, and bounded by the same short probe
        /// budget as the rest of it. It was unbounded until 2026-08-16, which quietly
        /// undid health's whole non-blocking guarantee: against an unresponsive Outlook
        /// this one best-effort ENRICHMENT step spent the full 120 s operation budget, and
        /// health took 126 s to report - measured - while the store probe beside it
        /// correctly gave up after 5 s. A step whose result is optional must never cost
        /// more than the step whose result is the point.
        /// </para>
        /// <para>
        /// Its INDEX half was still unbounded until the same rule was applied to it: two
        /// searches plus a probe per '@'-named store, each falling through to the 30 s index
        /// default. It now runs under health's index clock and stops the moment
        /// <see cref="HealthPerStoreIndexBudgetMs"/> is spent - dropping an enrichment is
        /// invisible in the report, where taking minutes is not.
        /// </para>
        /// </summary>
        /// <summary>What the index freshness probe established about the index as a whole.</summary>
        public enum IndexCurrency
        {
            /// <summary>A frontier was measured: the index holds mail and its age is known.</summary>
            Measured = 0,

            /// <summary>
            /// The index is reachable and holds NO mail whatsoever. Not a lag and not a quiet
            /// mailbox - a state in which the index tier cannot serve a single search.
            /// </summary>
            NoMailAtAll = 1,

            /// <summary>The index could not be reached, so nothing about it was established.</summary>
            Unavailable = 2,
        }

        /// <summary>
        /// The three things a freshness probe can establish about the index, told apart in one
        /// pure place because two of them look identical in the data: an index with no mail and
        /// an index that could not be read BOTH report a null frontier and a null age.
        /// <para>
        /// Reading the null alone is what made outlook_health say "Index is current; searches
        /// run at index speed" over an index holding zero mail rows - the one report an
        /// operator would open to discover exactly that, denying it. Public so T1 pins all
        /// three without an index.
        /// </para>
        /// </summary>
        public static IndexCurrency ClassifyIndexCurrency(string provider, DateTime? newestIndexedUtc)
        {
            if (provider == null || provider.StartsWith("unavailable", StringComparison.Ordinal))
            {
                return IndexCurrency.Unavailable;
            }

            return newestIndexedUtc.HasValue ? IndexCurrency.Measured : IndexCurrency.NoMailAtAll;
        }

        /// <summary>
        /// Which stores OUTLOOK reports the INDEX holds nothing for: the comparison the health
        /// report never made. <paramref name="probe"/> answers per store - true indexed, false
        /// not, null not established - and only a FALSE is reported.
        /// <para>
        /// Absence from <paramref name="indexKnownStores"/> is not evidence and must never be
        /// treated as any: that list comes from an unordered 2000-row discovery sample which
        /// misses small stores, and a delegate mailbox is indexed under its OWNER's subtree so
        /// it never appears under its own name at all. Name-comparison alone would report every
        /// shared mailbox on a delegate-heavy profile as unindexed, which is the failure mode
        /// that makes a completeness flag worthless.
        /// </para>
        /// <para>
        /// Pure over an injected probe, so T1 pins the rule - including the case that matters
        /// most, a store the probe cannot settle - without a Windows Search index.
        /// </para>
        /// </summary>
        public static IReadOnlyList<string> StoresMissingFromIndex(
            IReadOnlyList<string>? comStoreNames,
            IReadOnlyCollection<string>? indexKnownStores,
            Func<string, bool?> probe)
        {
            if (probe == null)
            {
                throw new ArgumentNullException(nameof(probe));
            }

            if (comStoreNames == null || comStoreNames.Count == 0)
            {
                return Array.Empty<string>();
            }

            HashSet<string> known = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string store in indexKnownStores ?? Array.Empty<string>())
            {
                known.Add(store);
            }

            List<string> missing = new List<string>();
            foreach (string store in comStoreNames)
            {
                if (store == null || known.Contains(store) || missing.Contains(store, StringComparer.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (probe(store) == false)
                {
                    missing.Add(store);
                }
            }

            return missing;
        }

        /// <summary>
        /// Adds one <see cref="StoreStaleness"/> row per store OUTLOOK reports that the index
        /// does not know, and raises a problem naming them. The report held both lists and
        /// compared neither, so the disagreement between them - the single condition that
        /// makes searches of a store fall back to a fixed window - had no field anywhere.
        /// <para>
        /// Every "no" here is a probe result, never an absence from the catalog: the catalog
        /// comes from an unordered 2000-row sample that misses small stores, and a delegate
        /// mailbox is indexed under its OWNER's subtree so it never appears under its own name
        /// at all. Comparing names alone would have reported every shared mailbox on this
        /// profile as unindexed. A store that could not be probed inside the budget is left
        /// out entirely rather than reported either way.
        /// </para>
        /// <para>
        /// Skipped when the per-store loop was cut short (<paramref name="perStoreComplete"/>)
        /// or Outlook gave no store list: a comparison against half a list would invent
        /// missing stores, and the truncation is already reported in its own words.
        /// </para>
        /// </summary>
        private void AddStoresMissingFromIndex(
            List<StoreStaleness> perStore,
            IReadOnlyList<string>? comStoreNames,
            bool perStoreComplete,
            System.Diagnostics.Stopwatch indexClock,
            List<string> problems)
        {
            if (!perStoreComplete || comStoreNames == null || comStoreNames.Count == 0)
            {
                return;
            }

            List<string> indexed = new List<string>(perStore.Count);
            foreach (StoreStaleness row in perStore)
            {
                indexed.Add(row.Store);
            }

            IReadOnlyList<string> missing = StoresMissingFromIndex(
                comStoreNames,
                indexed,

                // A store left unprobed because the budget ran out answers null, never false:
                // "not established" must not become "not indexed".
                store => indexClock.ElapsedMilliseconds > HealthPerStoreIndexBudgetMs
                    ? (bool?)null
                    : StoreHasIndexRows(store, HealthIndexTimeoutSeconds));

            foreach (string store in missing)
            {
                perStore.Add(new StoreStaleness { Store = store, InLocalIndex = false });
            }

            if (missing.Count > 0)
            {
                problems.Add("Outlook has " + missing.Count.ToString(CultureInfo.InvariantCulture)
                    + " store(s) the local index holds nothing for (" + string.Join(", ", missing)
                    + "). Searches covering them fall back to a live sweep of the last "
                    + EmptyIndexSweepWindow.TotalDays.ToString("F0", CultureInfo.InvariantCulture)
                    + " days in their arrival-path folders; older mail there is not findable through search. Add the "
                    + "data file(s) to Windows Indexing Options, or use exhaustive:true with store plus folder/after.");
            }
        }

        private void EnsureCatalogCoverageFromCom(System.Diagnostics.Stopwatch indexClock)
        {
            IReadOnlyList<ComStoreDetail> stores = _gateway.Run(GetStoreDetails, HealthProbeBudgetMs);
            foreach (ComStoreDetail store in stores)
            {
                if (indexClock.ElapsedMilliseconds > HealthPerStoreIndexBudgetMs)
                {
                    return;
                }

                if (store.DisplayName.IndexOf('@') < 0)
                {
                    continue;
                }

                bool known = GetCatalog(HealthIndexTimeoutSeconds).Any(s =>
                    string.Equals(s.StoreDisplayName, store.DisplayName, StringComparison.OrdinalIgnoreCase));
                if (known)
                {
                    continue;
                }

                StoreScopeInfo? targeted = _index.Value.TryDiscoverStoreScopeByAddress(
                    store.DisplayName, HealthIndexTimeoutSeconds);
                if (targeted != null)
                {
                    InvalidateCatalog(targeted);
                }
            }
        }

        private bool ProbeStoreInIndex(string displayName, bool isDelegate, int? commandTimeoutSeconds = null)
        {
            int timeout = commandTimeoutSeconds ?? SearchIndexTimeoutSeconds;
            IReadOnlyList<StoreScopeInfo> catalog = GetCatalog(timeout);
            foreach (StoreScopeInfo scopeInfo in catalog)
            {
                if (string.Equals(scopeInfo.StoreDisplayName, displayName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            if (isDelegate)
            {
                // Delegate items are indexed under the OWNER's /1/<delegate display name>
                // subtree (Phase-1 fact 3).
                foreach (StoreScopeInfo owner in catalog)
                {
                    if (_index.Value.ScopeHasAnyItem(owner.StorePrefix + "/1/" + displayName, timeout))
                    {
                        return true;
                    }
                }

                return false;
            }

            if (displayName.IndexOf('@') >= 0)
            {
                StoreScopeInfo? targeted = _index.Value.TryDiscoverStoreScopeByAddress(displayName, timeout);
                if (targeted != null)
                {
                    InvalidateCatalog(targeted);
                    return true;
                }
            }

            return false;
        }

        private string ResolveScope(string store, string? folder)
        {
            return ResolveFolderScope(store, folder, includeSubfolders: true).Scope;
        }

        /// <summary>
        /// Turns a store display name + folder path + include_subfolders into the index
        /// predicates, choosing the shape the store's namespace actually supports.
        /// <para>
        /// ⚠ THE DELEGATE FIX (soak fix 15). This method used to build a NESTED delegate
        /// URL (<c>&lt;host&gt;/1/&lt;delegate&gt;/&lt;path&gt;</c>) for a delegate
        /// folder. The delegate index namespace is FLAT - intermediate folders are dropped
        /// from both the mapi URL and System.ItemFolderPathDisplay - so that URL addressed
        /// a folder that does not exist and every delegate SUBFOLDER search returned 0
        /// rows, silently (measured: 8/8 probed subfolders, ~3,871 items across 15
        /// subfolders unreachable on this profile). Delegate scopes now go through
        /// <see cref="FolderScopeResolver.ForDelegateStore"/>: delegate STORE ROOT scope +
        /// folder-NAME equality, verified exact against COM ground truth.
        /// </para>
        /// </summary>
        private FolderScopeResolution ResolveFolderScope(string store, string? folder, bool includeSubfolders)
        {
            IReadOnlyList<StoreScopeInfo> catalog = GetCatalog(SearchIndexTimeoutSeconds);
            StoreScopeInfo? match = catalog.FirstOrDefault(s =>
                string.Equals(s.StoreDisplayName, store, StringComparison.OrdinalIgnoreCase));

            if (match == null && store.IndexOf('@') >= 0)
            {
                match = _index.Value.TryDiscoverStoreScopeByAddress(store, SearchIndexTimeoutSeconds);
                if (match != null)
                {
                    InvalidateCatalog(match);
                }
            }

            if (match != null)
            {
                return FolderScopeResolver.ForPrimaryStore(match.StorePrefix, folder, includeSubfolders);
            }

            // Delegate store: scope under an owner's /1/<name> subtree.
            foreach (StoreScopeInfo owner in catalog)
            {
                string delegateScope = owner.StorePrefix + "/1/" + store;
                bool exists;
                try
                {
                    exists = _index.Value.ScopeHasAnyItem(delegateScope, SearchIndexTimeoutSeconds);
                }
                catch (Exception ex) when (ex is not OutOfMemoryException)
                {
                    exists = false;
                }

                if (exists)
                {
                    // The COM tree is needed only for a delegate FOLDER scope: to map the
                    // requested subtree onto the flat leaf names it contains, and to spot
                    // leaf names the flat namespace merges.
                    IReadOnlyList<string>? tree = folder == null ? null : TryGetFolderPaths(store);
                    return FolderScopeResolver.ForDelegateStore(delegateScope, folder, includeSubfolders, tree);
                }
            }

            string known = string.Join(", ", catalog.Select(s => s.StoreDisplayName));
            throw new ArgumentException(
                "Store '" + store + "' was not found in the local index. Known stores: " + known
                + ". Use list_accounts for the full store list.");
        }

        /// <summary>
        /// Store-relative folder paths of one store from COM, cached for
        /// <see cref="FolderPathCacheTtl"/>. Returns null when Outlook cannot be reached -
        /// the caller then widens rather than narrowing on a guess.
        /// </summary>
        private IReadOnlyList<string>? TryGetFolderPaths(string store)
        {
            lock (_catalogLock)
            {
                // MonotonicClock: the stamp is only ever subtracted from a later reading of
                // the same clock to get the entry's age, never compared with anything outside
                // this process.
                if (_folderPaths.TryGetValue(store, out (IReadOnlyList<string> Paths, DateTime FetchedUtc) cached)
                    && MonotonicClock.UtcNow - cached.FetchedUtc <= FolderPathCacheTtl)
                {
                    return cached.Paths;
                }
            }

            IReadOnlyList<string> paths;
            try
            {
                paths = _gateway.Run(s => s.ListFolderPaths(store, FolderWalkAbsoluteCap));
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                return null;
            }

            if (paths.Count == 0)
            {
                return null;
            }

            lock (_catalogLock)
            {
                _folderPaths[store] = (paths, MonotonicClock.UtcNow);
            }

            return paths;
        }

        /// <summary>
        /// The store catalog, discovered once per process and then reused.
        /// <para>
        /// <paramref name="commandTimeoutSeconds"/> bounds the discovery statements when
        /// this is the call that performs it. Whichever caller gets there first supplies the
        /// timeout, which is why the two that can - search and health - both pass their own
        /// rather than letting it fall through to the 30 s index default underneath a budget
        /// measured in single-digit seconds.
        /// </para>
        /// </summary>
        private IReadOnlyList<StoreScopeInfo> GetCatalog(int? commandTimeoutSeconds)
        {
            lock (_catalogLock)
            {
                _catalog ??= _index.Value.DiscoverStoreScopes(StoreDiscoverySampleSize, commandTimeoutSeconds);
                return _catalog;
            }
        }

        /// <summary>
        /// Item URLs sampled to discover store scopes. The 2000-row pull measured 552 ms in
        /// the section-5 probes; a busy store dominates smaller samples, and the tiny idle
        /// ones it still misses are found by targeted per-address discovery instead.
        /// </summary>
        private const int StoreDiscoverySampleSize = 2000;

        private void InvalidateCatalog(StoreScopeInfo addition)
        {
            lock (_catalogLock)
            {
                if (_catalog == null)
                {
                    return;
                }

                List<StoreScopeInfo> updated = _catalog.ToList();
                if (!updated.Any(s => string.Equals(s.StorePrefix, addition.StorePrefix, StringComparison.OrdinalIgnoreCase)))
                {
                    updated.Add(addition);
                }

                _catalog = updated;
            }
        }

        private IReadOnlyList<ComStoreDetail> GetStoreDetails(IOutlookSession session)
        {
            lock (_catalogLock)
            {
                // MonotonicClock, for the same reason as the folder-path cache above: this
                // stamp exists only to be subtracted from a later reading of the same clock.
                if (_storeDetails == null || MonotonicClock.UtcNow - _storeDetailsFetchedUtc > StoreDetailsCacheTtl)
                {
                    _storeDetails = session.GetStoreDetails();
                    _storeDetailsFetchedUtc = MonotonicClock.UtcNow;
                }

                return _storeDetails;
            }
        }

        private string NextHitId()
        {
            int n = System.Threading.Interlocked.Increment(ref _nextHitId);
            return "h" + n.ToString(CultureInfo.InvariantCulture);
        }

        // ------------------------------------------------------------------ helpers

        private static string? DescribeHitFolder(IndexHit hit)
        {
            if (hit.StoreType == 1 && hit.FolderSegments.Count > 1)
            {
                return string.Join("/", hit.FolderSegments.Skip(1));
            }

            return hit.FolderSegments.Count > 0 ? string.Join("/", hit.FolderSegments) : null;
        }

        /// <summary>
        /// How far behind the index is, as prose, over the WIDEST lag in the search's scope.
        /// <paramref name="widestFrontierUtc"/> is the oldest per-store frontier when the
        /// sweep opened one window per store; the profile-wide frontier is the MAXIMUM across
        /// stores, so it is never the honest figure for a multi-store answer.
        /// </summary>
        private static string DescribeAge(IndexStalenessReport staleness, DateTime? widestFrontierUtc = null)
        {
            TimeSpan? age = widestFrontierUtc.HasValue
                ? staleness.ClockUtc - widestFrontierUtc.Value
                : staleness.Age;
            return age.HasValue
                ? age.Value.TotalMinutes.ToString("F0", CultureInfo.InvariantCulture) + " minutes"
                : "unknown span";
        }

        private static IReadOnlyList<string> SplitTerms(string? query)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                return Array.Empty<string>();
            }

            return query!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        }

        private static bool Contains(string? haystack, string needle)
        {
            return haystack != null && haystack.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static DateTime? ToUtc(DateTime? comLocal)
        {
            if (!comLocal.HasValue)
            {
                return null;
            }

            return DateTime.SpecifyKind(comLocal.Value, DateTimeKind.Local).ToUniversalTime();
        }

        private static int Clamp(int value, int min, int max)
        {
            return value < min ? min : value > max ? max : value;
        }

        private static bool IsHex(string value)
        {
            foreach (char c in value)
            {
                bool hex = (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
                if (!hex)
                {
                    return false;
                }
            }

            return true;
        }

        private sealed class CachedHit
        {
            public IndexHit? IndexHit { get; set; }

            public ComMailBrief? Live { get; set; }

            public string? LocatedEntryId { get; set; }

            public string? LocatedStoreId { get; set; }

            public string? LocatedVia { get; set; }
        }

        /// <summary>Releases the COM gateway (Outlook itself keeps running - S7/D17).</summary>
        public void Dispose()
        {
            _gateway.Dispose();
        }
    }
}
