using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;

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
        /// <para>
        /// Since 2026-08-19 it derives from the scan's OWN deadline class
        /// (<c>ComHostOperationClass.ExhaustiveScan</c>, 615 s) rather than from the shared
        /// operation deadline, so ten minutes of scanning does not become ten minutes of
        /// hang detection for <c>read</c> and <c>move_mail</c>. The measurement behind the
        /// number: on the maintainer's real profile a 60-day whole-store scan reached 3
        /// folders of 32 before the old 105 s stopped it, so the budget - not the result cap
        /// - was what bounded completeness.
        /// </para>
        /// </summary>
        public const int ExhaustiveTimeBudgetMs = ComOperationBudgets.ExhaustiveScanWorkBudgetMs;

        /// <summary>
        /// What a non-folder-scoped sweep covers, echoed in the sweep block so an agent
        /// can see its freshness coverage (soak fix 13). Kept in sync with
        /// <see cref="OutlookComSession.DefaultSweepFolderKinds"/>.
        /// </summary>
        public const string DefaultSweepScopeDescription =
            "default folders (Inbox, Sent Items, Deleted Items, Junk Email)";

        /// <summary>
        /// <see cref="SweepInfo.ScopeShape"/> when the sweep covered the arrival-path default
        /// folders of the store(s) in scope - four folders per store, and NOT their
        /// subfolders (gap E2).
        /// </summary>
        public const string SweepScopeDefaultFolders = "default_folders";

        /// <summary>
        /// <see cref="SweepInfo.ScopeShape"/> when the sweep followed a folder scope, subtree
        /// included. Deliberately the token <c>SearchScopeInfo.Shape</c> already uses for the
        /// same breadth, so an agent learns one vocabulary for "which folders".
        /// </summary>
        public const string SweepScopeFolder = "folder";

        /// <summary>
        /// <see cref="SweepInfo.ScopeShape"/> when the sweep followed a folder scope with
        /// <c>include_subfolders: false</c>. Same token as <c>SearchScopeInfo.Shape</c>'s.
        /// </summary>
        public const string SweepScopeFolderOnly = "folder_only";

        /// <summary>
        /// Which folders a sweep set out to cover, as a token software can branch on - gap
        /// E2, and the reason it is a token rather than a code or a flag is worth having in
        /// one place.
        /// <para>
        /// THE HOLE. The default set is four arrival-path folders per store and it does not
        /// descend, so mail a server-side rule files into a subfolder BEFORE the indexer
        /// reaches it is in neither tier - not in the index, which has not seen it, and not
        /// in the sweep, which does not look there. Rules filing mail is ordinary rather than
        /// exotic, which makes this the likeliest of the remaining holes to be hit. The fact
        /// was in the payload, but only as <see cref="SweepInfo.Scope"/>: an English sentence,
        /// so the only way to branch on it was to compare against a sentence, which is not a
        /// contract anybody can rely on.
        /// </para>
        /// <para>
        /// IT RAISES NO COVERAGE CODE AND DOES NOT DEGRADE THE SEARCH, on the B2 precedent
        /// and for the same arithmetic: <c>default_folders</c> is the shape of nearly every
        /// search anyone runs, and <see cref="FreshMerge.ClassifyFreshness"/> derives
        /// <c>partial</c> from the code list, so a code here would make almost every answer
        /// permanently degraded and devalue the codes that fire rarely. What is missing is
        /// also not a hole in what the sweep was ASKED to do - it is the breadth of the tier
        /// itself, which is a fact to read rather than an alarm to raise. An agent that needs
        /// a subfolder covered live passes <c>folder</c>, and the tool description says so.
        /// </para>
        /// <para>
        /// Pure, and the ONE source of both renderings: <see cref="DescribeSweepScope"/>
        /// turns the token into the sentence, so the field and the prose cannot drift into
        /// describing different breadths.
        /// </para>
        /// </summary>
        public static string ClassifySweepScope(bool folderScoped, bool includeSubfolders)
        {
            if (!folderScoped)
            {
                // The flag is meaningless without a folder: the default set is shallow by
                // construction (SweepFolder, not SweepFolderTree), so include_subfolders
                // cannot widen it and must not be allowed to imply that it did.
                return SweepScopeDefaultFolders;
            }

            return includeSubfolders ? SweepScopeFolder : SweepScopeFolderOnly;
        }

        /// <summary>
        /// The human rendering of a <see cref="ClassifySweepScope"/> token - what
        /// <see cref="SweepInfo.Scope"/> has always carried, now computed FROM the token
        /// rather than beside it. Pure, so T1 pins that every token has a sentence and that
        /// the sentences are the ones callers already read.
        /// </summary>
        public static string DescribeSweepScope(string shape)
        {
            return shape switch
            {
                SweepScopeFolder => "folder",
                SweepScopeFolderOnly => "folder (no subfolders)",
                SweepScopeDefaultFolders => DefaultSweepScopeDescription,

                // A token with no sentence would be a breadth the payload names and the prose
                // cannot explain. T1 pins that every declared token is handled, so this is
                // reachable only by a token added without its wording - say so rather than
                // returning something that reads as a real scope.
                _ => "unknown sweep scope (" + shape + ")",
            };
        }

        /// <summary>
        /// Above this many swept folders the sweep block reports the count only - the
        /// list exists to make a narrow scope legible, not to bloat every payload
        /// (section-12 compact-payload discipline).
        /// </summary>
        public const int SweptFolderListCap = 12;

        /// <summary>
        /// Stores <see cref="SweepInfo.StoresWithoutIndex"/> names before it reports a count
        /// instead (Q7b). DERIVED from <see cref="SweptFolderListCap"/> rather than a second
        /// 12: both bound a name list inside the same sweep block for the same
        /// compact-payload reason, and two independent numbers doing one job is how a pair
        /// starts drifting. Twelve is far above any real profile's unindexed-store count, so
        /// this is a guard against a pathological shape, not a routine cut.
        /// <para>
        /// Public so T1 pins it beside every other payload cap - it exists because this list
        /// was the one that had no cap at all.
        /// </para>
        /// </summary>
        public const int UnindexedStoreListCap = SweptFolderListCap;

        /// <summary>
        /// Distinct item classes the "not ordinary mail" advice names before it trails off.
        /// Small because the sentence is a heads-up, not an inventory: the per-hit
        /// <c>itemClass</c> fields are the complete answer and are right there. Public so
        /// T1 pins it with the other payload caps - a cap in prose is still a cap.
        /// </summary>
        public const int NonMailClassAdviceCap = 4;

        /// <summary>
        /// Items ONE folder may contribute to a freshness sweep.
        /// <para>
        /// The sweep reads newest-first, so hitting this cap drops the OLDEST
        /// not-yet-indexed mail in that folder - which is exactly why it is not a silent
        /// cap. The affected folders are named in the response sweep block
        /// (<see cref="SweepInfo.ItemCappedFolders"/>, carried up from
        /// <see cref="ComSweepResult.ItemCappedFolders"/> by <c>ApplySweepCounters</c>) and
        /// in an advice line that quotes this number (<see cref="DescribeSweepCoverage"/>) -
        /// the same shape the folder cap, the time budget and the top clamp already use.
        /// </para>
        /// <para>
        /// "Newest-first" is a property of the folder's TABLE, not of this cap, and Outlook
        /// can refuse to sort one (gap H2). Such a folder is named in
        /// <see cref="SweepInfo.ItemCappedFoldersUnsorted"/> instead and gets a different
        /// sentence, because there the cap kept an arbitrary slice and the claim above is
        /// simply false.
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
        private readonly DraftUpdateIntents _updateIntents = new DraftUpdateIntents();

        /// <summary>
        /// Walk state for paged exhaustive scans (F2). Per process and unpersisted, on the
        /// rule the other three of its kind already state: a restarted server never ran the
        /// earlier pages and must not claim it can continue one.
        /// </summary>
        private readonly ExhaustiveScanCursors _scanCursors = new ExhaustiveScanCursors();
        private readonly SweepCache _sweepCache = new SweepCache();
        private readonly BodyCache _bodies = new BodyCache();
        private readonly ConcurrentDictionary<string, CachedHit> _hits =
            new ConcurrentDictionary<string, CachedHit>(StringComparer.Ordinal);
        private readonly object _catalogLock = new object();
        private readonly Dictionary<string, (ComFolderPathList Paths, DateTime FetchedUtc)> _folderPaths =
            new Dictionary<string, (ComFolderPathList, DateTime)>(StringComparer.OrdinalIgnoreCase);

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
            : this(gateway, sendTokens, null)
        {
        }

        /// <summary>
        /// Creates the service over an explicit index client. The second test seam beside
        /// <see cref="IComGateway"/>, and it exists for a defect that shipped BECAUSE it was
        /// missing: every test until 2026-08-18 ran against an index that knew about every
        /// store, so nothing exercised a profile whose store catalog is empty or short - the
        /// ordinary shape of a PST-only machine, a fresh install, or one where indexing is
        /// off or still building - and a store-scoped search failed outright on all of them.
        /// <para>
        /// <paramref name="indexClient"/> null keeps production behaviour: the real provider,
        /// attached lazily on first use so nothing touches Windows Search until a search does.
        /// </para>
        /// </summary>
        public MailService(IComGateway gateway, SendConfirmationTokens? sendTokens, IIndexClient? indexClient)
        {
            _gateway = gateway ?? throw new ArgumentNullException(nameof(gateway));
            _sendTokens = sendTokens ?? new SendConfirmationTokens();
            _index = indexClient == null
                ? new Lazy<IndexSearchService>(
                    () => IndexSearchService.CreateDefault(out _providerReport),
                    System.Threading.LazyThreadSafetyMode.ExecutionAndPublication)
                : new Lazy<IndexSearchService>(
                    () => new IndexSearchService(indexClient),
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
        /// Shorter than an ordinary operation, because the sweep is an ENHANCEMENT: search
        /// already has its indexed answer in hand before the sweep runs, and the tool's own
        /// contract calls search "sub-second and cheap". Measured healthy on this machine at
        /// 0.5-6 s; measured against a wedged Outlook it spent the full operation budget
        /// before degrading, which made every search feel broken even though the answer was
        /// already computed and waiting.
        /// </para>
        /// <para>
        /// It is a budget for the SWEEP, and the sweep call passes allowConnectFloor so the
        /// COM host may still add its cold-start connect allowance on a fresh host. Without
        /// that the very first search had to fit the COM attach AND the whole sweep into
        /// this budget - on a machine where attaching to a large OST takes longer than that
        /// (the reason ConnectDeadlineMs is what it is at all) the sweep could never
        /// succeed: every attempt timed out, killed the host, bumped the restart count and
        /// blamed the sweep.
        /// </para>
        /// <para>
        /// RAISED FROM 30 000 TO 180 000 on 2026-08-19, and this is a measurement, not a
        /// preference. Measured on a purpose-built corpus - one PST outside the local index,
        /// 20 000 items across the four arrival-path folders with real received dates, 1 612
        /// of them inside the seven-day fallback window so the per-folder cap engages: four
        /// sweeps of that ONE store took 13.6 s, 11.8 s, 10.7 s and 11.9 s, i.e. about 12 s
        /// per store with the cap engaged. The maintainer's profile mounts FIVE stores and
        /// the sweep covers four folders in each, so the extrapolation is ~60 s - against a
        /// 30 s budget. That is the direct explanation for the sweep timeout observed on
        /// their real profile, where the supervisor then killed and replaced the COM host,
        /// and it agrees with the earlier whole-store 7-day figure of 36.6 s there.
        /// </para>
        /// <para>
        /// 180 000 is 3x that measured extrapolation, and the margin is headroom rather than
        /// luxury: the corpus is a fast LOCAL PST, and the same per-item work against
        /// Exchange is slower. The other half of the fix is that expiry is no longer fatal -
        /// the sweep stops at a folder boundary and returns what it covered
        /// (<see cref="SweepWorkBudgetMs"/>), and an expiring caller budget no longer counts
        /// as evidence that Outlook is unresponsive.
        /// </para>
        /// <para>
        /// It is COUPLED to <c>OutlookComSession.SweepBodyBytesBudget</c>. The 432 KB frame
        /// high-water previously measured on the real profile was bounded by the old 30 s
        /// timeout, not by any item cap; the same corpus measured 10.2 MB from one store's
        /// sweep once the sweep was allowed to finish. Giving the sweep time is what lets it
        /// build frames large enough for the body budget to matter.
        /// </para>
        /// </summary>
        public const int SweepBudgetMs = 180_000;

        /// <summary>
        /// The sweep's INNER budget - the one the COM child measures against its own clock
        /// and stops on gracefully. Derived from <see cref="SweepBudgetMs"/> exactly as the
        /// exhaustive scan's is derived from its class deadline, and for the same reason:
        /// an inner budget equal to its outer one can never degrade, because the outer
        /// watchdog fires while the inner walk is still serializing its answer.
        /// <para>
        /// WHAT IT BUYS (maintainer decision (d), 2026-08-19). Before it, the whole-profile
        /// sweep had no budget of its own at all - only the outer gateway deadline - so a
        /// sweep that ran long produced a <c>TimeoutException</c>, the supervisor concluded
        /// the host was wedged, the child was killed and replaced, and every folder the
        /// sweep HAD covered was thrown away. Observed on the maintainer's real profile.
        /// Now the walk checks this budget at each store and each folder boundary, stops,
        /// and returns the folders it did cover with <c>sweep_time_budget</c> in
        /// <c>coverageGaps</c> - the same discipline the exhaustive scan has had since
        /// 2026-08-18.
        /// </para>
        /// <para>
        /// The clock starts inside the COM child, after the session is connected, so a cold
        /// start is paid out of the outer budget's headroom rather than out of this one.
        /// </para>
        /// </summary>
        public const int SweepWorkBudgetMs = SweepBudgetMs - ComOperationBudgets.ResultReturnHeadroomMs;

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
        /// <para>
        /// RAISED FROM 15 TO 60 on 2026-08-19. The measurement it was sized against is
        /// unchanged and so is its meaning - what changed is the judgement about what to do
        /// when it is exceeded. On a ~50 GB profile a saturated or still-building index is
        /// an ordinary state rather than a fault, and a search that gives up after 15 s
        /// hands back a degraded answer the caller then has to work around; one that waits
        /// a minute hands back the real one. The headroom over the healthy measurement is
        /// now ~110x, which is the point: this is a backstop against an indexer that has
        /// stopped answering at all, not a service-level target.
        /// </para>
        /// </summary>
        public const int SearchIndexTimeoutSeconds = 60;

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
        /// <para>
        /// IT USED TO EQUAL <c>ComOperationBudgets.OperationDeadlineMs</c>, and an aggregate
        /// equal to its own unit of work is not an aggregate. The check runs BEFORE each
        /// item, and each item was a fresh gateway call carrying a full operation deadline
        /// of its own, so a batch sitting at 119.9 s could start one more item and run to
        /// ~240 s. T1 pinned the equality, which meant the test enforced the flaw. Two
        /// things fix it: the value is now strictly below the operation deadline, and each
        /// item is dispatched with what is LEFT of this budget
        /// (<see cref="MinimumItemBudgetMs"/> is the floor below which the item is reported
        /// as not attempted instead), so the batch is bounded by this number plus one
        /// result-return rather than by this number plus a whole extra deadline.
        /// </para>
        /// <para>
        /// 240 000 is 80% of the operation deadline: a full 50-id batch stays comfortably
        /// servable on a slow profile - the maintainer's instruction is that finishing
        /// slowly beats giving up - while leaving the hang detector strictly above it, so
        /// "the batch ran long" and "Outlook stopped answering" remain different events
        /// with different reports.
        /// </para>
        /// </summary>
        public const int MoveBatchBudgetMs = 240_000;

        /// <summary>
        /// Least budget one move/archive item is dispatched with. Below this the item is
        /// reported as not attempted rather than started: a sub-second deadline would be
        /// refused by the COM host's own dispatch floor anyway, and the refusal would
        /// surface as a bare timeout instead of as the legible per-item "re-issue the rest
        /// as a smaller batch" the batch short circuit already produces.
        /// <para>
        /// Public so T1 pins it beside the budget it divides, exactly like every other value
        /// in this block.
        /// </para>
        /// </summary>
        public const int MinimumItemBudgetMs = 1_000;

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

            // A continuation handle only means anything to the scan that issued it. Ignoring
            // one here would be the silent half of the failure the refusals exist to prevent:
            // the caller believes they are continuing a walk, and gets a fresh indexed search
            // that answers a different question and says nothing about it.
            if (!string.IsNullOrWhiteSpace(request.ResumeToken))
            {
                throw new ArgumentException(
                    "resume_token continues an EXHAUSTIVE scan and has no meaning without it. Pass exhaustive:true "
                    + "with the same arguments the scan started with, or drop resume_token.",
                    nameof(request));
            }

            FolderScopeResolution? folderScope = null;
            if (request.Store != null)
            {
                folderScope = ResolveFolderScope(request.Store, request.Folder, request.IncludeSubfolders);
            }

            // A store the PROFILE has but the index cannot address (a PST, an archive-only
            // data file, indexing off or still building) has no SCOPE to query with, and the
            // three things one could do about that are not equivalent:
            //
            //   REFUSE   - what this used to do, and wrong: the sweep tier can answer, and
            //              an unscoped search on the same profile already answers.
            //   WIDEN    - what 'thread' does with its store, and wrong HERE: thread's store
            //              only makes a lookup faster (the conversation is pinned by id
            //              either way), whereas search's 'store' decides WHICH MAIL MAY COME
            //              BACK. Widening it returns another account's mail under a scope the
            //              caller chose, which is a wrong answer rather than a slow one.
            //   PROCEED  - skip the index tier, sweep the store, report both. This.
            //
            // The reporting needs nothing new: no_index_frontier, sweep.storesWithoutIndex,
            // degraded and freshness:"partial" are exactly this state, and they already fire
            // on the unscoped path for the same store.
            bool indexAddressable = folderScope == null || folderScope.IndexAddressable;

            IndexQuery query = new IndexQuery
            {
                Scope = folderScope?.Scope,
                FolderPathsAnyOf = folderScope?.FolderPaths,
                Terms = terms.Count > 0 ? terms : null,
                SearchIn = request.SearchIn,
                // Message rows of EVERY item class in both message-bearing shapes (gap B3):
                // meeting requests, NDRs, receipts and posts are mail a user asks about by
                // name, and the freshness sweep beside this query has always returned them.
                Kinds = request.AttachmentHitsOnly
                    ? KindFilter.AttachmentsOnly
                    : request.IncludeAttachmentHits ? KindFilter.MessagesAndAttachments : KindFilter.MessagesOnly,
                SenderContains = request.From,
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

            IndexSearchResult indexResult = indexAddressable
                ? _index.Value.Search(query, SearchIndexTimeoutSeconds)
                : IndexSearchResult.NotQueried();

            // The frontier is measured over the SCOPE being searched, not over the profile
            // (StalenessScopeFor): it sets this search's sweep window, and a busy store's
            // frontier would otherwise pin a quiet store's window to the last few minutes
            // while that store's own index lagged by hours.
            //
            // An unaddressable store is not probed AT ALL, and falling back to the profile
            // probe would be the same defect one level down: the frontier of the stores that
            // ARE indexed says nothing about this one, and handing it their window would give
            // the store that needs the widest sweep on the profile the narrowest one - and
            // clear indexFrontierMissing while doing it. A no-rows report is the truth, and it
            // is what routes this search to the fallback window and the honest flags.
            IndexStalenessReport staleness = indexAddressable
                ? _index.Value.GetStaleness(StalenessScopeFor(folderScope), SearchIndexTimeoutSeconds)
                : IndexStalenessReport.NoRows();

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

            // Hoisted out of the sweep block below because the staleness block needs it too
            // (Q7a): it is the EARLIEST per-store frontier the sweep planner measured, which
            // is the answer to "how far behind is the worst store" and the number the
            // freshness advice has already been quoting.
            DateTime? oldestPerStoreFrontierUtc = null;

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
                oldestPerStoreFrontierUtc = widestFrontierUtc;

                // Before anything reads the list: the coverage advice below JOINS these
                // names into a sentence, so an uncapped list would reach the caller twice
                // over (Q7b). Applied here rather than inside the sweep because the sweep
                // adds to the list in two places and this is the first point past both.
                ApplyUnindexedStoreCap(sweep);

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
                string? attachmentTextAdvice = DescribeAttachmentTextGap(sweep);
                if (attachmentTextAdvice != null)
                {
                    advice.Add(attachmentTextAdvice);
                }

                string? unnamedStoreAdvice = DescribeUnnamedStores(sweep.StoresUnnamed);
                if (unnamedStoreAdvice != null)
                {
                    advice.Add(unnamedStoreAdvice);
                }
            }

            // Gap G4: a scope hole, so it is said whether or not a sweep ran.
            string? scopeAdvice = DescribeTruncatedFolderNames(folderScope);
            if (scopeAdvice != null)
            {
                advice.Add(scopeAdvice);
            }

            // Gap G5. Judged on the INDEX tier's own row count, not on the merged answer:
            // the question is whether the index can see the folder at all, and a single
            // swept item - which comes from COM, where the folder resolved fine - used to
            // silence it completely.
            bool? folderNotIndexed = DescribeFolderMissingFromIndex(
                advice, folderScope, request, indexResult.Hits.Count);

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

            SortForOrder(summaries, request.OrderBySizeDescending);
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

            // After the list is sorted and trimmed: it describes what the caller GETS, not
            // what the tiers passed through.
            string? nonMailAdvice = DescribeNonMailHits(summaries);
            if (nonMailAdvice != null)
            {
                advice.Add(nonMailAdvice);
            }

            if (indexResult.CandidatesExhausted)
            {
                // The index tier admits rows in code over an over-fetched candidate list, so
                // running out of candidates - or failing the follow-up query that recovers
                // rows the result ordering may have hidden - is how it could hide matches.
                // Say so (D42), in a FIELD as well as here since 2026-08-18 (gap G6).
                advice.Add("The index tier could not confirm this list holds every match: its candidate rows ran out "
                    + "while filtering, or the follow-up query that would have proved otherwise failed. "
                    + "Narrow with store/folder/after, or lower top.");
            }

            // Say the live check's outcome in a FIELD, not only in prose: a result that
            // looks complete but silently lags recent mail is the one failure here that can
            // mislead rather than merely inconvenience. Three states, because a sweep that
            // ran and covered part of its scope is neither of the other two - it did run, so
            // it is not "index-only", and it left mail unchecked, so it is not "live".
            string freshness = FreshMerge.ClassifyFreshness(sweep);

            // Gap G4. A delegate folder scope built from a folder walk that hit its own cap
            // covers less than it was asked to, which is what this flag has always meant -
            // so it is raised here too, by a fact about the SCOPE rather than about a tier.
            // freshness stays what the tiers measured: this is not a lag behind the index,
            // and folding it into "partial" would make that word mean two different holes
            // with two different remedies.
            bool scopeTruncated = folderScope != null && folderScope.FolderTreeTruncated;

            return new SearchOutcome
            {
                Hits = summaries,
                Truncated = truncated,
                Degraded = freshness != FreshMerge.FreshnessLive || scopeTruncated ? true : (bool?)null,
                Freshness = freshness,
                IndexElapsedMs = indexResult.ElapsedMilliseconds,
                Sweep = sweep,
                Index = new IndexTierInfo
                {
                    RowsScanned = indexResult.RowsScanned,
                    RowsDropped = indexResult.RowsDropped,
                    CandidatesExhausted = indexResult.CandidatesExhausted ? true : (bool?)null,
                    StoreNotIndexed = indexAddressable ? (bool?)null : true,
                    FolderNotIndexed = folderNotIndexed,
                },
                Scope = DescribeSearchScope(folderScope, request),
                Staleness = new StalenessInfo
                {
                    NewestIndexedUtc = staleness.NewestIndexedReceivedUtc,
                    OldestStoreFrontierUtc = OldestStoreFrontier(
                        request.Store != null, staleness.NewestIndexedReceivedUtc, oldestPerStoreFrontierUtc),
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
        /// The value of <c>staleness.oldestStoreFrontierUtc</c> (Q7a): how far behind the
        /// WORST store in this search's scope is, or null when no per-store frontier was
        /// measured.
        /// <para>
        /// A store-scoped search has exactly one store in scope, so its own frontier
        /// (<paramref name="scopeFrontierUtc"/>) is both the newest and the oldest, and the
        /// two staleness fields agree by construction. That is deliberate rather than
        /// redundant: a caller reading only the new field gets a true answer on every search
        /// shape instead of having to know which shape it asked for.
        /// </para>
        /// <para>
        /// An unscoped search reports the earliest of the per-store frontiers the sweep
        /// planner measured, and reports NOTHING when it measured none - a profile-wide
        /// maximum is not an "oldest store" and substituting it would put a number in the
        /// field that no store's index actually stands at. Absence therefore means "not
        /// measured", which is why this is a nullable answer rather than a fallback chain.
        /// </para>
        /// <para>
        /// Pure and public so T1 pins all three branches: the alternatives (always the
        /// maximum, or the maximum as a fallback) are one line away and neither is
        /// distinguishable from this at the call site.
        /// </para>
        /// </summary>
        public static DateTime? OldestStoreFrontier(
            bool storeScoped, DateTime? scopeFrontierUtc, DateTime? oldestPerStoreFrontierUtc)
        {
            return storeScoped ? scopeFrontierUtc : oldestPerStoreFrontierUtc;
        }

        /// <summary>
        /// Applies <see cref="UnindexedStoreListCap"/> to <see cref="SweepInfo.StoresWithoutIndex"/>
        /// and reports the cut (Q7b).
        /// <para>
        /// This list was the only one in the server with no bound - in the payload AND in
        /// the <c>no_index_frontier</c> advice sentence, which joins the names into prose -
        /// so a profile of many unindexed data files would have printed its whole store list
        /// twice per search. That is an omission rather than a design choice: every other
        /// cap here is named, applied and reported.
        /// </para>
        /// <para>
        /// The list is TRUNCATED rather than dropped, which is where it differs from
        /// <see cref="SweepInfo.Folders"/>: a swept-folder list is a legibility aid and is
        /// worth nothing in part, whereas each unindexed store NAME is separately actionable
        /// (search it with <c>exhaustive:true</c>), so the first few keep their whole value.
        /// The order is the order the stores were found, which is the sweep's store order -
        /// stable, and not a ranking, so the flags below are the only honest way to say
        /// there are more.
        /// </para>
        /// <para>
        /// Pure and public: T1 pins the cap, the flags and the advice sentence together,
        /// over a payload shape that needs a profile full of unindexed PSTs to produce.
        /// </para>
        /// </summary>
        public static void ApplyUnindexedStoreCap(SweepInfo sweep)
        {
            if (sweep == null)
            {
                throw new ArgumentNullException(nameof(sweep));
            }

            sweep.StoresWithoutIndex = CapUnindexedStoreList(
                sweep.StoresWithoutIndex, out int total, out bool truncated);
            sweep.StoresWithoutIndexTruncated = truncated ? true : (bool?)null;
            sweep.StoresWithoutIndexTotal = truncated ? total : (int?)null;
        }

        /// <summary>
        /// How a list of stores the index holds nothing for is SAID, wherever it is said -
        /// the sweep's <c>no_index_frontier</c> sentence and the conversation walk's
        /// <c>unindexed_store</c> one. Null for an empty list, so each caller supplies its
        /// own wording for "we know it happened and cannot name the stores".
        /// <para>
        /// The truncation clause is the substance. A capped list read aloud as "store A,
        /// store B" claims to be the whole set, which is a quieter lie than the unbounded
        /// list the cap replaced (Q7b) - so the remainder is counted, the cap is quoted from
        /// its constant rather than restated, and the payload field carrying the true total
        /// is named. That field differs per block, which is the one thing
        /// <paramref name="totalFieldPath"/> exists for.
        /// </para>
        /// </summary>
        public static string? DescribeUnindexedStoreList(
            IReadOnlyList<string>? stores, bool truncated, int? total, string totalFieldPath)
        {
            if (stores == null || stores.Count == 0)
            {
                return null;
            }

            string named = "store(s) " + string.Join(", ", stores);
            if (!truncated)
            {
                return named;
            }

            return named + " and " + ((total ?? stores.Count) - stores.Count).ToString(CultureInfo.InvariantCulture)
                + " more (list capped at " + UnindexedStoreListCap.ToString(CultureInfo.InvariantCulture)
                + "; " + totalFieldPath + " is the true count)";
        }

        /// <summary>
        /// The list half of <see cref="ApplyUnindexedStoreCap"/>, pure and shared, so the
        /// sweep block and the conversation walk's block
        /// (<see cref="ThreadLiveInfo.StoresWithoutIndex"/>) cut the same list the same way
        /// at the same number. Two blocks applying one cap by two copies of the code is how
        /// the pair drifts.
        /// </summary>
        public static IReadOnlyList<string>? CapUnindexedStoreList(
            IReadOnlyList<string>? stores, out int total, out bool truncated)
        {
            total = stores?.Count ?? 0;
            truncated = total > UnindexedStoreListCap;
            if (!truncated)
            {
                return stores;
            }

            List<string> capped = new List<string>(UnindexedStoreListCap);
            for (int i = 0; i < UnindexedStoreListCap; i++)
            {
                capped.Add(stores![i]);
            }

            return capped;
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
        /// (<see cref="DescribeFolderMissingFromIndex"/>): it judges the INDEX tier's own
        /// row count, which is a different question from what the merged answer holds.
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
        /// The non-silent zero-row guard (v3.MD constraint C7, widened by gap G5). Two TOP-1
        /// probes: first "does this folder bound match ANY row" - so a folder that merely
        /// holds no match stays quiet, which the one-probe form would not - then "does the
        /// store". Rows in the store but none for the folder means the PATH did not resolve
        /// in the index, which is the failure mode that hid the delegate defect.
        /// <para>
        /// Returns the value of <c>index.folderNotIndexed</c>: <c>true</c> when the index
        /// tier holds nothing addressable by this folder bound while it holds rows for the
        /// store, null in every other case - not probed, probe failed, or the folder
        /// resolves. The advice sentence is emitted from the same branch that decides the
        /// field, so the payload and the prose are one decision rather than two.
        /// </para>
        /// <para>
        /// WHAT CHANGED, and it is the whole of gap G5. This ran only on a FULLY EMPTY
        /// merged answer - index rows plus swept items - so a single item the freshness
        /// sweep happened to return silenced it outright. That condition is not defensible:
        /// the swept item comes from COM, where the folder resolved perfectly well, and it
        /// says nothing whatever about whether the INDEX can address that folder. The state
        /// it hid is the ordinary one for a renamed or localized folder - the sweep answers
        /// for the last few minutes and the index answers for nothing at all - so the answer
        /// looked like a thin result set rather than a folder scope that matched no indexed
        /// row. The gate is now the index tier's own row count, which is exactly the
        /// question the two probes go on to ask.
        /// </para>
        /// <para>
        /// COST: two TOP-1 statements, on folder-scoped searches whose index tier returned
        /// nothing. That is more often than "the whole answer was empty" and still only on
        /// searches the index did not answer, which is the trade this project's standing
        /// rule already settles - completeness outranks speed.
        /// </para>
        /// <para>
        /// It deliberately does NOT set <c>degraded</c>. The probe establishes that the
        /// index holds no row under this folder bound; it cannot distinguish "the path does
        /// not resolve" from "the folder is genuinely new and everything in it is still
        /// only in the sweep window", and degrading the second would be a false alarm on
        /// every freshly created folder - the same cry-wolf trap <c>foldersAbsent</c> exists
        /// to avoid one level up. The fact is reported instead, in a field and in a
        /// sentence, which is what gap G5 asks for.
        /// </para>
        /// </summary>
        private bool? DescribeFolderMissingFromIndex(
            List<string> advice, FolderScopeResolution? folderScope, SearchRequest request, int indexRowCount)
        {
            // Not for a store the index cannot address: the guard's whole question is "did
            // the FOLDER bound match nothing while the STORE matched something", and with no
            // index tier there is no folder bound to have failed. Both probes would also run
            // on a null scope, i.e. profile-wide, and answer about other stores entirely.
            if (indexRowCount > 0
                || folderScope == null
                || folderScope.RequestedFolder == null
                || !folderScope.IndexAddressable)
            {
                return null;
            }

            try
            {
                if (_index.Value.FolderScopeHasAnyItem(folderScope.Scope, folderScope.FolderPaths, SearchIndexTimeoutSeconds))
                {
                    return null;
                }

                if (!_index.Value.FolderScopeHasAnyItem(folderScope.StoreScope, null, SearchIndexTimeoutSeconds))
                {
                    // The STORE has no indexed rows either, so this says nothing about the
                    // folder. That state has its own reporting (index.storeNotIndexed,
                    // sweep.storesWithoutIndex, no_index_frontier) and claiming the folder
                    // failed to resolve here would blame the path for the store's gap.
                    return null;
                }

                advice.Add(FolderScopeResolver.DescribeUnresolvedFolder(request.Folder, request.Store!));
                return true;
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                // The guard is a diagnostic; a failing probe must never fail the search -
                // and must not answer either, which is what null means here.
                return null;
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
                            + (DescribeUnindexedStoreList(
                                sweep.StoresWithoutIndex,
                                sweep.StoresWithoutIndexTruncated == true,
                                sweep.StoresWithoutIndexTotal,
                                "sweep.storesWithoutIndexTotal") ?? "this profile")
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

                    case FreshMerge.GapSweepBudget:
                        advice.Add("Freshness sweep ran out of its "
                            + (SweepWorkBudgetMs / 1000).ToString(CultureInfo.InvariantCulture)
                            + " s budget after " + sweep.FoldersSwept.ToString(CultureInfo.InvariantCulture)
                            + " folder(s) and stopped there, so stores or folders it had not reached yet have no freshness "
                            + "coverage - index results still cover them, but brand-new mail there may be missing. It "
                            + "returned what it did cover rather than failing. Name a 'store' (and a 'folder' if you know "
                            + "one) to give the sweep less ground, or re-run: the window it could not finish is still open "
                            + "next time.");
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
                        // Named from the SORTED subset, not from the whole capped set: the
                        // clause after it asserts newest-first, and over a folder Outlook
                        // would not sort that assertion is false (gap H2).
                        advice.Add("Freshness sweep hit its per-folder cap of "
                            + SweepPerFolderCap.ToString(CultureInfo.InvariantCulture) + " items in: "
                            + string.Join(", ", FreshMerge.SortedItemCappedFolders(sweep))
                            + ". It reads newest-first, so the OLDEST not-yet-indexed mail in those folders is not covered - "
                            + "narrow the window with 'after' or search those folders directly.");
                        break;

                    case FreshMerge.GapItemCapUnsorted:
                        advice.Add("Freshness sweep hit its per-folder cap of "
                            + SweepPerFolderCap.ToString(CultureInfo.InvariantCulture) + " items in: "
                            + string.Join(", ", sweep.ItemCappedFoldersUnsorted ?? Array.Empty<string>())
                            + " - and Outlook would not sort those folders by received time, so the cap kept an ARBITRARY "
                            + SweepPerFolderCap.ToString(CultureInfo.InvariantCulture)
                            + " of the window rather than the newest. WHICH not-yet-indexed mail is missing there is "
                            + "unknown: do not assume it is the oldest. Narrow the window with 'after' until the folder "
                            + "fits under the cap, or read the folder with exhaustive:true.");
                        break;

                    case FreshMerge.GapRowsUnreadable:
                        advice.Add("Freshness sweep could not open "
                            + sweep.RowsUnreadable.ToString(CultureInfo.InvariantCulture)
                            + " item(s) it found in the swept folders, so mail that arrived there in the last " + indexAge
                            + " may be missing even though the folder itself was read. Usually a transient Outlook "
                            + "state - retry the search, or use exhaustive:true for that folder.");
                        break;

                    case FreshMerge.GapFilterUnreadable:
                        advice.Add("Freshness sweep dropped "
                            + sweep.ItemsFilterUnreadable.ToString(CultureInfo.InvariantCulture)
                            + " newly arrived item(s) because Outlook reported no usable value for "
                            + (sweep.FiltersUnevaluated != null && sweep.FiltersUnevaluated.Count > 0
                                ? string.Join(" / ", sweep.FiltersUnevaluated)
                                : "one of the filters you passed")
                            + " on them, so they could neither be matched nor ruled out. Re-run without that filter to "
                            + "see them and judge them yourself; index results are unaffected.");
                        break;

                    case FreshMerge.GapBodyCap:
                        // Both halves are read from the payload rather than restated: the two
                        // counts, and which bound did the cutting. The sentence deliberately
                        // says "may be" - see FreshMerge.GapBodyCap for why nothing here can
                        // say more than that.
                        advice.Add("Freshness sweep matched only the first "
                            + OutlookComSession.SweepBodyCharsCap.ToString(CultureInfo.InvariantCulture)
                            + " characters of "
                            + (sweep.ItemsBodyCapped ?? 0).ToString(CultureInfo.InvariantCulture)
                            + " just-arrived item(s) - a mail body is capped on its way back from Outlook so one "
                            + "answer cannot outgrow what the connection can carry. "
                            + (sweep.ItemsBodyCappedUnmatched ?? 0).ToString(CultureInfo.InvariantCulture)
                            + " of them did not match your terms, so a term appearing only in the part that was "
                            + "cut would have been missed and those items MAY be hits this answer does not have. "
                            + (sweep.BodyBudgetExhausted == true
                                ? "The sweep as a whole carried more body text than one answer holds, so later items "
                                  + "kept little or none of theirs: narrow the search with store, folder or 'after' so "
                                  + "fewer items are swept."
                                : "These are individually enormous mails; read one with 'read' (its body pages via "
                                  + "body_offset) to search it in full.")
                            + " Anything already indexed is matched in full and is unaffected.");
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

        /// <summary>
        /// The one sentence for <see cref="SweepInfo.AttachmentTextCovered"/>, or null when
        /// there is nothing this gap could be hiding (gap B2).
        /// <para>
        /// Generated FROM the field, so the prose cannot claim a coverage hole the payload
        /// does not carry, nor stay quiet about one it does - the same pairing the coverage
        /// codes and <see cref="DescribeSweepCoverage"/> already have.
        /// </para>
        /// <para>
        /// Two further conditions, and both are about whether the hole is REAL rather than
        /// merely structural. The sweep must have RUN: a sweep that was refused or could not
        /// run has already said the whole window is uncovered, and repeating a subset of
        /// that is noise. And it must have SEEN something: <see cref="SweepInfo.ItemsSeen"/>
        /// counts the items in the freshness window before term filtering, so zero means no
        /// mail arrived after the index frontier at all - and a window with no mail in it
        /// cannot be hiding a mail whose match is in an attachment. That is the same
        /// reasoning <c>foldersAbsent</c> uses, and it is emphatically NOT the "suppress it
        /// when there is a hit" shape gap G5 is about: this suppresses when there is
        /// provably nothing to report, not when the answer happens to look full.
        /// </para>
        /// <para>
        /// Public and pure so T1 can pin all four states without a mailbox.
        /// </para>
        /// </summary>
        public static string? DescribeAttachmentTextGap(SweepInfo sweep)
        {
            if (sweep == null)
            {
                throw new ArgumentNullException(nameof(sweep));
            }

            if (sweep.AttachmentTextCovered != false || !sweep.Performed || sweep.ItemsSeen <= 0)
            {
                return null;
            }

            return "Attachment CONTENT is matched by the index tier alone: the freshness sweep reads subject and body "
                + "through Outlook and never opens an attachment. It found "
                + sweep.ItemsSeen.ToString(CultureInfo.InvariantCulture)
                + " item(s) newer than the index, and a query term appearing ONLY inside an attachment of one of those "
                + "was not matched - so these results are complete for subject and body but NOT for attachment text in "
                + "that window, even though freshness reads 'live'. The gap closes by itself once those items are "
                + "indexed; to check now, read the candidates or re-run with attachment_hits_only once they are.";
        }

        /// <summary>
        /// One advice sentence per EXHAUSTIVE coverage code, from the code list alone - the
        /// third tier's <see cref="DescribeSweepCoverage"/>, pure and public for the same
        /// reason (T1 pins that every code declared has prose, over a payload block that
        /// only a real multi-GB store can produce).
        /// <para>
        /// The three sentences that were emitted inline moved here unchanged in wording. The
        /// change is where they come from: they used to be written from the scan counters by
        /// three separate <c>if</c>s, next to <c>freshness</c> computed from a fourth copy of
        /// the same conditions. Now the codes are the single decision and both renderings
        /// read them, so a hole cannot appear in the prose and not in the fields, or the
        /// other way round.
        /// </para>
        /// </summary>
        public static IReadOnlyList<string> DescribeExhaustiveCoverage(
            ExhaustiveInfo exhaustive, int top, int summaryCount = 0)
        {
            if (exhaustive == null)
            {
                throw new ArgumentNullException(nameof(exhaustive));
            }

            List<string> advice = new List<string>();
            foreach (string gap in exhaustive.CoverageGaps ?? Array.Empty<string>())
            {
                switch (gap)
                {
                    case FreshMerge.ScanGapTimeBudget:
                        advice.Add("The " + (ExhaustiveTimeBudgetMs / 1000).ToString(CultureInfo.InvariantCulture)
                            + " s time budget stopped the scan after "
                            + exhaustive.FoldersScanned.ToString(CultureInfo.InvariantCulture)
                            + " folder(s) - results are partial. "
                            + ResumeRemedy(exhaustive, "Nothing is cheaper than continuing here: the budget stopped "
                                + "this page mid-walk, so re-running the same search would re-walk the same folders "
                                + "and stop in the same place.")
                            + " Narrowing the folder/date bounds, or include_subfolders:false to scan just the named "
                            + "folder, makes each page cover more of what you actually want.");
                        break;

                    case FreshMerge.ScanGapResultCap:
                        advice.Add("Result cap (" + top.ToString(CultureInfo.InvariantCulture)
                            + ") stopped the scan - results may be incomplete. "
                            + ResumeRemedy(exhaustive, "Raising top will not help beyond "
                                + SearchTopCap.ToString(CultureInfo.InvariantCulture) + ".")
                            + " Narrowing the folder/date bounds is the CHEAPER remedy when you know what you are "
                            + "looking for: it stops the walk having to reach these matches at all.");
                        break;

                    case FreshMerge.ScanGapFoldersSkipped:
                        advice.Add("The scan SKIPPED " + exhaustive.FoldersSkipped.ToString(CultureInfo.InvariantCulture)
                            + " folder(s) Outlook would not filter or enumerate (of "
                            + (exhaustive.FoldersScanned + exhaustive.FoldersSkipped).ToString(CultureInfo.InvariantCulture)
                            + " reached) - mail in them is NOT covered by these results.");
                        break;

                    case FreshMerge.ScanGapDepthLimit:
                        advice.Add("The scan refused folders deeper than "
                            + Com.OutlookComSession.FolderWalkDepthGuard.ToString(CultureInfo.InvariantCulture)
                            + " levels and never opened them, so mail in them is NOT covered by these results. A real "
                            + "mailbox is nowhere near that deep, so this points at a damaged or looping folder tree "
                            + "rather than at the bound being too low - list_folders on this store shows where. Scan "
                            + "the subtree directly with folder set to a path inside it.");
                        break;

                    case FreshMerge.ScanGapRowsUnreadable:
                        advice.Add("The scan matched " + exhaustive.RowsUnreadable.ToString(CultureInfo.InvariantCulture)
                            + " item(s) it could not then open or identify, so they are missing from these results even "
                            + "though the folders holding them were scanned. Usually a transient Outlook state - retry, "
                            + "or narrow the scan to the folder in question.");
                        break;

                    case FreshMerge.ScanGapFilterUnreadable:
                        advice.Add("The scan dropped "
                            + exhaustive.ItemsFilterUnreadable.ToString(CultureInfo.InvariantCulture)
                            + " matching item(s) because Outlook reported no usable value for "
                            + (exhaustive.FiltersUnevaluated != null && exhaustive.FiltersUnevaluated.Count > 0
                                ? string.Join(" / ", exhaustive.FiltersUnevaluated)
                                : "one of the filters you passed")
                            + " on them, so they could neither be matched nor ruled out. Re-run without that filter to "
                            + "see them and judge them yourself.");
                        break;

                    case FreshMerge.ScanGapPostCapFilter:
                        advice.Add("The result cap counted CANDIDATES, not results: the scan stopped after "
                            + top.ToString(CultureInfo.InvariantCulture) + " item(s) matched its subject/body filter, "
                            + "and your "
                            + string.Join(" / ", exhaustive.PostCapFilters ?? Array.Empty<string>())
                            + " filter was applied only afterwards, discarding "
                            + exhaustive.ItemsFilteredOut.ToString(CultureInfo.InvariantCulture)
                            + " of them. So the "
                            + summaryCount.ToString(CultureInfo.InvariantCulture)
                            + " result(s) here are NOT the first "
                            + top.ToString(CultureInfo.InvariantCulture)
                            + " matches of your whole query, and far more may exist further into the scan. Put the "
                            + "narrowing into the scan itself - a tighter folder or 'after' bound - rather than into a "
                            + "filter that runs after the cap.");
                        break;

                    case FreshMerge.ScanGapResumed:
                        advice.Add("These results CONTINUE an earlier exhaustive scan (page "
                            + (exhaustive.Position?.Page ?? 0).ToString(CultureInfo.InvariantCulture)
                            + "; " + (exhaustive.ItemsReturnedTotal ?? 0).ToString(CultureInfo.InvariantCulture)
                            + " item(s) returned across the chain so far, "
                            + (exhaustive.Position == null
                                ? "folder progress unknown"
                                : exhaustive.Position.FoldersDone.ToString(CultureInfo.InvariantCulture) + " of "
                                    + exhaustive.Position.FoldersTotal.ToString(CultureInfo.InvariantCulture)
                                    + " folder(s) finished")
                            + "). This page is NOT the whole answer, and top counts per page - every page you pull "
                            + "adds to the total above, so decide deliberately when to stop rather than paging until "
                            + "the token runs out.");
                        break;

                    case FreshMerge.ScanGapTreeChanged:
                        advice.Add("Folders were added, moved, renamed or removed inside the scope between pages of "
                            + "this scan ("
                            + (exhaustive.TreeChangedFolders ?? 0).ToString(CultureInfo.InvariantCulture)
                            + "). Ones that appeared were SCANNED rather than skipped and ones that left had already "
                            + "been covered, so ordinarily nothing is missing"
                            + (exhaustive.CursorFolderMissing == true
                                ? " - EXCEPT that the folder this scan had stopped inside is gone, so whatever it "
                                    + "still held is NOT covered by these results. Check list_folders for where it "
                                    + "went and scan that path directly."
                                : ". Mail that arrived in an already-finished folder after that folder was scanned is "
                                    + "not covered either; a fresh search with a tighter 'after' bound picks it up."));
                        break;

                    case FreshMerge.ScanGapResumedUnsorted:
                        advice.Add("A folder had to be re-read from its beginning to resume, because Outlook would "
                            + "not sort its table. Nothing is duplicated and nothing is lost - items already returned "
                            + "are suppressed by id - but each further page of that folder re-reads everything before "
                            + "it, so the cost grows page by page. A narrower 'after'/'before' window, or a folder "
                            + "scope aimed at that folder alone, is the cheap remedy.");
                        break;

                    case FreshMerge.ScanGapResumePositionLost:
                        advice.Add("A resumed folder's recorded position no longer identified the same item, so the "
                            + "table's row order had not held between pages - which MAPI allows, since an unsorted "
                            + "table has no guaranteed order and its rows follow the folder live. The folder was "
                            + "restarted with duplicate suppression rather than resumed into an unknown position, so "
                            + "these results are still correct and this page cost more than it should have.");
                        break;

                    case FreshMerge.ScanGapDedupCapacity:
                        advice.Add("Duplicate suppression for the folder being re-read has reached its capacity, so "
                            + "items already returned by an earlier page of this scan MAY appear again from here on. "
                            + "De-duplicate by id if you are accumulating results. Reaching this means one folder has "
                            + "been paged through many times - narrow the scan (folder, after/before) instead of "
                            + "continuing.");
                        break;

                    default:
                        // Same rule as the sweep's: a code with no sentence is a partial
                        // result an agent can see and cannot explain. T1 pins that every
                        // declared code is handled, so this is only reachable by a code
                        // added without prose - say so rather than dropping it.
                        advice.Add("The exhaustive scan reported partial coverage (" + gap
                            + ") with no further detail available; treat these results as incomplete.");
                        break;
                }
            }

            return advice;
        }

        /// <summary>
        /// The one clause both stop-reason sentences share: what to do about a walk that
        /// stopped, given whether a continuation handle could be issued for it.
        /// <para>
        /// Both branches matter. WITH a token, the remedy is concrete and the sentence names
        /// the field, because "narrow it yourself" was the only answer this mode had for a
        /// large store and it is no longer the best one. WITHOUT a token the answer that
        /// looks the same - a missing <c>nextToken</c> - would otherwise read as completeness,
        /// so the absence is stated rather than left to be inferred.
        /// </para>
        /// </summary>
        private static string ResumeRemedy(ExhaustiveInfo exhaustive, string context)
        {
            if (!string.IsNullOrEmpty(exhaustive.NextToken))
            {
                return "Continue where it stopped: pass exhaustive.nextToken back as resume_token with every other "
                    + "argument unchanged"
                    + (exhaustive.Position?.ResumeFolder == null
                        ? string.Empty
                        : " (next up: '" + exhaustive.Position.ResumeFolder + "')")
                    + ". " + context;
            }

            return "This scan could NOT be made resumable, so exhaustive.nextToken is absent - do not read that "
                + "absence as completeness. " + context;
        }

        /// <summary>
        /// Why a <c>resume_token</c> was refused, and what to do instead. FIVE distinct
        /// messages for five distinct causes, because the remedies genuinely differ: a
        /// malformed handle is a caller bug, an unknown one usually means the server
        /// restarted, an expired one means the chain aged out, a superseded one means a later
        /// page already exists, and a changed request means the caller asked something else.
        /// <para>
        /// Every refusal that can name the chain's POSITION does, because the alternative is
        /// throwing away work that cost minutes: the position is expressed in <c>folder</c>
        /// and <c>before</c>, which are parameters <c>search</c> already has, so recovery
        /// needs no token at all. That is also what makes a superseded replay safe to refuse
        /// outright - keeping older tokens live would need a snapshot of the suppression set
        /// per position, which is a lot of memory for a rare case, and honouring one without
        /// that snapshot would suppress exactly the rows the replay exists to return.
        /// </para>
        /// </summary>
        public static string DescribeResumeRefusal(
            ScanTokenDecision decision,
            ExhaustiveScanSession? session,
            IReadOnlyList<string>? changedArguments)
        {
            string recovery = session == null ? string.Empty : " " + session.DescribeRecovery();
            switch (decision)
            {
                case ScanTokenDecision.Malformed:
                    return "resume_token is not a continuation handle this server issues (they look like "
                        + "'scan-' followed by 32 hex characters). Pass back exactly the value from a previous "
                        + "result's exhaustive.nextToken, or omit it to start a fresh scan.";

                case ScanTokenDecision.Unknown:
                    return "resume_token is not known to this server: continuation state lives in the running "
                        + "process and is lost when it restarts, and it is also dropped once its scan finishes. "
                        + "Continue by re-running this search with 'folder' and 'before' set from the previous "
                        + "result's exhaustive.position, or omit resume_token to start a fresh scan.";

                case ScanTokenDecision.Expired:
                    return "resume_token has expired - a paged scan stays resumable for a limited time after its "
                        + "last page, and this one aged out." + recovery
                        + " Or omit resume_token to start a fresh scan.";

                case ScanTokenDecision.Superseded:
                    return "resume_token has been superseded: a later page of the same scan was already served, so "
                        + "this handle names a position the scan has moved past. Use the newest result's "
                        + "exhaustive.nextToken." + recovery;

                case ScanTokenDecision.RequestChanged:
                    return "resume_token belongs to a scan of a DIFFERENT query, so continuing it would answer a "
                        + "different question than the one it started: "
                        + (changedArguments == null || changedArguments.Count == 0
                            ? "one or more arguments changed"
                            : string.Join(", ", changedArguments) + " changed")
                        + ". Re-issue the call with the original arguments (only top and snippet_chars may differ "
                        + "between pages), or drop resume_token to start a fresh scan for the new question."
                        + recovery;

                default:
                    return "resume_token could not be used.";
            }
        }

        /// <summary>
        /// One sentence when a result set contains something other than ordinary mail, and
        /// nothing at all when it does not.
        /// <para>
        /// This is where the widening announces itself (gap B3). All three tiers now admit
        /// every item class, so a search can return a bounce report, a read receipt, a
        /// meeting request or - from the index tier, which has no folder-type column to
        /// narrow by - a calendar or contact item. That is the answer being MORE complete,
        /// not less, so it degrades nothing and raises no coverage code; but an agent
        /// relaying "you have 4 mails about the invoice" when one of them is a delivery
        /// receipt is relaying something false, and nothing else in the payload would have
        /// told it.
        /// </para>
        /// <para>
        /// It is emitted from the hits themselves rather than from a counter, so it can
        /// never disagree with the <c>itemClass</c> fields beside it, and it costs exactly
        /// nothing on the ordinary search where every hit is mail. The tool description says
        /// none of this: it is 1791 of its 2048 client-truncation units and this is a fact
        /// about ONE answer, which is what advice is for.
        /// </para>
        /// <para>
        /// Pure and public so T1 pins it - the states it describes need a mailbox holding
        /// meeting requests and NDRs to produce naturally.
        /// </para>
        /// </summary>
        public static string? DescribeNonMailHits(IReadOnlyList<HitSummary>? hits)
        {
            if (hits == null || hits.Count == 0)
            {
                return null;
            }

            int count = 0;
            bool omitted = false;
            List<string> classes = new List<string>();
            foreach (HitSummary hit in hits)
            {
                if (hit == null || hit.ItemClass == null)
                {
                    continue;
                }

                count++;
                if (classes.Contains(hit.ItemClass, StringComparer.OrdinalIgnoreCase))
                {
                    continue;
                }

                // The "..." has to mean something was actually left out. Deriving it from
                // "the list is full" would make it fire on EXACTLY the cap, which is a
                // has-more flag that lies - the shape this whole payload discipline exists
                // to prevent.
                if (classes.Count >= NonMailClassAdviceCap)
                {
                    omitted = true;
                    continue;
                }

                classes.Add(hit.ItemClass);
            }

            if (count == 0)
            {
                return null;
            }

            return count.ToString(CultureInfo.InvariantCulture)
                + " of these hits are not ordinary mail (" + string.Join(", ", classes)
                + (omitted ? ", ..." : string.Empty)
                + ") - every search tier returns bounce reports, read receipts, meeting requests and posts "
                + "alongside mail, and each such hit carries itemClass (a 'kind:' value means the index "
                + "inferred it rather than reading it). Say which is which if you relay a count.";
        }

        /// <summary>
        /// Compact scope block: present for a folder-scoped search, and for a store whose
        /// scope the index cannot address - the two cases where what was COVERED differs from
        /// what was asked for.
        /// </summary>
        private static SearchScopeInfo? DescribeSearchScope(FolderScopeResolution? folderScope, SearchRequest request)
        {
            // A store scope with no folder is ordinarily uninteresting - it covers exactly
            // what was asked for - so it stays out of the payload. StoreNotIndexed does not:
            // there the shape of the answer changed (one tier instead of two), and a block
            // that appeared only when a folder happened to be named would report that on
            // some calls and not others.
            if (folderScope == null
                || (folderScope.RequestedFolder == null && folderScope.Kind != FolderScopeKind.StoreNotIndexed))
            {
                return null;
            }

            string shape = folderScope.Kind switch
            {
                FolderScopeKind.PrimaryRecursive => "folder",
                FolderScopeKind.PrimaryNonRecursive => "folder_only",
                FolderScopeKind.DelegateFlat => "delegate_folders",
                FolderScopeKind.DelegateWidened => "delegate_store_widened",
                // Reported even when a folder WAS named: no index folder scope was built,
                // so calling this "folder" or "folder_only" would describe a narrowing that
                // never happened. The folder still bounds the sweep, and is still echoed.
                FolderScopeKind.StoreNotIndexed => "store_not_indexed",
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

                // Gap G4. folderNamesMatched is a count, and a count reads the same whether
                // the walk behind it saw the whole mailbox or stopped at its cap - so the
                // one thing that makes the number untrustworthy needs its own field.
                FolderNamesTruncated = folderScope.FolderTreeTruncated ? true : (bool?)null,
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

            // Gap B2, and set BEFORE any early return: it is a fact about what this tier can
            // read, so it is equally true of a sweep that ran, one that was refused and one
            // that was not needed. Deciding it here rather than at each exit is what keeps
            // it from going missing on the paths that report the least.
            info.AttachmentTextCovered =
                FreshMerge.AttachmentTextMatchable(request.SearchIn, terms.Count > 0) ? false : (bool?)null;

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
                //
                // COMPLETE ONLY WHERE THE INDEX ACTUALLY COVERS IT. "Not needed" is a claim
                // about the index tier, and it is false for a store the index holds no mail
                // for - which is precisely the store a 'before'-bounded search is reaching
                // into. So the profile is checked before that claim is made.
                info.Performed = false;
                info.NotNeeded = true;
                info.Error = null;
                return NoteProfileStoresWithoutIndex(info, request.Store, perStoreBase);
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
                // Two independent holes, and the answer owes the caller both: this filter has
                // no live tier at all, AND part of the profile has no index tier. Naming only
                // the refusal points at a remedy - drop the filter - that does not reach the
                // store the index cannot see.
                info.Performed = false;
                info.Error = refusal;
                return NoteProfileStoresWithoutIndex(info, request.Store, perStoreBase);
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

            // Gap E2. The token and the sentence are set in ONE statement from ONE classifier,
            // so the branchable field and the prose describe the same breadth by construction.
            // Both stay paired with the sweep that was PLANNED - a refused or failed sweep
            // returns before this and carries neither, because "it covered the default
            // folders" is a claim about coverage, and those paths covered nothing. That is
            // the one way this differs from B2's attachmentTextCovered, which survives every
            // early return precisely because it states a capability rather than a coverage.
            info.ScopeShape = ClassifySweepScope(folderKey != null, request.IncludeSubfolders);
            info.Scope = DescribeSweepScope(info.ScopeShape);

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
                    // Two budgets, one relationship. The inner one is what the walk itself
                    // measures and stops on gracefully; the outer one is the gateway
                    // deadline that reclaims a host which did not answer at all. The gap
                    // between them is the return trip, so a slow sweep always ends as
                    // partial coverage rather than as a timeout and a killed host.
                    sweepResult = _gateway.Run(
                    s => s.SweepFoldersNewerThan(
                        gapStart, SweepPerFolderCap, includeBodies: true, request.Store, sweepFolderPath, sweepRecursive,
                        perStoreGapStart, SweepWorkBudgetMs),
                    SweepBudgetMs,
                    allowConnectFloor: true);
                }
                // Each of the three failures below still names the stores with no index
                // behind them: "the sweep could not run" and "the index has nothing here" are
                // independent facts with different remedies - retry versus exhaustive:true -
                // and an answer missing BOTH tiers over a store has to say so.
                catch (OutlookUnavailableException ex)
                {
                    info.Performed = false;
                    info.Error = ex.Message;
                    return NoteProfileStoresWithoutIndex(info, request.Store, perStoreBase);
                }
                catch (TimeoutException ex)
                {
                    // A bounded COM failure: the operation exceeded its budget and the COM
                    // host was restarted. Its message already says what timed out and what
                    // was done about it, so surface that rather than a bare type name -
                    // this text reaches the agent, and through it the user.
                    info.Performed = false;
                    info.Error = ex.Message;
                    return NoteProfileStoresWithoutIndex(info, request.Store, perStoreBase);
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
                    return NoteProfileStoresWithoutIndex(info, request.Store, perStoreBase);
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
            //
            // The store list comes from the sweep's own counters here, so this path spends no
            // COM beyond the sweep it just did. The three paths that never reach a sweep ask
            // Outlook instead (NoteProfileStoresWithoutIndex); the rule they apply is the same
            // one, in the same method.
            NoteStoresWithoutIndex(
                info,
                effectiveResult.PerStore.Select(counters => counters.StoreDisplayName).ToList(),
                request.Store,
                perStoreBase);

            List<ComMailBrief> filtered = new List<ComMailBrief>();
            HashSet<string> unevaluatedFilters = new HashSet<string>(StringComparer.Ordinal);

            // The COM layer bounds each swept body so a frame too big to send cannot be built
            // (OutlookComSession.SweepBodyCharsCap / SweepBodyBytesBudget). A cut body can
            // only ever cost a hit where the terms actually go against the body, so that
            // condition is evaluated once here rather than per item - and where it is false,
            // the cut is free and no code is raised for it.
            bool termsReachTheBody = terms != null && terms.Count > 0 && request.SearchIn != SearchIn.SubjectOnly;
            int bodiesCapped = 0;
            int bodiesCappedUnmatched = 0;
            foreach (ComMailBrief item in sweptItems)
            {
                if (request.Store != null
                    && !string.Equals(item.StoreDisplayName, request.Store, StringComparison.OrdinalIgnoreCase))
                {
                    continue; // Cached all-stores sweep serving a store-scoped request.
                }

                info.ItemsSeen++;

                // Counted INSIDE the store filter above, so a cached all-stores sweep serving
                // a store-scoped request reports this store's cuts rather than the sweep's.
                bool bodyCut = item.BodyTruncated == true;
                if (bodyCut)
                {
                    bodiesCapped++;
                }

                if (!FreshMerge.MatchesTerms(item, terms, request.SearchIn))
                {
                    // The one case a body cut can have cost a result: this item was cut AND
                    // did not match. Whether the term really sat past the cut is exactly what
                    // the un-carried remainder would have answered, so this counts candidates
                    // and the advice says "may be".
                    if (bodyCut && termsReachTheBody)
                    {
                        bodiesCappedUnmatched++;
                    }

                    continue;
                }

                if (request.From != null
                    && !(Contains(item.SenderAddress, request.From) || Contains(item.SenderName, request.From)))
                {
                    continue;
                }

                // Gap I1. Four filters below need a property the COM snapshot may not
                // carry, and each of them used to drop such an item SILENTLY. The drop
                // itself stays - a filter the caller asked for has to be honoured, and
                // admitting an item that cannot be shown to match would corrupt the answer
                // the other way round - but it is now counted, and the filter that could not
                // be evaluated is named, so the caller can re-run without it. See
                // FreshMerge.GapFilterUnreadable for the decision in full.
                DateTime? receivedUtc = ToUtc(item.ReceivedTime);
                bool dateFilterRequested = request.BeforeUtc.HasValue || request.AfterUtc.HasValue;
                if (dateFilterRequested && receivedUtc == null)
                {
                    // Deliberately BEFORE the comparisons: an item with no usable timestamp
                    // fails both of them, and the reason is the missing value, not the
                    // bound. Which of `before`/`after` is named follows what was asked.
                    if (request.BeforeUtc.HasValue)
                    {
                        NoteUnevaluatedFilter(unevaluatedFilters, "before");
                    }

                    if (request.AfterUtc.HasValue)
                    {
                        NoteUnevaluatedFilter(unevaluatedFilters, "after");
                    }

                    info.ItemsFilterUnreadable++;
                    continue;
                }

                if (request.BeforeUtc.HasValue && receivedUtc!.Value >= request.BeforeUtc.Value)
                {
                    continue;
                }

                if (request.AfterUtc.HasValue && receivedUtc!.Value < request.AfterUtc.Value)
                {
                    continue;
                }

                if (request.UnreadOnly == true && item.IsRead == null)
                {
                    NoteUnevaluatedFilter(unevaluatedFilters, "unread_only");
                    info.ItemsFilterUnreadable++;
                    continue;
                }

                if (request.UnreadOnly == true && item.IsRead != false)
                {
                    continue;
                }

                if (request.HasAttachments.HasValue && item.HasAttachments == null)
                {
                    NoteUnevaluatedFilter(unevaluatedFilters, "has_attachments");
                    info.ItemsFilterUnreadable++;
                    continue;
                }

                if (request.HasAttachments.HasValue && item.HasAttachments != request.HasAttachments.Value)
                {
                    continue;
                }

                filtered.Add(item);
            }

            // Reported in the request's own parameter order, so the names read as the
            // remedy they are: drop the one named and the dropped items come back.
            info.FiltersUnevaluated = OrderUnevaluatedFilters(unevaluatedFilters);

            // The body bounds, reported from the per-item flags this loop just read rather
            // than from the sweep's whole-sweep totals - the same store-attribution rule
            // ApplySweepCounters applies to every other counter, and the reason those totals
            // are deliberately not copied across (ComSweepResult.BodiesTruncated).
            info.ItemsBodyCapped = bodiesCapped > 0 ? bodiesCapped : (int?)null;
            info.ItemsBodyCappedUnmatched = bodiesCappedUnmatched > 0 ? bodiesCappedUnmatched : (int?)null;

            // WHICH bound cut is a fact about the frame, so it comes from the sweep result -
            // but it is only carried when this scope actually lost something, so a
            // store-scoped answer never reports a budget another account's mail exhausted.
            info.BodyBudgetExhausted = bodiesCapped > 0 && effectiveResult.BodyBudgetExhausted
                ? true
                : (bool?)null;

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
        /// The COM error token that means "the item could not be OPENED where we looked" -
        /// set by every by-EntryID operation at its <c>GetItemFromID</c> and nowhere else, so
        /// it is the one failure a second attempt against a different store can fix, and the
        /// one that is guaranteed to have changed nothing yet.
        /// <para>
        /// Aliased from <see cref="ComErrorTokens.ItemNotFound"/> rather than spelt again:
        /// the writer and the reader of this word sit in different namespaces, and while it
        /// was written twice a COM path that stopped setting it broke a retry loop silently.
        /// Now the two cannot disagree.
        /// </para>
        /// </summary>
        public const string ItemNotFoundToken = ComErrorTokens.ItemNotFound;

        /// <summary>
        /// Whether a failed by-EntryID operation should be re-attempted store by store: only
        /// for a bare EntryID (no store was known), and only when the item was not FOUND.
        /// <para>
        /// One rule in one place because the two loops that needed it had drifted apart in
        /// opposite directions, each in a way its own author could not see. <c>reply_draft</c>
        /// and friends retried on failure ALONE, so a compose or Save error fanned out across
        /// every store on the profile - and creating a draft is not idempotent, so a run that
        /// still ended in failure could leave an orphan per store, none of whose ids the
        /// caller ever learns. <c>update_draft</c> and <c>discard_draft</c> asked for this
        /// token, which their COM layer never produced, so their retry was dead code and a
        /// draft in a non-default store answered with an opaque <c>COMException 0x...</c>
        /// instead of being found.
        /// </para>
        /// <para>Pure, and public so T1 pins the rule without a mailbox.</para>
        /// </summary>
        public static bool ShouldSearchOtherStores(string? storeId, bool succeeded, string? error)
        {
            return !succeeded && storeId == null && string.Equals(error, ItemNotFoundToken, StringComparison.Ordinal);
        }

        /// <summary>
        /// Whether the store-by-store loop should try the NEXT store. It stops the moment a
        /// store answers - with the item, or with any refusal other than "not found". A store
        /// that opened the item has answered the question the loop is asking, and trying the
        /// rest would repeat work that is not free: an update re-applies attachments and
        /// signatures, a derived draft leaves another orphan behind.
        /// </summary>
        public static bool KeepSearchingStores(bool succeeded, string? error)
        {
            return !succeeded && string.Equals(error, ItemNotFoundToken, StringComparison.Ordinal);
        }

        /// <summary>
        /// Puts the MERGED hit list into the order the caller asked for.
        /// <para>
        /// The index tier asks the provider for an ORDER BY, but that only decides which
        /// rows come back - the list the caller receives is index hits plus freshness-sweep
        /// hits (or, on the exhaustive tier, a COM walk), re-sorted after the merge. That
        /// re-sort was unconditionally by date, so a size-ordered search was silently
        /// reordered one layer above the query that honoured it and the caller's ordering
        /// survived only as far as the SQL.
        /// </para>
        /// <para>
        /// An unknown key sorts LAST in both orders: a hit whose size the provider never
        /// reported is not a zero-byte mail, and treating it as one would put it at the top
        /// of an ascending comparison and the bottom of a descending one for no reason.
        /// Size ties fall back to date because <see cref="List{T}.Sort(Comparison{T})"/> is
        /// not stable, and two equally large mails should not swap places between runs.
        /// </para>
        /// <para>Pure, and public so T1 pins the ordering without a mailbox.</para>
        /// </summary>
        public static void SortForOrder(List<HitSummary> summaries, bool bySizeDescending)
        {
            if (summaries == null)
            {
                throw new ArgumentNullException(nameof(summaries));
            }

            if (!bySizeDescending)
            {
                summaries.Sort((a, b) => DateTime.Compare(b.ReceivedUtc ?? DateTime.MinValue, a.ReceivedUtc ?? DateTime.MinValue));
                return;
            }

            summaries.Sort((a, b) =>
            {
                int bySize = (b.SizeBytes ?? -1L).CompareTo(a.SizeBytes ?? -1L);
                return bySize != 0
                    ? bySize
                    : DateTime.Compare(b.ReceivedUtc ?? DateTime.MinValue, a.ReceivedUtc ?? DateTime.MinValue);
            });
        }

        /// <summary>
        /// The request-filter names that can fail to evaluate on a swept item, in the order
        /// they are reported (gap I1). Also the T1 pin that the reported names are exactly
        /// the names of the request parameters a caller passes - a name that does not match
        /// a parameter is advice nobody can act on.
        /// </summary>
        public static readonly IReadOnlyList<string> SweepFilterNames =
            new[] { "unread_only", "has_attachments", "before", "after" };

        private static void NoteUnevaluatedFilter(HashSet<string> names, string filter)
        {
            names.Add(filter);
        }

        /// <summary>
        /// The unevaluated-filter names in <see cref="SweepFilterNames"/> order, or null when
        /// every filter could be evaluated on every item. Pure, and public so T1 pins the
        /// ordering without a mailbox.
        /// </summary>
        public static IReadOnlyList<string>? OrderUnevaluatedFilters(IReadOnlyCollection<string>? names)
        {
            if (names == null || names.Count == 0)
            {
                return null;
            }

            List<string> ordered = new List<string>(names.Count);
            foreach (string candidate in SweepFilterNames)
            {
                if (names.Contains(candidate))
                {
                    ordered.Add(candidate);
                }
            }

            return ordered.Count == 0 ? null : ordered;
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
        /// Names the stores in scope that the index holds nothing for, so an answer resting on
        /// the sweep alone says which store it is resting on. Only ever ADDS to what the
        /// frontier probes already found.
        /// <para>
        /// The verdict is <see cref="StoresMissingFromIndex"/>, the same pure rule
        /// <c>outlook_health</c> compares its two store lists with (gap A3), so "the index
        /// holds nothing for this store" means one thing in this server rather than two:
        /// absence from the catalog is never evidence, only a probe's NO counts, and a store
        /// that could not be settled is reported neither way.
        /// </para>
        /// <para>
        /// <paramref name="storeNames"/> is the profile's stores from whichever source this
        /// call has cheaply in hand - the sweep's own per-store counters after it ran, and
        /// Outlook's store list when it did not (<see cref="NoteProfileStoresWithoutIndex"/>).
        /// Null means "could not be established", which reports nothing at all: inventing a
        /// missing store out of a list nobody could read would be the same defect pointing the
        /// other way.
        /// </para>
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
            IReadOnlyList<string>? storeNames,
            string? requestedStore,
            IReadOnlyDictionary<string, DateTime>? perStoreBase)
        {
            if (requestedStore != null)
            {
                return; // Scoped: the search's own frontier probe already settled it.
            }

            Stopwatch clock = Stopwatch.StartNew();
            IReadOnlyList<string> missing = StoresMissingFromIndex(
                storeNames,

                // A store with a per-store window HAS a frontier, so the index demonstrably
                // knows it and no probe is worth spending on it.
                perStoreBase?.Keys.ToList(),

                // A store left unprobed because the budget ran out answers null, never false:
                // "not established" must not become "not indexed".
                store => clock.ElapsedMilliseconds > StoreIndexProbeBudgetMs
                    ? (bool?)null
                    : StoreHasIndexRows(store));

            if (missing.Count == 0)
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
        /// The same naming pass for a sweep that never ran, sourcing the profile's stores from
        /// Outlook because there is no sweep result to read them out of. Returns
        /// <paramref name="info"/> so the three places that give up on a sweep can end on one
        /// statement.
        /// <para>
        /// THE DEFECT THIS CLOSES (gap A1, residue). The pass above ran only after a sweep,
        /// off the sweep's own store list. On a MIXED profile - an indexed mailbox plus an
        /// unindexed data file, the ordinary archive-PST shape - the profile-wide frontier
        /// probe succeeds and the per-store loop only walks the index's CATALOG, which has
        /// never heard of that data file, so nothing else names it. All three paths where the
        /// sweep does not run were therefore silent about it: the window was NOT NEEDED (a
        /// <c>before</c> bound older than the fallback span - which is how an agent searches
        /// for OLD mail, i.e. exactly what such a store holds - answering
        /// <c>freshness: "live"</c> with <c>degraded</c> absent); the sweep was REFUSED (a
        /// recipient or attachment-only filter); or it FAILED. The first is the one that
        /// misleads, and the other two were missing the half of the answer that says which
        /// tier is gone for good rather than until a retry.
        /// </para>
        /// <para>
        /// COST, and where it is not paid: nothing at all for a store-scoped search, whose own
        /// frontier probe already settled the question, and nothing on the ordinary path, which
        /// still reads the sweep's counters. Only a sweep that did not happen pays one COM
        /// store-list read, from the <see cref="StoreDetailsCacheTtl"/> cache. Two of those
        /// three paths - not needed, and refused - previously spent no COM at all, so this is
        /// a real addition to them rather than a rounding error on work already done; it is
        /// the completeness-over-cost trade taken deliberately, because a search that cannot
        /// see a whole store must say so. On the third the sweep has just failed, and the
        /// store list may fail with it: that answers null and reports nothing, which is
        /// correct - after a COM failure there is no evidence either way.
        /// </para>
        /// </summary>
        private SweepInfo NoteProfileStoresWithoutIndex(
            SweepInfo info,
            string? requestedStore,
            IReadOnlyDictionary<string, DateTime>? perStoreBase)
        {
            if (requestedStore == null)
            {
                NoteStoresWithoutIndex(info, TryGetProfileStoreNames(), requestedStore, perStoreBase);
            }

            return info;
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

            // The continuation contract (F2). The fingerprint is taken from the request the
            // caller just made, and a resume is honoured only where it matches the request
            // the chain was opened for - answering a different question under a claim of
            // continuity is the failure this whole mechanism is built against, and both ways
            // of not refusing (honour it silently, ignore it silently) are that failure.
            string fingerprint = ExhaustiveScanCursors.FingerprintOf(request, terms);
            ExhaustiveScanSession? session = null;
            ComScanCursor? resumeFrom = null;
            if (!string.IsNullOrWhiteSpace(request.ResumeToken))
            {
                ScanTokenDecision decision = _scanCursors.Resolve(request.ResumeToken, fingerprint, out session);
                if (decision != ScanTokenDecision.Valid)
                {
                    IReadOnlyList<string> changed = decision == ScanTokenDecision.RequestChanged && session != null
                        ? ExhaustiveScanCursors.DifferingArguments(session.Fingerprint, fingerprint)
                        : Array.Empty<string>();
                    throw new ArgumentException(
                        DescribeResumeRefusal(decision, session, changed), nameof(request));
                }

                resumeFrom = session!.Cursor;
            }

            Stopwatch stopwatch = Stopwatch.StartNew();

            // The scan runs on its OWN deadline, not the shared one. Two numbers, one
            // relationship: the outer budget below is the exhaustive class's hard deadline
            // and the inner timeBudgetMs is that same deadline less the return trip, so the
            // walk always stops and hands back partial results before the watchdog can
            // decide the host is wedged. Stated explicitly here as well as in
            // ComHostOperationClass because the enclosing gateway operation would otherwise
            // bound the lambda at the ORDINARY deadline and clip the scan back to it - an
            // aggregate is measured across the lambda, and this lambda's one call is
            // allowed to be longer than an ordinary one.
            ComExhaustiveResult scan = _gateway.Run(
                s => s.ExhaustiveScan(
                    request.Store!,
                    folderSegments,
                    terms,
                    request.AfterUtc,
                    request.BeforeUtc,
                    maxItems: top,
                    timeBudgetMs: ExhaustiveTimeBudgetMs,
                    searchIn: request.SearchIn,
                    includeSubfolders: request.IncludeSubfolders,
                    resumeFrom: resumeFrom),
                ComOperationBudgets.ExhaustiveScanDeadlineMs,
                allowConnectFloor: true);
            stopwatch.Stop();

            List<HitSummary> summaries = new List<HitSummary>();
            HashSet<string> scanFiltersUnevaluated = new HashSet<string>(StringComparer.Ordinal);
            int scanItemsFilterUnreadable = 0;

            // Gap F3. Everything below this line runs over scan.Items, which the COM scan
            // already capped at `top` - so these three filters thin a list the cap has
            // closed, and the number they remove is the difference between "top MATCHES
            // exist" and "top CANDIDATES were reached". Counted so the payload can say so.
            int scanItemsFilteredOut = 0;
            foreach (ComMailBrief item in scan.Items)
            {
                if (request.From != null
                    && !(Contains(item.SenderAddress, request.From) || Contains(item.SenderName, request.From)))
                {
                    scanItemsFilteredOut++;
                    continue;
                }

                // Gap I1's other tier. Same snapshots, same unreadable properties, same
                // deliberate drop - so the same counter and the same code, rather than a
                // second answer to one question. There is no before/after here: this mode's
                // date bounds go into the DASL filter and are never read back off the item.
                if (request.UnreadOnly == true && item.IsRead == null)
                {
                    NoteUnevaluatedFilter(scanFiltersUnevaluated, "unread_only");
                    scanItemsFilterUnreadable++;
                    continue;
                }

                if (request.UnreadOnly == true && item.IsRead != false)
                {
                    scanItemsFilteredOut++;
                    continue;
                }

                if (request.HasAttachments.HasValue && item.HasAttachments == null)
                {
                    NoteUnevaluatedFilter(scanFiltersUnevaluated, "has_attachments");
                    scanItemsFilterUnreadable++;
                    continue;
                }

                if (request.HasAttachments.HasValue && item.HasAttachments != request.HasAttachments.Value)
                {
                    scanItemsFilteredOut++;
                    continue;
                }

                summaries.Add(RegisterLiveHit(item, snippetChars: 0, source: "exhaustive"));
            }

            SortForOrder(summaries, request.OrderBySizeDescending);
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

            ExhaustiveInfo exhaustive = new ExhaustiveInfo
            {
                Engine = scan.Engine,
                InstantSearchEnabled = scan.InstantSearchEnabled,
                FoldersScanned = scan.FoldersScanned,
                FoldersSkipped = scan.FoldersSkipped,
                Truncated = scan.Truncated,
                TimedOut = scan.TimedOut,
                DepthLimitReached = scan.DepthLimitReached,
                RowsDropped = scan.RowsDropped,
                RowsUnreadable = scan.RowsUnreadable,
                ItemsFilterUnreadable = scanItemsFilterUnreadable,
                FiltersUnevaluated = OrderUnevaluatedFilters(scanFiltersUnevaluated),
                PostCapFilters = FreshMerge.PostCapFilters(
                    request.From != null, request.UnreadOnly == true, request.HasAttachments.HasValue),
                ItemsFilteredOut = scanItemsFilteredOut,
                ElapsedMs = stopwatch.ElapsedMilliseconds,
                StopReason = scan.StopReason,
                Resumed = resumeFrom != null ? true : (bool?)null,
                TreeChangedFolders = scan.TreeChangedFoldersAdded + scan.TreeChangedFoldersMissing > 0
                    ? scan.TreeChangedFoldersAdded + scan.TreeChangedFoldersMissing
                    : (int?)null,
                CursorFolderMissing = scan.CursorFolderMissing ? true : (bool?)null,
                ResumedUnsorted = scan.ResumedUnsorted ? true : (bool?)null,
                ResumePositionLost = scan.ResumePositionLost ? true : (bool?)null,
                DedupCapacityReached = scan.DedupCapacityReached ? true : (bool?)null,
            };

            // The token, and the three states it has. A walk that COVERED its scope closes
            // its chain (the state exists to make a next page possible, and there is no next
            // page). A walk that stopped WITH a position gets the handle for that position. A
            // walk that stopped WITHOUT one gets no handle at all, and the advice says so -
            // rather than leaving a caller to read a missing field as completeness, which is
            // exactly what the field's own absence means in the other case.
            if (string.Equals(scan.StopReason, ComScanStopReasons.Complete, StringComparison.Ordinal))
            {
                _scanCursors.Complete(session);
                exhaustive.ItemsReturnedTotal = (session?.ItemsReturnedTotal ?? 0) + summaries.Count;
            }
            else if (scan.Position != null)
            {
                exhaustive.NextToken = _scanCursors.Issue(
                    session, fingerprint, scan.Position, summaries.Count, out ExhaustiveScanSession issued);
                session = issued;
                exhaustive.ItemsReturnedTotal = issued.ItemsReturnedTotal;
            }
            else
            {
                _scanCursors.NoteUnresumablePage(session, summaries.Count);
                exhaustive.ItemsReturnedTotal = (session?.ItemsReturnedTotal ?? summaries.Count);
            }

            if (scan.Position != null)
            {
                exhaustive.Position = new ScanPositionInfo
                {
                    FoldersDone = scan.Position.FoldersDone,
                    FoldersTotal = scan.Position.FoldersTotal,
                    ResumeFolder = scan.Position.ResumeFolderPath,
                    ResumeWithinFolder = scan.Position.ResumeWithinFolder,
                    ResumeCursorUtc = scan.Position.ResumeCursorUtc,
                    ResumeTier = scan.Position.ResumeTier,
                    Page = session?.PagesServed ?? 1,
                };
            }

            // The codes first, then the prose from the codes, then the verdict recomputed
            // from the same codes - the order the sweep and the thread walk already use, so
            // this tier cannot report a hole in one rendering and full coverage in another.
            exhaustive.CoverageGaps = FreshMerge.DescribeExhaustiveCoverageGaps(exhaustive);
            advice.AddRange(DescribeExhaustiveCoverage(exhaustive, top, summaries.Count));

            // The same sentence the merged path emits, from the same helper: this tier
            // gained the most from the widening (it was the only one that could not find a
            // bounce report at all), so it is the last place that should stay quiet about it.
            string? scanNonMailAdvice = DescribeNonMailHits(summaries);
            if (scanNonMailAdvice != null)
            {
                advice.Add(scanNonMailAdvice);
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

            // The SAME two top-level flags the merged path sets, from this mode's own
            // counters (FreshMerge.ClassifyExhaustiveFreshness). They used to be absent
            // here, so the one search mode a caller reaches for BECAUSE completeness
            // matters was the one that never said it had fallen short - the partial-scan
            // facts lived in exhaustive.* while degraded and freshness, the two fields the
            // tool description tells agents to branch on, stayed empty.
            string freshness = FreshMerge.ClassifyExhaustiveFreshness(exhaustive);

            return new SearchOutcome
            {
                Hits = summaries,
                Truncated = scan.Truncated,
                Degraded = freshness == FreshMerge.FreshnessLive ? (bool?)null : true,
                Freshness = freshness,
                IndexElapsedMs = 0,
                Sweep = null,
                Exhaustive = exhaustive,
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
                if (ShouldSearchOtherStores(storeId, d != null, error))
                {
                    // Direct EntryID without a known store: retry across stores, and ONLY
                    // when the item could not be OPENED. The guard used to be `d == null`
                    // alone, which asked every store on the profile to re-open an item that
                    // one of them had already opened and then failed to snapshot - a body
                    // Outlook would not render reads the same as an id that is not there.
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        d = s.TryReadItem(entryId, store.StoreId, includeHeaders, includeBody: !haveCachedBody, out error, includeHtml);
                        if (!KeepSearchingStores(d != null, error))
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
                if (ShouldSearchOtherStores(storeId, saved != null, error))
                {
                    // Direct EntryID without a known store, and ONLY when the item could not
                    // be OPENED. This loop writes to disk, which is why save_attachment is
                    // classified MUTATING (ComSessionOperations). Under the old `saved ==
                    // null` guard, an attachment index that was out of range, or a save that
                    // failed on permissions, re-ran the whole save against every store on
                    // the profile - so a partially written or oddly named file could be left
                    // behind per store while the call still reported failure.
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        saved = s.TrySaveAttachment(entryId, store.StoreId, attachmentIndex, directory, out sizeBytes, out error);
                        if (!KeepSearchingStores(saved != null, error))
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
        /// Time budget for thread's live conversation walk. The same budget the freshness
        /// sweep runs under, and stated as that relationship rather than as a second
        /// literal, because it is the same kind of work: a bounded live COM check layered
        /// over an answer the index already produced. It carries the sweep's connect floor
        /// too - the walk may be the call that starts Outlook.
        /// </summary>
        public const int ThreadWalkBudgetMs = SweepBudgetMs;

        /// <summary>
        /// Resolves a conversation: index ConversationID query, then the live Outlook
        /// Conversation walk, merged and deduped (v3.MD section 0.6 Phase 2).
        /// <para>
        /// THE WALK IS NO LONGER A FALLBACK, and that is the whole of gap C1. It used to run
        /// only when the index returned zero rows, so ONE indexed row for a conversation was
        /// enough to skip the live check entirely - and a reply that arrived after the index
        /// frontier was absent from a payload whose tool description promises the FULL
        /// conversation, with no <c>degraded</c>, no <c>freshness</c> and no staleness block
        /// to hint at it. <c>thread</c> was the only tool on this surface with no way to
        /// express a partial answer.
        /// </para>
        /// <para>
        /// SWEEPING IS THE RIGHT ANSWER HERE AND IT IS CHEAP, which is why this reports AND
        /// checks rather than only reporting. A search has to guess where new mail might be,
        /// which is what makes its sweep a window over a folder SET with per-folder caps and
        /// a 30 s budget. A conversation is a scope Outlook can enumerate directly: one
        /// GetConversation plus one table walk over the members, which is why this path has
        /// existed as the COM fallback since Phase 2 and is already exercised on every
        /// conversation the index has never seen. The cost it adds to the previously
        /// index-only case is one COM round trip plus one GetItemFromID per member, bounded
        /// by <see cref="ThreadWalkBudgetMs"/>, plus - on the first use of a hit id that has
        /// not been located yet - the same ~2 s HitLocator probe that <c>read</c> pays on
        /// that hit, cached thereafter. Against that: <c>search</c> already sweeps folders
        /// through COM on EVERY call and starts Outlook headless to do it, so a
        /// <c>thread</c> that skipped the live check was not being consistent with the
        /// product's stance, it was the exception to it.
        /// </para>
        /// <para>
        /// The walk needs a concrete item to start from - COM cannot look up a conversation
        /// by id string - so a caller who passed only <c>conversation_id</c> gets
        /// <c>freshness: "index-only"</c> with the remedy named. That is the one degraded
        /// state here the CALLER can clear, and it is why it is a reported state rather than
        /// an error.
        /// </para>
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

            string? scope = null;
            bool scopeWidened = false;
            if (effectiveStore != null)
            {
                try
                {
                    scope = ResolveScope(effectiveStore, null);

                    // The second way to get no scope, added with the unindexed-store fix: a
                    // store the profile HAS and the index cannot address resolves to null
                    // instead of throwing. The lookup then widens exactly as an unknown name
                    // does, so it must be reported exactly as one - reading the flag off the
                    // catch alone would have made this half silent again.
                    scopeWidened = scope == null;
                }
                catch (ArgumentException)
                {
                    // C3: the store did not resolve to an index scope, so the lookup runs
                    // profile-wide. Over-returning members of the right conversation is the
                    // safe direction and the widening stays - but it used to happen in
                    // silence, and a scope the caller chose that was not applied is a fact
                    // the caller is entitled to.
                    scope = null;
                    scopeWidened = true;
                }
            }

            List<IndexHit> indexHits = new List<IndexHit>();
            List<HitSummary> hits = new List<HitSummary>();
            bool truncated = false;
            if (conversationId != null)
            {
                IndexSearchResult result = _index.Value.Search(
                    new IndexQuery
                    {
                        Scope = scope,
                        // C2: message rows of EVERY item class, not 'email' alone. A meeting
                        // request and its acceptances share the mail's ConversationID and
                        // index as 'calendar', so a kind-narrowed query dropped real members
                        // of the conversation this tool exists to return whole. Since B3 it
                        // is the same shape `search` uses, minus the attachment rows a
                        // conversation has no use for.
                        Kinds = KindFilter.MessagesOnly,
                        ConversationIdEquals = conversationId,
                        Top = top + 1, // Over-fetch by one: definite has-more flag.
                    },
                    SearchIndexTimeoutSeconds);

                truncated = result.Hits.Count > top;
                foreach (IndexHit hit in result.Hits)
                {
                    if (indexHits.Count >= top)
                    {
                        break; // The over-fetched row is evidence, not a member.
                    }

                    indexHits.Add(hit);
                    hits.Add(RegisterIndexHit(hit, snippetChars: SnippetCharsDefault));
                }
            }

            ThreadLiveInfo live = RunConversationWalk(id, top, indexHits, hits);

            // The stores the index says this conversation reaches, which is the only
            // evidence available that the walk (one store, by Outlook's construction) left
            // part of it unchecked.
            HashSet<string?> indexStores = new HashSet<string?>(StringComparer.OrdinalIgnoreCase);
            foreach (IndexHit hit in indexHits)
            {
                indexStores.Add(FreshMerge.ResolveHitStore(hit));
            }

            // C4's silent half, and it has to come BEFORE the codes are computed because
            // they are read off the block. The index rows above can only ever reveal a store
            // the INDEX knows, so on the profile shape this product keeps meeting - one
            // indexed mailbox plus a data file Windows Search has never opened - a
            // conversation reaching into that data file produced no rows to compare and
            // nothing said the walk had covered one store of two.
            NoteThreadStoresWithoutIndex(live);

            live.CoverageGaps = FreshMerge.DescribeThreadCoverageGaps(live, indexStores);
            string freshness = FreshMerge.ClassifyThreadFreshness(live, indexStores);

            // Newest-first to decide WHICH members survive the cap - the index tier returns
            // the newest top rows and the live tier exists to add newer ones still, so
            // trimming from the old end is the only trim that cannot throw away the reply
            // this whole path was added to find. Oldest-first afterwards, which is the
            // documented member order.
            hits.Sort((a, b) => DateTime.Compare(b.ReceivedUtc ?? DateTime.MinValue, a.ReceivedUtc ?? DateTime.MinValue));
            if (hits.Count > top)
            {
                hits.RemoveRange(top, hits.Count - top);
                truncated = true;
            }

            hits.Reverse();

            // Inside the clock, not after it: elapsedMs is what this call cost, and a probe
            // billed to nobody is how a tool's reported cost drifts from its real one.
            StalenessInfo staleness = ProbeStaleness(scope);

            stopwatch.Stop();
            return new ThreadOutcome
            {
                ConversationId = conversationId,
                Degraded = freshness == FreshMerge.FreshnessLive ? (bool?)null : true,
                Freshness = freshness,
                Source = indexHits.Count > 0 || !live.Performed ? "index" : "com",
                Hits = hits,
                Truncated = truncated,
                ScopeWidened = scopeWidened ? true : (bool?)null,
                Live = live,
                Staleness = staleness,
                Advice = DescribeThreadCoverage(live, freshness, effectiveStore, scopeWidened, top),
                ElapsedMs = stopwatch.ElapsedMilliseconds,
            };
        }

        /// <summary>
        /// Runs the live conversation walk and merges what it found into
        /// <paramref name="hits"/>, or reports why it could not run.
        /// <para>
        /// A failure DEGRADES this one tier, exactly as a failed freshness sweep degrades a
        /// search rather than failing it - but only while the index tier has produced an
        /// answer to degrade. With <paramref name="indexHits"/> empty the walk is not an
        /// enhancement, it is the whole lookup, and swallowing its failure would turn a
        /// caller's bad id or a wedged Outlook into an empty conversation reported as a
        /// success. That case throws, exactly as it did when this walk was the fallback -
        /// the new reporting is added to the merged path, it does not replace an error that
        /// was already right.
        /// </para>
        /// </summary>
        private ThreadLiveInfo RunConversationWalk(
            string? id,
            int top,
            IReadOnlyList<IndexHit> indexHits,
            List<HitSummary> hits)
        {
            ThreadLiveInfo live = new ThreadLiveInfo();
            if (id == null)
            {
                // The one degraded state the CALLER can clear: Outlook's conversation graph
                // is reachable only from a concrete item, so conversation_id alone leaves
                // this tier with nothing to start from.
                live.Performed = false;
                live.Error = "NoAnchorItem";
                return live;
            }

            bool degradeOnly = indexHits.Count > 0;
            Stopwatch clock = Stopwatch.StartNew();
            string entryId;
            string? storeId;
            try
            {
                (entryId, storeId, _, _, _) = ResolveToEntryId(id);
            }
            catch (ArgumentException) when (degradeOnly)
            {
                // The id is not a hit id from this session and not a plausible EntryID.
                clock.Stop();
                live.Performed = false;
                live.Error = "UnknownAnchorId";
                live.ElapsedMs = clock.ElapsedMilliseconds;
                return live;
            }
            catch (Exception ex) when (degradeOnly && ex is not OutOfMemoryException)
            {
                // A real id Outlook would not open: a stale index row, a moved item, a
                // folder that no longer exists. LocateFailureAdvice's text is the remedy and
                // it reaches the caller when this is the only tier; here it is one token.
                clock.Stop();
                live.Performed = false;
                live.Error = DescribeWalkFailure(ex, "AnchorNotLocatable");
                live.ElapsedMs = clock.ElapsedMilliseconds;
                return live;
            }

            IReadOnlyList<ComMailBrief> briefs;
            try
            {
                briefs = _gateway.Run(
                    s =>
                    {
                        IReadOnlyList<ComMailBrief>? items =
                            s.TryGetConversationItems(entryId, storeId, top + 1, out string? error);
                        return items ?? throw new InvalidOperationException(
                            "Conversation walk failed (" + (error ?? "unknown") + ").");
                    },
                    ThreadWalkBudgetMs,
                    allowConnectFloor: true);
            }
            catch (Exception ex) when (degradeOnly && ex is not OutOfMemoryException)
            {
                clock.Stop();
                live.Performed = false;
                live.Error = DescribeWalkFailure(ex, "ConversationWalkFailed");
                live.ElapsedMs = clock.ElapsedMilliseconds;
                return live;
            }

            live.Performed = true;
            live.MemberCapReached = briefs.Count > top;
            List<ComMailBrief> walked = new List<ComMailBrief>(Math.Min(briefs.Count, top));
            foreach (ComMailBrief brief in briefs)
            {
                if (walked.Count >= top)
                {
                    break; // The over-fetched member is evidence, not a member.
                }

                walked.Add(brief);
                live.AnchorStore ??= string.IsNullOrEmpty(brief.StoreDisplayName) ? null : brief.StoreDisplayName;
            }

            live.MembersWalked = walked.Count;
            IReadOnlyList<ComMailBrief> freshOnly = FreshMerge.SelectFreshOnly(
                walked, indexHits, DedupeToleranceSeconds, out int _);
            live.MembersAdded = freshOnly.Count;
            foreach (ComMailBrief brief in freshOnly)
            {
                hits.Add(RegisterLiveHit(brief, snippetChars: 0, source: "com"));
            }

            clock.Stop();
            live.ElapsedMs = clock.ElapsedMilliseconds;
            return live;
        }

        /// <summary>
        /// Names the stores neither tier covered for this conversation (gap C4's silent
        /// half), by the SAME rule the sweep names them with: Outlook's own store list,
        /// probed store by store, verdict from the pure <see cref="StoresMissingFromIndex"/>.
        /// Reusing it is the point - "the index holds nothing for this store" already means
        /// one thing in this server, and a second rule for <c>thread</c> would be a second
        /// thing to keep true.
        /// <para>
        /// THE HOLE IT CLOSES. <c>unwalked_store</c> is raised from the stores this
        /// conversation's INDEX ROWS name, so it is blind exactly where the index is: on a
        /// mixed profile the walk covers the anchor's store, the data file contributes no
        /// rows to compare against, and half a conversation can be missing under
        /// <c>freshness: "live"</c>. This pass asks Outlook instead of asking the index about
        /// itself.
        /// </para>
        /// <para>
        /// The gate below is a COST decision, not the rule: a walk that did not run or found
        /// nothing would be answered "nothing to report" by
        /// <see cref="FreshMerge.UnwalkedUnindexedStores"/> anyway, which is where the rule
        /// lives and where it is pinned - the gate only avoids paying for probes whose answer
        /// cannot be used. What it costs when it does run is one store-list read and one
        /// probe per store, both from the <see cref="StoreDetailsCacheTtl"/> caches the
        /// search path already fills, under <see cref="StoreIndexProbeBudgetMs"/>; a store
        /// left unprobed answers null and is reported neither way, because silence here means
        /// "not established" and never "indexed".
        /// </para>
        /// </summary>
        private void NoteThreadStoresWithoutIndex(ThreadLiveInfo live)
        {
            if (!live.Performed || live.MembersWalked <= 0 || string.IsNullOrEmpty(live.AnchorStore))
            {
                return;
            }

            Stopwatch clock = Stopwatch.StartNew();
            IReadOnlyList<string> missing = StoresMissingFromIndex(
                TryGetProfileStoreNames(),

                // No catalog short-circuit here: unlike the sweep, this path has no per-store
                // frontier map in hand, so every store is settled by its own probe.
                indexKnownStores: null,
                store => clock.ElapsedMilliseconds > StoreIndexProbeBudgetMs
                    ? (bool?)null
                    : StoreHasIndexRows(store));

            IReadOnlyList<string> unwalked = FreshMerge.UnwalkedUnindexedStores(live, missing);
            if (unwalked.Count == 0)
            {
                return;
            }

            live.StoresWithoutIndex = CapUnindexedStoreList(unwalked, out int total, out bool truncated);
            live.StoresWithoutIndexTruncated = truncated ? true : (bool?)null;
            live.StoresWithoutIndexTotal = truncated ? total : (int?)null;
        }

        /// <summary>
        /// One content-free token for a failed conversation walk (S4: never a subject, never
        /// a body). <c>OutlookUnavailableException</c> and <c>TimeoutException</c> carry
        /// messages the product wrote - they already name what failed and what was done
        /// about it - so those pass through; everything else becomes a HRESULT or a type
        /// name, and an exception carrying arbitrary text becomes <paramref name="fallback"/>.
        /// </summary>
        private static string DescribeWalkFailure(Exception ex, string fallback)
        {
            return ex switch
            {
                OutlookUnavailableException => ex.Message,
                TimeoutException => ex.Message,
                System.Runtime.InteropServices.COMException com =>
                    string.Format(CultureInfo.InvariantCulture, "COMException 0x{0:X8}", com.HResult),
                InvalidOperationException => fallback,
                _ => ex.GetType().Name,
            };
        }

        /// <summary>
        /// Best-effort staleness snapshot for a thread lookup. Best-effort because the
        /// conversation itself does not depend on it: it is context for a caller deciding
        /// how much to trust an index-only answer, so an unreachable index reports unknown
        /// rather than failing a lookup the COM walk may have answered completely.
        /// </summary>
        private StalenessInfo ProbeStaleness(string? scope)
        {
            DateTime? newestIndexed = null;
            double? ageMinutes = null;
            try
            {
                IndexStalenessReport staleness = _index.Value.GetStaleness(scope, SearchIndexTimeoutSeconds);
                newestIndexed = staleness.NewestIndexedReceivedUtc;
                ageMinutes = staleness.Age?.TotalMinutes;
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                // Unknown, reported as unknown (nulls are omitted from the payload).
            }

            return new StalenessInfo
            {
                NewestIndexedUtc = newestIndexed,
                AgeMinutes = ageMinutes,
                OutlookRunning = ComGateway.IsOutlookRunning(),
            };
        }

        /// <summary>
        /// The advice sentences for one thread lookup - one per way this conversation may be
        /// short of members, in the same severity-first order and the same voice the sweep's
        /// coverage advice uses. Null when there is nothing to say.
        /// <para>
        /// Public and pure for the same reason <see cref="DescribeSweepCoverage"/> is: the
        /// states it narrates need a real conversation spanning two stores, or one longer
        /// than the member cap, which no CI runner has - and a code that reaches the payload
        /// without a sentence is a partial answer an agent can see but cannot explain.
        /// </para>
        /// </summary>
        public static IReadOnlyList<string>? DescribeThreadCoverage(
            ThreadLiveInfo live,
            string freshness,
            string? store,
            bool scopeWidened,
            int top)
        {
            List<string> advice = new List<string>();
            if (scopeWidened)
            {
                // Two causes, one sentence, since 2026-08-18: the name may be wrong, or the
                // store may be real and simply not in the local index (a PST, indexing off).
                // Naming only the first told a caller with a correct name to go and check it.
                advice.Add("The store '" + (store ?? "?") + "' did not resolve to an index scope - the name may be "
                    + "wrong, or the store may be one Windows Search does not index - so this conversation was "
                    + "looked up across the WHOLE profile and the answer may include members from another account. "
                    + "list_accounts settles which: it lists the names, and reports inLocalIndex per store.");
            }

            if (freshness == FreshMerge.FreshnessIndexOnly)
            {
                const string lead = "INCOMPLETE CONVERSATION - TELL THE USER: these members are indexed results only";
                advice.Add(live.Error switch
                {
                    // The remedy is the caller's, so it is stated as an instruction.
                    "NoAnchorItem" => lead + ", so replies newer than the last index update are missing. Outlook's "
                        + "conversation graph can only be walked from a concrete mail, so call thread again with id set "
                        + "to any member's hit id to get the live check.",
                    "UnknownAnchorId" => lead + " and may be missing the newest replies: the id passed is neither a hit "
                        + "id from this session nor a full EntryID, so the live check had nothing to walk from. Hit ids "
                        + "expire when the server restarts - re-run the search and pass a fresh one.",
                    _ => lead + " and may be missing the newest replies. The live check against Outlook could not run ("
                        + (live.Error ?? "unknown") + "). Retry, or check outlook_health.",
                });
            }

            foreach (string gap in live.CoverageGaps ?? Array.Empty<string>())
            {
                switch (gap)
                {
                    case FreshMerge.ThreadGapUnwalkedStore:
                        advice.Add("The live check covered '" + (live.AnchorStore ?? "?")
                            + "' only - Outlook walks a conversation inside one store - while this conversation also has "
                            + "members in another account. Those are indexed results, so a reply that arrived there "
                            + "moments ago may be missing; call thread again with an id from that account to check it live.");
                        break;

                    case FreshMerge.ThreadGapUnindexedStore:
                        advice.Add("INCOMPLETE CONVERSATION - TELL THE USER: the local index holds no mail for "
                            + (DescribeUnindexedStoreList(
                                live.StoresWithoutIndex,
                                live.StoresWithoutIndexTruncated == true,
                                live.StoresWithoutIndexTotal,
                                "live.storesWithoutIndexTotal") ?? "part of this profile")
                            + ", and Outlook walks a conversation inside ONE store (here '"
                            + (live.AnchorStore ?? "?")
                            + "'), so a member sitting there is covered by NEITHER tier. Whether this conversation "
                            + "reaches into it cannot be established from here - it is not that the members were "
                            + "checked and found absent. Call thread again with an id from that account to walk it "
                            + "live, or search it with exhaustive:true.");
                        break;

                    case FreshMerge.ThreadGapMemberCap:
                        advice.Add("The live check stopped at top=" + top.ToString(CultureInfo.InvariantCulture)
                            + " members, and it reads the conversation in Outlook's own order rather than newest-first, "
                            + "so the members it did not reach could include recent ones. Raise top (max "
                            + ThreadTopCap.ToString(CultureInfo.InvariantCulture) + ") for full live coverage.");
                        break;

                    default:
                        // A code with no sentence would be a silent partial result, which is
                        // what all of this exists to remove. T1 pins that every code is
                        // handled, so this can only be reached by a code added without its
                        // advice - say so rather than dropping it.
                        advice.Add("The conversation walk reported partial coverage (" + gap
                            + ") with no further detail available; treat this thread as incomplete.");
                        break;
                }
            }

            return advice.Count > 0 ? advice : null;
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
                if (ShouldSearchOtherStores(storeId, d != null, error))
                {
                    // Direct EntryID without a known store, and ONLY when the item could not
                    // be OPENED. open_in_outlook is classified MUTATING (ComSessionOperations)
                    // because Display() puts a window on the user's screen and can mark the
                    // mail read. Under the old `d == null` guard a failing Display was retried
                    // against every store, so an item that WAS found could be displayed - or
                    // marked read - once per store before the call reported failure.
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        d = s.TryDisplayItem(entryId, store.StoreId, out error);
                        if (!KeepSearchingStores(d != null, error))
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
            //
            // THE FALLBACK IS GONE, and its absence is the guard. `perStore` used to also
            // require `result.PerStore.Count > 0`, which contradicted the paragraph above:
            // a store-scoped request over a result with NO per-store entries at all read
            // the whole sweep's totals, i.e. exactly the pre-c515565 cross-store
            // attribution this method exists to prevent. Unreachable today - both
            // ComSweepResult construction sites populate PerStore - so it was a latent seam
            // rather than a live defect, and it is closed by construction instead of by
            // comment: `store != null` alone decides, a missing entry answers zero, and
            // zeroes make DescribeCoverageGaps raise `nothing_swept` with degraded: true.
            // That is the loud, safe direction. The empty-PerStore case now says "no
            // coverage attributable to this store" rather than lending it another
            // account's, and a future third construction site that forgets PerStore fails
            // visibly instead of silently mis-attributing.
            ComStoreSweepCounters? scoped = store == null ? null : FindStoreCounters(result, store);
            bool perStore = store != null;

            info.FoldersSwept = perStore ? scoped?.FoldersSwept ?? 0 : result.FoldersSwept;
            info.FoldersSkipped = perStore ? scoped?.FoldersSkipped ?? 0 : result.FoldersSkipped;
            info.FoldersFailed = perStore ? scoped?.FoldersFailed ?? 0 : result.FoldersFailed;
            info.RowsUnreadable = perStore ? scoped?.RowsUnreadable ?? 0 : result.RowsUnreadable;

            // Gap G2, and attributed by the same rule as everything above it. A sweep that
            // was SCOPED never reaches an unnameable store (it can be neither matched nor
            // ruled out), so a non-zero count here can only have come from an all-stores
            // sweep - which the cache may well be serving to a store-scoped request. Reading
            // it there would report another store's naming failure inside this store's
            // answer, which is the cross-store leak the per-store buckets exist to prevent.
            info.StoresUnnamed = !perStore && result.StoresUnnamed > 0 ? result.StoresUnnamed : (int?)null;
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

            // Bound of the WHOLE sweep rather than of a subtree walk, so it is deliberately
            // NOT attributed per store: the budget is spent across every store the sweep
            // visited, and the stores it never reached are exactly the ones with no entry to
            // attribute it to. A store-scoped request served by a cached all-stores sweep
            // has to see this - its own zero coverage would otherwise read as "nothing was
            // there" instead of "we ran out of time before reaching you".
            info.SweepBudgetExpired = result.SweepBudgetExpired ? true : (bool?)null;

            List<string> capped = new List<string>();
            foreach (string entry in result.ItemCappedFolders)
            {
                if (InStoreScope(entry, store))
                {
                    capped.Add(entry);
                }
            }

            info.ItemCappedFolders = capped.Count == 0 ? null : capped;

            // The H2 subset, filtered by the SAME store rule as the list it is a subset of -
            // any other rule and the difference the advice is built from would stop being a
            // subset at all, and one of the two sentences would name a folder the other
            // never saw.
            List<string> arbitrary = new List<string>();
            foreach (string entry in result.ItemCappedFoldersUnsorted)
            {
                if (InStoreScope(entry, store))
                {
                    arbitrary.Add(entry);
                }
            }

            info.ItemCappedFoldersUnsorted = arbitrary.Count == 0 ? null : arbitrary;

            // NOT store-filtered, unlike the two lists above, and deliberately: it is a
            // count rather than a set of labels, so there is nothing to attribute, and the
            // question it answers - does Table.Sort work at all - is about the call and not
            // about any one store.
            info.SortRefusedFolders = result.SortRefusedFolders > 0 ? result.SortRefusedFolders : (int?)null;
        }

        /// <summary>
        /// Whether a <c>store/folder</c> sweep label belongs to the requested store, or to
        /// any store when the request named none. The store filter both capped-folder lists
        /// share, so the H2 subset cannot be filtered by a rule its superset was not.
        /// </summary>
        private static bool InStoreScope(string entry, string? store)
        {
            if (store == null)
            {
                return true;
            }

            int separator = entry.IndexOf('/');
            return separator >= 0
                && string.Equals(entry.Substring(0, separator), store, StringComparison.OrdinalIgnoreCase);
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
                if (ShouldSearchOtherStores(storeId, r != null, error))
                {
                    // Direct EntryID without a known store: retry across stores (same
                    // pattern as read/open_in_outlook), and ONLY when the source mail could
                    // not be OPENED.
                    //
                    // The guard used to be `r == null` alone, so any failure fanned out
                    // across every store on the profile. That is the one loop where a
                    // needless retry is not free: TryCreateDerivedDraft is classified as
                    // MUTATING for exactly this reason (ComSessionOperations) - a re-run
                    // can leave a second draft behind, and the caller only ever learns the
                    // id of the last one, so the earlier ones are orphaned where no cleanup
                    // will find them. A compose or Save failure could therefore leave up to
                    // one orphan per store and still report a failure.
                    // "ItemNotFound" is set before any item is created, so a retry under it
                    // repeats nothing. Every other failure now stops at the first attempt,
                    // which is also what the sibling loops in update_draft/discard_draft/
                    // move_mail do.
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        r = s.TryCreateDerivedDraft(entryId, store.StoreId, kind, toList, draftBody, display, signatureOverride, options, out error);
                        if (!KeepSearchingStores(r != null, error))
                        {
                            // A store that opened the item and then failed answers the
                            // question the loop is asking - the item is not missing - so
                            // trying the rest would create drafts for nothing.
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
            IReadOnlyList<string>? unresolved = CapUnresolvedRecipients(
                created.UnresolvedRecipients, out int unresolvedTotal, out bool unresolvedTruncated);

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
                UnresolvedRecipients = unresolved,
                UnresolvedRecipientsTruncated = unresolvedTruncated ? true : (bool?)null,
                UnresolvedRecipientsTotal = unresolvedTruncated ? unresolvedTotal : (int?)null,
                UnresolvedRecipientsAdvice = DescribeUnresolvedRecipientCap(unresolvedTotal, unresolvedTruncated),
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
            IReadOnlyList<string> attachmentPaths = files.Select(f => f.Path).ToList();

            // THE IDEMPOTENCE KEY. Derived here, from the request as it will reach Outlook -
            // never supplied by the caller (see DraftUpdateIntents.KeyFor). It is what makes
            // a repeat identifiable as a repeat rather than as a second, identical edit.
            string intentKey = DraftUpdateIntents.KeyFor(
                entryId, draftBody, subject?.Trim(), toList, ccList, bccList, parsedImportance,
                requestReadReceipt, signatureOverride, attachmentPaths, removeNames, display);
            ComDraftUpdateResume? resume = _updateIntents.Resume(intentKey, entryId);
            bool resumed = resume != null;
            string? resolvedStoreId = storeId;

            ComDraftUpdateResult updated;
            string? lastComError = null;
            try
            {
                updated = _gateway.Run(s =>
                {
                    // Record the intent BEFORE the first mutating call, because the whole
                    // point is that it outlives the process making that call. The pre-image
                    // read is the same READ the resume needs, costs at most two round trips
                    // on a ~20-call operation, and is skipped when nothing in the request is
                    // something a blind repeat could get wrong.
                    if (!resumed)
                    {
                        ComDraftUpdateResume? preImage = CapturePreImage(
                            s, entryId, ref resolvedStoreId, subject, attachmentPaths);
                        if (preImage != null)
                        {
                            _updateIntents.Begin(intentKey, entryId, preImage);
                        }
                    }

                    string? error = null;
                    ComDraftUpdateResult? r = s.TryUpdateDraft(
                        entryId, resolvedStoreId, draftBody, subject?.Trim(), toList, ccList, bccList,
                        parsedImportance, requestReadReceipt, signatureOverride,
                        attachmentPaths, removeNames, resume, display, out error);

                    if (ShouldSearchOtherStores(resolvedStoreId, r != null, error))
                    {
                        // This loop was unreachable until the COM layer began setting
                        // "ItemNotFound" on a failed open: TryUpdateDraft reported that failure
                        // as a bare "COMException 0x...", so a direct EntryID naming a draft
                        // outside the DEFAULT store never got its second attempt and the caller
                        // saw an opaque COM code instead. Now that it runs, it stops the moment
                        // a store OPENS the draft - update_draft appends attachments and can
                        // re-apply a signature, so carrying on past a store that answered would
                        // risk doing that twice. In practice the pre-image read above has
                        // already resolved the store by READING, so this is now the fallback
                        // for a draft no read could find rather than the ordinary path.
                        foreach (ComStoreDetail store in GetStoreDetails(s))
                        {
                            r = s.TryUpdateDraft(
                                entryId, store.StoreId, draftBody, subject?.Trim(), toList, ccList, bccList,
                                parsedImportance, requestReadReceipt, signatureOverride,
                                attachmentPaths, removeNames, resume, display, out error);
                            if (!KeepSearchingStores(r != null, error))
                            {
                                break;
                            }
                        }
                    }

                    lastComError = error;
                    return r ?? throw BuildDraftRefusal("update_draft", error, entryId);
                });
            }
            catch (Exception) when (string.Equals(lastComError, ComErrorTokens.ItemNotFound, StringComparison.Ordinal))
            {
                // The draft was never OPENED - by any store - so nothing was applied and there
                // is nothing to finish. It is the one failure outside DraftRefusedException
                // that still proves the negative, and without this it would be swallowed by
                // the unknown-outcome branch below and told to re-issue a call against an id
                // that does not resolve.
                _updateIntents.Settle(intentKey);
                throw;
            }
            catch (DraftRefusedException refusal) when (string.Equals(refusal.Reason, ComFailureRefusal, StringComparison.Ordinal))
            {
                // The ONE refusal that does not prove the draft was left alone: it comes from
                // the catch-all around the whole ~20-call sequence, so it can arrive after the
                // body was committed through the inspector or after an attachment went. The
                // intent therefore stays PENDING, and the refusal keeps its type - the tool
                // layer branches on that - while gaining the advice that fits what is known.
                throw new DraftRefusedException(
                    refusal.Reason, AuditUpdateOutcomeUnknown(intentKey, entryId, hitId, resumed, refusal));
            }
            catch (DraftRefusedException)
            {
                // Every OTHER refusal is proof of the negative: each is decided before anything
                // is written (BuildDraftRefusal), and the one that is not - a body replace that
                // failed - discards its own inspector so the draft survives untouched. So the
                // intent is settled: there is nothing left to complete.
                _updateIntents.Settle(intentKey);
                throw;
            }
            catch (Exception ex) when (ex is not ArgumentException)
            {
                // Everything else - the deadline expiring and the COM host being killed, the
                // pipe dying under a sibling request, an unclassified COM failure part-way
                // through the sequence - leaves the intent PENDING on purpose. Nobody can
                // state what was applied, so the record has to outlive the failure for the
                // repeat that can finish it.
                throw new InvalidOperationException(
                    AuditUpdateOutcomeUnknown(intentKey, entryId, hitId, resumed, ex), ex);
            }

            _updateIntents.Settle(intentKey);

            // EntryIDs are not stable - re-key the registry so a following discard_draft
            // still recognises the draft this call just rewrote (D46/C2).
            _draftRegistry.Replace(entryId, updated.Draft.EntryId);
            if (!string.Equals(entryId, updated.Draft.EntryId, StringComparison.OrdinalIgnoreCase))
            {
                // A pre-image is addressed by EntryID, and a re-keyed draft is no longer at
                // that address - keeping one would offer a resume over an id that is gone.
                _updateIntents.Forget(entryId);
            }

            AuditUpdate(updated, draftBody, hitId, resumed);
            return ToUpdateOutcome(updated, hitId, draftBody, htmlAdjustments, removeNames, resumed);
        }

        /// <summary>The one refusal code that does NOT prove the draft was left alone.</summary>
        internal const string ComFailureRefusal = "com_failure";

        /// <summary>
        /// Reads what a repeat of this update would otherwise be unable to know, BEFORE the
        /// first attempt writes anything. Null means no pre-image could be taken, and no
        /// intent is then recorded at all - a resume this server cannot vouch for is worse
        /// than no resume.
        /// <para>
        /// Both halves are conditional, because each is only needed by a step a blind repeat
        /// would get wrong. The conversation index/topic matter only when the subject
        /// changes, since assigning Subject is what makes Outlook regenerate the index. The
        /// attachment names matter only when files are being attached, because that is the
        /// one case where "is this file already on?" cannot be answered from the request
        /// alone: by name, the copy the first attempt added and a copy the draft always had
        /// are the same thing.
        /// </para>
        /// <para>
        /// It doubles as the store resolver. A bare EntryID whose draft lives outside the
        /// default store is now found by a READ, so the mutating call goes straight to the
        /// right store instead of being offered to each store in turn.
        /// </para>
        /// </summary>
        private ComDraftUpdateResume? CapturePreImage(
            IOutlookSession session,
            string entryId,
            ref string? storeId,
            string? subject,
            IReadOnlyList<string> attachmentPaths)
        {
            if (subject == null && attachmentPaths.Count == 0)
            {
                // Nothing in this request is order-coupled or accumulating, so every step of
                // it is safe to replay and an empty pre-image is the whole truth.
                return new ComDraftUpdateResume();
            }

            string? error = null;
            ComDraftInfo? info = session.TryGetMailInfo(entryId, storeId, out error);
            if (info == null && ShouldSearchOtherStores(storeId, false, error))
            {
                foreach (ComStoreDetail store in GetStoreDetails(session))
                {
                    info = session.TryGetMailInfo(entryId, store.StoreId, out error);
                    if (!KeepSearchingStores(info != null, error))
                    {
                        if (info != null)
                        {
                            storeId = store.StoreId;
                        }

                        break;
                    }
                }
            }

            if (info == null)
            {
                return null;
            }

            storeId ??= info.StoreId;

            // An attachment snapshot that FAILED and one of a draft with no attachments are
            // both an empty list - the contract has no error channel here. The consequence is
            // bounded and it falls the safe way: only a RESUME reads this list (a first
            // attempt uses what it sees on the live item), and an understated pre-image makes
            // a resume attach too little rather than twice. That is visible in the attachment
            // list the call returns and fixable by attaching again; the opposite mistake is a
            // silent duplicate.
            IReadOnlyList<string> attachmentNames = attachmentPaths.Count == 0
                ? Array.Empty<string>()
                : session.SnapshotAttachmentsById(info.EntryId, storeId)
                    .Select(a => a.FileName ?? string.Empty)
                    .ToList();

            return new ComDraftUpdateResume(
                attachmentNames,
                subject == null ? null : info.ConversationIndex,
                subject == null ? null : info.ConversationTopic);
        }

        /// <summary>
        /// Records an update whose outcome nobody can state, and builds the message the
        /// caller gets. The same shape as the send path's <c>send_outcome_unknown</c> and
        /// for the same reason: the audit ordering is mutate-then-record, so a kill leaves a
        /// MISSING line rather than a wrong one, and a missing line is a gap nothing looks
        /// for. This states it instead.
        /// <para>
        /// The append is best-effort, as it is for a send: the operation this line describes
        /// has already happened or already failed, and replacing the one message the caller
        /// needs with a message about a log file would be the wrong trade.
        /// </para>
        /// </summary>
        private string AuditUpdateOutcomeUnknown(string intentKey, string entryId, string? hitId, bool wasResume, Exception cause)
        {
            bool resumable = _updateIntents.Resume(intentKey, entryId) != null;
            try
            {
                Audit.AuditLog.Append(
                    "update_draft_outcome_unknown",
                    ("entryId", entryId),
                    ("hitId", hitId),
                    ("resumable", resumable ? "true" : "false"),
                    ("wasResume", wasResume ? "true" : null),
                    ("reason", cause.GetType().Name));
            }
            catch (InvalidOperationException)
            {
                // outlook_health reports an unwritable audit log; this message must survive it.
            }

            return DescribeUpdateOutcomeUnknown(cause.Message, resumable);
        }

        /// <summary>
        /// What a caller is told when an update did not answer. Pure and public so T1 pins
        /// it: the words are the whole point, and the path cannot be exercised without a real
        /// Outlook wedged mid-sequence.
        /// <para>
        /// The advice INVERTS when an intent was recorded, and that inversion is the whole of
        /// the re-entrancy work as the caller experiences it. Without a record the honest
        /// answer is the one every killed mutation gets - the outcome is unknown, look before
        /// you act. With one, the identical call IS the remedy: it converges on the end state
        /// the first attempt was aiming at instead of performing it again, so files it had
        /// already removed are not lost and files it had already attached are not doubled.
        /// </para>
        /// </summary>
        public static string DescribeUpdateOutcomeUnknown(string causeMessage, bool resumable)
        {
            string opening = "The draft revision did not answer, so WHETHER IT WAS APPLIED IS UNKNOWN: update_draft is a "
                + "sequence of steps inside Outlook and it may have completed only some of them. ";
            string remedy = resumable
                ? "RE-ISSUE THIS EXACT CALL - same id, same arguments - and it will FINISH what the first attempt started "
                + "rather than repeat it: a file it had already attached is not attached twice, a file it had already "
                + "removed stays removed, and the draft's place in its conversation is restored from what was recorded "
                + "before the first attempt ran. Change any argument and it becomes a NEW update instead, so read the "
                + "draft first if you want something different. "
                : "Do NOT simply retry it: nothing was recorded that a retry could resume from, so read the draft and "
                + "decide from what it actually contains. ";

            return opening + remedy + "(Underlying failure: " + causeMessage + ")";
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
                if (ShouldSearchOtherStores(storeId, r != null, error))
                {
                    // Unreachable until the COM layer began setting "ItemNotFound" here too
                    // (see update_draft above). It stops at the first store that opens the
                    // draft: this is the one mail-deleting path in the product, and a store
                    // that answered with a refusal has answered - the draft is not missing.
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        r = s.TryDiscardDraft(entryId, store.StoreId, out error);
                        if (!KeepSearchingStores(r != null, error))
                        {
                            break;
                        }
                    }
                }

                return r ?? throw BuildDraftRefusal("discard_draft", error, entryId);
            });

            AuditDiscard(discarded, hitId);
            _draftRegistry.Forget(entryId);

            // The draft is in Deleted Items and its EntryID is dead, so any pre-image held for
            // it describes an item at an address nothing will resolve. Dropping it here keeps
            // the resume offer from outliving the thing it would complete.
            _updateIntents.Forget(entryId);
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
                    // "Nothing was changed" is true of every NAMED refusal above and was not
                    // true here: an unclassified COM failure is the one that can arrive
                    // part-way through the sequence, after the body has been committed
                    // through the inspector or after an attachment has gone. update_draft
                    // therefore says what it actually knows, and points at the repeat that
                    // can finish the job; discard_draft keeps its own wording because its
                    // sequence is a different shape.
                    return RefuseDraft(ComFailureRefusal, operation, entryId,
                        operation == "discard_draft"
                            ? "The draft could not be discarded (" + (comError ?? "unknown")
                                + "). Nothing was changed. Check outlook_health and retry."
                            : "Outlook failed part-way through the revision (" + (comError ?? "unknown")
                                + "). Check outlook_health.");
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

        private static void AuditUpdate(ComDraftUpdateResult updated, ComDraftBody? body, string? hitId, bool resumed)
        {
            try
            {
                Audit.AuditLog.Append(
                    "update_draft",
                    ("entryId", updated.Draft.EntryId),
                    ("hitId", hitId),
                    ("resumed", resumed ? "true" : null),
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
            IReadOnlyList<string> requestedRemovals,
            bool resumed)
        {
            IReadOnlyList<RecipientView> recipients = CapRecipients(updated.Draft.Recipients, out int total, out bool truncated);
            IReadOnlyList<string>? unresolved = CapUnresolvedRecipients(
                updated.UnresolvedRecipients, out int unresolvedTotal, out bool unresolvedTruncated);
            IReadOnlyList<AttachmentView> attachmentViews = CapAttachments(
                VerifiedAttachments(updated.Draft, updated.Attachments), out int _, out bool _);
            IReadOnlyList<string>? notRemoved = requestedRemovals
                .Where(n => !updated.AttachmentsRemoved.Any(r => string.Equals(r, n, StringComparison.OrdinalIgnoreCase)))
                .ToList();

            return new UpdateDraftOutcome
            {
                Status = "updated",
                Resumed = resumed ? true : (bool?)null,
                ResumedAdvice = resumed ? ResumedUpdateAdvice : null,
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
                UnresolvedRecipients = unresolved,
                UnresolvedRecipientsTruncated = unresolvedTruncated ? true : (bool?)null,
                UnresolvedRecipientsTotal = unresolvedTruncated ? unresolvedTotal : (int?)null,
                UnresolvedRecipientsAdvice = DescribeUnresolvedRecipientCap(unresolvedTotal, unresolvedTruncated),
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
        /// <summary>
        /// What a caller is told when this update FINISHED an earlier one rather than
        /// performing a fresh revision. It matters because the two look identical in the
        /// payload otherwise, and the difference decides what "changed" means: the fields
        /// listed are the end state this call converged on, not necessarily the writes it
        /// made itself.
        /// </summary>
        internal const string ResumedUpdateAdvice =
            "This call COMPLETED an earlier update_draft whose outcome was unknown - the COM host ended before it "
            + "answered - rather than revising the draft a second time. Attachments were reconciled against what the "
            + "draft actually held, and the conversation index was restored from the state recorded before that first "
            + "attempt, so the draft is in the state the original request asked for. Nothing was applied twice.";

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
                if (ShouldSearchOtherStores(storeId, st != null, error))
                {
                    // Direct EntryID without a known store, and ONLY when the draft could not
                    // be OPENED. The read itself is harmless to repeat, but the answer is
                    // not: "not a mail item" is a verdict about the draft that WAS found, and
                    // re-asking it store by store used to end on whichever store failed last,
                    // so the refusal the caller finally saw named the wrong reason.
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        st = s.TryGetSendableDraftState(entryId, store.StoreId, out error);
                        if (!KeepSearchingStores(st != null, error))
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
            ComSendResult? sent;
            try
            {
                sent = _gateway.Run(s => s.TrySendDraft(state.EntryId, state.StoreId, contentHash, sentOnBehalfOf, out sendError));
            }
            catch (TimeoutException ex)
            {
                // THE ONE PATH THAT MOST NEEDS THE OUTBOX WARNING WAS THE ONE NOT GETTING
                // IT. A deadline expiry here kills the COM host somewhere between
                // MailItem.Send() executing inside Outlook and the answer reaching us, so
                // the mail may already have been submitted - Outlook creates and submits a
                // message in a folder, usually the Outbox - and the draft's EntryID is gone
                // with it. The neighbouring SendCallFailed branch has said "The mail MAY be
                // sitting in the Outbox - verify before retrying" since it was written; the
                // kill path handed back the generic "Outlook did not respond ... the COM
                // host was restarted" instead, which says nothing about mail at all.
                //
                // The confirm token is already consumed at this point (it is consumed in
                // this process, before the child is even asked), so the friction is
                // pointing the wrong way too: re-confirming after an unknown-outcome send
                // is exactly how a duplicate gets sent. Hence the explicit instruction to
                // look before re-sending.
                throw new InvalidOperationException(AuditSendOutcomeUnknown(state, hitId, ex), ex);
            }

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

        /// <summary>
        /// Records a send whose outcome nobody can state, and builds the message the caller
        /// gets. Both halves, in one place, because they are one fact.
        /// <para>
        /// The audit trail was one line away from being a real detector here. Its ordering
        /// is mutate-then-record, so a kill produces a MISSING line, never a corrupt one -
        /// which meant a killed send left <c>send_token_issued</c> with no matching
        /// <c>send</c> and no <c>send_refused</c>, and nothing in the product looked for
        /// that shape. <c>send_outcome_unknown</c> turns the gap into a record: the same
        /// diagnosis, stated rather than inferred.
        /// </para>
        /// <para>
        /// The audit write is best-effort here, unlike everywhere else in the send path. A
        /// failed append normally refuses the operation (D4: no send without its line), but
        /// the operation this describes has ALREADY happened or already failed; throwing an
        /// audit error over it would replace the one message the caller most needs with a
        /// message about a log file.
        /// </para>
        /// </summary>
        private static string AuditSendOutcomeUnknown(ComSendableDraftState state, string? hitId, Exception cause)
        {
            try
            {
                Audit.AuditLog.Append(
                    "send_outcome_unknown",
                    ("entryId", state.EntryId),
                    ("store", state.StoreDisplayName),
                    ("account", state.ResolvedAccountSmtp),
                    ("recipients", state.Recipients.Count.ToString(CultureInfo.InvariantCulture)),
                    ("reason", cause.GetType().Name),
                    ("hitId", hitId));
            }
            catch (InvalidOperationException)
            {
                // The caller-facing message below is the important half and it is built
                // regardless. An unwritable audit log is reported by outlook_health.
            }

            return DescribeSendOutcomeUnknown(cause.Message);
        }

        /// <summary>
        /// What a caller is told when a send did not answer and the COM host was reclaimed.
        /// Pure and public so T1 pins it: the whole point of this message is a set of words
        /// that must be present, and the path it belongs to cannot be exercised without
        /// wedging a real Outlook mid-send.
        /// <para>
        /// It says what the <c>SendCallFailed</c> branch beside it has always said - "the
        /// mail MAY be sitting in the Outbox" - because the two describe the same state.
        /// Only one of them used to say it, and it was the branch that fires when Outlook
        /// ANSWERS with an error; the kill path, where the caller knows least, got the
        /// generic "Outlook did not respond ... the COM host was restarted", which mentions
        /// no mail at all.
        /// </para>
        /// </summary>
        public static string DescribeSendOutcomeUnknown(string causeMessage)
        {
            return "Outlook did not answer the send within its budget, and the COM host was restarted to reclaim the call. "
                + "WHETHER THE MAIL WAS SENT IS UNKNOWN: the send may have executed inside Outlook without the answer "
                + "reaching us, in which case the mail is on its way or MAY BE SITTING IN THE OUTBOX and will go out when "
                + "Outlook next runs. Do NOT simply send again - check Sent Items and the Outbox for this message first, "
                + "and only re-create and re-send if it is in neither. The confirm token is spent either way. (Underlying "
                + "failure: " + causeMessage + ")";
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

                int itemBudgetMs = RemainingBatchBudgetMs(batchClock);
                if (itemBudgetMs < MinimumItemBudgetMs)
                {
                    items.Add(FailedItem(id, BatchBudgetExhaustedMessage));
                    continue;
                }

                MoveItemView item = MoveOne(
                    id, segments, createFolder, requiredStore, targetFolderEcho, createdFolders, itemBudgetMs,
                    out bool auditFailed);
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

                int itemBudgetMs = RemainingBatchBudgetMs(batchClock);
                if (itemBudgetMs < MinimumItemBudgetMs)
                {
                    items.Add(FailedItem(id, BatchBudgetExhaustedMessage));
                    continue;
                }

                MoveItemView item = ArchiveOne(id, archiveByStore, itemBudgetMs, out bool auditFailed);
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

        /// <summary>
        /// What is left of <see cref="MoveBatchBudgetMs"/>, which is the budget the NEXT
        /// item is dispatched with.
        /// <para>
        /// Handing it to the item is what makes the batch budget an aggregate rather than a
        /// label. Without it the check before each item was the only bound, and each item
        /// then ran under a full operation deadline of its own - so a batch sitting one
        /// millisecond inside the budget could start one more item and run to the budget
        /// plus a whole extra deadline.
        /// </para>
        /// </summary>
        private static int RemainingBatchBudgetMs(Stopwatch batchClock)
        {
            long remaining = MoveBatchBudgetMs - batchClock.ElapsedMilliseconds;
            return remaining > 0 ? (int)remaining : 0;
        }

        /// <summary>
        /// Per-item reason when an item's own budget expired mid-move. It is a MUTATION
        /// whose outcome the caller cannot infer: the COM host was ended to reclaim the
        /// call, so the move may or may not have happened, and re-issuing it blindly is how
        /// a caller ends up hunting an item that already moved.
        /// </summary>
        internal static string BatchItemTimedOutMessage(string detail)
        {
            return "The move did not answer within the batch's remaining time budget (" + detail
                + "). Whether it took effect is UNKNOWN - the COM host was ended to reclaim the call. Check where the item "
                + "is now (search for it, or read it) before re-issuing, and re-issue the rest as a smaller batch.";
        }

        private MoveItemView ArchiveOne(
            string id,
            Dictionary<string, (ComArchiveFolderInfo? Info, string? Error)> archiveByStore,
            int itemBudgetMs,
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
                ComDraftInfo? info = _gateway.Run(
                    s =>
                {
                    ComDraftInfo? r = s.TryGetMailInfo(entryId, storeId, out string? infoError);
                    if (ShouldSearchOtherStores(storeId, r != null, infoError))
                    {
                        // Direct EntryID without a known store, and ONLY when the item could
                        // not be OPENED - the same rule the move loop below already follows,
                        // which matters because archive_mail runs both in sequence over the
                        // same id and they disagreed about when to fan out.
                        foreach (ComStoreDetail candidate in GetStoreDetails(s))
                        {
                            r = s.TryGetMailInfo(entryId, candidate.StoreId, out infoError);
                            if (!KeepSearchingStores(r != null, infoError))
                            {
                                break;
                            }
                        }
                    }

                    return r;
                },
                    itemBudgetMs,
                    allowConnectFloor: true);
                if (info?.StoreDisplayName == null)
                {
                    return FailedItem(id, "The item could not be opened. Re-run search - it may have moved (EntryIDs change on moves).");
                }

                if (!archiveByStore.TryGetValue(info.StoreDisplayName, out (ComArchiveFolderInfo? Info, string? Error) archive))
                {
                    archive = _gateway.Run(
                        s =>
                        {
                            ComArchiveFolderInfo? resolvedInfo = s.TryResolveArchiveFolder(info.StoreDisplayName, out string? resolveError);
                            return (resolvedInfo, resolveError);
                        },
                        itemBudgetMs,
                        allowConnectFloor: true);
                    archiveByStore[info.StoreDisplayName] = archive;
                }

                if (archive.Info == null)
                {
                    return FailedItem(id, DescribeArchiveResolutionFailure(info.StoreDisplayName, archive.Error));
                }

                ComArchiveFolderInfo target = archive.Info;
                (ComMoveItemResult? moved, string? moveError) = _gateway.Run(
                    s =>
                    {
                        ComMoveItemResult? r = s.TryMoveItemToFolderId(entryId, info.StoreId ?? storeId, target.EntryId, target.StoreId, out string? e);
                        return (r, e);
                    },
                    itemBudgetMs,
                    allowConnectFloor: true);
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
            catch (TimeoutException ex)
            {
                // Per item, not per batch. The item's own budget is what expired, so the
                // rest of the batch is still worth attempting - and the item is a mutation
                // whose outcome nobody can now state, which the message has to say.
                return FailedItem(id, BatchItemTimedOutMessage(ex.Message));
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
            int itemBudgetMs,
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
                (ComMoveItemResult? moved, string? comError) = _gateway.Run(
                    s =>
                {
                    ComMoveItemResult? r = s.TryMoveItemToPath(entryId, storeId, segments, createFolder, requestedStore, out string? e);
                    if (ShouldSearchOtherStores(storeId, r != null, e))
                    {
                        // Direct EntryID without a known store: retry across stores
                        // (same pattern as read/draft ops).
                        foreach (ComStoreDetail candidate in GetStoreDetails(s))
                        {
                            r = s.TryMoveItemToPath(entryId, candidate.StoreId, segments, createFolder, requestedStore, out e);
                            if (!KeepSearchingStores(r != null, e))
                            {
                                break;
                            }
                        }
                    }

                    return (r, e);
                },
                    itemBudgetMs,
                    allowConnectFloor: true);

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
            catch (TimeoutException ex)
            {
                // Per item, not per batch - see ArchiveOne's twin.
                return FailedItem(id, BatchItemTimedOutMessage(ex.Message));
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
        /// Caps the UNRESOLVABLE address list at <see cref="UnresolvedRecipientsCap"/> and
        /// says what the cap hid - the has-more pair <see cref="CapRecipients"/> two
        /// properties away has always carried and this list did not.
        /// <para>
        /// WHY IT MATTERS MORE HERE THAN ON THE RESOLVED LIST, which is what made a silent
        /// <c>Take(20)</c> the wrong shape on the drafting surface. The resolved list is
        /// context: the draft holds the recipients whether or not the payload lists them all,
        /// and the operations that matter (the send hash, the identity checks, transport)
        /// always read the full COM-side list. An UNRESOLVED address is the opposite - it is
        /// a defect the agent is expected to ACT on, by asking the user about each one, and a
        /// short list read as complete means the addresses past the cap are never mentioned
        /// to anybody. The draft then goes out, or fails to, with a fault its own report
        /// declared absent.
        /// </para>
        /// <para>
        /// Returns null for an empty list, because that is what the payload wants: an absent
        /// field for "nothing failed to resolve" rather than an empty array to be
        /// interpreted. <paramref name="total"/> and <paramref name="truncated"/> are still
        /// set, so a caller never has to infer either from the returned list's length - which
        /// is exactly the inference this pair exists to remove.
        /// </para>
        /// </summary>
        public static IReadOnlyList<string>? CapUnresolvedRecipients(
            IReadOnlyList<string>? unresolved, out int total, out bool truncated)
        {
            total = unresolved?.Count ?? 0;
            truncated = total > UnresolvedRecipientsCap;
            if (total == 0)
            {
                return null;
            }

            return unresolved!.Take(truncated ? UnresolvedRecipientsCap : total).ToList();
        }

        /// <summary>
        /// What a caller is told when the unresolvable-address list was cut, or null when it
        /// was not. Pure, and it reads the SAME two values the payload fields are set from,
        /// so the count in the sentence and the count in the field cannot disagree.
        /// <para>
        /// It quotes <see cref="UnresolvedRecipientsCap"/> rather than restating 20: a cap
        /// named in prose beside a constant nothing compares it with is how the two drift
        /// (the shape <c>SubjectCharsCap</c> was fixed into).
        /// </para>
        /// </summary>
        public static string? DescribeUnresolvedRecipientCap(int total, bool truncated)
        {
            if (!truncated)
            {
                return null;
            }

            return "Outlook could not resolve " + total.ToString(CultureInfo.InvariantCulture)
                + " of this draft's addresses, and unresolvedRecipients names only the first "
                + UnresolvedRecipientsCap.ToString(CultureInfo.InvariantCulture) + " - the other "
                + (total - UnresolvedRecipientsCap).ToString(CultureInfo.InvariantCulture)
                + " are on the draft and are NOT listed here. Do not tell the user this list is all of them. "
                + "Every one of them stays on the draft and will fail on send; open_in_outlook shows the user "
                + "the whole recipient line, and re-setting the recipients with update_draft in batches of "
                + UnresolvedRecipientsCap.ToString(CultureInfo.InvariantCulture)
                + " or fewer names each batch's failures in full.";
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

            if (comHost.FramesRefusedTooLarge > 0)
            {
                // A refusal is a request that got no answer, so it belongs with the
                // problems even though nothing is broken: the caller must ask for LESS,
                // and a retry of the same request refuses again. The high-water mark
                // beside it in the payload is a measurement and stays out of here.
                problems.Add("A message on the Outlook connection was too large to send and was refused "
                    + comHost.FramesRefusedTooLarge.Value.ToString(CultureInfo.InvariantCulture)
                    + " time(s) this session (limit "
                    + (comHost.FrameLimitBytes / (1024 * 1024))?.ToString(CultureInfo.InvariantCulture)
                    + " MB) - in practice an answer carrying too much mail. Narrow the request that caused it "
                    + "- a shorter time window, fewer folders, or a smaller limit - rather than retrying it "
                    + "unchanged.");
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

                    // Gap G2: this tool is where an agent is told what the 'store' argument
                    // may be, so a label printed here without the flag beside it would be
                    // read as a usable store name - and it is the one name that is not.
                    NameUnreadable = store.NameUnreadable ? true : (bool?)null,
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

            ComFolderTree tree = _gateway.Run(s => s.ListFolders(store, FolderWalkAbsoluteCap));

            // Gap G1. A store name that matched nothing produced folderTotal: 0,
            // truncated: false and no error - indistinguishable from a store that is there
            // and empty, so a typo came back looking like an answer. search has always
            // refused an unresolvable store loudly; this was the odd one out, and the fix is
            // that same refusal rather than a second shape.
            //
            // Only on the EMPTY result, so the ordinary call pays nothing: the store list is
            // fetched (from a cache with a 5-minute TTL) exactly when there is a question to
            // answer. An empty tree from a store that IS in the list stays an empty tree with
            // no error - a store whose root has no subfolders is a real thing, and inventing
            // a failure for it would be the mirror of the defect.
            if (store != null && tree.Folders.Count == 0)
            {
                IReadOnlyList<ComStoreDetail> details = _gateway.Run(s => GetStoreDetails(s));
                string? refusal = DescribeUnresolvedFolderStore(
                    store,
                    details.Where(d => !d.NameUnreadable).Select(d => d.DisplayName).ToList(),
                    details.Count(d => d.NameUnreadable));
                if (refusal != null)
                {
                    throw new ArgumentException(refusal, nameof(store));
                }
            }

            return PageFolders(tree, offset);
        }

        /// <summary>
        /// What to say when the freshness sweep covered a store Outlook would not name (gap
        /// G2), or null when every store named itself.
        /// <para>
        /// It raises no coverage code and does not degrade the search on purpose: the store
        /// IS swept, so nothing is missing from the answer. What the caller loses is the
        /// ability to ASK about it again - every scope in this server is keyed by a display
        /// name - and that is a fact about the next call rather than a hole in this one.
        /// </para>
        /// </summary>
        public static string? DescribeUnnamedStores(int? storesUnnamed)
        {
            if (storesUnnamed == null || storesUnnamed <= 0)
            {
                return null;
            }

            return storesUnnamed.Value.ToString(CultureInfo.InvariantCulture)
                + " store(s) in this profile would not report a display name to Outlook. Their mail IS in this answer "
                + "(the freshness sweep covered them under '" + StoreNaming.UnnamedStorePrefix
                + "N)' labels), but they cannot be used as a 'store' scope, and hits from them carry that label "
                + "instead of a real store name. list_folders shows their folder trees.";
        }

        /// <summary>
        /// What to say when a delegate folder scope was built from a folder walk that hit its
        /// own bound (gap G4), or null when it was not.
        /// <para>
        /// The remedy is the shape of the defect: the delegate index namespace is flat, so
        /// the scope is an OR of folder NAMES, and a name the walk never reached is a folder
        /// no tier looks in. Naming one folder at a time re-walks a smaller tree; exhaustive
        /// bypasses the index and its name matching entirely.
        /// </para>
        /// </summary>
        public static string? DescribeTruncatedFolderNames(FolderScopeResolution? folderScope)
        {
            if (folderScope == null || !folderScope.FolderTreeTruncated)
            {
                return null;
            }

            return "INCOMPLETE SCOPE - TELL THE USER: this shared/delegate mailbox's folder tree could not be walked "
                + "in full (it hit the " + FolderWalkAbsoluteCap.ToString(CultureInfo.InvariantCulture)
                + "-folder walk cap or the "
                + OutlookComSession.FolderWalkDepthGuard.ToString(CultureInfo.InvariantCulture)
                + "-level depth guard), and a delegate folder scope can only be matched by folder "
                + "NAME - so folders the walk never reached were searched by no tier. Narrow to a single folder with "
                + "include_subfolders:false, or use exhaustive:true with store plus folder/after bounds.";
        }

        /// <summary>
        /// Why a <c>list_folders</c> store scope resolved to nothing, or null when the empty
        /// tree is the honest answer (gap G1). Pure, and public so T1 pins both verdicts -
        /// reaching them for real needs a live profile.
        /// <para>
        /// The message deliberately matches the shape the index-tier store resolver already
        /// throws: what was asked for, what exists, and the tool that lists the rest. Two
        /// tools refusing the same mistake in two different shapes is how an agent learns to
        /// handle one and not the other.
        /// </para>
        /// <para>
        /// <paramref name="knownStores"/> comes from the same COM enumeration
        /// <c>ListFolders</c> itself walks, so the two agree by construction. It holds the
        /// stores that reported a NAME; <paramref name="unnamedStores"/> counts the rest,
        /// and it is what stops this refusal asserting more than it knows. Until gap G2 was
        /// closed such a store was absent from every list in this server, so "not found in
        /// Outlook" was said with certainty about a profile that might well have held it
        /// under a name nothing could read.
        /// </para>
        /// </summary>
        /// <param name="requestedStore">The store scope as the caller wrote it.</param>
        /// <param name="knownStores">Stores Outlook reported a display name for.</param>
        /// <param name="unnamedStores">
        /// How many stores Outlook has whose display name could not be read (gap G2). Any of
        /// them could be the requested one, so a non-zero count turns the refusal from a
        /// verdict into a report of what could and could not be established.
        /// </param>
        public static string? DescribeUnresolvedFolderStore(
            string requestedStore, IReadOnlyList<string>? knownStores, int unnamedStores = 0)
        {
            if (requestedStore == null)
            {
                throw new ArgumentNullException(nameof(requestedStore));
            }

            if (StoreNaming.IsUnnamedStoreLabel(requestedStore))
            {
                // The caller read a label out of a payload and passed it back as a scope,
                // which is the one wrong turn this label makes possible - so it gets the
                // answer that fits, rather than the "check for a typo" a generic refusal
                // would send it hunting for. A scope is matched against DisplayName, and
                // DisplayName is exactly what this store would not report.
                return "'" + requestedStore + "' is a placeholder this server prints for a store whose display name "
                    + "Outlook would not report - not a name, and it cannot be used as a scope, because a store scope "
                    + "is matched against the display name that could not be read. Search without 'store' to include "
                    + "that store's mail, or use exhaustive:true with a folder bound.";
            }

            IReadOnlyList<string> known = knownStores ?? Array.Empty<string>();
            foreach (string candidate in known)
            {
                if (string.Equals(candidate, requestedStore, StringComparison.OrdinalIgnoreCase))
                {
                    // It is there and it has no folders below its root. Nothing was lost, so
                    // nothing is reported.
                    return null;
                }
            }

            return "Store '" + requestedStore + "' was not found in Outlook. "
                + (known.Count > 0
                    ? "Known stores: " + string.Join(", ", known) + ". "
                    : "Outlook reported no stores at all. ")
                + (unnamedStores > 0
                    ? unnamedStores.ToString(CultureInfo.InvariantCulture)
                        + " further store(s) would not report a display name, so this one cannot be ruled out among "
                        + "them; list_folders without 'store' lists them under '" + StoreNaming.UnnamedStorePrefix
                        + "N)' labels. "
                    : string.Empty)
                + "Use list_accounts for the full store list.";
        }

        /// <summary>
        /// Pure paging step of list_folders (public for T1): slices the stable-order
        /// flattened walk at [offset, offset + <see cref="FoldersPerCallCap"/>) and
        /// derives the has-more contract (truncated + nextOffset + total).
        /// </summary>
        public static FoldersOutcome PageFolders(ComFolderTree tree, int offset)
        {
            if (tree == null)
            {
                throw new ArgumentNullException(nameof(tree));
            }

            if (offset < 0)
            {
                offset = 0;
            }

            IReadOnlyList<ComFolderInfo> folders = tree.Folders;
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

                    // Gap G2: the label is in the payload, so what it IS has to be there
                    // beside it - an agent that reads it as a store name will pass it back
                    // as a scope and be refused for a reason it has no way to guess.
                    NameUnreadable = StoreNaming.IsUnnamedStoreLabel(g.Key) ? true : (bool?)null,
                    Folders = g.Select(f => new FolderView
                    {
                        Path = f.Path,
                        Items = f.ItemCount,
                        Unread = f.UnreadCount,
                    }).ToList(),
                })
                .ToList();

            // Gap G3. This was `end < folders.Count` alone - computed against the list the
            // WALK had already truncated, so the one truncation it could never see was the
            // one that lost whole folders rather than merely deferring them to a later page.
            bool morePages = end < folders.Count;
            bool walkCut = tree.WalkCapReached || tree.DepthLimitReached;
            return new FoldersOutcome
            {
                Stores = byStore,
                FolderTotal = folders.Count,
                Offset = offset > 0 ? offset : (int?)null,
                Truncated = morePages || walkCut,

                // Only the pageable half gets a continuation: the next call re-walks and
                // stops at the same cap, so offering an offset past a walk cut would be an
                // instruction that cannot work.
                NextOffset = morePages ? end : (int?)null,
                WalkCapReached = tree.WalkCapReached ? true : (bool?)null,
                DepthLimitReached = tree.DepthLimitReached ? true : (bool?)null,
                StoresUnnamed = tree.StoresUnnamed > 0 ? tree.StoresUnnamed : (int?)null,
                StoresUnnamedExcluded = tree.StoresUnnamedExcluded > 0 ? tree.StoresUnnamedExcluded : (int?)null,
                Advice = DescribeFolderListingCoverage(
                    tree.WalkCapReached, tree.DepthLimitReached, tree.StoresUnnamed, tree.StoresUnnamedExcluded),
            };
        }

        /// <summary>
        /// The prose half of what <c>list_folders</c> could not deliver - one sentence per
        /// fact, each naming the remedy, in the same relationship to the flags beside it that
        /// <see cref="DescribeSweepCoverage"/> has to the sweep's coverage codes: the flags
        /// are what software branches on, these are what a person is told.
        /// <para>
        /// Pure and public so T1 pins the wording; reaching any of these states for real
        /// needs a profile with 10 000 folders, a cyclic folder tree, or a store whose
        /// DisplayName read fails - none of which a CI runner has.
        /// </para>
        /// </summary>
        public static IReadOnlyList<string>? DescribeFolderListingCoverage(
            bool walkCapReached, bool depthLimitReached, int storesUnnamed, int storesUnnamedExcluded)
        {
            List<string> advice = new List<string>();
            if (walkCapReached)
            {
                advice.Add("INCOMPLETE LISTING - the folder walk stopped at its "
                    + FolderWalkAbsoluteCap.ToString(CultureInfo.InvariantCulture)
                    + "-folder cap, so folders (and possibly whole stores, which are walked one after another) are "
                    + "missing from this tree. Paging cannot reach them - the next call re-walks and stops in the "
                    + "same place - so list one store at a time with 'store'.");
            }

            if (depthLimitReached)
            {
                advice.Add("INCOMPLETE LISTING - folders nested deeper than "
                    + OutlookComSession.FolderWalkDepthGuard.ToString(CultureInfo.InvariantCulture)
                    + " levels were refused, so a subtree is missing. That depth means a pathological or cyclic "
                    + "folder tree; the guard is what keeps the walk from taking the process down.");
            }

            if (storesUnnamed > 0)
            {
                advice.Add(storesUnnamed.ToString(CultureInfo.InvariantCulture)
                    + " store(s) would not report a display name and are listed under '" + StoreNaming.UnnamedStorePrefix
                    + "N)' labels (nameUnreadable: true). Their folders are real and are listed; the LABEL is not a "
                    + "name and cannot be passed back as 'store' to search or list_folders, because a store scope is "
                    + "matched against the display name that could not be read. Search without 'store' to include "
                    + "their mail.");
            }

            if (storesUnnamedExcluded > 0)
            {
                advice.Add(storesUnnamedExcluded.ToString(CultureInfo.InvariantCulture)
                    + " store(s) would not report a display name, so this store-scoped listing could neither include "
                    + "them nor rule them out. Call list_folders without 'store' to see them.");
            }

            return advice.Count == 0 ? null : advice;
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
                // Present only when this row is not ordinary mail. The index never opens the
                // item, so what it can say is the row's System.Kind, marked as such (gap B3
                // widened this tier to admit message rows of every class).
                ItemClass = MailItemAdmission.DescribeIndexRowClass(hit.Kinds, hit.IsAttachmentHit),
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
                // The real MAPI class, and only when it is not ordinary mail: this tier
                // opened the item, so it can name a bounce report or a meeting request
                // outright rather than guessing from a kind.
                ItemClass = MailItemAdmission.DescribeComItemClass(item.MessageClass),
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

        /// <summary>
        /// The index SCOPE for a whole store, or null when the store exists in the profile
        /// but the index cannot address it. Null is an ANSWER here, not a failure - the
        /// caller decides what an unaddressable store means for its own tool.
        /// </summary>
        private string? ResolveScope(string store, string? folder)
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
                    ComFolderPathList? tree = folder == null ? null : TryGetFolderPaths(store);
                    return FolderScopeResolver.ForDelegateStore(
                        delegateScope, folder, includeSubfolders, tree?.Paths, tree?.Incomplete ?? false);
                }
            }

            // NOTHING IN THE INDEX ADDRESSES THIS STORE. That is a fact about the INDEX, and
            // it used to be reported as if it were a fact about the profile: the refusal
            // named the index catalog as "Known stores", which on a profile of unindexed
            // data files is EMPTY, and sent the caller to list_accounts - which returns the
            // very name that just failed. A real store and a typo were the same message.
            //
            // So the verdict comes from Outlook, which is the profile the caller is actually
            // searching, through the same pure classifier list_folders refuses with (gap G1).
            // Fetched only here, on the failed lookup, and from a 5-minute cache, so an
            // ordinary store-scoped search pays nothing for it.
            IReadOnlyList<string>? profileStores = TryGetProfileStoreNames();
            if (profileStores == null)
            {
                // Outlook is unreachable, so the two cases genuinely cannot be told apart
                // right now. Say that, rather than picking one and sounding certain.
                throw new ArgumentException(
                    "Store '" + store + "' is not in the local search index, and Outlook could not be reached to "
                    + "check whether the profile has it - so a store that exists but is not indexed cannot be told "
                    + "apart from a misspelled name. Check outlook_health, then retry; search without 'store' works "
                    + "meanwhile.",
                    nameof(store));
            }

            string? refusal = DescribeUnresolvedFolderStore(store, profileStores);
            if (refusal != null)
            {
                throw new ArgumentException(refusal, nameof(store));
            }

            // Outlook HAS it; only the index does not. The search proceeds with the index
            // tier skipped - never widened, because 'store' filters which mail may come back
            // and a widened scope would answer with another store's mail.
            return FolderScopeResolver.ForUnindexedStore(folder);
        }

        /// <summary>
        /// The store display names OUTLOOK reports, or null when Outlook could not be
        /// reached. The same list <c>list_folders</c> refuses an unknown store against
        /// (gap G1), from the same <see cref="StoreDetailsCacheTtl"/> cache.
        /// <para>
        /// Null is a THIRD answer, not an empty list: "the profile has no stores" and "we
        /// could not ask" lead to different messages, and collapsing them would let a wedged
        /// Outlook produce a confident "that store does not exist".
        /// </para>
        /// </summary>
        private IReadOnlyList<string>? TryGetProfileStoreNames()
        {
            try
            {
                return _gateway.Run(s => GetStoreDetails(s)).Select(d => d.DisplayName).ToList();
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                return null;
            }
        }

        /// <summary>
        /// Store-relative folder paths of one store from COM, cached for
        /// <see cref="FolderPathCacheTtl"/>. Returns null when Outlook cannot be reached -
        /// the caller then widens rather than narrowing on a guess.
        /// <para>
        /// The walk's BOUNDS travel with the paths (gap G4). A delegate folder scope is an
        /// OR of folder names taken from this list, so a list the walk cut short narrows the
        /// search silently - and a bare list cannot say whether it ended because the mailbox
        /// did. The truncation flag is cached with the paths for the same reason they are:
        /// the next caller inside the TTL is answered from this entry and must be told the
        /// same thing.
        /// </para>
        /// </summary>
        private ComFolderPathList? TryGetFolderPaths(string store)
        {
            lock (_catalogLock)
            {
                // MonotonicClock: the stamp is only ever subtracted from a later reading of
                // the same clock to get the entry's age, never compared with anything outside
                // this process.
                if (_folderPaths.TryGetValue(store, out (ComFolderPathList Paths, DateTime FetchedUtc) cached)
                    && MonotonicClock.UtcNow - cached.FetchedUtc <= FolderPathCacheTtl)
                {
                    return cached.Paths;
                }
            }

            ComFolderPathList paths;
            try
            {
                paths = _gateway.Run(s => s.ListFolderPaths(store, FolderWalkAbsoluteCap));
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                return null;
            }

            if (paths.Paths.Count == 0)
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
