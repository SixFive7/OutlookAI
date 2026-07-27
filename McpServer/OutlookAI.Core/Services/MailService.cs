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
        private const int SweepPerFolderCap = 200;
        private const int ExhaustiveTimeBudgetMs = 120_000;

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
        private const double VeryStaleAdviceMinutes = 720; // 12 h - suggest exhaustive:true
        private static readonly TimeSpan SweepSafetyMargin = TimeSpan.FromMinutes(10);
        private static readonly TimeSpan EmptyIndexSweepWindow = TimeSpan.FromDays(7);

        /// <summary>
        /// How long a store's COM folder-path list is reused. Only DELEGATE folder scopes
        /// need it (the flat index namespace has to be mapped back onto the real tree), so
        /// this keeps a per-search COM walk off the hot path. Matches the store-details TTL.
        /// </summary>
        private static readonly TimeSpan FolderPathCacheTtl = TimeSpan.FromMinutes(5);

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
        private readonly ComGateway _gateway;
        private readonly SendConfirmationTokens _sendTokens;
        private readonly SweepCache _sweepCache = new SweepCache();
        private readonly BodyCache _bodies = new BodyCache();
        private readonly ConcurrentDictionary<string, CachedHit> _hits =
            new ConcurrentDictionary<string, CachedHit>(StringComparer.Ordinal);
        private readonly object _catalogLock = new object();
        private readonly Dictionary<string, (IReadOnlyList<string> Paths, DateTime FetchedUtc)> _folderPaths =
            new Dictionary<string, (IReadOnlyList<string>, DateTime)>(StringComparer.OrdinalIgnoreCase);
        private string? _providerReport;
        private IReadOnlyList<StoreScopeInfo>? _catalog;
        private IReadOnlyList<ComStoreDetail>? _storeDetails;
        private DateTime _storeDetailsFetchedUtc;
        private int _nextHitId;

        /// <summary>Creates the service; both the index client and the COM session attach lazily.</summary>
        public MailService(ComGateway gateway)
            : this(gateway, null)
        {
        }

        /// <summary>
        /// Creates the service with an explicit send-confirmation token store (tests
        /// inject short-TTL stores; production uses the 120 s default).
        /// </summary>
        public MailService(ComGateway gateway, SendConfirmationTokens? sendTokens)
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

            IndexSearchResult indexResult = _index.Value.Search(query);
            IndexStalenessReport staleness = _index.Value.GetStaleness();

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
                sweep = RunGapSweep(request, terms, staleness, indexResult.Hits, summaries, snippetChars);
                if (sweep.Error == "RecipientFilterNotSweepable")
                {
                    advice.Add("Freshness sweep skipped: recipient ('to') filters cannot be matched by the sweep, so results are "
                        + "index-only and may lag the last " + DescribeAge(staleness) + " of mail.");
                }
                else if (sweep.Error != null)
                {
                    advice.Add("Freshness sweep unavailable (" + sweep.Error + "); results are index-only and may lag the last "
                        + DescribeAge(staleness) + " of mail. " + (ComGateway.IsInstallerMutexHeld()
                            ? "An add-in update is in progress - retry shortly (D17)."
                            : "Retry later, check outlook_health, or search again with exhaustive:true plus store + folder/after bounds for an index-free COM search."));
                }
                else if (sweep.Performed && sweep.FoldersSwept == 0 && request.Folder != null)
                {
                    // The folder-scoped sweep could not resolve the folder through COM
                    // (renamed, gone, or not a mail folder): index results still stand,
                    // but this query has no freshness coverage - say so rather than
                    // implying the last few minutes were checked.
                    advice.Add("Freshness sweep covered no folder: '" + request.Folder
                        + "' could not be opened in Outlook, so results for it are index-only and may lag the last "
                        + DescribeAge(staleness) + " of mail. Check the path with list_folders.");
                }

                AddSweepCoverageAdvice(advice, sweep, staleness);
            }

            AddUnresolvedFolderAdvice(advice, folderScope, request, summaries.Count);

            // Snapshot AFTER the sweep: the sweep may have just autostarted Outlook
            // (D17) and the staleness block must reflect that reality, not the
            // pre-autostart state (D34 self-consistency fix).
            bool outlookRunning = ComGateway.IsOutlookRunning();
            if (!outlookRunning && (sweep == null || (!sweep.Performed && sweep.Error == null)))
            {
                advice.Add("Outlook is not running, so the index is frozen; recent mail may be missing until Outlook runs again.");
            }

            double staleMinutes = staleness.Age?.TotalMinutes ?? 0;
            if (staleMinutes > VeryStaleAdviceMinutes)
            {
                advice.Add("The index is very stale (" + (staleMinutes / 60).ToString("F0", CultureInfo.InvariantCulture)
                    + " h behind). For correctness-critical queries search again with exhaustive:true (bounded COM scan, store + folder/after required) - it bypasses the index entirely.");
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

            return new SearchOutcome
            {
                Hits = summaries,
                Truncated = truncated,
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
                if (_index.Value.FolderScopeHasAnyItem(folderScope.Scope, folderScope.FolderPaths))
                {
                    return;
                }

                if (_index.Value.FolderScopeHasAnyItem(folderScope.StoreScope, null))
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
        /// </summary>
        private static void AddSweepCoverageAdvice(List<string> advice, SweepInfo sweep, IndexStalenessReport staleness)
        {
            if (!sweep.Performed)
            {
                return;
            }

            if (sweep.FoldersFailed > 0)
            {
                advice.Add("Freshness sweep FAILED on " + sweep.FoldersFailed.ToString(CultureInfo.InvariantCulture)
                    + " folder(s) - Outlook would not enumerate them, so mail that arrived there in the last "
                    + DescribeAge(staleness) + " is missing from these results. Retry, or use exhaustive:true for that folder.");
            }

            if (sweep.ItemCappedFolders != null && sweep.ItemCappedFolders.Count > 0)
            {
                advice.Add("Freshness sweep hit its per-folder cap of "
                    + SweepPerFolderCap.ToString(CultureInfo.InvariantCulture) + " items in: "
                    + string.Join(", ", sweep.ItemCappedFolders)
                    + ". It reads newest-first, so the OLDEST not-yet-indexed mail in those folders is not covered - "
                    + "narrow the window with 'after' or search those folders directly.");
            }

            if (sweep.FolderCapReached == true)
            {
                advice.Add("Freshness sweep stopped at its folder cap ("
                    + OutlookComSession.MaxScopedSweepFolders.ToString(CultureInfo.InvariantCulture)
                    + "), so deeper subfolders were never visited - index results still cover them, but brand-new "
                    + "mail there may be missing. Scope the search to a narrower folder for full freshness coverage.");
            }
            else if (sweep.FoldersSkipped > 0)
            {
                advice.Add("Freshness sweep skipped " + sweep.FoldersSkipped.ToString(CultureInfo.InvariantCulture)
                    + " folder(s) it could not resolve or enumerate, so mail that arrived there in the last "
                    + DescribeAge(staleness) + " may be missing. Check paths with list_folders.");
            }

            if (sweep.FolderListOmitted == true)
            {
                advice.Add("The swept-folder list is omitted above " + SweptFolderListCap.ToString(CultureInfo.InvariantCulture)
                    + " folders (payload discipline); sweep.foldersSwept is the true count.");
            }
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

        private SweepInfo RunGapSweep(
            SearchRequest request,
            IReadOnlyList<string> terms,
            IndexStalenessReport staleness,
            IReadOnlyList<IndexHit> indexHits,
            List<HitSummary> summaries,
            int snippetChars)
        {
            SweepInfo info = new SweepInfo();
            DateTime baseGapStart = (staleness.NewestIndexedReceivedUtc ?? DateTime.UtcNow - EmptyIndexSweepWindow) - SweepSafetyMargin;
            DateTime gapStart = baseGapStart;
            if (request.AfterUtc.HasValue && request.AfterUtc.Value > gapStart)
            {
                gapStart = request.AfterUtc.Value;
            }

            info.GapStartUtc = gapStart;
            if (request.BeforeUtc.HasValue && request.BeforeUtc.Value <= gapStart)
            {
                info.Performed = false;
                info.Error = null; // Window empty by request - nothing to sweep, not an error.
                return info;
            }

            if (request.To != null)
            {
                info.Performed = false;
                info.Error = "RecipientFilterNotSweepable";
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
            DateTime nowUtc = DateTime.UtcNow;
            IReadOnlyList<ComMailBrief> sweptItems;
            if (_sweepCache.TryGet(baseGapStart, request.Store, folderKey, sweepRecursive, nowUtc, out SweepCache.CachedSweep? cachedSweep)
                && cachedSweep != null)
            {
                info.Performed = true;
                info.Cached = true;
                info.CacheAgeSeconds = Math.Round((nowUtc - cachedSweep.FetchedAtUtc).TotalSeconds, 1);
                info.ElapsedMs = 0;
                ApplySweepCounters(info, cachedSweep.Result, request.Store);
                sweptItems = cachedSweep.Result.Items;
            }
            else
            {
                Stopwatch stopwatch = Stopwatch.StartNew();
                ComSweepResult sweepResult;
                try
                {
                    sweepResult = _gateway.Run(s => s.SweepFoldersNewerThan(
                        gapStart, SweepPerFolderCap, includeBodies: true, request.Store, sweepFolderPath, sweepRecursive));
                }
                catch (OutlookUnavailableException ex)
                {
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
                sweptItems = sweepResult.Items;

                // Only unclamped windows are cacheable: an After-narrowed sweep must
                // not poison wider follow-up searches.
                if (gapStart == baseGapStart)
                {
                    _sweepCache.Store(
                        baseGapStart, request.Store, folderKey, sweepRecursive, sweepResult, info.ElapsedMs, nowUtc);
                }
            }

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
                IndexStalenessReport staleness = _index.Value.GetStaleness();
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
        public ReadOutcome Read(string id, int maxBodyChars = BodyCharsDefault, bool includeHeaders = false, int maxHeaderChars = HeaderCharsDefault, int bodyOffset = 0)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                throw new ArgumentException("id is required.", nameof(id));
            }

            maxBodyChars = Clamp(maxBodyChars, 0, BodyCharsCap);
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
                ComItemDetail? d = s.TryReadItem(entryId, storeId, includeHeaders, includeBody: !haveCachedBody, out string? error);
                if (d == null && storeId == null)
                {
                    // Direct EntryID without a known store: retry across stores.
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        d = s.TryReadItem(entryId, store.StoreId, includeHeaders, includeBody: !haveCachedBody, out error);
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

                IndexSearchResult result = _index.Value.Search(new IndexQuery
                {
                    Scope = scope,
                    Kinds = KindFilter.EmailOnly,
                    ConversationIdEquals = conversationId,
                    Top = top + 1, // Over-fetch by one: definite has-more flag.
                });
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

            if (query.Length > 256)
            {
                throw new ArgumentException("query is too long for the Outlook search box (max 256 chars).", nameof(query));
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
        /// The swept-folder list for the sweep block: the folders the sweep actually
        /// covered, narrowed to the requested store when a cached all-stores sweep is
        /// serving a store-scoped request, and dropped entirely once the list is too
        /// long to be worth carrying (the count and scope still describe it).
        /// </summary>
        private static IReadOnlyList<string>? DescribeSweptFolders(ComSweepResult result, string? store)
        {
            if (result.SweptFolders.Count == 0)
            {
                return null;
            }

            List<string> folders = new List<string>(result.SweptFolders.Count);
            foreach (string entry in result.SweptFolders)
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

                folders.Add(entry);
            }

            return folders.Count == 0 || folders.Count > SweptFolderListCap ? null : folders;
        }

        /// <summary>
        /// Copies the sweep's coverage counters onto the response block, including the
        /// ones added by soak fix 15 (failed folders, per-folder item truncation, folder
        /// cap, and whether the folder LIST was dropped by its own cap). Every one of
        /// these was previously either absent or indistinguishable from success.
        /// </summary>
        private static void ApplySweepCounters(SweepInfo info, ComSweepResult result, string? store)
        {
            info.FoldersSwept = result.FoldersSwept;
            info.FoldersSkipped = result.FoldersSkipped;
            info.FoldersFailed = result.FoldersFailed;
            info.Folders = DescribeSweptFolders(result, store);
            info.FolderListOmitted = info.Folders == null && result.SweptFolders.Count > SweptFolderListCap
                ? true
                : (bool?)null;
            info.FolderCapReached = result.FolderCapReached ? true : (bool?)null;

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
        /// Creates a new draft in <paramref name="account"/>'s Drafts folder with that
        /// account's identity and signature (v3.MD section 3 mechanics), optionally
        /// displayed for the user (D4 default). Never sends. Audit-logged (load-bearing).
        /// </summary>
        public DraftOutcome NewDraft(string account, string? to, string? cc, string? subject, string? body, bool display = true, string? signature = null)
        {
            if (string.IsNullOrWhiteSpace(account))
            {
                throw new ArgumentException("account is required (a sending account SMTP address from list_accounts).", nameof(account));
            }

            IReadOnlyList<string> toList = Text.HtmlBodyComposer.SplitRecipients(to);
            IReadOnlyList<string> ccList = Text.HtmlBodyComposer.SplitRecipients(cc);
            if (toList.Count == 0)
            {
                throw new ArgumentException("to is required: one or more recipient addresses separated by ';' or ','.", nameof(to));
            }

            if (string.IsNullOrWhiteSpace(subject))
            {
                throw new ArgumentException("subject is required.", nameof(subject));
            }

            if (subject!.Length > 255)
            {
                throw new ArgumentException("subject is too long (max 255 characters).", nameof(subject));
            }

            if (string.IsNullOrWhiteSpace(body))
            {
                throw new ArgumentException("body is required (plain text; it is placed above the signature).", nameof(body));
            }

            ComSignatureOverride? signatureOverride = ResolveSignatureOverride(signature);
            ComDraftCreateResult created = _gateway.Run(s =>
            {
                ComDraftCreateResult? r = s.TryCreateNewDraft(account, toList, ccList, subject!, body!, display, signatureOverride, out string? error);
                return r ?? throw new InvalidOperationException(BuildDraftError(error, account));
            });

            AuditDraft("new_draft", created, requestedAccount: account, sourceEntryId: null);
            return ToDraftOutcome("new", created, hitId: null, sourceEntryId: null);
        }

        /// <summary>
        /// Creates a reply (or reply-all) draft for a hit id / EntryID via COM
        /// <c>Reply()</c>/<c>ReplyAll()</c> - threading and quoted history preserved,
        /// agent text above the quote, saved to the source store's Drafts (D4). Never sends.
        /// </summary>
        public DraftOutcome ReplyDraft(string id, string? body, bool replyAll = false, bool display = true, string? signature = null)
        {
            (string? hitId, string sourceEntryId, ComDraftCreateResult created) = CreateDerived(
                id, replyAll ? ComDerivedDraftKind.ReplyAll : ComDerivedDraftKind.Reply, to: null, body, display, signature);
            string op = replyAll ? "replyall_draft" : "reply_draft";
            AuditDraft(op, created, requestedAccount: null, sourceEntryId);
            return ToDraftOutcome(replyAll ? "replyall" : "reply", created, hitId, sourceEntryId);
        }

        /// <summary>
        /// Creates a forward draft for a hit id / EntryID via COM <c>Forward()</c> -
        /// quoted content and attachments preserved, agent text above the quote, saved to
        /// the source store's Drafts (D4). Never sends.
        /// </summary>
        public DraftOutcome ForwardDraft(string id, string? body, string? to, bool display = true, string? signature = null)
        {
            IReadOnlyList<string> toList = Text.HtmlBodyComposer.SplitRecipients(to);
            if (toList.Count == 0)
            {
                throw new ArgumentException("to is required for forward_draft: one or more recipient addresses separated by ';' or ','.", nameof(to));
            }

            (string? hitId, string sourceEntryId, ComDraftCreateResult created) = CreateDerived(
                id, ComDerivedDraftKind.Forward, toList, body, display, signature);
            AuditDraft("forward_draft", created, requestedAccount: null, sourceEntryId);
            return ToDraftOutcome("forward", created, hitId, sourceEntryId);
        }

        private (string? HitId, string SourceEntryId, ComDraftCreateResult Created) CreateDerived(
            string id,
            ComDerivedDraftKind kind,
            IReadOnlyList<string>? to,
            string? body,
            bool display,
            string? signature)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                throw new ArgumentException("id is required (a hit id from search/thread or a full EntryID).", nameof(id));
            }

            if (string.IsNullOrWhiteSpace(body))
            {
                throw new ArgumentException("body is required (plain text; it is placed above the quoted mail).", nameof(body));
            }

            ComSignatureOverride? signatureOverride = ResolveSignatureOverride(signature);
            (string entryId, string? storeId, string? _, long _, string? hitId) = ResolveToEntryId(id);
            IReadOnlyList<string> toList = to ?? Array.Empty<string>();
            ComDraftCreateResult created = _gateway.Run(s =>
            {
                ComDraftCreateResult? r = s.TryCreateDerivedDraft(entryId, storeId, kind, toList, body!, display, signatureOverride, out string? error);
                if (r == null && storeId == null)
                {
                    // Direct EntryID without a known store: retry across stores (same
                    // pattern as read/open_in_outlook).
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        r = s.TryCreateDerivedDraft(entryId, store.StoreId, kind, toList, body!, display, signatureOverride, out error);
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

            return (hitId, entryId, created);
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
        private static void AuditDraft(string operation, ComDraftCreateResult created, string? requestedAccount, string? sourceEntryId)
        {
            try
            {
                Audit.AuditLog.Append(
                    operation,
                    ("entryId", created.Draft.EntryId),
                    ("store", created.Draft.StoreDisplayName),
                    ("account", created.Draft.SendUsingAccountSmtp ?? requestedAccount),
                    ("accountResolved", created.AccountResolved ? "true" : "false"),
                    ("signatureInjected", created.SignatureInjected ? "true" : "false"),
                    ("signature", created.SignatureOverrideName),
                    ("signatureApplied", created.SignatureOverrideName != null ? (created.SignatureOverrideApplied ? "true" : "false") : null),
                    ("displayed", created.Displayed ? "true" : "false"),
                    ("recipients", created.Draft.Recipients.Count.ToString(CultureInfo.InvariantCulture)),
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

        private static DraftOutcome ToDraftOutcome(string kind, ComDraftCreateResult created, string? hitId, string? sourceEntryId)
        {
            IReadOnlyList<RecipientView> recipients = CapRecipients(created.Draft.Recipients, out int total, out bool truncated);
            return new DraftOutcome
            {
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
            };
        }

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

            string contentHash = SendContentHash.Compute(state.Subject, state.Recipients, state.BodyText, sentOnBehalfOf);

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
            foreach (string id in requestIds)
            {
                if (auditBroken)
                {
                    items.Add(FailedItem(id, "Not attempted: the audit log is unavailable (every move must be audited)."));
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
            foreach (string id in requestIds)
            {
                if (auditBroken)
                {
                    items.Add(FailedItem(id, "Not attempted: the audit log is unavailable (every move must be audited)."));
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

            // Outlook + stores (attach-only while running; never a cold start).
            int? storesReachable = null;
            List<string>? storeNames = null;
            if (outlookRunning)
            {
                try
                {
                    IReadOnlyList<ComStoreDetail> stores = _gateway.Run(GetStoreDetails);
                    storesReachable = stores.Count;
                    storeNames = stores.Select(s => s.DisplayName).ToList();
                    if (stores.Count == 0)
                    {
                        problems.Add("Outlook is running but reports no stores.");
                    }
                }
                catch (Exception ex) when (ex is not OutOfMemoryException)
                {
                    problems.Add("Outlook is running but the COM attach failed (" + ex.GetType().Name + ").");
                }
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
                IndexStalenessReport staleness = index.GetStaleness();
                newestIndexed = staleness.NewestIndexedReceivedUtc;
                ageMinutes = staleness.Age?.TotalMinutes;

                // The unordered discovery sample misses tiny idle stores (Phase-1 fact
                // 5); when Outlook is already running, its store list closes the gap via
                // targeted per-address discovery. Never STARTS Outlook here.
                if (outlookRunning)
                {
                    try
                    {
                        EnsureCatalogCoverageFromCom();
                    }
                    catch (Exception ex) when (ex is not OutOfMemoryException)
                    {
                        // Best-effort enrichment only.
                    }
                }

                perStore = new List<StoreStaleness>();
                foreach (StoreScopeInfo scopeInfo in GetCatalog())
                {
                    IndexStalenessReport scoped = index.GetStaleness(scopeInfo.StorePrefix);
                    perStore.Add(new StoreStaleness
                    {
                        Store = scopeInfo.StoreDisplayName,
                        NewestIndexedUtc = scoped.NewestIndexedReceivedUtc,
                    });
                }
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                provider = "unavailable: " + ex.GetType().Name;
                problems.Add("The SystemIndex is unreachable - index-backed search cannot run.");
            }

            if (!outlookRunning)
            {
                advice.Add("Outlook is not running: the index stops advancing and search's freshness sweep will start Outlook"
                    + (mutexHeld
                        ? " - but an add-in update is in progress, so the sweep degrades to index-only results until it finishes."
                        : "."));
            }
            else if (ageMinutes.HasValue && ageMinutes.Value > 30)
            {
                advice.Add("Newest indexed mail is " + ageMinutes.Value.ToString("F0", CultureInfo.InvariantCulture)
                    + " minutes old. search covers the gap automatically with its freshness sweep.");
            }

            if (advice.Count == 0 && !provider.StartsWith("unavailable", StringComparison.Ordinal))
            {
                advice.Add("Index is current" + (ageMinutes.HasValue
                    ? " (newest mail " + ageMinutes.Value.ToString("F1", CultureInfo.InvariantCulture) + " min ago)"
                    : string.Empty) + "; searches run at index speed.");
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
                    InstallerMutexHeld = mutexHeld,
                    // PROBED liveness (SF-1 fix): never report a dead held session as
                    // connected; the probe also releases a dead session's refs.
                    ComConnected = _gateway.ProbeConnected(),
                    StoresReachable = storesReachable,
                    Stores = storeNames,
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
                    throw new InvalidOperationException(
                        "Hit could not be located in Outlook (" + (location.Error ?? "unknown") + "). Re-run search - the item may have moved.");
                }

                string? storeId = null;
                if (location.StoreDisplayName != null)
                {
                    storeId = _gateway.Run(s => GetStoreDetails(s)
                        .FirstOrDefault(d => string.Equals(d.DisplayName, location.StoreDisplayName, StringComparison.OrdinalIgnoreCase))?.StoreId);
                }

                cached.LocatedEntryId = location.Located.EntryId;
                cached.LocatedStoreId = storeId;
                cached.LocatedVia = location.Tier == HitLocationTier.UrlSegments ? "urlSegments" : "itemPathDisplay";
                return (cached.LocatedEntryId, storeId, cached.LocatedVia, stopwatch.ElapsedMilliseconds, id);
            }

            // Raw EntryID hex (real Outlook EntryIDs are 70+ bytes = 140+ hex chars, but
            // accept anything plausibly hex and long enough to not be a hit id).
            if (id.Length >= 48 && id.Length % 2 == 0 && IsHex(id))
            {
                return (id, null, "directEntryId", 0, null);
            }

            throw new ArgumentException(
                "Unknown id '" + id + "'. Pass a hit id from a previous search/thread call in this session, or a full EntryID hex string.");
        }

        private void EnsureCatalogCoverageFromCom()
        {
            IReadOnlyList<ComStoreDetail> stores = _gateway.Run(GetStoreDetails);
            foreach (ComStoreDetail store in stores)
            {
                if (store.DisplayName.IndexOf('@') < 0)
                {
                    continue;
                }

                bool known = GetCatalog().Any(s =>
                    string.Equals(s.StoreDisplayName, store.DisplayName, StringComparison.OrdinalIgnoreCase));
                if (known)
                {
                    continue;
                }

                StoreScopeInfo? targeted = _index.Value.TryDiscoverStoreScopeByAddress(store.DisplayName);
                if (targeted != null)
                {
                    InvalidateCatalog(targeted);
                }
            }
        }

        private bool ProbeStoreInIndex(string displayName, bool isDelegate)
        {
            IReadOnlyList<StoreScopeInfo> catalog = GetCatalog();
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
                    if (_index.Value.ScopeHasAnyItem(owner.StorePrefix + "/1/" + displayName))
                    {
                        return true;
                    }
                }

                return false;
            }

            if (displayName.IndexOf('@') >= 0)
            {
                StoreScopeInfo? targeted = _index.Value.TryDiscoverStoreScopeByAddress(displayName);
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
            IReadOnlyList<StoreScopeInfo> catalog = GetCatalog();
            StoreScopeInfo? match = catalog.FirstOrDefault(s =>
                string.Equals(s.StoreDisplayName, store, StringComparison.OrdinalIgnoreCase));

            if (match == null && store.IndexOf('@') >= 0)
            {
                match = _index.Value.TryDiscoverStoreScopeByAddress(store);
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
                    exists = _index.Value.ScopeHasAnyItem(delegateScope);
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
                if (_folderPaths.TryGetValue(store, out (IReadOnlyList<string> Paths, DateTime FetchedUtc) cached)
                    && DateTime.UtcNow - cached.FetchedUtc <= FolderPathCacheTtl)
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
                _folderPaths[store] = (paths, DateTime.UtcNow);
            }

            return paths;
        }

        private IReadOnlyList<StoreScopeInfo> GetCatalog()
        {
            lock (_catalogLock)
            {
                _catalog ??= _index.Value.DiscoverStoreScopes(2000);
                return _catalog;
            }
        }

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

        private IReadOnlyList<ComStoreDetail> GetStoreDetails(OutlookComSession session)
        {
            lock (_catalogLock)
            {
                if (_storeDetails == null || DateTime.UtcNow - _storeDetailsFetchedUtc > TimeSpan.FromMinutes(5))
                {
                    _storeDetails = session.GetStoreDetails();
                    _storeDetailsFetchedUtc = DateTime.UtcNow;
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

        private static string DescribeAge(IndexStalenessReport staleness)
        {
            return staleness.Age.HasValue
                ? staleness.Age.Value.TotalMinutes.ToString("F0", CultureInfo.InvariantCulture) + " minutes"
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
