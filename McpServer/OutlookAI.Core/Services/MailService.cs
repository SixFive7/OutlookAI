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
    /// search (fast) + COM gap-sweep merge (fresh, D19), lazy hit location with caching
    /// (Phase-1 guidance: locate cost avg ~2 s - never locate eagerly, always cache),
    /// EntryID-based reads, attachment saving, thread lookup, account/folder listing and
    /// staleness self-reporting. No MCP types, no console assumptions; per-process hit
    /// cache only (a server process lives for one agent session).
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
        private const double VeryStaleAdviceMinutes = 720; // 12 h - suggest mode=exhaustive
        private static readonly TimeSpan SweepSafetyMargin = TimeSpan.FromMinutes(10);
        private static readonly TimeSpan EmptyIndexSweepWindow = TimeSpan.FromDays(7);

        private readonly Lazy<IndexSearchService> _index;
        private readonly ComGateway _gateway;
        private readonly SendConfirmationTokens _sendTokens;
        private readonly ConcurrentDictionary<string, CachedHit> _hits =
            new ConcurrentDictionary<string, CachedHit>(StringComparer.Ordinal);
        private readonly object _catalogLock = new object();
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

        /// <summary>Runs one search (v3.MD section 8 L1). Fast = index only; fresh adds the COM gap sweep.</summary>
        public SearchOutcome Search(SearchRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            int top = Clamp(request.Top, 1, 100);
            int snippetChars = Clamp(request.SnippetChars, 0, 1000);
            IReadOnlyList<string> terms = SplitTerms(request.Query);
            if (request.Folder != null && request.Store == null)
            {
                throw new ArgumentException("folder requires store.", nameof(request));
            }

            if (request.Mode == SearchMode.Exhaustive)
            {
                return RunExhaustive(request, terms, top);
            }

            string? scope = null;
            if (request.Store != null)
            {
                scope = ResolveScope(request.Store, request.Folder);
            }

            IndexQuery query = new IndexQuery
            {
                Scope = scope,
                Terms = terms.Count > 0 ? terms : null,
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
                Top = top,
            };

            IndexSearchResult indexResult = _index.Value.Search(query);
            IndexStalenessReport staleness = _index.Value.GetStaleness();
            bool outlookRunning = ComGateway.IsOutlookRunning();

            List<HitSummary> summaries = new List<HitSummary>(indexResult.Hits.Count);
            foreach (IndexHit hit in indexResult.Hits)
            {
                summaries.Add(RegisterIndexHit(hit, snippetChars));
            }

            SweepInfo? sweep = null;
            List<string> advice = new List<string>();
            if (request.Mode == SearchMode.Fresh)
            {
                sweep = RunGapSweep(request, terms, staleness, indexResult.Hits, summaries, snippetChars);
                if (sweep.Error != null)
                {
                    advice.Add("Freshness sweep unavailable (" + sweep.Error + "); results are index-only and may miss the last "
                        + DescribeAge(staleness) + " of mail. " + (ComGateway.IsInstallerMutexHeld()
                            ? "An add-in update is in progress - retry shortly (D17)."
                            : "Retry later, check index_status, or use mode=exhaustive with store + folder/after bounds for an index-free COM search."));
                }
            }
            else
            {
                double ageMinutes = staleness.Age?.TotalMinutes ?? 0;
                if (!outlookRunning)
                {
                    advice.Add("Outlook is not running, so the index is frozen; mode=fresh sweeps recent mail via COM (it may start Outlook).");
                }
                else if (ageMinutes > 30)
                {
                    advice.Add("Newest indexed mail is " + ageMinutes.ToString("F0", CultureInfo.InvariantCulture)
                        + " minutes old; use mode=fresh to also catch not-yet-indexed mail.");
                }
            }

            double staleMinutes = staleness.Age?.TotalMinutes ?? 0;
            if (staleMinutes > VeryStaleAdviceMinutes)
            {
                advice.Add("The index is very stale (" + (staleMinutes / 60).ToString("F0", CultureInfo.InvariantCulture)
                    + " h behind). For correctness-critical queries use mode=exhaustive (bounded COM scan, store + folder/after required) - it bypasses the index entirely.");
            }

            summaries.Sort((a, b) => DateTime.Compare(b.ReceivedUtc ?? DateTime.MinValue, a.ReceivedUtc ?? DateTime.MinValue));
            if (summaries.Count > top)
            {
                summaries.RemoveRange(top, summaries.Count - top);
            }

            return new SearchOutcome
            {
                Mode = request.Mode == SearchMode.Fresh ? "fresh" : "fast",
                Hits = summaries,
                IndexElapsedMs = indexResult.ElapsedMilliseconds,
                Sweep = sweep,
                Staleness = new StalenessInfo
                {
                    NewestIndexedUtc = staleness.NewestIndexedReceivedUtc,
                    AgeMinutes = staleness.Age?.TotalMinutes,
                    OutlookRunning = outlookRunning,
                },
                Advice = advice.Count > 0 ? advice : null,
            };
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
            DateTime gapStart = (staleness.NewestIndexedReceivedUtc ?? DateTime.UtcNow - EmptyIndexSweepWindow) - SweepSafetyMargin;
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

            Stopwatch stopwatch = Stopwatch.StartNew();
            ComSweepResult sweepResult;
            try
            {
                bool needBodies = terms.Count > 0;
                sweepResult = _gateway.Run(s => s.SweepDefaultFoldersNewerThan(
                    gapStart, SweepPerFolderCap, needBodies, request.Store));
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
            info.FoldersSwept = sweepResult.FoldersSwept;
            info.FoldersSkipped = sweepResult.FoldersSkipped;
            info.ItemsSeen = sweepResult.Items.Count;

            List<ComMailBrief> filtered = new List<ComMailBrief>();
            foreach (ComMailBrief item in sweepResult.Items)
            {
                if (!FreshMerge.MatchesTerms(item, terms))
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
        /// mode=exhaustive: folder/date-bounded COM scan that bypasses the index
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
                    "mode=exhaustive requires 'store' (a display name from list_accounts) - it scans Outlook folders directly instead of the index.",
                    nameof(request));
            }

            if (request.Folder == null && !request.AfterUtc.HasValue)
            {
                throw new ArgumentException(
                    "mode=exhaustive requires a bound: pass 'folder' (scan one folder) and/or 'after' (date-bounded store scan). Unbounded store scans take minutes - use mode=fast/fresh for those.",
                    nameof(request));
            }

            if (request.To != null)
            {
                throw new ArgumentException(
                    "'to' filtering is not supported in mode=exhaustive (scanned items carry no recipient list). Use mode=fast/fresh or filter after read.",
                    nameof(request));
            }

            if (request.AttachmentHitsOnly)
            {
                throw new ArgumentException(
                    "Attachment-content matching requires the index; mode=exhaustive scans mail subject/body only.",
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
                timeBudgetMs: ExhaustiveTimeBudgetMs));
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
                    + " s time budget stopped the scan - results are partial. Narrow the folder/date bounds.");
            }

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
                Mode = "exhaustive",
                Hits = summaries,
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
        public ReadOutcome Read(string id, int maxBodyChars = 20000, bool includeHeaders = false, int maxHeaderChars = 8192)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                throw new ArgumentException("id is required.", nameof(id));
            }

            maxBodyChars = Clamp(maxBodyChars, 0, 500000);
            maxHeaderChars = Clamp(maxHeaderChars, 256, 65536);

            (string entryId, string? storeId, string? locatedVia, long locateMs, string? hitId) = ResolveToEntryId(id);
            ComItemDetail detail = _gateway.Run(s =>
            {
                ComItemDetail? d = s.TryReadItem(entryId, storeId, includeHeaders, maxBodyChars, out string? error);
                if (d == null && storeId == null)
                {
                    // Direct EntryID without a known store: retry across stores.
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        d = s.TryReadItem(entryId, store.StoreId, includeHeaders, maxBodyChars, out error);
                        if (d != null)
                        {
                            break;
                        }
                    }
                }

                return d ?? throw new InvalidOperationException("Item could not be opened (" + (error ?? "unknown") + ").");
            });

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

            List<RecipientView> recipients = detail.Recipients
                .Select(r => new RecipientView { Kind = r.Kind, Name = r.Name, Address = r.Address })
                .ToList();
            List<AttachmentView> attachments = detail.Attachments
                .Select(a => new AttachmentView { Index = a.Index, FileName = a.FileName, SizeBytes = a.SizeBytes })
                .ToList();

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
                Body = detail.Body,
                BodyTotalChars = detail.BodyTotalChars,
                BodyTruncated = detail.BodyTotalChars > detail.Body.Length,
                BodyOrigin = detail.BodyOrigin,
                SizeBytes = detail.SizeBytes,
                IsRead = detail.IsRead,
                ConversationId = detail.ConversationId,
                InternetMessageId = detail.InternetMessageId,
                Headers = headers,
                HeadersTruncated = headersTruncated,
                Attachments = attachments,
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
        public ThreadOutcome Thread(string? conversationId, string? id, string? store, int top = 50)
        {
            top = Clamp(top, 1, 200);
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
                    Top = top,
                });
                if (result.Hits.Count > 0)
                {
                    List<HitSummary> hits = result.Hits
                        .Select(h => RegisterIndexHit(h, snippetChars: 200))
                        .OrderBy(h => h.ReceivedUtc ?? DateTime.MinValue)
                        .ToList();
                    stopwatch.Stop();
                    return new ThreadOutcome
                    {
                        ConversationId = conversationId,
                        Source = "index",
                        Hits = hits,
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

            // COM fallback: walk the Outlook Conversation of the referenced item.
            (string entryId, string? storeId, string? _, long _, string? _) = ResolveToEntryId(id);
            IReadOnlyList<ComMailBrief> briefs = _gateway.Run(s =>
            {
                IReadOnlyList<ComMailBrief>? items = s.TryGetConversationItems(entryId, storeId, top, out string? error);
                return items ?? throw new InvalidOperationException("Conversation walk failed (" + (error ?? "unknown") + ").");
            });

            List<HitSummary> comHits = briefs.Select(b => RegisterLiveHit(b, snippetChars: 0, source: "com")).ToList();
            stopwatch.Stop();
            return new ThreadOutcome
            {
                ConversationId = conversationId,
                Source = "com",
                Hits = comHits,
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
        /// the current_folder/subfolders scopes apply to.
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

            return new ShowSearchResultsOutcome
            {
                Query = query,
                Scope = NormalizeScopeName(scope),
                ExplorerFolderPath = state.CurrentFolderPath,
                ExplorerCaption = state.Caption,
                Displayed = true,
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
        public DraftOutcome NewDraft(string account, string? to, string? cc, string? subject, string? body, bool display = true)
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

            ComDraftCreateResult created = _gateway.Run(s =>
            {
                ComDraftCreateResult? r = s.TryCreateNewDraft(account, toList, ccList, subject!, body!, display, out string? error);
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
        public DraftOutcome ReplyDraft(string id, string? body, bool replyAll = false, bool display = true)
        {
            (string? hitId, string sourceEntryId, ComDraftCreateResult created) = CreateDerived(
                id, replyAll ? ComDerivedDraftKind.ReplyAll : ComDerivedDraftKind.Reply, to: null, body, display);
            string op = replyAll ? "replyall_draft" : "reply_draft";
            AuditDraft(op, created, requestedAccount: null, sourceEntryId);
            return ToDraftOutcome(replyAll ? "replyall" : "reply", created, hitId, sourceEntryId);
        }

        /// <summary>
        /// Creates a forward draft for a hit id / EntryID via COM <c>Forward()</c> -
        /// quoted content and attachments preserved, agent text above the quote, saved to
        /// the source store's Drafts (D4). Never sends.
        /// </summary>
        public DraftOutcome ForwardDraft(string id, string? body, string? to, bool display = true)
        {
            IReadOnlyList<string> toList = Text.HtmlBodyComposer.SplitRecipients(to);
            if (toList.Count == 0)
            {
                throw new ArgumentException("to is required for forward_draft: one or more recipient addresses separated by ';' or ','.", nameof(to));
            }

            (string? hitId, string sourceEntryId, ComDraftCreateResult created) = CreateDerived(
                id, ComDerivedDraftKind.Forward, toList, body, display);
            AuditDraft("forward_draft", created, requestedAccount: null, sourceEntryId);
            return ToDraftOutcome("forward", created, hitId, sourceEntryId);
        }

        private (string? HitId, string SourceEntryId, ComDraftCreateResult Created) CreateDerived(
            string id,
            ComDerivedDraftKind kind,
            IReadOnlyList<string>? to,
            string? body,
            bool display)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                throw new ArgumentException("id is required (a hit id from search/thread or a full EntryID).", nameof(id));
            }

            if (string.IsNullOrWhiteSpace(body))
            {
                throw new ArgumentException("body is required (plain text; it is placed above the quoted mail).", nameof(body));
            }

            (string entryId, string? storeId, string? _, long _, string? hitId) = ResolveToEntryId(id);
            IReadOnlyList<string> toList = to ?? Array.Empty<string>();
            ComDraftCreateResult created = _gateway.Run(s =>
            {
                ComDraftCreateResult? r = s.TryCreateDerivedDraft(entryId, storeId, kind, toList, body!, display, out string? error);
                if (r == null && storeId == null)
                {
                    // Direct EntryID without a known store: retry across stores (same
                    // pattern as read/open_in_outlook).
                    foreach (ComStoreDetail store in GetStoreDetails(s))
                    {
                        r = s.TryCreateDerivedDraft(entryId, store.StoreId, kind, toList, body!, display, out error);
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
                Displayed = created.Displayed,
                ConversationId = created.Draft.ConversationId,
                Recipients = created.Draft.Recipients
                    .Select(r => new RecipientView { Kind = r.Kind, Name = r.Name, Address = r.Address })
                    .ToList(),
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
                Recipients = ToRecipientViews(sent.Recipients),
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
                Recipients = ToRecipientViews(state.Recipients),
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

        private static IReadOnlyList<RecipientView> ToRecipientViews(IReadOnlyList<ComRecipientInfo> recipients)
        {
            return recipients
                .Select(r => new RecipientView { Kind = r.Kind, Name = r.Name, Address = r.Address })
                .ToList();
        }

        // ------------------------------------------------------------------ index_status

        /// <summary>Staleness + availability self-report (R7/D19). Never starts Outlook.</summary>
        public IndexStatusOutcome IndexStatus()
        {
            bool outlookRunning = ComGateway.IsOutlookRunning();
            bool mutexHeld = ComGateway.IsInstallerMutexHeld();
            List<string> advice = new List<string>();

            string provider;
            DateTime? newest = null;
            double? ageMinutes = null;
            List<StoreStaleness>? perStore = null;
            try
            {
                IndexSearchService index = _index.Value;
                provider = index.Provider.ToString();
                IndexStalenessReport report = index.GetStaleness();
                newest = report.NewestIndexedReceivedUtc;
                ageMinutes = report.Age?.TotalMinutes;

                // The unordered discovery sample misses tiny idle stores (Phase-1 fact
                // 5); when Outlook is already running, its store list closes the gap via
                // targeted per-address discovery. Never STARTS Outlook here (D17:
                // index_status is an index-only tool).
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
                advice.Add("The SystemIndex is not reachable; search cannot run on this machine state.");
            }

            if (!outlookRunning)
            {
                advice.Add("Outlook is not running: the index stops advancing and fresh-mode sweeps will start Outlook (D17)"
                    + (mutexHeld ? " - but an add-in update is in progress, so COM tools will ask you to retry later." : "."));
            }
            else if (ageMinutes.HasValue && ageMinutes.Value > 30)
            {
                advice.Add("Newest indexed mail is " + ageMinutes.Value.ToString("F0", CultureInfo.InvariantCulture)
                    + " minutes old. Use search mode=fresh for anything newer.");
            }

            if (advice.Count == 0)
            {
                advice.Add("Index is current" + (ageMinutes.HasValue
                    ? " (newest mail " + ageMinutes.Value.ToString("F1", CultureInfo.InvariantCulture) + " min ago)"
                    : string.Empty) + "; fast mode is safe.");
            }

            return new IndexStatusOutcome
            {
                Provider = provider,
                OutlookRunning = outlookRunning,
                InstallerMutexHeld = mutexHeld,
                NewestIndexedUtc = newest,
                IndexAgeMinutes = ageMinutes,
                PerStore = perStore,
                Advice = advice,
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

        /// <summary>Folder trees (list_folders), depth- and count-capped.</summary>
        public FoldersOutcome ListFolders(string? store = null, int depth = 2, int maxFolders = 300)
        {
            depth = Clamp(depth, 1, 6);
            maxFolders = Clamp(maxFolders, 1, 1000);
            IReadOnlyList<ComFolderInfo> folders = _gateway.Run(s => s.ListFolders(store, depth, maxFolders));

            List<StoreFoldersView> byStore = folders
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

            return new FoldersOutcome
            {
                Stores = byStore,
                Truncated = folders.Count >= maxFolders,
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
                if (folder == null)
                {
                    return match.StorePrefix;
                }

                return match.StorePrefix + "/0/" + folder.Trim('/');
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
                    return folder == null ? delegateScope : delegateScope + "/" + folder.Trim('/');
                }
            }

            string known = string.Join(", ", catalog.Select(s => s.StoreDisplayName));
            throw new ArgumentException(
                "Store '" + store + "' was not found in the local index. Known stores: " + known
                + ". Use list_accounts for the full store list.");
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
