using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;

using OutlookAI.Core.Mapi;

namespace OutlookAI.Core.IndexSearch
{
    /// <summary>Result of one index search: compact hits plus timing/provenance.</summary>
    public sealed class IndexSearchResult
    {
        internal IndexSearchResult(
            IReadOnlyList<IndexHit> hits,
            long elapsedMilliseconds,
            string sql,
            IndexProviderKind provider,
            int rowsScanned,
            int rowsDropped,
            bool candidatesExhausted)
        {
            Hits = hits;
            ElapsedMilliseconds = elapsedMilliseconds;
            Sql = sql;
            Provider = provider;
            RowsScanned = rowsScanned;
            RowsDropped = rowsDropped;
            CandidatesExhausted = candidatesExhausted;
        }

        /// <summary>
        /// The result of a query that was never SENT, for a caller that must skip the index
        /// tier rather than run it wrong - today only a store the profile has and the index
        /// cannot address (<see cref="OutlookAI.Core.Services.FolderScopeKind.StoreNotIndexed"/>).
        /// <para>
        /// Every counter is zero because nothing happened, which is the honest reading:
        /// <see cref="RowsScanned"/> 0 with <see cref="Sql"/> empty says no statement ran,
        /// where a fabricated non-empty statement would read as one that ran and matched
        /// nothing. <see cref="Provider"/> has no "none" member and nothing on this path
        /// reads it; the empty statement beside it is what tells a diagnostic reader which
        /// case this is.
        /// </para>
        /// </summary>
        internal static IndexSearchResult NotQueried()
        {
            return new IndexSearchResult(
                Array.Empty<IndexHit>(),
                elapsedMilliseconds: 0,
                sql: string.Empty,
                provider: IndexProviderKind.OleDb,
                rowsScanned: 0,
                rowsDropped: 0,
                candidatesExhausted: false);
        }

        /// <summary>
        /// Mapped hits, newest first, with rows the ordering cannot rank last
        /// (<see cref="IndexOrderGuard.RankableFirst"/>) rather than wherever the provider's
        /// NULL collation put them.
        /// </summary>
        public IReadOnlyList<IndexHit> Hits { get; }

        /// <summary>
        /// Wall-clock cost of the query including connection open and row drain, and of the
        /// displacement refetch when one ran (<see cref="IndexOrderGuard"/>).
        /// </summary>
        public long ElapsedMilliseconds { get; }

        /// <summary>
        /// The WS-SQL statement that ANSWERED the search (diagnostics/tests). A displacement
        /// refetch, when one runs, is this statement plus
        /// <see cref="WsSqlBuilder.BuildOrderKeyPresence"/>; it is a recovery query rather
        /// than the search, so it is not reported here.
        /// </summary>
        public string Sql { get; }

        /// <summary>Provider that served the query.</summary>
        public IndexProviderKind Provider { get; }

        /// <summary>
        /// Rows the statement returned before <see cref="IndexRowFilter"/> ran, plus the
        /// refetch's rows when one ran.
        /// </summary>
        public int RowsScanned { get; }

        /// <summary>
        /// Rows the statement offered and <see cref="IndexRowFilter"/> refused: rows outside
        /// the mapi namespace (only reachable without a SCOPE), and rows of the wrong SHAPE
        /// for the requested <see cref="KindFilter"/> - an attachment row under
        /// <see cref="KindFilter.MessagesOnly"/>, a message row under
        /// <see cref="KindFilter.AttachmentsOnly"/>.
        /// <para>
        /// It no longer counts message rows dropped for their item class, because no search
        /// drops any (gap B3). It is the index tier's half of the counter the exhaustive
        /// scan reports as <c>exhaustive.rowsDropped</c> - one counter shape across tiers -
        /// and since 2026-08-18 it reaches the payload as <c>index.rowsDropped</c> instead
        /// of dying here.
        /// </para>
        /// <para>
        /// Counted over every row both statements returned. It used to stop at the first
        /// <c>Top</c> admitted rows, so it under-reported the refusals by exactly the tail
        /// nobody looked at.
        /// </para>
        /// </summary>
        public int RowsDropped { get; }

        /// <summary>
        /// True when this tier could not establish that the list holds every match it should:
        /// either the over-fetched candidate list ran out before enough rows were admitted,
        /// or the displacement refetch (<see cref="IndexOrderGuard"/>) failed, which leaves
        /// the rows the ordering may have hidden unrecovered. Both mean "possibly short of
        /// matches" and both are reported rather than silent (no-silent-caps discipline, D42).
        /// </summary>
        public bool CandidatesExhausted { get; }
    }

    /// <summary>
    /// Staleness self-report (v3.MD R7/D19): the index only advances while classic Outlook
    /// runs, so the newest indexed DateReceived versus the clock is surfaced to callers -
    /// stale results are served, never silently.
    /// </summary>
    public sealed class IndexStalenessReport
    {
        internal IndexStalenessReport(DateTime? newestIndexedReceivedUtc, DateTime clockUtc)
        {
            NewestIndexedReceivedUtc = newestIndexedReceivedUtc;
            ClockUtc = clockUtc;
        }

        /// <summary>
        /// The report for a scope the index provably holds no mail for, WITHOUT probing -
        /// used where probing is impossible rather than merely expensive: a store the
        /// profile has and the index cannot address has no scope to probe, and the only
        /// probe available (the profile-wide one) would answer about OTHER stores. Same
        /// shape as a probe that ran and found nothing, because that is the same fact.
        /// </summary>
        internal static IndexStalenessReport NoRows()
        {
            return new IndexStalenessReport(null, DateTime.UtcNow);
        }

        /// <summary>Newest System.Message.DateReceived in the (scoped) index, UTC; null if no rows.</summary>
        public DateTime? NewestIndexedReceivedUtc { get; }

        /// <summary>UTC clock at probe time.</summary>
        public DateTime ClockUtc { get; }

        /// <summary>Age of the newest indexed mail relative to the probe clock.</summary>
        public TimeSpan? Age => NewestIndexedReceivedUtc.HasValue ? ClockUtc - NewestIndexedReceivedUtc.Value : (TimeSpan?)null;
    }

    /// <summary>A store subtree discovered in the index.</summary>
    public sealed class StoreScopeInfo
    {
        internal StoreScopeInfo(string storePrefix, string storeDisplayName, int sampleCount, bool hasDelegateSubtree)
        {
            StorePrefix = storePrefix;
            StoreDisplayName = storeDisplayName;
            SampleCount = sampleCount;
            HasDelegateSubtree = hasDelegateSubtree;
        }

        /// <summary>Whole-store SCOPE prefix (mapi16://{SID}/store($hash)).</summary>
        public string StorePrefix { get; }

        /// <summary>Store display name parsed from the prefix.</summary>
        public string StoreDisplayName { get; }

        /// <summary>How many sampled item URLs fell under this prefix.</summary>
        public int SampleCount { get; }

        /// <summary>True when SCOPE='&lt;prefix&gt;/1' returned rows (delegate-store subtree, v3.MD section 5).</summary>
        public bool HasDelegateSubtree { get; }
    }

    /// <summary>
    /// The Phase-1 index search core (v3.MD section 0.6): WS-SQL over the SystemIndex with
    /// compact hit mapping, store-scope discovery and the staleness probe. Stateless and
    /// parallel-safe; works with Outlook closed (results go stale, never wrong).
    /// </summary>
    public sealed class IndexSearchService
    {
        private readonly IIndexClient _client;

        /// <summary>Creates a service over an explicit client (tests inject fakes here).</summary>
        public IndexSearchService(IIndexClient client)
        {
            _client = client ?? throw new ArgumentNullException(nameof(client));
        }

        /// <summary>Provider that serves this service's queries.</summary>
        public IndexProviderKind Provider => _client.Provider;

        /// <summary>
        /// Creates a service on the auto-selected provider (OleDb primary, ADODB COM
        /// fallback); <paramref name="providerReport"/> records the choice.
        /// </summary>
        public static IndexSearchService CreateDefault(out string providerReport)
        {
            return new IndexSearchService(IndexClientFactory.CreateAuto(out providerReport));
        }

        /// <summary>
        /// Runs one search; see <see cref="WsSqlBuilder.Build"/> for the emitted shape.
        /// <paramref name="commandTimeoutSeconds"/> defaults to
        /// <see cref="OleDbIndexClient.DefaultCommandTimeoutSeconds"/>; callers with a
        /// tool-level budget above them pass their own, so one saturated-indexer query
        /// cannot consume a budget that was meant for the whole call.
        /// </summary>
        public IndexSearchResult Search(IndexQuery query, int? commandTimeoutSeconds = null)
        {
            if (query == null)
            {
                throw new ArgumentNullException(nameof(query));
            }

            // The Kind predicate no longer fences the statement (attachment rows carry the
            // attachment's kind - v3.MD block (q)), so admission happens in code and the
            // statement over-fetches candidates to keep "TOP n" meaning n ADMITTED rows.
            int sqlTop = IndexRowFilter.ComputeSqlTop(query.Top, query.Scope != null, WsSqlBuilder.MaxTop);
            string sql = WsSqlBuilder.Build(query, sqlTop);
            Stopwatch stopwatch = Stopwatch.StartNew();
            IReadOnlyList<IReadOnlyDictionary<string, object?>> rows = _client.ExecuteRows(sql, sqlTop, commandTimeoutSeconds);

            // EVERY returned row is mapped and judged, not just the first Top of them: the
            // trim below happens after the ordering guard has run, and a row it never looked
            // at cannot be ranked. This is also what makes RowsScanned mean what it says.
            List<IndexHit> mapped = MapRows(rows);
            List<IndexHit> admitted = Admit(mapped, query.Kinds);
            IReadOnlyList<IndexHit> candidates = admitted;
            int scanned = mapped.Count;
            int dropped = scanned - admitted.Count;
            bool statementFilledTop = rows.Count >= sqlTop;
            bool refetchFilledTop = false;
            bool refetchFailed = false;

            if (IndexOrderGuard.NeedsOrderKeyRefetch(
                rows.Count, sqlTop, IndexOrderGuard.AnyRowMissingOrderKey(mapped, query.OrderBy)))
            {
                // Rows the ORDER BY cannot rank took slots in a statement that was cut off,
                // so rows it CAN rank may never have left the provider (IndexOrderGuard says
                // why this is the exact condition). Ask again for the rankable ones only and
                // union the two answers, which can add rows and can never remove any.
                try
                {
                    IReadOnlyList<IReadOnlyDictionary<string, object?>> refetchRows =
                        _client.ExecuteRows(WsSqlBuilder.Build(query, sqlTop, true), sqlTop, commandTimeoutSeconds);
                    List<IndexHit> refetchMapped = MapRows(refetchRows);
                    List<IndexHit> refetchAdmitted = Admit(refetchMapped, query.Kinds);
                    scanned += refetchMapped.Count;
                    dropped += refetchMapped.Count - refetchAdmitted.Count;
                    candidates = IndexOrderGuard.Merge(admitted, refetchAdmitted);
                    refetchFilledTop = refetchRows.Count >= sqlTop;
                }
                catch (Exception ex) when (ex is not OutOfMemoryException)
                {
                    // The recovery query failed, so this answer is the un-recovered one -
                    // today's rows, possibly short of mail the ordering hid. Reported through
                    // CandidatesExhausted rather than thrown: the caller still gets the rows
                    // the search found, and a search that dies outright because a SECOND
                    // statement was rejected would be a worse answer than a flagged one.
                    refetchFailed = true;
                }
            }

            // Unrankable rows go last HERE, before the trim, because the trim is the second
            // place they could displace mail: taking the provider's first Top admitted rows
            // keeps them even when the dated rows are already in hand.
            IReadOnlyList<IndexHit> ordered = IndexOrderGuard.RankableFirst(candidates, query.OrderBy);
            stopwatch.Stop();

            List<IndexHit> hits = new List<IndexHit>(Math.Min(ordered.Count, query.Top));
            for (int i = 0; i < ordered.Count && hits.Count < query.Top; i++)
            {
                hits.Add(ordered[i]);
            }

            bool exhausted = refetchFailed
                || (hits.Count < query.Top && (statementFilledTop || refetchFilledTop));
            return new IndexSearchResult(
                hits, stopwatch.ElapsedMilliseconds, sql, _client.Provider, scanned, dropped, exhausted);
        }

        private static List<IndexHit> MapRows(IReadOnlyList<IReadOnlyDictionary<string, object?>> rows)
        {
            List<IndexHit> mapped = new List<IndexHit>(rows.Count);
            foreach (IReadOnlyDictionary<string, object?> row in rows)
            {
                mapped.Add(IndexRowMapper.Map(row));
            }

            return mapped;
        }

        private static List<IndexHit> Admit(IReadOnlyList<IndexHit> mapped, KindFilter kinds)
        {
            List<IndexHit> admitted = new List<IndexHit>(mapped.Count);
            for (int i = 0; i < mapped.Count; i++)
            {
                if (IndexRowFilter.Keep(mapped[i], kinds))
                {
                    admitted.Add(mapped[i]);
                }
            }

            return admitted;
        }

        /// <summary>
        /// Probes the newest indexed DateReceived (optionally scoped to one store/folder)
        /// against the current clock.
        /// </summary>
        public IndexStalenessReport GetStaleness(string? scope = null, int? commandTimeoutSeconds = null)
        {
            string sql = WsSqlBuilder.BuildNewestReceivedProbe(scope);
            IReadOnlyList<IReadOnlyDictionary<string, object?>> rows = _client.ExecuteRows(sql, 1, commandTimeoutSeconds);

            DateTime? newest = null;
            if (rows.Count > 0 && rows[0].TryGetValue("System.Message.DateReceived", out object? value)
                && value is DateTime dt)
            {
                newest = dt.Kind == DateTimeKind.Unspecified ? DateTime.SpecifyKind(dt, DateTimeKind.Utc) : dt.ToUniversalTime();
            }

            return new IndexStalenessReport(newest, DateTime.UtcNow);
        }

        /// <summary>True when at least one indexed item exists under <paramref name="scope"/>.</summary>
        public bool ScopeHasAnyItem(string scope, int? commandTimeoutSeconds = null)
        {
            string sql = WsSqlBuilder.BuildScopeExistenceProbe(scope);
            return _client.ExecuteRows(sql, 1, commandTimeoutSeconds).Count > 0;
        }

        /// <summary>
        /// True when the index holds at least one row for a folder scope (scope plus
        /// optional folder-path equalities), ignoring every search filter. The non-silent
        /// zero-row guard uses this to tell "this folder holds no match" apart from "this
        /// folder path matched nothing in the index" (v3.MD constraint C7).
        /// </summary>
        public bool FolderScopeHasAnyItem(string? scope, IReadOnlyList<string>? folderPaths, int? commandTimeoutSeconds = null)
        {
            string sql = WsSqlBuilder.BuildFolderScopeExistenceProbe(scope, folderPaths);
            return _client.ExecuteRows(sql, 1, commandTimeoutSeconds).Count > 0;
        }

        /// <summary>
        /// Targeted store-scope discovery for one account: unordered URL samples are
        /// dominated by the big stores (a 30k sample missed the tiny idle store on this
        /// machine), and SCOPE demands the exact store segment including the ($hash)
        /// suffix, so small stores are found by querying mail addressed to / sent by the
        /// account (per-column CONTAINS, index-backed, ~60 ms) and taking the store
        /// prefix of hits whose store display name equals the address.
        /// <para>
        /// The sender probe matches display NAME as well as address since B1, which only
        /// ever offers this loop MORE candidate rows. It cannot widen the answer: the store
        /// prefix is taken from a hit whose store display name EQUALS the address, and that
        /// test is unchanged.
        /// </para>
        /// </summary>
        public StoreScopeInfo? TryDiscoverStoreScopeByAddress(string smtpAddress, int? commandTimeoutSeconds = null)
        {
            if (string.IsNullOrWhiteSpace(smtpAddress))
            {
                throw new ArgumentException("Address must not be blank.", nameof(smtpAddress));
            }

            IndexQuery[] probes =
            {
                new IndexQuery { Kinds = KindFilter.MailKindOnly, RecipientContains = smtpAddress, Top = 50 },
                new IndexQuery { Kinds = KindFilter.MailKindOnly, SenderContains = smtpAddress, Top = 50 },
            };

            foreach (IndexQuery probe in probes)
            {
                foreach (IndexHit hit in Search(probe, commandTimeoutSeconds).Hits)
                {
                    if (hit.StorePrefix == null
                        || !string.Equals(hit.StoreDisplayName, smtpAddress, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    bool hasDelegates = false;
                    try
                    {
                        string delegateProbe = WsSqlBuilder.BuildScopeExistenceProbe(hit.StorePrefix + "/1");
                        hasDelegates = _client.ExecuteRows(delegateProbe, 1, commandTimeoutSeconds).Count > 0;
                    }
                    catch (Exception ex) when (ex is not OutOfMemoryException)
                    {
                        // Best-effort delegate probe.
                    }

                    return new StoreScopeInfo(hit.StorePrefix, hit.StoreDisplayName!, 1, hasDelegates);
                }
            }

            return null;
        }

        /// <summary>
        /// Discovers store scopes by sampling email item URLs and grouping their store
        /// prefixes (WS-SQL has no DISTINCT). For each discovered store, probes
        /// SCOPE='&lt;prefix&gt;/1' to detect delegate subtrees. A busy store can dominate
        /// small samples - use <see cref="TryDiscoverStoreScopeByAddress"/> for stores the
        /// sample misses (the 2000-row pull measured 552 ms in the section-5 probes).
        /// <para>
        /// <paramref name="commandTimeoutSeconds"/> reaches BOTH statements - the sample and
        /// the per-store delegate probe. It was declared and then passed to neither, so
        /// outlook_health's 8 s per-store index budget silently escaped to the 30 s default
        /// on every one of them; on a delegate-heavy profile with a saturated indexer the
        /// tool whose whole promise is answering fast could spend minutes here.
        /// </para>
        /// </summary>
        public IReadOnlyList<StoreScopeInfo> DiscoverStoreScopes(int sampleSize = 2000, int? commandTimeoutSeconds = null)
        {
            string sql = WsSqlBuilder.BuildStoreDiscoverySample(sampleSize);
            IReadOnlyList<IReadOnlyDictionary<string, object?>> rows = _client.ExecuteRows(sql, sampleSize, commandTimeoutSeconds);

            Dictionary<string, (string DisplayName, int Count)> byPrefix =
                new Dictionary<string, (string, int)>(StringComparer.OrdinalIgnoreCase);
            foreach (IReadOnlyDictionary<string, object?> row in rows)
            {
                if (!row.TryGetValue("System.ItemUrl", out object? value) || value is not string url)
                {
                    continue;
                }

                if (!MapiItemUrl.TryParse(url, out MapiItemUrl? parsed) || parsed == null)
                {
                    continue;
                }

                if (byPrefix.TryGetValue(parsed.StorePrefix, out (string DisplayName, int Count) existing))
                {
                    byPrefix[parsed.StorePrefix] = (existing.DisplayName, existing.Count + 1);
                }
                else
                {
                    byPrefix[parsed.StorePrefix] = (parsed.StoreDisplayName, 1);
                }
            }

            List<StoreScopeInfo> result = new List<StoreScopeInfo>(byPrefix.Count);
            foreach (KeyValuePair<string, (string DisplayName, int Count)> entry in byPrefix.OrderByDescending(e => e.Value.Count))
            {
                bool hasDelegates = false;
                try
                {
                    string probe = WsSqlBuilder.BuildScopeExistenceProbe(entry.Key + "/1");
                    hasDelegates = _client.ExecuteRows(probe, 1, commandTimeoutSeconds).Count > 0;
                }
                catch (Exception ex) when (ex is not OutOfMemoryException)
                {
                    // Delegate probe is best-effort; a failing scope probe means "none seen".
                }

                result.Add(new StoreScopeInfo(entry.Key, entry.Value.DisplayName, entry.Value.Count, hasDelegates));
            }

            return result;
        }
    }
}
