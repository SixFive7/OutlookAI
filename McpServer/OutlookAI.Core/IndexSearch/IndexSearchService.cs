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
        internal IndexSearchResult(IReadOnlyList<IndexHit> hits, long elapsedMilliseconds, string sql, IndexProviderKind provider)
        {
            Hits = hits;
            ElapsedMilliseconds = elapsedMilliseconds;
            Sql = sql;
            Provider = provider;
        }

        /// <summary>Mapped hits, newest first (ORDER BY System.Message.DateReceived DESC).</summary>
        public IReadOnlyList<IndexHit> Hits { get; }

        /// <summary>Wall-clock cost of the query including connection open and row drain.</summary>
        public long ElapsedMilliseconds { get; }

        /// <summary>The executed WS-SQL statement (diagnostics/tests).</summary>
        public string Sql { get; }

        /// <summary>Provider that served the query.</summary>
        public IndexProviderKind Provider { get; }
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

        /// <summary>Runs one search; see <see cref="WsSqlBuilder.Build"/> for the emitted shape.</summary>
        public IndexSearchResult Search(IndexQuery query)
        {
            if (query == null)
            {
                throw new ArgumentNullException(nameof(query));
            }

            string sql = WsSqlBuilder.Build(query);
            Stopwatch stopwatch = Stopwatch.StartNew();
            IReadOnlyList<IReadOnlyDictionary<string, object?>> rows = _client.ExecuteRows(sql, query.Top);
            stopwatch.Stop();

            List<IndexHit> hits = new List<IndexHit>(rows.Count);
            foreach (IReadOnlyDictionary<string, object?> row in rows)
            {
                hits.Add(IndexRowMapper.Map(row));
            }

            return new IndexSearchResult(hits, stopwatch.ElapsedMilliseconds, sql, _client.Provider);
        }

        /// <summary>
        /// Probes the newest indexed DateReceived (optionally scoped to one store/folder)
        /// against the current clock.
        /// </summary>
        public IndexStalenessReport GetStaleness(string? scope = null)
        {
            string sql = WsSqlBuilder.BuildNewestReceivedProbe(scope);
            IReadOnlyList<IReadOnlyDictionary<string, object?>> rows = _client.ExecuteRows(sql, 1);

            DateTime? newest = null;
            if (rows.Count > 0 && rows[0].TryGetValue("System.Message.DateReceived", out object? value)
                && value is DateTime dt)
            {
                newest = dt.Kind == DateTimeKind.Unspecified ? DateTime.SpecifyKind(dt, DateTimeKind.Utc) : dt.ToUniversalTime();
            }

            return new IndexStalenessReport(newest, DateTime.UtcNow);
        }

        /// <summary>True when at least one indexed item exists under <paramref name="scope"/>.</summary>
        public bool ScopeHasAnyItem(string scope)
        {
            string sql = WsSqlBuilder.BuildScopeExistenceProbe(scope);
            return _client.ExecuteRows(sql, 1).Count > 0;
        }

        /// <summary>
        /// Targeted store-scope discovery for one account: unordered URL samples are
        /// dominated by the big stores (a 30k sample missed the tiny idle store on this
        /// machine), and SCOPE demands the exact store segment including the ($hash)
        /// suffix, so small stores are found by querying mail addressed to / sent by the
        /// account (per-column CONTAINS, index-backed, ~60 ms) and taking the store
        /// prefix of hits whose store display name equals the address.
        /// </summary>
        public StoreScopeInfo? TryDiscoverStoreScopeByAddress(string smtpAddress)
        {
            if (string.IsNullOrWhiteSpace(smtpAddress))
            {
                throw new ArgumentException("Address must not be blank.", nameof(smtpAddress));
            }

            IndexQuery[] probes =
            {
                new IndexQuery { Kinds = KindFilter.EmailOnly, RecipientContains = smtpAddress, Top = 50 },
                new IndexQuery { Kinds = KindFilter.EmailOnly, FromAddressContains = smtpAddress, Top = 50 },
            };

            foreach (IndexQuery probe in probes)
            {
                foreach (IndexHit hit in Search(probe).Hits)
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
                        hasDelegates = _client.ExecuteRows(delegateProbe, 1).Count > 0;
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
        /// </summary>
        public IReadOnlyList<StoreScopeInfo> DiscoverStoreScopes(int sampleSize = 2000)
        {
            string sql = WsSqlBuilder.BuildStoreDiscoverySample(sampleSize);
            IReadOnlyList<IReadOnlyDictionary<string, object?>> rows = _client.ExecuteRows(sql, sampleSize);

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
                    hasDelegates = _client.ExecuteRows(probe, 1).Count > 0;
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
