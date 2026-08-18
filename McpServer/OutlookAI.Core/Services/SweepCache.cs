using System;
using System.Collections.Generic;

using OutlookAI.Core.Com;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// Short-lived cache for freshness gap-sweep results (D34): since the sweep is
    /// ALWAYS on (the fast/fresh mode split was dropped), rapid-fire iterative searches
    /// would otherwise pay a COM sweep per call. Entries are keyed on the sweep window
    /// base (the index frontier minus the safety margin) plus the store scope plus the
    /// SWEPT FOLDER SET, so a cached sweep is reused only while the frontier has not
    /// advanced and only for a request whose coverage it actually satisfies; entries
    /// expire after <see cref="DefaultTimeToLive"/> (~10 s, T1-pinned). Swept items are
    /// pure data snapshots (no COM refs), safe to hold. Pure logic - no COM, no clock
    /// reads (callers pass UTC now), fully unit-testable.
    /// <para>
    /// The folder set is part of the key because the sweep follows the search scope
    /// (soak fix 13): a folder-scoped sweep covers ONE folder subtree, so serving it to
    /// a store-wide query would answer from a fraction of the coverage - a silent recall
    /// bug in the opposite direction from the one that fix repaired.
    /// </para>
    /// </summary>
    public sealed class SweepCache
    {
        /// <summary>
        /// How long a sweep result may serve subsequent searches (D34 trade-off: within
        /// this window a brand-new arrival can be invisible to repeat searches; a first
        /// search after idle always sweeps live). Pinned by a T1 test.
        /// </summary>
        public static readonly TimeSpan DefaultTimeToLive = TimeSpan.FromSeconds(10);

        private readonly object _lock = new object();
        private readonly TimeSpan _timeToLive;
        private readonly Dictionary<string, CachedSweep> _entries =
            new Dictionary<string, CachedSweep>(StringComparer.OrdinalIgnoreCase);

        /// <summary>Creates the cache with the production TTL.</summary>
        public SweepCache()
            : this(null)
        {
        }

        /// <summary>Creates the cache with an explicit TTL (tests inject short/zero TTLs).</summary>
        public SweepCache(TimeSpan? timeToLive)
        {
            _timeToLive = timeToLive ?? DefaultTimeToLive;
        }

        /// <summary>One cached sweep with its provenance.</summary>
        public sealed class CachedSweep
        {
            /// <summary>Creates a cached sweep record.</summary>
            public CachedSweep(
                ComSweepResult result,
                DateTime baseGapStartUtc,
                string? store,
                string? folder,
                bool includeSubfolders,
                DateTime fetchedAtUtc,
                long elapsedMs,
                IReadOnlyDictionary<string, DateTime>? perStoreBaseUtc = null)
            {
                Result = result ?? throw new ArgumentNullException(nameof(result));
                BaseGapStartUtc = baseGapStartUtc;
                Store = store;
                Folder = folder;
                IncludeSubfolders = includeSubfolders;
                FetchedAtUtc = fetchedAtUtc;
                ElapsedMs = elapsedMs;
                PerStoreBaseUtc = Freeze(perStoreBaseUtc);
            }

            /// <summary>The swept items + folder counters.</summary>
            public ComSweepResult Result { get; }

            /// <summary>
            /// The unclamped sweep window start this result covers for a store NOT named in
            /// <see cref="PerStoreBaseUtc"/>.
            /// </summary>
            public DateTime BaseGapStartUtc { get; }

            /// <summary>
            /// The per-store window starts this sweep actually ran with (each store's own
            /// index frontier, minus the safety margin). Empty for a sweep that used one
            /// window everywhere, which is every store-scoped sweep.
            /// </summary>
            public IReadOnlyDictionary<string, DateTime> PerStoreBaseUtc { get; }

            /// <summary>
            /// The window this entry covers for one store: its own if the sweep opened one,
            /// otherwise the fallback base. Public so T1 pins the reuse rule from the
            /// outside - this is the value the broad-entry check compares, and the whole
            /// question the per-store windows raise for this cache.
            /// </summary>
            public DateTime WindowFor(string store)
            {
                return PerStoreBaseUtc.TryGetValue(store, out DateTime since) ? since : BaseGapStartUtc;
            }

            /// <summary>Store the sweep was scoped to (null = all stores).</summary>
            public string? Store { get; }

            /// <summary>
            /// Folder subtree the sweep was scoped to (null = the default folder set of
            /// every store in scope). Part of the cache key - see the class remarks.
            /// </summary>
            public string? Folder { get; }

            /// <summary>
            /// Whether the sweep covered the folder's SUBTREE. Part of the cache key
            /// (v3.MD constraint C6): a shallow sweep answering a recursive query would
            /// report one folder of coverage as a whole subtree's worth.
            /// </summary>
            public bool IncludeSubfolders { get; }

            /// <summary>When the sweep ran (UTC).</summary>
            public DateTime FetchedAtUtc { get; }

            /// <summary>Wall-clock cost of the original sweep.</summary>
            public long ElapsedMs { get; }
        }

        /// <summary>
        /// Looks up a reusable sweep for a request. Reuse requires: entry younger than
        /// the TTL, the SAME window base (a frontier advance invalidates - the index
        /// ingested something, so re-sweeping keeps the cache honest), the SAME folder
        /// scope, and a compatible store scope: the exact store entry, or the all-stores
        /// entry (whose items the caller filters down by store display name - sound only
        /// because every store gets the identical default folder set).
        /// <para>
        /// "The same window" now means the same window PER STORE
        /// (<paramref name="perStoreBaseUtc"/>), and that is the whole of what the per-store
        /// window change costs this cache. An all-stores entry is a BROAD sweep, and the
        /// request it may serve is a NARROW one, so the question is whether the broad entry
        /// covers the narrow request as well as a fresh sweep would: it does exactly when
        /// its window FOR THAT STORE is the window the request would have opened. Comparing
        /// the scalar base alone would have compared the fallback windows of two different
        /// stores and reused a sweep taken over a span this store's frontier had already
        /// moved past.
        /// </para>
        /// </summary>
        public bool TryGet(
            DateTime baseGapStartUtc,
            string? store,
            string? folder,
            bool includeSubfolders,
            DateTime nowUtc,
            out CachedSweep? cached,
            IReadOnlyDictionary<string, DateTime>? perStoreBaseUtc = null)
        {
            lock (_lock)
            {
                Prune(nowUtc);
                if (TryGetUsable(
                    KeyFor(store, folder, includeSubfolders), baseGapStartUtc, perStoreBaseUtc, nowUtc, out cached))
                {
                    return true;
                }

                // A store-scoped request can be served from an all-stores sweep (the
                // caller filters items by store); never the other way around, and never
                // across folder scopes - a folder-scoped sweep covers one subtree only.
                // The default folder set is shallow by construction, so an all-stores
                // entry may serve only a shallow-equivalent request.
                //
                // A store-scoped request has ONE window - its own store's - so the entry is
                // usable when it swept that store from the same instant, whatever windows it
                // used for the others.
                if (store != null && folder == null
                    && _entries.TryGetValue(KeyFor(null, null, includeSubfolders), out CachedSweep? broad)
                    && broad != null
                    && IsFresh(broad, nowUtc)
                    && broad.WindowFor(store) == baseGapStartUtc)
                {
                    cached = broad;
                    return true;
                }

                cached = null;
                return false;
            }
        }

        /// <summary>Records a completed live sweep (overwrites the scope's previous entry).</summary>
        public void Store(
            DateTime baseGapStartUtc,
            string? store,
            string? folder,
            bool includeSubfolders,
            ComSweepResult result,
            long elapsedMs,
            DateTime nowUtc,
            IReadOnlyDictionary<string, DateTime>? perStoreBaseUtc = null)
        {
            if (result == null)
            {
                throw new ArgumentNullException(nameof(result));
            }

            if (_timeToLive <= TimeSpan.Zero)
            {
                return; // Cache disabled (test injection).
            }

            lock (_lock)
            {
                Prune(nowUtc);
                _entries[KeyFor(store, folder, includeSubfolders)] = new CachedSweep(
                    result, baseGapStartUtc, store, folder, includeSubfolders, nowUtc, elapsedMs, perStoreBaseUtc);
            }
        }

        /// <summary>Drops every entry (tests; also useful before latency-critical polls).</summary>
        public void Clear()
        {
            lock (_lock)
            {
                _entries.Clear();
            }
        }

        /// <summary>
        /// Cache key: store scope + folder scope + subtree flag. The unit separator cannot
        /// occur in a store display name or a folder path, so the parts can never blur
        /// into each other (a store literally named "x/y" must not collide with store "x",
        /// folder "y"). The flag belongs in the key because a SHALLOW sweep must never
        /// answer a RECURSIVE query - it would report one folder of coverage as a whole
        /// subtree's worth (v3.MD constraint C6).
        /// </summary>
        private static string KeyFor(string? store, string? folder, bool includeSubfolders)
        {
            return (store ?? string.Empty) + "\u001F" + (folder ?? string.Empty)
                + "\u001F" + (includeSubfolders ? "1" : "0");
        }

        private bool TryGetUsable(
            string key,
            DateTime baseGapStartUtc,
            IReadOnlyDictionary<string, DateTime>? perStoreBaseUtc,
            DateTime nowUtc,
            out CachedSweep? cached)
        {
            if (_entries.TryGetValue(key, out CachedSweep? entry)
                && entry != null
                && entry.BaseGapStartUtc == baseGapStartUtc
                && SameWindows(entry.PerStoreBaseUtc, perStoreBaseUtc)
                && IsFresh(entry, nowUtc))
            {
                cached = entry;
                return true;
            }

            cached = null;
            return false;
        }

        private bool IsFresh(CachedSweep entry, DateTime nowUtc)
        {
            return nowUtc - entry.FetchedAtUtc <= _timeToLive && nowUtc >= entry.FetchedAtUtc;
        }

        /// <summary>
        /// Whether two per-store window sets are the same sweep. Equality, not containment:
        /// a differing entry means one store's frontier moved, and the existing rule is that
        /// a frontier advance invalidates so the cache cannot outlive what the index has
        /// ingested. Containment would keep a wider-but-older sweep alive across exactly that
        /// event.
        /// </summary>
        private static bool SameWindows(
            IReadOnlyDictionary<string, DateTime> entry, IReadOnlyDictionary<string, DateTime>? request)
        {
            int requestCount = request?.Count ?? 0;
            if (entry.Count != requestCount)
            {
                return false;
            }

            if (requestCount == 0)
            {
                return true;
            }

            foreach (KeyValuePair<string, DateTime> pair in request!)
            {
                if (!entry.TryGetValue(pair.Key, out DateTime since) || since != pair.Value)
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// A private, case-insensitive copy: the caller's map may be mutated after the call,
        /// and store display names are matched case-insensitively wherever they are compared.
        /// </summary>
        private static IReadOnlyDictionary<string, DateTime> Freeze(IReadOnlyDictionary<string, DateTime>? source)
        {
            if (source == null || source.Count == 0)
            {
                return EmptyWindows;
            }

            Dictionary<string, DateTime> copy =
                new Dictionary<string, DateTime>(source.Count, StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, DateTime> pair in source)
            {
                copy[pair.Key] = pair.Value;
            }

            return copy;
        }

        private static readonly Dictionary<string, DateTime> EmptyWindows =
            new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);

        private void Prune(DateTime nowUtc)
        {
            List<string>? dead = null;
            foreach (KeyValuePair<string, CachedSweep> pair in _entries)
            {
                if (nowUtc - pair.Value.FetchedAtUtc > _timeToLive || nowUtc < pair.Value.FetchedAtUtc)
                {
                    (dead ??= new List<string>()).Add(pair.Key);
                }
            }

            if (dead != null)
            {
                foreach (string key in dead)
                {
                    _entries.Remove(key);
                }
            }
        }
    }
}
