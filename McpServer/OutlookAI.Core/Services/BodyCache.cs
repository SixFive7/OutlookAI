using System;
using System.Collections.Generic;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// Per-process cache of FULLY extracted plain-text mail bodies, keyed by EntryID
    /// (soak fix D37). Purpose: true body paging - read's body_offset windows are
    /// served from the one-time extraction instead of re-transferring the whole body
    /// over COM for every window. Policy: an offset-0 read always extracts fresh (and
    /// refreshes the cache), so plain re-reads keep today's always-fresh semantics;
    /// only offset&gt;0 continuation reads are served from the snapshot the paging
    /// session started with. Bounded (entries + total chars + TTL - T1-pinned) and
    /// thread-safe.
    /// </summary>
    public sealed class BodyCache
    {
        /// <summary>Maximum cached bodies (oldest evicted first).</summary>
        public const int MaxEntries = 8;

        /// <summary>Maximum total cached characters across all entries (the newest entry always fits).</summary>
        public const int MaxTotalChars = 8_000_000;

        /// <summary>
        /// Maximum age served to continuation reads. Generous - it only guards a paging
        /// session against arbitrarily stale snapshots; offset-0 reads always refresh.
        /// </summary>
        public static readonly TimeSpan TimeToLive = TimeSpan.FromMinutes(15);

        private readonly object _lock = new object();
        private readonly Dictionary<string, Entry> _entries = new Dictionary<string, Entry>(StringComparer.OrdinalIgnoreCase);
        private readonly Func<DateTime> _utcNow;

        /// <summary>Creates a cache on the system clock.</summary>
        public BodyCache()
            : this(null)
        {
        }

        /// <summary>
        /// Creates a cache with an injectable clock (T1 expiry tests). The default is
        /// <see cref="MonotonicClock"/>, not <see cref="DateTime.UtcNow"/>: the only thing
        /// this clock is asked for is the AGE of an entry, and on the wall clock a backwards
        /// jump would hold a 15-minute paging snapshot open for the size of the jump.
        /// </summary>
        public BodyCache(Func<DateTime>? utcNow)
        {
            _utcNow = utcNow ?? (() => MonotonicClock.UtcNow);
        }

        /// <summary>Number of cached bodies (diagnostics/tests).</summary>
        public int Count
        {
            get
            {
                lock (_lock)
                {
                    return _entries.Count;
                }
            }
        }

        /// <summary>
        /// Returns the cached full body for an EntryID, or false when absent/expired
        /// (an expired entry is evicted on the spot).
        /// </summary>
        public bool TryGet(string entryId, out string body, out string origin)
        {
            body = string.Empty;
            origin = "none";
            if (string.IsNullOrEmpty(entryId))
            {
                return false;
            }

            lock (_lock)
            {
                if (!_entries.TryGetValue(entryId, out Entry? entry))
                {
                    return false;
                }

                if (_utcNow() - entry.InsertedUtc > TimeToLive)
                {
                    _entries.Remove(entryId);
                    return false;
                }

                body = entry.Body;
                origin = entry.Origin;
                return true;
            }
        }

        /// <summary>
        /// Stores (or refreshes) the full body of an EntryID, then evicts oldest-first
        /// while over the entry/char bounds - never the entry just inserted.
        /// </summary>
        public void Put(string entryId, string body, string origin)
        {
            if (string.IsNullOrEmpty(entryId) || body == null)
            {
                return;
            }

            lock (_lock)
            {
                _entries[entryId] = new Entry(body, origin ?? "none", _utcNow());

                while (_entries.Count > 1 && IsOverBounds(entryId))
                {
                    string? oldestKey = null;
                    DateTime oldest = DateTime.MaxValue;
                    foreach (KeyValuePair<string, Entry> pair in _entries)
                    {
                        if (!string.Equals(pair.Key, entryId, StringComparison.OrdinalIgnoreCase)
                            && pair.Value.InsertedUtc < oldest)
                        {
                            oldest = pair.Value.InsertedUtc;
                            oldestKey = pair.Key;
                        }
                    }

                    if (oldestKey == null)
                    {
                        break;
                    }

                    _entries.Remove(oldestKey);
                }
            }
        }

        /// <summary>Drops one entry (e.g. after an operation known to change the body).</summary>
        public void Invalidate(string entryId)
        {
            if (string.IsNullOrEmpty(entryId))
            {
                return;
            }

            lock (_lock)
            {
                _entries.Remove(entryId);
            }
        }

        private bool IsOverBounds(string keepKey)
        {
            if (_entries.Count > MaxEntries)
            {
                return true;
            }

            long total = 0;
            foreach (KeyValuePair<string, Entry> pair in _entries)
            {
                total += pair.Value.Body.Length;
            }

            // The just-inserted body may alone exceed the char bound - it still stays
            // (the cache exists to page exactly such giants); everything else goes.
            return total > MaxTotalChars && _entries.Count > 1 && HasOtherEntry(keepKey);
        }

        private bool HasOtherEntry(string keepKey)
        {
            foreach (string key in _entries.Keys)
            {
                if (!string.Equals(key, keepKey, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private sealed class Entry
        {
            public Entry(string body, string origin, DateTime insertedUtc)
            {
                Body = body;
                Origin = origin;
                InsertedUtc = insertedUtc;
            }

            public string Body { get; }

            public string Origin { get; }

            public DateTime InsertedUtc { get; }
        }
    }
}
