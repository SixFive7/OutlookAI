using System;
using System.Collections.Generic;
using System.Linq;

using OutlookAI.Core.Mapi;

namespace OutlookAI.Core.IndexSearch
{
    /// <summary>How an Outlook store was tied to its slice of the Windows Search index.</summary>
    public enum StoreIndexMatchKind
    {
        /// <summary>
        /// The store is NOT in the index: a PST file - the one store family whose hash input is
        /// measured - and no index store carries its hash. Final: a same-named index store is
        /// another store's, never this one's.
        /// </summary>
        None = 0,

        /// <summary>
        /// The index store's <c>($hash)</c> equals a hash computed from the store's own documented
        /// input (<see cref="OutlookAI.Core.Mapi.StoreHash"/>). Deterministic: the same function
        /// Outlook ran to build the URL, so a rename, a second profile or a same-named store
        /// cannot move it.
        /// </summary>
        StoreHash = 1,

        /// <summary>
        /// Left to the NAME RULE the product used before the store hash: no index store carries
        /// one of the store's hashes and its hash input is not measured (an Exchange store, an
        /// IMAP or Outlook.com .ost, any store that is not a PST file), or the hash was claimed
        /// twice. The caller resolves the store exactly as it did before, so no store is ever
        /// matched worse than it was.
        /// </summary>
        NameRule = 2,

        /// <summary>
        /// A delegate store: it is never a store of its own in the index - its cached items sit
        /// under the owner's store as <c>/1/&lt;name&gt;</c> - so the caller resolves it there,
        /// as before.
        /// </summary>
        Delegate = 3,
    }

    /// <summary>One Outlook store as the matcher sees it: a name to report and the hashes it may carry.</summary>
    public sealed class StoreIndexIdentity
    {
        /// <summary>Creates an identity.</summary>
        public StoreIndexIdentity(
            string displayName,
            int? exchangeStoreType,
            IReadOnlyList<StoreHashCandidate> candidates,
            bool nameUnreadable = false,
            string? filePath = null)
        {
            DisplayName = displayName ?? throw new ArgumentNullException(nameof(displayName));
            ExchangeStoreType = exchangeStoreType;
            Candidates = candidates ?? throw new ArgumentNullException(nameof(candidates));
            NameUnreadable = nameUnreadable;
            FilePath = filePath;
        }

        /// <summary><c>Store.DisplayName</c> - the name every tool takes and reports.</summary>
        public string DisplayName { get; }

        /// <summary>Raw OlExchangeStoreType: 0 primary, 1 delegate, 2 public folders, 3 not Exchange, 4 additional.</summary>
        public int? ExchangeStoreType { get; }

        /// <summary>The hashes Microsoft's algorithm can give this store (<see cref="StoreHash.Candidates"/>).</summary>
        public IReadOnlyList<StoreHashCandidate> Candidates { get; }

        /// <summary>True when <see cref="DisplayName"/> is a label because Outlook would not report the name.</summary>
        public bool NameUnreadable { get; }

        /// <summary><c>Store.FilePath</c>: the store's .pst or .ost, or null when it has none or it would not read.</summary>
        public string? FilePath { get; }

        /// <summary>
        /// Whether this store's hash is AUTHORITATIVE: a PST file - not Exchange, a <c>.pst</c>
        /// path - whose id read. The rule for those is MEASURED (Q92/Q99: Unicode and ANSI, renamed,
        /// copied, re-keyed, in two profiles, after a catalog reset), so when no index store
        /// carries the hash the store is not in the index, full stop; falling back to its name
        /// could only pick another store that happens to share it (two stores named "Outlook
        /// Data File" sat in the test guest's index). Every other store - an IMAP or Outlook.com
        /// .ost included, which is "not Exchange" too - has a documented but UNMEASURED input,
        /// and falls back to the name rule instead.
        /// </summary>
        public bool HashIsAuthoritative =>
            ExchangeStoreType == 3
            && FilePath != null
            && FilePath.EndsWith(".pst", StringComparison.OrdinalIgnoreCase)
            && Candidates.Any(c => c.Input == StoreHashInput.EntryId);
    }

    /// <summary>The outcome for one store.</summary>
    public sealed class StoreIndexMatch
    {
        internal StoreIndexMatch(
            StoreIndexIdentity store, StoreScopeInfo? root, StoreIndexMatchKind kind, StoreHashInput? input, string? note)
        {
            Store = store;
            Root = root;
            Kind = kind;
            Input = input;
            Note = note;
        }

        /// <summary>The store.</summary>
        public StoreIndexIdentity Store { get; }

        /// <summary>The index store it was tied to: set for <see cref="StoreIndexMatchKind.StoreHash"/> only.</summary>
        public StoreScopeInfo? Root { get; }

        /// <summary>How it was tied.</summary>
        public StoreIndexMatchKind Kind { get; }

        /// <summary>For a hash match, the documented input whose hash the index carries.</summary>
        public StoreHashInput? Input { get; }

        /// <summary>Why a hash that was present could not be used, when that is what happened.</summary>
        public string? Note { get; }
    }

    /// <summary>
    /// Every Outlook store tied to the index store that is its own, and the index stores no
    /// Outlook store's hash claimed. Built by <see cref="StoreIndexMatcher.Match"/>.
    /// </summary>
    public sealed class StoreIndexMap
    {
        internal StoreIndexMap(IReadOnlyList<StoreIndexMatch> stores, IReadOnlyList<StoreScopeInfo> unclaimed)
        {
            Stores = stores;
            Unclaimed = unclaimed;
        }

        /// <summary>One entry per store, in the order the stores were given.</summary>
        public IReadOnlyList<StoreIndexMatch> Stores { get; }

        /// <summary>
        /// Index stores no Outlook store's hash claimed: another Outlook profile's stores (one
        /// Windows user has ONE index across all their profiles), a store moved or removed since
        /// its rows were pushed - and the index store a <see cref="StoreIndexMatchKind.NameRule"/>
        /// store resolves to by name, which only the caller's name rule can tell.
        /// </summary>
        public IReadOnlyList<StoreScopeInfo> Unclaimed { get; }

        /// <summary>
        /// The match for the store with this display name, or null when no store has it - or when
        /// SEVERAL do: two stores sharing a name cannot be told apart by it, and choosing one would
        /// answer with the other's mail.
        /// </summary>
        public StoreIndexMatch? ForStore(string displayName)
        {
            StoreIndexMatch? found = null;
            foreach (StoreIndexMatch m in Stores)
            {
                if (!string.Equals(m.Store.DisplayName, displayName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (found != null)
                {
                    return null;
                }

                found = m;
            }

            return found;
        }

        /// <summary>
        /// The Outlook store name for an index store prefix, or null when no store's hash claimed
        /// it. What turns an index hit's URL back into the name a caller passes as <c>store</c>.
        /// </summary>
        public string? StoreNameForPrefix(string? storePrefix)
        {
            if (string.IsNullOrEmpty(storePrefix))
            {
                return null;
            }

            foreach (StoreIndexMatch m in Stores)
            {
                if (m.Root != null && string.Equals(m.Root.StorePrefix, storePrefix, StringComparison.OrdinalIgnoreCase))
                {
                    return m.Store.DisplayName;
                }
            }

            return null;
        }
    }

    /// <summary>
    /// Ties each Outlook store to the index store that is its own (Q92/Q99). Pure.
    /// <para>
    /// THE RULE. An index store matches an Outlook store when its <c>($hash)</c> equals one of
    /// the hashes Microsoft's algorithm gives that store (<see cref="StoreHash.Candidates"/>), and
    /// no other store and no other index store makes the same claim. A claim made twice is never
    /// resolved by picking: a 32-bit hash is unique among one user's stores in practice, and a
    /// collision is exactly the case where picking would answer with the wrong mail.
    /// </para>
    /// <para>
    /// WITHOUT A MATCH, the measured case and the unmeasured ones part ways. A PST file whose hash
    /// no index store carries is not in the index (<see cref="StoreIndexMatchKind.None"/>). Every
    /// other store - and any store whose claim was refused - is handed back to the name rule
    /// (<see cref="StoreIndexMatchKind.NameRule"/>), which the caller runs exactly as before, so
    /// a store whose hash input is only documented is never matched worse than it was. The
    /// matcher itself never matches by name.
    /// </para>
    /// </summary>
    public static class StoreIndexMatcher
    {
        /// <summary>Matches <paramref name="stores"/> against the index's <paramref name="roots"/>.</summary>
        public static StoreIndexMap Match(IReadOnlyList<StoreIndexIdentity> stores, IReadOnlyList<StoreScopeInfo> roots)
        {
            if (stores == null)
            {
                throw new ArgumentNullException(nameof(stores));
            }

            if (roots == null)
            {
                throw new ArgumentNullException(nameof(roots));
            }

            // Every (store, root, input) whose hashes agree. A delegate has no store of its own
            // in the index, so it claims nothing.
            var claims = new List<(int Store, int Root, StoreHashInput Input)>();
            for (int s = 0; s < stores.Count; s++)
            {
                if (stores[s].ExchangeStoreType == 1)
                {
                    continue;
                }

                for (int r = 0; r < roots.Count; r++)
                {
                    uint? rootHash = roots[r].StoreHash;
                    if (rootHash == null)
                    {
                        continue;
                    }

                    StoreHashCandidate? hit = stores[s].Candidates.FirstOrDefault(c => c.Hash == rootHash.Value);
                    if (hit != null)
                    {
                        claims.Add((s, r, hit.Input));
                    }
                }
            }

            var results = new StoreIndexMatch?[stores.Count];
            var notes = new string?[stores.Count];
            var claimedRoots = new bool[roots.Count];

            foreach (IGrouping<int, (int Store, int Root, StoreHashInput Input)> byStore in claims.GroupBy(c => c.Store))
            {
                int s = byStore.Key;
                var mine = byStore.ToList();
                foreach ((int _, int r, StoreHashInput _) in mine)
                {
                    claimedRoots[r] = true;
                }

                if (mine.Select(c => c.Root).Distinct().Count() > 1)
                {
                    notes[s] = "more than one index store carries this store's hash ("
                        + string.Join(", ", mine.Select(c => roots[c.Root].StoreSegment).Distinct(StringComparer.Ordinal)) + ")";
                    continue;
                }

                int root = mine[0].Root;
                bool shared = claims.Any(c => c.Root == root && c.Store != s);
                if (shared)
                {
                    notes[s] = "another Outlook store's hash is also " + roots[root].StoreSegment;
                    continue;
                }

                results[s] = new StoreIndexMatch(stores[s], roots[root], StoreIndexMatchKind.StoreHash, mine[0].Input, null);
            }

            for (int s = 0; s < stores.Count; s++)
            {
                if (results[s] != null)
                {
                    continue;
                }

                StoreIndexIdentity store = stores[s];
                StoreIndexMatchKind kind = store.ExchangeStoreType == 1
                    ? StoreIndexMatchKind.Delegate
                    : notes[s] == null && store.HashIsAuthoritative
                        ? StoreIndexMatchKind.None
                        : StoreIndexMatchKind.NameRule;
                results[s] = new StoreIndexMatch(store, null, kind, null, notes[s]);
            }

            var unclaimed = new List<StoreScopeInfo>();
            for (int r = 0; r < roots.Count; r++)
            {
                if (!claimedRoots[r])
                {
                    unclaimed.Add(roots[r]);
                }
            }

            return new StoreIndexMap(results.Select(m => m!).ToList(), unclaimed);
        }
    }
}
