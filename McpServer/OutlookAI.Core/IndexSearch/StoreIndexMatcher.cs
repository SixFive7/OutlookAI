using System;
using System.Collections.Generic;
using System.Linq;

using OutlookAI.Core.Mapi;

namespace OutlookAI.Core.IndexSearch
{
    /// <summary>How an Outlook store was tied to its slice of the Windows Search index.</summary>
    /// <remarks>
    /// <see cref="NameAndHash"/>, <see cref="None"/> and <see cref="Ambiguous"/> are THE RULE (Q113 (a),
    /// decided by the maintainer 2026-10-03): deterministic, never a guess. <see cref="StoreHash"/> and
    /// <see cref="NameRule"/> are reached by an Exchange store only - THE ONE OPEN EXCEPTION, kept until
    /// Q113 (b) measures which hash input cached Exchange uses; see
    /// <see cref="StoreIndexIdentity.InExchangeException"/>, the one switch that routes a store there.
    /// </remarks>
    public enum StoreIndexMatchKind
    {
        /// <summary>
        /// The store is tied to NO index store: it is matched by name and hash (Outlook reports it
        /// as not Exchange), and no index store carries its own name together with its hash - or
        /// its id or its own name would not read, so it has no name and hash to match. Never
        /// widened to a same-named index store: that is another store's. For a <c>.pst</c> file
        /// this is final - its rule is measured, so it is not in the index; for an IMAP or
        /// Outlook.com <c>.ost</c>, whose input is documented but not measured, it is what can be
        /// established, and <see cref="StoreIndexMatch.Note"/> says so.
        /// </summary>
        None = 0,

        /// <summary>
        /// EXCHANGE ONLY - part of the one open exception (Q113 (b)). The index store's
        /// <c>($hash)</c> equals a hash computed from one of the store's documented inputs
        /// (<see cref="OutlookAI.Core.Mapi.StoreHash"/>), and no other store or index store makes
        /// the same claim. By hash alone, because the name a cached Exchange store is filed under
        /// is not measured either.
        /// </summary>
        StoreHash = 1,

        /// <summary>
        /// EXCHANGE ONLY - the other half of the one open exception (Q113 (b)). Left to the NAME
        /// RULE the product used before the store hash: no index store carries one of the store's
        /// hashes, or its hash was claimed twice. The caller resolves the store exactly as it did
        /// before. Delete this, and <see cref="StoreHash"/>, once Q113 (b) is measured.
        /// </summary>
        NameRule = 2,

        /// <summary>
        /// A delegate store: it is never a store of its own in the index - its cached items sit
        /// under the owner's store as <c>/1/&lt;name&gt;</c> - so the caller resolves it there,
        /// as before.
        /// </summary>
        Delegate = 3,

        /// <summary>
        /// THE RULE (Q113 (a)). The index store's name is the store's OWN name (its root folder's,
        /// which the index files it under) and its <c>($hash)</c> is the store's documented hash,
        /// and no other Outlook store makes the same claim. The index URL carries both, so where
        /// two stores differ in either, this tells them apart with certainty: a rename's leftover
        /// root (the same hash under the old name) and another store of the same name (another
        /// hash) are both passed over.
        /// </summary>
        NameAndHash = 4,

        /// <summary>
        /// REFUSED (Q113 (a)). The index store carrying this store's name and hash is claimed by
        /// another Outlook store too - two stores share both name and hash - or more than one index
        /// store carries them, so the index itself cannot tell this store's mail from another's.
        /// Nothing is picked: a search scoped to the store refuses, an unscoped search leaves the
        /// contested index store's rows out (<see cref="StoreIndexMatch.Contested"/>), and
        /// outlook_health says so for the store.
        /// </summary>
        Ambiguous = 5,
    }

    /// <summary>One Outlook store as the matcher sees it: a name to report, its own name, and the hashes it may carry.</summary>
    public sealed class StoreIndexIdentity
    {
        /// <summary>Creates an identity.</summary>
        public StoreIndexIdentity(
            string displayName,
            int? exchangeStoreType,
            IReadOnlyList<StoreHashCandidate> candidates,
            bool nameUnreadable = false,
            string? filePath = null,
            string? ownName = null)
        {
            DisplayName = displayName ?? throw new ArgumentNullException(nameof(displayName));
            ExchangeStoreType = exchangeStoreType;
            Candidates = candidates ?? throw new ArgumentNullException(nameof(candidates));
            NameUnreadable = nameUnreadable;
            FilePath = filePath;
            OwnName = string.IsNullOrEmpty(ownName) ? null : ownName;
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
        /// The store's OWN name - its root folder's, which is the name the index files it under
        /// (measured, <c>Docs/live-tier-on-the-vm.md</c> section 8 item 24) and which need not be
        /// <see cref="DisplayName"/>: a store named in the profile only, and an ANSI PST, whose
        /// <c>Store.DisplayName</c> Outlook reports one character short. Read for the stores the
        /// name-and-hash rule applies to; null for the rest and when it would not read.
        /// </summary>
        public string? OwnName { get; }

        /// <summary>
        /// THE ONE OPEN EXCEPTION (Q113 (b)), in one place so it is easy to remove: true for every
        /// store Outlook does not report as olNotExchange (3) - an Exchange mailbox, public folders,
        /// an additional mailbox, and a store whose type would not read, which cannot be shown not to
        /// be Exchange - except a delegate (1), which is never a store of its own in the index. Such a
        /// store keeps the behaviour it had before Q113: tied by its hash alone when exactly one
        /// documented input fits (<see cref="StoreIndexMatchKind.StoreHash"/>), otherwise left to the
        /// old name rule (<see cref="StoreIndexMatchKind.NameRule"/>). Every other store is matched by
        /// name AND hash, or not at all. Once Q113 (b) has measured the input - and the name - a
        /// cached Exchange store is filed under, this returns false and the two kinds above go.
        /// </summary>
        public bool InExchangeException => ExchangeStoreType != 3 && ExchangeStoreType != 1;

        /// <summary>
        /// Whether this store is a PST FILE - not Exchange and a <c>.pst</c> path - the one family
        /// whose hash input is MEASURED (Q92/Q99: Unicode and ANSI, renamed, copied, re-keyed, in two
        /// profiles, after a catalog reset). It changes no decision: every non-Exchange store is
        /// matched by name and hash. It changes what an unmatched store's note can claim - "not in
        /// the index" for a PST, "not tied, input unmeasured" for an IMAP or Outlook.com <c>.ost</c>.
        /// </summary>
        public bool IsPstFile =>
            ExchangeStoreType == 3
            && FilePath != null
            && FilePath.EndsWith(".pst", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>The outcome for one store.</summary>
    public sealed class StoreIndexMatch
    {
        internal StoreIndexMatch(
            StoreIndexIdentity store,
            StoreScopeInfo? root,
            StoreIndexMatchKind kind,
            StoreHashInput? input,
            string? note,
            IReadOnlyList<StoreScopeInfo>? contested = null,
            IReadOnlyList<string>? sharedWith = null)
        {
            Store = store;
            Root = root;
            Kind = kind;
            Input = input;
            Note = note;
            Contested = contested ?? Array.Empty<StoreScopeInfo>();
            SharedWith = sharedWith ?? Array.Empty<string>();
        }

        /// <summary>The store.</summary>
        public StoreIndexIdentity Store { get; }

        /// <summary>
        /// The index store it was tied to: set for <see cref="StoreIndexMatchKind.NameAndHash"/> and
        /// <see cref="StoreIndexMatchKind.StoreHash"/> only.
        /// </summary>
        public StoreScopeInfo? Root { get; }

        /// <summary>How it was tied.</summary>
        public StoreIndexMatchKind Kind { get; }

        /// <summary>For a tied store, the documented input whose hash the index carries.</summary>
        public StoreHashInput? Input { get; }

        /// <summary>
        /// Why the store was not tied, when there is something to say: a claim refused (two stores on
        /// one index store, one store on two), or the near misses of an untied store - its hash under
        /// another name (a rename's leftover), its name under another hash (another store's) - and,
        /// for a store that is not a PST file, that its hash input is unmeasured.
        /// </summary>
        public string? Note { get; }

        /// <summary>
        /// For <see cref="StoreIndexMatchKind.Ambiguous"/>: the index store(s) carrying this store's
        /// name and hash that the index cannot attribute to it alone. Their rows are left out of an
        /// unscoped search. Empty for every other kind.
        /// </summary>
        public IReadOnlyList<StoreScopeInfo> Contested { get; }

        /// <summary>
        /// For <see cref="StoreIndexMatchKind.Ambiguous"/>: the display names of the other Outlook
        /// stores that claim the same index store. Empty for every other kind, and when the store is
        /// ambiguous only because two index stores carry its name and hash.
        /// </summary>
        public IReadOnlyList<string> SharedWith { get; }
    }

    /// <summary>
    /// Every Outlook store tied to the index store that is its own, and the index stores no
    /// Outlook store claimed. Built by <see cref="StoreIndexMatcher.Match"/>.
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
        /// Index stores no Outlook store claimed: another Outlook profile's stores (one Windows
        /// user has ONE index across all their profiles), a store moved or removed since its rows
        /// were pushed, a rename's leftover root - and the index store a
        /// <see cref="StoreIndexMatchKind.NameRule"/> store resolves to by name, which only the
        /// caller's name rule can tell.
        /// </summary>
        public IReadOnlyList<StoreScopeInfo> Unclaimed { get; }

        /// <summary>
        /// The stores the index cannot tell apart (<see cref="StoreIndexMatchKind.Ambiguous"/>), in
        /// store order: an unscoped search leaves their <see cref="StoreIndexMatch.Contested"/> rows
        /// out and reports these stores as unmatched.
        /// </summary>
        public IReadOnlyList<StoreIndexMatch> Unmatchable =>
            Stores.Where(m => m.Kind == StoreIndexMatchKind.Ambiguous).ToList();

        /// <summary>
        /// The match for the store with this display name, or null when no store has it - or when
        /// SEVERAL do: two stores sharing a name cannot be told apart by it, and choosing one would
        /// answer with the other's mail. <see cref="StoresNamed"/> returns them all.
        /// </summary>
        public StoreIndexMatch? ForStore(string displayName)
        {
            IReadOnlyList<StoreIndexMatch> named = StoresNamed(displayName);
            return named.Count == 1 ? named[0] : null;
        }

        /// <summary>Every store whose display name is <paramref name="displayName"/> (ordinal, case-insensitive), in store order.</summary>
        public IReadOnlyList<StoreIndexMatch> StoresNamed(string displayName)
        {
            if (displayName == null)
            {
                throw new ArgumentNullException(nameof(displayName));
            }

            return Stores
                .Where(m => string.Equals(m.Store.DisplayName, displayName, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        /// <summary>
        /// The Outlook store name for an index store prefix, or null when no store was tied to it.
        /// What turns an index hit's URL back into the name a caller passes as <c>store</c>.
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
    /// Ties each Outlook store to the index store that is its own (Q92/Q99, Q113). Pure.
    /// <para>
    /// THE RULE (Q113 (a), decided by the maintainer 2026-10-03: deterministic matching, never a
    /// guess). An index store - <c>mapi16://{SID}/&lt;own name&gt;($hash)</c> - is a store's when its
    /// NAME is the store's own name and its HASH is one Microsoft's algorithm gives the store
    /// (<see cref="StoreHash.Candidates"/>), and no other Outlook store makes the same claim. The URL
    /// carries both, so two stores that differ in either are told apart with certainty. Two that
    /// share both cannot be told apart by the index itself, and are refused
    /// (<see cref="StoreIndexMatchKind.Ambiguous"/>) rather than resolved by picking. A store
    /// matched by this rule never falls back to its name alone: without a match it is tied to
    /// nothing (<see cref="StoreIndexMatchKind.None"/>).
    /// </para>
    /// <para>
    /// THE ONE OPEN EXCEPTION (<see cref="StoreIndexIdentity.InExchangeException"/>): an Exchange
    /// store keeps the pre-Q113 behaviour until Q113 (b) is measured - its hash alone when exactly
    /// one documented input fits and nothing else claims that index store
    /// (<see cref="StoreIndexMatchKind.StoreHash"/>), otherwise handed back to the name rule the
    /// caller runs exactly as before (<see cref="StoreIndexMatchKind.NameRule"/>). The matcher itself
    /// never matches by name alone.
    /// </para>
    /// <para>
    /// NAMES are compared the way the index stores them: the URL's name, percent-decoded
    /// (<see cref="MapiUrlSegment"/>), against the store's own name, ordinal and case-insensitive -
    /// the server treats an index store prefix as case-insensitive everywhere, and a SCOPE matches
    /// one so - with white space at either end not counted, because the URL parser trims a
    /// trailing space and nothing has measured how the index spells one. Two stores whose names
    /// differ only there would need the same 32-bit hash as well to be confused, and are then
    /// refused, never picked.
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

            // Every (store, root, input) that is a CLAIM: for a store under the rule, the root's
            // name is its own name and the root's hash is its hash; for a store in the Exchange
            // exception, the hash alone, as before Q113. A delegate has no store of its own in the
            // index, so it claims nothing.
            var claims = new List<(int Store, int Root, StoreHashInput Input)>();
            for (int s = 0; s < stores.Count; s++)
            {
                if (stores[s].ExchangeStoreType == 1)
                {
                    continue;
                }

                for (int r = 0; r < roots.Count; r++)
                {
                    StoreHashCandidate? hit = HashHit(stores[s], roots[r]);
                    if (hit == null)
                    {
                        continue;
                    }

                    if (!stores[s].InExchangeException && !NamesAgree(stores[s].OwnName, roots[r]))
                    {
                        continue;
                    }

                    claims.Add((s, r, hit.Input));
                }
            }

            var results = new StoreIndexMatch?[stores.Count];
            var notes = new string?[stores.Count];
            var claimedRoots = new bool[roots.Count];
            foreach ((int _, int r, StoreHashInput _) in claims)
            {
                claimedRoots[r] = true;
            }

            foreach (IGrouping<int, (int Store, int Root, StoreHashInput Input)> byStore in claims.GroupBy(c => c.Store))
            {
                int s = byStore.Key;
                StoreIndexIdentity store = stores[s];
                List<int> mine = byStore.Select(c => c.Root).Distinct().ToList();
                List<int> rivals = claims
                    .Where(c => c.Store != s && mine.Contains(c.Root))
                    .Select(c => c.Store)
                    .Distinct()
                    .ToList();

                if (store.InExchangeException)
                {
                    // ---- THE ONE OPEN EXCEPTION (Q113 (b)): Exchange, exactly as before Q113 (D57) ----
                    if (mine.Count > 1)
                    {
                        notes[s] = "more than one index store carries this store's hash (" + Segments(roots, mine) + ")";
                        continue;
                    }

                    if (rivals.Count > 0)
                    {
                        notes[s] = "another Outlook store's hash is also " + roots[mine[0]].StoreSegment;
                        continue;
                    }

                    results[s] = new StoreIndexMatch(
                        store, roots[mine[0]], StoreIndexMatchKind.StoreHash, byStore.First().Input, null);
                    continue;
                }

                // ---- THE RULE (Q113 (a)): name and hash, or refused - never picked ----
                if (mine.Count > 1 || rivals.Count > 0)
                {
                    List<StoreScopeInfo> contested = mine.Select(r => roots[r]).ToList();
                    List<string> sharedWith = rivals.Select(o => stores[o].DisplayName).ToList();
                    results[s] = new StoreIndexMatch(
                        store,
                        null,
                        StoreIndexMatchKind.Ambiguous,
                        null,
                        DescribeAmbiguity(roots, mine, sharedWith),
                        contested,
                        sharedWith);
                    continue;
                }

                results[s] = new StoreIndexMatch(
                    store, roots[mine[0]], StoreIndexMatchKind.NameAndHash, byStore.First().Input, null);
            }

            for (int s = 0; s < stores.Count; s++)
            {
                if (results[s] != null)
                {
                    continue;
                }

                StoreIndexIdentity store = stores[s];
                if (store.ExchangeStoreType == 1)
                {
                    results[s] = new StoreIndexMatch(store, null, StoreIndexMatchKind.Delegate, null, null);
                }
                else if (store.InExchangeException)
                {
                    // THE ONE OPEN EXCEPTION (Q113 (b)): the name rule, as before.
                    results[s] = new StoreIndexMatch(store, null, StoreIndexMatchKind.NameRule, null, notes[s]);
                }
                else
                {
                    results[s] = new StoreIndexMatch(store, null, StoreIndexMatchKind.None, null, DescribeUntied(store, roots));
                }
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

        /// <summary>
        /// The name the index files <paramref name="root"/> under, exactly as its URL spells it but
        /// percent-decoded - unlike <see cref="StoreScopeInfo.StoreDisplayName"/>, which the URL
        /// parser trims - or null when the segment carries no <c>($hash)</c>.
        /// </summary>
        internal static string? IndexName(StoreScopeInfo root)
        {
            if (root == null)
            {
                throw new ArgumentNullException(nameof(root));
            }

            int open = root.StoreSegment.LastIndexOf("($", StringComparison.Ordinal);
            return open < 0 || root.StoreHash == null
                ? null
                : MapiUrlSegment.Decode(root.StoreSegment.Substring(0, open));
        }

        /// <summary>Whether <paramref name="ownName"/> is the name <paramref name="root"/> is filed under (see the class remarks).</summary>
        internal static bool NamesAgree(string? ownName, StoreScopeInfo root)
        {
            string? indexName = IndexName(root);
            return ownName != null
                && indexName != null
                && string.Equals(ownName.Trim(), indexName.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private static StoreHashCandidate? HashHit(StoreIndexIdentity store, StoreScopeInfo root)
        {
            uint? rootHash = root.StoreHash;
            return rootHash == null ? null : store.Candidates.FirstOrDefault(c => c.Hash == rootHash.Value);
        }

        private static string Segments(IReadOnlyList<StoreScopeInfo> roots, IEnumerable<int> indexes)
            => string.Join(", ", indexes.Select(r => roots[r].StoreSegment).Distinct(StringComparer.Ordinal));

        private static string DescribeAmbiguity(IReadOnlyList<StoreScopeInfo> roots, IReadOnlyList<int> mine, IReadOnlyList<string> sharedWith)
        {
            if (sharedWith.Count > 0)
            {
                return "the index cannot tell this store apart from "
                    + string.Join(", ", sharedWith.Select(n => "'" + n + "'"))
                    + ": the index store " + Segments(roots, mine)
                    + " carries the name and hash of both, so its mail cannot be attributed to either. "
                    + "Renaming one of them (its data file name in Outlook) files it apart";
            }

            return "more than one index store carries this store's own name and hash (" + Segments(roots, mine)
                + "), so the index cannot say which holds its mail";
        }

        /// <summary>
        /// Why a store under the rule is tied to nothing, when there is more to say than "not in the
        /// index": an input that would not read, the near misses that a name lookup or a hash lookup
        /// alone would have taken, and - for a store that is not a PST file - that its input is
        /// unmeasured. Null for a PST with no near miss: not in the index, and nothing to add.
        /// </summary>
        private static string? DescribeUntied(StoreIndexIdentity store, IReadOnlyList<StoreScopeInfo> roots)
        {
            var parts = new List<string>();
            if (store.Candidates.Count == 0)
            {
                parts.Add("its store ID would not read, so its hash cannot be computed");
            }
            else if (store.OwnName == null)
            {
                parts.Add("its own name (its root folder's) would not read, so it cannot be matched by name and hash");
            }
            else
            {
                List<int> hashElsewhere = Enumerable.Range(0, roots.Count)
                    .Where(r => HashHit(store, roots[r]) != null)
                    .ToList();
                List<int> nameElsewhere = Enumerable.Range(0, roots.Count)
                    .Where(r => HashHit(store, roots[r]) == null && NamesAgree(store.OwnName, roots[r]))
                    .ToList();
                if (hashElsewhere.Count > 0)
                {
                    parts.Add("the index carries this store's hash only under another name (" + Segments(roots, hashElsewhere)
                        + ") - what a rename leaves until the index files the store again - and that is not used");
                }

                if (nameElsewhere.Count > 0)
                {
                    parts.Add("the index store of its name (" + Segments(roots, nameElsewhere)
                        + ") carries another hash - another store's (another profile's, or this file at another path) - and is not used");
                }
            }

            if (!store.IsPstFile)
            {
                parts.Add("no index store carries its own name with its hash; for a store that is not a .pst file the hash "
                    + "input (its entry ID) is documented, not measured");
            }

            return parts.Count == 0 ? null : string.Join("; ", parts);
        }
    }
}
