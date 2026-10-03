using OutlookAI.Core.IndexSearch;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Finds an indexed store's index scope for the live tier, in three steps: the unordered 2000-row
/// mail sample; for an address-named store the sample missed, the mail addressed to it; and then
/// the index's own listing of store roots - one shallow <c>DIRECTORY</c> traversal of the user's
/// MAPI root, the listing the product's store map reads (Q92/Q99,
/// <see cref="IndexSearchService.ListStoreRoots"/>).
/// <para>
/// THE THIRD STEP IS WHY THIS EXISTS. The first live run with the 160,000-item <c>Corpus A</c>
/// (OutlookAI-Indexed, 2026-10-03) sampled 1,941 rows of another profile's 20,000-item store, 59 of
/// the hub's and 1 of the bystander's - and not one of Corpus A's. No address names Corpus A, so
/// the targeted step could not find it either, and 13 tests failed on "Store 'Corpus A' not found
/// among 3 discovered index scopes" before measuring anything. The sample decides nothing about
/// which stores exist; the root listing does.
/// </para>
/// <para>
/// Two rules keep the listing from answering wrongly. A name the listing holds more than once is
/// never guessed between - two profiles can each have a store of that name, and a scope of the
/// other one would measure the wrong store. And a root with nothing indexed below it does not
/// count: after a catalog reset or a store rename the root can be back while nothing below it is
/// (runbook section 8 item 24), and a test scoped to it would measure an empty store.
/// </para>
/// </summary>
internal static class LiveIndexScopes
{
    /// <summary>The sample size every live discovery uses.</summary>
    public const int SampleSize = 2000;

    /// <summary>
    /// The scope of <paramref name="storeDisplayName"/>, or null when none of the three steps finds
    /// one. <paramref name="sample"/> is a discovery sample already taken, so a caller resolving
    /// several stores samples once; null takes one.
    /// </summary>
    public static StoreScopeInfo? Find(
        IndexSearchService index,
        string storeDisplayName,
        IReadOnlyList<StoreScopeInfo>? sample = null,
        string? userRoot = null)
    {
        ArgumentNullException.ThrowIfNull(index);
        ArgumentException.ThrowIfNullOrWhiteSpace(storeDisplayName);

        sample ??= index.DiscoverStoreScopes(SampleSize);
        StoreScopeInfo? found = sample.FirstOrDefault(s =>
            string.Equals(s.StoreDisplayName, storeDisplayName, StringComparison.OrdinalIgnoreCase));
        if (found != null)
        {
            return found;
        }

        if (storeDisplayName.IndexOf('@', StringComparison.Ordinal) >= 0)
        {
            found = index.TryDiscoverStoreScopeByAddress(storeDisplayName);
            if (found != null)
            {
                return found;
            }
        }

        StoreScopeInfo? root = FindRoot(
            index.ListStoreRoots(userRoot ?? IndexSearchService.CurrentUserMapiRoot()), storeDisplayName);
        return root != null && index.ScopeHasAnyItem(root.StorePrefix) ? root : null;
    }

    /// <summary>
    /// The one root of that name in <paramref name="roots"/>, or null when there is none - or more
    /// than one, which is never guessed between.
    /// </summary>
    public static StoreScopeInfo? FindRoot(IReadOnlyList<StoreScopeInfo> roots, string storeDisplayName)
    {
        ArgumentNullException.ThrowIfNull(roots);
        List<StoreScopeInfo> named = roots
            .Where(r => string.Equals(r.StoreDisplayName, storeDisplayName, StringComparison.OrdinalIgnoreCase))
            .ToList();
        return named.Count == 1 ? named[0] : null;
    }
}
