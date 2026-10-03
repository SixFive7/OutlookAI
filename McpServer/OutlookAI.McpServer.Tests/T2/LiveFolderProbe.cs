using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The read-only folder probes the folder-scope live tests share: a store's whole folder tree, a
/// row count for one index query, and a delegate mailbox's index root. Split out of
/// <see cref="LiveFolderScopeTests"/> when its two delegate tests moved to
/// <see cref="LiveDelegateFolderScopeTests"/> (Q74 B1), so the two classes cannot come to probe
/// the same thing two ways. Every member only reads.
/// </summary>
internal static class LiveFolderProbe
{
    /// <summary>Flattened folder tree of one store (list_folders, all pages).</summary>
    public static IReadOnlyList<FolderView> FolderTree(MailService service, string store)
    {
        ArgumentNullException.ThrowIfNull(service);
        List<FolderView> all = new();
        int offset = 0;
        while (true)
        {
            FoldersOutcome page = service.ListFolders(store, offset);
            foreach (StoreFoldersView view in page.Stores)
            {
                all.AddRange(view.Folders);
            }

            if (!page.Truncated || page.NextOffset is not int next || next <= offset)
            {
                return all;
            }

            offset = next;
        }
    }

    /// <summary>How many rows one index query returns.</summary>
    public static int DrainCount(IndexSearchService index, IndexQuery query)
    {
        ArgumentNullException.ThrowIfNull(index);
        return index.Search(query).Hits.Count;
    }

    /// <summary>
    /// The delegate mailbox's index root (<c>&lt;host&gt;/1/&lt;delegate&gt;</c>), discovered by
    /// probing every host store - delegate mailboxes hang off a HOST account, which is exactly the
    /// fact the naive display-name construction gets wrong.
    /// </summary>
    public static string ResolveDelegateRootScope(IndexSearchService index, string delegateStore)
    {
        ArgumentNullException.ThrowIfNull(index);
        foreach (StoreScopeInfo host in index.DiscoverStoreScopes(2000))
        {
            string candidate = host.StorePrefix + "/1/" + delegateStore;
            if (index.ScopeHasAnyItem(candidate))
            {
                return candidate;
            }
        }

        throw new InvalidOperationException($"no index root found for delegate store '{delegateStore}'");
    }
}
