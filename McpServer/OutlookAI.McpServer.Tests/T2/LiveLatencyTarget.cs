using System.Globalization;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>One indexed store and how many items it holds, as its own folders report them.</summary>
/// <param name="Store">The store's display name, exactly as the indexed list names it.</param>
/// <param name="Items">Total items across the store's folders (PR_CONTENT_COUNT), or null when they could not be counted.</param>
/// <param name="Complete">False when the folder walk stopped short of the tree, so <paramref name="Items"/> is only a lower bound.</param>
public sealed record LiveStoreSize(string Store, long? Items, bool Complete = true);

/// <summary>
/// Which indexed store a LATENCY bound is timed against: the LARGEST one. Decided by the maintainer
/// 2026-09-24 (Q70 follow-up, option (c)).
/// <para>
/// <b>Why.</b> The index tier holds every query to <c>MaxQueryMs = 2000</c>. Those bounds were set
/// against the maintainer's profile - about 160,000 items over five stores - and several tests timed
/// them against the FIRST indexed store, because that is the one whose content they read. On a test
/// guest the first indexed store is the hub, a few dozen items, and a two-second bound on it is met by
/// construction: the test passes and has measured nothing. The content half of those tests stays on
/// the first store - that is where the attachments and senders are - and the LATENCY half is timed on
/// the largest, where a slow shape would actually show.
/// </para>
/// <para>
/// <b>Largest by what.</b> By the items the store's own folders report, summed - read once per run
/// over COM. Not by the index's 2000-row discovery sample, which is unordered and can be dominated by
/// whichever store the indexer crawled first; and not by draining every store's rows, which on a
/// 160,000-item store is itself the slow query this exists to time. A store whose size could not be
/// read, or whose folder walk stopped short, is REFUSED rather than guessed at: a latency bound timed
/// on a store that only looked largest is the silent weakening this exists to prevent.
/// </para>
/// <para>
/// Pure: sizes in, a store name out. T1 pins it (<c>T1/LatencyTargetTests</c>), and pins from the
/// compiled IL that every latency-timing test reads its target from here.
/// </para>
/// </summary>
public static class LiveLatencyTarget
{
    private static readonly object Gate = new();
    private static (string Key, IReadOnlyList<LiveStoreSize> Sizes, string Largest)? _measured;

    /// <summary>
    /// Every indexed store's size and the largest of them, measured ONCE per test process for the
    /// indexed list <paramref name="settings"/> names: one COM session of its own, one folder walk per
    /// indexed store, READ-ONLY. Every latency test calls this rather than a fixture, because the
    /// tests that time the index live in three different collections.
    /// </summary>
    public static (IReadOnlyList<LiveStoreSize> Sizes, string Largest) Measure(LiveTestSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        IReadOnlyList<string> indexed = settings.RequireIndexedStores();
        string key = string.Join("\n", indexed);
        lock (Gate)
        {
            if (_measured is { } known && string.Equals(known.Key, key, StringComparison.Ordinal))
            {
                return (known.Sizes, known.Largest);
            }

            using OutlookComSession session = OutlookComSession.Connect(allowStartingOutlook: true);
            List<LiveStoreSize> sizes = indexed
                .Select(store => SizeOf(store, session.ListFolders(store, MailService.FolderWalkAbsoluteCap)))
                .ToList();
            string largest = Largest(sizes);
            _measured = (key, sizes, largest);
            return (sizes, largest);
        }
    }

    /// <summary>
    /// A store's size from its folder walk: the sum of every folder's item count. A walk that stopped
    /// at the cap or the depth guard is marked incomplete, and <see cref="Largest"/> refuses it; a
    /// store the walk returned no folder of has no size at all. Pure.
    /// </summary>
    public static LiveStoreSize SizeOf(string store, ComFolderTree tree)
    {
        ArgumentNullException.ThrowIfNull(tree);
        List<ComFolderInfo> folders = tree.Folders
            .Where(f => string.Equals(f.StoreDisplayName, store, StringComparison.OrdinalIgnoreCase))
            .ToList();
        long? items = folders.Count == 0 ? null : folders.Sum(f => f.ItemCount ?? 0);
        return new LiveStoreSize(store, items, !tree.WalkCapReached && !tree.DepthLimitReached);
    }

    /// <summary>
    /// The store to time a latency bound on: the one holding the most items. Ties go to the store
    /// the indexed list names FIRST, so the choice is deterministic and a reordering of the settings
    /// cannot make it flap between two equal stores from run to run.
    /// </summary>
    /// <param name="sizes">Every indexed store, in the indexed list's order.</param>
    public static string Largest(IReadOnlyList<LiveStoreSize> sizes)
    {
        ArgumentNullException.ThrowIfNull(sizes);
        if (sizes.Count == 0)
        {
            throw new InvalidOperationException(
                "There is no indexed store to time a latency bound on. The index tier's latency tests measure "
                + "the largest indexed store, and this machine's settings name none.");
        }

        foreach (LiveStoreSize size in sizes)
        {
            if (size.Items == null || !size.Complete)
            {
                throw new InvalidOperationException(
                    $"Cannot tell which indexed store is largest: '{size.Store}' "
                    + (size.Items == null ? "could not be counted" : "was only partly counted - its folder walk stopped short")
                    + ". A latency bound timed on a store that merely LOOKED largest proves less than it says, so this "
                    + "refuses instead of guessing. Every indexed store's folders must report their item counts.");
            }
        }

        LiveStoreSize largest = sizes[0];
        foreach (LiveStoreSize size in sizes)
        {
            if (size.Items!.Value > largest.Items!.Value)
            {
                largest = size;
            }
        }

        return largest.Store;
    }

    /// <summary>The line a latency test prints so the log says what it was timed on and why.</summary>
    public static string Describe(IReadOnlyList<LiveStoreSize> sizes, string chosen)
    {
        ArgumentNullException.ThrowIfNull(sizes);
        return "latency target: '" + chosen + "', the largest indexed store ("
            + string.Join(", ", sizes.Select(s => s.Store + "=" + (s.Items?.ToString("N0", CultureInfo.InvariantCulture) ?? "?")))
            + " items)";
    }
}
