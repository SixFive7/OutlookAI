using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// D34 sweep-cache logic (pure, no COM): the ~10 s TTL constant is PINNED (product
/// contract - the accepted staleness window for rapid-fire searches), and the reuse
/// rules are covered: fresh entry reuse, TTL expiry, frontier-advance invalidation,
/// store-scope compatibility (all-stores serves store-scoped, never the reverse),
/// FOLDER-scope separation (soak fix 13: a folder-scoped sweep covers one subtree and
/// must never answer a broader query), and Clear().
/// </summary>
public sealed class SweepCacheTests
{
    private static readonly DateTime Frontier = new(2026, 7, 24, 10, 0, 0, DateTimeKind.Utc);
    private static readonly DateTime Now = new(2026, 7, 24, 12, 0, 0, DateTimeKind.Utc);

    [Fact]
    public void TimeToLive_IsPinnedAtTenSeconds()
    {
        // D34 product constant: rapid-fire iterative searches reuse one sweep for at
        // most this long. Changing it is a decision, not a refactor (20 s -> 10 s
        // per user order, 2026-07-24).
        Assert.Equal(TimeSpan.FromSeconds(10), SweepCache.DefaultTimeToLive);
    }

    [Fact]
    public void FreshEntry_IsReused_WithinTtl()
    {
        SweepCache cache = new();
        cache.Store(Frontier, store: null, folder: null, MakeResult(3), elapsedMs: 120, Now);

        bool hit = cache.TryGet(Frontier, store: null, folder: null, Now.AddSeconds(9), out SweepCache.CachedSweep? cached);

        Assert.True(hit);
        Assert.NotNull(cached);
        Assert.Equal(3, cached!.Result.Items.Count);
        Assert.Equal(120, cached.ElapsedMs);
        Assert.Equal(Frontier, cached.BaseGapStartUtc);
    }

    [Fact]
    public void Entry_Expires_AfterTtl()
    {
        SweepCache cache = new();
        cache.Store(Frontier, store: null, folder: null, MakeResult(1), elapsedMs: 50, Now);

        Assert.False(cache.TryGet(Frontier, store: null, folder: null, Now + SweepCache.DefaultTimeToLive + TimeSpan.FromMilliseconds(1), out _));
    }

    [Fact]
    public void FrontierAdvance_InvalidatesTheEntry()
    {
        SweepCache cache = new();
        cache.Store(Frontier, store: null, folder: null, MakeResult(1), elapsedMs: 50, Now);

        // The index ingested new mail: the window base moved - the cached sweep no
        // longer represents the current gap and must not be reused (D34 rule).
        Assert.False(cache.TryGet(Frontier.AddMinutes(2), store: null, folder: null, Now.AddSeconds(1), out _));
    }

    [Fact]
    public void AllStoresEntry_ServesStoreScopedRequest_ButNotViceVersa()
    {
        SweepCache cache = new();
        cache.Store(Frontier, store: null, folder: null, MakeResult(2), elapsedMs: 80, Now);

        // All-stores sweep covers any single store (caller filters items by store) -
        // sound because every store gets the identical default folder set.
        Assert.True(cache.TryGet(Frontier, store: "someone@example.com", folder: null, Now.AddSeconds(5), out SweepCache.CachedSweep? cached));
        Assert.Null(cached!.Store);

        // A store-scoped sweep must never serve an all-stores request.
        SweepCache scoped = new();
        scoped.Store(Frontier, store: "someone@example.com", folder: null, MakeResult(2), elapsedMs: 80, Now);
        Assert.False(scoped.TryGet(Frontier, store: null, folder: null, Now.AddSeconds(5), out _));

        // Nor a request for a DIFFERENT store.
        Assert.False(scoped.TryGet(Frontier, store: "other@example.com", folder: null, Now.AddSeconds(5), out _));

        // But the exact store matches, case-insensitively.
        Assert.True(scoped.TryGet(Frontier, store: "SOMEONE@example.com", folder: null, Now.AddSeconds(5), out _));
    }

    // ------------------------------------------- folder scope in the key (soak fix 13)

    [Fact]
    public void FolderScopedEntry_NeverServesABroaderRequest()
    {
        // The correctness bug this pins: a sweep of ONE folder subtree answering a
        // store-wide or all-stores query would report a fraction of the coverage as if
        // it were the whole freshness gap.
        SweepCache cache = new();
        cache.Store(Frontier, store: "someone@example.com", folder: "Projects/2026", MakeResult(2), elapsedMs: 40, Now);

        Assert.False(cache.TryGet(Frontier, store: "someone@example.com", folder: null, Now.AddSeconds(1), out _));
        Assert.False(cache.TryGet(Frontier, store: null, folder: null, Now.AddSeconds(1), out _));
        Assert.False(cache.TryGet(Frontier, store: "someone@example.com", folder: "Projects", Now.AddSeconds(1), out _));
        Assert.False(cache.TryGet(Frontier, store: "someone@example.com", folder: "Projects/2025", Now.AddSeconds(1), out _));

        // Only the identical folder scope is served.
        Assert.True(cache.TryGet(Frontier, store: "someone@example.com", folder: "Projects/2026", Now.AddSeconds(1), out _));
    }

    [Fact]
    public void DefaultFolderEntry_DoesNotServeAFolderScopedRequest()
    {
        // The default set sweeps Inbox/Sent/Deleted/Junk NON-recursively, so it cannot
        // stand in for a folder-scoped sweep - not even for one of those folders,
        // whose scoped sweep also covers subfolders.
        SweepCache cache = new();
        cache.Store(Frontier, store: null, folder: null, MakeResult(2), elapsedMs: 80, Now);

        Assert.False(cache.TryGet(Frontier, store: "someone@example.com", folder: "Inbox", Now.AddSeconds(1), out _));
        Assert.False(cache.TryGet(Frontier, store: null, folder: "Inbox", Now.AddSeconds(1), out _));
    }

    [Fact]
    public void FolderScopes_OfTheSameStore_AreSeparateEntries()
    {
        SweepCache cache = new();
        cache.Store(Frontier, store: "someone@example.com", folder: "Inbox", MakeResult(1), elapsedMs: 10, Now);
        cache.Store(Frontier, store: "someone@example.com", folder: "Deleted Items", MakeResult(4), elapsedMs: 15, Now);

        Assert.True(cache.TryGet(Frontier, store: "someone@example.com", folder: "Inbox", Now.AddSeconds(1), out SweepCache.CachedSweep? inbox));
        Assert.Single(inbox!.Result.Items);
        Assert.Equal("Inbox", inbox.Folder);

        Assert.True(cache.TryGet(Frontier, store: "someone@example.com", folder: "Deleted Items", Now.AddSeconds(1), out SweepCache.CachedSweep? deleted));
        Assert.Equal(4, deleted!.Result.Items.Count);
        Assert.Equal("Deleted Items", deleted.Folder);
    }

    [Fact]
    public void KeyParts_CannotBlurIntoEachOther()
    {
        // Naive concatenation would let store "a" + folder "b" collide with a store
        // literally named "a/b" (or "ab").
        SweepCache cache = new();
        cache.Store(Frontier, store: "a", folder: "b", MakeResult(1), elapsedMs: 10, Now);

        Assert.False(cache.TryGet(Frontier, store: "a/b", folder: null, Now.AddSeconds(1), out _));
        Assert.False(cache.TryGet(Frontier, store: "ab", folder: null, Now.AddSeconds(1), out _));
    }

    [Fact]
    public void ExactStoreEntry_WinsOverAllStoresEntry()
    {
        SweepCache cache = new();
        cache.Store(Frontier, store: null, folder: null, MakeResult(5), elapsedMs: 200, Now);
        cache.Store(Frontier, store: "someone@example.com", folder: null, MakeResult(1), elapsedMs: 30, Now.AddSeconds(2));

        Assert.True(cache.TryGet(Frontier, store: "someone@example.com", folder: null, Now.AddSeconds(4), out SweepCache.CachedSweep? cached));
        Assert.Equal("someone@example.com", cached!.Store);
        Assert.Single(cached.Result.Items);
    }

    [Fact]
    public void Clear_DropsEverything()
    {
        SweepCache cache = new();
        cache.Store(Frontier, store: null, folder: null, MakeResult(1), elapsedMs: 10, Now);
        cache.Store(Frontier, store: "someone@example.com", folder: null, MakeResult(1), elapsedMs: 10, Now);
        cache.Store(Frontier, store: "someone@example.com", folder: "Inbox", MakeResult(1), elapsedMs: 10, Now);

        cache.Clear();

        Assert.False(cache.TryGet(Frontier, store: null, folder: null, Now.AddSeconds(1), out _));
        Assert.False(cache.TryGet(Frontier, store: "someone@example.com", folder: null, Now.AddSeconds(1), out _));
        Assert.False(cache.TryGet(Frontier, store: "someone@example.com", folder: "Inbox", Now.AddSeconds(1), out _));
    }

    [Fact]
    public void ZeroTtl_DisablesTheCache()
    {
        SweepCache cache = new(TimeSpan.Zero);
        cache.Store(Frontier, store: null, folder: null, MakeResult(1), elapsedMs: 10, Now);

        Assert.False(cache.TryGet(Frontier, store: null, folder: null, Now, out _));
    }

    [Fact]
    public void ClockSkew_EntryFromTheFuture_IsNotServed()
    {
        SweepCache cache = new();
        cache.Store(Frontier, store: null, folder: null, MakeResult(1), elapsedMs: 10, Now);

        // A caller clock BEHIND the entry timestamp must not serve the entry (defensive).
        Assert.False(cache.TryGet(Frontier, store: null, folder: null, Now.AddSeconds(-1), out _));
    }

    private static ComSweepResult MakeResult(int items)
    {
        var list = new List<ComMailBrief>(items);
        for (int i = 0; i < items; i++)
        {
            list.Add(new ComMailBrief(
                entryId: "00AB" + i.ToString("X4"),
                storeDisplayName: "someone@example.com",
                storeId: null,
                folderName: "Inbox",
                folderKind: "inbox",
                subject: "s" + i,
                senderName: null,
                senderAddress: null,
                receivedTime: Frontier,
                isRead: null,
                hasAttachments: null,
                sizeBytes: null,
                body: null));
        }

        return new ComSweepResult(list, foldersSwept: 2, foldersSkipped: 0, sweptFolders: new[]
        {
            "someone@example.com/Inbox", "someone@example.com/Sent Items",
        });
    }
}
