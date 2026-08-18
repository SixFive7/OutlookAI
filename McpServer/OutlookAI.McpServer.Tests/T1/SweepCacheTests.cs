using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// D34 sweep-cache logic (pure, no COM): the ~10 s TTL constant is PINNED (product
/// contract - the accepted staleness window for rapid-fire searches), and the reuse
/// rules are covered: fresh entry reuse, TTL expiry, frontier-advance invalidation,
/// store-scope compatibility (all-stores serves store-scoped, never the reverse, and only
/// while both describe the same window - the check that keeps a scope-aware frontier from
/// needing a key change),
/// FOLDER-scope separation (soak fix 13: a folder-scoped sweep covers one subtree and
/// must never answer a broader query), the SUBTREE FLAG in the key (soak fix 15 /
/// constraint C6: a shallow sweep must never answer a recursive query), and Clear().
/// </summary>
public sealed class SweepCacheTests
{
    private static readonly DateTime Frontier = new(2026, 7, 24, 10, 0, 0, DateTimeKind.Utc);
    private static readonly DateTime Now = new(2026, 7, 24, 12, 0, 0, DateTimeKind.Utc);

    // The default folder set is swept shallowly, so a non-folder-scoped request always
    // carries the shallow flag (MailService: sweepRecursive = folder != null && flag).
    private const bool Shallow = false;
    private const bool Recursive = true;

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
        cache.Store(Frontier, store: null, folder: null, Shallow, MakeResult(3), elapsedMs: 120, Now);

        bool hit = cache.TryGet(Frontier, store: null, folder: null, Shallow, Now.AddSeconds(9), out SweepCache.CachedSweep? cached);

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
        cache.Store(Frontier, store: null, folder: null, Shallow, MakeResult(1), elapsedMs: 50, Now);

        Assert.False(cache.TryGet(Frontier, store: null, folder: null, Shallow, Now + SweepCache.DefaultTimeToLive + TimeSpan.FromMilliseconds(1), out _));
    }

    [Fact]
    public void FrontierAdvance_InvalidatesTheEntry()
    {
        SweepCache cache = new();
        cache.Store(Frontier, store: null, folder: null, Shallow, MakeResult(1), elapsedMs: 50, Now);

        // The index ingested new mail: the window base moved - the cached sweep no
        // longer represents the current gap and must not be reused (D34 rule).
        Assert.False(cache.TryGet(Frontier.AddMinutes(2), store: null, folder: null, Shallow, Now.AddSeconds(1), out _));
    }

    [Fact]
    public void AllStoresEntry_ServesStoreScopedRequest_ButNotViceVersa()
    {
        SweepCache cache = new();
        cache.Store(Frontier, store: null, folder: null, Shallow, MakeResult(2), elapsedMs: 80, Now);

        // All-stores sweep covers any single store (caller filters items by store) -
        // sound because every store gets the identical default folder set.
        Assert.True(cache.TryGet(Frontier, store: "someone@example.com", folder: null, Shallow, Now.AddSeconds(5), out SweepCache.CachedSweep? cached));
        Assert.Null(cached!.Store);

        // A store-scoped sweep must never serve an all-stores request.
        SweepCache scoped = new();
        scoped.Store(Frontier, store: "someone@example.com", folder: null, Shallow, MakeResult(2), elapsedMs: 80, Now);
        Assert.False(scoped.TryGet(Frontier, store: null, folder: null, Shallow, Now.AddSeconds(5), out _));

        // Nor a request for a DIFFERENT store.
        Assert.False(scoped.TryGet(Frontier, store: "other@example.com", folder: null, Shallow, Now.AddSeconds(5), out _));

        // But the exact store matches, case-insensitively.
        Assert.True(scoped.TryGet(Frontier, store: "SOMEONE@example.com", folder: null, Shallow, Now.AddSeconds(5), out _));
    }

    [Fact]
    public void AllStoresEntry_ServesAStoreScopedRequest_OnlyWhenTheWindowIsTheSame()
    {
        // Why the key needed no change when the staleness frontier became scope-aware.
        //
        // The window base is now measured over the store being searched, so a quiet store
        // and the profile disagree about it whenever another store is further ahead. The
        // base is not part of the KEY - it is checked for EQUALITY on the entry - so the
        // broad-serves-narrow reuse now fires only when the two windows are identical,
        // which is exactly when the broad sweep covered what the narrow one would have.
        SweepCache cache = new();
        cache.Store(Frontier, store: null, folder: null, Shallow, MakeResult(2), elapsedMs: 80, Now);

        // The quiet store's own frontier lags the profile's: its window starts earlier and
        // is WIDER, so the all-stores sweep (which started later) cannot answer it.
        Assert.False(cache.TryGet(
            Frontier.AddHours(-45), store: "quiet@example.com", folder: null, Shallow, Now.AddSeconds(1), out _));

        // The store that set the profile frontier shares the window, and is served.
        Assert.True(cache.TryGet(
            Frontier, store: "busy@example.com", folder: null, Shallow, Now.AddSeconds(1), out _));
    }

    // ------------------------------------------- folder scope in the key (soak fix 13)

    [Fact]
    public void FolderScopedEntry_NeverServesABroaderRequest()
    {
        // The correctness bug this pins: a sweep of ONE folder subtree answering a
        // store-wide or all-stores query would report a fraction of the coverage as if
        // it were the whole freshness gap.
        SweepCache cache = new();
        cache.Store(Frontier, store: "someone@example.com", folder: "Projects/2026", Recursive, MakeResult(2), elapsedMs: 40, Now);

        Assert.False(cache.TryGet(Frontier, store: "someone@example.com", folder: null, Shallow, Now.AddSeconds(1), out _));
        Assert.False(cache.TryGet(Frontier, store: null, folder: null, Shallow, Now.AddSeconds(1), out _));
        Assert.False(cache.TryGet(Frontier, store: "someone@example.com", folder: "Projects", Recursive, Now.AddSeconds(1), out _));
        Assert.False(cache.TryGet(Frontier, store: "someone@example.com", folder: "Projects/2025", Recursive, Now.AddSeconds(1), out _));

        // Only the identical folder scope is served.
        Assert.True(cache.TryGet(Frontier, store: "someone@example.com", folder: "Projects/2026", Recursive, Now.AddSeconds(1), out _));
    }

    [Fact]
    public void DefaultFolderEntry_DoesNotServeAFolderScopedRequest()
    {
        // The default set sweeps Inbox/Sent/Deleted/Junk NON-recursively, so it cannot
        // stand in for a folder-scoped sweep - not even for one of those folders,
        // whose scoped sweep also covers subfolders.
        SweepCache cache = new();
        cache.Store(Frontier, store: null, folder: null, Shallow, MakeResult(2), elapsedMs: 80, Now);

        Assert.False(cache.TryGet(Frontier, store: "someone@example.com", folder: "Inbox", Recursive, Now.AddSeconds(1), out _));
        Assert.False(cache.TryGet(Frontier, store: null, folder: "Inbox", Recursive, Now.AddSeconds(1), out _));

        // Not even for a SHALLOW folder request: the default set is a fixed four folders
        // of every store, not "whatever folder you name".
        Assert.False(cache.TryGet(Frontier, store: "someone@example.com", folder: "Inbox", Shallow, Now.AddSeconds(1), out _));
    }

    [Fact]
    public void FolderScopes_OfTheSameStore_AreSeparateEntries()
    {
        SweepCache cache = new();
        cache.Store(Frontier, store: "someone@example.com", folder: "Inbox", Recursive, MakeResult(1), elapsedMs: 10, Now);
        cache.Store(Frontier, store: "someone@example.com", folder: "Deleted Items", Recursive, MakeResult(4), elapsedMs: 15, Now);

        Assert.True(cache.TryGet(Frontier, store: "someone@example.com", folder: "Inbox", Recursive, Now.AddSeconds(1), out SweepCache.CachedSweep? inbox));
        Assert.Single(inbox!.Result.Items);
        Assert.Equal("Inbox", inbox.Folder);

        Assert.True(cache.TryGet(Frontier, store: "someone@example.com", folder: "Deleted Items", Recursive, Now.AddSeconds(1), out SweepCache.CachedSweep? deleted));
        Assert.Equal(4, deleted!.Result.Items.Count);
        Assert.Equal("Deleted Items", deleted.Folder);
    }

    // ------------------------------- include_subfolders in the key (soak fix 15 / C6)

    [Fact]
    public void ShallowSweep_NeverAnswersARecursiveQuery_AndViceVersa()
    {
        // The bug this pins: with include_subfolders in the request but NOT in the key,
        // a search that swept ONE folder would answer the next search that asked for
        // that folder AND its subtree - reporting a fraction of the coverage as the
        // whole freshness gap (v3.MD constraint C6).
        SweepCache cache = new();
        cache.Store(Frontier, store: "someone@example.com", folder: "Projects", Shallow, MakeResult(1), elapsedMs: 10, Now);

        Assert.False(cache.TryGet(Frontier, store: "someone@example.com", folder: "Projects", Recursive, Now.AddSeconds(1), out _));
        Assert.True(cache.TryGet(Frontier, store: "someone@example.com", folder: "Projects", Shallow, Now.AddSeconds(1), out SweepCache.CachedSweep? shallow));
        Assert.False(shallow!.IncludeSubfolders);

        SweepCache deep = new();
        deep.Store(Frontier, store: "someone@example.com", folder: "Projects", Recursive, MakeResult(9), elapsedMs: 90, Now);

        // The reverse direction is also refused: a recursive sweep over-covers a shallow
        // request, which would inflate the reported folder count for that query.
        Assert.False(deep.TryGet(Frontier, store: "someone@example.com", folder: "Projects", Shallow, Now.AddSeconds(1), out _));
        Assert.True(deep.TryGet(Frontier, store: "someone@example.com", folder: "Projects", Recursive, Now.AddSeconds(1), out SweepCache.CachedSweep? recursive));
        Assert.True(recursive!.IncludeSubfolders);
    }

    [Fact]
    public void BothSubtreeScopes_OfOneFolder_CoexistAsSeparateEntries()
    {
        SweepCache cache = new();
        cache.Store(Frontier, store: "someone@example.com", folder: "Projects", Shallow, MakeResult(1), elapsedMs: 10, Now);
        cache.Store(Frontier, store: "someone@example.com", folder: "Projects", Recursive, MakeResult(7), elapsedMs: 70, Now);

        Assert.True(cache.TryGet(Frontier, store: "someone@example.com", folder: "Projects", Shallow, Now.AddSeconds(1), out SweepCache.CachedSweep? one));
        Assert.Single(one!.Result.Items);

        Assert.True(cache.TryGet(Frontier, store: "someone@example.com", folder: "Projects", Recursive, Now.AddSeconds(1), out SweepCache.CachedSweep? many));
        Assert.Equal(7, many!.Result.Items.Count);
    }

    [Fact]
    public void KeyParts_CannotBlurIntoEachOther()
    {
        // Naive concatenation would let store "a" + folder "b" collide with a store
        // literally named "a/b" (or "ab") - and, with the flag appended, with a store
        // "a" + folder "b0"/"b1".
        SweepCache cache = new();
        cache.Store(Frontier, store: "a", folder: "b", Recursive, MakeResult(1), elapsedMs: 10, Now);

        Assert.False(cache.TryGet(Frontier, store: "a/b", folder: null, Shallow, Now.AddSeconds(1), out _));
        Assert.False(cache.TryGet(Frontier, store: "ab", folder: null, Shallow, Now.AddSeconds(1), out _));
        Assert.False(cache.TryGet(Frontier, store: "a", folder: "b1", Recursive, Now.AddSeconds(1), out _));
        Assert.False(cache.TryGet(Frontier, store: "a", folder: "b0", Shallow, Now.AddSeconds(1), out _));
    }

    [Fact]
    public void ExactStoreEntry_WinsOverAllStoresEntry()
    {
        SweepCache cache = new();
        cache.Store(Frontier, store: null, folder: null, Shallow, MakeResult(5), elapsedMs: 200, Now);
        cache.Store(Frontier, store: "someone@example.com", folder: null, Shallow, MakeResult(1), elapsedMs: 30, Now.AddSeconds(2));

        Assert.True(cache.TryGet(Frontier, store: "someone@example.com", folder: null, Shallow, Now.AddSeconds(4), out SweepCache.CachedSweep? cached));
        Assert.Equal("someone@example.com", cached!.Store);
        Assert.Single(cached.Result.Items);
    }

    [Fact]
    public void Clear_DropsEverything()
    {
        SweepCache cache = new();
        cache.Store(Frontier, store: null, folder: null, Shallow, MakeResult(1), elapsedMs: 10, Now);
        cache.Store(Frontier, store: "someone@example.com", folder: null, Shallow, MakeResult(1), elapsedMs: 10, Now);
        cache.Store(Frontier, store: "someone@example.com", folder: "Inbox", Recursive, MakeResult(1), elapsedMs: 10, Now);

        cache.Clear();

        Assert.False(cache.TryGet(Frontier, store: null, folder: null, Shallow, Now.AddSeconds(1), out _));
        Assert.False(cache.TryGet(Frontier, store: "someone@example.com", folder: null, Shallow, Now.AddSeconds(1), out _));
        Assert.False(cache.TryGet(Frontier, store: "someone@example.com", folder: "Inbox", Recursive, Now.AddSeconds(1), out _));
    }

    [Fact]
    public void ZeroTtl_DisablesTheCache()
    {
        SweepCache cache = new(TimeSpan.Zero);
        cache.Store(Frontier, store: null, folder: null, Shallow, MakeResult(1), elapsedMs: 10, Now);

        Assert.False(cache.TryGet(Frontier, store: null, folder: null, Shallow, Now, out _));
    }

    [Fact]
    public void ClockSkew_EntryFromTheFuture_IsNotServed()
    {
        SweepCache cache = new();
        cache.Store(Frontier, store: null, folder: null, Shallow, MakeResult(1), elapsedMs: 10, Now);

        // A caller clock BEHIND the entry timestamp must not serve the entry (defensive).
        Assert.False(cache.TryGet(Frontier, store: null, folder: null, Shallow, Now.AddSeconds(-1), out _));
    }

    // ---------------------------------- per-store windows: what "the same sweep" now means

    // A sweep no longer has ONE window: an unscoped sweep opens one per store, from that
    // store's own index frontier. The cache key never carried them, so the scalar base alone
    // decided reuse - and the scalar base of a broad sweep is the FALLBACK window, which
    // describes no particular store. The audit's question, answered below: a broad entry
    // must never serve a narrow request under weaker coverage than a fresh sweep would give.

    private static readonly Dictionary<string, DateTime> TwoStoreWindows = new(StringComparer.OrdinalIgnoreCase)
    {
        ["alice@example.com"] = Frontier,
        ["bob@example.com"] = Frontier.AddHours(-9),
    };

    [Fact]
    public void OneStoresFrontierAdvancing_InvalidatesTheBroadEntry_EvenThoughTheScalarBaseIsUnchanged()
    {
        // The exact hole the scalar key left: the fallback base is identical across both
        // calls, so the old comparison saw "same window" while store B's own window had
        // moved. Equality, not containment - a frontier advance means the index ingested
        // something and the cache must not outlive that.
        SweepCache cache = new();
        cache.Store(Frontier, null, null, Shallow, MakeResult(2), 100, Now, TwoStoreWindows);

        Dictionary<string, DateTime> moved = new(TwoStoreWindows, StringComparer.OrdinalIgnoreCase)
        {
            ["bob@example.com"] = Frontier.AddHours(-8),
        };

        Assert.False(cache.TryGet(Frontier, null, null, Shallow, Now.AddSeconds(1), out _, moved));
        Assert.True(cache.TryGet(Frontier, null, null, Shallow, Now.AddSeconds(1), out _, TwoStoreWindows));
    }

    [Fact]
    public void AStoreAppearingOrVanishingFromTheWindowSet_IsADifferentSweep()
    {
        SweepCache cache = new();
        cache.Store(Frontier, null, null, Shallow, MakeResult(2), 100, Now, TwoStoreWindows);

        Dictionary<string, DateTime> fewer = new(StringComparer.OrdinalIgnoreCase)
        {
            ["alice@example.com"] = Frontier,
        };

        Assert.False(cache.TryGet(Frontier, null, null, Shallow, Now.AddSeconds(1), out _, fewer));

        // And an entry taken WITHOUT per-store windows cannot serve a request that has them.
        SweepCache legacy = new();
        legacy.Store(Frontier, null, null, Shallow, MakeResult(2), 100, Now);
        Assert.False(legacy.TryGet(Frontier, null, null, Shallow, Now.AddSeconds(1), out _, TwoStoreWindows));
    }

    [Fact]
    public void ABroadEntryServesAStoreScopedRequest_OnlyWhenItSweptThatStoreFromTheSameInstant()
    {
        // A store-scoped request has ONE window - its own store's - and the broad entry's
        // window for that store is what decides, not its fallback base. Store B's window in
        // the entry is 9 h before the fallback, so a B-scoped request whose own base IS that
        // instant is served, and one whose base has moved on is not.
        SweepCache cache = new();
        cache.Store(Frontier, null, null, Shallow, MakeResult(2), 100, Now, TwoStoreWindows);

        Assert.True(cache.TryGet(
            Frontier.AddHours(-9), "bob@example.com", null, Shallow, Now.AddSeconds(2), out SweepCache.CachedSweep? served));
        Assert.NotNull(served);
        Assert.Equal(Frontier.AddHours(-9), served!.WindowFor("bob@example.com"));

        // Its frontier advanced by a minute: a fresh sweep would open a NARROWER window than
        // the entry has, and reusing it would answer from a sweep taken before that ingest.
        Assert.False(cache.TryGet(Frontier.AddHours(-9).AddMinutes(1), "bob@example.com", null, Shallow, Now.AddSeconds(2), out _));

        // A store the broad sweep had no window for falls back to the entry's scalar base,
        // which is exactly the window such a store was swept with.
        Assert.True(cache.TryGet(Frontier, "Archive 2019.pst", null, Shallow, Now.AddSeconds(2), out _));
    }

    [Fact]
    public void TheBroadEntrysWindowForAStore_IsMatchedCaseInsensitively()
    {
        SweepCache cache = new();
        cache.Store(Frontier, null, null, Shallow, MakeResult(2), 100, Now, TwoStoreWindows);

        Assert.True(cache.TryGet(Frontier.AddHours(-9), "BOB@EXAMPLE.COM", null, Shallow, Now.AddSeconds(1), out _));
    }

    [Fact]
    public void ABroadEntryStillNeverServesAcrossFolderScopes_WhateverTheWindows()
    {
        // Unchanged by per-store windows, and restated because the store-scoped path now has
        // its own comparison: a folder-scoped sweep covers ONE subtree and answering a
        // store-wide query from it would report a fraction of the coverage as all of it.
        SweepCache cache = new();
        cache.Store(Frontier, null, null, Shallow, MakeResult(2), 100, Now, TwoStoreWindows);

        Assert.False(cache.TryGet(Frontier.AddHours(-9), "bob@example.com", "Inbox", Recursive, Now.AddSeconds(1), out _));
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
