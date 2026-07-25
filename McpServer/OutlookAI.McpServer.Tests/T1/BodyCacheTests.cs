using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Soak fix D37 (read body paging): the per-process full-body cache that serves
/// body_offset continuation windows without re-transferring the body over COM, plus
/// the pure window math read builds its payload from.
/// </summary>
public sealed class BodyCacheTests
{
    [Fact]
    public void PutAndGet_RoundTrips_BodyAndOrigin()
    {
        BodyCache cache = new();

        cache.Put("AABB01", "hello body", "text");

        Assert.True(cache.TryGet("AABB01", out string body, out string origin));
        Assert.Equal("hello body", body);
        Assert.Equal("text", origin);
        // EntryID hex is matched case-insensitively.
        Assert.True(cache.TryGet("aabb01", out _, out _));
    }

    [Fact]
    public void Get_Miss_ReturnsFalse()
    {
        BodyCache cache = new();

        Assert.False(cache.TryGet("DOESNOTEXIST", out string body, out string origin));
        Assert.Equal(string.Empty, body);
        Assert.Equal("none", origin);
    }

    [Fact]
    public void Put_RefreshesExistingEntry()
    {
        BodyCache cache = new();
        cache.Put("AA", "old", "text");

        cache.Put("AA", "new", "html-converted");

        Assert.True(cache.TryGet("AA", out string body, out string origin));
        Assert.Equal("new", body);
        Assert.Equal("html-converted", origin);
        Assert.Equal(1, cache.Count);
    }

    [Fact]
    public void EntryBound_EvictsOldestFirst_NeverTheJustInserted()
    {
        DateTime now = new(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        BodyCache cache = new(() => now);

        for (int i = 1; i <= BodyCache.MaxEntries + 2; i++)
        {
            cache.Put("ID" + i, "body" + i, "text");
            now = now.AddSeconds(1);
        }

        Assert.Equal(BodyCache.MaxEntries, cache.Count);
        Assert.False(cache.TryGet("ID1", out _, out _), "oldest entry must be evicted");
        Assert.False(cache.TryGet("ID2", out _, out _), "second-oldest entry must be evicted");
        Assert.True(cache.TryGet("ID" + (BodyCache.MaxEntries + 2), out _, out _), "newest entry must survive");
    }

    [Fact]
    public void CharBound_EvictsOthers_ButAlwaysKeepsTheNewGiant()
    {
        DateTime now = new(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        BodyCache cache = new(() => now);
        cache.Put("SMALL", "x", "text");
        now = now.AddSeconds(1);

        // A single body larger than the whole char bound still gets cached (paging a
        // giant body is exactly what the cache exists for); the older entry goes.
        string giant = new('g', BodyCache.MaxTotalChars + 1);
        cache.Put("GIANT", giant, "text");

        Assert.True(cache.TryGet("GIANT", out string body, out _));
        Assert.Equal(giant.Length, body.Length);
        Assert.False(cache.TryGet("SMALL", out _, out _));
        Assert.Equal(1, cache.Count);
    }

    [Fact]
    public void ExpiredEntry_IsMissAndEvicted()
    {
        DateTime now = new(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        BodyCache cache = new(() => now);
        cache.Put("AA", "body", "text");

        now = now.Add(BodyCache.TimeToLive).AddSeconds(1);

        Assert.False(cache.TryGet("AA", out _, out _));
        Assert.Equal(0, cache.Count);
    }

    [Fact]
    public void WithinTtl_EntryStillServes()
    {
        DateTime now = new(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        BodyCache cache = new(() => now);
        cache.Put("AA", "body", "text");

        now = now.Add(BodyCache.TimeToLive).AddSeconds(-1);

        Assert.True(cache.TryGet("AA", out _, out _));
    }

    [Fact]
    public void Invalidate_RemovesEntry()
    {
        BodyCache cache = new();
        cache.Put("AA", "body", "text");

        cache.Invalidate("AA");

        Assert.False(cache.TryGet("AA", out _, out _));
    }

    // ------------------------------------------------------------------ window math

    [Fact]
    public void Window_AtStart_ReturnsPrefix_FlagsMore()
    {
        (int start, string window, bool more) = MailService.ComputeBodyWindow("0123456789", offset: 0, maxChars: 4);

        Assert.Equal(0, start);
        Assert.Equal("0123", window);
        Assert.True(more);
    }

    [Fact]
    public void Window_Continuation_TilesExactly()
    {
        string body = "0123456789";
        (int s1, string w1, bool m1) = MailService.ComputeBodyWindow(body, 0, 4);
        (int s2, string w2, bool m2) = MailService.ComputeBodyWindow(body, s1 + w1.Length, 4);
        (int s3, string w3, bool m3) = MailService.ComputeBodyWindow(body, s2 + w2.Length, 4);

        Assert.True(m1);
        Assert.True(m2);
        Assert.False(m3, "the final window must clear the has-more flag");
        Assert.Equal(body, w1 + w2 + w3);
        Assert.Equal(8, s3);
    }

    [Fact]
    public void Window_OffsetBeyondEnd_EmptyNotMore()
    {
        (int start, string window, bool more) = MailService.ComputeBodyWindow("abc", offset: 99, maxChars: 10);

        Assert.Equal(3, start);
        Assert.Equal(string.Empty, window);
        Assert.False(more);
    }

    [Fact]
    public void Window_ZeroMaxChars_MetadataOnly_FlagsMoreWhenBodyExists()
    {
        (int start, string window, bool more) = MailService.ComputeBodyWindow("abc", offset: 0, maxChars: 0);

        Assert.Equal(0, start);
        Assert.Equal(string.Empty, window);
        Assert.True(more);
    }

    [Fact]
    public void Window_WholeBodyInOneWindow_NotMore()
    {
        (int start, string window, bool more) = MailService.ComputeBodyWindow("abc", offset: 0, maxChars: 3);

        Assert.Equal("abc", window);
        Assert.False(more);
    }

    [Fact]
    public void Window_NegativeInputs_AreClamped()
    {
        (int start, string window, bool more) = MailService.ComputeBodyWindow("abc", offset: -4, maxChars: -1);

        Assert.Equal(0, start);
        Assert.Equal(string.Empty, window);
        Assert.True(more);
    }
}
