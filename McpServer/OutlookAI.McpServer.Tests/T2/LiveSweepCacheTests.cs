using System.Diagnostics;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// D34 live acceptance for the sweep cache: two rapid-fire searches pay ONE COM sweep -
/// the first performs it live, the second is served from the cache (cached=true,
/// elapsedMs=0) at index speed. Also proves ClearSweepCache() forces the next sweep
/// live again. Store-scoped to the test hub (S2); logging is counts/timings only (S4).
/// A frontier advance between calls (new mail indexed mid-test on this live machine)
/// legitimately invalidates the cache, so the rapid pair retries a few times.
/// </summary>
[Collection("LivePhase2")]
[Trait("Category", "Live")]
public sealed class LiveSweepCacheTests
{
    private readonly LivePhase2Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveSweepCacheTests(LivePhase2Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    public void RapidSearches_PayOneSweep_SecondServedFromCache()
    {
        MailService service = _fixture.Service;
        string hub = _fixture.Settings.TestHubStoreDisplayName;

        SearchRequest MakeRequest() => new()
        {
            Store = hub,
            Top = 5,
            SnippetChars = 0,
        };

        bool proven = false;
        for (int attempt = 0; attempt < 3 && !proven; attempt++)
        {
            service.ClearSweepCache();

            Stopwatch first = Stopwatch.StartNew();
            SearchOutcome one = service.Search(MakeRequest());
            first.Stop();

            Assert.NotNull(one.Sweep);
            Assert.True(one.Sweep!.Performed, "first sweep must run live (error: " + (one.Sweep.Error ?? "-") + ")");
            Assert.Null(one.Sweep.Cached);

            Stopwatch second = Stopwatch.StartNew();
            SearchOutcome two = service.Search(MakeRequest());
            second.Stop();

            Assert.NotNull(two.Sweep);
            Assert.True(two.Sweep!.Performed);
            proven = two.Sweep.Cached == true;
            _output.WriteLine($"attempt {attempt + 1}: firstCallMs={first.ElapsedMilliseconds} "
                + $"(sweepMs={one.Sweep.ElapsedMs}) secondCallMs={second.ElapsedMilliseconds} "
                + $"cached={two.Sweep.Cached} cacheAgeSeconds={two.Sweep.CacheAgeSeconds}");

            if (proven)
            {
                // Cached sweeps report zero sweep cost and carry the cache age.
                Assert.Equal(0, two.Sweep.ElapsedMs);
                Assert.NotNull(two.Sweep.CacheAgeSeconds);
                Assert.InRange(two.Sweep.CacheAgeSeconds!.Value, 0, SweepCache.DefaultTimeToLive.TotalSeconds);
                Assert.True(second.ElapsedMilliseconds <= Math.Max(2000, first.ElapsedMilliseconds),
                    "the cached call must not be slower than the live-sweep call");
            }
        }

        Assert.True(proven, "a rapid follow-up search was never served from the sweep cache (D34 regression)");

        // ClearSweepCache() drops the entry: the next sweep runs live again.
        _fixture.Service.ClearSweepCache();
        SearchOutcome three = _fixture.Service.Search(MakeRequest());
        Assert.NotNull(three.Sweep);
        Assert.True(three.Sweep!.Performed);
        Assert.Null(three.Sweep.Cached);
        _output.WriteLine($"after ClearSweepCache: live sweep again (sweepMs={three.Sweep.ElapsedMs})");
    }
}
