using System.Reflection;
using System.Text.Json.Nodes;
using OutlookAI.Core.IndexSearch;
using OutlookAI.McpServer.Tests.T2;
using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins Q130 (b), decided by the maintainer 2026-10-03 (<c>Docs/overnight-review-2026-10-03.md</c>):
/// the live tier's one corpus-dependent timing test asks for a window anchored to the DATA, not to the
/// clock - <see cref="LiveDataWindow"/> - so on the indexed guest it times a date predicate over the same
/// Corpus A rows on every run, whatever the clock says, and on a machine with no corpus it asks what it
/// always asked. Pinned twice: the window's semantics, and from the compiled IL, that the test takes its
/// window from there and computes none of its own. Pure: no Outlook, no index, no settings file.
/// </summary>
public sealed class LiveDataWindowTests
{
    private static readonly DateTime Anchor = new(2026, 10, 3, 0, 0, 0, DateTimeKind.Utc);

    private static CorpusSettings Corpus(string anchorUtc = "2026-10-03T00:00:00Z") => new()
    {
        StoreDisplayName = "Corpus A",
        ManifestPath = @"C:\nowhere\corpus-vm-indexed.jsonl",
        CorpusId = "vm-indexed",
        Seed = 7777,
        AnchorUtc = anchorUtc,
        ItemCount = 160000,
        WindowDays = new List<int> { 30, 60 },
    };

    // ================================================================ the window

    [Fact]
    public void WithACorpus_TheWindowIsTheDaysBeforeItsAnchor_ClosedAtTheAnchor()
    {
        LiveDateWindow window = LiveDataWindow.Before(Corpus(), 30, new DateTime(2026, 10, 3, 17, 42, 0, DateTimeKind.Utc));

        Assert.Equal(Anchor.AddDays(-30), window.OnOrAfterUtc);
        Assert.Equal(Anchor, window.BeforeUtc);
        Assert.Equal(DateTimeKind.Utc, window.OnOrAfterUtc.Kind);
        Assert.Contains("vm-indexed", window.Basis, StringComparison.Ordinal);
        Assert.Contains("not the clock's", window.Basis, StringComparison.Ordinal);
    }

    [Fact]
    public void WithACorpus_TheClockDoesNotMoveTheWindow()
    {
        // The whole point: the same window the day the corpus was built, on a frozen guest, and a year
        // after the corpus's 30-day window emptied on a real clock (2026-11-01 23:59:16Z).
        DateTime[] clocks =
        {
            new(2026, 10, 3, 0, 1, 0, DateTimeKind.Utc),
            new(2026, 11, 2, 0, 0, 0, DateTimeKind.Utc),
            new(2027, 11, 2, 0, 0, 0, DateTimeKind.Utc),
            new(2026, 8, 1, 0, 0, 0, DateTimeKind.Local),
        };

        LiveDateWindow first = LiveDataWindow.Before(Corpus(), 30, clocks[0]);
        foreach (DateTime clock in clocks)
        {
            LiveDateWindow window = LiveDataWindow.Before(Corpus(), 30, clock);
            Assert.Equal(first.OnOrAfterUtc, window.OnOrAfterUtc);
            Assert.Equal(first.BeforeUtc, window.BeforeUtc);
        }
    }

    [Fact]
    public void WithADateOnlyAnchor_TheWindowIsTheSame()
    {
        // corpus-plan writes 'yyyy-MM-dd' as readily as the full instant; both mean midnight UTC.
        LiveDateWindow window = LiveDataWindow.Before(Corpus("2026-10-03"), 30, DateTime.UtcNow);
        Assert.Equal(Anchor.AddDays(-30), window.OnOrAfterUtc);
        Assert.Equal(Anchor, window.BeforeUtc);
    }

    [Fact]
    public void WithoutACorpus_TheWindowIsTheLastDaysOfTheClock_OpenEnded()
    {
        // A machine that declares no corpus - the maintainer's workstation - keeps the window it had.
        var now = new DateTime(2026, 10, 3, 12, 0, 0, DateTimeKind.Utc);
        LiveDateWindow window = LiveDataWindow.Before(null, 30, now);

        Assert.Equal(now.AddDays(-30), window.OnOrAfterUtc);
        Assert.Null(window.BeforeUtc);
        Assert.Contains("declares no corpus", window.Basis, StringComparison.Ordinal);
    }

    [Fact]
    public void ACorpusWithNoAnchor_IsTreatedAsNoCorpus()
    {
        // The loader refuses a partial corpus block before any test runs; this only says what the
        // window does with one, so the helper never has to guess an anchor.
        var now = new DateTime(2026, 10, 3, 12, 0, 0, DateTimeKind.Utc);
        LiveDateWindow window = LiveDataWindow.Before(Corpus(anchorUtc: " "), 30, now);
        Assert.Equal(now.AddDays(-30), window.OnOrAfterUtc);
        Assert.Null(window.BeforeUtc);
    }

    [Fact]
    public void AnAnchorThatIsNotAnInstant_IsRefused_NeverReplacedByTheClock()
    {
        InvalidOperationException refused = Assert.Throws<InvalidOperationException>(
            () => LiveDataWindow.Before(Corpus("last Tuesday"), 30, DateTime.UtcNow));
        Assert.Contains("last Tuesday", refused.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-30)]
    public void AWindowIsAtLeastOneDayWide(int days)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => LiveDataWindow.Before(Corpus(), days, DateTime.UtcNow));
    }

    [Fact]
    public void OnTheIndexedGuest_TheWindowSelectsCorpusAsRecordedThirtyDays_ForEver()
    {
        // The indexed guest's corpus as testbed.json declares it: the window the test asks for selects
        // exactly the items 'corpus-plan' printed for 30 days before the build (expectedPlan,
        // selectedByWindowDays) - big-store rows, on every run, frozen clock or real one.
        JsonObject root = JsonNode.Parse(File.ReadAllText(Path.Combine(RepoRoot(), "Testbed", "testbed.json")))!.AsObject();
        JsonObject declared = root["liveTestSettings"]!["OutlookAI-Indexed"]!["corpus"]!.AsObject();
        var corpus = new CorpusSettings
        {
            StoreDisplayName = declared["storeDisplayName"]!.GetValue<string>(),
            ManifestPath = declared["manifestPath"]!.GetValue<string>(),
            CorpusId = declared["corpusId"]!.GetValue<string>(),
            Seed = declared["seed"]!.GetValue<long>(),
            AnchorUtc = declared["anchorUtc"]!.GetValue<string>(),
            ItemCount = declared["itemCount"]!.GetValue<int>(),
        };

        JsonObject record = root["corpusIdConvention"]!["assigned"]!.AsArray()
            .Select(n => n!.AsObject())
            .Single(o => o["corpusId"]!.GetValue<string>() == corpus.CorpusId);
        int recorded = record["expectedPlan"]!["selectedByWindowDays"]!["30"]!.GetValue<int>();

        LiveDateWindow window = LiveDataWindow.Before(corpus, 30, new DateTime(2027, 1, 1, 0, 0, 0, DateTimeKind.Utc));
        DateTime anchor = CorpusManifest.ParseUtc(corpus.AnchorUtc)!.Value;
        var plan = new CorpusPlan(new CorpusPlanOptions(corpus.CorpusId, corpus.Seed, anchor));
        int selected = 0;
        for (int ordinal = 1; ordinal <= corpus.ItemCount; ordinal++)
        {
            CorpusItemSpec spec = plan.Describe(ordinal);
            if (!spec.IsUndated && spec.ReceivedUtc >= window.OnOrAfterUtc && spec.ReceivedUtc < window.BeforeUtc!.Value)
            {
                selected++;
            }
        }

        Assert.Equal(recorded, selected);
        Assert.True(selected > 10_000, $"the window selects {selected} of Corpus A's items - not a big-store predicate");
    }

    // ================================================================ the test that uses it (IL)

    [Fact]
    public void TheDateRangeTest_TakesItsWindowFromTheData_AndComputesNoneOfItsOwn()
    {
        MethodInfo test = typeof(LiveIndexSearchTests).GetMethod(nameof(LiveIndexSearchTests.ProbeParity_DateRangeQuery_HitsUnder2s))!;
        IReadOnlyList<MethodBase> calls = IlReader.Read(test).Where(i => i.Method != null).Select(i => i.Method!).ToList();

        Assert.Contains(calls, m => m.DeclaringType == typeof(LiveDataWindow) && m.Name == nameof(LiveDataWindow.Before));
        Assert.Contains(calls, m => m.DeclaringType == typeof(LiveTestSettings) && m.Name == "get_" + nameof(LiveTestSettings.Corpus));
        Assert.Contains(calls, m => m.DeclaringType == typeof(IndexQuery) && m.Name == "set_" + nameof(IndexQuery.ReceivedOnOrAfterUtc));
        Assert.Contains(calls, m => m.DeclaringType == typeof(IndexQuery) && m.Name == "set_" + nameof(IndexQuery.ReceivedBeforeUtc));

        // The window it used to compute inline - DateTime.UtcNow.AddDays(-30) - is gone, and with it any
        // other way of moving the window by the clock.
        Assert.DoesNotContain(calls, m => m.DeclaringType == typeof(DateTime) && m.Name is "AddDays" or "AddHours" or "Subtract");
    }

    private static string RepoRoot()
    {
        string testProjectDir =
            typeof(LiveTestSettings).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");

        // <repo>/McpServer/OutlookAI.McpServer.Tests/ -> <repo>
        return Path.GetFullPath(Path.Combine(testProjectDir, "..", ".."));
    }
}
