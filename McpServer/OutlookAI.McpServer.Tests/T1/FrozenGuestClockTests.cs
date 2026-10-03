using System.Reflection;
using System.Text.Json.Nodes;
using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tests.T2;
using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins Q130 (a), decided by the maintainer 2026-10-03 (<c>Docs/overnight-review-2026-10-03.md</c>): the two
/// Outlook test guests run with Hyper-V time synchronisation off, and every live run restores the frozen
/// checkpoint <c>Testbed/testbed.json</c> records under <c>frozenClocks</c>, so it starts at the same instant.
/// What makes that safe is that EVERY date check the tier makes holds at every instant a run can reach -
/// from the frozen instant to <c>checksHoldForMinutes</c> after it - and this proves it from the committed
/// records, with the same code the tier itself runs: the frontier margin
/// (<see cref="LiveHubPopulationFreshness"/>), Corpus A's declared windows (<see cref="CorpusFreshness"/>)
/// and the unindexed guest's seven-day reach (<see cref="MailService.EmptyIndexSweepWindow"/>). A new frozen
/// checkpoint recorded with an instant too far from its data fails here, before any guest runs it.
/// Pure: the committed testbed.json, no Outlook, no VM.
/// </summary>
public sealed class FrozenGuestClockTests
{
    private static readonly string[] OutlookGuests = { "OutlookAI-Indexed", "OutlookAI-Unindexed" };

    public static TheoryData<string> Guests => new() { "OutlookAI-Indexed", "OutlookAI-Unindexed" };

    // ================================================================ the records

    [Fact]
    public void ExactlyTheTwoOutlookGuestsAreFrozen_NeverTheBuildVm_NorTheExchangeVm()
    {
        // The build VM's runner requires the host's clock within 2 s (Invoke-TestsOnBuildVm.ps1), and
        // Microsoft 365 sign-in on OutlookAI-Exchange needs real time.
        string[] frozen = FrozenClocks().Select(p => p.Key).Where(k => !k.StartsWith('_')).OrderBy(k => k, StringComparer.Ordinal).ToArray();
        Assert.Equal(OutlookGuests.OrderBy(k => k, StringComparer.Ordinal), frozen);
    }

    [Theory]
    [MemberData(nameof(Guests))]
    public void TheRecord_NamesItsCheckpoints_AndItsTwoInstantsAgree(string guest)
    {
        JsonObject r = Record(guest);
        string checkpoint = r["checkpoint"]!.GetValue<string>();
        string parent = r["parentCheckpoint"]!.GetValue<string>();
        Assert.False(string.IsNullOrWhiteSpace(checkpoint));
        Assert.False(string.IsNullOrWhiteSpace(parent));
        Assert.NotEqual(parent, checkpoint);

        // Scripts read the Unix count (ConvertFrom-Json loses an ISO string's zone - Testbed/README.md 5b);
        // people read the text. They must be the same instant.
        DateTime text = ParseUtc(r["frozenUtc"]!.GetValue<string>());
        DateTime unix = DateTime.UnixEpoch.AddSeconds(r["frozenUnix"]!.GetValue<long>());
        Assert.Equal(text, unix);

        int start = r["suiteStartWithinMinutes"]!.GetValue<int>();
        int hold = r["checksHoldForMinutes"]!.GetValue<int>();
        Assert.InRange(start, 1, hold - 1);
    }

    // ================================================================ every check holds for every run

    [Theory]
    [MemberData(nameof(Guests))]
    public void TheFrontierCheck_HoldsFromTheFrozenInstant_ToTheEndOfAnyRun(string guest)
    {
        JsonObject r = Record(guest);
        (DateTime frozen, DateTime last) = Span(r);
        HubPopulationFact fact = LiveHubPopulationFreshness.Read(HubManifestHeader(guest, r));
        Assert.Equal(ParseUtc(r["hubPopulation"]!["anchorUtc"]!.GetValue<string>()), fact.AnchorUtc);

        // The guest's offset is the frozen date's - summer time, +2 h - and it is the same at the end of
        // any run: no run can cross 2026-10-25, so the margin never shrinks to the winter 55 minutes.
        TimeZoneInfo zone = TimeZoneInfo.FindSystemTimeZoneById(r["timeZone"]!.GetValue<string>());
        TimeSpan offset = zone.GetUtcOffset(frozen);
        Assert.Equal(offset, zone.GetUtcOffset(last));
        Assert.Equal(TimeSpan.FromHours(2), offset);

        // A restore comes up seconds before the instant; the guard allows two minutes either side of that.
        foreach (DateTime clock in new[] { frozen.AddMinutes(-2), frozen, last })
        {
            (bool proceed, string message) = LiveHubPopulationFreshness.Decide(fact, clock, offset);
            Assert.True(proceed, $"{guest} at {CorpusManifest.FormatUtc(clock)}: {message}");
        }
    }

    [Fact]
    public void OnTheIndexedGuest_CorpusAIsFreshForEveryRun_WhereARealClockWouldHaveStoppedTheTier()
    {
        JsonObject r = Record("OutlookAI-Indexed");
        (DateTime frozen, DateTime last) = Span(r);
        JsonObject corpus = LiveSettings("OutlookAI-Indexed")["corpus"]!.AsObject();
        int[] windows = corpus["windowDays"]!.AsArray().Select(n => n!.GetValue<int>()).ToArray();
        DateTime anchor = ParseUtc(corpus["anchorUtc"]!.GetValue<string>());
        var plan = new CorpusPlan(new CorpusPlanOptions(corpus["corpusId"]!.GetValue<string>(), corpus["seed"]!.GetValue<long>(), anchor));
        int count = corpus["itemCount"]!.GetValue<int>();

        Assert.Equal(CorpusFreshnessVerdict.Fresh, CorpusFreshness.Evaluate(plan, count, TimeSpan.Zero, frozen, windows).Verdict);
        Assert.Equal(CorpusFreshnessVerdict.Fresh, CorpusFreshness.Evaluate(plan, count, TimeSpan.Zero, last, windows).Verdict);

        // What the frozen clock spares this guest: on a real clock the tier refuses from 2026-11-02.
        Assert.Equal(
            CorpusFreshnessVerdict.WindowsEmptied,
            CorpusFreshness.Evaluate(plan, count, TimeSpan.Zero, new DateTime(2026, 11, 2, 0, 0, 0, DateTimeKind.Utc), windows).Verdict);
    }

    [Fact]
    public void OnTheUnindexedGuest_NoHubItemLeavesTheSevenDayReach_DuringAnyRun()
    {
        // An unindexed store's search reaches back MailService.EmptyIndexSweepWindow from the clock, so the
        // hub-reading tests there see exactly the hub items younger than that. The frozen clock keeps that
        // set the one the hub was built with, for the whole span a run can reach.
        JsonObject r = Record("OutlookAI-Unindexed");
        DateTime last = Span(r).Last;
        JsonObject hub = r["hubPopulation"]!.AsObject();
        DateTime anchor = ParseUtc(hub["anchorUtc"]!.GetValue<string>());
        var options = HubOptions(hub, anchor);
        var plan = new CorpusPlan(options);

        List<DateTime> dated = new();
        for (int ordinal = 1; ordinal <= plan.FixedItemCount!.Value; ordinal++)
        {
            CorpusItemSpec spec = plan.Describe(ordinal);
            if (!spec.IsUndated)
            {
                dated.Add(spec.ReceivedUtc);
            }
        }

        int atTheBuild = dated.Count(d => d > anchor - MailService.EmptyIndexSweepWindow);
        int atTheEndOfAnyRun = dated.Count(d => d > last - MailService.EmptyIndexSweepWindow);
        Assert.True(atTheBuild > 0, "the hub has no item inside the seven-day reach at all");
        Assert.Equal(atTheBuild, atTheEndOfAnyRun);
    }

    // ================================================================ helpers

    private static (DateTime Frozen, DateTime Last) Span(JsonObject record)
    {
        DateTime frozen = ParseUtc(record["frozenUtc"]!.GetValue<string>());
        return (frozen, frozen.AddMinutes(record["checksHoldForMinutes"]!.GetValue<int>()));
    }

    private static CorpusPlanOptions HubOptions(JsonObject hub, DateTime anchor)
    {
        var options = new CorpusPlanOptions(hub["corpusId"]!.GetValue<string>(), hub["seed"]!.GetValue<long>(), anchor)
        {
            Population = CorpusPopulationKind.Hub,
            Owner = CorpusMailboxOwner.ForStore(hub["store"]!.GetValue<string>()),
            IncludeUndatedContacts = hub["includeUndatedContacts"]!.GetValue<bool>(),
            IncludeAllKinds = hub["includeAllKinds"]!.GetValue<bool>(),
        };

        // The recorded shape key is the guest manifest's; the generator must reproduce it from the record.
        Assert.Equal(hub["shapeKey"]!.GetValue<string>(), options.ShapeKey);
        return options;
    }

    /// <summary>The hub manifest's header line as the frozen checkpoint holds it, from the record.</summary>
    private static string[] HubManifestHeader(string guest, JsonObject record)
    {
        JsonObject hub = record["hubPopulation"]!.AsObject();
        DateTime anchor = ParseUtc(hub["anchorUtc"]!.GetValue<string>());
        CorpusPlanOptions options = HubOptions(hub, anchor);
        Assert.Equal(LiveSettings(guest)["testHubStoreDisplayName"]!.GetValue<string>(), hub["store"]!.GetValue<string>());
        var header = new CorpusManifestHeader(
            CorpusManifest.CurrentVersion, options.CorpusId, options.Seed, CorpusManifest.FormatUtc(anchor),
            options.ShapeKey, hub["store"]!.GetValue<string>(), null, "PropertyAccessorDates", "PostAsNote");
        return new[] { CorpusManifest.RenderLine(header) };
    }

    private static DateTime ParseUtc(string text)
        => CorpusManifest.ParseUtc(text) ?? throw new InvalidOperationException($"'{text}' is not a UTC instant");

    private static JsonObject Record(string guest) => FrozenClocks()[guest]!.AsObject();

    private static JsonObject FrozenClocks() => Testbed()["frozenClocks"]!.AsObject();

    private static JsonObject LiveSettings(string guest) => Testbed()["liveTestSettings"]![guest]!.AsObject();

    private static JsonObject Testbed()
        => JsonNode.Parse(File.ReadAllText(Path.Combine(RepoRoot(), "Testbed", "testbed.json")))!.AsObject();

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
