using OutlookAI.Core.Com;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins how the index-hit locator matches an item to its index row by time (the coordinator's item 7, 2026-11-18):
/// <c>LiveMailServiceTests.Thread_IndexPath_AndComFallback</c> could not open its one index member on the Exchange VM,
/// even on a rebuilt index. Measured there the same day (Testbed/guest/Measure-FolderIdResolution.ps1): the test
/// account's newest index rows are holiday APPOINTMENTS, whose <c>System.Message.DateReceived</c> is their creation
/// time - as Q98 (f) measured for appointments and tasks in a PST - while the locator compared the item's
/// <c>ReceivedTime</c>, which an appointment does not have; and one subject comes back two to five times, the
/// occurrences created tenths of a second apart. So an item without a ReceivedTime is matched on its CreationTime,
/// and the NEAREST candidate within the tolerance wins, not the first.
/// </summary>
public sealed class HitLocatorTimeMatchTests
{
    private static readonly DateTime IndexUtc = new(2026, 11, 17, 18, 0, 0, 500, DateTimeKind.Utc);

    [Fact]
    public void AnItemWithoutAReceivedTime_IsMatchedOnItsCreationTime()
    {
        DateTime received = new(2026, 1, 1, 9, 0, 0);
        DateTime created = new(2026, 1, 2, 9, 0, 0);

        Assert.Equal(received, HitLocator.IndexTimeOf(received, created));
        Assert.Equal(created, HitLocator.IndexTimeOf(null, created));
        Assert.Null(HitLocator.IndexTimeOf(null, null));
    }

    [Fact]
    public void SecondsApart_IsTheNearerOfTheLocalAndTheUtcReading()
    {
        // The zone is held still (ComDateValueTests.OddFixedZone, UTC+07:13), so the answer is the same on any machine.
        TimeZoneInfo zone = ComDateValueTests.OddFixedZone;
        DateTime wallTime = new DateTime(2026, 11, 18, 1, 13, 2, 500, DateTimeKind.Unspecified);
        Assert.Equal(2.0, HitLocator.SecondsApart(wallTime, IndexUtc, zone)!.Value, 3);

        // A time Outlook handed over already in UTC: the raw reading is the near one.
        DateTime rawUtc = DateTime.SpecifyKind(IndexUtc.AddSeconds(-0.25), DateTimeKind.Unspecified);
        Assert.Equal(0.25, HitLocator.SecondsApart(rawUtc, IndexUtc, zone)!.Value, 3);

        Assert.Null(HitLocator.SecondsApart(null, IndexUtc, zone));
        Assert.Throws<ArgumentNullException>(() => HitLocator.SecondsApart(rawUtc, IndexUtc, null!));
    }

    [Fact]
    public void TheNearestWithinTheTolerance_Wins_NotTheFirst()
    {
        // The measured shape: occurrences of one holiday created 0.8 s apart; the row's own one is 0.0 s away.
        ClosestByTime<string> nearest = new(5);
        Assert.False(nearest.Offer("one year", 0.8));
        Assert.False(nearest.Offer("another year", 0.8 + 0.8));
        Assert.True(nearest.Offer("the row's own", 0.0));
        Assert.Equal("the row's own", nearest.Best);
        Assert.Equal(0.0, nearest.BestSecondsApart);
    }

    [Fact]
    public void ACandidateOutsideTheTolerance_OrWithoutATime_IsNeverChosen_AndATieKeepsTheFirst()
    {
        ClosestByTime<string> nearest = new(5);
        Assert.False(nearest.Offer("too far", 5.01));
        Assert.False(nearest.Offer("no time", null));
        Assert.Null(nearest.Best);

        Assert.False(nearest.Offer("first", 1.0));
        Assert.False(nearest.Offer("second", 1.0));
        Assert.Equal("first", nearest.Best);

        // Exactly the tolerance still counts, as ReceivedTimeMatches always did.
        ClosestByTime<string> edge = new(5);
        edge.Offer("edge", 5.0);
        Assert.Equal("edge", edge.Best);
    }

    [Fact]
    public void AnExactMatch_EndsTheSearch()
    {
        ClosestByTime<string> nearest = new(120);
        Assert.True(nearest.Offer("exact", HitLocator.ExactMatchSeconds));
        Assert.False(new ClosestByTime<string>(120).Offer("near", HitLocator.ExactMatchSeconds + 0.01));
        Assert.Throws<ArgumentOutOfRangeException>(() => new ClosestByTime<string>(-1));
    }

    [Fact]
    [Trait("TimeZone", "Machine")]
    public void ReceivedTimeMatches_KeepsItsContract_ThroughSecondsApart()
    {
        DateTime local = DateTime.SpecifyKind(IndexUtc.ToLocalTime(), DateTimeKind.Unspecified);
        Assert.True(OutlookComSession.ReceivedTimeMatches(local.AddSeconds(4), IndexUtc, 5));
        Assert.False(OutlookComSession.ReceivedTimeMatches(local.AddSeconds(6), IndexUtc, 5));
        Assert.False(OutlookComSession.ReceivedTimeMatches(null, IndexUtc, 5));
    }
}
