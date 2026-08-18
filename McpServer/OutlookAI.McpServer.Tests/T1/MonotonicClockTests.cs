using System.Diagnostics;
using System.Reflection;

using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// <see cref="MonotonicClock"/> exists so that a TTL cannot be extended or expired wholesale
/// by the system clock moving - an NTP correction, a user editing the clock, a VM resuming
/// from a snapshot. Three things in the product measure elapsed time through it: the send
/// confirmation token, the body cache, and the sweep/folder-path/store-detail caches.
/// <para>
/// WHAT CAN BE PROVEN WITHOUT MOVING THE MACHINE'S CLOCK, and it is more than it looks.
/// The defect the class prevents needs a clock jump to observe, and a test may not cause
/// one. But the PROPERTY that rules the jump out is checkable here: the wall clock is read
/// exactly once, into a field that never changes again, and every reading after that is
/// that fixed anchor plus a stopwatch. A reading composed that way cannot move backwards,
/// whatever the system clock does, because nothing in it consults the system clock a second
/// time. So this pins composition and monotonicity, and says plainly which part it cannot
/// reach.
/// </para>
/// <para>
/// It also pins the OTHER half of the class doc - that the value is a real UTC instant on
/// the way in, not process uptime and not zero - because that is what lets it drop into the
/// existing <see cref="DateTime"/>-shaped TTL seams unchanged, and what would make a
/// mistaken use of it for an absolute instant merely drift rather than be nonsense.
/// </para>
/// </summary>
public sealed class MonotonicClockTests
{
    /// <summary>
    /// How far the anchored reading may sit from the system clock. Real drift is the sum of
    /// the clock corrections this process has accepted since its first reading - microseconds
    /// on a healthy machine, and the whole point of the class when it is not - so a minute is
    /// loose enough never to flake and tight enough to prove the value is calendar time.
    /// </summary>
    private static readonly TimeSpan AnchorDriftAllowance = TimeSpan.FromMinutes(1);

    [Fact]
    public void TwoReadingsNeverGoBackwards()
    {
        // The one property every caller depends on, over enough readings that a ragged
        // implementation would show. Equality is allowed: DateTime carries 100 ns ticks and
        // two reads can land inside one.
        DateTime previous = MonotonicClock.UtcNow;
        for (int i = 0; i < 200_000; i++)
        {
            DateTime current = MonotonicClock.UtcNow;
            if (current < previous)
            {
                // Built only on failure: formatting a message 200,000 times would cost more
                // than the loop it is checking.
                Assert.Fail($"reading {i} went backwards: {current:O} after {previous:O}");
            }

            previous = current;
        }
    }

    [Fact]
    public void ReadingsFromSeveralThreadsAtOnce_AreStillNeverBackwardsWithinAThread()
    {
        // Both fields are read without a lock. That is safe only because neither is ever
        // written after initialisation, which is exactly what this would catch if it changed.
        Exception? failure = null;
        Thread[] threads = new Thread[4];
        for (int t = 0; t < threads.Length; t++)
        {
            threads[t] = new Thread(() =>
            {
                try
                {
                    DateTime previous = MonotonicClock.UtcNow;
                    for (int i = 0; i < 50_000; i++)
                    {
                        DateTime current = MonotonicClock.UtcNow;
                        if (current < previous)
                        {
                            Assert.Fail($"went backwards: {current:O} after {previous:O}");
                        }

                        previous = current;
                    }
                }
                catch (Exception ex)
                {
                    Interlocked.CompareExchange(ref failure, ex, null);
                }
            });
            threads[t].Start();
        }

        foreach (Thread thread in threads)
        {
            thread.Join();
        }

        Assert.Null(failure);
    }

    [Fact]
    public void AReadingIsAUtcInstant_NotProcessUptime()
    {
        // If this were the stopwatch alone it would read as year 1. The anchor is what makes
        // it legible in a debugger and droppable into the DateTime-shaped TTL seams.
        DateTime reading = MonotonicClock.UtcNow;

        Assert.Equal(DateTimeKind.Utc, reading.Kind);
        Assert.True(
            (reading - DateTime.UtcNow).Duration() < AnchorDriftAllowance,
            $"anchored reading {reading:O} is not near the system clock {DateTime.UtcNow:O}");
    }

    [Fact]
    public void TheWallClockIsReadOnce_AndEveryReadingAfterThatIsAnchorPlusStopwatch()
    {
        // THE COMPOSITION PIN, and the closest a test can get to the failure the class doc
        // spends two paragraphs on. Read straight off the type: an anchor that is captured
        // once and a stopwatch that supplies every subsequent movement is what makes a
        // backwards clock jump unobservable to the callers. Reached by reflection because
        // the alternative - proving it by moving the machine's clock - is not something a
        // test may do.
        FieldInfo anchor = FieldOrFail("AnchorUtc");
        FieldInfo since = FieldOrFail("SinceAnchor");

        Assert.Equal(typeof(DateTime), anchor.FieldType);
        Assert.Equal(typeof(Stopwatch), since.FieldType);
        Assert.True(anchor.IsInitOnly, "the wall-clock anchor must be readonly - a second reading is the defect");
        Assert.True(since.IsInitOnly, "the stopwatch must be readonly - restarting it would rewind every TTL");

        DateTime anchorBefore = (DateTime)anchor.GetValue(null)!;
        Stopwatch stopwatch = (Stopwatch)since.GetValue(null)!;
        Assert.Equal(DateTimeKind.Utc, anchorBefore.Kind);
        Assert.True(stopwatch.IsRunning, "the stopwatch must still be running, or readings would freeze");

        DateTime reading = MonotonicClock.UtcNow;
        TimeSpan elapsedJustAfter = stopwatch.Elapsed;

        // The reading is anchor + elapsed and nothing else: recomposing it from the two
        // fields a moment later can only be AHEAD, and only by the time between the two
        // lines. A reading that consulted the system clock instead would agree here too, so
        // the assertion below carries the part this cannot: the anchor never moves.
        TimeSpan overshoot = (anchorBefore + elapsedJustAfter) - reading;
        Assert.InRange(overshoot, TimeSpan.Zero, TimeSpan.FromSeconds(1));

        Thread.Sleep(25);

        DateTime anchorAfter = (DateTime)anchor.GetValue(null)!;
        Assert.Equal(anchorBefore, anchorAfter);
        Assert.True(MonotonicClock.UtcNow > reading, "the reading must move even though the anchor does not");
    }

    [Fact]
    public void ADifferenceBetweenTwoReadings_MeasuresTheSameElapsedTimeAStopwatchDoes()
    {
        // What every caller actually uses: the DIFFERENCE. It has to be real elapsed time,
        // or a 120 s send-token TTL is not 120 s. Compared against an independent stopwatch
        // rather than against a fixed expectation, so a loaded machine cannot flake it.
        Stopwatch independent = Stopwatch.StartNew();
        DateTime start = MonotonicClock.UtcNow;
        Thread.Sleep(60);
        DateTime end = MonotonicClock.UtcNow;
        independent.Stop();

        TimeSpan measured = end - start;
        Assert.True(measured > TimeSpan.Zero, "no time passed across a 60 ms sleep");
        Assert.InRange(
            (measured - independent.Elapsed).Duration(),
            TimeSpan.Zero,
            TimeSpan.FromMilliseconds(50));
    }

    private static FieldInfo FieldOrFail(string name)
    {
        FieldInfo? field = typeof(MonotonicClock).GetField(
            name, BindingFlags.NonPublic | BindingFlags.Static);
        Assert.True(
            field != null,
            $"MonotonicClock no longer has a static '{name}' field. If it was renamed, rename it here too; "
            + "if the anchor-plus-stopwatch composition itself changed, this test is the record of why it "
            + "was that way and has to be rewritten deliberately.");
        return field!;
    }
}
