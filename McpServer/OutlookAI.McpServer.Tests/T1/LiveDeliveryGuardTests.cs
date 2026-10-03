using System.Reflection;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins what the first live run on a test guest (2026-10-03, <c>OutlookAI-Unindexed</c>) found
/// about mail delivery through the guests' local sink, so neither half can quietly come back.
/// <para>
/// <b>The arming defect.</b> The mail-sink probe, the Outbox-drained check and the corpus
/// freshness check were armed by <c>LivePhase1Fixture</c> and nowhere else. Every Phase-1 test
/// needs the search index, so the unindexed guest's filter selects none of them: that run never
/// probed its sink, never checked its Outbox and - because the probe is what switches it on -
/// never let an arrival wait nudge Outlook to fetch. One <c>SendAndReceive</c> per send then
/// raced the submission, and the sink's own log showed the POP3 fetch closing before the message
/// was stored, three times; three tests timed out with their mail sitting in the sink. They are
/// armed now in <see cref="LiveStoreCountTripwire.EnsureBaseline"/>, the funnel every guarded
/// collection's fixture passes.
/// </para>
/// <para>
/// <b>The copy defect.</b> Two of those three waited through PRIVATE copies of the arrival loop
/// that never nudged at all - the sixth and seventh copies of a loop already consolidated once
/// into <see cref="LiveInboxArrival"/>. Every live test that sends mail now waits through it, or
/// through <see cref="LiveInboxArrival.NudgeIfDue"/> where it looks for the arrival through the
/// product instead.
/// </para>
/// <para>
/// Read out of the sources, as <c>LiveEarlyReturnGuardTests</c> does, because the callers are
/// <c>Category=Live</c> and CI can never run them. Nothing here touches Outlook or a mailbox.
/// </para>
/// </summary>
public sealed class LiveDeliveryGuardTests
{
    [Fact]
    public void TheMachineGuards_AreArmedInTheFunnelEveryGuardedCollectionPasses()
    {
        string body = MemberBody("LiveStoreCountTripwire.cs", "public static void EnsureBaseline(");

        int freshness = body.IndexOf("LiveCorpusFreshness.EnsureFresh(settings);", StringComparison.Ordinal);
        int sink = body.IndexOf("LiveMailSink.EnsureReachable(settings);", StringComparison.Ordinal);
        int connect = body.IndexOf("OutlookComSession.Connect(", StringComparison.Ordinal);
        int outbox = body.IndexOf("LiveMailSink.EnsureOutboxDrained(", StringComparison.Ordinal);
        int census = body.IndexOf("Capture(WatchedStores(settings), \"baseline\"", StringComparison.Ordinal);
        int soundness = body.IndexOf("TripwireWatchSoundness.Require(", StringComparison.Ordinal);

        Assert.True(freshness >= 0, "EnsureBaseline no longer arms the corpus freshness check");
        Assert.True(sink >= 0, "EnsureBaseline no longer arms the mail-sink probe (and with it the delivery nudge)");
        Assert.True(outbox >= 0, "EnsureBaseline no longer checks the Outbox is drained before anything sends");

        // The settings soundness gate stays FIRST (the two Tripwire*Tests reach it from CI only
        // because nothing ahead of it touches the machine); the sink is probed before COM; the
        // Outbox is read once COM is up and before the census - before any collection sends.
        Assert.True(soundness >= 0 && soundness < freshness && soundness < sink, "the soundness gate must run first");
        Assert.True(sink < connect, "the sink must be probed before Outlook is connected");
        Assert.True(connect < outbox && outbox < census, "the Outbox must be checked after connecting and before the baseline census");
    }

    [Fact]
    public void NoFixtureArmsTheMachineGuardsOnItsOwn()
    {
        // The funnel is the ONE place. A fixture arming them itself is how one collection came to
        // be the only one that did, and how a filter that skipped it skipped them too.
        foreach (string file in LiveSourceFiles("T2"))
        {
            if (Path.GetFileName(file) is "LiveStoreCountTripwire.cs" or "LiveMailSink.cs" or "LiveCorpusFreshness.cs")
            {
                continue;
            }

            string text = File.ReadAllText(file);
            Assert.False(
                text.Contains("LiveMailSink.EnsureReachable(", StringComparison.Ordinal)
                    || text.Contains("LiveMailSink.EnsureOutboxDrained(", StringComparison.Ordinal)
                    || text.Contains("LiveCorpusFreshness.EnsureFresh(", StringComparison.Ordinal),
                Path.GetFileName(file) + " arms a machine guard itself - arm it in LiveStoreCountTripwire.EnsureBaseline");
        }
    }

    [Fact]
    public void EveryLiveTestThatSendsMail_WaitsThroughTheNudgingHelper()
    {
        List<string> offenders = new List<string>();
        int senders = 0;
        foreach (string file in LiveSourceFiles("T2").Concat(LiveSourceFiles("T3")))
        {
            string name = Path.GetFileName(file);
            if (name == "LiveOutlookTestMailer.cs")
            {
                continue;
            }

            string text = File.ReadAllText(file);
            if (!text.Contains("[Trait(\"Category\", \"Live\")]", StringComparison.Ordinal))
            {
                continue;
            }

            bool sends = text.Contains("SendSelfMail(", StringComparison.Ordinal)
                || text.Contains("CallToolAsync(\"send\"", StringComparison.Ordinal);
            if (!sends)
            {
                continue;
            }

            senders++;
            if (!text.Contains("LiveInboxArrival.", StringComparison.Ordinal))
            {
                offenders.Add(name + " (sends, and never waits through LiveInboxArrival)");
            }

            if (text.Contains("did not arrive in the hub Inbox", StringComparison.Ordinal))
            {
                offenders.Add(name + " (keeps a private copy of the arrival wait)");
            }
        }

        // A scan that found no sender at all has stopped proving anything.
        Assert.True(senders >= 10, "only " + senders + " live test file(s) send mail - the scan has stopped finding them");
        Assert.Empty(offenders);
    }

    [Fact]
    public void TheArrivalNudge_FiresOnTheFirstPollAndOnItsCadence()
    {
        // NudgeIfDue is a no-op unless the machine declares a sink, so what can be pinned without
        // one is the cadence it gates on: the first poll and every NudgeEveryPolls-th after it.
        Assert.Equal(5, LiveInboxArrival.NudgeEveryPolls);
        string body = MemberBody("LiveInboxArrival.cs", "internal static void NudgeIfDue(");
        Assert.Contains("pollsSoFar % NudgeEveryPolls == 0", body, StringComparison.Ordinal);
        Assert.Contains("LiveMailSink.NudgeDelivery();", body, StringComparison.Ordinal);
    }

    private static IEnumerable<string> LiveSourceFiles(string folder)
    {
        return Directory.EnumerateFiles(Path.Combine(TestProjectDir(), folder), "*.cs", SearchOption.TopDirectoryOnly);
    }

    private static string MemberBody(string t2File, string declarationStart)
    {
        string[] lines = File.ReadAllText(Path.Combine(TestProjectDir(), "T2", t2File))
            .Split('\n').Select(l => l.TrimEnd('\r')).ToArray();
        int declaration = Array.FindIndex(lines, l => l.TrimStart().StartsWith(declarationStart, StringComparison.Ordinal));
        Assert.True(declaration >= 0, declarationStart + " was not found in " + t2File + " - this test has stopped proving anything");
        int open = Array.FindIndex(lines, declaration, l => string.Equals(l, "    {", StringComparison.Ordinal));
        int close = open < 0 ? -1 : Array.FindIndex(lines, open + 1, l => string.Equals(l, "    }", StringComparison.Ordinal));
        Assert.True(open > declaration && close > open, "could not find the body of " + declarationStart + " in " + t2File);
        return string.Join("\n", lines[(open + 1)..close]);
    }

    private static string TestProjectDir()
    {
        return typeof(LiveTestSettings).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
    }
}
