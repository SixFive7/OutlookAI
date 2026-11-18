using System.Diagnostics;
using System.Globalization;
using OutlookAI.Core.Audit;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The start protections against a real Outlook (Q175, Q176 (c), Q178 (e), decided by the maintainer
/// 2026-11-18, after the hang of 2026-11-15): a connection waits for Outlook to settle before it opens a
/// mailbox, takes the start lock, and leaves the start record and the start log behind - and this measures
/// what the wait costs.
/// <para>
/// <b>What it cannot show.</b> The guests have no Teams Meeting add-in and no Exchange store, so the
/// deadlock the wait exists to make rarer cannot happen on them (Q181 (d): no reproduction). It shows that
/// the wait works and what it costs; the cold-start cost is measured by <see cref="LiveDisconnectRecoveryTests"/>,
/// the one test that stops and starts Outlook, which prints its connection trace.
/// </para>
/// <para>
/// The test itself reads only: a connection, the profile's name, health; the start log it prints holds no
/// mail content by construction. It does not carry <c>Writes=Nothing</c>, because its collection's fixture
/// sweeps the run's tagged artifacts when it ends - a write - and that trait speaks for the fixture too
/// (<c>T1/ReadOnlyLiveTestTests</c>).
/// </para>
/// </summary>
[Collection(LiveCollections.Lifecycle)]
[Trait("Category", "Live")]
public sealed class LiveOutlookStartTests
{
    private readonly LiveLifecycleFixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveOutlookStartTests(LiveLifecycleFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    [Trait("Requires", "OutlookInstance")]
    public void AConnection_WaitsForOutlookToSettle_AndLeavesTheRecordAndTheLogBehind()
    {
        // Outlook up the way every lifecycle test gets it: through the fixture's own session.
        Assert.True(_fixture.Service.ListAccounts().Accounts.Count > 0, "the fixture's session reached no account");
        OutlookConnectTrace? fixtureConnect = OutlookComSession.LastConnectTrace;
        _output.WriteLine("the fixture's connection: " + (fixtureConnect?.Describe() ?? "none in this process yet"));

        OutlookStartObservation before = OutlookStartObserver.Observe();
        _output.WriteLine("Outlook before: " + before.Describe() + "; add-in stamp: " + (before.AddIn.Source ?? "none")
            + (before.AddIn.StampUtc is DateTime stamped && before.ProcessStartUtc is DateTime started
                ? string.Format(CultureInfo.InvariantCulture, " ({0:0.0} s after Outlook started)", (stamped - started).TotalSeconds)
                : string.Empty));

        // A connection of its own: the start lock, the settle check, the attach.
        Stopwatch clock = Stopwatch.StartNew();
        string profile;
        using (ComGateway gateway = new(allowStartingOutlook: false))
        {
            profile = gateway.Run(session => session.GetProfileName());
        }

        clock.Stop();
        Assert.False(string.IsNullOrEmpty(profile));
        OutlookConnectTrace? trace = OutlookComSession.LastConnectTrace;
        Assert.NotNull(trace);
        _output.WriteLine(string.Format(CultureInfo.InvariantCulture, "this connection: {0} ({1} ms to the profile's name)", trace!.Describe(), clock.ElapsedMilliseconds));

        Assert.False(trace.StartedOutlook, "a connection to a running Outlook claims to have started it");
        Assert.Equal(before.Windows.ProcessId, trace.OutlookProcessId);
        Assert.Contains(trace.SettleOutcome, new[] { OutlookStartPolicy.SettledCode, OutlookStartPolicy.SettledByAgeCode });

        // An Outlook past the starting limit is settled by its age alone, at the cost of one look when it was
        // past it before this connection began; a younger one has had to meet every condition, its age among them.
        if (trace.SettleOutcome == OutlookStartPolicy.SettledCode)
        {
            Assert.True(trace.AgeAtSettleMilliseconds >= OutlookStartPolicy.MinimumAgeMilliseconds,
                "a young Outlook was connected to before it was " + OutlookStartPolicy.MinimumAgeMilliseconds + " ms old");
        }
        else
        {
            Assert.True(trace.AgeAtSettleMilliseconds >= OutlookStartPolicy.StartingLimitMilliseconds,
                "an Outlook was settled by its age before it was " + OutlookStartPolicy.StartingLimitMilliseconds + " ms old");
            if (before.AgeMilliseconds >= OutlookStartPolicy.StartingLimitMilliseconds)
            {
                Assert.True(trace.SettleWaitMilliseconds < 1_000, "an Outlook settled by its age still cost " + trace.SettleWaitMilliseconds + " ms");
            }
        }

        // The record: this connection, finished, on this Outlook.
        OutlookConnectRecord? record = OutlookStartRecord.ReadConnect();
        Assert.NotNull(record);
        Assert.Equal(OutlookStartSteps.Connected, record!.Step);
        Assert.Equal(trace.OutlookProcessId, record.OutlookProcessId);
        Assert.Equal(ComHostIdentity.HostProcessId, record.HostProcessId);

        // The log: this process's lines - written to the test run's own directory - and tokens only.
        Assert.True(AuditLog.IsRedirected, "not redirected - the start log lines below would be the real ones");
        string host = " host=" + ComHostIdentity.HostProcessId.ToString(CultureInfo.InvariantCulture) + " ";
        List<string> mine = File.ReadAllLines(OutlookStartLog.EffectivePath).Where(l => l.Contains(host, StringComparison.Ordinal)).ToList();
        foreach (string line in mine)
        {
            _output.WriteLine("start log: " + line);
            Assert.DoesNotContain("@", line, StringComparison.Ordinal);
            Assert.DoesNotContain("'", line, StringComparison.Ordinal);
        }

        Assert.Contains(mine, l => l.Contains(" connect.end ", StringComparison.Ordinal) && l.Contains("outcome=connected", StringComparison.Ordinal));
        Assert.Contains(mine, l => l.Contains(" settle.end ", StringComparison.Ordinal));

        // Health's start block, from the same look every tool's gate takes.
        HealthOutcome health = _fixture.Service.Health();
        OutlookStartHealthView? start = health.Outlook.Start;
        Assert.NotNull(start);
        _output.WriteLine(string.Format(
            CultureInfo.InvariantCulture,
            "health: state={0} pid={1} startedBy={2} hasWindow={3} everShowedWindow={4} step={5} addInStarted={6}",
            start!.State, start.ProcessId, start.StartedBy, start.HasWindow, start.EverShowedWindow, start.StartStep, start.AddInStarted));
        Assert.Equal("ready", start.State);
        Assert.Equal(before.Windows.ProcessId, start.ProcessId);
        Assert.Null(start.NotRespondingSinceUtc);
        Assert.Equal(OutlookStartLog.EffectivePath, start.LogPath);
    }
}
