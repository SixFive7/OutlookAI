using System.Globalization;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Q118 (2026-10-03): pins the fact the multi-session measurement of D49's lifetime pin rests on - on
/// <c>OAI-UNINDEXED</c>, Office LTSC 2024 16.0.17932.20996, with <c>Testbed/guest/Measure-PinWindows.ps1</c>
/// driving real server processes, each with its own COM host, against one Outlook.
/// <para>
/// <b>A session that finds an Explorer adds no pin, and so can leave none behind.</b> Against a user's
/// visible Outlook, 11 server sessions - one at a time, three and two at once, two whose COM host was
/// killed - left <c>Explorers.Count</c> at 1 after every one of them, and Application.Quit() and File &gt;
/// Exit then ended OUTLOOK.EXE about as fast as in the control with no session at all. Only a session that
/// finds NO Explorer pins (<c>ComposeSurface.TryPinProcess</c>'s count guard), and only such pins were
/// ever left behind: a starter whose COM host was killed after it pinned, and a session that attached
/// during a cold start, before the starter's pin existed - it pinned too, did not start Outlook, and so
/// kept its pin on a normal exit. A pin left in a window-less Outlook kept OUTLOOK.EXE running past
/// Application.Quit() and past the user's close of the window they had opened over it.
/// </para>
/// <para>
/// This pins the guard, across two sessions in one process - the fixture's, which brings Outlook up,
/// and a second one connected only once an Explorer is there, so that its count guard is what keeps
/// it from pinning. It reads, and creates nothing when its precondition holds.
/// </para>
/// </summary>
[Collection(LiveCollections.Lifecycle)]
[Trait("Category", "Live")]
public sealed class LivePinWindowTests
{
    private readonly LiveLifecycleFixture _fixture;
    private readonly ITestOutputHelper _output;

    public LivePinWindowTests(LiveLifecycleFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    [Trait("Requires", "OutlookInstance")]
    public void ASecondSession_FindingAnExplorer_AddsNoPin_AndLeavesNoExplorerBehind()
    {
        // Outlook up the way every lifecycle test gets it: through the fixture's own session (D17
        // autostart when none runs - and a session that starts Outlook pins it, D49).
        Assert.True(_fixture.Service.ListAccounts().Accounts.Count > 0, "the fixture's session reached no account");

        ExplorerWindowReading? before = LiveOutlookTestMailer.ReadExplorerWindows();
        _output.WriteLine("before: " + (before?.Describe() ?? "unreadable"));
        Assert.True(
            before is { Explorers: > 0 },
            "precondition: Outlook held no Explorer (" + (before?.Describe() ?? "unreadable") + ") after the fixture's "
            + "session reached it - neither that session's lifetime pin nor a user's window. A session connected now "
            + "WOULD pin, and since it did not start Outlook its pin would outlive it (the Q118 leak), so this test "
            + "does not connect one.");

        bool pinned;
        bool startedOutlook;
        int? during;
        using (OutlookComSession second = OutlookComSession.Connect(allowStartingOutlook: false))
        {
            pinned = second.ComposeSurfacePinned;
            startedOutlook = second.StartedOutlook;
            during = LiveOutlookTestMailer.CountExplorers();
        }

        ExplorerWindowReading? after = LiveOutlookTestMailer.ReadExplorerWindows();
        _output.WriteLine(string.Format(
            CultureInfo.InvariantCulture,
            "second session: pinned={0} startedOutlook={1}; Explorers while it was connected: {2}",
            pinned,
            startedOutlook,
            during?.ToString(CultureInfo.InvariantCulture) ?? "unreadable"));
        _output.WriteLine("after it ended: " + (after?.Describe() ?? "unreadable"));

        Assert.False(startedOutlook, "the second session claims to have started an Outlook that was already running");
        Assert.False(
            pinned,
            "D49's count guard: a session that finds an Explorer must not add a lifetime pin - it did not start Outlook, "
            + "so OutlookComSession.Dispose would leave that pin behind, and a pin left in a window-less Outlook keeps "
            + "OUTLOOK.EXE running past Application.Quit() (Q118).");
        Assert.Equal(before!.Explorers, during);
        Assert.True(after != null, "Outlook's Explorers could not be read after the second session ended");
        Assert.Equal(before.Explorers, after!.Explorers);
        Assert.Equal(before.Hidden, after.Hidden);
    }
}

/// <summary>
/// One reading of the running Outlook's Explorers by <see cref="LiveOutlookTestMailer.ReadExplorerWindows"/>:
/// how many, how many with a visible / hidden / unreadable window, and what <c>ActiveExplorer()</c> returned.
/// </summary>
public sealed record ExplorerWindowReading(int Explorers, int Visible, int Hidden, int Unknown, bool ActiveIsNull, bool? ActiveVisible)
{
    public string Describe()
    {
        string active = ActiveIsNull ? "null" : ActiveVisible switch
        {
            true => "a visible Explorer",
            false => "a HIDDEN Explorer",
            _ => "an Explorer whose window could not be read",
        };
        return string.Format(
            CultureInfo.InvariantCulture,
            "Explorers {0} (visible {1}, hidden {2}, unreadable {3}); ActiveExplorer(): {4}",
            Explorers,
            Visible,
            Hidden,
            Unknown,
            active);
    }
}
