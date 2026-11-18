using System.Globalization;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Q175 (b) and Q118 (a), decided 2026-11-18: <b>a connect makes no lifetime pin, and every session joins the
/// per-user session registry while it is attached.</b>
/// <para>
/// Until then every session that found no Explorer pinned Outlook at connect, and the Q118 measurement
/// (<c>OAI-UNINDEXED</c>, Office LTSC 2024 16.0.17932.20996, 2026-10-03, <c>Testbed/guest/Measure-PinWindows.ps1</c>)
/// found such pins left behind - by a session that attached during a cold start and did not start Outlook, and by a
/// starter whose COM host was killed - each keeping a window-less OUTLOOK.EXE running past Application.Quit() and past
/// the user's close of the window they had opened over it. Now the pin is made only when a draft needs one, is shared
/// through the registry, and is closed by the last session out (the Q149 measurement's P cells exercise that; this
/// collection never drafts).
/// </para>
/// <para>
/// This pins the connect half, across two sessions in one process - the fixture's, which brings Outlook up, and a
/// second one - in whatever state Outlook is: the second session makes no pin, the Explorers are the same before,
/// during and after it, and it is in the registry while it is connected and out of it once it is gone. It reads, and
/// creates nothing.
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
    public void AConnect_MakesNoPin_JoinsTheRegistry_AndLeavesTheExplorersAsTheyWere()
    {
        // Outlook up the way every lifecycle test gets it: through the fixture's own session (D17 autostart when none
        // runs). Since Q175 (b) that session made no pin either.
        Assert.True(_fixture.Service.ListAccounts().Accounts.Count > 0, "the fixture's session reached no account");

        ExplorerWindowReading? before = LiveOutlookTestMailer.ReadExplorerWindows();
        _output.WriteLine("before: " + (before?.Describe() ?? "unreadable"));
        Assert.True(before != null, "Outlook's Explorers could not be read before the second session connected");

        bool pinned;
        bool startedOutlook;
        int? during;
        string? key;
        bool registeredWhileConnected;
        using (OutlookComSession second = OutlookComSession.Connect(allowStartingOutlook: false))
        {
            pinned = second.ComposeSurfacePinned;
            startedOutlook = second.StartedOutlook;
            during = LiveOutlookTestMailer.CountExplorers();
            key = second.RegistryKey;
            SessionRegistryState? state = OutlookSessionRegistry.Default.TryRead();
            registeredWhileConnected = key != null && state != null && state.Sessions.Any(s => s.Key == key);
            _output.WriteLine("second session: registry access " + (second.LastRegistryAccess ?? "none") + ", key " + (key ?? "none"));
        }

        ExplorerWindowReading? after = LiveOutlookTestMailer.ReadExplorerWindows();
        SessionRegistryState? afterState = OutlookSessionRegistry.Default.TryRead();
        bool registeredAfter = key != null && afterState != null && afterState.Sessions.Any(s => s.Key == key);
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
            "Q175 (b): a connect must not make a lifetime pin - only a draft that needs one does. A pin made at connect "
            + "by a session that did not start Outlook is the Q118 leak: it kept a window-less OUTLOOK.EXE running past "
            + "Application.Quit().");
        Assert.Equal(before!.Explorers, during);
        Assert.True(after != null, "Outlook's Explorers could not be read after the second session ended");
        Assert.Equal(before.Explorers, after!.Explorers);
        Assert.Equal(before.Hidden, after.Hidden);
        Assert.True(key != null, "Q118 (a): the second session did not join the session registry (its Outlook process was not found)");
        Assert.True(registeredWhileConnected, "Q118 (a): the second session is not in the session registry while it is connected");
        Assert.False(registeredAfter, "Q118 (a): the second session is still in the session registry after it ended");
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
