using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins <see cref="LiveWriteAccess"/>, the one decision every Q74 write gate asks: which machine
/// profile may write at all, and what the test-side MCP client may send in a given process.
/// Pure, so every branch is driven here on a runner with no Outlook and no settings file.
/// </summary>
public sealed class LiveWriteAccessTests
{
    [Fact]
    public void OnlyAPortableMachine_MayWrite()
    {
        Assert.False(LiveWriteAccess.RefusesEveryWrite(LiveMachineProfile.Portable));
        Assert.True(LiveWriteAccess.RefusesEveryWrite(LiveMachineProfile.Production));
    }

    [Fact]
    public void AnyProfileNobodyHasThoughtAbout_IsReadOnly()
    {
        // "machineProfile": 7 parses - the enum is numeric underneath - and so would a value added
        // later. Every one of them is a machine nobody decided may write.
        foreach (int value in new[] { -1, 3, 7, int.MaxValue })
        {
            Assert.True(LiveWriteAccess.RefusesEveryWrite((LiveMachineProfile)value), "profile value " + value);
        }

        Assert.Contains("the undefined value 7", LiveWriteAccess.ReadOnlyReason((LiveMachineProfile)7), StringComparison.Ordinal);
    }

    [Fact]
    public void TheReason_NamesTheRule_TheProfile_AndWhereWritingTestsRunInstead()
    {
        string reason = LiveWriteAccess.ReadOnlyReason(LiveMachineProfile.Production);

        Assert.StartsWith(LiveWriteAccess.ReadOnlyMachine, reason, StringComparison.Ordinal);
        Assert.Contains("'Production'", reason, StringComparison.Ordinal);
        Assert.Contains("read-only for live tests ALWAYS", reason, StringComparison.Ordinal);
        Assert.Contains("test mailbox included", reason, StringComparison.Ordinal);
        Assert.Contains("test guest", reason, StringComparison.Ordinal);
        Assert.Contains("Testbed/README.md section 4d", reason, StringComparison.Ordinal);
    }

    [Fact]
    public void TheSettings_AskTheDecision_RatherThanKeepingTheirOwnCopy()
    {
        Assert.True(new LiveTestSettings().RefusesEveryWrite);
        Assert.False(new LiveTestSettings { MachineProfile = LiveMachineProfile.Portable }.RefusesEveryWrite);

        // And the run says so at the top of its log.
        Assert.Contains("writes=NONE (read-only machine)", new LiveTestSettings().Describe(), StringComparison.Ordinal);
        Assert.Contains(
            "writes=hub allowlist",
            new LiveTestSettings { MachineProfile = LiveMachineProfile.Portable }.Describe(),
            StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ the MCP client's posture

    [Theory]
    [InlineData(LiveRunOptIn.Verdict.Missing)]
    [InlineData(LiveRunOptIn.Verdict.OtherMachine)]
    [InlineData(LiveRunOptIn.Verdict.Persisted)]
    public void WithoutTheOptIn_NoLiveRunIsHappening_AndNoSettingsFileIsRead(LiveRunOptIn.Verdict verdict)
    {
        StdioWritePosture posture = LiveWriteAccess.StdioPostureFor(
            verdict, () => throw new InvalidOperationException("a run that did not opt in read the settings file"));

        Assert.Equal(StdioWritePosture.NotALiveRun, posture);
    }

    [Fact]
    public void AnOptedInRun_OnTheWorkstation_IsReadOnly()
    {
        Assert.Equal(
            StdioWritePosture.ReadOnly,
            LiveWriteAccess.StdioPostureFor(LiveRunOptIn.Verdict.Open, () => new LiveTestSettings()));
    }

    [Fact]
    public void AnOptedInRun_OnATestGuest_MayWrite()
    {
        Assert.Equal(
            StdioWritePosture.Writable,
            LiveWriteAccess.StdioPostureFor(
                LiveRunOptIn.Verdict.Open, () => new LiveTestSettings { MachineProfile = LiveMachineProfile.Portable }));
    }

    [Fact]
    public void AnOptedInRun_WhoseSettingsCannotBeRead_IsReadOnly()
    {
        // A posture nobody could establish must not be the one that writes.
        Assert.Equal(
            StdioWritePosture.ReadOnly,
            LiveWriteAccess.StdioPostureFor(
                LiveRunOptIn.Verdict.Open, () => throw new InvalidOperationException("settings not found")));
    }

    [Fact]
    public void TheZeroPosture_IsTheReadOnlyOne()
    {
        // default(StdioWritePosture) is what a field nobody assigned holds.
        Assert.Equal(StdioWritePosture.ReadOnly, default(StdioWritePosture));
    }
}
