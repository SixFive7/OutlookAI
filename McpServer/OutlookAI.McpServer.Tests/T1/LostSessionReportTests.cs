using System.Runtime.InteropServices;
using System.Text.Json;
using ModelContextProtocol.Protocol;
using OutlookAI.ComHost.Client;
using OutlookAI.Core.Com;
using OutlookAI.McpServer.Tools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// A call whose Outlook went away under it is reported to the agent as a COM failure, whatever shape the loss
/// took (decided 2026-11-18). A COM disconnect, a session the exit watcher had already disposed, and one whose
/// COM object it had already separated are one event to <c>ComGateway</c>; the agent gets the same
/// <c>ComFailure</c> error, advice and outcome for each, for a read and a write alike. Until then the two
/// managed shapes reached the agent as "Cannot access a disposed object", with no advice and no outcome - the
/// shape stream E2's Exchange run hit on 2026-11-18.
/// </summary>
public sealed class LostSessionReportTests
{
    private const int RpcDisconnected = unchecked((int)0x80010108);

    public static TheoryData<string, string?> LostShapesAtEachKindOfCall()
    {
        TheoryData<string, string?> cases = new();
        foreach (string shape in new[] { "disposed", "separated" })
        {
            cases.Add(shape, null);
            cases.Add(shape, nameof(IOutlookSession.TryReadItem));
            cases.Add(shape, nameof(IOutlookSession.TrySendDraft));
        }

        return cases;
    }

    [Theory]
    [MemberData(nameof(LostShapesAtEachKindOfCall))]
    public async Task ALostSession_IsReportedExactlyAsACOMDisconnectIs(string shape, string? operation)
    {
        JsonElement lost = await ErrorOf(Lost(shape), operation);
        JsonElement disconnect = await ErrorOf(new COMException("disconnected", RpcDisconnected), operation);

        Assert.Equal("ComFailure", lost.GetProperty("type").GetString());
        Assert.Equal(disconnect.GetProperty("type").GetString(), lost.GetProperty("type").GetString());
        Assert.Equal(disconnect.GetProperty("advice").GetString(), lost.GetProperty("advice").GetString());
        Assert.Equal(OutcomeOf(disconnect), OutcomeOf(lost));
    }

    [Fact]
    public async Task ARead_ThatLostItsSession_IsUnchanged_AndToldToRetry()
    {
        JsonElement error = await ErrorOf(Lost("disposed"), nameof(IOutlookSession.TryReadItem));

        Assert.Equal(MutationOutcome.Unchanged, OutcomeOf(error));
        Assert.Equal(OutlookTools.ComFailureAdvice(nameof(IOutlookSession.TryReadItem)), error.GetProperty("advice").GetString());
        Assert.Equal("ObjectDisposedException 0x80131622", error.GetProperty("message").GetString());
    }

    [Fact]
    public async Task AWrite_ThatLostItsSession_StatesWhatTheInterruptionMayHaveDone_AndIsNotToldToRepeat()
    {
        JsonElement error = await ErrorOf(Lost("separated"), nameof(IOutlookSession.TrySendDraft));

        Assert.Equal(MutationOutcome.ForInterrupted(nameof(IOutlookSession.TrySendDraft)), OutcomeOf(error));
        Assert.NotEqual(MutationOutcome.Unchanged, OutcomeOf(error));
        Assert.Contains("do not simply repeat this call", error.GetProperty("advice").GetString(), StringComparison.Ordinal);
        Assert.Equal("InvalidComObjectException 0x80131527", error.GetProperty("message").GetString());
    }

    [Fact]
    public async Task ALostSession_BeforeAnyCallWasDispatched_StatesNoOutcome()
    {
        JsonElement error = await ErrorOf(Lost("disposed"), null);

        Assert.Null(OutcomeOf(error));
        Assert.Equal(OutlookTools.ComFailureAdvice(null), error.GetProperty("advice").GetString());
    }

    [Fact]
    public void OnlyTheTwoLostSessionShapes_AreCounted()
    {
        Assert.True(OutlookTools.IsLostSession(new ObjectDisposedException("OutlookComSession")));
        Assert.True(OutlookTools.IsLostSession(new InvalidComObjectException("separated")));
        Assert.False(OutlookTools.IsLostSession(new InvalidOperationException("not a lost session")));
        Assert.False(OutlookTools.IsLostSession(new COMException("its own arm", RpcDisconnected)));
    }

    private static Exception Lost(string shape)
    {
        return shape == "disposed"
            ? new ObjectDisposedException("OutlookComSession")
            : new InvalidComObjectException("COM object that has been separated from its underlying RCW cannot be used.");
    }

    /// <summary>The error a tool returns when <paramref name="failure"/> ends it after <paramref name="operation"/> was dispatched.</summary>
    private static async Task<JsonElement> ErrorOf(Exception failure, string? operation)
    {
        CallToolResult result = await OutlookTools.GuardAsync<int>(CancellationToken.None, () =>
        {
            if (operation != null)
            {
                ComHostRequestContext.NoteOperation(operation);
            }

            throw failure;
        });

        Assert.True(result.IsError);
        string text = Assert.IsType<TextContentBlock>(Assert.Single(result.Content)).Text;
        using JsonDocument payload = JsonDocument.Parse(text);
        return payload.RootElement.GetProperty("error").Clone();
    }

    private static string? OutcomeOf(JsonElement error)
    {
        return error.TryGetProperty("outcome", out JsonElement outcome) && outcome.ValueKind == JsonValueKind.String
            ? outcome.GetString()
            : null;
    }
}
