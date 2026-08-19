using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 argument validation for the send operation: rejections that must happen BEFORE
/// any COM work. The gateway is constructed with autostart disabled, so an accidental
/// COM touch fails loudly instead of passing (same discipline as DraftValidationTests).
/// </summary>
public sealed class SendValidationTests : IDisposable
{
    private readonly MailService _service = new(new ComGateway(allowStartingOutlook: false));

    public void Dispose()
    {
        _service.Dispose();
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Send_RequiresId(string? id)
    {
        Assert.Throws<ArgumentException>(() => _service.Send(id!));
    }

    [Fact]
    public void Send_UnknownHitId_RejectedBeforeCom()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() => _service.Send("h424242"));
        Assert.Contains("Unknown id", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Send_NonHexShortId_RejectedBeforeCom()
    {
        Assert.Throws<ArgumentException>(() => _service.Send("not-a-draft-id"));
    }

    /// <summary>
    /// A send that is KILLED mid-flight tells the caller the mail may be in the Outbox.
    /// <para>
    /// This is the one path where the caller knows least and used to be told least. A
    /// deadline expiry lands somewhere between <c>MailItem.Send()</c> executing inside
    /// Outlook and the answer reaching us, so the mail may already have been submitted -
    /// Outlook creates and submits a message in a folder, usually the Outbox - and the
    /// draft's EntryID is gone with it. The neighbouring <c>SendCallFailed</c> branch has
    /// said "The mail MAY be sitting in the Outbox - verify before retrying" since it was
    /// written; the kill path handed back the generic "Outlook did not respond ... the COM
    /// host was restarted" instead, which mentions no mail at all. The confirm token is
    /// already spent by then, so re-confirming after an unknown outcome is exactly how a
    /// duplicate goes out.
    /// </para>
    /// <para>
    /// Pinned as a pure function because reaching the real path needs a real Outlook wedged
    /// in a window of milliseconds.
    /// </para>
    /// </summary>
    [Fact]
    public void AKilledSend_TellsTheCallerTheMailMayBeInTheOutbox()
    {
        string message = MailService.DescribeSendOutcomeUnknown("Outlook did not respond to 'TrySendDraft' within 300000 ms.");

        Assert.Contains("Outbox", message, StringComparison.Ordinal);
        Assert.Contains("UNKNOWN", message, StringComparison.Ordinal);

        // The instruction, not just the diagnosis: an agent that only reads "it failed" will
        // re-send, and that is the outcome this exists to prevent.
        Assert.Contains("Do NOT simply send again", message, StringComparison.Ordinal);
        Assert.Contains("Sent Items", message, StringComparison.Ordinal);

        // And the underlying failure survives, so the report still says what actually broke.
        Assert.Contains("TrySendDraft", message, StringComparison.Ordinal);
    }
}
