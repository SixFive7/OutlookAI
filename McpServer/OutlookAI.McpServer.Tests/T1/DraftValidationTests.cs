using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 argument validation for the draft operations: every rejection must happen BEFORE
/// any COM work. The gateway is constructed with autostart disabled, so an accidental
/// COM touch on a machine without a running Outlook fails loudly instead of passing.
/// </summary>
public sealed class DraftValidationTests : IDisposable
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
    public void NewDraft_RequiresAccount(string? account)
    {
        Assert.Throws<ArgumentException>(() =>
            _service.NewDraft(account!, "a@b.example", null, "subject", "body", display: false));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" ; , ")]
    public void NewDraft_RequiresAtLeastOneToRecipient(string? to)
    {
        Assert.Throws<ArgumentException>(() =>
            _service.NewDraft("hub@example.com", to, null, "subject", "body", display: false));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("  ")]
    public void NewDraft_RequiresSubject(string? subject)
    {
        Assert.Throws<ArgumentException>(() =>
            _service.NewDraft("hub@example.com", "a@b.example", null, subject, "body", display: false));
    }

    [Fact]
    public void NewDraft_RejectsOverlongSubject()
    {
        string subject = new('x', 256);
        Assert.Throws<ArgumentException>(() =>
            _service.NewDraft("hub@example.com", "a@b.example", null, subject, "body", display: false));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("  ")]
    public void NewDraft_RequiresBody(string? body)
    {
        Assert.Throws<ArgumentException>(() =>
            _service.NewDraft("hub@example.com", "a@b.example", null, "subject", body, display: false));
    }

    [Theory]
    [InlineData(null)]
    [InlineData(" ")]
    public void ReplyDraft_RequiresId(string? id)
    {
        Assert.Throws<ArgumentException>(() => _service.ReplyDraft(id!, "body", replyAll: false, display: false));
    }

    [Theory]
    [InlineData(null)]
    [InlineData(" ")]
    public void ReplyDraft_RequiresBody(string? body)
    {
        Assert.Throws<ArgumentException>(() => _service.ReplyDraft("h1", body, replyAll: false, display: false));
    }

    [Fact]
    public void ReplyDraft_UnknownHitId_ThrowsInstructiveArgumentError()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            _service.ReplyDraft("h424242", "body", replyAll: true, display: false));
        Assert.Contains("Unknown id", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null)]
    [InlineData(" ; ")]
    public void ForwardDraft_RequiresToRecipients(string? to)
    {
        Assert.Throws<ArgumentException>(() => _service.ForwardDraft("h1", "body", to, display: false));
    }

    [Fact]
    public void ForwardDraft_RequiresBody()
    {
        Assert.Throws<ArgumentException>(() => _service.ForwardDraft("h1", " ", "a@b.example", display: false));
    }

    [Fact]
    public void ForwardDraft_UnknownHitId_ThrowsInstructiveArgumentError()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            _service.ForwardDraft("h424242", "body", "a@b.example", display: false));
        Assert.Contains("Unknown id", ex.Message, StringComparison.Ordinal);
    }
}
