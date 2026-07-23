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
}
