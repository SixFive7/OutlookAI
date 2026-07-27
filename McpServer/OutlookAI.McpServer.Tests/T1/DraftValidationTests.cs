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

    // ---------------------------------------------------------- batch A (A3/A4)

    [Theory]
    [InlineData("low", 0)]
    [InlineData("normal", 1)]
    [InlineData("high", 2)]
    [InlineData("HIGH", 2)]
    [InlineData("  High  ", 2)]
    public void ParseImportance_MapsTheThreeWireValues(string importance, int expected)
    {
        Assert.Equal(expected, MailService.ParseImportance(importance));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void ParseImportance_BlankLeavesOutlookDefault(string? importance)
    {
        Assert.Null(MailService.ParseImportance(importance));
    }

    [Theory]
    [InlineData("urgent")]
    [InlineData("2")]
    [InlineData("medium")]
    public void ParseImportance_RejectsUnknownValue_ListingTheAllowedOnes(string importance)
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() => MailService.ParseImportance(importance));
        Assert.Contains("low", ex.Message, StringComparison.Ordinal);
        Assert.Contains("normal", ex.Message, StringComparison.Ordinal);
        Assert.Contains("high", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void NewDraft_RejectsUnknownImportance_BeforeAnyComWork()
    {
        Assert.Throws<ArgumentException>(() => _service.NewDraft(
            "hub@example.com", "a@b.example", null, "subject", "body", display: false, signature: null,
            bcc: null, importance: "urgent"));
    }

    [Theory]
    [InlineData("reply")]
    [InlineData("replyall")]
    [InlineData("forward")]
    public void DerivedDrafts_RejectUnknownImportance_BeforeAnyComWork(string kind)
    {
        Assert.Throws<ArgumentException>(() => CallDerived(kind, subject: null, importance: "urgent"));
    }

    [Theory]
    [InlineData("reply")]
    [InlineData("replyall")]
    [InlineData("forward")]
    public void DerivedDrafts_RejectOverlongSubjectOverride_BeforeAnyComWork(string kind)
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            CallDerived(kind, subject: new string('x', 256), importance: null));
        Assert.Contains("subject", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("reply")]
    [InlineData("replyall")]
    [InlineData("forward")]
    public void DerivedDrafts_AcceptA255CharSubjectOverride_AndFailOnlyOnTheUnknownId(string kind)
    {
        // Proof that the length gate is 255-inclusive: the ONLY complaint left is the
        // unknown hit id, which is resolved after validation and still before COM.
        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            CallDerived(kind, subject: new string('x', 255), importance: null));
        Assert.Contains("Unknown id", ex.Message, StringComparison.Ordinal);
    }

    // ---------------------------------------------------------------- body / body_html (batch B - B1)

    [Theory]
    [InlineData("new")]
    [InlineData("reply")]
    [InlineData("replyall")]
    [InlineData("forward")]
    public void ExactlyOneBodyForm_IsRequired_AndTheErrorNamesBoth(string kind)
    {
        ArgumentException missing = Assert.Throws<ArgumentException>(() => CallWithBody(kind, body: null, bodyHtml: null));
        Assert.Contains("body", missing.Message, StringComparison.Ordinal);
        Assert.Contains("body_html", missing.Message, StringComparison.Ordinal);

        ArgumentException both = Assert.Throws<ArgumentException>(() => CallWithBody(kind, body: "text", bodyHtml: "<p>html</p>"));
        Assert.Contains("mutually exclusive", both.Message, StringComparison.Ordinal);
        Assert.Contains("body_html", both.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("new")]
    [InlineData("reply")]
    [InlineData("replyall")]
    [InlineData("forward")]
    public void BodyHtmlWithNothingVisible_IsRejectedBeforeAnyComWork(string kind)
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            CallWithBody(kind, body: null, bodyHtml: "<script>alert(1)</script><!-- nothing -->"));

        Assert.Contains("no usable content", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("new")]
    [InlineData("reply")]
    [InlineData("replyall")]
    [InlineData("forward")]
    public void OversizeBodyHtml_IsRejectedBeforeAnyComWork(string kind)
    {
        string huge = "<p>" + new string('a', OutlookAI.Core.Text.HtmlFragmentNormalizer.MaxInputChars) + "</p>";

        ArgumentException ex = Assert.Throws<ArgumentException>(() => CallWithBody(kind, body: null, bodyHtml: huge));

        Assert.Contains("body_html", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("reply")]
    [InlineData("replyall")]
    [InlineData("forward")]
    public void AcceptedBodyHtml_PassesValidation_AndOnlyThenReachesIdResolution(string kind)
    {
        // A well-formed fragment must get PAST every body gate; the derived tools then
        // fail on the unknown hit id - still before any COM work - which proves body_html
        // was accepted rather than silently rejected. (new_draft has no further pre-COM
        // gate after the body, so its acceptance is pinned on the wire and live instead.)
        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            CallWithBody(kind, body: null, bodyHtml: "<h1>Title</h1><p>Hello</p>"));

        Assert.Contains("Unknown id", ex.Message, StringComparison.Ordinal);
    }

    private void CallWithBody(string kind, string? body, string? bodyHtml)
    {
        switch (kind)
        {
            case "new":
                _service.NewDraft("hub@example.com", "a@b.example", null, "subject", body, display: false,
                    signature: null, bcc: null, importance: null, requestReadReceipt: null, bodyHtml: bodyHtml);
                break;
            case "reply":
                _service.ReplyDraft("h424242", body, replyAll: false, display: false, signature: null,
                    cc: null, bcc: null, subject: null, importance: null, requestReadReceipt: null, bodyHtml: bodyHtml);
                break;
            case "replyall":
                _service.ReplyDraft("h424242", body, replyAll: true, display: false, signature: null,
                    cc: null, bcc: null, subject: null, importance: null, requestReadReceipt: null, bodyHtml: bodyHtml);
                break;
            default:
                _service.ForwardDraft("h424242", body, "a@b.example", display: false, signature: null,
                    cc: null, bcc: null, subject: null, importance: null, requestReadReceipt: null, bodyHtml: bodyHtml);
                break;
        }
    }

    private void CallDerived(string kind, string? subject, string? importance)
    {
        switch (kind)
        {
            case "reply":
                _service.ReplyDraft("h424242", "body", replyAll: false, display: false, signature: null,
                    cc: null, bcc: null, subject: subject, importance: importance);
                break;
            case "replyall":
                _service.ReplyDraft("h424242", "body", replyAll: true, display: false, signature: null,
                    cc: null, bcc: null, subject: subject, importance: importance);
                break;
            default:
                _service.ForwardDraft("h424242", "body", "a@b.example", display: false, signature: null,
                    cc: null, bcc: null, subject: subject, importance: importance);
                break;
        }
    }
}
