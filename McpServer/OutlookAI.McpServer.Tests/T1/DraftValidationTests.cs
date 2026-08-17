using System.Globalization;

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
        string subject = new('x', MailService.SubjectCharsCap + 1);
        Assert.Throws<ArgumentException>(() =>
            _service.NewDraft("hub@example.com", "a@b.example", null, subject, "body", display: false));
    }

    // ------------------------------------------------- the subject cap teaches by erroring
    //
    // new_draft does not advertise the 255-character cap on its subject argument, unlike
    // the three derived tools: an over-long subject is rare, and the established pattern
    // here is a refusal the model can self-correct from without asking - the same shape as
    // the writing-rules gate, send's confirm-token refusal and the fail-closed attachment
    // validation. That trade only holds while the error carries everything the retry needs,
    // so what the description does not say is pinned here instead. The message used to be
    // "subject is too long (max 255 characters)." - the limit but not the breach, so a
    // model that had built the subject from a variable could not tell whether it had missed
    // by one character or by a thousand.

    /// <summary>Every draft path that enforces the cap, called with the same over-long subject.</summary>
    private static readonly string[] SubjectCapPaths = { "new", "reply", "replyall", "forward", "update" };

    private ArgumentException RejectOverlongSubject(string path, int length)
    {
        string subject = new('x', length);
        return Assert.Throws<ArgumentException>(() =>
        {
            switch (path)
            {
                case "new":
                    _service.NewDraft("hub@example.com", "a@b.example", null, subject, "body", display: false);
                    break;
                case "update":
                    _service.UpdateDraft("h424242", subject: subject, display: false);
                    break;
                default:
                    CallDerived(path, subject, importance: null);
                    break;
            }
        });
    }

    [Fact]
    public void OverlongSubject_IsRefusedWithAnErrorTheModelCanSelfCorrectFrom_OnEveryDraftPath()
    {
        int supplied = MailService.SubjectCharsCap + 57;
        foreach (string path in SubjectCapPaths)
        {
            string message = RejectOverlongSubject(path, supplied).Message;

            // 1. It names the argument at fault - a draft call carries several strings.
            Assert.Contains("subject", message, StringComparison.OrdinalIgnoreCase);

            // 2. The LIMIT, as a number, derived from the constant the service enforces.
            Assert.Contains(
                "max " + MailService.SubjectCharsCap.ToString(CultureInfo.InvariantCulture) + " characters",
                message,
                StringComparison.Ordinal);

            // 3. What was WRONG with what was supplied, including how long it actually was.
            //    Without this the model cannot tell a near miss from an order-of-magnitude
            //    one, and cannot compute how much to cut.
            Assert.Contains(
                supplied.ToString(CultureInfo.InvariantCulture) + " characters supplied",
                message,
                StringComparison.Ordinal);

            // 4. That nothing was written, and that the SAME call with a shorter subject
            //    works - so the retry is obvious and needs no user round trip.
            Assert.Contains("Nothing was created or changed", message, StringComparison.Ordinal);
            Assert.Contains("call again", message, StringComparison.OrdinalIgnoreCase);

            // The subject itself is never echoed: its LENGTH is the whole diagnosis, and an
            // error is not a place to reflect content back.
            Assert.DoesNotContain(new string('x', 40), message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void OverlongSubject_IsRefusedIdentically_ByEveryDraftPath()
    {
        // One gate, one sentence: four paths that disagree about the same mistake teach a
        // model that the limit depends on which tool it picked.
        string[] messages = SubjectCapPaths
            .Select(path => RejectOverlongSubject(path, MailService.SubjectCharsCap + 1).Message)
            .Distinct(StringComparer.Ordinal)
            .ToArray();

        string only = Assert.Single(messages);
        Assert.StartsWith(
            MailService.BuildOverlongSubjectMessage(MailService.SubjectCharsCap + 1),
            only,
            StringComparison.Ordinal);
    }

    [Fact]
    public void OverlongSubject_ReportsTheLengthItWasActuallyGiven()
    {
        Assert.Contains("256 characters supplied", MailService.BuildOverlongSubjectMessage(256), StringComparison.Ordinal);
        Assert.Contains("9001 characters supplied", MailService.BuildOverlongSubjectMessage(9001), StringComparison.Ordinal);
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

    // --------------------------------------------- D46/C1+C2: update_draft / discard_draft

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void UpdateDraft_RequiresAnId(string? id)
    {
        Assert.Throws<ArgumentException>(() => _service.UpdateDraft(id!, subject: "s", display: false));
    }

    [Fact]
    public void UpdateDraft_WithNoChangeRequested_IsRejectedBeforeAnyIdResolution()
    {
        // Rejected BEFORE the id is even resolved: a no-op update would still open the
        // item and re-save it, which is a mailbox write for nothing.
        ArgumentException ex = Assert.Throws<ArgumentException>(() => _service.UpdateDraft("h424242", display: false));
        Assert.Contains("Nothing to update", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void UpdateDraft_BothBodyForms_AreMutuallyExclusive()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(
            () => _service.UpdateDraft("h424242", body: "text", bodyHtml: "<p>x</p>", display: false));
        Assert.Contains("mutually exclusive", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void UpdateDraft_BlankSubject_IsRejected_RatherThanClearingTheSubject()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(
            () => _service.UpdateDraft("h424242", subject: "   ", display: false));
        Assert.Contains("must not be blank", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void UpdateDraft_OverlongSubject_IsRejected()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(
            () => _service.UpdateDraft("h424242", subject: new string('s', MailService.SubjectCharsCap + 1), display: false));
        Assert.Contains("too long", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void UpdateDraft_EmptyToList_IsRejected_BecauseToIsReplaceNotAppend()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(
            () => _service.UpdateDraft("h424242", to: " ; , ", display: false));
        Assert.Contains("REPLACES the To list", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void UpdateDraft_EmptyCcOrBcc_IsAccepted_BecauseThatIsHowYouClearThem()
    {
        // An empty cc/bcc is a legitimate REPLACE-with-nothing; only the id may complain.
        foreach (Func<UpdateDraftOutcome> call in new Func<UpdateDraftOutcome>[]
                 {
                     () => _service.UpdateDraft("h424242", cc: "", display: false),
                     () => _service.UpdateDraft("h424242", bcc: "", display: false),
                 })
        {
            ArgumentException ex = Assert.Throws<ArgumentException>(() => call());
            Assert.Contains("Unknown id", ex.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void UpdateDraft_InvalidImportance_IsRejectedBeforeAnyComWork()
    {
        Assert.Throws<ArgumentException>(() => _service.UpdateDraft("h424242", importance: "urgent", display: false));
    }

    [Fact]
    public void UpdateDraft_BadAttachmentPath_IsRejectedBeforeAnyComWork()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(
            () => _service.UpdateDraft("h424242", attachments: new[] { "relative.pdf" }, display: false));
        Assert.Contains("absolute path", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void UpdateDraft_AcceptedArguments_ReachIdResolution_ProvingTheyPassedValidation()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(
            () => _service.UpdateDraft(
                "h424242", body: "new text", subject: "New subject", to: "a@b.example",
                importance: "high", requestReadReceipt: true,
                removeAttachments: new[] { "old.pdf" }, display: false));
        Assert.Contains("Unknown id", ex.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void DiscardDraft_RequiresAnId(string? id)
    {
        Assert.Throws<ArgumentException>(() => _service.DiscardDraft(id!));
    }

    [Fact]
    public void DiscardDraft_UnknownHitId_FailsIdResolution()
    {
        ArgumentException ex = Assert.Throws<ArgumentException>(() => _service.DiscardDraft("h424242"));
        Assert.Contains("Unknown id", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DiscardDraft_EntryIdNotFromThisServer_IsRefusedBeforeAnyComWork()
    {
        // THE guardrail (S1 v3): a well-formed EntryID no draft tool of this process ever
        // returned is refused by the registry, with Outlook never touched. The gateway
        // here cannot even start Outlook, so reaching COM would fail loudly instead.
        DraftRefusedException ex = Assert.Throws<DraftRefusedException>(
            () => _service.DiscardDraft(new string('A', 96)));

        Assert.Equal("not_created_by_this_server", ex.Reason);
        Assert.Contains("not created or last updated by this server session", ex.Message, StringComparison.Ordinal);
        Assert.Contains("Delete it in Outlook instead", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void DiscardDraft_RegistryStartsEmpty_SoAFreshServerCanDeleteNothing()
    {
        // A server restart must not inherit deletion rights over items it cannot vouch for.
        Assert.Equal(0, _service.DraftRegistry.Count);
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
            CallDerived(kind, subject: new string('x', MailService.SubjectCharsCap + 1), importance: null));
        Assert.Contains("subject", ex.Message, StringComparison.OrdinalIgnoreCase);

        // The message quotes the cap the service actually enforces, so an agent that hit it
        // is told the real limit rather than a number someone typed once.
        Assert.Contains(
            "max " + MailService.SubjectCharsCap.ToString(System.Globalization.CultureInfo.InvariantCulture)
            + " characters",
            ex.Message,
            StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("reply")]
    [InlineData("replyall")]
    [InlineData("forward")]
    public void DerivedDrafts_AcceptASubjectOverrideExactlyAtTheCap_AndFailOnlyOnTheUnknownId(string kind)
    {
        // Proof that the length gate is INCLUSIVE at MailService.SubjectCharsCap: the ONLY
        // complaint left is the unknown hit id, which is resolved after validation and
        // still before COM.
        ArgumentException ex = Assert.Throws<ArgumentException>(() =>
            CallDerived(kind, subject: new string('x', MailService.SubjectCharsCap), importance: null));
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
