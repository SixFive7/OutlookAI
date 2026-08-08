using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1 for the Phase-5 high-friction send policy (v3.MD D4/L5): the one-time confirm
/// token store (issue/consume/expiry/binding/single-use/replacement/cap) and the
/// canonical content hash the token is bound to. Pure logic - no COM, no Outlook.
/// </summary>
public sealed class SendConfirmationTests
{
    private const string DraftA = "00AA11BB22CC33DD44EE55FF66AA77BB88CC99DD00AA11BB";
    private const string DraftB = "FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF00FF01";
    private const string HashA = "aaaa1111";
    private const string HashB = "bbbb2222";

    // ------------------------------------------------------------------ token store

    [Fact]
    public void Issue_ReturnsPrefixedUnpredictableTokens()
    {
        var store = new SendConfirmationTokens();

        string t1 = store.Issue(DraftA, HashA);
        string t2 = store.Issue(DraftB, HashB);

        Assert.StartsWith("confirm-", t1, StringComparison.Ordinal);
        Assert.Equal("confirm-".Length + 32, t1.Length);
        Assert.NotEqual(t1, t2);
    }

    [Fact]
    public void Consume_ValidToken_ExactlyOnce()
    {
        var store = new SendConfirmationTokens();
        string token = store.Issue(DraftA, HashA);

        Assert.Equal(SendTokenDecision.Valid, store.Consume(token, DraftA, HashA));

        // SINGLE-USE: the same token never validates twice.
        Assert.Equal(SendTokenDecision.UnknownOrUsed, store.Consume(token, DraftA, HashA));
        Assert.Equal(0, store.PendingCount);
    }

    [Fact]
    public void Consume_IsCaseInsensitiveOnEntryId_ButNotOnHash()
    {
        var store = new SendConfirmationTokens();
        string token = store.Issue(DraftA.ToUpperInvariant(), "abcdef");

        Assert.Equal(SendTokenDecision.Valid, store.Consume(token, DraftA.ToLowerInvariant(), "abcdef"));
    }

    [Fact]
    public void Consume_WrongDraft_RefusesAndBurnsTheToken()
    {
        var store = new SendConfirmationTokens();
        string token = store.Issue(DraftA, HashA);

        // Bound to draft id: using it on another draft refuses...
        Assert.Equal(SendTokenDecision.DraftMismatch, store.Consume(token, DraftB, HashA));

        // ...and BURNS it - even the originally-bound draft cannot use it anymore.
        Assert.Equal(SendTokenDecision.UnknownOrUsed, store.Consume(token, DraftA, HashA));
    }

    [Fact]
    public void Consume_ChangedContent_RefusesAndBurnsTheToken()
    {
        var store = new SendConfirmationTokens();
        string token = store.Issue(DraftA, HashA);

        // Bound to content hash: a modified draft invalidates the token.
        Assert.Equal(SendTokenDecision.ContentChanged, store.Consume(token, DraftA, HashB));
        Assert.Equal(SendTokenDecision.UnknownOrUsed, store.Consume(token, DraftA, HashA));
    }

    [Fact]
    public void Consume_UnknownOrBlankToken_Refuses()
    {
        var store = new SendConfirmationTokens();

        Assert.Equal(SendTokenDecision.UnknownOrUsed, store.Consume("confirm-doesnotexist", DraftA, HashA));
        Assert.Equal(SendTokenDecision.UnknownOrUsed, store.Consume("", DraftA, HashA));
        Assert.Equal(SendTokenDecision.UnknownOrUsed, store.Consume("   ", DraftA, HashA));
    }

    [Fact]
    public void Consume_ExpiredToken_Refuses()
    {
        DateTime now = new DateTime(2026, 7, 23, 12, 0, 0, DateTimeKind.Utc);
        var store = new SendConfirmationTokens(TimeSpan.FromSeconds(120), () => now);
        string token = store.Issue(DraftA, HashA);

        now = now.AddSeconds(121);
        Assert.Equal(SendTokenDecision.Expired, store.Consume(token, DraftA, HashA));

        // Expired = burned, like every other consumption.
        Assert.Equal(SendTokenDecision.UnknownOrUsed, store.Consume(token, DraftA, HashA));
    }

    [Fact]
    public void Consume_JustInsideTtl_StillValid()
    {
        DateTime now = new DateTime(2026, 7, 23, 12, 0, 0, DateTimeKind.Utc);
        var store = new SendConfirmationTokens(TimeSpan.FromSeconds(120), () => now);
        string token = store.Issue(DraftA, HashA);

        now = now.AddSeconds(119);
        Assert.Equal(SendTokenDecision.Valid, store.Consume(token, DraftA, HashA));
    }

    [Fact]
    public void Issue_ReplacesPendingTokenForTheSameDraft()
    {
        var store = new SendConfirmationTokens();
        string first = store.Issue(DraftA, HashA);
        string second = store.Issue(DraftA, HashA);

        // Only the newest token per draft is live.
        Assert.Equal(SendTokenDecision.UnknownOrUsed, store.Consume(first, DraftA, HashA));
        Assert.Equal(SendTokenDecision.Valid, store.Consume(second, DraftA, HashA));
    }

    [Fact]
    public void Issue_PrunesExpired_AndCapsPendingTokens()
    {
        DateTime now = new DateTime(2026, 7, 23, 12, 0, 0, DateTimeKind.Utc);
        var store = new SendConfirmationTokens(TimeSpan.FromSeconds(60), () => now);
        store.Issue(DraftA, HashA);
        Assert.Equal(1, store.PendingCount);

        // Expired entries are pruned on the next issue.
        now = now.AddSeconds(61);
        store.Issue(DraftB, HashB);
        Assert.Equal(1, store.PendingCount);

        // Hard cap: distinct drafts never accumulate beyond the bound.
        for (int i = 0; i < 40; i++)
        {
            store.Issue("draft" + i.ToString("D4"), HashA);
        }

        Assert.True(store.PendingCount <= 32, "pending token store must stay bounded");
    }

    [Fact]
    public void Issue_ValidatesArguments()
    {
        var store = new SendConfirmationTokens();

        Assert.Throws<ArgumentException>(() => store.Issue(" ", HashA));
        Assert.Throws<ArgumentException>(() => store.Issue(DraftA, " "));
        Assert.Throws<ArgumentOutOfRangeException>(() => new SendConfirmationTokens(TimeSpan.Zero));
    }

    [Fact]
    public void Invalidate_RemovesPendingToken()
    {
        var store = new SendConfirmationTokens();
        string token = store.Issue(DraftA, HashA);

        store.Invalidate(token);

        Assert.Equal(SendTokenDecision.UnknownOrUsed, store.Consume(token, DraftA, HashA));
        Assert.Equal(0, store.PendingCount);
    }

    [Fact]
    public void DefaultTimeToLive_IsShort()
    {
        // D4: the token must expire QUICKLY - pin the product default.
        Assert.Equal(TimeSpan.FromSeconds(120), SendConfirmationTokens.DefaultTimeToLive);
        Assert.Equal(TimeSpan.FromSeconds(120), new SendConfirmationTokens().TimeToLive);
    }

    // ------------------------------------------------------------------ content hash

    private static ComRecipientInfo To(string address) => new("to", "Name", address);

    [Fact]
    public void Hash_IsDeterministic_AndRecipientOrderInsensitive()
    {
        var r1 = new[] { To("a@example.com"), To("b@example.com") };
        var r2 = new[] { To("b@example.com"), To("a@example.com") };

        string h1 = SendContentHash.Compute("Subject", r1, "body", null, null, null);
        string h2 = SendContentHash.Compute("Subject", r2, "body", null, null, null);

        Assert.Equal(h1, h2);
        Assert.Equal(64, h1.Length);
        Assert.Matches("^[0-9a-f]{64}$", h1);
    }

    [Theory]
    [InlineData("Subject2", "body", null)] // subject change
    [InlineData("Subject", "body CHANGED", null)] // body change
    [InlineData("Subject", "body", "boss@example.com")] // on-behalf-of change
    public void Hash_ChangesWhenAnyBoundPartChanges(string subject, string body, string? onBehalfOf)
    {
        var recipients = new[] { To("a@example.com") };
        string baseline = SendContentHash.Compute("Subject", recipients, "body", null, null, null);

        string changed = SendContentHash.Compute(subject, recipients, body, onBehalfOf, null, null);

        Assert.NotEqual(baseline, changed);
    }

    [Fact]
    public void Hash_ChangesOnRecipientSetOrKindChange()
    {
        string baseline = SendContentHash.Compute("S", new[] { To("a@example.com") }, "b", null, null, null);

        string added = SendContentHash.Compute("S", new[] { To("a@example.com"), To("c@example.com") }, "b", null, null, null);
        string ccInstead = SendContentHash.Compute(
            "S", new[] { new ComRecipientInfo("cc", "Name", "a@example.com") }, "b", null, null, null);

        Assert.NotEqual(baseline, added);
        Assert.NotEqual(baseline, ccInstead);
    }

    [Fact]
    public void Hash_NormalizesLineEndings_AndCase_Robustly()
    {
        var recipients = new[] { To("A@Example.COM") };

        // CRLF vs LF for the SAME body must not flip the hash (COM read paths differ).
        string crlf = SendContentHash.Compute("S", recipients, "line1\r\nline2", null, null, null);
        string lf = SendContentHash.Compute("S", new[] { To("a@example.com") }, "line1\nline2", null, null, null);
        Assert.Equal(crlf, lf);

        // But an actual body difference does.
        Assert.NotEqual(crlf, SendContentHash.Compute("S", recipients, "line1\n\nline2", null, null, null));

        // On-behalf-of is trimmed + case-normalized (same both calls contract).
        Assert.Equal(
            SendContentHash.Compute("S", recipients, "b", " Boss@Example.com ", null, null),
            SendContentHash.Compute("S", recipients, "b", "boss@example.com", null, null));
    }

    [Fact]
    public void Hash_NullSafeOnSubjectAndBody()
    {
        string h = SendContentHash.Compute(null, Array.Empty<ComRecipientInfo>(), null, null, null, null);
        Assert.Equal(64, h.Length);
    }

    // --------------------------------------------- D46/C3: the attachment + HTML interlock

    private static ComAttachmentInfo Att(int index, string name, long size) => new(index, name, size);

    [Fact]
    public void Hash_ChangesWhenAnAttachmentIsAdded_SoAPendingTokenIsInvalidated()
    {
        var recipients = new[] { To("a@example.com") };
        string withoutFile = SendContentHash.Compute("S", recipients, "b", null, null, null);

        string withFile = SendContentHash.Compute(
            "S", recipients, "b", null, new[] { Att(1, "offer.pdf", 1234) }, null);

        Assert.NotEqual(withoutFile, withFile);
    }

    [Fact]
    public void Hash_ChangesWhenAnAttachmentIsRemoved()
    {
        var recipients = new[] { To("a@example.com") };
        var both = new[] { Att(1, "a.pdf", 10), Att(2, "b.pdf", 20) };

        string withBoth = SendContentHash.Compute("S", recipients, "b", null, both, null);
        string withOne = SendContentHash.Compute("S", recipients, "b", null, new[] { Att(1, "a.pdf", 10) }, null);

        Assert.NotEqual(withBoth, withOne);
    }

    [Fact]
    public void Hash_ChangesWhenAnAttachmentKeepsItsNameButNotItsSize()
    {
        // The "swap the file, keep the name" attack: size is hashed as well as the name.
        var recipients = new[] { To("a@example.com") };

        string original = SendContentHash.Compute("S", recipients, "b", null, new[] { Att(1, "offer.pdf", 1000) }, null);
        string swapped = SendContentHash.Compute("S", recipients, "b", null, new[] { Att(1, "offer.pdf", 2000) }, null);

        Assert.NotEqual(original, swapped);
    }

    [Fact]
    public void Hash_IsAttachmentOrderInsensitive_AndNameCaseInsensitive()
    {
        // Outlook may enumerate the collection differently after a reopen - the hash must
        // depend on the SET, exactly like the recipient list, or a send would spuriously
        // refuse.
        var recipients = new[] { To("a@example.com") };
        string ab = SendContentHash.Compute(
            "S", recipients, "b", null, new[] { Att(1, "a.pdf", 10), Att(2, "B.pdf", 20) }, null);
        string ba = SendContentHash.Compute(
            "S", recipients, "b", null, new[] { Att(1, "b.pdf", 20), Att(2, "A.pdf", 10) }, null);

        Assert.Equal(ab, ba);
    }

    [Fact]
    public void Hash_TreatsNullAndEmptyAttachmentListsAsTheSame()
    {
        var recipients = new[] { To("a@example.com") };

        Assert.Equal(
            SendContentHash.Compute("S", recipients, "b", null, null, null),
            SendContentHash.Compute("S", recipients, "b", null, Array.Empty<ComAttachmentInfo>(), null));
    }

    [Fact]
    public void Hash_ChangesOnAMarkupOnlyEdit_ThatThePlainBodyCannotShow()
    {
        // WHY THIS EXISTS (D46/C3): Outlook derives MailItem.Body from HTMLBody, and that
        // derivation is lossy - retargeting a hyperlink while its visible label stays put
        // leaves the plain text byte-identical. Without the HTML digest a confirm token
        // would survive exactly that edit.
        var recipients = new[] { To("a@example.com") };
        const string PlainBody = "Click here to pay.";
        string honest = SendContentHash.DigestHtml("<p>Click <a href=\"https://bank.example/pay\">here</a> to pay.</p>")!;
        string swapped = SendContentHash.DigestHtml("<p>Click <a href=\"https://evil.example/pay\">here</a> to pay.</p>")!;

        Assert.NotEqual(honest, swapped);
        Assert.NotEqual(
            SendContentHash.Compute("S", recipients, PlainBody, null, null, honest),
            SendContentHash.Compute("S", recipients, PlainBody, null, null, swapped));
    }

    [Fact]
    public void DigestHtml_IsStableAndNullSafe()
    {
        Assert.Null(SendContentHash.DigestHtml(null));
        Assert.Null(SendContentHash.DigestHtml(string.Empty));

        string a = SendContentHash.DigestHtml("<p>x</p>")!;
        Assert.Equal(a, SendContentHash.DigestHtml("<p>x</p>"));
        Assert.Matches("^[0-9a-f]{64}$", a);
    }

    // ------------------------------------------------------------------ refusal type

    [Fact]
    public void SendRefusedException_CarriesReasonCode()
    {
        var ex = new SendRefusedException("token_expired", "msg");
        Assert.Equal("token_expired", ex.Reason);
        Assert.Equal("msg", ex.Message);
    }
}
