using OutlookAI.Core.IndexSearch;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the attachment-recall fix (soak fix 16, v3.MD section 0.8 block (q)): the SQL no
/// longer filters on System.Kind under a mapi scope, so admission happens here - a
/// message-level row must be kind 'email', an attachment-content row (/at=) is kept
/// whatever its kind, and nothing outside the mapi namespace is ever kept.
/// </summary>
public sealed class IndexRowFilterTests
{
    private const string MessageUrl =
        "mapi16://{SID}/alice@example.com($ab12)/0/Inbox/\uD5B4\uD5B4\uD5B4";

    private const string AttachmentUrl = MessageUrl + "/at=1:photo.jpg";

    private static IndexHit Hit(string url, params string[] kinds)
    {
        // The mapper is the only writer of these (internal setters); build the same shape
        // it would produce for the given row.
        return IndexRowMapper.Map(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = url,
            ["System.Kind"] = kinds,
        });
    }

    [Theory]
    [InlineData("picture")]
    [InlineData("communication")]
    [InlineData("calendar")]
    [InlineData("music")]
    [InlineData("video")]
    [InlineData("document")]
    public void AttachmentRows_AreKeptWhateverTheirKind(string kind)
    {
        // THE DEFECT: (Kind='email' OR Kind='document') kept only 'document', so 709 of
        // 3,139 measured attachment rows (22.6%) could never find their parent mail.
        Assert.True(IndexRowFilter.Keep(Hit(AttachmentUrl, kind), KindFilter.EmailAndDocuments));
        Assert.True(IndexRowFilter.Keep(Hit(AttachmentUrl, kind), KindFilter.DocumentsOnly));
    }

    [Fact]
    public void MessageRows_AreKeptOnlyWhenTheyAreMail()
    {
        Assert.True(IndexRowFilter.Keep(Hit(MessageUrl, "email"), KindFilter.EmailAndDocuments));

        // Meeting requests/responses live in mail folders and index as 'calendar'. They
        // were excluded before and stay excluded - the fix must not widen message rows.
        Assert.False(IndexRowFilter.Keep(Hit(MessageUrl, "calendar"), KindFilter.EmailAndDocuments));
        Assert.False(IndexRowFilter.Keep(Hit(MessageUrl, "document"), KindFilter.EmailAndDocuments));
        Assert.False(IndexRowFilter.Keep(Hit(MessageUrl), KindFilter.EmailAndDocuments));
    }

    [Fact]
    public void KindComparisonIsCaseInsensitive_TheProviderReturnsMixedCase()
    {
        Assert.True(IndexRowFilter.Keep(Hit(MessageUrl, "Email"), KindFilter.EmailAndDocuments));
        Assert.True(IndexRowFilter.Keep(Hit(MessageUrl, "EMAIL"), KindFilter.EmailAndDocuments));
        Assert.True(IndexRowFilter.HasEmailKind(new[] { "Document", "Email" }));
        Assert.False(IndexRowFilter.HasEmailKind(new[] { "Document" }));
        Assert.False(IndexRowFilter.HasEmailKind(null));
    }

    [Fact]
    public void NonMapiRows_AreNeverKept_EvenWhenTheirKindIsEmail()
    {
        // An .eml/.msg file on disk indexes as kind 'email'. Without a SCOPE the statement
        // can reach it, so the namespace check here is load-bearing, not decoration.
        Assert.False(IndexRowFilter.Keep(Hit("file:C:/mail/archive.eml", "email"), KindFilter.EmailAndDocuments));
        Assert.False(IndexRowFilter.Keep(Hit("file:C:/pictures/holiday.jpg", "picture"), KindFilter.EmailAndDocuments));
        Assert.False(IndexRowFilter.IsMapiRow(null));
        Assert.True(IndexRowFilter.IsMapiRow("MAPI16://x"));
    }

    [Fact]
    public void EmailOnly_KeepsMessagesAndRejectsAttachmentRows()
    {
        Assert.True(IndexRowFilter.Keep(Hit(MessageUrl, "email"), KindFilter.EmailOnly));
        Assert.False(IndexRowFilter.Keep(Hit(AttachmentUrl, "email"), KindFilter.EmailOnly));
    }

    [Fact]
    public void DocumentsOnly_KeepsOnlyAttachmentRows()
    {
        Assert.False(IndexRowFilter.Keep(Hit(MessageUrl, "email"), KindFilter.DocumentsOnly));
        Assert.True(IndexRowFilter.Keep(Hit(AttachmentUrl, "picture"), KindFilter.DocumentsOnly));
    }

    [Fact]
    public void AttachmentDetectionIsAUrlTest_NotAParseResult()
    {
        // A malformed attachment URL must still count as an attachment: promoting it to a
        // message row would then judge it on the ATTACHMENT's kind and drop it.
        Assert.True(IndexRowFilter.IsAttachmentRow("mapi16://x/0/Inbox/notavalidid/at=9:x.png"));
        Assert.False(IndexRowFilter.IsAttachmentRow(MessageUrl));
        Assert.False(IndexRowFilter.IsAttachmentRow(null));
    }

    [Fact]
    public void Keep_RejectsAnUnknownKindFilter()
    {
        Assert.Throws<ArgumentException>(() => IndexRowFilter.Keep(Hit(MessageUrl, "email"), (KindFilter)99));
        Assert.Throws<ArgumentNullException>(() => IndexRowFilter.Keep(null!, KindFilter.EmailAndDocuments));
    }

    [Fact]
    public void ComputeSqlTop_OverFetchesEnoughToSurviveFiltering()
    {
        // Scoped: modest widening - the only rows lost are message-level calendar items
        // (0.3-1.2% of a real folder). Unscoped: wider, the file system is in play.
        Assert.Equal(62, IndexRowFilter.ComputeSqlTop(26, scoped: true, maxTop: 5000));
        Assert.Equal(124, IndexRowFilter.ComputeSqlTop(26, scoped: false, maxTop: 5000));
        Assert.Equal(12, IndexRowFilter.ComputeSqlTop(1, scoped: true, maxTop: 5000));

        // Never above the provider ceiling, never below what the caller asked for.
        Assert.Equal(5000, IndexRowFilter.ComputeSqlTop(5000, scoped: true, maxTop: 5000));
        Assert.Equal(5000, IndexRowFilter.ComputeSqlTop(3000, scoped: true, maxTop: 5000));
        Assert.Throws<ArgumentException>(() => IndexRowFilter.ComputeSqlTop(0, true, 5000));
        Assert.Throws<ArgumentException>(() => IndexRowFilter.ComputeSqlTop(10, true, 0));
    }

    [Fact]
    public void UnscopedKindList_CoversEveryMeasuredAttachmentKind()
    {
        foreach (string kind in new[] { "email", "document", "picture", "communication", "calendar", "music", "video" })
        {
            Assert.Contains(kind, IndexRowFilter.UnscopedKinds);
        }
    }
}
