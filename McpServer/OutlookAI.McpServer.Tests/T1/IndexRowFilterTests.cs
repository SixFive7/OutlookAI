using OutlookAI.Core.IndexSearch;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the attachment-recall fix (soak fix 16, v3.MD section 0.8 block (q)): the SQL no
/// longer filters on System.Kind under a mapi scope, so admission happens here - an
/// attachment-content row (/at=) is kept whatever its kind, and nothing outside the mapi
/// namespace is ever kept.
/// <para>
/// AND, since gap B3 (maintainer decision 2026-08-18), a message-level row is kept whatever
/// its kind too. It used to need <c>email</c>, which dropped meeting requests - they index
/// as <c>calendar</c> - from every search while the freshness sweep beside them returned
/// all of them. The assertions that changed are marked where they changed.
/// </para>
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
        Assert.True(IndexRowFilter.Keep(Hit(AttachmentUrl, kind), KindFilter.MessagesAndAttachments));
        Assert.True(IndexRowFilter.Keep(Hit(AttachmentUrl, kind), KindFilter.AttachmentsOnly));
    }

    /// <summary>
    /// THE ANSWER CHANGED HERE (gap B3). This test used to be called
    /// <c>MessageRows_AreKeptOnlyWhenTheyAreMail</c> and asserted the opposite of the last
    /// three lines: a message-level row of kind <c>calendar</c>, <c>document</c> or no kind
    /// at all was refused. That is what dropped every meeting request and response from a
    /// search - while the freshness sweep beside it, which never filtered by class, handed
    /// them back - so the same query returned a different item set depending on which tier
    /// reached the mail first, with nothing in the payload saying so.
    /// </summary>
    [Fact]
    public void MessageRows_AreKeptWhateverTheirKind()
    {
        Assert.True(IndexRowFilter.Keep(Hit(MessageUrl, "email"), KindFilter.MessagesAndAttachments));

        Assert.True(IndexRowFilter.Keep(Hit(MessageUrl, "calendar"), KindFilter.MessagesAndAttachments));
        Assert.True(IndexRowFilter.Keep(Hit(MessageUrl, "document"), KindFilter.MessagesAndAttachments));
        Assert.True(IndexRowFilter.Keep(Hit(MessageUrl), KindFilter.MessagesAndAttachments));
    }

    [Fact]
    public void KindComparisonIsCaseInsensitive_TheProviderReturnsMixedCase()
    {
        // Still exercised through the one shape that reads the kind at all.
        Assert.True(IndexRowFilter.Keep(Hit(MessageUrl, "Email"), KindFilter.MailKindOnly));
        Assert.True(IndexRowFilter.Keep(Hit(MessageUrl, "EMAIL"), KindFilter.MailKindOnly));
        Assert.True(IndexRowFilter.HasEmailKind(new[] { "Document", "Email" }));
        Assert.False(IndexRowFilter.HasEmailKind(new[] { "Document" }));
        Assert.False(IndexRowFilter.HasEmailKind(null));
    }

    [Fact]
    public void NonMapiRows_AreNeverKept_EvenWhenTheirKindIsEmail()
    {
        // An .eml/.msg file on disk indexes as kind 'email'. Without a SCOPE the statement
        // can reach it, so the namespace check here is load-bearing, not decoration - and
        // since message rows stopped being judged on their kind it is the ONLY thing
        // standing between the widest shape and the file system.
        Assert.False(IndexRowFilter.Keep(Hit("file:C:/mail/archive.eml", "email"), KindFilter.MessagesAndAttachments));
        Assert.False(IndexRowFilter.Keep(Hit("file:C:/pictures/holiday.jpg", "picture"), KindFilter.MessagesAndAttachments));
        Assert.False(IndexRowFilter.IsMapiRow(null));
        Assert.True(IndexRowFilter.IsMapiRow("MAPI16://x"));
    }

    [Fact]
    public void MailKindOnly_KeepsMailMessagesAndRejectsAttachmentRows()
    {
        // The store-discovery shape: it wants a row that is certainly mail, so this is the
        // one place a kind still decides admission.
        Assert.True(IndexRowFilter.Keep(Hit(MessageUrl, "email"), KindFilter.MailKindOnly));
        Assert.False(IndexRowFilter.Keep(Hit(AttachmentUrl, "email"), KindFilter.MailKindOnly));
        Assert.False(IndexRowFilter.Keep(Hit(MessageUrl, "calendar"), KindFilter.MailKindOnly));
    }

    [Fact]
    public void AttachmentsOnly_KeepsOnlyAttachmentRows()
    {
        Assert.False(IndexRowFilter.Keep(Hit(MessageUrl, "email"), KindFilter.AttachmentsOnly));
        Assert.True(IndexRowFilter.Keep(Hit(AttachmentUrl, "picture"), KindFilter.AttachmentsOnly));
    }

    /// <summary>
    /// Gap C2, then gap B3: a message row is admitted whatever its item class. A meeting
    /// request indexes as <c>calendar</c> and carries the surrounding mail's ConversationID,
    /// so a kind-narrowed filter dropped a real member of a conversation the tool promises
    /// whole - and, once B3 unified the tiers, a real hit of a search too. Attachment rows
    /// stay out of this shape: a thread member is a message, and admitting the attachment
    /// rows of its own members would return each of them twice.
    /// </summary>
    [Theory]
    [InlineData("calendar")]
    [InlineData("email")]
    [InlineData("document")]
    public void MessagesOnly_KeepsEveryMessageRow_WhateverItsKind(string kind)
    {
        Assert.True(IndexRowFilter.Keep(Hit(MessageUrl, kind), KindFilter.MessagesOnly));
    }

    [Fact]
    public void MessagesOnly_KeepsAMessageRowWithNoKindAtAll()
    {
        // A kind test needs the column to SAY email, so a row whose System.Kind did not come
        // back was dropped. A message with an unreadable kind column is still a message.
        Assert.True(IndexRowFilter.Keep(Hit(MessageUrl), KindFilter.MessagesOnly));
        Assert.False(IndexRowFilter.Keep(Hit(MessageUrl), KindFilter.MailKindOnly));
    }

    [Fact]
    public void MessagesOnly_StillRejectsAttachmentRowsAndTheFileSystem()
    {
        Assert.False(IndexRowFilter.Keep(Hit(AttachmentUrl, "picture"), KindFilter.MessagesOnly));

        // Without a kind test the mapi-namespace check is the ONLY thing keeping the file
        // system out of this filter, so it is pinned here rather than assumed.
        Assert.False(IndexRowFilter.Keep(Hit("file:C:/mail/archive.eml", "email"), KindFilter.MessagesOnly));
        Assert.False(IndexRowFilter.Keep(Hit("file:C:/notes/agenda.ics", "calendar"), KindFilter.MessagesOnly));
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
        Assert.Throws<ArgumentNullException>(() => IndexRowFilter.Keep(null!, KindFilter.MessagesAndAttachments));
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
