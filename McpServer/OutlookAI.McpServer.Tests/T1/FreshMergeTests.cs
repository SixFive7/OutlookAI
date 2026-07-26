using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// T1: fresh-mode merge logic (v3.MD D19) - term re-application over gap-swept items
/// and boundary de-duplication against index hits. Index hits are fabricated through
/// the public IndexRowMapper (synthetic URLs, S6-safe).
/// </summary>
public sealed class FreshMergeTests
{
    private static readonly DateTime BaseUtc = new(2026, 07, 23, 10, 00, 00, DateTimeKind.Utc);

    private static ComMailBrief Brief(
        string subject = "Quarterly invoice",
        string store = "alice@example.com",
        string folder = "Inbox",
        DateTime? receivedLocal = null,
        string? body = null,
        string? senderName = null,
        string? senderAddress = null)
    {
        return new ComMailBrief(
            entryId: "AA" + Guid.NewGuid().ToString("N"),
            storeDisplayName: store,
            storeId: null,
            folderName: folder,
            folderKind: "inbox",
            subject: subject,
            senderName: senderName,
            senderAddress: senderAddress,
            receivedTime: receivedLocal ?? BaseUtc.ToLocalTime(),
            isRead: false,
            hasAttachments: false,
            sizeBytes: 1000,
            body: body);
    }

    private static IndexHit Hit(string subject = "Quarterly invoice", string store = "alice@example.com", string folder = "Inbox", DateTime? receivedUtc = null)
    {
        var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] =
                $"mapi16://{{S-1-5-21-1-2-3-1001}}/{store}($abcd1234)/0/{folder}/{EntryIdCodecTests.SyntheticEncodedTail()}",
            ["System.Subject"] = subject,
            ["System.Message.DateReceived"] = (receivedUtc ?? BaseUtc).ToUniversalTime(),
        };
        return IndexRowMapper.Map(row);
    }

    // ---------------------------------------------------------------- MatchesTerms

    private const SearchIn Default = SearchInValues.Default;

    [Fact]
    public void MatchesTerms_NullOrEmpty_MatchesEverything()
    {
        Assert.True(FreshMerge.MatchesTerms(Brief(), null, Default));
        Assert.True(FreshMerge.MatchesTerms(Brief(), Array.Empty<string>(), Default));
    }

    [Fact]
    public void MatchesTerms_AndSemantics_AllTermsMustHit()
    {
        ComMailBrief item = Brief(subject: "Invoice for April", body: "total 100 euro");
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "invoice", "euro" }, Default));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "invoice", "missingterm" }, Default));
    }

    [Fact]
    public void MatchesTerms_CaseInsensitive_AcrossSubjectAndBody()
    {
        ComMailBrief item = Brief(subject: "hello", body: "WORLD");
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "HELLO" }, Default));
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "world" }, Default));
    }

    [Fact]
    public void MatchesTerms_SenderIsNotMatchedByTerms_MatchingSenderAloneDoesNotHit()
    {
        // D40/SF-6 tier alignment: the index tier never matched senders by term, so the
        // sweep must not either - otherwise a hit would vanish once the frontier passed
        // the item. Sender matching is the 'from' filter's job (applied by MailService).
        ComMailBrief item = Brief(
            subject: "hello", body: "world", senderName: "Charlie", senderAddress: "c@example.com");
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "charlie" }, Default));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "c@example.com" }, Default));
    }

    [Fact]
    public void MatchesTerms_PrefixStar_MatchesStem()
    {
        ComMailBrief item = Brief(subject: "factuur 2026-001");
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "fact*" }, Default));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "xyz*" }, Default));
    }

    [Fact]
    public void MatchesTerms_NoBody_StillMatchesOnSubject()
    {
        ComMailBrief item = Brief(subject: "Order confirmation", body: null);
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "order" }, Default));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "invoice" }, Default));
    }

    // ------------------------------------------------- search_in scopes (D40, user 2026-07-26)

    [Fact]
    public void MatchesTerms_SubjectOnlyScope_IgnoresBody()
    {
        ComMailBrief item = Brief(subject: "alert prefix", body: "requeued backend");

        Assert.True(FreshMerge.MatchesTerms(item, new[] { "alert" }, SearchIn.SubjectOnly));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "requeued" }, SearchIn.SubjectOnly));
    }

    [Fact]
    public void MatchesTerms_BodyOnlyScope_IgnoresSubject()
    {
        ComMailBrief item = Brief(subject: "alert prefix", body: "requeued backend");

        Assert.True(FreshMerge.MatchesTerms(item, new[] { "requeued" }, SearchIn.BodyOnly));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "alert" }, SearchIn.BodyOnly));
    }

    [Fact]
    public void MatchesTerms_DefaultScope_FindsSubjectOnlyAndBodyOnlyTerms()
    {
        // The SF-6 shape in miniature: a term that lives only in the subject must be
        // found by the default scope (that is the whole point of the fix).
        ComMailBrief item = Brief(subject: "alert prefix", body: "requeued backend");

        Assert.True(FreshMerge.MatchesTerms(item, new[] { "alert" }, Default));
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "requeued" }, Default));
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "alert", "requeued" }, Default));
    }

    [Fact]
    public void MatchesTerms_ScopesHonorPrefixStems()
    {
        ComMailBrief item = Brief(subject: "factuur 2026-001", body: "betaling ontvangen");

        Assert.True(FreshMerge.MatchesTerms(item, new[] { "fact*" }, SearchIn.SubjectOnly));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "fact*" }, SearchIn.BodyOnly));
        Assert.True(FreshMerge.MatchesTerms(item, new[] { "betal*" }, SearchIn.BodyOnly));
        Assert.False(FreshMerge.MatchesTerms(item, new[] { "betal*" }, SearchIn.SubjectOnly));
    }

    // ---------------------------------------------------------------- IsDuplicate

    [Fact]
    public void IsDuplicate_SameStoreFolderSubjectAndTime_IsTrue()
    {
        Assert.True(FreshMerge.IsDuplicate(Brief(), Hit(), toleranceSeconds: 15));
    }

    [Fact]
    public void IsDuplicate_DifferentFolder_SentVsInboxCopy_IsFalse()
    {
        // A self-send: identical subject + near-identical time, but Sent Items vs Inbox
        // must remain two distinct hits.
        ComMailBrief sentCopy = Brief(folder: "Sent Items");
        Assert.False(FreshMerge.IsDuplicate(sentCopy, Hit(folder: "Inbox"), toleranceSeconds: 15));
    }

    [Fact]
    public void IsDuplicate_DifferentStore_IsFalse()
    {
        Assert.False(FreshMerge.IsDuplicate(Brief(store: "bob@example.com"), Hit(store: "alice@example.com"), 15));
    }

    [Fact]
    public void IsDuplicate_DifferentSubject_IsFalse()
    {
        Assert.False(FreshMerge.IsDuplicate(Brief(subject: "Other"), Hit(subject: "Quarterly invoice"), 15));
    }

    [Fact]
    public void IsDuplicate_TimeOutsideTolerance_IsFalse()
    {
        ComMailBrief item = Brief(receivedLocal: BaseUtc.AddMinutes(10).ToLocalTime());
        Assert.False(FreshMerge.IsDuplicate(item, Hit(receivedUtc: BaseUtc), toleranceSeconds: 15));
    }

    // ---------------------------------------------------------------- SelectFreshOnly

    [Fact]
    public void SelectFreshOnly_DropsIndexDuplicates_KeepsNewItems()
    {
        var swept = new List<ComMailBrief>
        {
            Brief(subject: "Quarterly invoice"),               // duplicate of the index hit
            Brief(subject: "Brand new mail", folder: "Inbox"), // genuinely fresh
        };
        var hits = new List<IndexHit> { Hit(subject: "Quarterly invoice") };

        IReadOnlyList<ComMailBrief> fresh = FreshMerge.SelectFreshOnly(swept, hits, 15, out int duplicates);

        Assert.Single(fresh);
        Assert.Equal("Brand new mail", fresh[0].Subject);
        Assert.Equal(1, duplicates);
    }

    [Fact]
    public void SelectFreshOnly_DropsRepeatedEntryIds()
    {
        ComMailBrief item = Brief(subject: "Once");
        var swept = new List<ComMailBrief> { item, item };

        IReadOnlyList<ComMailBrief> fresh = FreshMerge.SelectFreshOnly(swept, Array.Empty<IndexHit>(), 15, out int duplicates);

        Assert.Single(fresh);
        Assert.Equal(1, duplicates);
    }

    // ---------------------------------------------------------------- ResolveHitStore

    [Fact]
    public void ResolveHitStore_DelegateSubtree_UsesFirstFolderSegment()
    {
        var row = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = "mapi16://{S-1-5-21-1-2-3-1001}/owner@example.com($abcd1234)/1/Delegate Name/Postvak IN/"
                + EntryIdCodecTests.SyntheticEncodedTail(),
            ["System.Subject"] = "x",
        };
        IndexHit hit = IndexRowMapper.Map(row);
        Assert.Equal(1, hit.StoreType);
        Assert.Equal("Delegate Name", FreshMerge.ResolveHitStore(hit));
    }

    [Fact]
    public void ResolveHitStore_PrimaryStore_UsesStoreDisplayName()
    {
        Assert.Equal("alice@example.com", FreshMerge.ResolveHitStore(Hit()));
    }
}
