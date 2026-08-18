using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Phase-7 payload-size review (v3.MD section 12 compact-payload discipline): PINS
/// every tool-facing cap so accidental cap creep fails the build, and covers the
/// capped list mappers (recipients/attachments) that keep read/draft/send payloads
/// bounded with has-more indicators.
/// </summary>
public sealed class PayloadDisciplineTests
{
    [Fact]
    public void Caps_ArePinned()
    {
        // Search hits: small pages, iterate instead of dumping (section 8/12).
        Assert.Equal(100, MailService.SearchTopCap);
        Assert.Equal(25, MailService.SearchTopDefault);
        Assert.Equal(1000, MailService.SnippetCharsCap);
        Assert.Equal(200, MailService.SnippetCharsDefault);

        // Thread members.
        Assert.Equal(200, MailService.ThreadTopCap);
        Assert.Equal(50, MailService.ThreadTopDefault);

        // Read: body/header text caps (flags: bodyTruncated/headersTruncated).
        Assert.Equal(500_000, MailService.BodyCharsCap);
        Assert.Equal(20_000, MailService.BodyCharsDefault);
        Assert.Equal(100_000, MailService.HtmlCharsDefault);
        Assert.True(MailService.HtmlCharsDefault <= MailService.BodyCharsCap, "the HTML budget may not exceed the hard body cap");
        Assert.Equal(256, MailService.HeaderCharsMin);
        Assert.Equal(65_536, MailService.HeaderCharsCap);
        Assert.Equal(8_192, MailService.HeaderCharsDefault);

        // List caps with has-more indicators (Phase-7 additions).
        Assert.Equal(100, MailService.RecipientsCap);
        Assert.Equal(100, MailService.AttachmentsCap);

        // Folder listing bounds (soak fix D37: full tree, offset-paged - no depth knob;
        // D38 raised the per-call page 500 -> 1000).
        Assert.Equal(1000, MailService.FoldersPerCallCap);
        Assert.Equal(10_000, MailService.FolderWalkAbsoluteCap);

        // Read body paging (soak fix D37: window served from the per-process body cache).
        Assert.Equal(8, BodyCache.MaxEntries);
        Assert.Equal(8_000_000, BodyCache.MaxTotalChars);
        Assert.Equal(TimeSpan.FromMinutes(15), BodyCache.TimeToLive);

        // Freshness-sweep coverage reporting (soak fix 13): the folder list is a
        // legibility aid for narrow scopes, not a payload every search carries.
        Assert.Equal(12, MailService.SweptFolderListCap);
        Assert.Equal(40, OutlookComSession.MaxScopedSweepFolders);

        // The unindexed-store list, which until 2026-08-18 had no cap at all (Q7b) - in the
        // payload or in the advice sentence that joins the names into prose. Derived from
        // the swept-folder cap rather than a second 12; StalenessAndUnindexedStoreTests
        // pins the truncation flags and the sentence.
        Assert.Equal(12, MailService.UnindexedStoreListCap);
        Assert.Equal(MailService.SweptFolderListCap, MailService.UnindexedStoreListCap);

        // Distinct item classes the "not ordinary mail" advice names before trailing off
        // (gap B3). A cap in prose is still a cap; the per-hit itemClass fields are the
        // complete answer beside it.
        Assert.Equal(4, MailService.NonMailClassAdviceCap);

        // The sweep's per-folder item cap. Pinned here because it was the one cap in the
        // service that was private and covered by no test at all, which made it the only
        // cap whose creep nothing would have noticed - and raising it silently is a COM
        // cost change, while lowering it silently hides just-arrived mail.
        Assert.Equal(200, MailService.SweepPerFolderCap);

        // Longest accepted subject. Three copies of this literal in MailService plus a
        // fourth as tool prose; see BudgetCompositionTests for the prose half.
        Assert.Equal(255, MailService.SubjectCharsCap);

        // Longest query show_search_results puts in Outlook's search box. Outlook's limit,
        // not ours; the rejection message quotes the constant rather than restating it.
        Assert.Equal(256, MailService.ShowSearchQueryCharsCap);

        // The two COM locate-probe bounds. Neither said what it counted before it was named:
        // the first is ITEMS IN A FOLDER (a subjectless hit can only be found by walking the
        // folder), the second is TABLE ROWS READ in the GetTable fallback.
        Assert.Equal(1000, OutlookComSession.TimeOnlyProbeMaxFolderItems);
        Assert.Equal(500, OutlookComSession.GetTableProbeMaxRows);

        // How far back the sweep looks when there is NO index frontier to look up to. Not a
        // payload cap - it is the reachable history of an unindexed store, and the whole
        // content of the no_index_frontier advice, so lowering it silently shrinks what such
        // a search can find and raising it silently makes every such search slower.
        Assert.Equal(TimeSpan.FromDays(7), MailService.EmptyIndexSweepWindow);

        // Budget for the post-sweep "is this store in the index at all" probes. They refine
        // what the answer SAYS, never what it contains, so this stays small.
        Assert.Equal(1_500, MailService.StoreIndexProbeBudgetMs);
    }

    /// <summary>
    /// The two index-staleness thresholds are ONE decision with two wordings, so they are
    /// pinned together and in order. The 30-minute one only says "the index is a bit behind
    /// and the sweep covers it"; the 12-hour one tells the agent to search differently. If
    /// the notice threshold ever rose above the advice threshold, outlook_health would call
    /// an index "current" that search is simultaneously calling very stale.
    /// </summary>
    [Fact]
    public void IndexStalenessThresholds_ArePinnedAndOrdered()
    {
        Assert.Equal(30, MailService.StaleIndexNoticeMinutes);
        Assert.Equal(720, MailService.VeryStaleAdviceMinutes);
        Assert.True(
            MailService.StaleIndexNoticeMinutes < MailService.VeryStaleAdviceMinutes,
            $"the notice threshold ({MailService.StaleIndexNoticeMinutes} min) must stay below the "
            + $"very-stale advice threshold ({MailService.VeryStaleAdviceMinutes} min): the first says the "
            + "freshness sweep is handling the gap, the second says stop trusting the index.");
    }

    [Fact]
    public void SweepCoverage_DefaultFolderSet_IsPinnedAndSelfDescribing()
    {
        // The default set is the freshness contract of every non-folder-scoped search
        // (soak fix 13): the four folders mail lands in without user action. Widening
        // it is a decision (cost: ~10 ms per folder per store), not a refactor - and
        // the scope string agents read must keep naming exactly these folders.
        Assert.Equal(
            new[] { "inbox", "sent", "deleted", "junk" },
            OutlookComSession.DefaultSweepFolderKinds);

        Assert.Equal(
            "default folders (Inbox, Sent Items, Deleted Items, Junk Email)",
            MailService.DefaultSweepScopeDescription);
    }

    [Fact]
    public void PageFolders_UnderCap_SinglePage_WithTotal()
    {
        var walk = MakeFolders(7);

        FoldersOutcome outcome = MailService.PageFolders(walk, offset: 0);

        Assert.Equal(7, outcome.FolderTotal);
        Assert.False(outcome.Truncated);
        Assert.Null(outcome.NextOffset);
        Assert.Null(outcome.Offset);
        Assert.Equal(7, outcome.Stores.Sum(s => s.Folders.Count));
    }

    [Fact]
    public void PageFolders_OverCap_TruncatesWithNextOffset()
    {
        var walk = MakeFolders(MailService.FoldersPerCallCap + 3);

        FoldersOutcome first = MailService.PageFolders(walk, offset: 0);
        Assert.Equal(MailService.FoldersPerCallCap, first.Stores.Sum(s => s.Folders.Count));
        Assert.True(first.Truncated);
        Assert.Equal(MailService.FoldersPerCallCap, first.NextOffset);
        Assert.Equal(walk.Count, first.FolderTotal);

        FoldersOutcome second = MailService.PageFolders(walk, offset: first.NextOffset!.Value);
        Assert.Equal(3, second.Stores.Sum(s => s.Folders.Count));
        Assert.False(second.Truncated);
        Assert.Null(second.NextOffset);
        Assert.Equal(first.NextOffset, second.Offset);

        // The two pages tile the walk exactly - no folder lost or repeated.
        var pagedPaths = first.Stores.SelectMany(s => s.Folders).Select(f => f.Path)
            .Concat(second.Stores.SelectMany(s => s.Folders).Select(f => f.Path))
            .ToList();
        Assert.Equal(walk.Select(f => f.Path), pagedPaths);
    }

    [Fact]
    public void PageFolders_OffsetBeyondEnd_ReturnsEmptyNotTruncated()
    {
        var walk = MakeFolders(5);

        FoldersOutcome outcome = MailService.PageFolders(walk, offset: 99);

        Assert.Empty(outcome.Stores);
        Assert.Equal(5, outcome.FolderTotal);
        Assert.False(outcome.Truncated);
        Assert.Null(outcome.NextOffset);
        Assert.Equal(99, outcome.Offset);
    }

    [Fact]
    public void PageFolders_NegativeOffset_IsClampedToZero()
    {
        var walk = MakeFolders(2);

        FoldersOutcome outcome = MailService.PageFolders(walk, offset: -5);

        Assert.Equal(2, outcome.Stores.Sum(s => s.Folders.Count));
        Assert.Null(outcome.Offset);
    }

    private static List<ComFolderInfo> MakeFolders(int count)
    {
        var list = new List<ComFolderInfo>(count);
        for (int i = 1; i <= count; i++)
        {
            // Two stores to prove grouping survives page slicing.
            string store = i <= count / 2 ? "Store A" : "Store B";
            list.Add(new ComFolderInfo(store, $"F{i:D5}", $"F{i:D5}", i, 0, 0));
        }

        return list;
    }

    [Fact]
    public void SearchRequest_Defaults_AreCompact()
    {
        SearchRequest request = new();

        Assert.Equal(MailService.SearchTopDefault, request.Top);
        Assert.Equal(MailService.SnippetCharsDefault, request.SnippetChars);
        // D34: fresh is THE behavior - no mode field exists; exhaustive is an opt-in
        // boolean and IndexOnly is a Core-only test escape hatch, both off by default.
        Assert.False(request.Exhaustive);
        Assert.False(request.IndexOnly);
        Assert.True(request.IncludeAttachmentHits);
    }

    [Fact]
    public void CapRecipients_UnderCap_ReturnsAll_NotTruncated()
    {
        var recipients = MakeRecipients(3);

        var views = MailService.CapRecipients(recipients, out int total, out bool truncated);

        Assert.Equal(3, views.Count);
        Assert.Equal(3, total);
        Assert.False(truncated);
        Assert.Equal("to", views[0].Kind);
        Assert.Equal("r1@example.com", views[0].Address);
    }

    [Fact]
    public void CapRecipients_OverCap_CapsWithHasMore()
    {
        var recipients = MakeRecipients(MailService.RecipientsCap + 57);

        var views = MailService.CapRecipients(recipients, out int total, out bool truncated);

        Assert.Equal(MailService.RecipientsCap, views.Count);
        Assert.Equal(MailService.RecipientsCap + 57, total);
        Assert.True(truncated);
        // Order preserved: the first cap-many recipients are listed.
        Assert.Equal("r1@example.com", views[0].Address);
        Assert.Equal($"r{MailService.RecipientsCap}@example.com", views[^1].Address);
    }

    [Fact]
    public void CapRecipients_ExactlyCap_NotTruncated()
    {
        var recipients = MakeRecipients(MailService.RecipientsCap);

        var views = MailService.CapRecipients(recipients, out int total, out bool truncated);

        Assert.Equal(MailService.RecipientsCap, views.Count);
        Assert.Equal(MailService.RecipientsCap, total);
        Assert.False(truncated);
    }

    [Fact]
    public void CapAttachments_OverCap_CapsAndKeepsOriginalIndexes()
    {
        var attachments = new List<ComAttachmentInfo>();
        for (int i = 1; i <= MailService.AttachmentsCap + 5; i++)
        {
            attachments.Add(new ComAttachmentInfo(i, $"file{i}.txt", i * 10));
        }

        var views = MailService.CapAttachments(attachments, out int total, out bool truncated);

        Assert.Equal(MailService.AttachmentsCap, views.Count);
        Assert.Equal(MailService.AttachmentsCap + 5, total);
        Assert.True(truncated);
        // 1-based ORIGINAL indexes survive the cap - save_attachment still addresses
        // unlisted attachments by index.
        Assert.Equal(1, views[0].Index);
        Assert.Equal(MailService.AttachmentsCap, views[^1].Index);
    }

    [Fact]
    public void CapAttachments_UnderCap_ReturnsAll_NotTruncated()
    {
        var attachments = new List<ComAttachmentInfo> { new(1, "a.pdf", 100), new(2, "b.pdf", null) };

        var views = MailService.CapAttachments(attachments, out int total, out bool truncated);

        Assert.Equal(2, views.Count);
        Assert.Equal(2, total);
        Assert.False(truncated);
        Assert.Null(views[1].SizeBytes);
    }

    private static List<ComRecipientInfo> MakeRecipients(int count)
    {
        var list = new List<ComRecipientInfo>(count);
        for (int i = 1; i <= count; i++)
        {
            list.Add(new ComRecipientInfo("to", $"Recipient {i}", $"r{i}@example.com"));
        }

        return list;
    }
}
