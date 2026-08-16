using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the incident-7 remediation decision rules (RemediationTools). The load-bearing
/// property is ORDINAL matching: the incident was caused by shell-side wildcard
/// matching where the tag's brackets became a character class - these tests prove the
/// sanctioned predicates treat every character literally, and that the delete-time
/// verification of the duplicate cleanup fails CLOSED (skip, never delete) on any
/// missing signal.
/// </summary>
public class RemediationRulesTests
{
    // ---------------------------------------------------------- purge folder sets

    [Fact]
    public void PurgeSweep_CoversTheOutbox()
    {
        // Regression guard, 2026-08-16. The Outbox was missing from both purge sets, and
        // nothing pinned them, so nothing failed. The live suite's own sweep covers the
        // Outbox but deletes only items in its per-run EntryID allowlist (the S3 guard),
        // so a tagged item left there by an EARLIER run was reachable by nothing at all.
        // One sat in the test hub's Outbox from 2026-07-30 to 2026-08-16.
        Assert.Contains(4, ComMailbox.SweepFolderIds);
        Assert.Contains(4, ComMailbox.HubSweepFolderIdsWithArchive);
    }

    [Fact]
    public void PurgeSweep_CoversEveryFolderATestCanWriteTo()
    {
        // Drafts(16), Inbox(6), Sent Items(5), Outbox(4), Deleted Items(3).
        foreach (int folderId in new[] { 16, 6, 5, 4, 3 })
        {
            Assert.Contains(folderId, ComMailbox.SweepFolderIds);
            Assert.Contains(folderId, ComMailbox.HubSweepFolderIdsWithArchive);
        }

        // The hub set additionally covers its designated Archive.
        Assert.Contains(39, ComMailbox.HubSweepFolderIdsWithArchive);
        Assert.DoesNotContain(39, ComMailbox.SweepFolderIds);
    }

    [Fact]
    public void PurgeSweep_VisitsDeletedItemsLast()
    {
        // Load-bearing ordering: deleting from any other folder SOFT-moves the item into
        // Deleted Items. Sweeping Deleted Items first would leave every one of those
        // moved copies behind, and the purge would never reach stable zero.
        Assert.Equal(3, ComMailbox.SweepFolderIds[^1]);
        Assert.Equal(3, ComMailbox.HubSweepFolderIdsWithArchive[^1]);
    }

    [Fact]
    public void PurgeSweep_ListsEachFolderOnce()
    {
        // A duplicated id would double-delete and, worse, make the stable-zero loop
        // disagree with itself about what it had already covered.
        Assert.Equal(ComMailbox.SweepFolderIds.Length, ComMailbox.SweepFolderIds.Distinct().Count());
        Assert.Equal(
            ComMailbox.HubSweepFolderIdsWithArchive.Length,
            ComMailbox.HubSweepFolderIdsWithArchive.Distinct().Count());
    }

    [Fact]
    public void IsTagged_MatchesFullTagOrdinally()
    {
        Assert.True(RemediationRules.IsTagged("[OutlookAI-McpTest] p5-something"));
        Assert.True(RemediationRules.IsTagged("RE: prefixed [OutlookAI-McpTest] mid-subject"));
        Assert.False(RemediationRules.IsTagged(null));
        Assert.False(RemediationRules.IsTagged(string.Empty));
        // The 7d bug class: subjects that MATCH a bracket wildcard-class pattern but
        // do NOT contain the literal tag must never match the ordinal predicate.
        Assert.False(RemediationRules.IsTagged("O")); // matched "[OutlookAI-McpTest]" as -like class
        Assert.False(RemediationRules.IsTagged("RE: Telefonie storing"));
        Assert.False(RemediationRules.IsTagged("OutlookAI-McpTest without brackets"));
        Assert.False(RemediationRules.IsTagged("[outlookai-mcptest] wrong case"));
    }

    [Fact]
    public void DaslFragment_IsBracketFreeSubsetOfTag()
    {
        // LIKE prefilter soundness: fragment must appear in the tag itself (superset
        // prefilter) and carry no DASL wildcard/quote characters.
        Assert.Contains(RemediationRules.DaslCountFragment, RemediationRules.SubjectTag, StringComparison.Ordinal);
        Assert.DoesNotContain('[', RemediationRules.DaslCountFragment);
        Assert.DoesNotContain(']', RemediationRules.DaslCountFragment);
        Assert.DoesNotContain('%', RemediationRules.DaslCountFragment);
        Assert.DoesNotContain('\'', RemediationRules.DaslCountFragment);
    }

    [Theory]
    [InlineData("hub@example.test", false, RemediationRules.TelefonieOrigin.SentOrigin)]
    [InlineData("HUB@EXAMPLE.TEST", false, RemediationRules.TelefonieOrigin.SentOrigin)]
    [InlineData("other@example.test", true, RemediationRules.TelefonieOrigin.InboxOrigin)]
    [InlineData(null, true, RemediationRules.TelefonieOrigin.InboxOrigin)]
    public void ClassifyOrigin_AgreeingSignalsClassify(string? sender, bool receivedBy, RemediationRules.TelefonieOrigin expected)
    {
        Assert.Equal(expected, RemediationRules.ClassifyOrigin(sender, receivedBy, "hub@example.test"));
    }

    [Theory]
    [InlineData("hub@example.test", true)]   // hub sender but received-by present
    [InlineData("other@example.test", false)] // foreign sender but no received-by
    [InlineData(null, false)]
    public void ClassifyOrigin_DisagreeingSignalsReturnNull(string? sender, bool receivedBy)
    {
        Assert.Null(RemediationRules.ClassifyOrigin(sender, receivedBy, "hub@example.test"));
    }

    [Fact]
    public void ParseDeletionLog_ReadsStoreFolderPrefix()
    {
        var entries = RemediationRules.ParseDeletionLog(new[]
        {
            "delete: store=hub@example.test folder=6 markerPrefix=RE:",
            "delete: store=hub@example.test folder=5 markerPrefix=Telefoni",
            string.Empty,
            "delete: store=someone@example.test folder=6 markerPrefix=Uw",
        });
        Assert.Equal(3, entries.Count);
        Assert.Equal(new RemediationRules.DeletionLogEntry("hub@example.test", 6, "RE:"), entries[0]);
        Assert.Equal(new[] { "Telefoni" }, RemediationRules.ExpectedPrefixes(entries, "hub@example.test", 5));
        Assert.Equal(new[] { "RE:" }, RemediationRules.ExpectedPrefixes(entries, "HUB@EXAMPLE.TEST", 6));
        Assert.Throws<FormatException>(() => RemediationRules.ParseDeletionLog(new[] { "delete: folder=6" }));
    }

    [Fact]
    public void TryConsumePrefixMatch_ConsumesOrdinalPrefixOnce()
    {
        var remaining = new List<string> { "RE:", "Telefoni", "test" };
        Assert.Equal("Telefoni", RemediationRules.TryConsumePrefixMatch(remaining, "Telefonie storing xyz"));
        Assert.Equal("RE:", RemediationRules.TryConsumePrefixMatch(remaining, "RE: something"));
        Assert.Null(RemediationRules.TryConsumePrefixMatch(remaining, "RE: second RE has no slot left"));
        Assert.Equal("test", RemediationRules.TryConsumePrefixMatch(remaining, "  test with leading spaces"));
        Assert.Empty(remaining);
        Assert.Null(RemediationRules.TryConsumePrefixMatch(remaining, null));
    }

    [Fact]
    public void DecideDuplicateDelete_FailsClosed()
    {
        var inbox = new HashSet<string>(StringComparer.Ordinal) { "<twin@id.example>" };
        // Only the fully verified case deletes.
        Assert.Equal(RemediationRules.DedupeDecision.Delete,
            RemediationRules.DecideDuplicateDelete("any subject", "<twin@id.example>", inbox));
        Assert.Equal(RemediationRules.DedupeDecision.Delete,
            RemediationRules.DecideDuplicateDelete("any subject", "  <twin@id.example>  ", inbox));
        // Every missing signal skips.
        Assert.Equal(RemediationRules.DedupeDecision.SkipTagged,
            RemediationRules.DecideDuplicateDelete("[OutlookAI-McpTest] x", "<twin@id.example>", inbox));
        Assert.Equal(RemediationRules.DedupeDecision.SkipEmptyMessageId,
            RemediationRules.DecideDuplicateDelete("s", null, inbox));
        Assert.Equal(RemediationRules.DedupeDecision.SkipEmptyMessageId,
            RemediationRules.DecideDuplicateDelete("s", "   ", inbox));
        Assert.Equal(RemediationRules.DedupeDecision.SkipNoInboxTwin,
            RemediationRules.DecideDuplicateDelete("s", "<absent@id.example>", inbox));
        // Message-ID identity is ordinal - no case folding.
        Assert.Equal(RemediationRules.DedupeDecision.SkipNoInboxTwin,
            RemediationRules.DecideDuplicateDelete("s", "<TWIN@ID.EXAMPLE>", inbox));
    }

    [Fact]
    public void NormalizeMessageId_TrimsToNull()
    {
        Assert.Equal("<a@b>", RemediationRules.NormalizeMessageId(" <a@b> "));
        Assert.Null(RemediationRules.NormalizeMessageId("   "));
        Assert.Null(RemediationRules.NormalizeMessageId(null));
    }
}
