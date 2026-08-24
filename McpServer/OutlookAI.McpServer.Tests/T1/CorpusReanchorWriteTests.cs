using System;
using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the two things that let a re-anchor destroy the corpus it exists to repair.
/// <para>
/// <b>What happened, measured on the test VM 2026-08-24.</b> <c>corpus-reanchor --execute</c>
/// ran over 20,000 items and reported <c>rewritten 20,000, already correct 0, refused by rule
/// 0, already gone 0, failed 0</c>. Afterwards every sampled item in Inbox and Sent Items was
/// dated inside the six minutes the tool had been running: the age-band structure the corpus
/// exists for - last-24h, 1d-7d, 7d-60d, 60d-1y, 1y-4y - had become "everything arrived while
/// the tool ran". The VM had to be restored from a checkpoint.
/// </para>
/// <para>
/// Neither failure was exotic. The write path ALREADY read each item's date back after
/// writing it, and threw that read-back away except to record it into the manifest as though
/// it were the intention - so the manifest faithfully recorded the wrong answer too. And the
/// replacement manifest lines left <c>FolderId</c> and <c>BodyBytes</c> at zero on the
/// reasoning that a re-anchor does not know them; the manifest reader is last-writer-wins per
/// ordinal WHOLESALE (<c>_items[ordinal] = item</c>), so that deleted what the build recorded.
/// </para>
/// <para>
/// <b>Why a mutation score of 39/39 did not catch it.</b> The pass that shipped this ran 39
/// mutations and killed all 39 - every one against the pure decision logic: shift derivation,
/// the agreement floor, the ordinal and allowlist guards. Nothing pinned the value actually
/// written to an item, because that path needs COM. A perfect score bounded to the half of the
/// code CI can reach says nothing about the half it cannot. These tests move the two decisions
/// into that reachable half.
/// </para>
/// <para>Pure: no Outlook, no COM, no mailbox, no settings file.</para>
/// </summary>
public sealed class CorpusReanchorWriteTests
{
    private static readonly DateTime Intended = new(2026, 8, 23, 23, 59, 51, DateTimeKind.Utc);

    [Fact]
    public void AWriteThatLanded_IsAccepted()
    {
        Assert.True(CorpusReanchor.WriteLanded(Intended, Intended));
    }

    [Fact]
    public void AWriteInsideTheSecondTolerance_IsAccepted()
    {
        // Outlook stores delivery times to the second and the manifest renders them to the
        // second, so a sub-second difference is rendering, not a failed write.
        Assert.True(CorpusReanchor.WriteLanded(Intended, Intended.AddMilliseconds(400)));
        Assert.True(CorpusReanchor.WriteLanded(Intended, Intended.AddMilliseconds(-400)));
    }

    [Fact]
    public void TheStoreReportingNow_InsteadOfTheIntendedInstant_IsRefused()
    {
        // THE DEFECT ITSELF. The store re-stamped the item with the moment of the write; the
        // old code accepted that silently, for twenty thousand items in a row.
        DateTime whenTheToolRan = new(2026, 8, 24, 0, 3, 12, DateTimeKind.Utc);
        Assert.False(CorpusReanchor.WriteLanded(Intended, whenTheToolRan));
    }

    [Fact]
    public void AnUnreadableDate_IsRefusedRatherThanAssumedGood()
    {
        // Null is "the item would not tell us", which is not evidence that the write worked.
        Assert.False(CorpusReanchor.WriteLanded(Intended, null));
    }

    [Fact]
    public void ADriftOfSecondsBeyondTolerance_IsRefused()
    {
        Assert.False(CorpusReanchor.WriteLanded(Intended, Intended.AddSeconds(5)));
    }

    [Fact]
    public void TheRefusalNamesBothInstants_BecauseTheDifferenceIsTheDiagnosis()
    {
        DateTime got = new(2026, 8, 24, 0, 3, 12, DateTimeKind.Utc);
        string message = CorpusReanchor.DescribeWriteRefusal(1, Intended, got);

        Assert.Contains("2026-08-23T23:59:51Z", message, StringComparison.Ordinal);
        Assert.Contains("2026-08-24T00:03:12Z", message, StringComparison.Ordinal);
        Assert.Contains("Item 1", message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheRefusalSaysItIsStopping_NotThatOneItemWasSkipped()
    {
        // The distinction that matters: this must abort the run. A message that reads like a
        // per-item skip invites exactly the "rewritten 20,000, failed 0" outcome.
        string message = CorpusReanchor.DescribeWriteRefusal(1, Intended, null);

        Assert.Contains("Stopping", message, StringComparison.Ordinal);
        Assert.Contains("unreadable", message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheReplacementLine_KeepsTheFolderAndBodySize_RatherThanZeroingThem()
    {
        // THE SECOND DEFECT. The manifest reader is last-writer-wins per ordinal WHOLESALE, so
        // a replacement line carrying zeroes deletes what the build recorded. Measured on the
        // VM: after a re-anchor, every entry read back as FolderId 0, BodyBytes 0.
        var item = new CorpusReanchorItem(3, "0000BEEF", Intended, Intended.AddMinutes(-2), 6, 4018);

        CorpusManifestItem line = CorpusReanchor.ReplacementLine(item, Intended, Intended);

        Assert.Equal(6, line.FolderId);
        Assert.Equal(4018, line.BodyBytes);
        Assert.Equal(3, line.Ordinal);
        Assert.Equal("0000BEEF", line.EntryId);
    }

    [Fact]
    public void TheReplacementLine_RecordsWhatTheStoreActuallyReports()
    {
        // What the store has is the truth worth recording - but only because the caller has
        // already refused to continue when it disagrees with the intention.
        DateTime readBack = Intended.AddMilliseconds(600);
        var item = new CorpusReanchorItem(4, "0000CAFE", Intended, Intended, 5, 900);

        CorpusManifestItem line = CorpusReanchor.ReplacementLine(item, Intended, readBack);

        Assert.Equal(CorpusManifest.FormatUtc(readBack), line.ReceivedUtc);
    }

    [Fact]
    public void TheReplacementLine_FallsBackToTheIntention_WhenTheStoreWillNotSay()
    {
        var item = new CorpusReanchorItem(5, "0000F00D", Intended, Intended, 5, 900);

        CorpusManifestItem line = CorpusReanchor.ReplacementLine(item, Intended, null);

        Assert.Equal(CorpusManifest.FormatUtc(Intended), line.ReceivedUtc);
    }

    [Fact]
    public void APlannedItem_CarriesTheFolderAndBodySizeTheManifestAlreadyRecorded()
    {
        // Without this the replacement line zeroes both, and the wholesale last-writer-wins
        // reader deletes what the build recorded.
        var item = new CorpusReanchorItem(
            Ordinal: 7,
            EntryId: "0000ABCD",
            ReceivedUtc: Intended,
            SentUtc: Intended.AddMinutes(-3),
            FolderId: 6,
            BodyBytes: 4018);

        Assert.Equal(6, item.FolderId);
        Assert.Equal(4018, item.BodyBytes);
    }
}
