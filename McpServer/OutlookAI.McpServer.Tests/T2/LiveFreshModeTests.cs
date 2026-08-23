using System.Diagnostics;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Fresh-mode proof (v3.MD section 0.6 Phase 2 acceptance, D19/D20): a mail sent
/// telefonie-to-telefonie via COM must be findable through search(fresh) within
/// seconds - even before the Windows Search index has caught up - because the COM
/// gap-sweep covers items newer than the index frontier. The round trip continues into
/// read + save_attachment + content grep, and every artifact is deleted afterwards
/// (S3: tag + unique marker double-match, this run's items only).
/// </summary>
[Collection(LiveCollections.Phase2)]
[Trait("Category", "Live")]
public sealed class LiveFreshModeTests
{
    private static readonly TimeSpan ArrivalTimeout = TimeSpan.FromSeconds(120);
    private static readonly TimeSpan PollInterval = TimeSpan.FromSeconds(3);

    private readonly LivePhase2Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveFreshModeTests(LivePhase2Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    [Trait("Requires", "Transport")]
    public void FreshSearch_FindsSelfSentMail_BeforeIndexCatchesUp_ThenCleansUp()
    {
        string hub = _fixture.Settings.TestHubStoreDisplayName;
        string marker = "oaimcp" + Guid.NewGuid().ToString("N");
        string subject = $"{LiveOutlookTestMailer.SubjectTag} fresh-proof {marker}";
        string attachmentPath = Path.Combine(Path.GetTempPath(), $"OutlookAI-fresh-{marker}.txt");
        File.WriteAllText(attachmentPath, $"OutlookAI Phase-2 fresh-mode attachment payload. marker={marker}\n");

        DateTime sentAtUtc = DateTime.MinValue;
        try
        {
            sentAtUtc = LiveOutlookTestMailer.SendSelfMail(
                hub,
                subject,
                $"OutlookAI Phase-2 fresh-mode proof body. marker={marker}",
                attachmentPath);
            _output.WriteLine($"sent at {sentAtUtc:O}");

            Stopwatch overall = Stopwatch.StartNew();
            HitSummary? inboxFind = null;
            long inboxFindMs = -1;
            string? inboxFindSource = null;
            DateTime? stalenessAtInboxFind = null;
            bool sentCopySeen = false;
            int lastLoggedMatches = -1;

            while (overall.Elapsed < ArrivalTimeout)
            {
                // D34: the sweep cache would blind rapid re-polls for its ~10 s TTL by
                // design - this test measures RAW sweep arrival latency, so it clears
                // the cache before each poll (dedicated cache behavior test:
                // LiveSweepCacheTests).
                _fixture.Service.ClearSweepCache();
                SearchOutcome outcome = _fixture.Service.Search(new SearchRequest
                {
                    Query = marker,
                    Store = hub,
                    Top = 10,
                });

                List<HitSummary> matches = outcome.Hits
                    .Where(h => h.Subject != null && h.Subject.Contains(marker, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (matches.Count != lastLoggedMatches)
                {
                    lastLoggedMatches = matches.Count;
                    _output.WriteLine($"t+{overall.ElapsedMilliseconds}ms matches={matches.Count} "
                        + $"[{string.Join("; ", matches.Select(m => $"{m.Source}/{m.FolderKind ?? m.Folder}"))}] "
                        + $"sweep: performed={outcome.Sweep?.Performed} folders={outcome.Sweep?.FoldersSwept}/{outcome.Sweep?.FoldersSkipped} "
                        + $"seen={outcome.Sweep?.ItemsSeen} dups={outcome.Sweep?.Duplicates} err={outcome.Sweep?.Error ?? "-"} "
                        + $"gapStart={outcome.Sweep?.GapStartUtc:HH:mm:ss} indexNewest={outcome.Staleness.NewestIndexedUtc:HH:mm:ss}");
                }

                sentCopySeen |= matches.Any(m =>
                    string.Equals(m.FolderKind, "sent", StringComparison.OrdinalIgnoreCase));

                // ACCEPTANCE (doc row): the ARRIVED (inbox) copy is findable within
                // seconds, even before the index catches up.
                HitSummary? inboxMatch = matches.FirstOrDefault(m =>
                    string.Equals(m.FolderKind, "inbox", StringComparison.OrdinalIgnoreCase)
                    || (m.Source == "index"
                        && !ContainsOutboundFolder(m.Folder)
                        && !ContainsOutboundFolder(m.FolderKind)));
                if (inboxMatch != null && inboxFind == null)
                {
                    inboxFind = inboxMatch;
                    inboxFindMs = overall.ElapsedMilliseconds;
                    inboxFindSource = inboxMatch.Source;
                    stalenessAtInboxFind = outcome.Staleness.NewestIndexedUtc;
                    break;
                }

                Thread.Sleep(PollInterval);
            }

            overall.Stop();
            Assert.NotNull(inboxFind);
            _output.WriteLine($"inbox copy found after {inboxFindMs} ms: source={inboxFindSource} "
                + $"indexNewestAtFind={stalenessAtInboxFind:O} sentCopySeen={sentCopySeen}");

            // The fresh-mode claim: the arrival was served by the LIVE sweep before
            // indexing, or - if the index genuinely raced ahead - its frontier must
            // already cover the send instant.
            bool sweptLive = inboxFindSource == "live";
            bool indexCaughtUp = stalenessAtInboxFind.HasValue && stalenessAtInboxFind.Value >= sentAtUtc;
            Assert.True(sweptLive || indexCaughtUp,
                $"hit came from '{inboxFindSource}' while index frontier {stalenessAtInboxFind:O} predates the send {sentAtUtc:O}");
            _output.WriteLine($"fresh-mode proof: sweptLive={sweptLive} indexCaughtUp={indexCaughtUp} inboxFindMs={inboxFindMs}");

            HitSummary inboxCopy = inboxFind!;

            // Round trip continues: read the arrived mail, save its attachment, grep it.
            ReadOutcome read = _fixture.Service.Read(inboxCopy.Id, maxBodyChars: 5000);
            Assert.Contains(marker, read.Body, StringComparison.OrdinalIgnoreCase);
            Assert.True(read.Attachments.Count >= 1, "the test mail carries one attachment");
            _output.WriteLine($"read ok: folder set={read.Folder != null} attachments={read.Attachments.Count} bodyChars={read.BodyTotalChars}");

            AttachmentView attachment = read.Attachments.First(a =>
                a.FileName != null && a.FileName.Contains("OutlookAI-fresh", StringComparison.OrdinalIgnoreCase));
            SaveAttachmentOutcome saved = _fixture.Service.SaveAttachment(inboxCopy.Id, attachment.Index);
            try
            {
                Assert.True(File.Exists(saved.SavedPath));
                string content = File.ReadAllText(saved.SavedPath);
                Assert.Contains(marker, content, StringComparison.OrdinalIgnoreCase);
                _output.WriteLine($"attachment saved+grepped: bytes={saved.SizeBytes}");
            }
            finally
            {
                TryDelete(saved.SavedPath);
            }
        }
        finally
        {
            TryDelete(attachmentPath);

            // S3 cleanup: delete ONLY this run's artifacts (tag + marker double match)
            // from Inbox, Sent Items and Deleted Items of the test hub.
            if (sentAtUtc != DateTime.MinValue)
            {
                int deleted = TryCleanup(hub, marker);
                _output.WriteLine($"cleanup deleted {deleted} artifacts (expect >= 2: sent + inbox copies)");
                Assert.True(deleted >= 2, $"expected to delete both self-send copies, deleted {deleted}");
            }
        }
    }

    private static bool ContainsOutboundFolder(string? folder)
    {
        if (folder == null)
        {
            return false;
        }

        // Localized OUTBOUND-folder names in this profile's languages: Sent Items AND
        // the Outbox. A self-send can linger in the Outbox long enough for the index to
        // pick that copy up (observed 2026-07-27); that row is the outgoing copy, never
        // the arrival, so it must not satisfy the inbox-arrival acceptance.
        return folder.Contains("Sent", StringComparison.OrdinalIgnoreCase)
            || folder.Contains("Verzonden", StringComparison.OrdinalIgnoreCase)
            || folder.Contains("Outbox", StringComparison.OrdinalIgnoreCase)
            || folder.Contains("Postvak UIT", StringComparison.OrdinalIgnoreCase);
    }

    private int TryCleanup(string hub, string marker)
    {
        // The inbox copy can lag the search visibility by a moment; retry briefly so the
        // run never leaves artifacts behind (S3).
        int total = 0;
        for (int attempt = 0; attempt < 10; attempt++)
        {
            total += LiveOutlookTestMailer.DeleteTaggedArtifacts(hub, marker);
            if (total >= 2)
            {
                // One more pass for the Deleted Items copies created by the deletes above.
                total += LiveOutlookTestMailer.DeleteTaggedArtifacts(hub, marker);
                return total;
            }

            Thread.Sleep(TimeSpan.FromSeconds(5));
        }

        return total;
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch (IOException)
        {
        }
        catch (UnauthorizedAccessException)
        {
        }
    }
}
