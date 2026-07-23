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
[Collection("LivePhase2")]
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
            HitSummary? firstFind = null;
            long firstFindMs = -1;
            string? firstFindSource = null;
            DateTime? stalenessAtFirstFind = null;
            List<HitSummary> bothCopies = new();

            while (overall.Elapsed < ArrivalTimeout)
            {
                SearchOutcome outcome = _fixture.Service.Search(new SearchRequest
                {
                    Query = marker,
                    Store = hub,
                    Mode = SearchMode.Fresh,
                    Top = 10,
                });

                List<HitSummary> matches = outcome.Hits
                    .Where(h => h.Subject != null && h.Subject.Contains(marker, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                if (matches.Count > 0 && firstFind == null)
                {
                    firstFind = matches[0];
                    firstFindMs = overall.ElapsedMilliseconds;
                    firstFindSource = matches[0].Source;
                    stalenessAtFirstFind = outcome.Staleness.NewestIndexedUtc;
                    _output.WriteLine($"first find after {firstFindMs} ms: source={firstFindSource} folderKind={matches[0].FolderKind} "
                        + $"sweepPerformed={outcome.Sweep?.Performed} sweepMs={outcome.Sweep?.ElapsedMs} indexNewest={outcome.Staleness.NewestIndexedUtc:O}");
                }

                // The SENT copy is visible immediately; the INBOX copy proves arrival.
                bothCopies = matches
                    .GroupBy(m => (m.Folder ?? string.Empty) + "|" + (m.FolderKind ?? string.Empty), StringComparer.OrdinalIgnoreCase)
                    .Select(g => g.First())
                    .ToList();
                bool inboxSeen = bothCopies.Any(m =>
                    string.Equals(m.FolderKind, "inbox", StringComparison.OrdinalIgnoreCase)
                    || (m.Source == "index" && m.FolderKind == null && !string.Equals(m.FolderKind, "sent", StringComparison.OrdinalIgnoreCase)
                        && bothCopies.Count >= 2));
                if (firstFind != null && inboxSeen && bothCopies.Count >= 2)
                {
                    break;
                }

                Thread.Sleep(PollInterval);
            }

            overall.Stop();
            Assert.NotNull(firstFind);
            _output.WriteLine($"copies seen: {bothCopies.Count} after {overall.ElapsedMilliseconds} ms "
                + $"({string.Join(", ", bothCopies.Select(c => c.FolderKind ?? "index:" + (c.Folder ?? "?")))})");

            // The fresh-mode claim: found via the LIVE sweep before indexing, or - if the
            // index genuinely raced ahead - the index frontier must already cover the send.
            bool sweptLive = firstFindSource == "live";
            bool indexCaughtUp = stalenessAtFirstFind.HasValue && stalenessAtFirstFind.Value >= sentAtUtc.AddSeconds(-30);
            Assert.True(sweptLive || indexCaughtUp,
                $"hit came from '{firstFindSource}' while index frontier {stalenessAtFirstFind:O} predates the send {sentAtUtc:O}");
            _output.WriteLine($"fresh-mode proof: sweptLive={sweptLive} indexCaughtUp={indexCaughtUp} firstFindMs={firstFindMs}");

            // Arrival: an inbox-side copy must exist within the timeout.
            HitSummary inboxCopy = bothCopies.FirstOrDefault(m =>
                    string.Equals(m.FolderKind, "inbox", StringComparison.OrdinalIgnoreCase))
                ?? bothCopies.FirstOrDefault(m => m.FolderKind == null)
                ?? bothCopies[0];
            Assert.True(bothCopies.Count >= 2 || string.Equals(inboxCopy.FolderKind, "inbox", StringComparison.OrdinalIgnoreCase),
                "expected the inbox copy (arrival) plus the sent copy within the timeout");

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
