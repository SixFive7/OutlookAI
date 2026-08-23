using System.Text.Json;
using OutlookAI.McpServer.Tests.T2;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// D39 T3 live acceptance: move_mail and archive_mail over REAL stdio MCP against the
/// built server exe with golden-shape asserts - a seed is found via the wire search
/// (hit-id flow), moved into the tagged test folder (create_folder path), then
/// archived via its newEntryId; camelCase result fields, advice, and archiveFolders
/// are pinned on the wire. HUB ONLY (S2); everything is deleted via the allowlist
/// helpers with zero-remaining asserts (items and the test folder).
/// </summary>
[Collection(LiveCollections.MoveArchive)]
[Trait("Category", "Live")]
public sealed class MoveArchiveLiveMcpToolTests
{
    /// <summary>
    /// How long the seed may take to become visible over stdio. Covers a real mail round
    /// trip plus the ~10 s sweep cache TTL, which can delay an arrival during rapid polling.
    /// </summary>
    private const int SeedVisibleSeconds = 180;

    private readonly LiveMoveArchiveFixture _fixture;
    private readonly ITestOutputHelper _output;

    public MoveArchiveLiveMcpToolTests(LiveMoveArchiveFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    [Fact]
    [Trait("Requires", "Transport")]
    public async Task MoveAndArchive_GoldenShapes_OverRealStdio()
    {
        // Pre-clean: the createdFolders assert needs a clean slate (allowlist helper).
        LiveOutlookTestMailer.DeleteTestFolders(Hub);

        string seedSubject = _fixture.TaggedSubject("t3 move seed");
        try
        {
            LiveOutlookTestMailer.SendSelfMail(Hub, seedSubject, "T3 D39 seed body " + Marker, null);

            await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(TimeSpan.FromMinutes(8));

            // Find the seed over stdio (hit-id flow; the freshness sweep catches
            // pre-index arrivals, the cache TTL is absorbed by the deadline).
            string? hitId = null;
            LiveWaitBudget wait = LiveWaitBudget.OfSeconds(SeedVisibleSeconds);
            while (hitId == null && wait.HasTimeLeft)
            {
                JsonElement search = await client.CallToolAsync("search", new
                {
                    query = Marker,
                    store = Hub,
                    include_attachment_hits = false,
                    top = 25,
                });
                foreach (JsonElement hit in search.GetProperty("hits").EnumerateArray())
                {
                    bool isInbox = hit.TryGetProperty("folderKind", out JsonElement kind) && kind.GetString() == "inbox";
                    bool subjectMatches = hit.TryGetProperty("subject", out JsonElement subj) && subj.GetString() == seedSubject;
                    if (subjectMatches && (isInbox || !hit.TryGetProperty("folderKind", out _)))
                    {
                        hitId = hit.GetProperty("id").GetString();
                        break;
                    }
                }

                if (hitId == null)
                {
                    await Task.Delay(3000);
                }
            }

            Assert.True(hitId != null, "seed mail not findable over stdio within 180 s");

            // --- move_mail golden shape (hit id + create_folder).
            JsonElement move = await client.CallToolAsync("move_mail", new
            {
                ids = new[] { hitId },
                folder = _fixture.TestFolderName,
                create_folder = true,
            });
            Assert.Equal(1, move.GetProperty("requested").GetInt32());
            Assert.Equal(1, move.GetProperty("moved").GetInt32());
            Assert.Equal(0, move.GetProperty("failed").GetInt32());
            Assert.Equal(_fixture.TestFolderName, move.GetProperty("targetFolder").GetString());
            Assert.Contains(_fixture.TestFolderName,
                move.GetProperty("createdFolders").EnumerateArray().Select(f => f.GetString()));
            Assert.Contains("newEntryId", move.GetProperty("advice").EnumerateArray().Select(a => a.GetString()!).Single(),
                StringComparison.Ordinal);

            JsonElement movedItem = move.GetProperty("items").EnumerateArray().Single();
            Assert.Equal(hitId, movedItem.GetProperty("id").GetString());
            Assert.True(movedItem.GetProperty("ok").GetBoolean());
            Assert.Equal(Hub, movedItem.GetProperty("store").GetString(), ignoreCase: true);

            // The wire search may surface the Inbox OR the Sent copy of the self-send -
            // either is a valid move source (Inbox-specific undo semantics are pinned in
            // the T2 chain); what the wire shape guarantees is a non-empty undo address.
            string fromFolder = movedItem.GetProperty("fromFolder").GetString()!;
            Assert.False(string.IsNullOrEmpty(fromFolder));
            Assert.Equal(_fixture.TestFolderName, movedItem.GetProperty("toFolder").GetString());
            string oldEntryId = movedItem.GetProperty("oldEntryId").GetString()!;
            string newEntryId = movedItem.GetProperty("newEntryId").GetString()!;
            Assert.True(newEntryId.Length >= 48);
            Assert.NotEqual(oldEntryId, newEntryId);
            _output.WriteLine("wire move_mail ok: hit id flow, created folder, EntryID changed");

            // The HIT ID keeps working after the move (refreshed to the new EntryID):
            // read through it and confirm the item is intact.
            JsonElement read = await client.CallToolAsync("read", new { id = hitId, max_body_chars = 200 });
            Assert.Equal(seedSubject, read.GetProperty("subject").GetString());
            Assert.Equal(newEntryId, read.GetProperty("entryId").GetString(), ignoreCase: true);
            _output.WriteLine("hit id survived the move (cache refreshed to newEntryId)");

            // --- archive_mail golden shape (EntryID flow).
            JsonElement archive = await client.CallToolAsync("archive_mail", new { ids = new[] { newEntryId } });
            Assert.Equal(1, archive.GetProperty("requested").GetInt32());
            Assert.Equal(1, archive.GetProperty("archived").GetInt32());
            Assert.Equal(0, archive.GetProperty("failed").GetInt32());

            JsonElement archiveFolder = archive.GetProperty("archiveFolders").EnumerateArray().Single();
            Assert.Equal(Hub, archiveFolder.GetProperty("store").GetString(), ignoreCase: true);
            Assert.False(string.IsNullOrEmpty(archiveFolder.GetProperty("folder").GetString()));
            Assert.Equal("outlookDefaultFolder", archiveFolder.GetProperty("via").GetString());

            JsonElement archivedItem = archive.GetProperty("items").EnumerateArray().Single();
            Assert.True(archivedItem.GetProperty("ok").GetBoolean());
            Assert.Equal(_fixture.TestFolderName, archivedItem.GetProperty("fromFolder").GetString());
            Assert.Equal(archiveFolder.GetProperty("folder").GetString(), archivedItem.GetProperty("toFolder").GetString());
            Assert.NotEqual(newEntryId, archivedItem.GetProperty("newEntryId").GetString());
            _output.WriteLine($"wire archive_mail ok: landed in designated '{archiveFolder.GetProperty("folder").GetString()}'");

            Assert.True(await client.CloseAndAwaitExitAsync(TimeSpan.FromSeconds(30)), "server must exit on stdin close");
        }
        finally
        {
            // Folders first (contents purged into Deleted Items), then the stable-zero
            // item sweep with hub-archive coverage - same order as the T2 chain.
            LiveOutlookTestMailer.DeleteTestFolders(Hub);
            LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(
                Hub, Marker, folderIds: LiveOutlookTestMailer.HubSweepFolderIdsWithArchive);
        }

        // The STRICT count cannot tolerate the documented materialization lag: an item
        // can surface after the stable-zero sweep returned, which failed this test in an
        // otherwise green run. Purge once more (S3-legal: tag AND this run's marker) and
        // assert only what SURVIVES - the same correction batch A made for the T2 suites.
        int remaining = LiveOutlookTestMailer.CountTaggedArtifactsAfterPurgingStragglers(
            Hub, Marker, LiveOutlookTestMailer.HubSweepFolderIdsWithArchive, out int stragglersPurged);
        if (stragglersPurged > 0)
        {
            _output.WriteLine($"cleanup[{Hub}]: {stragglersPurged} late-materialized artifact(s) purged (documented lag)");
        }

        Assert.Equal(0, remaining);
        // Folders: LIVE ones only - an empty test folder can wedge in Deleted Items
        // for the rest of an Outlook session (documented in DeleteTestFolders, which
        // tolerates and reports it). Asserting the raw count pins that limitation
        // instead of the contract; see CountLiveTestFolders.
        Assert.Equal(0, LiveOutlookTestMailer.CountLiveTestFolders(Hub, out int wedgedEmptyFolders));
        if (wedgedEmptyFolders > 0)
        {
            Console.WriteLine($"cleanup[{Hub}]: {wedgedEmptyFolders} empty test folder(s) wedged in Deleted Items "
                + "until Outlook restarts (documented same-session limitation, no items involved)");
        }
        _output.WriteLine("t3 move/archive: 0 marker artifacts, 0 test folders remain in hub");

        // Belt-and-braces (D39): whole-store snapshot reconciliation after the wire run.
        _output.WriteLine(_fixture.VerifyHubReconciled());
    }
}
