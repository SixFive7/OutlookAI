using System.Text.Json;
using OutlookAI.Core.Audit;
using OutlookAI.McpServer.Tests.T2;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// T3 live acceptance for the Phase-4 draft tools (v3.MD section 0.6): all four called
/// over REAL stdio MCP against the built server exe with golden-shape asserts, and an
/// audit-log line asserted for EVERY write (the Phase-4 audit goes live). All artifacts
/// target the test hub (S2), carry tag + run marker, use display:false, and are deleted
/// after assert with a 0-remaining count (S3). Output is content-free for business
/// stores (S4) - everything logged here is agent-authored hub content or booleans.
/// </summary>
[Collection("LivePhase4")]
[Trait("Category", "Live")]
public sealed class Phase4LiveMcpToolShapeTests
{
    /// <summary>
    /// How long the seed may take to become visible over stdio. Covers a real mail round
    /// trip plus the ~10 s sweep cache TTL, which can delay an arrival during rapid polling.
    /// </summary>
    private const int SeedVisibleSeconds = 180;

    private readonly LivePhase4Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public Phase4LiveMcpToolShapeTests(LivePhase4Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    [Fact]
    public async Task DraftTools_GoldenShapes_OverRealStdio_WithAuditLines()
    {
        int auditLinesBefore = CountAuditLines();
        string seedSubject = LiveOutlookTestMailer.SubjectTag + " " + Marker + " t3seed";
        var draftEntryIds = new List<string>();
        try
        {
            // Seed (D20: hub -> itself) and wait for the Inbox copy.
            DateTime sentUtc = LiveOutlookTestMailer.SendSelfMail(Hub, seedSubject, "T3 seed body " + Marker, null);

            await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(TimeSpan.FromMinutes(8));

            // Find the seed over stdio: the always-on freshness sweep (D34) catches
            // not-yet-indexed arrivals. The ~10 s sweep cache can delay visibility of
            // an arrival by up to its TTL during rapid polling - the 180 s deadline
            // absorbs that by design.
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
                    bool subjectMatches = hit.TryGetProperty("subject", out JsonElement subj)
                        && subj.GetString() == seedSubject;
                    if (subjectMatches && IsArrivedCopy(hit))
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
            _output.WriteLine($"seed found over stdio: secondsAfterSend={(DateTime.UtcNow - sentUtc).TotalSeconds:F1}");

            // The first drafting call of a server process is answered with the user's writing
            // rules instead of the work (WritingRulesGate). Spend it here: the four golden
            // shapes below are this test's subject, and its priming call cannot reach Outlook.
            await client.PrimeWritingRulesGateAsync();

            // --- reply_draft / replyall_draft / forward_draft / new_draft golden shapes.
            JsonElement reply = await client.CallToolAsync("reply_draft", new
            {
                id = hitId,
                body = "T3 reply body " + Marker,
                display = false,
            });
            draftEntryIds.Add(AssertDraftShape(reply, "reply"));

            JsonElement replyAll = await client.CallToolAsync("replyall_draft", new
            {
                id = hitId,
                body = "T3 replyall body " + Marker,
                display = false,
            });
            draftEntryIds.Add(AssertDraftShape(replyAll, "replyall"));

            JsonElement forward = await client.CallToolAsync("forward_draft", new
            {
                id = hitId,
                body = "T3 forward body " + Marker,
                to = Hub,
                display = false,
            });
            draftEntryIds.Add(AssertDraftShape(forward, "forward"));

            JsonElement fresh = await client.CallToolAsync("new_draft", new
            {
                account = Hub,
                to = Hub,
                subject = LiveOutlookTestMailer.SubjectTag + " " + Marker + " t3new",
                body = "T3 new body " + Marker,
                display = false,
            });
            draftEntryIds.Add(AssertDraftShape(fresh, "new"));

            // --- audit: one structured line for EVERY write, with the draft's EntryID.
            IReadOnlyList<string> newLines = ReadAuditLinesAfter(auditLinesBefore);
            AssertAuditLine(newLines, "reply_draft", draftEntryIds[0]);
            AssertAuditLine(newLines, "replyall_draft", draftEntryIds[1]);
            AssertAuditLine(newLines, "forward_draft", draftEntryIds[2]);
            AssertAuditLine(newLines, "new_draft", draftEntryIds[3]);
            _output.WriteLine($"audit: {newLines.Count} new lines, all four draft ops present with matching entryIds");

            Assert.True(await client.CloseAndAwaitExitAsync(TimeSpan.FromSeconds(30)), "server must exit on stdin close");
        }
        finally
        {
            // Self-send copies can arrive AFTER a one-shot cleanup (delivery lag,
            // Phase-4 live finding) - loop delete+count until stable zero.
            int deleted = 0;
            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    deleted = LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(Hub, Marker);
                    break;
                }
                catch (Exception) when (attempt < 2)
                {
                    await Task.Delay(1000);
                }
            }

            _output.WriteLine($"cleanup: taggedArtifactsDeleted={deleted} (stable zero)");
        }

        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
        _output.WriteLine("t3 drafts: 0 marker artifacts remain in hub");
    }

    /// <summary>
    /// True when a hit is the seed's ARRIVED (Inbox) copy - the only copy the draft tools
    /// can safely be pointed at.
    /// <para>
    /// A self-send briefly exists as a Drafts and then an Outbox copy, and the index can
    /// ingest one of those within a second or two (the same fact soak fix 14 hit in the
    /// freshness test). The old acceptance took ANY index hit - no folderKind means
    /// "index row" - so the test could reply to a copy that had already moved on, and
    /// the id then failed to locate: "Hit could not be located in Outlook
    /// (url:NoSubjectTimeMatch)". Requiring the Inbox copy removes the race instead of
    /// widening a timeout around it.
    /// </para>
    /// </summary>
    private static bool IsArrivedCopy(JsonElement hit)
    {
        if (hit.TryGetProperty("folderKind", out JsonElement kind))
        {
            return kind.GetString() == "inbox"; // live (freshness sweep) hit
        }

        if (!hit.TryGetProperty("folder", out JsonElement folderElement))
        {
            return false;
        }

        string? folder = folderElement.GetString();
        if (string.IsNullOrEmpty(folder))
        {
            return false;
        }

        string leaf = folder!.Split('/')[^1];
        return leaf.Equals("Inbox", StringComparison.OrdinalIgnoreCase)
            || leaf.Equals("Postvak IN", StringComparison.OrdinalIgnoreCase);
    }

    private string AssertDraftShape(JsonElement outcome, string expectedKind)
    {
        if (!outcome.TryGetProperty("kind", out _))
        {
            // Domain error payload instead of a draft outcome - surface WHAT the server
            // said (test-artifact content only, S4-safe) instead of a bare
            // KeyNotFoundException (live-bitten: a reply against a ~1 s-old seed).
            Assert.Fail($"{expectedKind}: tool returned no draft outcome: {outcome.GetRawText()}");
        }

        Assert.Equal(expectedKind, outcome.GetProperty("kind").GetString());
        string entryId = outcome.GetProperty("entryId").GetString()!;
        Assert.True(entryId.Length >= 48, "draft outcome must carry the real EntryID");
        Assert.Equal(Hub, outcome.GetProperty("store").GetString(), ignoreCase: true);
        Assert.False(string.IsNullOrEmpty(outcome.GetProperty("folder").GetString()));
        Assert.True(outcome.GetProperty("accountResolved").GetBoolean(), expectedKind + ": accountResolved must be true");
        Assert.Equal(Hub, outcome.GetProperty("account").GetString(), ignoreCase: true);
        Assert.False(outcome.GetProperty("displayed").GetBoolean(), "tests suppress display");
        Assert.True(outcome.TryGetProperty("signatureInjected", out JsonElement sig)
            && (sig.ValueKind == JsonValueKind.True || sig.ValueKind == JsonValueKind.False));
        Assert.True(outcome.GetProperty("recipients").GetArrayLength() >= 1, expectedKind + ": recipients expected");
        if (expectedKind != "new")
        {
            Assert.False(string.IsNullOrEmpty(outcome.GetProperty("sourceEntryId").GetString()));
        }

        _output.WriteLine($"{expectedKind}_draft shape ok: store=hub displayed=false "
            + $"signatureInjected={sig.GetBoolean()} recipients={outcome.GetProperty("recipients").GetArrayLength()}");
        return entryId;
    }

    private static int CountAuditLines()
    {
        string path = AuditLog.DefaultLogPath;
        return File.Exists(path) ? File.ReadAllLines(path).Length : 0;
    }

    private static IReadOnlyList<string> ReadAuditLinesAfter(int skipCount)
    {
        string path = AuditLog.DefaultLogPath;
        Assert.True(File.Exists(path), "audit log must exist after write operations");
        return File.ReadAllLines(path).Skip(skipCount).ToList();
    }

    private static void AssertAuditLine(IReadOnlyList<string> lines, string operation, string entryId)
    {
        Assert.Contains(lines, l =>
            l.Contains(" op=" + operation + " ", StringComparison.Ordinal)
            && l.Contains("entryId=\"" + entryId + "\"", StringComparison.OrdinalIgnoreCase)
            && l.Contains("displayed=\"false\"", StringComparison.Ordinal));
    }
}
