using System.Diagnostics;
using System.Text.Json;
using OutlookAI.Core.Audit;
using OutlookAI.Core.Com;
using OutlookAI.McpServer.Tests.T2;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T3;

/// <summary>
/// T3 live acceptance for the Phase-5 high-friction send (v3.MD section 0.6 Phase 5):
/// the FULL two-step flow through the real send tool over REAL stdio MCP against the
/// built server exe. A tagged hub draft is created via the draft layer, the first send
/// call MUST refuse and issue a token, a bogus token MUST refuse, the token call sends;
/// arrival is verified on the INBOX side (Phase-2 fact: the Sent copy lags) via COM
/// read-back AND the product surface (search fresh + read), the From identity is
/// asserted = the hub account (Phase-4 putref lesson), Sent Items filing is verified
/// lag-tolerantly, and an audit line is asserted for EVERY step. All artifacts are
/// hub-only (S2/D20), tagged + marker'd and deleted to stable zero (S3).
/// </summary>
[Collection("LivePhase5")]
[Trait("Category", "Live")]
public sealed class Phase5LiveMcpToolShapeTests
{
    private readonly LivePhase5Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public Phase5LiveMcpToolShapeTests(LivePhase5Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    [Fact]
    public async Task SendTool_TwoStepFlow_RoundTrip_OverRealStdio_WithAuditLines()
    {
        int auditBefore = CountAuditLines();
        string sendSubject = _fixture.TaggedSubject("send-rt");
        try
        {
            await using McpStdioClient client = await McpStdioClient.StartAndInitializeAsync(TimeSpan.FromMinutes(10));

            // The first drafting call of a server process is answered with the user's writing
            // rules instead of the work (WritingRulesGate); spend it before the real draft.
            await client.PrimeWritingRulesGateAsync();

            // Draft via the draft layer (D4 default path), display suppressed for tests.
            JsonElement draft = await client.CallToolAsync("new_draft", new
            {
                account = Hub,
                to = Hub,
                subject = sendSubject,
                body = "Phase-5 send round-trip body " + Marker,
                display = false,
            });
            string draftEntryId = draft.GetProperty("entryId").GetString()!;
            Assert.True(draft.GetProperty("accountResolved").GetBoolean());

            // --- STEP 1: send WITHOUT token MUST refuse and issue a one-time token.
            Stopwatch step1Watch = Stopwatch.StartNew();
            JsonElement step1 = await client.CallToolAsync("send", new { id = draftEntryId });
            step1Watch.Stop();
            Assert.Equal("confirmation_required", step1.GetProperty("status").GetString());
            Assert.False(step1.GetProperty("sent").GetBoolean(), "step 1 must never send");
            string token = step1.GetProperty("confirmToken").GetString()!;
            Assert.StartsWith("confirm-", token, StringComparison.Ordinal);
            Assert.Contains("NOT SENT", step1.GetProperty("warning").GetString(), StringComparison.Ordinal);
            double ttl = step1.GetProperty("tokenExpiresInSeconds").GetDouble();
            Assert.InRange(ttl, 30, 600);
            Assert.Equal(Hub, step1.GetProperty("account").GetString(), ignoreCase: true);
            Assert.Equal(Hub, step1.GetProperty("store").GetString(), ignoreCase: true);
            Assert.Equal(draftEntryId, step1.GetProperty("entryId").GetString(), ignoreCase: true);
            Assert.True(step1.GetProperty("recipients").GetArrayLength() >= 1);
            _output.WriteLine($"step1 golden: refused+token in {step1Watch.ElapsedMilliseconds} ms, ttl={ttl:F0}s");

            // --- NEGATIVE over stdio: a bogus token refuses with the SendRefused shape.
            JsonElement bogus = await client.CallToolAsync("send", new
            {
                id = draftEntryId,
                confirm_token = "confirm-ffffffffffffffffffffffffffffffff",
            });
            JsonElement bogusError = bogus.GetProperty("error");
            Assert.Equal("SendRefused", bogusError.GetProperty("type").GetString());
            Assert.Contains("Nothing was sent", bogusError.GetProperty("advice").GetString(), StringComparison.Ordinal);

            // --- STEP 2: the real token sends; identity readback-verified server-side.
            Stopwatch sendWatch = Stopwatch.StartNew();
            JsonElement step2 = await client.CallToolAsync("send", new { id = draftEntryId, confirm_token = token });
            sendWatch.Stop();
            DateTime sentUtc = DateTime.UtcNow;
            Assert.Equal("sent", step2.GetProperty("status").GetString());
            Assert.True(step2.GetProperty("sent").GetBoolean());
            Assert.True(step2.GetProperty("accountVerified").GetBoolean(), "SendUsingAccount readback must be verified");
            Assert.Equal(Hub, step2.GetProperty("account").GetString(), ignoreCase: true);
            Assert.Equal(draftEntryId, step2.GetProperty("entryId").GetString(), ignoreCase: true);
            _output.WriteLine($"step2: sent in {sendWatch.ElapsedMilliseconds} ms, account={step2.GetProperty("account").GetString()} verified=true");

            // --- ARRIVAL: Inbox side via independent COM read-back (Phase-2 fact 4:
            // never verify a send via Sent Items). From identity = hub (Phase-4 lesson).
            ComMailBrief arrived = WaitForInboxArrival(sendSubject, sentUtc);
            double arrivalSeconds = (DateTime.UtcNow - sentUtc).TotalSeconds;
            Assert.True(arrived.SenderAddress != null
                && string.Equals(arrived.SenderAddress, Hub, StringComparison.OrdinalIgnoreCase),
                "the RECEIVED copy must report the hub as sender (From identity)");
            _output.WriteLine($"arrival: inbox copy {arrivalSeconds:F1} s after send, sender={arrived.SenderAddress}");

            // --- Product surface: the sent mail is findable via search(fresh) over
            // stdio and reads back with the hub From identity.
            (string hitId, JsonElement read) = await FindReadableHitAsync(client, sendSubject);
            Assert.Equal(Hub, read.GetProperty("fromAddress").GetString(), ignoreCase: true);
            _output.WriteLine($"stdio: hit {hitId} read back fromAddress=hub");

            // --- NEGATIVE: sending the ARRIVED (already sent) copy refuses.
            JsonElement refuseSent = await client.CallToolAsync("send", new { id = arrived.EntryId });
            Assert.Equal("SendRefused", refuseSent.GetProperty("error").GetProperty("type").GetString());

            // --- Sent Items filing (lag-tolerant per Phase-2 fact).
            double sentFilingSeconds = WaitForSentItemsCopy(sentUtc);
            _output.WriteLine($"sent items: copy visible {sentFilingSeconds:F1} s after send");

            // --- AUDIT: one line for EVERY step of the flow.
            IReadOnlyList<string> lines = ReadAuditLinesAfter(auditBefore);
            int issuedIndex = IndexOfLine(lines, " op=send_token_issued ", draftEntryId, null);
            int sendIndex = IndexOfLine(lines, " op=send ", draftEntryId, "accountVerified=\"true\"");
            Assert.True(issuedIndex >= 0, "send_token_issued audit line missing");
            Assert.True(sendIndex >= 0, "send audit line missing (accountVerified)");
            Assert.True(issuedIndex < sendIndex, "token issue must be audited before the send");
            Assert.True(IndexOfLine(lines, " op=send_refused ", draftEntryId, "reason=\"unknown_or_used_token\"") >= 0,
                "bogus-token refusal audit line missing");
            Assert.True(IndexOfLine(lines, " op=send_refused ", arrived.EntryId, "reason=\"not_an_unsent_draft\"") >= 0,
                "already-sent refusal audit line missing");
            _output.WriteLine($"audit: {lines.Count} new lines - token_issued -> send ordered, both refusals present");

            Assert.True(await client.CloseAndAwaitExitAsync(TimeSpan.FromSeconds(30)), "server must exit on stdin close");
        }
        finally
        {
            // Inbox/Sent copies can materialize AFTER a one-shot cleanup pass
            // (Phase-4 live finding) - loop delete+count until stable zero.
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
        _output.WriteLine("send round-trip: 0 marker artifacts remain in hub");
    }

    // ------------------------------------------------------------------ helpers

    private ComMailBrief WaitForInboxArrival(string subject, DateTime sentUtc)
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(180);
        while (DateTime.UtcNow < deadline)
        {
            ComSweepResult sweep = _fixture.VerifySession.SweepFoldersNewerThan(
                sentUtc.AddMinutes(-2), perFolderCap: 100, includeBodies: false, onlyStoreDisplayName: Hub);
            ComMailBrief? hit = sweep.Items.FirstOrDefault(i =>
                i.FolderKind == "inbox" && string.Equals(i.Subject, subject, StringComparison.Ordinal));
            if (hit != null)
            {
                return hit;
            }

            Thread.Sleep(3000);
        }

        throw new TimeoutException("Sent mail did not arrive in the hub Inbox within 180 s (D20 round trip).");
    }

    /// <summary>
    /// Polls search(fresh) over stdio for the subject and returns the first hit that
    /// READS successfully (index rows of just-deleted drafts can linger - Phase-2
    /// fact 9 - so unlocatable hits are skipped, not failed).
    /// </summary>
    private async Task<(string HitId, JsonElement Read)> FindReadableHitAsync(McpStdioClient client, string subject)
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(120);
        while (DateTime.UtcNow < deadline)
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
                if (!hit.TryGetProperty("subject", out JsonElement subjectProp)
                    || subjectProp.GetString() != subject)
                {
                    continue;
                }

                string hitId = hit.GetProperty("id").GetString()!;
                JsonElement read = await client.CallToolAsync("read", new { id = hitId, max_body_chars = 0 });
                if (read.TryGetProperty("error", out _))
                {
                    continue; // stale index row of the deleted draft - try the next hit
                }

                return (hitId, read);
            }

            await Task.Delay(3000);
        }

        throw new TimeoutException("Sent mail not readable over stdio within 120 s.");
    }

    /// <summary>Sent Items filing check via direct folder count (lag-tolerant, Phase-2 fact).</summary>
    private double WaitForSentItemsCopy(DateTime sentUtc)
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(180);
        while (DateTime.UtcNow < deadline)
        {
            if (LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker, new[] { 5 }) >= 1)
            {
                return (DateTime.UtcNow - sentUtc).TotalSeconds;
            }

            Thread.Sleep(3000);
        }

        throw new TimeoutException("Sent Items copy of the sent mail not visible within 180 s.");
    }

    private static int CountAuditLines()
    {
        string path = AuditLog.DefaultLogPath;
        return File.Exists(path) ? File.ReadAllLines(path).Length : 0;
    }

    private static IReadOnlyList<string> ReadAuditLinesAfter(int skipCount)
    {
        string path = AuditLog.DefaultLogPath;
        Assert.True(File.Exists(path), "audit log must exist after send operations");
        return File.ReadAllLines(path).Skip(skipCount).ToList();
    }

    private static int IndexOfLine(IReadOnlyList<string> lines, string opFragment, string entryId, string? extraFragment)
    {
        for (int i = 0; i < lines.Count; i++)
        {
            if (lines[i].Contains(opFragment, StringComparison.Ordinal)
                && lines[i].Contains("entryId=\"" + entryId + "\"", StringComparison.OrdinalIgnoreCase)
                && (extraFragment == null || lines[i].Contains(extraFragment, StringComparison.Ordinal)))
            {
                return i;
            }
        }

        return -1;
    }
}
