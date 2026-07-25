using OutlookAI.Core.Audit;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Phase-5 T2 live negatives for the high-friction send policy (v3.MD D4/L5): token
/// binding to draft id + content, strict single-use, expiry, and refusal of missing
/// drafts - all WITHOUT any transport (every path here refuses BEFORE Send(); the
/// audit log is asserted to contain NO send line for this run's drafts). All artifacts
/// are hub-only drafts (S2/D20), tagged + marker'd, deleted after assert (S3).
/// </summary>
[Collection("LivePhase5")]
[Trait("Category", "Live")]
public sealed class LiveSendTests
{
    private readonly LivePhase5Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveSendTests(LivePhase5Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    private string Marker => _fixture.RunMarker;

    [Fact]
    public void TokenFlow_BindingSingleUseAndModification_AllRefuse_NothingSent()
    {
        int auditBefore = CountAuditLines();
        string hubStoreId = _fixture.GetStoreId(Hub);
        var draftIds = new List<string>();
        try
        {
            DraftOutcome draftA = Service.NewDraft(
                Hub, Hub, cc: null, _fixture.TaggedSubject("neg-a"), "Negative-test draft A body " + Marker, display: false);
            draftIds.Add(draftA.EntryId);
            DraftOutcome draftB = Service.NewDraft(
                Hub, Hub, cc: null, _fixture.TaggedSubject("neg-b"), "Negative-test draft B body " + Marker, display: false);
            draftIds.Add(draftB.EntryId);

            // --- Step 1 golden shape: refuse + issue a bound one-time token.
            SendOutcome step1 = Service.Send(draftA.EntryId);
            Assert.Equal("confirmation_required", step1.Status);
            Assert.False(step1.Sent);
            Assert.NotNull(step1.ConfirmToken);
            Assert.StartsWith("confirm-", step1.ConfirmToken!, StringComparison.Ordinal);
            Assert.Contains("NOT SENT", step1.Warning, StringComparison.Ordinal);
            Assert.Equal(draftA.EntryId, step1.EntryId, ignoreCase: true);
            Assert.Equal(Hub, step1.Store, ignoreCase: true);
            Assert.Equal(Hub, step1.Account, ignoreCase: true);
            Assert.True(step1.TokenExpiresInSeconds is >= 30 and <= 600, "token TTL must be short (D4)");
            Assert.NotNull(step1.Recipients);
            Assert.True(step1.Recipients!.Count >= 1);
            _output.WriteLine($"step1: token issued, ttl={step1.TokenExpiresInSeconds:F0}s account={step1.Account}");

            // --- Token is bound to the DRAFT ID: draft B refuses (and burns the token).
            SendRefusedException mismatch = Assert.Throws<SendRefusedException>(
                () => Service.Send(draftB.EntryId, step1.ConfirmToken));
            Assert.Equal("token_draft_mismatch", mismatch.Reason);

            // --- STRICT single-use: even the bound draft cannot use the burned token.
            SendRefusedException burned = Assert.Throws<SendRefusedException>(
                () => Service.Send(draftA.EntryId, step1.ConfirmToken));
            Assert.Equal("unknown_or_used_token", burned.Reason);

            // --- Token is bound to CONTENT: a modified draft invalidates a fresh token.
            SendOutcome reissued = Service.Send(draftA.EntryId);
            LiveOutlookTestMailer.AppendToDraftBody(Hub, draftA.EntryId, Marker, "\r\nModified after token issue.");
            SendRefusedException changed = Assert.Throws<SendRefusedException>(
                () => Service.Send(draftA.EntryId, reissued.ConfirmToken));
            Assert.Equal("draft_changed", changed.Reason);

            // --- Garbage token refuses.
            SendRefusedException garbage = Assert.Throws<SendRefusedException>(
                () => Service.Send(draftA.EntryId, "confirm-ffffffffffffffffffffffffffffffff"));
            Assert.Equal("unknown_or_used_token", garbage.Reason);

            _output.WriteLine("negatives: draft_mismatch + single_use + draft_changed + garbage all refused");

            // --- Audit: one line per step, and NO send line for this run's drafts.
            IReadOnlyList<string> lines = ReadAuditLinesAfter(auditBefore);
            Assert.True(CountLines(lines, " op=send_token_issued ", draftA.EntryId) >= 2, "both token issues must be audited");
            AssertRefusalLine(lines, draftB.EntryId, "token_draft_mismatch");
            AssertRefusalLine(lines, draftA.EntryId, "unknown_or_used_token");
            AssertRefusalLine(lines, draftA.EntryId, "draft_changed");
            Assert.DoesNotContain(lines, l => l.Contains(" op=send ", StringComparison.Ordinal));
            _output.WriteLine($"audit: {lines.Count} new lines, refusals audited, zero op=send lines");
        }
        finally
        {
            foreach (string entryId in draftIds)
            {
                CleanupDraft(entryId);
            }
        }

        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
        _output.WriteLine("token negatives: 0 marker artifacts remain in hub");
    }

    [Fact]
    public void TokenExpiry_ShortTtlService_RefusesExpiredToken()
    {
        int auditBefore = CountAuditLines();
        // Same wiring as production, but a 2 s TTL so expiry is provable live.
        using var shortTtlService = new MailService(
            new ComGateway(allowStartingOutlook: true),
            new SendConfirmationTokens(TimeSpan.FromSeconds(2)));

        string? draftEntryId = null;
        try
        {
            DraftOutcome draft = shortTtlService.NewDraft(
                Hub, Hub, cc: null, _fixture.TaggedSubject("neg-exp"), "Expiry-test draft body " + Marker, display: false);
            draftEntryId = draft.EntryId;

            SendOutcome issued = shortTtlService.Send(draft.EntryId);
            Assert.Equal(2, issued.TokenExpiresInSeconds);

            Thread.Sleep(2600);

            SendRefusedException expired = Assert.Throws<SendRefusedException>(
                () => shortTtlService.Send(draft.EntryId, issued.ConfirmToken));
            Assert.Equal("token_expired", expired.Reason);

            IReadOnlyList<string> lines = ReadAuditLinesAfter(auditBefore);
            AssertRefusalLine(lines, draft.EntryId, "token_expired");
            Assert.DoesNotContain(lines, l => l.Contains(" op=send ", StringComparison.Ordinal));
            _output.WriteLine("expiry: 2 s token refused after 2.6 s, audited, nothing sent");
        }
        finally
        {
            if (draftEntryId != null)
            {
                CleanupDraft(draftEntryId);
            }
        }

        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
    }

    [Fact]
    public void Send_DeletedDraft_FailsClosed_WithoutSending()
    {
        int auditBefore = CountAuditLines();
        DraftOutcome draft = Service.NewDraft(
            Hub, Hub, cc: null, _fixture.TaggedSubject("neg-del"), "Deleted-draft test body " + Marker, display: false);
        string entryId = draft.EntryId;
        CleanupDraft(entryId);

        // The EntryID no longer resolves anywhere: the send flow must fail closed.
        Assert.ThrowsAny<InvalidOperationException>(() => Service.Send(entryId));

        IReadOnlyList<string> lines = ReadAuditLinesAfter(auditBefore);
        Assert.DoesNotContain(lines, l => l.Contains(" op=send ", StringComparison.Ordinal));
        Assert.Equal(0, LiveOutlookTestMailer.CountTaggedArtifacts(Hub, Marker));
        _output.WriteLine("deleted draft: send failed closed, nothing sent, 0 artifacts remain");
    }

    [Fact]
    public void ArtifactSweep_AllThreeAccounts_ZeroTaggedRemain()
    {
        // S3 post-suite proof: NO item tagged [OutlookAI-McpTest] - from this or any
        // earlier run - remains in Drafts/Inbox/Sent/Deleted of ANY account. Late-
        // materializing self-send copies of earlier collections (documented sent-copy
        // lag) are purged first, then stable zero is asserted (counts only, S4).
        foreach (string store in _fixture.Settings.ExpectedStoreDisplayNames)
        {
            int count = LiveOutlookTestMailer.CountTaggedArtifacts(store, "OutlookAI-McpTest");
            if (count > 0)
            {
                _output.WriteLine($"sweep[{store}]: {count} late-materialized tagged artifact(s) found - purging (documented sent-copy lag)");
                LiveOutlookTestMailer.DeleteTaggedArtifactsUntilStableZero(store, "OutlookAI-McpTest");
                count = LiveOutlookTestMailer.CountTaggedArtifacts(store, "OutlookAI-McpTest");
            }

            _output.WriteLine($"sweep[{store}]: taggedArtifacts={count}");
            Assert.Equal(0, count);
        }
    }

    // ------------------------------------------------------------------ helpers

    private void CleanupDraft(string entryId)
    {
        for (int attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                LiveOutlookTestMailer.DeleteItemByEntryId(Hub, entryId, Marker);
                LiveOutlookTestMailer.DeleteTaggedArtifacts(Hub, Marker);
                return;
            }
            catch (Exception) when (attempt < 2)
            {
                Thread.Sleep(1000);
            }
        }
    }

    private static int CountAuditLines()
    {
        string path = AuditLog.DefaultLogPath;
        return File.Exists(path) ? File.ReadAllLines(path).Length : 0;
    }

    private static IReadOnlyList<string> ReadAuditLinesAfter(int skipCount)
    {
        string path = AuditLog.DefaultLogPath;
        Assert.True(File.Exists(path), "audit log must exist after send-policy operations");
        return File.ReadAllLines(path).Skip(skipCount).ToList();
    }

    private static int CountLines(IReadOnlyList<string> lines, string opFragment, string entryId)
    {
        return lines.Count(l => l.Contains(opFragment, StringComparison.Ordinal)
            && l.Contains("entryId=\"" + entryId + "\"", StringComparison.OrdinalIgnoreCase));
    }

    private static void AssertRefusalLine(IReadOnlyList<string> lines, string entryId, string reason)
    {
        Assert.Contains(lines, l => l.Contains(" op=send_refused ", StringComparison.Ordinal)
            && l.Contains("entryId=\"" + entryId + "\"", StringComparison.OrdinalIgnoreCase)
            && l.Contains("reason=\"" + reason + "\"", StringComparison.Ordinal));
    }
}
