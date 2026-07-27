using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Live verification of the orphan-index-row handling (soak fix 16 part B2). The index
/// legitimately outlives the mailbox: block (q) measured ~458 rows in one delegate store
/// filed under a leaf name with NO Outlook folder - search returns them, nothing can open
/// them, and the shipped error said "Re-run search - the item may have moved", the one
/// remedy that provably cannot work.
/// <para>
/// STRICTLY READ-ONLY. The delegate probe reads counts and flags only - no subject,
/// sender or body reaches the output (S4) - and writes nothing anywhere (S1).
/// </para>
/// </summary>
[Collection("LivePhase2")]
[Trait("Category", "Live")]
public sealed class LiveStaleIndexRowTests
{
    private readonly LivePhase2Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveStaleIndexRowTests(LivePhase2Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    public void OrphanRows_AreFlagged_Advised_AndFailReadWithAnActionableMessage()
    {
        OrphanFolderProbeSettings? probe = _fixture.Settings.OrphanFolderProbe;
        if (probe == null || string.IsNullOrWhiteSpace(probe.StoreDisplayName)
            || string.IsNullOrWhiteSpace(probe.FolderName))
        {
            _output.WriteLine("no orphanFolderProbe configured - covered by the no-false-positive test only.");
            return;
        }

        SearchOutcome outcome = _fixture.Service.Search(new SearchRequest
        {
            Store = probe.StoreDisplayName,
            Folder = probe.FolderName,
            Top = 5,
            SnippetChars = 0,
        });

        _output.WriteLine(
            $"orphan probe: {outcome.Hits.Count} hit(s), "
            + $"{outcome.Hits.Count(h => h.StaleIndexRow == true)} flagged staleIndexRow.");

        Assert.NotEmpty(outcome.Hits);

        // (1) every hit under a folder Outlook does not have is flagged...
        Assert.All(outcome.Hits, h => Assert.True(h.StaleIndexRow));

        // (2) ...and the agent is told, in advice, before it burns a read on one.
        Assert.NotNull(outcome.Advice);
        Assert.Contains(outcome.Advice!, a => a.Contains("staleIndexRow", StringComparison.Ordinal));

        // (3) reading one fails with a diagnosis, not a misleading retry suggestion.
        InvalidOperationException error = Assert.Throws<InvalidOperationException>(
            () => _fixture.Service.Read(outcome.Hits[0].Id, maxBodyChars: 0));
        _output.WriteLine("read error: " + error.Message);

        Assert.Contains("no longer exists in Outlook", error.Message, StringComparison.Ordinal);
        Assert.Contains("stale index row", error.Message, StringComparison.Ordinal);
        Assert.Contains("exhaustive:true", error.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("may have moved", error.Message, StringComparison.Ordinal);

        // (4) the diagnostic token survives for debugging.
        Assert.Contains("FolderNotFound", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void OrdinaryHits_AreNeverFlaggedStale()
    {
        // The flag must be rare and right: a false positive would teach agents to ignore
        // real mail. Probe the hub, whose folders certainly exist.
        SearchOutcome outcome = _fixture.Service.Search(new SearchRequest
        {
            Store = _fixture.Settings.TestHubStoreDisplayName,
            Top = 25,
            SnippetChars = 0,
        });

        _output.WriteLine($"hub probe: {outcome.Hits.Count} hit(s), none may be flagged.");
        Assert.All(outcome.Hits, h => Assert.Null(h.StaleIndexRow));
        if (outcome.Advice != null)
        {
            Assert.DoesNotContain(outcome.Advice, a => a.Contains("staleIndexRow", StringComparison.Ordinal));
        }
    }
}
