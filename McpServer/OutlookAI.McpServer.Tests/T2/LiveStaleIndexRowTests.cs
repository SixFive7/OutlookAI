using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Live cover for delegate hits whose folder the index publishes FLAT (soak fix 16 part
/// B2) - and for the stale-row marker that remains once they resolve.
/// <para>
/// THE DEFECT this pins: the delegate index namespace drops every intermediate folder, so
/// an item in the delegate's <c>Archive/SomeFolder</c> is indexed as
/// <c>&lt;host&gt;/1/&lt;delegate&gt;/SomeFolder</c>. The locator walked that path from the
/// delegate store root, found nothing, and EVERY such hit failed to open - read,
/// save_attachment, open_in_outlook and the thread COM fallback alike. D42 fixed searching
/// those folders; opening what the search returned stayed broken.
/// </para>
/// <para>
/// STRICTLY READ-ONLY on the delegate mailbox: counts, booleans and the locator tier only -
/// no subject, sender or body reaches the output (S4), and nothing is written (S1).
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
    public void DelegateHitsInANestedFolder_AreReadable_ViaTheFlatLeafName()
    {
        DelegateNestedFolderProbeSettings? probe = _fixture.Settings.DelegateNestedFolderProbe;
        if (probe == null || string.IsNullOrWhiteSpace(probe.StoreDisplayName)
            || string.IsNullOrWhiteSpace(probe.FolderName))
        {
            _output.WriteLine("no delegateNestedFolderProbe configured - skipping the positive half.");
            return;
        }

        SearchOutcome outcome = _fixture.Service.Search(new SearchRequest
        {
            Store = probe.StoreDisplayName,
            Folder = probe.FolderName,
            Top = 3,
            SnippetChars = 0,
        });

        _output.WriteLine(
            $"delegate nested probe: {outcome.Hits.Count} hit(s), "
            + $"{outcome.Hits.Count(h => h.StaleIndexRow == true)} flagged staleIndexRow.");

        Assert.NotEmpty(outcome.Hits);

        // (1) The folder EXISTS in Outlook (nested), so nothing here may be called stale.
        Assert.All(outcome.Hits, h => Assert.Null(h.StaleIndexRow));
        if (outcome.Advice != null)
        {
            Assert.DoesNotContain(outcome.Advice, a => a.Contains("staleIndexRow", StringComparison.Ordinal));
        }

        // (2) ...and the hit OPENS - the regression this test exists for. Before the fix
        // this threw "FolderNotFound" for every delegate item in a subfolder.
        ReadOutcome read = _fixture.Service.Read(outcome.Hits[0].Id, maxBodyChars: 0);
        _output.WriteLine($"read locatedVia={read.LocatedVia} in {read.LocateMs} ms; folder='{read.Folder}'.");

        Assert.False(string.IsNullOrEmpty(read.EntryId));
        Assert.Equal("delegateLeafName", read.LocatedVia);

        // (3) The COM folder really is nested - i.e. the flat URL genuinely addressed
        // nothing and the leaf-name resolution is what found it.
        using OutlookComSession verify = OutlookComSession.Connect(allowStartingOutlook: true);
        IReadOnlyList<IReadOnlyList<string>> matches = verify.FindFolderPathsByLeafName(
            probe.StoreDisplayName, probe.FolderName, 5000);
        Assert.NotEmpty(matches);
        Assert.All(matches, m => Assert.True(m.Count > 1, "the probe folder must be nested, not top-level"));
        _output.WriteLine("COM leaf matches: " + string.Join(" | ", matches.Select(m => string.Join("/", m))));
    }

    [Fact]
    public void OrdinaryHits_AreNeverFlaggedStale()
    {
        // The marker must be rare and right: a false positive would teach agents to ignore
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
