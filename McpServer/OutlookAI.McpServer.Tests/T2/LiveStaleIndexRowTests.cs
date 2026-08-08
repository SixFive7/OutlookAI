using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Live cover for delegate hits whose folder the index publishes FLAT (soak fix 16 part B2).
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

        _output.WriteLine($"delegate nested probe: {outcome.Hits.Count} hit(s).");
        Assert.NotEmpty(outcome.Hits);

        // A delegate/shared mailbox syncs its folder HIERARCHY lazily: the same store
        // enumerated this nested folder in one walk and not in the next, minutes apart
        // (measured during soak fix 16). Resolution can only work while the tree exposes
        // it, so the tree is asked FIRST and the read is asserted only when it does.
        using OutlookComSession verify = OutlookComSession.Connect(allowStartingOutlook: true);
        IReadOnlyList<IReadOnlyList<string>> matches = verify.FindFolderPathsByLeafName(
            probe.StoreDisplayName, probe.FolderName, HitLocator.DelegateLeafWalkCap);
        _output.WriteLine("COM leaf matches: "
            + (matches.Count == 0
                ? "(none - the delegate hierarchy is not enumerable right now)"
                : string.Join(" | ", matches.Select(m => string.Join("/", m)))));

        if (matches.Count == 0 || matches.All(m => m.Count <= 1))
        {
            _output.WriteLine(
                "the delegate folder tree does not currently expose the nested folder - there is nothing to "
                + "resolve against, so the locator assertion is skipped this run.");
            return;
        }

        // THE REGRESSION: before the fix this threw FolderNotFound for every delegate item
        // in a subfolder, because the flat index path was walked from the store root.
        ReadOutcome read = _fixture.Service.Read(outcome.Hits[0].Id, maxBodyChars: 0);
        _output.WriteLine($"read locatedVia={read.LocatedVia} in {read.LocateMs} ms; folder='{read.Folder}'.");

        Assert.False(string.IsNullOrEmpty(read.EntryId));
        Assert.Equal("delegateLeafName", read.LocatedVia);
    }

}
