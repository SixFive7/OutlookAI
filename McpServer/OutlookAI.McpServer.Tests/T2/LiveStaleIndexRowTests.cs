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
/// <para>
/// <b>It cannot pass without resolving something (Q76, 2026-09-27).</b> Both of its early returns
/// - no probe configured, and a delegate folder tree that does not list the probe folder NESTED
/// right now - refuse on a Production profile and print <c>PROVED NOTHING:</c> on a Portable one.
/// The second used to log a line and return GREEN. Pinned by <c>T1/LiveEarlyReturnGuardTests</c>.
/// </para>
/// </summary>
[Collection(LiveCollections.Phase2)]
[Trait("Category", "Live")]
public sealed class LiveStaleIndexRowTests
{
    /// <summary>
    /// What the locator assertion resolves against, named as the Production refusal wraps it. A path
    /// of one segment is the folder at the top of the tree, where there is nothing flat to resolve.
    /// </summary>
    internal const string NestedPathPopulation =
        "a nested path for the probe folder in the delegate mailbox's folder tree as Outlook lists it right now";

    /// <summary>What a reader of a PROVED NOTHING line here is to do about it.</summary>
    internal const string NestedPathRemedy =
        "Outlook syncs a delegate mailbox's folder hierarchy lazily - the same nested folder was listed in one walk "
        + "and missing from the next, minutes apart (soak fix 16) - so re-run once the tree has synced. If it never "
        + "lists the folder nested, 'delegateNestedFolderProbe' names a folder that is not nested (any more) and "
        + "needs one that is.";

    private readonly LivePhase2Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveStaleIndexRowTests(LivePhase2Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    [Trait("Requires", "SearchIndex")]
    [Trait("Requires", "DelegateStore")]
    [Trait("Requires", "ProbePopulation")]
    [Trait("Writes", "None")]
    public void DelegateHitsInANestedFolder_AreReadable_ViaTheFlatLeafName()
    {
        DelegateNestedFolderProbeSettings? probe = _fixture.Settings.DelegateNestedFolderProbe;
        if (probe == null || string.IsNullOrWhiteSpace(probe.StoreDisplayName)
            || string.IsNullOrWhiteSpace(probe.FolderName))
        {
            // On the machine this test was written for, an absent probe means the settings
            // have drifted, and returning green would hide it: the whole point here is a
            // delegate folder Outlook nests and the index publishes flat, and without one
            // the test proves nothing at all. On a machine that has no delegate mailbox the
            // absence is simply true - and this test declares Requires=DelegateStore, so it
            // should not have been selected there in the first place.
            _fixture.Settings.RequireProductionPopulation("a delegateNestedFolderProbe population");
            _output.WriteLine(
                "PROVED NOTHING: no delegateNestedFolderProbe configured on this machine, so the "
                + "positive half of this test did not run.");
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

        // Nothing nested in the tree right now means nothing to resolve against, so the assertion
        // below cannot run. That used to log a line and return GREEN; it is the same question the
        // probe guard above asks, and it gets the same answer (Q76).
        IReadOnlyList<IReadOnlyList<string>> nested = LivePopulationCoverage.Require(
            _fixture.Settings,
            NestedPaths(matches),
            NestedPathPopulation,
            "the delegate leaf-name locator assertion",
            NestedPathRemedy,
            _output.WriteLine);
        if (nested.Count == 0)
        {
            return;
        }

        // THE REGRESSION: before the fix this threw FolderNotFound for every delegate item
        // in a subfolder, because the flat index path was walked from the store root.
        ReadOutcome read = _fixture.Service.Read(outcome.Hits[0].Id, maxBodyChars: 0);
        _output.WriteLine($"read locatedVia={read.LocatedVia} in {read.LocateMs} ms; folder='{read.Folder}'.");

        Assert.False(string.IsNullOrEmpty(read.EntryId));
        Assert.Equal("delegateLeafName", read.LocatedVia);
    }

    /// <summary>
    /// The leaf matches the locator assertion can resolve against: the ones BELOW the store root.
    /// Empty exactly when the test used to log "not currently exposed" and return green - no match
    /// at all, or only matches at the top of the tree - which <c>T1/LiveEarlyReturnGuardTests</c>
    /// holds it to, shape by shape.
    /// </summary>
    internal static IReadOnlyList<IReadOnlyList<string>> NestedPaths(IReadOnlyList<IReadOnlyList<string>> matches)
    {
        ArgumentNullException.ThrowIfNull(matches);
        return matches.Where(m => m.Count > 1).ToList();
    }
}
