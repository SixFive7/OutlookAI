using System.Globalization;
using OutlookAI.RemediationTools;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// T2, read-only, cached Exchange only: the zero-artifact half of the Exchange VM's read-only run -
/// that the hub, a REAL mailbox (Q108), holds no item carrying either of the testbed's subject tags in
/// any folder the artifact sweep covers.
/// <para>
/// <b>Why it exists (2026-10-03).</b> On a test guest the zero-artifact proof is the last step of the
/// write tests' own sweep, which counts and then purges. The Exchange VM runs only <c>Writes=Nothing</c>
/// tests until the maintainer approves its Phase 2 write-safety design, so nothing there ever asked the
/// question at all - and it is the one machine where a leftover would sit in real mail. This counts,
/// and never purges: a leftover found here fails the run and is removed afterwards by the tested sweep
/// on an approved write run, never by this test and never by hand (AGENTS.md, Mailbox Safety rule 1).
/// </para>
/// <para>
/// The folders are the sweep's own set - Drafts, Inbox, Sent Items, the archive folder, Outbox, the Sync
/// Issues subtree (Conflicts, Local Failures, Server Failures) and Deleted Items - and the match is the
/// same server-side subject restriction the sweep uses, once per tag: <see cref="LiveOutlookTestMailer.SubjectTag"/>
/// for live-tier artifacts and <see cref="CorpusPlan.SubjectTag"/> for corpus items, which must never be
/// in a real mailbox either.
/// </para>
/// <para>
/// Content-free (S4): counts only - never a subject, a sender or a folder's name.
/// </para>
/// </summary>
[Collection(LiveCollections.Phase1)]
[Trait("Category", "Live")]
public sealed class LiveExchangeHubArtifactTests
{
    private readonly LivePhase1Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveExchangeHubArtifactTests(LivePhase1Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    [Trait("Requires", "CachedExchange")]
    [Trait("Writes", "Nothing")]
    public void TheCachedExchangeHub_HoldsNoTaggedItem_InAnyFolderTheSweepCovers()
    {
        string hub = _fixture.Settings.TestHubStoreDisplayName;
        ComStoreDetailCheck(hub);

        int liveTier = LiveOutlookTestMailer.CountTaggedArtifacts(
            hub, Fragment(LiveOutlookTestMailer.SubjectTag), LiveOutlookTestMailer.HubSweepFolderIdsWithArchive);
        int corpus = LiveOutlookTestMailer.CountTaggedArtifacts(
            hub, Fragment(CorpusPlan.SubjectTag), LiveOutlookTestMailer.HubSweepFolderIdsWithArchive);

        _output.WriteLine("hub (cached Exchange): " + liveTier.ToString(CultureInfo.InvariantCulture)
            + " live-tier artifact(s), " + corpus.ToString(CultureInfo.InvariantCulture) + " corpus item(s) in the "
            + LiveOutlookTestMailer.HubSweepFolderIdsWithArchive.Length.ToString(CultureInfo.InvariantCulture)
            + " folders the sweep covers");

        Assert.True(liveTier == 0, liveTier + " item(s) tagged " + LiveOutlookTestMailer.SubjectTag
            + " are in the hub. Remove them with the tested sweep on an approved write run - never by hand.");
        Assert.True(corpus == 0, corpus + " item(s) tagged " + CorpusPlan.SubjectTag
            + " are in a real mailbox, where no corpus may ever be built.");
    }

    /// <summary>The tag without its brackets: the sweep's restriction matches a plain fragment.</summary>
    private static string Fragment(string tag)
    {
        return tag.Trim('[', ']');
    }

    /// <summary>Refuses unless the hub is a cached Exchange store, which is what this test is about.</summary>
    private void ComStoreDetailCheck(string hub)
    {
        IReadOnlyList<OutlookAI.Core.Com.ComStoreDetail> details = _fixture.Session.GetStoreDetails();
        OutlookAI.Core.Com.ComStoreDetail? detail = details.FirstOrDefault(
            d => string.Equals(d.DisplayName, hub, StringComparison.OrdinalIgnoreCase));
        Assert.True(detail != null, "the hub is not a store of the running profile");
        Assert.True(
            ShortDecodedIdExpectation.FormatOf(detail!.ExchangeStoreType, detail.IsCachedExchange) == DecodedIdStoreFormat.CachedExchange,
            "the hub is not a cached Exchange store (Requires=CachedExchange)");
    }
}
