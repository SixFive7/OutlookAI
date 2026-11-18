using System.Globalization;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// T2, cached Exchange, ONE-OFF: removes the one tagged leftover the maintainer confirmed by its EntryID from the
/// Exchange VM's bystander mailbox (Q138, confirmed 2026-11-18) - the only write that mailbox ever gets - and
/// proves it went and nothing else did.
/// <para>
/// <b>Why it is a test.</b> AGENTS.md, Mailbox Safety rules 1 and 8: an item is deleted only by the project's
/// tested helper code, run through the suite, never by hand. So the removal is
/// <see cref="LiveOutlookTestMailer.RemoveConfirmedLeftover"/>, whose decisions T1 pins
/// (<c>T1/ConfirmedLeftoverTests</c>), and it runs inside a guarded collection: the count tripwire censuses the
/// bystander item by item before and after it, and accepts this item's departure and no other.
/// </para>
/// <para>
/// <b>Never in a derived filter.</b> It declares <see cref="LiveRunFilters.ConfirmedLeftover"/>, which the
/// Exchange VM's run filter leaves out, and <c>CachedExchange</c>, which the PST guests' filters leave out; it
/// does not carry <c>Writes=Nothing</c>, so the read-only Exchange filter never selects it either. It runs only
/// when named: <c>-Filter 'Category=Live&amp;Requires=ConfirmedLeftover'</c> on the Exchange VM's runner. A second
/// run finds nothing, writes nothing, and passes on the same read-back proof.
/// </para>
/// <para>
/// Content-free (S4): counts, the pinned subject - test data the live tier once wrote - and the first twelve
/// characters of the pin's hash; never another item's subject or id.
/// </para>
/// </summary>
[Collection(LiveCollections.Phase2)]
[Trait("Category", "Live")]
public sealed class LiveConfirmedLeftoverTests
{
    private readonly LivePhase2Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveConfirmedLeftoverTests(LivePhase2Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    [Trait("Requires", "CachedExchange")]
    [Trait("Requires", "ConfirmedLeftover")]
    public void TheConfirmedLeftover_LeavesTheBystander_AndNoTaggedItemIsLeftThere()
    {
        LiveTestSettings settings = _fixture.Settings;
        ConfirmedLeftover leftover = ConfirmedLeftover.Q138;
        string? refusal = leftover.RefusalFor(
            settings.MachineProfile, settings.BystanderStoreDisplayNames, settings.TestHubStoreDisplayName,
            settings.BystanderStoreDisplayNames.FirstOrDefault() ?? string.Empty);
        Assert.True(refusal == null, refusal);
        string bystander = settings.BystanderStoreDisplayNames[0];

        int before = LiveArtifactSweep.CountTagged(settings, bystander);
        _output.WriteLine("bystander before: " + before.ToString(CultureInfo.InvariantCulture)
            + " tagged item(s) in the folders the sweep covers");

        ConfirmedLeftoverRemoval removal = LiveOutlookTestMailer.RemoveConfirmedLeftover(settings, bystander, leftover);
        _output.WriteLine(removal.Describe(leftover));

        int after = LiveArtifactSweep.CountTagged(settings, bystander);
        _output.WriteLine("bystander after: " + after.ToString(CultureInfo.InvariantCulture)
            + " tagged item(s) in the folders the sweep covers");

        Assert.True(removal.Gone, "the confirmed leftover still reads back in its folder or in Deleted Items: " + removal.Describe(leftover));
        Assert.True(removal.OtherTaggedInFolder == 0, removal.OtherTaggedInFolder.ToString(CultureInfo.InvariantCulture)
            + " other tagged item(s) are in the leftover's folder - none was confirmed, so none was touched; they are the maintainer's to decide.");
        Assert.True(after == 0, after.ToString(CultureInfo.InvariantCulture)
            + " tagged item(s) are still in the bystander after the removal - nothing but the confirmed one is ever removed there.");
        if (removal.Removed)
        {
            Assert.True(before == after + 1, "the removal took " + (before - after).ToString(CultureInfo.InvariantCulture)
                + " tagged item(s) out of the bystander's sweep folders, where it must take exactly one.");
        }
    }
}
