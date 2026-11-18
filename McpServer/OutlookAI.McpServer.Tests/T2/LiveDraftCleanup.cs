using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The clean-up a live test that makes a tagged draft in the hub ends with (S3), written once
/// (Q154): delete the draft by its EntryID and sweep the run's other tagged artifacts, then prove
/// none is left. The same steps as the private helpers of <c>LiveUpdateDiscardTests</c>,
/// <c>LiveHtmlDraftTests</c>, <c>LiveSendTests</c>, <c>LiveDraftTests</c> and
/// <c>LiveDraftOptionsTests</c>, which can call these instead. Everything goes through
/// <see cref="LiveOutlookTestMailer"/>, whose guards decide what may be deleted.
/// </summary>
internal static class LiveDraftCleanup
{
    /// <summary>
    /// Deletes <paramref name="entryId"/> - which must carry the tag and <paramref name="marker"/> - and
    /// sweeps the run's other tagged artifacts out of <paramref name="hub"/>, trying three times when
    /// Outlook is busy.
    /// </summary>
    internal static void DeleteDraft(string hub, string marker, string entryId)
    {
        for (int attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                LiveOutlookTestMailer.DeleteItemByEntryId(hub, entryId, marker);
                LiveOutlookTestMailer.DeleteTaggedArtifacts(hub, marker);
                return;
            }
            catch (Exception) when (attempt < 2)
            {
                Thread.Sleep(1000);
            }
        }
    }

    /// <summary>Purges late-materialised stragglers and asserts that no tagged artifact of the run is left in <paramref name="hub"/>.</summary>
    internal static void AssertNoTaggedArtifactsRemain(string hub, string marker, ITestOutputHelper output)
    {
        int remaining = LiveOutlookTestMailer.CountTaggedArtifactsAfterPurgingStragglers(
            hub, marker, folderIds: null, out int stragglers);
        if (stragglers > 0)
        {
            output.WriteLine($"cleanup[{hub}]: {stragglers} late-materialized artifact(s) purged (documented lag)");
        }

        Assert.Equal(0, remaining);
    }
}
