using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The clean-up a live test that makes a tagged draft ends with (S3), written once
/// (Q154): delete the draft by its EntryID and sweep the run's other tagged artifacts, then prove
/// none is left. <c>LiveUpdateDiscardTests</c>, <c>LiveHtmlDraftTests</c>, <c>LiveSendTests</c>,
/// <c>LiveDraftTests</c> and <c>LiveDraftOptionsTests</c> each restated these steps and now call
/// them, as the save_attachment tests do. Everything goes through
/// <see cref="LiveOutlookTestMailer"/>, whose guards decide what may be deleted.
/// </summary>
internal static class LiveDraftCleanup
{
    /// <summary>
    /// Deletes <paramref name="entryId"/> - which must carry the tag and <paramref name="marker"/> - and
    /// sweeps the run's other tagged artifacts out of <paramref name="store"/>, the hub or a business
    /// account the draft was made in, trying three times when Outlook is busy.
    /// </summary>
    internal static void DeleteDraft(string store, string marker, string entryId)
    {
        for (int attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                LiveOutlookTestMailer.DeleteItemByEntryId(store, entryId, marker);
                LiveOutlookTestMailer.DeleteTaggedArtifacts(store, marker);
                return;
            }
            catch (Exception) when (attempt < 2)
            {
                Thread.Sleep(1000);
            }
        }
    }

    /// <summary>
    /// Purges late-materialised stragglers and asserts that no tagged artifact of the run is left in
    /// <paramref name="store"/> - in the mailer's default folders, or in <paramref name="folderIds"/>.
    /// </summary>
    internal static void AssertNoTaggedArtifactsRemain(string store, string marker, ITestOutputHelper output, int[]? folderIds = null)
    {
        int remaining = LiveOutlookTestMailer.CountTaggedArtifactsAfterPurgingStragglers(
            store, marker, folderIds, out int stragglers);
        if (stragglers > 0)
        {
            output.WriteLine($"cleanup[{store}]: {stragglers} late-materialized artifact(s) purged (documented lag)");
        }

        Assert.Equal(0, remaining);
    }
}
