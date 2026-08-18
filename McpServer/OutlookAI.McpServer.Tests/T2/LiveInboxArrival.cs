using System.Linq;

using OutlookAI.Core.Com;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The D20 round-trip wait, once.
/// <para>
/// This was five verbatim copies - LiveDraftTests, LiveDraftOptionsTests,
/// LiveHtmlDraftTests, LiveSignatureTests, LiveUpdateDiscardTests - each with its own
/// <c>AddSeconds(180)</c> and <c>Thread.Sleep(3000)</c>. Real mail crossing a real mail
/// server is the slowest thing the live tier waits on, so if the deadline ever needs
/// raising it needs raising everywhere; five copies means four of them stay wrong and the
/// suite goes flaky in whichever file was forgotten.
/// </para>
/// <para>
/// A SIXTH copy was missed by that consolidation and proved the point: LiveMoveArchiveTests
/// waited its own <c>120</c> seconds through a store walk, and a real round trip that took
/// longer than two minutes failed a 17-minute live run. It now calls this helper - the copy
/// is gone rather than synchronised, which is the only version of this fix that stays fixed.
/// </para>
/// </summary>
internal static class LiveInboxArrival
{
    /// <summary>
    /// How long a seeded mail may take to come back through the mail server. Generous
    /// because the failure mode of being too tight is a flaky suite that blames the code
    /// under test for the mail server's latency.
    /// </summary>
    internal const int DeadlineSeconds = 180;

    /// <summary>Gap between sweeps while waiting. Each sweep is a real COM folder read.</summary>
    internal const int PollIntervalMs = 3000;

    /// <summary>
    /// How far BEFORE the send instant the sweep window starts, absorbing clock skew
    /// between this machine and the store's DateReceived.
    /// </summary>
    internal const int WindowLeadMinutes = 2;

    /// <summary>Items read per folder per sweep.</summary>
    private const int PerFolderCap = 100;

    /// <summary>
    /// Waits for <paramref name="subject"/> to land in <paramref name="hubStore"/>'s Inbox,
    /// sweeping through <paramref name="session"/>. READ-ONLY: it sweeps, it never mutates.
    /// </summary>
    /// <exception cref="TimeoutException">The mail did not arrive inside the deadline.</exception>
    internal static ComMailBrief WaitFor(OutlookComSession session, string hubStore, string subject, DateTime sentUtc)
    {
        LiveWaitBudget wait = LiveWaitBudget.OfSeconds(DeadlineSeconds);
        while (wait.HasTimeLeft)
        {
            ComSweepResult sweep = session.SweepFoldersNewerThan(
                sentUtc.AddMinutes(-WindowLeadMinutes),
                perFolderCap: PerFolderCap,
                includeBodies: false,
                onlyStoreDisplayName: hubStore);
            ComMailBrief? hit = sweep.Items.FirstOrDefault(
                i => i.FolderKind == "inbox" && string.Equals(i.Subject, subject, StringComparison.Ordinal));
            if (hit != null)
            {
                return hit;
            }

            Thread.Sleep(PollIntervalMs);
        }

        throw new TimeoutException(
            $"Seed mail did not arrive in the hub Inbox within {DeadlineSeconds} s (D20 round trip).");
    }
}
