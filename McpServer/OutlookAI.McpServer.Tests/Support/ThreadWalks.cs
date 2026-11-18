using OutlookAI.Core.Com;
using System.Globalization;
using OutlookAI.Core.Services;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// What a conversation walk reports, built the way the thread suites need it: a walk that read three
/// members and added all three, anchored in one store - and the same walk naming one store it never asked
/// about. The one copy (Q154) of what <c>ThreadScopedStoreTests</c>, <c>ThreadUnindexedStoreTests</c> and
/// <c>ThreadDerivedStoreReportingTests</c> each built, imported through <c>using static</c>.
/// </summary>
internal static class ThreadWalks
{
    /// <summary>
    /// Member <paramref name="i"/> of the thread two thread suites walk in the profile's store, received
    /// <paramref name="i"/> minutes after its frontier - each suite wrote it out for itself.
    /// </summary>
    public static ComMailBrief Member(int i)
    {
        return new ComMailBrief(
            entryId: "MEMBER" + i.ToString(CultureInfo.InvariantCulture),
            storeDisplayName: AliceMailbox.Store,
            storeId: "store-alice",
            folderName: "Inbox",
            folderKind: "inbox",
            subject: "a test thread member",
            senderName: "Bob",
            senderAddress: "bob@example.com",
            receivedTime: AliceMailbox.Frontier.AddMinutes(i),
            isRead: true,
            hasAttachments: false,
            sizeBytes: 2048,
            body: "test body");
    }

    /// <summary>A walk that read three members and added all three, anchored in <paramref name="anchorStore"/>.</summary>
    public static ThreadLiveInfo Walked(string? anchorStore)
    {
        return new ThreadLiveInfo
        {
            Performed = true,
            MembersWalked = 3,
            MembersAdded = 3,
            AnchorStore = anchorStore,
        };
    }

    /// <summary>A walk of <paramref name="walkedStore"/> that names <paramref name="unqueriedStore"/> as never asked about.</summary>
    public static ThreadLiveInfo Unqueried(string walkedStore, string unqueriedStore)
    {
        ThreadLiveInfo live = Walked(walkedStore);
        live.StoresNotQueried = new[] { unqueriedStore };
        live.CoverageGaps = new[] { FreshMerge.ThreadGapUnqueriedStore };
        return live;
    }
}
