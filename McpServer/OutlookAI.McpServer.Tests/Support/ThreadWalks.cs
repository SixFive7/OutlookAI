using Xunit;
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
    /// The rule both thread suites' store lists keep - the stores the walk could not ask (the scoped suite) and the
    /// stores without an index (the unindexed one): a store the <paramref name="list"/> names raises that list's
    /// <paramref name="gap"/> code and makes the lookup partial, never live. Each failure names the list.
    /// </summary>
    public static void AssertNamedStoresRaiseTheCode(string list, ThreadLiveInfo live, string gap, string walkedStore)
    {
        var gaps = FreshMerge.DescribeThreadCoverageGaps(live, new[] { walkedStore })!;
        Assert.True(
            gaps.Contains(gap),
            list + ": a store it names must raise " + gap + "; the walk's gaps were [" + string.Join(", ", gaps) + "].");
        var freshness = FreshMerge.ClassifyThreadFreshness(live, new[] { walkedStore });
        Assert.True(
            freshness == FreshMerge.FreshnessPartial,
            list + ": a store it names must make the lookup " + FreshMerge.FreshnessPartial + ", not " + freshness + ".");
    }

    /// <summary>
    /// What both thread suites' coverage sentence says of a store the walk could not cover: it names that store and
    /// the walked one, tells the agent to tell the user, and withdraws the stronger claim - "we could not ask",
    /// never "we asked and there was nothing there"; an agent that reports the second has invented a fact. Each
    /// failure names the <paramref name="list"/> the store came from.
    /// </summary>
    public static void AssertSentenceNamesTheStoresAndWithdrawsTheStrongerClaim(
        string list, string line, string walkedStore, string uncoveredStore)
    {
        foreach (string expected in new[] { uncoveredStore, walkedStore, "TELL THE USER", "cannot be established" })
        {
            Assert.True(
                line.Contains(expected, StringComparison.Ordinal),
                list + ": the coverage sentence must contain \"" + expected + "\". It was: " + line);
        }
    }

    /// <summary>
    /// The Q7b rule both thread suites' capped store lists keep: naming the first twelve of thirty and stopping reads
    /// as "these twelve", so the coverage line says how many more there were and names the field that holds the
    /// total. Each failure names the <paramref name="list"/>.
    /// </summary>
    public static void AssertCappedListSaysHowManyMore(string list, ThreadLiveInfo live, string? scopeStore, string totalField)
    {
        string line = Assert.Single(MailService.DescribeThreadCoverage(
            live, FreshMerge.FreshnessPartial, scopeStore, scopeWidened: false, top: 50)!);
        Assert.True(
            line.Contains("and 18 more", StringComparison.Ordinal),
            list + ": the coverage line must say how many more there were (\"and 18 more\"). It was: " + line);
        Assert.True(
            line.Contains(totalField, StringComparison.Ordinal),
            list + ": the coverage line must name the field that holds the total, " + totalField + ". It was: " + line);
    }

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
