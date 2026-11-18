using System.Reflection;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tests.T1;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// The one-mailbox profile the suites around the sweep and the exhaustive scan search: <c>alice@example.com</c>,
/// indexed under one <c>mapi16</c> prefix whose freshness frontier stands at a fixed instant, and what they
/// fabricate around it - a mail, the freshness sweep that finds it or finds nothing, an exhaustive search of
/// its Inbox, the index, and a session that answers the store list, the sweep and the scan. The one copy
/// (Q154): <c>ScanDepthAndSweepScopeTests</c>, <c>TierSemanticsReportingTests</c> and
/// <c>SearchCoverageClaimTests</c> search this profile as it is, through <c>using static</c>;
/// <c>SweepBodyCapTests</c>, <c>SweepCacheKeyTests</c> and <c>ResumableScanTests</c>, whose profiles add a
/// store or index it under another prefix, name the parts they share.
/// </summary>
internal static class AliceMailbox
{
    /// <summary>The user's SID in the index's store URLs.</summary>
    public const string Sid = "{S-1-5-21-1111111111-2222222222-3333333333-1001}";

    /// <summary>The mailbox's store display name.</summary>
    public const string Store = "alice@example.com";

    /// <summary>The URL prefix the index files the mailbox's rows under.</summary>
    public const string StorePrefix = "mapi16://" + Sid + "/" + Store + "($deadbeef)";

    /// <summary>The index's freshness frontier, and when a mail of this profile arrived.</summary>
    public static readonly DateTime Frontier = new(2026, 8, 18, 9, 30, 0, DateTimeKind.Utc);

    /// <summary>The profile's store list: the mailbox, alone.</summary>
    public static readonly IReadOnlyList<ComStoreDetail> ProfileStores = new[] { new ComStoreDetail(Store, "store-alice", 0, true) };

    /// <summary>
    /// A mail in the mailbox's Inbox, from <paramref name="sender"/>, that arrived at <paramref name="received"/> -
    /// the <see cref="Frontier"/> unless a suite whose index stands elsewhere says otherwise.
    /// </summary>
    public static ComMailBrief Mail(string entryId, string sender = "bob@example.com", DateTime? received = null)
    {
        return new ComMailBrief(
            entryId: entryId,
            storeDisplayName: Store,
            storeId: "store-alice",
            folderName: "Inbox",
            folderKind: "inbox",
            subject: "a test mail",
            senderName: "Bob",
            senderAddress: sender,
            receivedTime: received ?? Frontier,
            isRead: true,
            hasAttachments: false,
            sizeBytes: 2048,
            body: "test body");
    }

    /// <summary>A sweep that reaches the store and finds one mail.</summary>
    public static ComSweepResult Sweep(string? onlyStore)
    {
        return new ComSweepResult(
            new[] { Mail("AA1") },
            foldersSwept: 4,
            foldersSkipped: 0,
            sweptFolders: new[] { Store + "/Inbox" },
            perStore: new[] { new ComStoreSweepCounters(Store, 4, 0, 0, 0) });
    }

    /// <summary>A sweep that ran, covered its scope and found no arrival at all.</summary>
    public static ComSweepResult EmptySweep(string? onlyStore)
    {
        return new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 4,
            foldersSkipped: 0,
            sweptFolders: new[] { Store + "/Inbox" },
            perStore: new[] { new ComStoreSweepCounters(Store, 4, 0, 0, 0) });
    }

    /// <summary>An exhaustive search of the mailbox's Inbox, from <paramref name="from"/> when it is given.</summary>
    public static SearchRequest ExhaustiveRequest(string? from = null)
    {
        return new SearchRequest
        {
            Query = "test",
            StoreId = TestFolders.StoreId(Store),

            // An exhaustive scan refuses to run unbounded, so the request carries the bound
            // its own validation demands.
            FolderId = TestFolders.FolderId(Store, "Inbox"),
            Exhaustive = true,
            From = from,
            Top = 25,
            SnippetChars = 0,
        };
    }

    /// <summary>The index of the mailbox, answering both probes for whatever they name.</summary>
    public static StubIndexClient Index() => new(new[] { StorePrefix }, Frontier, everyScope: true);

    /// <summary>
    /// An index that holds the profile's store alone, whose search answers <paramref name="searchRows"/> or nothing:
    /// what four suites around an unindexed store each built.
    /// </summary>
    public static StubIndexClient OnlyAliceIndexed(IReadOnlyList<IReadOnlyDictionary<string, object?>>? searchRows = null) =>
        new(new[] { StorePrefix }, Frontier) { SearchRows = searchRows ?? IndexRows.None };

    /// <summary>An index that holds no store, whose search answers <paramref name="searchRows"/> or nothing: three suites' own.</summary>
    public static StubIndexClient NothingIndexed(IReadOnlyList<IReadOnlyDictionary<string, object?>>? searchRows = null) =>
        new(Array.Empty<string>(), Frontier) { SearchRows = searchRows ?? IndexRows.None };

    /// <summary>
    /// The product over the profile: <paramref name="index"/> or <see cref="Index"/>, a session whose sweep is
    /// <paramref name="sweep"/> or <see cref="Sweep"/>, and whose exhaustive scan is <paramref name="scan"/> -
    /// or refused, when there is none.
    /// </summary>
    public static MailService Service(
        StubIndexClient? index = null,
        Func<string?, ComSweepResult>? sweep = null,
        Func<string?, ComExhaustiveResult>? scan = null)
    {
        return new MailService(
            new DirectGateway(StandInSession.Create(ProfileStores, sweep ?? Sweep, scan == null ? null
                // Argument 0 is storeDisplayName; see IOutlookSession.ExhaustiveScan.
                : standIn => standIn.Answering(nameof(IOutlookSession.ExhaustiveScan), args => scan(args?[0] as string)))),
            null,
            index ?? Index());
    }
}
