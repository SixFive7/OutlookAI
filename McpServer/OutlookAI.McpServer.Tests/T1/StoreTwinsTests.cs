using System.Text.Json;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins one mailbox in two store entries (decided 2026-11-18 under the maintainer's Q131/Q132 delegation, "most
/// stable"): an ID never resolves to two stores. Measured on the Exchange guest the same day
/// (Testbed/guest/Measure-FolderIdResolution.ps1): the test account's own cached store and the online entry Exchange
/// automapped beside it carry ONE StoreID, byte for byte, and top folders with ONE provider UID - so the account's
/// storeId named both, list_folders walked both trees and find_folder offered two Inboxes (runbook 4.4a, run 6).
/// The mailbox is now resolved once, through its kept (cached, indexed) entry. Every id here is synthetic.
/// </summary>
public sealed class StoreTwinsTests
{
    private const string MailboxUid = "A1A1A1A1A1A1A1A1A1A1A1A1A1A1A1A1";
    private const string OtherUid = "B2B2B2B2B2B2B2B2B2B2B2B2B2B2B2B2";

    /// <summary>A 46-byte Exchange top-folder entry id carrying <paramref name="uid"/> as its provider UID.</summary>
    private static string ExchangeRoot(string uid) => "00000000" + uid + "0100" + new string('2', 32) + "000000000001" + "0000";

    private static readonly string SharedStoreId = "0000000038A1BB1005E5101AA1BB08002B2A56C20000454D534D44422E444C4C00000000" + new string('C', 40);

    private static ComStoreDetail Primary() =>
        new("Primary", "0000000038A1BB1005E5101AA1BB08002B2A56C2" + new string('D', 40), 0, true, rootFolderEntryId: ExchangeRoot(OtherUid));

    /// <summary>The account's own cached store of the mailbox (ExchangeStoreType 4).</summary>
    private static ComStoreDetail CachedEntry() =>
        new("account@example.invalid", SharedStoreId, 4, true, rootFolderEntryId: ExchangeRoot(MailboxUid));

    /// <summary>The online entry Exchange automapped (ExchangeStoreType 1, not cached), same StoreID, same UID.</summary>
    private static ComStoreDetail OnlineEntry() =>
        new("Account Display Name", SharedStoreId, 1, false, rootFolderEntryId: ExchangeRoot(MailboxUid));

    [Fact]
    public void TheMeasuredPair_IsOneMailbox_KeptThroughItsCachedEntry()
    {
        IReadOnlyList<ComStoreDetail> stores = new[] { Primary(), CachedEntry(), OnlineEntry() };

        Assert.Equal(new[] { -1, -1, 1 }, StoreTwins.KeptEntryOf(stores));
        Assert.True(StoreTwins.IsOtherEntry(stores, 2));
        Assert.False(StoreTwins.IsOtherEntry(stores, 1));

        // Whatever order the profile lists them in: the cached entry is kept.
        IReadOnlyList<ComStoreDetail> reversed = new[] { OnlineEntry(), Primary(), CachedEntry() };
        Assert.Equal(new[] { 2, -1, -1 }, StoreTwins.KeptEntryOf(reversed));
    }

    [Fact]
    public void TwoEntriesAreOneMailbox_ByTheirTopFoldersUid_OrByTheirStoreIdWhenATopFolderIsUnread()
    {
        Assert.True(StoreTwins.SameMailbox(CachedEntry(), OnlineEntry()));

        // A top folder that would not read: the StoreID, byte for byte, decides - in either case of hex.
        ComStoreDetail unread = new("Account Display Name", SharedStoreId.ToLowerInvariant(), 1, false);
        Assert.True(StoreTwins.SameMailbox(CachedEntry(), unread));

        // Control: another mailbox, another UID; and two stores with nothing to compare are never one.
        Assert.False(StoreTwins.SameMailbox(CachedEntry(), Primary()));
        Assert.False(StoreTwins.SameMailbox(new ComStoreDetail("a", string.Empty, 3, null), new ComStoreDetail("b", string.Empty, 3, null)));
    }

    [Fact]
    public void TheKeptEntry_IsTheLocalOne_ThenNotADelegates_ThenTheFirst()
    {
        // Both cached (Download shared folders on): the account's own store, not the delegate's entry.
        ComStoreDetail cachedDelegate = new("Account Display Name", SharedStoreId, 1, true, rootFolderEntryId: ExchangeRoot(MailboxUid));
        Assert.Equal(new[] { 1, -1 }, StoreTwins.KeptEntryOf(new[] { cachedDelegate, CachedEntry() }));

        // Neither cached, neither a delegate's: the first.
        ComStoreDetail onlineA = new("a", SharedStoreId, 4, false, rootFolderEntryId: ExchangeRoot(MailboxUid));
        ComStoreDetail onlineB = new("b", SharedStoreId, 4, false, rootFolderEntryId: ExchangeRoot(MailboxUid));
        Assert.Equal(new[] { -1, 0 }, StoreTwins.KeptEntryOf(new[] { onlineA, onlineB }));
    }

    [Fact]
    public void StoresOfTheirOwn_AreNobodysOtherEntry()
    {
        IReadOnlyList<ComStoreDetail> stores = new[] { Primary(), CachedEntry() };
        Assert.Equal(new[] { -1, -1 }, StoreTwins.KeptEntryOf(stores));
        Assert.Empty(StoreTwins.KeptEntryOf(Array.Empty<ComStoreDetail>()));
    }

    // ============================================================ store IDs resolve once

    [Fact]
    public void TheMailboxsStoreId_ResolvesToItsCachedEntry_NeverRefusedAsTwoStores()
    {
        IReadOnlyList<ComStoreDetail> stores = new[] { Primary(), OnlineEntry(), CachedEntry() };
        string id = OutlookIds.StoreIdOf(ExchangeRoot(MailboxUid), SharedStoreId)!;
        Assert.True(OutlookIds.TryDecodeStoreId(id, "store_id", out StoreIdParts? parts, out _));

        // Both entries carry the UID; the ID names the mailbox and lands on the cached entry.
        Assert.Same(stores[2], MailService.PickStore(stores, parts!, "store_id"));

        // The T form - the StoreID itself, which both entries share - lands there too.
        Assert.True(OutlookIds.TryDecodeStoreId(OutlookIds.StoreIdOf(null, SharedStoreId), "store_id", out StoreIdParts? byStoreId, out _));
        Assert.Same(stores[2], MailService.PickStore(stores, byStoreId!, "store_id"));
    }

    [Fact]
    public void BothEntriesShowOneStoreId_TheS28Form()
    {
        // Option (b): the online entry's top folder is read too, so its ID is the mailbox's S ID, not T.
        Assert.Equal(MailService.StoreIdOf(CachedEntry()), MailService.StoreIdOf(OnlineEntry()));
        Assert.Equal(OutlookIds.StoreUidIdLength, MailService.StoreIdOf(OnlineEntry())!.Length);
        Assert.StartsWith("S", MailService.StoreIdOf(OnlineEntry()), StringComparison.Ordinal);
    }

    [Fact]
    public void TwoMailboxesAnsweringOneId_AreStillRefused_NeverPicked()
    {
        // Two stores of two mailboxes cannot share a UID; a T ID two different mailboxes' stores carry (not a real
        // shape, a guard): their kept entries differ, so nothing is picked.
        ComStoreDetail a = new("a", SharedStoreId, 0, true, rootFolderEntryId: ExchangeRoot(MailboxUid));
        ComStoreDetail b = new("b", SharedStoreId, 0, true, rootFolderEntryId: ExchangeRoot(OtherUid));
        IReadOnlyList<ComStoreDetail> stores = new[] { a, b };
        Assert.True(OutlookIds.TryDecodeStoreId(OutlookIds.StoreIdOf(null, SharedStoreId), "store_id", out StoreIdParts? parts, out _));

        ArgumentException refused = Assert.Throws<ArgumentException>(() => MailService.PickStore(stores, parts!, "store_id"));
        Assert.Contains("answer to the ID", refused.Message, StringComparison.Ordinal);
    }

    // ============================================================ decision 207 (b): a second view, not a degradation

    [Fact]
    public void ASecondView_ClaimsNothingInTheIndex_AndIsNeverAStoreToCover()
    {
        // The kept entry is tied by its hash; the second view claims nothing, whatever its kind would be.
        StoreHash.TryDecodeHex("0A0B0C0D", out byte[] signature);
        StoreScopeInfo root = StoreScopeInfo.FromStorePrefix(
            "mapi16://{S-1-5-21-1-2-3-1000}/account@example.invalid($" + StoreHash.Compute(signature).ToString("x", System.Globalization.CultureInfo.InvariantCulture) + ")")!;
        StoreIndexIdentity kept = new("account@example.invalid", 4, StoreHash.Candidates("00112233", true, false, "0A0B0C0D", null, null));
        StoreIndexIdentity view = new("Account Display Name", 1, StoreHash.Candidates("00112233", true, false, "0A0B0C0D", null, null),
            secondViewOf: "account@example.invalid");

        StoreIndexMap map = StoreIndexMatcher.Match(new[] { kept, view }, new[] { root });

        Assert.Equal(StoreIndexMatchKind.StoreHash, map.Stores[0].Kind);
        Assert.Equal(StoreIndexMatchKind.SecondView, map.Stores[1].Kind);
        Assert.Null(map.Stores[1].Root);
        Assert.Empty(map.Unclaimed);

        // outlook_health's row: a second view of an indexed mailbox, named, not a store the index lacks.
        StoreStaleness row = MailService.DescribeStoreMatch(map.Stores[1], null);
        Assert.Equal("secondView", row.MatchedBy);
        Assert.StartsWith("a second view of an indexed mailbox: the same mailbox is the store 'account@example.invalid'", row.MatchNote, StringComparison.Ordinal);
        Assert.Null(row.IndexStore);
    }

    [Fact]
    public void TheStoresToCover_AreEachMailboxOnce()
    {
        IReadOnlyList<ComStoreDetail> stores = new[] { Primary(), OnlineEntry(), CachedEntry() };
        Assert.Equal(new[] { "Primary", "account@example.invalid" }, MailService.WithoutSecondViews(stores).Select(d => d.DisplayName));

        // An online store with no cached twin is a store of its own - covered, and its own 28-character ID.
        ComStoreDetail alone = new("Shared Online", "0000000038A1BB1005E5101AA1BB08002B2A56C2" + new string('E', 40), 1, false,
            rootFolderEntryId: ExchangeRoot("C3C3C3C3C3C3C3C3C3C3C3C3C3C3C3C3"));
        IReadOnlyList<ComStoreDetail> withAlone = new[] { Primary(), alone };
        Assert.Equal(2, MailService.WithoutSecondViews(withAlone).Count);
        Assert.Equal(28, MailService.StoreIdOf(alone)!.Length);
        Assert.NotEqual(MailService.StoreIdOf(Primary()), MailService.StoreIdOf(alone));
    }

    [Fact]
    public void ListAccountsNamesTheMailboxASecondViewShows_UnderItsWireName()
    {
        string json = JsonSerializer.Serialize(new StoreView { DisplayName = "v", SecondViewOf = "account@example.invalid" },
            new JsonSerializerOptions(JsonSerializerDefaults.Web));
        Assert.Contains("\"secondViewOf\":\"account@example.invalid\"", json, StringComparison.Ordinal);
    }

    [Fact]
    public void ALookupByStoreId_LandsOnTheKeptEntry()
    {
        IReadOnlyList<ComStoreDetail> stores = new[] { Primary(), OnlineEntry(), CachedEntry() };
        Assert.Same(stores[2], MailService.StoreWithEntryId(stores, SharedStoreId));
        Assert.Same(stores[2], MailService.StoreWithEntryId(stores, SharedStoreId.ToLowerInvariant()));
        Assert.Same(stores[0], MailService.StoreWithEntryId(stores, stores[0].StoreId));
        Assert.Null(MailService.StoreWithEntryId(stores, "00"));
        Assert.Null(MailService.StoreWithEntryId(null, SharedStoreId));
    }
}
