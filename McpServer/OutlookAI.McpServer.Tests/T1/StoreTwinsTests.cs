using OutlookAI.Core.Com;
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
