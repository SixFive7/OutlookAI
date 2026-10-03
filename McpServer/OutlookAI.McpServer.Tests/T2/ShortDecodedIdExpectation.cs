namespace OutlookAI.McpServer.Tests.T2;

/// <summary>What kind of store a decoded short id is being tried on, as far as that question is concerned.</summary>
public enum DecodedIdStoreFormat
{
    /// <summary>The store's type could not be read. Deliberately the ZERO value: no expectation applies.</summary>
    Unknown = 0,

    /// <summary>A cached Exchange mailbox - primary or delegate - whose object model hands out 70-byte Exchange ids.</summary>
    CachedExchange = 1,

    /// <summary>A PST (<c>OlExchangeStoreType.olNotExchange</c>), whose message ids ARE the 24-byte short form.</summary>
    Pst = 2,

    /// <summary>An Exchange store that is not cached locally - not in the local index at all (D22/D25).</summary>
    OnlineExchange = 3,
}

/// <summary>What <c>Namespace.GetItemFromID</c> is expected to do with a decoded 24-byte short id.</summary>
public enum DecodedIdOutcome
{
    /// <summary>No expectation: the store's kind is unknown or not locally indexed.</summary>
    NotApplicable = 0,

    /// <summary>Refused with <see cref="ShortDecodedIdExpectation.InvalidEntryIdHResult"/> - the Phase-1 finding.</summary>
    RejectedAsAnInvalidEntryId = 1,

    /// <summary>Opens, and opens AS the item: the decoded id is the item's own entry id.</summary>
    OpensAsTheItemItself = 2,
}

/// <summary>
/// The two halves of the short-decoded-id check, decided in one pure place (Q74 C1 + C3, 2026-10-03).
/// <para>
/// <b>Why two halves.</b> <c>LiveDecodeVerifyTests.ShortDecodedId_IsRejectedByGetItemFromID_DiscoveryRecorded</c>
/// pins a Phase-1 platform finding that holds on CACHED EXCHANGE only: there the object model exposes
/// 70-byte Exchange entry ids, and <c>GetItemFromID</c> refuses the 24-byte id decoded from the index URL
/// with 0x80040107. On a PST the same 24 bytes are the entry-id format itself (flags, the store's record
/// key, the node id - <c>Mapi/EntryIdCodec.cs</c>), so there they should OPEN. The old single test therefore
/// failed on every test guest and could run only on the read-only workstation; it now says
/// <c>Requires=CachedExchange</c>, and a second test asserts the PST half on the guests.
/// </para>
/// <para>
/// <b>The PST half is INFERRED, not run.</b> It follows from the codec and from the Q92 measurement that a
/// PST's store UID in the URL is its record key (<c>Docs/live-tier-on-the-vm.md</c> section 8 item 24),
/// but no live run has tried a decoded id on a PST yet. It is written as an explicit assertion so the first
/// guest live run confirms it or fails it, loudly - never as an observation that passes either way.
/// </para>
/// </summary>
public static class ShortDecodedIdExpectation
{
    /// <summary>The hresult the cached-Exchange half expects, as the error text carries it.</summary>
    public const string InvalidEntryIdHResult = "80040107";

    /// <summary><c>OlExchangeStoreType.olNotExchange</c>.</summary>
    public const int NotExchange = 3;

    /// <summary>Classifies a store from what <c>OutlookComSession.GetStoreDetails</c> reports.</summary>
    /// <param name="exchangeStoreType">Raw <c>Store.ExchangeStoreType</c>; null when unreadable.</param>
    /// <param name="isCachedExchange">Raw <c>Store.IsCachedExchange</c>; null when unreadable.</param>
    public static DecodedIdStoreFormat FormatOf(int? exchangeStoreType, bool? isCachedExchange)
    {
        if (exchangeStoreType == NotExchange)
        {
            return DecodedIdStoreFormat.Pst;
        }

        if (exchangeStoreType == null || isCachedExchange == null)
        {
            return DecodedIdStoreFormat.Unknown;
        }

        return isCachedExchange.Value ? DecodedIdStoreFormat.CachedExchange : DecodedIdStoreFormat.OnlineExchange;
    }

    /// <summary>What a decoded short id must do on a store of <paramref name="format"/>.</summary>
    public static DecodedIdOutcome ExpectedOutcome(DecodedIdStoreFormat format)
    {
        return format switch
        {
            DecodedIdStoreFormat.CachedExchange => DecodedIdOutcome.RejectedAsAnInvalidEntryId,
            DecodedIdStoreFormat.Pst => DecodedIdOutcome.OpensAsTheItemItself,
            _ => DecodedIdOutcome.NotApplicable,
        };
    }
}
