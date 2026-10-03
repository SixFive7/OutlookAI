using System.Globalization;
using System.Text;
using OutlookAI.Core.Mapi;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>Which documented layout a folder's entry id has, read from its bytes alone.</summary>
public enum FolderEntryIdKind
{
    /// <summary>Neither layout below: a length or a fixed field that does not fit. Deliberately the ZERO value.</summary>
    Unknown = 0,

    /// <summary>
    /// [MS-PST] 2.4.3.2 (Mapping between EntryID and NID): 24 bytes - four flag bytes, the 16-byte
    /// provider UID that is the store's <c>PidTagRecordKey</c>, and the folder's 4-byte node id,
    /// little-endian. Every PST, and every OST that Outlook keeps in that format (IMAP, the older
    /// Outlook.com connector). [MS-PST] 2.2.2.6 (HEADER, <c>rgnid</c>): node ids are allocated by
    /// incrementing a per-type counter, so a deleted folder's node id is not handed out again.
    /// </summary>
    Pst = 1,

    /// <summary>
    /// [MS-OXCDATA] 2.2.4.1 (Folder EntryID): 46 bytes - four flag bytes, the 16-byte provider UID (a
    /// private mailbox's MailboxGuid), a 2-byte folder type, the 16-byte database GUID and 6-byte global
    /// counter of the folder's FID, and two pad bytes. Exchange, cached or online.
    /// </summary>
    ExchangeFolder = 2,
}

/// <summary>
/// A folder entry id taken apart along the documented layouts - Q114, 2026-10-03, for
/// <c>T2/LiveFolderIdentityTests</c>, which pins what each part does when a folder is renamed, moved
/// or recreated. Pure: no COM, no I/O. It never decides that an id is valid - Outlook does that - it
/// only names the parts, so a live test can say WHICH part changed.
/// </summary>
public sealed class FolderEntryIdLayout
{
    /// <summary>[MS-PST] 2.2.2.1: a NID's low five bits are its type; 0x02 is NID_TYPE_NORMAL_FOLDER.</summary>
    public const int NidTypeNormalFolder = 0x02;

    /// <summary>[MS-PST] 2.2.2.1: NID_TYPE_SEARCH_FOLDER.</summary>
    public const int NidTypeSearchFolder = 0x03;

    /// <summary>[MS-PST] 2.2.2.1: NID_TYPE_NORMAL_MESSAGE.</summary>
    public const int NidTypeNormalMessage = 0x04;

    /// <summary>[MS-OXCDATA] 2.2.4: eitLTPrivateFolder, the folder type of every folder in a private mailbox.</summary>
    public const int ExchangePrivateFolderType = 0x0001;

    /// <summary>[MS-OXCDATA] 2.2.4: eitLTPublicFolder.</summary>
    public const int ExchangePublicFolderType = 0x0003;

    private FolderEntryIdLayout(byte[] bytes, FolderEntryIdKind kind)
    {
        Bytes = bytes;
        Kind = kind;
    }

    /// <summary>The id's bytes.</summary>
    public byte[] Bytes { get; }

    /// <summary>Which layout the bytes fit.</summary>
    public FolderEntryIdKind Kind { get; }

    /// <summary>Length in bytes.</summary>
    public int Length => Bytes.Length;

    /// <summary>The four flag bytes as hex: 00000000 for every long-term id.</summary>
    public string FlagsHex => Hex(0, Math.Min(4, Bytes.Length));

    /// <summary>
    /// Bytes 4..19 as hex: the provider UID. On a PST it is the store's record key; on Exchange the
    /// mailbox GUID. Null when the id is shorter than 20 bytes.
    /// </summary>
    public string? ProviderUidHex => Bytes.Length >= 20 ? Hex(4, 16) : null;

    /// <summary>The part that names the folder WITHIN its store: the NID on a PST, the FID (database GUID plus global counter) on Exchange.</summary>
    public string? FolderPartHex => Kind switch
    {
        FolderEntryIdKind.Pst => Hex(20, 4),
        FolderEntryIdKind.ExchangeFolder => Hex(22, 22),
        _ => null,
    };

    /// <summary>A PST id's node-id type (low five bits of its first NID byte); null for any other layout.</summary>
    public int? NidType => Kind == FolderEntryIdKind.Pst ? Bytes[20] & 0x1F : null;

    /// <summary>An Exchange id's folder type (little-endian word at bytes 20..21); null for any other layout.</summary>
    public int? ExchangeFolderType => Kind == FolderEntryIdKind.ExchangeFolder ? Bytes[20] | (Bytes[21] << 8) : null;

    /// <summary>A PST id's node id as a number (little-endian at bytes 20..23); null for any other layout.</summary>
    public uint? Nid => Kind == FolderEntryIdKind.Pst ? BitConverter.ToUInt32(Bytes, 20) : null;

    /// <summary>
    /// The value the Windows Search index keeps in <c>System.ProviderItemID</c> for the object a PST id
    /// names: <c>N</c> and its node id in ten decimal digits (<c>N0000032898</c> for node 0x8082) - MEASURED
    /// 2026-10-03 (Q114/Q115) on every folder row and item row compared on the indexed test guest. It is
    /// the "provider item ID" Microsoft documents in the blob a store pushes with each MAPI URL ("send only
    /// the provider item ID for folders"). Unique only within its store: two PSTs' Inboxes share it. Null
    /// for any other layout.
    /// </summary>
    public string? PstProviderItemId => Nid is uint nid ? ProviderItemIdOf(nid) : null;

    /// <summary>The <c>System.ProviderItemID</c> spelling of a node id: <c>N</c> and ten decimal digits.</summary>
    public static string ProviderItemIdOf(uint nid)
    {
        return "N" + nid.ToString("D10", CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Parses a hex entry id. Null for anything that is not even-length hex. A string that is hex but fits
    /// neither layout comes back as <see cref="FolderEntryIdKind.Unknown"/>, so a caller can still report
    /// its length and flags.
    /// </summary>
    public static FolderEntryIdLayout? Parse(string? entryIdHex)
    {
        if (!StoreHash.TryDecodeHex(entryIdHex, out byte[] bytes) || bytes.Length == 0)
        {
            return null;
        }

        return new FolderEntryIdLayout(bytes, Classify(bytes));
    }

    /// <summary>
    /// Every spelling under which a row's value could carry these bytes: hex in either case, base64, and
    /// the Windows Search URL encoding (<see cref="EntryIdCodec.EncodeBytes"/>, each byte as U+AC00 plus
    /// it) - the encoding the index already uses for an ITEM's id in its URL. A row that holds the
    /// folder's id in any of them is a row that could address the folder without Outlook.
    /// </summary>
    public static IReadOnlyList<string> Spellings(byte[] bytes)
    {
        ArgumentNullException.ThrowIfNull(bytes);
        string hex = Convert.ToHexString(bytes);
        return new[]
        {
            hex,
            hex.ToLowerInvariant(),
            Convert.ToBase64String(bytes),
            EntryIdCodec.EncodeBytes(bytes),
        };
    }

    /// <summary>
    /// Whether a value read from an index row carries <paramref name="needle"/>: a string through any of
    /// <see cref="Spellings"/>, a byte array by containing the bytes, an array element by element.
    /// Ordinal throughout - the hex spellings are listed in both cases rather than compared loosely.
    /// </summary>
    public static bool ValueCarries(object? value, byte[] needle)
    {
        ArgumentNullException.ThrowIfNull(needle);
        if (value == null || needle.Length == 0)
        {
            return false;
        }

        if (value is byte[] blob)
        {
            return IndexOf(blob, needle) >= 0;
        }

        if (value is string text)
        {
            foreach (string spelling in Spellings(needle))
            {
                if (text.Contains(spelling, StringComparison.Ordinal))
                {
                    return true;
                }
            }

            return false;
        }

        if (value is System.Collections.IEnumerable many)
        {
            foreach (object? element in many)
            {
                if (ValueCarries(element, needle))
                {
                    return true;
                }
            }

            return false;
        }

        return ValueCarries(Convert.ToString(value, CultureInfo.InvariantCulture), needle);
    }

    /// <summary>Two hex entry ids as the same bytes - case does not matter, length does.</summary>
    public static bool SameBytes(string? a, string? b)
    {
        return StoreHash.TryDecodeHex(a, out byte[] x)
            && StoreHash.TryDecodeHex(b, out byte[] y)
            && x.AsSpan().SequenceEqual(y);
    }

    /// <summary>One line naming the parts, for a test's output.</summary>
    public string Describe()
    {
        StringBuilder sb = new StringBuilder();
        sb.Append(Kind).Append(", ").Append(Length).Append(" bytes, flags ").Append(FlagsHex);
        switch (Kind)
        {
            case FolderEntryIdKind.Pst:
                sb.Append(", nid type 0x").Append(NidType!.Value.ToString("X2", CultureInfo.InvariantCulture));
                break;
            case FolderEntryIdKind.ExchangeFolder:
                sb.Append(", folder type 0x").Append(ExchangeFolderType!.Value.ToString("X4", CultureInfo.InvariantCulture));
                break;
        }

        return sb.ToString();
    }

    private static FolderEntryIdKind Classify(byte[] bytes)
    {
        bool zeroFlags = bytes.Length >= 4 && bytes[0] == 0 && bytes[1] == 0 && bytes[2] == 0 && bytes[3] == 0;
        if (!zeroFlags)
        {
            return FolderEntryIdKind.Unknown;
        }

        if (bytes.Length == 24)
        {
            return FolderEntryIdKind.Pst;
        }

        if (bytes.Length == 46 && bytes[44] == 0 && bytes[45] == 0)
        {
            return FolderEntryIdKind.ExchangeFolder;
        }

        return FolderEntryIdKind.Unknown;
    }

    private static int IndexOf(byte[] haystack, byte[] needle)
    {
        return haystack.AsSpan().IndexOf(needle);
    }

    private string Hex(int offset, int count)
    {
        return Convert.ToHexString(Bytes, offset, count);
    }
}
