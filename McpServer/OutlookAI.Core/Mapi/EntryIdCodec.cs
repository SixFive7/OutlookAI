using System;
using System.Text;

namespace OutlookAI.Core.Mapi
{
    /// <summary>
    /// Pure logic for the Windows Search MAPI URL byte encoding (v3.MD sections 4/12).
    ///
    /// The Windows Search index stores each Outlook item's MAPI EntryID (and attachment
    /// IDs) inside the item URL, encoded per Microsoft's documented algorithm: every byte
    /// <c>b</c> becomes the WCHAR <c>b + 0xAC00</c> (which lands in the Hangul block).
    /// Decoding subtracts 0xAC00 per character.
    ///
    /// Message EntryIDs in OST/PST stores are exactly 24 bytes: 4 flag bytes,
    /// a 16-byte store UID, and a 4-byte node id (NID, little-endian). The first decoded
    /// byte observed in the wild (and in Microsoft's own documentation example) is the
    /// marker 0xEF rather than 0x00 - it must be treated as an encoding artifact and the
    /// four flag bytes rebuilt as 00 00 00 00 before handing the id to
    /// <c>Namespace.GetItemFromID</c>. Every decode must be verified on open
    /// (Subject/ReceivedTime comparison) with the ItemPathDisplay fallback as safety net.
    ///
    /// This type is host-neutral and side-effect free: no COM, no I/O.
    /// </summary>
    public static class EntryIdCodec
    {
        /// <summary>Offset added to each byte by the Windows Search MAPI URL encoding.</summary>
        public const int EncodingOffset = 0xAC00;

        /// <summary>Length in bytes of an OST/PST message EntryID (flags + store UID + NID).</summary>
        public const int MessageEntryIdLength = 24;

        /// <summary>
        /// Encodes raw bytes into the Windows Search MAPI URL representation
        /// (byte -> WCHAR b + 0xAC00). Used to fabricate deterministic synthetic test
        /// fixtures - live-recorded EntryIDs must never be committed (v3.MD S6).
        /// </summary>
        public static string EncodeBytes(byte[] bytes)
        {
            if (bytes == null)
            {
                throw new ArgumentNullException(nameof(bytes));
            }

            StringBuilder sb = new StringBuilder(bytes.Length);
            foreach (byte b in bytes)
            {
                sb.Append((char)(EncodingOffset + b));
            }

            return sb.ToString();
        }

        /// <summary>
        /// Decodes a Windows Search encoded segment back into raw bytes. Returns false if
        /// any character falls outside the encoded byte range [0xAC00, 0xACFF].
        /// </summary>
        public static bool TryDecodeEncodedBytes(string? encoded, out byte[] bytes)
        {
            bytes = Array.Empty<byte>();
            if (string.IsNullOrEmpty(encoded))
            {
                return false;
            }

            byte[] result = new byte[encoded!.Length];
            for (int i = 0; i < encoded.Length; i++)
            {
                int value = encoded[i] - EncodingOffset;
                if (value < 0 || value > 0xFF)
                {
                    return false;
                }

                result[i] = (byte)value;
            }

            bytes = result;
            return true;
        }

        /// <summary>
        /// Decodes the encoded tail segment of a message URL into an openable EntryID.
        /// Enforces the 24-byte message layout and rebuilds the four flag bytes as zeros
        /// (the observed first byte is the 0xEF encoding marker, not a real flag).
        /// </summary>
        public static bool TryDecodeMessageEntryId(string? encodedTail, out DecodedEntryId? decoded)
        {
            decoded = null;
            if (!TryDecodeEncodedBytes(encodedTail, out byte[] bytes) || bytes.Length != MessageEntryIdLength)
            {
                return false;
            }

            byte rawFirstByte = bytes[0];
            string storeUidHex = ToHex(bytes, 4, 16);
            string nidHex = ToHex(bytes, 20, 4);
            int nidLowFiveBits = bytes[20] & 0x1F;

            // Rebuild the flag bytes as 00 00 00 00: OST/PST EntryIDs carry zero flags,
            // and the URL encoding replaces the first byte with a marker (0xEF observed).
            string entryIdHex = "00000000" + storeUidHex + nidHex;

            decoded = new DecodedEntryId(entryIdHex, storeUidHex, nidHex, nidLowFiveBits, rawFirstByte);
            return true;
        }

        /// <summary>
        /// Decodes an attachment id segment (the part between <c>/at=</c> and <c>:</c> in an
        /// attachment URL). Attachment ids are variable-length (4 bytes observed).
        /// </summary>
        public static bool TryDecodeAttachmentId(string? encodedAttachmentId, out byte[] bytes)
        {
            return TryDecodeEncodedBytes(encodedAttachmentId, out bytes);
        }

        private static string ToHex(byte[] bytes, int offset, int count)
        {
            StringBuilder sb = new StringBuilder(count * 2);
            for (int i = offset; i < offset + count; i++)
            {
                sb.Append(bytes[i].ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
            }

            return sb.ToString();
        }
    }

    /// <summary>Result of decoding a message URL tail into an openable MAPI EntryID.</summary>
    public sealed class DecodedEntryId
    {
        internal DecodedEntryId(string entryIdHex, string storeUidHex, string nidHex, int nidLowFiveBits, byte rawFirstByte)
        {
            EntryIdHex = entryIdHex;
            StoreUidHex = storeUidHex;
            NidHex = nidHex;
            NidLowFiveBits = nidLowFiveBits;
            RawFirstByte = rawFirstByte;
        }

        /// <summary>
        /// 48-hex-char EntryID with flag bytes rebuilt as zeros - the exact string to pass
        /// to <c>Namespace.GetItemFromID(entryIdHex, storeId)</c>.
        /// </summary>
        public string EntryIdHex { get; }

        /// <summary>32-hex-char store UID (EntryID bytes 4..19); constant per store.</summary>
        public string StoreUidHex { get; }

        /// <summary>8-hex-char node id (EntryID bytes 20..23, little-endian NID).</summary>
        public string NidHex { get; }

        /// <summary>
        /// Low 5 bits of the NID's first byte; 0x04 (NID_TYPE_NORMAL_MESSAGE per [MS-PST])
        /// on every message sampled. Informational - not enforced.
        /// </summary>
        public int NidLowFiveBits { get; }

        /// <summary>Raw first decoded byte before the flag rebuild (0xEF marker observed).</summary>
        public byte RawFirstByte { get; }
    }
}
