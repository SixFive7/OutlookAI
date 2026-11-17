using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OutlookAI.Core.Mapi
{
    /// <summary>
    /// Which documented layout a folder's entry id has, read from its bytes alone. The kind decides
    /// which constant bytes a folder ID leaves out.
    /// </summary>
    public enum FolderIdLayout
    {
        /// <summary>Neither layout below: every byte is kept. Deliberately the ZERO value.</summary>
        Raw = 0,

        /// <summary>
        /// [MS-PST] 2.4.3.2: 24 bytes - four zero flag bytes, the 16-byte provider UID that is the
        /// store's record key, and the folder's 4-byte node id, little-endian. Every PST, and every
        /// OST Outlook keeps in that format (IMAP, Outlook.com).
        /// </summary>
        Pst = 1,

        /// <summary>
        /// [MS-OXCDATA] 2.2.4.1: 46 bytes - four zero flag bytes, the 16-byte provider UID (the
        /// mailbox GUID), a 2-byte folder type, the 22-byte FID (database GUID and global counter)
        /// and two zero pad bytes. Exchange, cached or online.
        /// </summary>
        Exchange = 2,
    }

    /// <summary>What a store ID names (<see cref="OutlookIds.TryDecodeStoreId"/>).</summary>
    public sealed class StoreIdParts
    {
        internal StoreIdParts(string id, string? providerUidHex, string? storeEntryIdHex)
        {
            Id = id;
            ProviderUidHex = providerUidHex;
            StoreEntryIdHex = storeEntryIdHex;
        }

        /// <summary>The ID in its canonical spelling.</summary>
        public string Id { get; }

        /// <summary>
        /// The 16-byte provider UID every entry id in the store carries at bytes 4..19, as upper-case
        /// hex - a PST's record key, an Exchange mailbox's GUID. Null for a store ID built from the
        /// store's own entry id instead.
        /// </summary>
        public string? ProviderUidHex { get; }

        /// <summary>The store's own entry id (<c>Store.StoreID</c>) as hex, when the ID carries it; null otherwise.</summary>
        public string? StoreEntryIdHex { get; }
    }

    /// <summary>
    /// The IDs every OutlookAI tool addresses folders and stores by (Q131-Q136, decided by the
    /// maintainer 2026-10-11): names are for display only.
    /// <para>
    /// <b>A folder ID is the folder's own entry id</b> - <c>Folder.EntryID</c>, which research
    /// measured unique and stable: it survives a rename, a move within the mail file, deletion into
    /// Deleted Items and restarts, and is never handed out again (<c>Docs/live-tier-on-the-vm.md</c>
    /// section 8 item 31) - <b>re-encoded shorter, with a check character</b> (Q131 (b)). Nothing is
    /// stored and there is no table: the ID IS the entry id, so any tool can open the folder from it
    /// with <c>NameSpace.GetFolderFromID</c>, and the Windows Search index can be asked for it with
    /// Outlook closed (Q132).
    /// </para>
    /// <para>
    /// <b>The spelling</b>: one KIND character, the payload in Crockford's base32 (the digits and
    /// the upper-case letters without I, L, O and U, most significant bit first), and one CHECK
    /// character. The kind says which constant bytes the payload leaves out:
    /// </para>
    /// <list type="table">
    /// <item><term><c>P</c></term><description>a PST-format folder id (<see cref="FolderIdLayout.Pst"/>):
    /// its four zero flag bytes left out, 20 bytes kept - 34 characters.</description></item>
    /// <item><term><c>X</c></term><description>an Exchange folder id (<see cref="FolderIdLayout.Exchange"/>):
    /// its zero flags and zero pad left out, 40 bytes kept - 66 characters.</description></item>
    /// <item><term><c>R</c></term><description>any other folder id, every byte kept.</description></item>
    /// <item><term><c>S</c></term><description>a store: the 16-byte provider UID every entry id in
    /// it carries (bytes 4..19 of its root folder's id) - 28 characters. So a folder ID names its
    /// store as well, with no lookup: <see cref="StoreIdOfEntryId"/>.</description></item>
    /// <item><term><c>T</c></term><description>a store whose root folder id carries no provider UID:
    /// its own entry id (<c>Store.StoreID</c>), every byte kept.</description></item>
    /// </list>
    /// <para>
    /// <b>The check character</b> is the polynomial check over GF(32) (generator <c>x</c>, modulus
    /// <c>x^5 + x^2 + 1</c>, which is primitive): every symbol before it, the kind included, is a
    /// coefficient, and the check makes the whole ID evaluate to zero. Any ONE wrong character and
    /// any swap of two characters fewer than 31 places apart change the value, so a mistyped ID is
    /// refused rather than read as another folder; the fixed lengths of P, X and S catch a dropped
    /// or added character. It stays inside the 32 symbols, so an ID is letters and digits only.
    /// </para>
    /// <para>
    /// <b>Reading one back</b> is forgiving the way Crockford specifies: case does not matter,
    /// <c>I</c> and <c>L</c> read as <c>1</c>, <c>O</c> as <c>0</c>, and hyphens are ignored.
    /// Anything else that is not an ID - a folder name, a path, a raw hex entry id - is refused with
    /// a pointer to <c>find_folder</c> (Q135 (i)): <see cref="DescribeFolderIdProblem"/>.
    /// </para>
    /// <para>Pure: no COM, no I/O, net48 and net10 alike.</para>
    /// </summary>
    public static class OutlookIds
    {
        /// <summary>Crockford's base32 alphabet: digits, then the letters without I, L, O and U.</summary>
        public const string Alphabet = "0123456789ABCDEFGHJKMNPQRSTVWXYZ";

        /// <summary>Kind of a PST-format folder ID.</summary>
        public const char PstFolderKind = 'P';

        /// <summary>Kind of an Exchange folder ID.</summary>
        public const char ExchangeFolderKind = 'X';

        /// <summary>Kind of a folder ID that keeps every byte of its entry id.</summary>
        public const char RawFolderKind = 'R';

        /// <summary>Kind of a store ID built from the store's provider UID.</summary>
        public const char StoreUidKind = 'S';

        /// <summary>Kind of a store ID built from the store's own entry id.</summary>
        public const char RawStoreKind = 'T';

        /// <summary>Bytes of a provider UID (a MAPIUID).</summary>
        public const int ProviderUidLength = 16;

        /// <summary>Length of a PST-format folder ID: kind, 32 symbols, check.</summary>
        public const int PstFolderIdLength = 34;

        /// <summary>Length of an Exchange folder ID: kind, 64 symbols, check.</summary>
        public const int ExchangeFolderIdLength = 66;

        /// <summary>Length of a store ID built from a provider UID: kind, 26 symbols, check.</summary>
        public const int StoreUidIdLength = 28;

        private const int PstEntryIdLength = 24;

        private const int ExchangeEntryIdLength = 46;

        /// <summary>The longest entry id an ID may carry. Store entry ids run to a few hundred bytes; this is far above any.</summary>
        private const int MaxEntryIdBytes = 4096;

        // x^5 + x^2 + 1 as a bit mask, the modulus of the check's field.
        private const int FieldModulus = 0x25;

        /// <summary>Which layout an entry id's bytes fit (<see cref="FolderIdLayout"/>).</summary>
        public static FolderIdLayout Classify(byte[] entryId)
        {
            if (entryId == null)
            {
                throw new ArgumentNullException(nameof(entryId));
            }

            bool zeroFlags = entryId.Length >= 4 && entryId[0] == 0 && entryId[1] == 0 && entryId[2] == 0 && entryId[3] == 0;
            if (!zeroFlags)
            {
                return FolderIdLayout.Raw;
            }

            if (entryId.Length == PstEntryIdLength)
            {
                return FolderIdLayout.Pst;
            }

            if (entryId.Length == ExchangeEntryIdLength && entryId[44] == 0 && entryId[45] == 0)
            {
                return FolderIdLayout.Exchange;
            }

            return FolderIdLayout.Raw;
        }

        /// <summary>
        /// The folder ID of a folder entry id given as hex (<c>Folder.EntryID</c>), or null when the
        /// value is not hex. Upper- and lower-case hex give the same ID.
        /// </summary>
        public static string? FolderIdFromEntryId(string? entryIdHex)
        {
            return StoreHash.TryDecodeHex(entryIdHex, out byte[] bytes) && bytes.Length > 0 && bytes.Length <= MaxEntryIdBytes
                ? FolderIdFromEntryId(bytes)
                : null;
        }

        /// <summary>The folder ID of a folder entry id.</summary>
        public static string FolderIdFromEntryId(byte[] entryId)
        {
            if (entryId == null)
            {
                throw new ArgumentNullException(nameof(entryId));
            }

            switch (Classify(entryId))
            {
                case FolderIdLayout.Pst:
                    return Compose(PstFolderKind, Slice(entryId, 4, 20));
                case FolderIdLayout.Exchange:
                    return Compose(ExchangeFolderKind, Slice(entryId, 4, 40));
                default:
                    return Compose(RawFolderKind, entryId);
            }
        }

        /// <summary>
        /// The store ID every entry id of a store implies: <c>S</c> and the provider UID at bytes
        /// 4..19 - for a folder, an item or the store's root folder alike. Null when the value is
        /// not hex, is shorter than 20 bytes, or carries an all-zero UID.
        /// </summary>
        public static string? StoreIdOfEntryId(string? entryIdHex)
        {
            string? uid = ProviderUidHexOf(entryIdHex);
            return uid == null ? null : StoreIdFromProviderUid(uid);
        }

        /// <summary>
        /// The store ID of a store, from what Outlook reports for it: <c>S</c> and the provider UID
        /// its root folder's id carries when there is one, else <c>T</c> and the store's own entry
        /// id; null when neither reads as hex.
        /// </summary>
        public static string? StoreIdOf(string? rootFolderEntryIdHex, string? storeEntryIdHex)
        {
            string? byUid = StoreIdOfEntryId(rootFolderEntryIdHex);
            if (byUid != null)
            {
                return byUid;
            }

            return StoreHash.TryDecodeHex(storeEntryIdHex, out byte[] bytes) && bytes.Length > 0 && bytes.Length <= MaxEntryIdBytes
                ? Compose(RawStoreKind, bytes)
                : null;
        }

        /// <summary>The store ID for a provider UID given as 32 hex digits.</summary>
        public static string StoreIdFromProviderUid(string providerUidHex)
        {
            if (!StoreHash.TryDecodeHex(providerUidHex, out byte[] uid) || uid.Length != ProviderUidLength)
            {
                throw new ArgumentException("A provider UID is 16 bytes of hex.", nameof(providerUidHex));
            }

            return Compose(StoreUidKind, uid);
        }

        /// <summary>
        /// Bytes 4..19 of an entry id as upper-case hex - the provider UID (MAPIUID) that routes it to
        /// its store. Null when the value is not hex, is shorter than 20 bytes, or the UID is all zero.
        /// </summary>
        public static string? ProviderUidHexOf(string? entryIdHex)
        {
            if (!StoreHash.TryDecodeHex(entryIdHex, out byte[] bytes) || bytes.Length < 4 + ProviderUidLength)
            {
                return null;
            }

            bool allZero = true;
            for (int i = 4; i < 4 + ProviderUidLength; i++)
            {
                allZero &= bytes[i] == 0;
            }

            return allZero ? null : ToHex(bytes, 4, ProviderUidLength);
        }

        /// <summary>
        /// The value the Windows Search index keeps in <c>System.ProviderItemID</c> for the folder a
        /// PST-format entry id names: <c>N</c> and its node id in ten decimal digits
        /// (<c>N0000032898</c> for node 0x8082) - MEASURED 2026-10-03 (Q114/Q115) on every folder and
        /// item row compared on the indexed test guest. Unique only within its store. Null for any
        /// other layout.
        /// </summary>
        public static string? PstProviderItemId(byte[] entryId)
        {
            if (entryId == null || Classify(entryId) != FolderIdLayout.Pst)
            {
                return null;
            }

            uint nid = (uint)(entryId[20] | (entryId[21] << 8) | (entryId[22] << 16) | (entryId[23] << 24));
            return ProviderItemIdOf(nid);
        }

        /// <summary>The <c>System.ProviderItemID</c> spelling of a node id: <c>N</c> and ten decimal digits.</summary>
        public static string ProviderItemIdOf(uint nid)
        {
            return "N" + nid.ToString("D10", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Reads a folder ID back into the folder's entry id (upper-case hex). False, with
        /// <paramref name="problem"/> saying why in words an agent can act on, for anything that is
        /// not a folder ID - <see cref="DescribeFolderIdProblem"/>.
        /// </summary>
        public static bool TryDecodeFolderId(string? id, string parameterName, out string entryIdHex, out string? problem)
        {
            entryIdHex = string.Empty;
            if (!TryReadSymbols(id, out string canonical, out int[] values) || values.Length < 3)
            {
                problem = DescribeFolderIdProblem(parameterName, id);
                return false;
            }

            if (!CheckHolds(values))
            {
                problem = DescribeFolderIdProblem(parameterName, id);
                return false;
            }

            char kind = canonical[0];
            byte[]? payload = DecodePayload(values);
            byte[]? entryId = null;
            if (payload != null)
            {
                switch (kind)
                {
                    case PstFolderKind when canonical.Length == PstFolderIdLength && payload.Length == 20:
                        entryId = new byte[PstEntryIdLength];
                        Array.Copy(payload, 0, entryId, 4, 20);
                        break;
                    case ExchangeFolderKind when canonical.Length == ExchangeFolderIdLength && payload.Length == 40:
                        entryId = new byte[ExchangeEntryIdLength];
                        Array.Copy(payload, 0, entryId, 4, 40);
                        break;
                    case RawFolderKind when payload.Length > 0 && payload.Length <= MaxEntryIdBytes:
                        entryId = payload;
                        break;
                }
            }

            if (entryId == null)
            {
                problem = DescribeFolderIdProblem(parameterName, id);
                return false;
            }

            entryIdHex = ToHex(entryId, 0, entryId.Length);
            problem = null;
            return true;
        }

        /// <summary>
        /// Reads a store ID back. False, with <paramref name="problem"/> saying why, for anything that
        /// is not a store ID - <see cref="DescribeStoreIdProblem"/>.
        /// </summary>
        public static bool TryDecodeStoreId(string? id, string parameterName, out StoreIdParts? parts, out string? problem)
        {
            parts = null;
            if (!TryReadSymbols(id, out string canonical, out int[] values) || values.Length < 3 || !CheckHolds(values))
            {
                problem = DescribeStoreIdProblem(parameterName, id);
                return false;
            }

            byte[]? payload = DecodePayload(values);
            if (payload != null && canonical[0] == StoreUidKind && canonical.Length == StoreUidIdLength && payload.Length == ProviderUidLength)
            {
                parts = new StoreIdParts(canonical, ToHex(payload, 0, payload.Length), null);
            }
            else if (payload != null && canonical[0] == RawStoreKind && payload.Length > 0 && payload.Length <= MaxEntryIdBytes)
            {
                parts = new StoreIdParts(canonical, null, ToHex(payload, 0, payload.Length));
            }

            if (parts == null)
            {
                problem = DescribeStoreIdProblem(parameterName, id);
                return false;
            }

            problem = null;
            return true;
        }

        /// <summary>
        /// The canonical spelling of an ID as typed - upper case, Crockford's look-alikes mapped,
        /// hyphens gone - or null when it holds a character no ID has. Used to compare an ID a caller
        /// passed with one this server computed.
        /// </summary>
        public static string? Canonical(string? id)
        {
            return TryReadSymbols(id, out string canonical, out _) ? canonical : null;
        }

        /// <summary>
        /// Why <paramref name="input"/> is not a folder ID, as a refusal an agent can act on: a name or
        /// path gets the pointer to <c>find_folder</c> (Q135 (i)), a store ID is named as one, a raw
        /// entry id and a mistyped ID each get their own sentence.
        /// </summary>
        public static string DescribeFolderIdProblem(string parameterName, string? input)
        {
            string shown = Shown(input);
            string pointer = "Folders are addressed by ID; names and paths are for display only. " + FindFolderPointer(input)
                + " Pass that folderId here.";
            if (string.IsNullOrWhiteSpace(input))
            {
                return "'" + parameterName + "' is empty. " + pointer;
            }

            if (LooksLikeHexEntryId(input!))
            {
                return "'" + parameterName + "' takes a folder ID, and " + shown + " is a raw Outlook entry id in hex, not one. "
                    + pointer;
            }

            if (TryReadSymbols(input, out string canonical, out int[] values) && values.Length >= 3)
            {
                if (!CheckHolds(values))
                {
                    return "'" + parameterName + "' " + shown + " is not a valid ID: its last character is a check "
                        + "character and it does not match the rest, so a character was mistyped, lost or swapped. Copy the "
                        + "folderId exactly as list_folders or find_folder returned it.";
                }

                if (canonical[0] == StoreUidKind || canonical[0] == RawStoreKind)
                {
                    return "'" + parameterName + "' takes a folder ID, and " + shown + " is a STORE ID (a whole mailbox or "
                        + "data file). Pass it as store_id where a tool takes one, or get the folder's own ID from "
                        + "find_folder or list_folders.";
                }

                if (canonical[0] == PstFolderKind || canonical[0] == ExchangeFolderKind || canonical[0] == RawFolderKind)
                {
                    return "'" + parameterName + "' " + shown + " is not a complete folder ID: it is too short or too long "
                        + "for its kind, so part of it was lost or added. Copy the folderId exactly as list_folders or "
                        + "find_folder returned it.";
                }
            }

            return "'" + parameterName + "' takes a folder ID, not a folder name or path, and " + shown + " is not one. "
                + pointer;
        }

        /// <summary>
        /// The pointer every refusal of a folder name or path ends with (Q135 (i)): the find_folder call that
        /// gets the folder's ID - with the last part of a path-like value as the name to ask for.
        /// </summary>
        public static string FindFolderPointer(string? nameOrPath)
        {
            return "Call find_folder with the folder's name" + LeafHint(nameOrPath) + " to get its folderId, or "
                + "list_folders for the whole tree.";
        }

        /// <summary>Why <paramref name="input"/> is not a store ID, as a refusal an agent can act on.</summary>
        public static string DescribeStoreIdProblem(string parameterName, string? input)
        {
            string shown = Shown(input);
            string pointer = "Stores are addressed by ID; display names are for reading only. list_accounts lists every "
                + "store with its storeId (list_folders and find_folder show it too) - pass that storeId here.";
            if (string.IsNullOrWhiteSpace(input))
            {
                return "'" + parameterName + "' is empty. " + pointer;
            }

            if (TryReadSymbols(input, out string canonical, out int[] values) && values.Length >= 3)
            {
                if (!CheckHolds(values))
                {
                    return "'" + parameterName + "' " + shown + " is not a valid ID: its last character is a check "
                        + "character and it does not match the rest, so a character was mistyped, lost or swapped. Copy the "
                        + "storeId exactly as list_accounts returned it.";
                }

                if (canonical[0] == PstFolderKind || canonical[0] == ExchangeFolderKind || canonical[0] == RawFolderKind)
                {
                    return "'" + parameterName + "' takes a store ID, and " + shown + " is a FOLDER ID. A folder ID already "
                        + "names its store: pass it as folder_id where the tool takes one, or take the storeId that "
                        + "list_folders or find_folder shows beside the folder.";
                }

                if (canonical[0] == StoreUidKind || canonical[0] == RawStoreKind)
                {
                    return "'" + parameterName + "' " + shown + " is not a complete store ID: it is too short or too long "
                        + "for its kind. Copy the storeId exactly as list_accounts returned it.";
                }
            }

            return "'" + parameterName + "' takes a store ID, not a store's name or address, and " + shown + " is not one. "
                + pointer;
        }

        /// <summary>
        /// The ID in its canonical spelling plus its check character, for a kind and a payload - the
        /// one composer every kind goes through.
        /// </summary>
        private static string Compose(char kind, byte[] payload)
        {
            StringBuilder sb = new StringBuilder(2 + ((payload.Length * 8) + 4) / 5);
            sb.Append(kind);
            int buffer = 0;
            int bits = 0;
            foreach (byte b in payload)
            {
                buffer = (buffer << 8) | b;
                bits += 8;
                while (bits >= 5)
                {
                    bits -= 5;
                    sb.Append(Alphabet[(buffer >> bits) & 0x1F]);
                }

                buffer &= (1 << bits) - 1;
            }

            if (bits > 0)
            {
                sb.Append(Alphabet[(buffer << (5 - bits)) & 0x1F]);
            }

            int[] values = new int[sb.Length];
            for (int i = 0; i < sb.Length; i++)
            {
                values[i] = Alphabet.IndexOf(sb[i]);
            }

            sb.Append(Alphabet[CheckValue(values)]);
            return sb.ToString();
        }

        /// <summary>
        /// The check value for a run of symbols: the run evaluated as a polynomial at the field's
        /// generator, times the generator once more - so the run plus this value evaluates to zero.
        /// </summary>
        internal static int CheckValue(IReadOnlyList<int> values)
        {
            int r = 0;
            foreach (int v in values)
            {
                r = TimesGenerator(r) ^ v;
            }

            return TimesGenerator(r);
        }

        /// <summary>Whether a whole ID, its check character last, evaluates to zero.</summary>
        private static bool CheckHolds(IReadOnlyList<int> values)
        {
            int r = 0;
            foreach (int v in values)
            {
                r = TimesGenerator(r) ^ v;
            }

            return r == 0;
        }

        /// <summary>Multiplies a field element by the generator x, modulo x^5 + x^2 + 1.</summary>
        private static int TimesGenerator(int value)
        {
            value <<= 1;
            if ((value & 0x20) != 0)
            {
                value ^= FieldModulus;
            }

            return value & 0x1F;
        }

        /// <summary>
        /// The symbols of an ID as typed, Crockford's way: surrounding whitespace and hyphens dropped,
        /// case ignored, I and L as 1, O as 0. False for anything holding another character.
        /// </summary>
        private static bool TryReadSymbols(string? input, out string canonical, out int[] values)
        {
            canonical = string.Empty;
            values = Array.Empty<int>();
            if (input == null)
            {
                return false;
            }

            string trimmed = input.Trim();
            if (trimmed.Length == 0 || trimmed.Length > 2 + ((MaxEntryIdBytes * 8) + 4) / 5)
            {
                return false;
            }

            StringBuilder sb = new StringBuilder(trimmed.Length);
            List<int> read = new List<int>(trimmed.Length);
            foreach (char raw in trimmed)
            {
                if (raw == '-')
                {
                    continue;
                }

                char c = char.ToUpperInvariant(raw);
                if (c == 'I' || c == 'L')
                {
                    c = '1';
                }
                else if (c == 'O')
                {
                    c = '0';
                }

                int value = c < 128 ? Alphabet.IndexOf(c) : -1;
                if (value < 0)
                {
                    return false;
                }

                sb.Append(c);
                read.Add(value);
            }

            if (read.Count == 0)
            {
                return false;
            }

            canonical = sb.ToString();
            values = read.ToArray();
            return true;
        }

        /// <summary>
        /// The payload between the kind and the check as bytes - null when its length is not one a
        /// whole number of bytes produces, or when the padding bits are not zero.
        /// </summary>
        private static byte[]? DecodePayload(int[] values)
        {
            int symbols = values.Length - 2;
            int byteCount = (symbols * 5) / 8;
            if (byteCount == 0 || ((byteCount * 8) + 4) / 5 != symbols)
            {
                return null;
            }

            byte[] bytes = new byte[byteCount];
            int buffer = 0;
            int bits = 0;
            int written = 0;
            for (int i = 1; i <= symbols; i++)
            {
                buffer = (buffer << 5) | values[i];
                bits += 5;
                if (bits >= 8)
                {
                    bits -= 8;
                    if (written < byteCount)
                    {
                        bytes[written++] = (byte)((buffer >> bits) & 0xFF);
                    }

                    buffer &= (1 << bits) - 1;
                }
            }

            return written == byteCount && buffer == 0 ? bytes : null;
        }

        private static bool LooksLikeHexEntryId(string input)
        {
            string trimmed = input.Trim();
            if (trimmed.Length < 40 || (trimmed.Length % 2) != 0)
            {
                return false;
            }

            foreach (char c in trimmed)
            {
                bool hex = (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
                if (!hex)
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>The input quoted for a message, cut so a pasted document cannot flood the answer.</summary>
        private static string Shown(string? input)
        {
            if (input == null)
            {
                return "(nothing)";
            }

            string trimmed = input.Trim();
            return trimmed.Length <= 80 ? "'" + trimmed + "'" : "'" + trimmed.Substring(0, 77) + "...'";
        }

        /// <summary>
        /// " (for example \"2026\")" - the last '/'-separated part of a value that looks like a path,
        /// as the name to ask find_folder for; empty when the value is no use as one.
        /// </summary>
        private static string LeafHint(string? input)
        {
            if (string.IsNullOrWhiteSpace(input) || LooksLikeHexEntryId(input!))
            {
                return string.Empty;
            }

            string trimmed = input!.Trim().TrimEnd('/', '\\');
            int cut = Math.Max(trimmed.LastIndexOf('/'), trimmed.LastIndexOf('\\'));
            string leaf = cut >= 0 ? trimmed.Substring(cut + 1) : trimmed;
            if (leaf.Length == 0 || leaf.Length > 80)
            {
                return string.Empty;
            }

            return " (for example name: \"" + leaf + "\")";
        }

        private static byte[] Slice(byte[] source, int offset, int count)
        {
            byte[] slice = new byte[count];
            Array.Copy(source, offset, slice, 0, count);
            return slice;
        }

        private static string ToHex(byte[] bytes, int offset, int count)
        {
            StringBuilder sb = new StringBuilder(count * 2);
            for (int i = offset; i < offset + count; i++)
            {
                sb.Append(bytes[i].ToString("X2", CultureInfo.InvariantCulture));
            }

            return sb.ToString();
        }
    }
}
