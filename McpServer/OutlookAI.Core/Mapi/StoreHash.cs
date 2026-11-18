using System;
using System.Collections.Generic;
using System.Globalization;
using OutlookAI.Core.AddIn;

namespace OutlookAI.Core.Mapi
{
    /// <summary>
    /// The store hash Outlook writes into every Windows Search URL it pushes:
    /// <c>mapi16://{SID}/&lt;StoreDisplayName&gt;($&lt;hash&gt;)/...</c>. Microsoft documents it
    /// (<i>Algorithm to Calculate the Store Hash Number</i>, Outlook MAPI reference; the same
    /// code is MFCMAPI's <c>ComputeStoreHash</c>): <c>h = h * 33 + x</c> over the input's
    /// little-endian DWORDs, then over its trailing bytes, then - for a public-folder store -
    /// over the constant <c>'.PUB'</c>, then optionally over the characters of a file path.
    /// <para>
    /// WHICH INPUT, per Microsoft: the store's <c>PR_ENTRYID</c> "in most cases" - which is
    /// <c>Store.StoreID</c> byte for byte - and, for a cached Exchange store, the
    /// <c>PR_MAPPING_SIGNATURE</c> found in the profile. MEASURED for the first (Q92/Q99, Office
    /// LTSC 2024): eight PST stores on a test guest - Unicode and ANSI, renamed, copied to a new
    /// path, re-keyed by Outlook, mounted in two profiles - each carried exactly
    /// <c>Compute(Store.StoreID)</c>, two of them predicted from the file path before Outlook
    /// was asked. The Exchange half is documented and NOT measured here: no test guest can
    /// have Exchange (Docs/live-tier-on-the-vm.md section 8 item 24).
    /// </para>
    /// <para>
    /// Pure: no COM, no I/O. The hash is a 32-bit DJB-style value, so it identifies a store
    /// among the handful one Windows user's index holds, not among all stores everywhere -
    /// <see cref="OutlookAI.Core.IndexSearch.StoreIndexMatcher"/> refuses any match it finds
    /// twice rather than picking one.
    /// </para>
    /// </summary>
    public static class StoreHash
    {
        /// <summary>
        /// <c>'.PUB'</c> as the DWORD Microsoft's code hashes in for a public-folder store, so its
        /// hash differs from the private store's built from the same signature.
        /// </summary>
        public const uint PublicStoreConstant = 0x2E505542;

        /// <summary>
        /// Microsoft's <c>ComputeHash</c>. <paramref name="fileName"/> is hashed as UTF-16 code
        /// units, the documented <c>pwzFileName</c> branch - Microsoft's code hashes a path only
        /// into a cached Exchange store's entry-ID variant, never into a PST's.
        /// </summary>
        public static uint Compute(byte[] blob, bool publicStore = false, string? fileName = null)
        {
            if (blob == null)
            {
                throw new ArgumentNullException(nameof(blob));
            }

            uint hash = 0;
            int words = blob.Length / 4;
            for (int i = 0; i < words; i++)
            {
                uint dw = (uint)(blob[i * 4] | (blob[(i * 4) + 1] << 8) | (blob[(i * 4) + 2] << 16) | (blob[(i * 4) + 3] << 24));
                hash = unchecked((hash << 5) + hash + dw);
            }

            for (int i = words * 4; i < blob.Length; i++)
            {
                hash = unchecked((hash << 5) + hash + blob[i]);
            }

            if (publicStore)
            {
                hash = unchecked((hash << 5) + hash + PublicStoreConstant);
            }

            if (fileName != null)
            {
                foreach (char c in fileName)
                {
                    hash = unchecked((hash << 5) + hash + c);
                }
            }

            return hash;
        }

        /// <summary>
        /// Parses the <c>($hash)</c> digits of a store segment. Compared as a NUMBER, never as
        /// text: Microsoft documents the digits "without any leading zeros" (mapi15), mapi16 has
        /// only been seen with eight, and the two spellings of one value must compare equal.
        /// </summary>
        public static bool TryParseUrlHash(string? digits, out uint value)
        {
            value = 0;
            if (string.IsNullOrEmpty(digits) || digits!.Length > 8)
            {
                return false;
            }

            return uint.TryParse(digits, NumberStyles.AllowHexSpecifier, CultureInfo.InvariantCulture, out value);
        }

        /// <summary>
        /// Every hash Microsoft's <c>ComputeStoreHash</c> can produce for one store, each named by
        /// its input, so the index's own <c>($hash)</c> decides which one Outlook used rather than
        /// this code guessing the store's type:
        /// <list type="bullet">
        /// <item><c>entryId</c> - <c>Compute(Store.StoreID)</c>; the documented rule for every
        /// non-Exchange store, and the only input measured.</item>
        /// <item><c>mappingSignature</c> / <c>profileMappingSignature</c> - the documented rule for
        /// a cached Exchange store, from the store object and from its profile section; with
        /// <c>'.PUB'</c> for a public-folder store.</item>
        /// <item><c>entryIdAndPath</c> - the documented entry-ID variant for a cached Exchange
        /// store, which also hashes in the offline-store path.</item>
        /// </list>
        /// Inputs that are absent are skipped. A 32-bit candidate that happens to equal another
        /// store's hash is what <see cref="OutlookAI.Core.IndexSearch.StoreIndexMatcher"/>'s
        /// two-claims rule exists for.
        /// </summary>
        public static IReadOnlyList<StoreHashCandidate> Candidates(
            string? storeIdHex,
            bool exchangeStore,
            bool publicStore,
            string? mappingSignatureHex,
            string? profileMappingSignatureHex,
            string? offlineStorePath)
        {
            var list = new List<StoreHashCandidate>(4);
            bool haveEntryId = Hex.TryDecode(storeIdHex, out byte[] entryId);
            if (haveEntryId)
            {
                list.Add(new StoreHashCandidate(StoreHashInput.EntryId, Compute(entryId)));
            }

            if (!exchangeStore)
            {
                return list;
            }

            if (Hex.TryDecode(profileMappingSignatureHex, out byte[] profileSignature))
            {
                list.Add(new StoreHashCandidate(StoreHashInput.ProfileMappingSignature, Compute(profileSignature, publicStore)));
            }

            if (Hex.TryDecode(mappingSignatureHex, out byte[] signature))
            {
                list.Add(new StoreHashCandidate(StoreHashInput.MappingSignature, Compute(signature, publicStore)));
            }

            if (haveEntryId && !string.IsNullOrEmpty(offlineStorePath))
            {
                list.Add(new StoreHashCandidate(StoreHashInput.EntryIdAndPath, Compute(entryId, publicStore, offlineStorePath)));
            }

            return list;
        }
    }

    /// <summary>Which documented input a <see cref="StoreHashCandidate"/> was computed from.</summary>
    public enum StoreHashInput
    {
        /// <summary><c>Store.StoreID</c> (<c>PR_ENTRYID</c>) - every non-Exchange store; measured.</summary>
        EntryId = 0,

        /// <summary>The cached Exchange store's <c>PR_MAPPING_SIGNATURE</c> as its profile section holds it - documented.</summary>
        ProfileMappingSignature = 1,

        /// <summary>The cached Exchange store's own <c>PR_MAPPING_SIGNATURE</c> - documented input, other route.</summary>
        MappingSignature = 2,

        /// <summary><c>PR_ENTRYID</c> plus the offline-store path - the documented entry-ID variant for cached Exchange.</summary>
        EntryIdAndPath = 3,
    }

    /// <summary>One hash a store might carry in its index URL, and the input it came from.</summary>
    public sealed class StoreHashCandidate
    {
        /// <summary>Creates a candidate.</summary>
        public StoreHashCandidate(StoreHashInput input, uint hash)
        {
            Input = input;
            Hash = hash;
        }

        /// <summary>The documented input the hash was computed from.</summary>
        public StoreHashInput Input { get; }

        /// <summary>The hash, as Microsoft's <c>ComputeHash</c> returns it.</summary>
        public uint Hash { get; }
    }
}
