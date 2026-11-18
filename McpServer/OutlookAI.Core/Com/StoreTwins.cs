using System;
using System.Collections.Generic;

using OutlookAI.Core.Mapi;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// One mailbox in two store entries of the profile. Exchange automaps a mailbox the profile's account has
    /// Full Access to as an online entry, and when that mailbox is an account of the profile too, its own
    /// cached store is beside it: measured on the Exchange guest (2026-11-18), the two entries carry the same
    /// <c>StoreID</c>, byte for byte, and top folders with the same provider UID - so every ID made from the
    /// mailbox's identity names both, and Outlook itself cannot tell them apart by <c>StoreID</c>. The
    /// maintainer's own profile has the same kind of pair.
    /// <para>
    /// DECIDED 2026-11-18 under the maintainer's Q131/Q132 delegation ("most stable"): an ID never resolves
    /// to two stores or two folders. Such a mailbox is listed and resolved ONCE, through its cached (indexed)
    /// entry; its other entry is not listed again by list_folders or find_folder, and no ID resolves into it
    /// while the cached entry exists. How list_accounts and outlook_health describe the other entry is the
    /// maintainer's open decision 207, so they still show it as they did.
    /// </para>
    /// <para>
    /// Pure, and in one place, because both sides need it: the COM host's folder walk skips the other entry,
    /// and the server picks the kept entry for a store ID. Two entries are one mailbox when their top folders
    /// carry the same provider UID (OutlookIds.ProviderUidHexOf), or - when either top folder was not read -
    /// when their StoreIDs are the same bytes. The KEPT entry of a mailbox is the one Outlook keeps locally
    /// (a cached Exchange store, or a store that is not Exchange), then one that is not a delegate's
    /// (ExchangeStoreType 1), then the first in the profile's order: deterministic, never a guess.
    /// </para>
    /// </summary>
    public static class StoreTwins
    {
        /// <summary>
        /// For each store, in the order given: the position of the entry that presents its mailbox when it is
        /// ANOTHER entry of a mailbox that entry presents, or -1 when it is its mailbox's kept entry or alone.
        /// </summary>
        public static int[] KeptEntryOf(IReadOnlyList<ComStoreDetail> stores)
        {
            if (stores == null)
            {
                throw new ArgumentNullException(nameof(stores));
            }

            int[] group = new int[stores.Count];
            for (int i = 0; i < stores.Count; i++)
            {
                group[i] = i;
                for (int j = 0; j < i; j++)
                {
                    if (group[j] == j && SameMailbox(stores[i], stores[j]))
                    {
                        group[i] = j;
                        break;
                    }
                }
            }

            // Every member of a group points at the group's first; the kept entry is its best member.
            int[] kept = new int[stores.Count];
            for (int i = 0; i < stores.Count; i++)
            {
                kept[i] = -1;
            }

            for (int first = 0; first < stores.Count; first++)
            {
                if (group[first] != first)
                {
                    continue;
                }

                int best = first;
                int members = 0;
                for (int k = first; k < stores.Count; k++)
                {
                    if (group[k] != first)
                    {
                        continue;
                    }

                    members++;
                    if (Rank(stores[k]) < Rank(stores[best]))
                    {
                        best = k;
                    }
                }

                if (members < 2)
                {
                    continue;
                }

                for (int k = first; k < stores.Count; k++)
                {
                    if (group[k] == first && k != best)
                    {
                        kept[k] = best;
                    }
                }
            }

            return kept;
        }

        /// <summary>Whether the store at <paramref name="index"/> is another entry of a mailbox an entry before or after it presents.</summary>
        public static bool IsOtherEntry(IReadOnlyList<ComStoreDetail> stores, int index)
        {
            return KeptEntryOf(stores)[index] >= 0;
        }

        /// <summary>
        /// Whether two store entries are one mailbox: their top folders carry the same provider UID, or - when
        /// either top folder is unknown - their StoreIDs are the same bytes. Two stores without either are
        /// never one.
        /// </summary>
        public static bool SameMailbox(ComStoreDetail a, ComStoreDetail b)
        {
            if (a == null)
            {
                throw new ArgumentNullException(nameof(a));
            }

            if (b == null)
            {
                throw new ArgumentNullException(nameof(b));
            }

            string? uidA = OutlookIds.ProviderUidHexOf(a.RootFolderEntryId);
            string? uidB = OutlookIds.ProviderUidHexOf(b.RootFolderEntryId);
            if (uidA != null && uidB != null)
            {
                return string.Equals(uidA, uidB, StringComparison.OrdinalIgnoreCase);
            }

            return !string.IsNullOrEmpty(a.StoreId)
                && !string.IsNullOrEmpty(b.StoreId)
                && StoreHash.TryDecodeHex(a.StoreId, out byte[] idA)
                && StoreHash.TryDecodeHex(b.StoreId, out byte[] idB)
                && BytesEqual(idA, idB);
        }

        /// <summary>Lower ranks first: kept locally, then not a delegate's. Ties go to the earlier entry.</summary>
        private static int Rank(ComStoreDetail store)
        {
            bool local = store.IsCachedExchange == true || store.ExchangeStoreType == 3;
            bool delegateEntry = store.ExchangeStoreType == 1;
            return (local ? 0 : 2) + (delegateEntry ? 1 : 0);
        }

        private static bool BytesEqual(byte[] a, byte[] b)
        {
            if (a.Length != b.Length)
            {
                return false;
            }

            for (int i = 0; i < a.Length; i++)
            {
                if (a[i] != b[i])
                {
                    return false;
                }
            }

            return true;
        }
    }
}
