using System;
using System.Globalization;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// What a store is called when Outlook will not say its name (gap G2).
    /// <para>
    /// EVERY scope, bucket, label and refusal message in this server is keyed by a store's
    /// <c>DisplayName</c>, so a store whose <c>DisplayName</c> read throws had nowhere to go
    /// and was dropped on the floor: absent from <c>list_folders</c>, absent from
    /// <c>list_accounts</c>, and in the freshness sweep its four default folders counted as
    /// skipped in the total while landing in no per-store bucket at all. Dropping it is the
    /// silent failure this whole audit is about - the folders behind that store hold mail,
    /// and nothing in any payload said they had not been looked at.
    /// </para>
    /// <para>
    /// So it gets a LABEL instead of being dropped. The label is deliberately not a guess at
    /// the real name: it is derived from the store's 1-based position in
    /// <c>Namespace.Stores</c>, which is the only thing about such a store that can still be
    /// read, and it is bracketed so it cannot be mistaken for a display name. A guess would
    /// be worse than the drop it replaces - two stores could collide into one sweep bucket,
    /// and a scope keyed by a guessed name would silently answer with another store's mail.
    /// </para>
    /// <para>
    /// The label CANNOT round-trip as a scope, and that is inherent rather than an omission:
    /// a store scope is resolved by comparing against <c>DisplayName</c>, which is the one
    /// property that failed. Anything derived from a different property - the root folder's
    /// name, the PST file path - would have the same problem while LOOKING usable, which is
    /// why none is used. <see cref="IsUnnamedStoreLabel"/> exists so a caller who reads the
    /// label out of a payload and passes it back as <c>store</c> gets told exactly that,
    /// rather than the generic "no such store" that would send it hunting for a typo.
    /// </para>
    /// </summary>
    public static class StoreNaming
    {
        /// <summary>
        /// Opening text of an unnamed store's label. Public because the refusal message,
        /// the payload flag and T1 all have to agree on one spelling.
        /// </summary>
        public const string UnnamedStorePrefix = "(unnamed store ";

        /// <summary>
        /// The label for the store at 1-based <paramref name="profilePosition"/> in
        /// <c>Namespace.Stores</c> whose display name could not be read.
        /// <para>
        /// Stable for as long as the profile's store order is - which is what makes two
        /// calls of <c>list_folders</c> agree, and is also the honest limit of it: adding or
        /// removing a store can renumber the label, so it identifies a store within one
        /// answer and never across a profile change.
        /// </para>
        /// </summary>
        public static string LabelForUnnamedStore(int profilePosition)
        {
            if (profilePosition < 1)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(profilePosition), "Store positions are 1-based, as Namespace.Stores is.");
            }

            return UnnamedStorePrefix + profilePosition.ToString(CultureInfo.InvariantCulture) + ")";
        }

        /// <summary>
        /// Whether <paramref name="storeName"/> is one of these labels rather than a real
        /// display name. Used to answer a caller that passed a label back as a scope, and to
        /// keep labels out of the "Known stores:" enumeration a refusal offers as remedies.
        /// </summary>
        public static bool IsUnnamedStoreLabel(string? storeName)
        {
            return storeName != null
                && storeName.StartsWith(UnnamedStorePrefix, StringComparison.Ordinal)
                && storeName.EndsWith(")", StringComparison.Ordinal);
        }
    }
}
