using System;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Turns a <see cref="HitLocationResult.Error"/> token into a sentence a calling agent
    /// can act on.
    /// <para>
    /// Why: the index legitimately outlives the mailbox. A read-only probe (2026-07-27,
    /// v3.MD section 0.8 block (q)) found ~458 rows in one delegate store filed under a leaf
    /// path with NO corresponding Outlook folder - searching returns them, opening them
    /// cannot work, and the shipped message said "Re-run search - the item may have moved",
    /// which is the one remedy that provably does not help: the next search returns the
    /// same orphan row. The tokens are content-free by construction (S4/S6) and stay in the
    /// message for diagnostics; only the remedy changes.
    /// </para>
    /// </summary>
    public static class LocateFailureAdvice
    {
        /// <summary>Composes the message for a failed locate; never returns null or empty.</summary>
        public static string Describe(string? locateError)
        {
            string token = locateError ?? "unknown";
            return "The item could not be opened in Outlook (" + token + "). " + Remedy(token);
        }

        /// <summary>The actionable half of <see cref="Describe"/>, chosen from the failure token.</summary>
        public static string Remedy(string? locateError)
        {
            string token = locateError ?? string.Empty;

            if (Mentions(token, "StoreNotFound"))
            {
                return "The mailbox that held it is not open in this Outlook profile, so the search index is serving "
                    + "rows for a store Outlook cannot reach. Open or reconnect that mailbox, or restrict the search "
                    + "with store= to a mailbox list_accounts reports.";
            }

            if (Mentions(token, "FolderNotFound"))
            {
                return "Its folder no longer exists in Outlook - this hit is a stale index row (the search index keeps "
                    + "entries for deleted or renamed folders), so re-running the same search will return it again and "
                    + "it can never be opened. Ignore this hit, or search the mailbox again without the folder bound; "
                    + "use exhaustive:true (store + folder/after) for an index-free COM search that only sees folders "
                    + "that really exist.";
            }

            if (Mentions(token, "NoSubjectTimeMatch") || Mentions(token, "NoTimeOnlyMatch"))
            {
                return "Its folder opened but the item is no longer in it - it was moved or deleted after the index "
                    + "recorded it. Re-run the search for fresh ids (EntryIDs change on moves).";
            }

            if (Mentions(token, "FolderTooLargeForTimeOnlyProbe"))
            {
                return "The item carries no subject, and its folder is too large to identify it by timestamp alone. "
                    + "Search again with a narrower folder or date bound so the hit can be pinned down.";
            }

            if (Mentions(token, "RootFolderUnavailable"))
            {
                return "Outlook could not open that mailbox's folder tree. Check outlook_health and retry.";
            }

            if (Mentions(token, "UrlNotParsable") || Mentions(token, "NoItemPathDisplay"))
            {
                return "The index row carries no usable location for it. Re-run the search; if the same hit comes back, "
                    + "that row is unusable - use exhaustive:true (store + folder/after) for an index-free COM search.";
            }

            return "Re-run the search for fresh ids; if the same hit keeps failing, use exhaustive:true "
                + "(store + folder/after) for an index-free COM search.";
        }

        /// <summary>True when the hit's folder is provably gone (the orphan-index-row case).</summary>
        public static bool IsMissingLocation(string? locateError)
        {
            return locateError != null
                && (Mentions(locateError, "FolderNotFound") || Mentions(locateError, "StoreNotFound"));
        }

        private static bool Mentions(string token, string needle)
        {
            return token.IndexOf(needle, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}
