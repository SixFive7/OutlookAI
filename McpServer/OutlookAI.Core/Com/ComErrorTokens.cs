namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Failure words that the COM layer writes into an <c>out string? error</c> and the
    /// service layer BRANCHES ON.
    /// <para>
    /// Most of what those parameters carry is prose for a human: it picks the advice
    /// sentence and nothing more, so a reworded value costs nothing. This one is
    /// different. <see cref="ItemNotFound"/> is the only failure word that decides
    /// control flow - it is what tells the cross-store retry that the item was never
    /// opened, and therefore that looking in another store can still find it.
    /// </para>
    /// <para>
    /// It lives here because it was a matched pair of string literals in two files with
    /// no compiler between them, and that pair had already come apart twice:
    /// <c>TryUpdateDraft</c> and <c>TryDiscardDraft</c> asked for the token their own COM
    /// layer never set, so their retries were dead code and a draft in a non-default store
    /// answered with an opaque COM code (fixed in eee02f2). A shared constant makes the
    /// same mistake a compile error rather than a silent behaviour change, in both
    /// directions: nothing can set a misspelt token, and nothing can wait for one that is
    /// no longer written.
    /// </para>
    /// <para>
    /// The draft tokens below it (2026-10-03, Q96) decide no control flow; they decide what a
    /// failed draft call may CLAIM - that no draft exists, and what became of the Drafts folder
    /// - which is the other thing a misspelt literal would silently get wrong.
    /// </para>
    /// </summary>
    public static class ComErrorTokens
    {
        /// <summary>
        /// The item could not be OPENED - set at <c>Namespace.GetItemFromID</c> and
        /// nowhere else.
        /// <para>
        /// "Nowhere else" is the contract, not a detail. A retry is only safe while
        /// nothing has happened yet: past the open, the operation may have created a
        /// draft, moved an item, written a file or opened a window, and a second attempt
        /// in another store would repeat that rather than find anything.
        /// </para>
        /// </summary>
        public const string ItemNotFound = "ItemNotFound";

        /// <summary>
        /// A draft call failed BEFORE anything that can save a draft had run - so no draft
        /// exists (Q96 (iii), 2026-10-03). Written as <c>DraftNotStarted:&lt;COM failure&gt;</c>.
        /// <para>
        /// The line it marks is the composition: <c>GetInspector</c> and its
        /// <c>Close(olSave)</c> are the first steps that can put the item in the mailbox. Above
        /// it only an unsaved item exists - <c>Items.Add</c>, <c>Reply()</c> and
        /// <c>Forward()</c> make one in memory, and pinning its account saves nothing. A failure
        /// there used to share the catch-all's "a draft may have been saved", which cannot be
        /// true of it, and its advice to go looking in Drafts first.
        /// </para>
        /// </summary>
        public const string DraftNotStarted = "DraftNotStarted";

        /// <summary>
        /// new_draft's Drafts lookup itself failed - the call that makes the folder on a mailbox
        /// without one threw, or answered no folder - so no draft exists, and the lookup's own
        /// re-check established whether it made a folder first: any it did travels as the
        /// call's created folders, and none means none (Q96 (ii) and (iii)). Written as
        /// <c>DraftsFolderUnavailable:&lt;detail&gt;</c>.
        /// </summary>
        public const string DraftsFolderUnavailable = "DraftsFolderUnavailable";

        /// <summary>
        /// As <see cref="DraftsFolderUnavailable"/>, except that whether the failed lookup made a
        /// folder first could NOT be established - the mailbox's top-level folders would not
        /// list - so nothing is claimed either way. Written as
        /// <c>DraftsFolderCreationUnverified:&lt;detail&gt;</c>.
        /// </summary>
        public const string DraftsFolderCreationUnverified = "DraftsFolderCreationUnverified";

        /// <summary>
        /// A token with its detail, in the <c>Token:detail</c> shape the service layer reads back
        /// with <see cref="TryRead"/>. The detail is content-free (a COM failure's type and
        /// HRESULT), never a subject, a body or a name.
        /// </summary>
        public static string With(string token, string? detail)
        {
            return token + ":" + (string.IsNullOrEmpty(detail) ? "unknown" : detail);
        }

        /// <summary>
        /// True when <paramref name="error"/> is <paramref name="token"/>, bare or with a detail
        /// (<see cref="With"/>), and then the detail - "unknown" for a bare token. Ordinal and
        /// exact: one token never matches another that merely starts with it.
        /// </summary>
        public static bool TryRead(string? error, string token, out string detail)
        {
            detail = string.Empty;
            if (error == null || token == null)
            {
                return false;
            }

            if (string.Equals(error, token, System.StringComparison.Ordinal))
            {
                detail = "unknown";
                return true;
            }

            if (error.Length > token.Length + 1
                && error.StartsWith(token, System.StringComparison.Ordinal)
                && error[token.Length] == ':')
            {
                detail = error.Substring(token.Length + 1);
                return true;
            }

            return false;
        }
    }
}
