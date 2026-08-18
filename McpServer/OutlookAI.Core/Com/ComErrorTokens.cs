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
    }
}
