using System;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// The two write-tool guards that only ASK whether a folder is a special one, decided without
    /// ever creating it (Q84 decision A, direction 1, 2026-09-24): move_mail's "the target is not
    /// the Outbox, and is neither Deleted Items nor under it", and update_draft's and
    /// discard_draft's "the item lives in the Drafts folder or under it".
    /// <para>
    /// <b>Why these two.</b> Both used to ask <c>Store.GetDefaultFolder</c>, which on a POP3, IMAP or
    /// data-file store that lacks the folder can MAKE it: measured for 23 = Junk Email and
    /// 39 = Archive on a POP3 PST, and for 16 = Drafts on a data file attached to a test guest
    /// (<c>Docs/live-tier-on-the-vm.md</c> section 4.1). The same data file was NOT given an
    /// Outbox by <c>GetDefaultFolder(4)</c>, and Deleted Items could not be tested - every data
    /// file is created with one - so for the move guard this is the rule rather than a measured
    /// creation. Either way a check that only compares folder identities never needs the folder
    /// to exist: a store with no Drafts folder holds no draft, and nothing can be moved into a
    /// Deleted Items or an Outbox a store does not have. So both ask
    /// <see cref="SpecialFolders.Resolve"/> instead - an Exchange store exactly as before
    /// (<c>GetDefaultFolder</c>, same ids, same order), any other store only once the folder is
    /// proven to exist.
    /// </para>
    /// <para>
    /// <b>Fail closed.</b> A lookup that cannot be completed - the resolver's
    /// <see cref="OutlookComSession.DefaultFolderResolution.Unreadable"/>, or a folder that
    /// resolves but will not say its EntryID - REFUSES, with a code that says the check could not
    /// be made rather than claiming an answer: <see cref="TargetGuardUnreadable"/> for a move,
    /// <see cref="DraftsFolderUnreadable"/> for a draft. For the move guard that is a change on
    /// every store type, Exchange included: it used to skip a check it could not make and let the
    /// move go ahead.
    /// </para>
    /// <para>
    /// The three places that PUT something into a special folder - a new draft's Drafts, a reply's
    /// source-store Drafts, and the Deleted Items a discarded draft is reported in - still ask
    /// <c>GetDefaultFolder</c> and are deliberately not here: whether they may create is the
    /// maintainer's open question.
    /// </para>
    /// <para>
    /// Pure over <see cref="ISpecialFolderStore"/> and a folder-chain predicate, so T1 pins every
    /// outcome against a fake store; <see cref="OutlookComSession"/> supplies the COM store and the
    /// walk up the folder tree.
    /// </para>
    /// </summary>
    public static class SpecialFolderGuards
    {
        /// <summary>move_mail: the target is the store's Outbox.</summary>
        public const string TargetIsOutbox = "TargetIsOutbox";

        /// <summary>move_mail: the target is the store's Deleted Items or a folder under it.</summary>
        public const string TargetIsDeletedItems = "TargetIsDeletedItems";

        /// <summary>move_mail: where the store's Deleted Items or Outbox is could not be established, so the move is refused.</summary>
        public const string TargetGuardUnreadable = "TargetGuardUnreadable";

        /// <summary>update_draft / discard_draft: the item is not in its store's Drafts folder (or the store has none).</summary>
        public const string NotInDraftsFolder = "NotInDraftsFolder";

        /// <summary>update_draft / discard_draft: which folder is the store's Drafts could not be established, so the item is refused.</summary>
        public const string DraftsFolderUnreadable = "DraftsFolderUnreadable";

        /// <summary>
        /// The special-folder half of move_mail's target guard: null when the target may be moved
        /// into, otherwise the refusal code. Deleted Items is looked up before the Outbox, as the
        /// guard always did. A refusal that is proven - the Outbox, or Deleted Items and below -
        /// is named even when the other lookup failed; otherwise any failed lookup refuses.
        /// </summary>
        /// <param name="store">The store the item and the target live in (moves are same-store).</param>
        /// <param name="targetEntryId">The target folder's EntryID.</param>
        /// <param name="targetIsOrIsUnder">
        /// True when the target folder, or any folder above it, has the given EntryID. Consulted
        /// only for a Deleted Items that resolved.
        /// </param>
        public static string? MoveTargetRefusal(ISpecialFolderStore store, string? targetEntryId, Func<string, bool> targetIsOrIsUnder)
        {
            if (store == null)
            {
                throw new ArgumentNullException(nameof(store));
            }

            if (targetIsOrIsUnder == null)
            {
                throw new ArgumentNullException(nameof(targetIsOrIsUnder));
            }

            OutlookComSession.DefaultFolderResolution deletedItems =
                ResolveEntryId(store, SpecialFolders.OlFolderDeletedItems, out string? deletedItemsEntryId);
            OutlookComSession.DefaultFolderResolution outbox =
                ResolveEntryId(store, SpecialFolders.OlFolderOutbox, out string? outboxEntryId);

            if (string.IsNullOrEmpty(targetEntryId))
            {
                // A target that will not say who it is cannot be shown not to be either folder.
                return TargetGuardUnreadable;
            }

            if (outbox == OutlookComSession.DefaultFolderResolution.Resolved
                && string.Equals(targetEntryId, outboxEntryId, StringComparison.OrdinalIgnoreCase))
            {
                return TargetIsOutbox;
            }

            if (deletedItems == OutlookComSession.DefaultFolderResolution.Resolved && targetIsOrIsUnder(deletedItemsEntryId!))
            {
                return TargetIsDeletedItems;
            }

            if (deletedItems == OutlookComSession.DefaultFolderResolution.Unreadable
                || outbox == OutlookComSession.DefaultFolderResolution.Unreadable)
            {
                return TargetGuardUnreadable;
            }

            // Each folder is either proven elsewhere or proven absent - and a folder the store
            // does not have cannot be the target.
            return null;
        }

        /// <summary>
        /// The special-folder half of the update_draft / discard_draft gate: null when the item
        /// lives in its store's Drafts folder or under it, otherwise the refusal code. A store
        /// with no Drafts folder holds no draft (<see cref="NotInDraftsFolder"/>); one whose
        /// Drafts folder cannot be established is refused as such (<see cref="DraftsFolderUnreadable"/>),
        /// never assumed either way.
        /// </summary>
        /// <param name="store">The store the item lives in.</param>
        /// <param name="itemFolderIsOrIsUnder">
        /// True when the item's folder, or any folder above it, has the given EntryID. Consulted
        /// only for a Drafts folder that resolved.
        /// </param>
        public static string? DraftsFolderRefusal(ISpecialFolderStore store, Func<string, bool> itemFolderIsOrIsUnder)
        {
            if (store == null)
            {
                throw new ArgumentNullException(nameof(store));
            }

            if (itemFolderIsOrIsUnder == null)
            {
                throw new ArgumentNullException(nameof(itemFolderIsOrIsUnder));
            }

            switch (ResolveEntryId(store, SpecialFolders.OlFolderDrafts, out string? draftsEntryId))
            {
                case OutlookComSession.DefaultFolderResolution.Resolved:
                    return itemFolderIsOrIsUnder(draftsEntryId!) ? null : NotInDraftsFolder;
                case OutlookComSession.DefaultFolderResolution.Absent:
                    return NotInDraftsFolder;
                default:
                    return DraftsFolderUnreadable;
            }
        }

        /// <summary>
        /// One special folder's EntryID, looked up without creating the folder; the folder itself
        /// is released here. A folder that resolves but will not say its EntryID is
        /// <see cref="OutlookComSession.DefaultFolderResolution.Unreadable"/>: an identity the
        /// guard cannot compare is a check it cannot make.
        /// </summary>
        private static OutlookComSession.DefaultFolderResolution ResolveEntryId(
            ISpecialFolderStore store,
            int olDefaultFolderId,
            out string? entryId)
        {
            entryId = null;
            OutlookComSession.DefaultFolderResolution resolution =
                SpecialFolders.Resolve(store, olDefaultFolderId, out object? folder, out _);
            if (resolution != OutlookComSession.DefaultFolderResolution.Resolved)
            {
                return resolution;
            }

            try
            {
                entryId = store.EntryIdOf(folder!);
            }
            finally
            {
                store.Release(folder);
            }

            return string.IsNullOrEmpty(entryId)
                ? OutlookComSession.DefaultFolderResolution.Unreadable
                : OutlookComSession.DefaultFolderResolution.Resolved;
        }
    }
}
