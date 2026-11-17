using System;
using System.Collections.Generic;
using OutlookAI.Core.IndexSearch;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Everything the service layer needs from a live Outlook session.
    /// <para>
    /// This interface is simultaneously two things, and it is worth knowing both. It is
    /// the TEST SEAM - <see cref="OutlookAI.Core.Services.MailService"/> depends on this
    /// rather than on a sealed COM class, so its logic is exercisable without a real
    /// Outlook. And it is the IPC CONTRACT - each method is one call across the pipe to
    /// the killable COM host (see McpServer/Docs/com-host.md). Those turned out to be the
    /// same artifact, which is why they are one interface rather than two.
    /// </para>
    /// <para>
    /// Deliberately free of any JSON attribute or serializer dependency: Core also
    /// targets net48, where System.Text.Json is not available. The wire encoding lives
    /// entirely in the OutlookAI.ComHost assembly.
    /// </para>
    /// <para>
    /// No member of this contract exposes a live COM object. Every implementation
    /// returns plain data snapshots, which is what makes the process split possible at
    /// all - an RCW cannot cross a process boundary, and cannot even be released from
    /// the wrong side of one.
    /// </para>
    /// </summary>
    public interface IOutlookSession
    {
        /// <summary>Current MAPI profile name. The cheap liveness ping - if this answers, the session is alive.</summary>
        string GetProfileName();

        /// <summary>Accounts configured in the profile.</summary>
        IReadOnlyList<ComAccountInfo> GetAccounts();

        /// <summary>Every store in the profile, with item counts where available.</summary>
        IReadOnlyList<ComStoreDetail> GetStoreDetails();

        /// <summary>
        /// Folder tree for one store - named by its own entry id, <c>Store.StoreID</c> - or all
        /// stores when null, bounded by <paramref name="absoluteWalkCap"/> - which the result REPORTS
        /// rather than merely obeying, because a list alone cannot say whether the tree or the cap
        /// ended it. Every folder carries its entry id, every store its root folder's (Q131, Q133).
        /// </summary>
        ComFolderTree ListFolders(string? storeEntryId, int absoluteWalkCap);

        /// <summary>
        /// Opens one folder by its entry id and describes it: its name, its names up to the store's
        /// root, its store, and whether mail may be moved into it (Q131-Q136). Read-only - it opens
        /// the folder with <c>GetFolderFromID</c>, which creates nothing. Null with
        /// <paramref name="error"/> <c>FolderNotFound</c> when no store of the profile has it.
        /// </summary>
        ComFolderDetail? TryDescribeFolder(string folderEntryId, out string? error);

        /// <summary>
        /// Creates one mail folder named <paramref name="name"/> - verbatim, a <c>/</c> included - directly
        /// under the folder <paramref name="parentFolderEntryId"/> (create_folder, Q134 (i)). Refused,
        /// with nothing created, when the parent cannot be opened, is not a mail folder, is Deleted
        /// Items or under it, or is the Outbox, and when the parent already has a folder of that name
        /// (<c>FolderExists</c>, the existing folder's entry id after the colon).
        /// </summary>
        ComFolderDetail? TryCreateFolder(string parentFolderEntryId, string name, out string? error);

        /// <summary>
        /// Store-relative folder paths for one store, bounded by
        /// <paramref name="absoluteWalkCap"/> and reporting that bound for the same reason.
        /// </summary>
        ComFolderPathList ListFolderPaths(string storeDisplayName, int absoluteWalkCap);

        /// <summary>Folder paths whose leaf segment matches <paramref name="leafName"/> (delegate-store resolution).</summary>
        IReadOnlyList<IReadOnlyList<string>> FindFolderPathsByLeafName(string storeDisplayName, string leafName, int absoluteWalkCap);

        /// <summary>
        /// The freshness sweep: mail newer than <paramref name="sinceUtc"/> that the index
        /// has not caught up with. <paramref name="perStoreSinceUtc"/> overrides that start
        /// per store (each store's own index frontier); <paramref name="sinceUtc"/> is what
        /// a store not named in it gets.
        /// </summary>
        /// <param name="timeBudgetMs">
        /// Wall clock the walk may spend before it stops at the next store or folder
        /// boundary and returns what it covered, with
        /// <see cref="ComSweepResult.SweepBudgetExpired"/> set. Zero or less means unbounded,
        /// which is what an in-process caller with its own bound asks for.
        /// <para>
        /// This is the SWEEP's own soft budget, not the gateway deadline above it. Without
        /// it, a sweep that ran long produced a timeout, the supervisor concluded the host
        /// was wedged, the child was killed, and every folder already swept was discarded -
        /// so a large mailbox lost its freshness tier entirely instead of getting a partial
        /// one. Same discipline as <see cref="ExhaustiveScan"/>'s.
        /// </para>
        /// </param>
        ComSweepResult SweepFoldersNewerThan(
            DateTime sinceUtc,
            int perFolderCap,
            bool includeBodies,
            string? onlyStoreDisplayName,
            IReadOnlyList<string>? folderPath = null,
            bool includeSubfolders = true,
            IReadOnlyDictionary<string, DateTime>? perStoreSinceUtc = null,
            int timeBudgetMs = 0,
            string? folderEntryId = null);

        /// <summary>Index-bypassing COM scan, bounded by item count and time budget.</summary>
        /// <param name="resumeFrom">
        /// Where a previous page of this same scan stopped, or null to start at the top of
        /// the scope. It carries the finished-folder set and the within-folder cursor, and
        /// the child treats it as authoritative rather than re-deriving it.
        /// <para>
        /// The state travels DOWN on every page and is held in the server PARENT between
        /// them, deliberately. The failure this mode is bounded against - a scan that runs
        /// past its deadline - ends with the supervisor killing the COM host child, so state
        /// kept in the child would be destroyed by exactly the event that makes resumption
        /// necessary. Re-running one page still changes nothing, so the operation stays
        /// classified read-only and retryable.
        /// </para>
        /// </param>
        ComExhaustiveResult ExhaustiveScan(
            string storeDisplayName,
            IReadOnlyList<string>? folderPath,
            IReadOnlyList<string>? terms,
            DateTime? sinceUtc,
            DateTime? beforeUtc,
            int maxItems,
            int timeBudgetMs,
            SearchIn searchIn = SearchInValues.Default,
            bool includeSubfolders = false,
            ComScanCursor? resumeFrom = null,
            string? folderEntryId = null);

        /// <summary>Opens one item and snapshots it. Null with <paramref name="error"/> set when it cannot be opened.</summary>
        ComItemDetail? TryReadItem(
            string entryIdHex,
            string? storeId,
            bool includeHeaders,
            bool includeBody,
            out string? error,
            bool includeHtml = false);

        /// <summary>Members of an item's conversation.</summary>
        IReadOnlyList<ComMailBrief>? TryGetConversationItems(string entryIdHex, string? storeId, int maxItems, out string? error);

        /// <summary>Attachment metadata for one item.</summary>
        IReadOnlyList<ComAttachmentInfo> SnapshotAttachmentsById(string entryId, string? storeId);

        /// <summary>Saves one attachment to disk; reports the written path and its size.</summary>
        /// <param name="attemptedPath">
        /// Where the write was aimed, reported on the failure path too: a save that throws
        /// part-way leaves a partial file at a name the caller cannot predict.
        /// </param>
        string? TrySaveAttachment(
            string entryIdHex,
            string? storeId,
            int attachmentIndex,
            string targetDirectory,
            out long sizeBytes,
            out string? attemptedPath,
            out string? error);

        /// <summary>Resolves an index hit to a real item by folder path and subject/time match.</summary>
        ComOpenResult? TryResolveByPath(
            string storeDisplayName,
            IReadOnlyList<string> folderPath,
            string itemSubject,
            DateTime? indexReceivedUtc,
            out string? error,
            int toleranceSeconds = 120);

        /// <summary>Opens an item in Outlook's own UI.</summary>
        ComOpenResult? TryDisplayItem(string entryIdHex, string? storeId, out string? error);

        /// <summary>
        /// Navigates the active Explorer to the folder <paramref name="folderEntryId"/>, or - when that is
        /// null - to the Inbox of the store <paramref name="storeEntryId"/>.
        /// </summary>
        ComExplorerState? TryGotoFolder(string? folderEntryId, string? storeEntryId, out string? error);

        /// <summary>
        /// Drives Outlook's own search UI, after navigating to the folder <paramref name="folderEntryId"/>
        /// or the Inbox of the store <paramref name="storeEntryId"/> when either is given.
        /// </summary>
        ComExplorerState? TryShowSearchResults(
            string query,
            int olSearchScope,
            string? folderEntryId,
            string? storeEntryId,
            out string? error);

        /// <summary>Basic mail properties, used to resolve a hit before acting on it.</summary>
        ComDraftInfo? TryGetMailInfo(string entryIdHex, string? storeId, out string? error);

        /// <summary>
        /// Resolves a store's designated Archive folder READ-ONLY: never creates anything. A
        /// store without one answers <c>NoDesignatedArchiveFolder</c>; a non-Exchange store
        /// whose designation could not be read answers <c>ArchiveDesignationUnreadable</c>.
        /// </summary>
        ComArchiveFolderInfo? TryResolveArchiveFolder(string storeDisplayName, out string? error);

        /// <summary>
        /// Resolves a store's designated Archive folder for archive_mail - the ONE lookup that
        /// may CREATE it (on a store that has none, Outlook makes it), because the caller asked
        /// for mail to be moved into it.
        /// </summary>
        /// <param name="createdFolderPath">
        /// Store-relative path of the Archive folder this call CREATED, reported whether the
        /// resolution then succeeded or was refused by verification; null when it created
        /// nothing. This server cannot delete folders, so the caller has to be told.
        /// </param>
        ComArchiveFolderInfo? TryResolveOrCreateArchiveFolder(
            string storeDisplayName,
            out string? createdFolderPath,
            out string? error);

        /// <summary>
        /// move_mail's move: one item into the folder <paramref name="targetFolderEntryId"/>, which must be
        /// in the item's own store (<c>CrossStoreTarget:</c> otherwise) and pass the move guard - a mail
        /// folder, not Deleted Items or anything under it, not the Outbox. Creates nothing: since Q134 (i)
        /// a folder is made only by <see cref="TryCreateFolder"/>.
        /// </summary>
        ComMoveItemResult? TryMoveItemToFolder(
            string entryIdHex,
            string? storeId,
            string targetFolderEntryId,
            out string? error);

        /// <summary>Moves an item to a folder identified by EntryID.</summary>
        ComMoveItemResult? TryMoveItemToFolderId(
            string entryIdHex,
            string? storeId,
            string targetFolderEntryId,
            string targetStoreId,
            out string? error);

        /// <summary>Creates a new draft.</summary>
        /// <param name="savedDraftEntryId">
        /// The EntryID of the draft as soon as <c>Save()</c> committed it, reported EVEN WHEN
        /// the call then fails. Four things happen after that save inside the same try - the
        /// hidden inspector close, the relocate to Drafts, the optional Display and the
        /// snapshot reads - and a COM failure in any of them used to lose the id, leaving a
        /// real draft in the mailbox that <c>discard_draft</c> was structurally unable to
        /// reach, because its registry gate only knows ids the caller was told. It is
        /// re-read after the relocate, since a move mints a new one.
        /// </param>
        ComDraftCreateResult? TryCreateNewDraft(
            string accountSmtpAddress,
            IReadOnlyList<string> toRecipients,
            string subject,
            ComDraftBody body,
            bool display,
            ComSignatureOverride? signatureOverride,
            ComDraftOptions? options,
            out string? savedDraftEntryId,
            out string? error);

        /// <summary>Creates a reply, reply-all or forward draft via Outlook's own derivation.</summary>
        /// <param name="savedDraftEntryId">
        /// Same contract as <see cref="TryCreateNewDraft"/>'s: the id of the saved draft,
        /// reported even when a later step fails, so a failed derivation cannot leave an
        /// orphan out of the cleanup tool's reach.
        /// </param>
        ComDraftCreateResult? TryCreateDerivedDraft(
            string sourceEntryIdHex,
            string? sourceStoreId,
            ComDerivedDraftKind kind,
            IReadOnlyList<string> toRecipients,
            ComDraftBody body,
            bool display,
            ComSignatureOverride? signatureOverride,
            ComDraftOptions? options,
            out string? savedDraftEntryId,
            out string? error);

        /// <summary>
        /// Edits an existing draft in place. <paramref name="resume"/> is null for an
        /// ordinary call and carries the pre-image recorded before an earlier attempt when
        /// this call is a REPEAT of one the COM host was killed part-way through.
        /// </summary>
        ComDraftUpdateResult? TryUpdateDraft(
            string entryIdHex,
            string? storeId,
            ComDraftBody? body,
            string? subject,
            IReadOnlyList<string>? toRecipients,
            IReadOnlyList<string>? ccRecipients,
            IReadOnlyList<string>? bccRecipients,
            int? importance,
            bool? requestReadReceipt,
            ComSignatureOverride? signatureOverride,
            IReadOnlyList<string> attachmentsToAdd,
            IReadOnlyList<string> attachmentsToRemove,
            ComDraftUpdateResume? resume,
            bool display,
            out string? error);

        /// <summary>Deletes a draft this server created.</summary>
        /// <param name="error">
        /// Null on success. A failure before the delete was issued is
        /// <see cref="ComErrorTokens.DiscardNotStarted"/> - the draft was NOT deleted (Q96 question
        /// 4 (b)); a failure from the delete on is the bare COM description, and then whether the
        /// draft was deleted is unknown.
        /// </param>
        ComDraftDiscardResult? TryDiscardDraft(
            string entryIdHex,
            string? storeId,
            out string? error);

        /// <summary>Reads the state a draft must be in before it may be sent, including its content hash.</summary>
        ComSendableDraftState? TryGetSendableDraftState(string entryIdHex, string? storeId, out string? error);

        /// <summary>
        /// Sends a draft, but only when its content still hashes to
        /// <paramref name="expectedContentHash"/>.
        /// </summary>
        ComSendResult? TrySendDraft(
            string entryIdHex,
            string? storeId,
            string expectedContentHash,
            string? sentOnBehalfOfName,
            out string? error);
    }
}
