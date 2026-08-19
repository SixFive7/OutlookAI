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
        /// Folder tree for one store (or all stores when null), bounded by
        /// <paramref name="absoluteWalkCap"/> - which the result REPORTS rather than merely
        /// obeying, because a list alone cannot say whether the tree or the cap ended it.
        /// </summary>
        ComFolderTree ListFolders(string? storeDisplayName, int absoluteWalkCap);

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
            int timeBudgetMs = 0);

        /// <summary>Index-bypassing COM scan, bounded by item count and time budget.</summary>
        ComExhaustiveResult ExhaustiveScan(
            string storeDisplayName,
            IReadOnlyList<string>? folderPath,
            IReadOnlyList<string>? terms,
            DateTime? sinceUtc,
            DateTime? beforeUtc,
            int maxItems,
            int timeBudgetMs,
            SearchIn searchIn = SearchInValues.Default,
            bool includeSubfolders = false);

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
        string? TrySaveAttachment(
            string entryIdHex,
            string? storeId,
            int attachmentIndex,
            string targetDirectory,
            out long sizeBytes,
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

        /// <summary>Navigates the active Explorer to a folder.</summary>
        ComExplorerState? TryGotoFolder(string storeDisplayName, IReadOnlyList<string>? folderPath, out string? error);

        /// <summary>Drives Outlook's own search UI.</summary>
        ComExplorerState? TryShowSearchResults(
            string query,
            int olSearchScope,
            string? storeDisplayName,
            IReadOnlyList<string>? folderPath,
            out string? error);

        /// <summary>Basic mail properties, used to resolve a hit before acting on it.</summary>
        ComDraftInfo? TryGetMailInfo(string entryIdHex, string? storeId, out string? error);

        /// <summary>Resolves the archive target for a store.</summary>
        ComArchiveFolderInfo? TryResolveArchiveFolder(string storeDisplayName, out string? error);

        /// <summary>Moves an item to a folder identified by store-relative path.</summary>
        ComMoveItemResult? TryMoveItemToPath(
            string entryIdHex,
            string? storeId,
            IReadOnlyList<string> targetSegments,
            bool createMissing,
            string? requireStoreDisplayName,
            out string? error);

        /// <summary>Moves an item to a folder identified by EntryID.</summary>
        ComMoveItemResult? TryMoveItemToFolderId(
            string entryIdHex,
            string? storeId,
            string targetFolderEntryId,
            string targetStoreId,
            out string? error);

        /// <summary>Creates a new draft.</summary>
        ComDraftCreateResult? TryCreateNewDraft(
            string accountSmtpAddress,
            IReadOnlyList<string> toRecipients,
            string subject,
            ComDraftBody body,
            bool display,
            ComSignatureOverride? signatureOverride,
            ComDraftOptions? options,
            out string? error);

        /// <summary>Creates a reply, reply-all or forward draft via Outlook's own derivation.</summary>
        ComDraftCreateResult? TryCreateDerivedDraft(
            string sourceEntryIdHex,
            string? sourceStoreId,
            ComDerivedDraftKind kind,
            IReadOnlyList<string> toRecipients,
            ComDraftBody body,
            bool display,
            ComSignatureOverride? signatureOverride,
            ComDraftOptions? options,
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
        ComDraftDiscardResult? TryDiscardDraft(string entryIdHex, string? storeId, out string? error);

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
