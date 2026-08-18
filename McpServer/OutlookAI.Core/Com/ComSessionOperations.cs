using System;
using System.Collections.Generic;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Which <see cref="IOutlookSession"/> operations may be RE-RUN after a disconnect, and
    /// which may not. The one-shot rebuild in <see cref="ComGateway.Run{T}(Func{IOutlookSession, T}, ComSessionRecovery)"/>
    /// reads this and nothing else.
    /// <para>
    /// <b>Why a re-run is not free.</b> The rebuild fires on the RPC_E_DISCONNECTED family,
    /// and that family includes <c>RPC_S_CALL_FAILED</c> (0x800706BE), whose documented
    /// meaning is that the call MAY OR MAY NOT have executed on the server. So a re-run is
    /// not "the call did not happen, do it again": it is "the call may already have
    /// happened, do it a second time". For a read that costs a duplicate answer, which is
    /// the same answer. For <c>TrySendDraft</c> it costs a second copy of a mail in the
    /// recipient's inbox, which nothing can take back - and it would arrive AFTER the
    /// confirm token was consumed, so not one of the send path's guards (two-step token,
    /// content hash, identity check) is still standing at that point.
    /// </para>
    /// <para>
    /// <b>The rule.</b> An operation is retryable only when running it twice can leave
    /// nothing behind that running it once would not: it reads, it answers, it changes
    /// nothing. Anything that can create, edit, move, delete or send a mail item, write a
    /// file to disk, or move the user's Outlook window is NOT retryable, however unlikely a
    /// partial execution seems. The classification is by effect, never by how the method is
    /// named.
    /// </para>
    /// <para>
    /// <b>The guard.</b> Both sets are written with <c>nameof</c>, so renaming a contract
    /// method breaks the build here, and T1 asserts that the two sets together are exactly
    /// the contract's method set - so ADDING a method fails that test until it is
    /// classified. And an unclassified name is treated as mutating by
    /// <see cref="IsRetryable"/>, so the failure mode of forgetting is a lost retry on a
    /// read, never an extra send.
    /// </para>
    /// </summary>
    public static class ComSessionOperations
    {
        /// <summary>
        /// Operations that only READ. Re-running one after a disconnect asks the same
        /// question of a rebuilt session and is the entire point of the recovery.
        /// <para>
        /// <see cref="IOutlookSession.TryGetSendableDraftState"/> is here despite belonging
        /// to the send path: it reads the draft and hashes its content, and changes nothing.
        /// The consumable step is the send itself.
        /// </para>
        /// </summary>
        private static readonly HashSet<string> ReadOnlyNames = new HashSet<string>(StringComparer.Ordinal)
        {
            nameof(IOutlookSession.GetProfileName),
            nameof(IOutlookSession.GetAccounts),
            nameof(IOutlookSession.GetStoreDetails),
            nameof(IOutlookSession.ListFolders),
            nameof(IOutlookSession.ListFolderPaths),
            nameof(IOutlookSession.FindFolderPathsByLeafName),
            nameof(IOutlookSession.SweepFoldersNewerThan),
            nameof(IOutlookSession.ExhaustiveScan),
            nameof(IOutlookSession.TryReadItem),
            nameof(IOutlookSession.TryGetConversationItems),
            nameof(IOutlookSession.SnapshotAttachmentsById),
            nameof(IOutlookSession.TryResolveByPath),
            nameof(IOutlookSession.TryGetMailInfo),
            nameof(IOutlookSession.TryResolveArchiveFolder),
            nameof(IOutlookSession.TryGetSendableDraftState),
        };

        /// <summary>
        /// Operations with an effect that outlives the call, each one for a stated reason:
        /// <list type="bullet">
        /// <item><description><see cref="IOutlookSession.TrySendDraft"/> - a duplicate send
        /// cannot be recalled. This is the one that made the whole classification
        /// necessary.</description></item>
        /// <item><description><see cref="IOutlookSession.TryCreateNewDraft"/>,
        /// <see cref="IOutlookSession.TryCreateDerivedDraft"/> - a re-run leaves a second
        /// draft in the mailbox, and the caller only learns the id of the second one, so
        /// the first is orphaned where no cleanup will find it.</description></item>
        /// <item><description><see cref="IOutlookSession.TryUpdateDraft"/> - not idempotent:
        /// it appends attachments and can re-apply a signature, so a re-run doubles
        /// them.</description></item>
        /// <item><description><see cref="IOutlookSession.TryDiscardDraft"/>,
        /// <see cref="IOutlookSession.TryMoveItemToPath"/>,
        /// <see cref="IOutlookSession.TryMoveItemToFolderId"/> - the item is gone from where
        /// the second attempt expects to find it, so the re-run reports a failure that did
        /// not happen, over a move that did.</description></item>
        /// <item><description><see cref="IOutlookSession.TrySaveAttachment"/> - it writes a
        /// file to disk outside this process's control.</description></item>
        /// <item><description><see cref="IOutlookSession.TryDisplayItem"/>,
        /// <see cref="IOutlookSession.TryGotoFolder"/>,
        /// <see cref="IOutlookSession.TryShowSearchResults"/> - they drive the user's own
        /// Outlook window, and displaying an item can mark it read. A duplicated window is
        /// only an annoyance, but it is still an effect, and the rule is easier to keep
        /// right than a list of exceptions to it.</description></item>
        /// </list>
        /// </summary>
        private static readonly HashSet<string> MutatingNames = new HashSet<string>(StringComparer.Ordinal)
        {
            nameof(IOutlookSession.TrySaveAttachment),
            nameof(IOutlookSession.TryDisplayItem),
            nameof(IOutlookSession.TryGotoFolder),
            nameof(IOutlookSession.TryShowSearchResults),
            nameof(IOutlookSession.TryMoveItemToPath),
            nameof(IOutlookSession.TryMoveItemToFolderId),
            nameof(IOutlookSession.TryCreateNewDraft),
            nameof(IOutlookSession.TryCreateDerivedDraft),
            nameof(IOutlookSession.TryUpdateDraft),
            nameof(IOutlookSession.TryDiscardDraft),
            nameof(IOutlookSession.TrySendDraft),
        };

        /// <summary>The read-only half of the contract, for the T1 completeness guard.</summary>
        public static IReadOnlyCollection<string> ReadOnlyOperations => ReadOnlyNames;

        /// <summary>The mutating half of the contract, for the T1 completeness guard.</summary>
        public static IReadOnlyCollection<string> MutatingOperations => MutatingNames;

        /// <summary>
        /// Whether <paramref name="operationName"/> may be re-run against a rebuilt session.
        /// Fail-closed: a name in neither set answers false, so a contract method added
        /// without a decision loses its retry rather than gaining a second execution.
        /// </summary>
        public static bool IsRetryable(string? operationName)
        {
            return operationName != null && ReadOnlyNames.Contains(operationName);
        }

        /// <summary>
        /// Whether <paramref name="operationName"/> has been classified at all. Used by the
        /// T1 guard, which compares both sets against the contract itself.
        /// </summary>
        public static bool IsClassified(string? operationName)
        {
            return operationName != null
                && (ReadOnlyNames.Contains(operationName) || MutatingNames.Contains(operationName));
        }
    }
}
