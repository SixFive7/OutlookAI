using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Designated-Archive-folder resolution (D39, research 2026-07-26; Q84, 2026-09-24).
    ///
    /// How the one-click Archive folder is found:
    ///
    ///  1. <c>Store.GetDefaultFolder(39)</c>. 39 is an UNDOCUMENTED OlDefaultFolders value (the
    ///     public enum and this build's own type library stop at olFolderSuggestedContacts=30)
    ///     that Outlook's implementation nevertheless honors: it returns exactly the mailbox's
    ///     designated Archive folder - the one the Archive button/Backspace, mobile
    ///     swipe-archive and OWA use, with the server-designated (localized) name. Live-verified
    ///     on all 5 Exchange stores of the dev machine, including the Dutch-localized delegate
    ///     whose folder is "Archiveren".
    ///     <b>It is also a CREATING call</b> (measured 2026-09-24 on a POP3 PST, Outlook LTSC
    ///     2024 16.0.17932): on a store with no Archive folder it made one at the store root.
    ///     So it is used as it always was on an EXCHANGE store - where Microsoft documents the
    ///     Archive folder as a default folder that "can't be deleted" - and by the
    ///     move-to-archive path (<see cref="ResolveForMove"/>), which may create the folder
    ///     because the caller asked for mail to be moved into it, and reports that it did. A
    ///     read-only lookup on any other store never calls it (<see cref="ResolveReadOnly"/>).
    ///  2. <c>PR_IPM_ARCHIVE_ENTRYID</c> (PidTagIpmArchiveEntryId, 0x35FF0102, MS-OXPROPS 2.752).
    ///     MS-OXOSFLD 2.2.3 has listed it among the binary identification properties since
    ///     revision 16.0 (2024-11-12) - one entry id, and "The implementation MUST use the
    ///     Inbox folder when the mailbox is that of the owner, and it MUST use the Root folder
    ///     when the mailbox is that of a delegate". That is how a read-only lookup finds the folder on a
    ///     non-Exchange store without asking Outlook for it. D39 read the same property off the
    ///     STORE object (live probe: absent on all 5 cached Exchange stores), which the spec
    ///     does not name; that read is kept as a second source, never as proof of absence.
    ///  3. NEVER by folder NAME - localization makes name guessing wrong by design.
    ///
    /// Researched-and-rejected carriers (documented so nobody re-walks this path):
    /// PR_ADDITIONAL_REN_ENTRYIDS_EX (0x36D90102, PersistData blocks per MS-OXOSFLD) - the
    /// documented PersistID list (RSF_PID_* up to 0x800B) has no archive value and the
    /// property is absent from all 5 live stores; PR_ADDITIONAL_REN_ENTRYIDS (0x36D81102, on
    /// the Inbox) - live-probed: carries exactly the classic 5 slots (Conflicts/Sync
    /// Issues/Local Failures/Server Failures/Junk) plus one 4-byte non-EntryID trailer, no
    /// archive slot.
    /// </summary>
    public static class ArchiveFolderResolution
    {
        /// <summary>
        /// Undocumented OlDefaultFolders value resolving the designated Archive folder
        /// (see class remarks; live-proven on this build, feature-detected per store).
        /// </summary>
        public const int OlFolderArchive = 39;

        /// <summary>PR_IPM_ARCHIVE_ENTRYID (PidTagIpmArchiveEntryId, 0x35FF, PT_BINARY) PropertyAccessor schema name.</summary>
        public const string ArchiveEntryIdPropertySchema = "http://schemas.microsoft.com/mapi/proptag/0x35FF0102";

        /// <summary>Resolved by <c>Store.GetDefaultFolder(39)</c>.</summary>
        public const string ViaOutlookDefaultFolder = "outlookDefaultFolder";

        /// <summary>Resolved from PR_IPM_ARCHIVE_ENTRYID on the store object.</summary>
        public const string ViaStoreArchiveProperty = "storeArchiveProperty";

        /// <summary>Resolved from PR_IPM_ARCHIVE_ENTRYID on the Inbox (MS-OXOSFLD 2.2.3), without asking Outlook for the folder.</summary>
        public const string ViaInboxArchiveProperty = "inboxArchiveProperty";

        /// <summary>The store has no designated Archive folder, and nothing was created.</summary>
        public const string NoDesignatedArchiveFolder = "NoDesignatedArchiveFolder";

        /// <summary>
        /// A read-only lookup could not establish whether a non-Exchange store has a designated
        /// Archive folder without asking Outlook for it - which would create one. Nothing was
        /// created.
        /// </summary>
        public const string ArchiveDesignationUnreadable = "ArchiveDesignationUnreadable";

        /// <summary>
        /// The move path could not list the store's top-level folders, so it could not tell
        /// whether resolving the Archive folder would CREATE one - and it did not try. Nothing
        /// was created.
        /// </summary>
        public const string ArchiveFolderStateUnreadable = "ArchiveFolderStateUnreadable";

        /// <summary>
        /// The move path's creating call returned no Archive folder, and the store's top-level
        /// folders no longer list as they did before it (or could not be listed again) - so it
        /// cannot be ruled out that the call created something. Nothing was moved.
        /// </summary>
        public const string ArchiveFolderCreationUnverified = "ArchiveFolderCreationUnverified";

        /// <summary>The core default folders a designated archive must never be: Deleted Items, Outbox, Sent Items, Inbox, Drafts, Junk.</summary>
        private static readonly int[] CoreDefaultFolderIds =
        {
            SpecialFolders.OlFolderDeletedItems,
            SpecialFolders.OlFolderOutbox,
            SpecialFolders.OlFolderSentMail,
            SpecialFolders.OlFolderInbox,
            SpecialFolders.OlFolderDrafts,
            SpecialFolders.OlFolderJunk,
        };

        /// <summary>
        /// Interprets a PropertyAccessor.GetProperty result as ONE EntryID hex string:
        /// accepts a byte[] (PT_BINARY), the first non-empty byte[] of an array
        /// (PT_MV_BINARY), or an existing plausible hex string. Returns null for
        /// anything else (missing, empty, junk) - resolution then falls through to the
        /// per-item "no designated archive folder" error instead of guessing.
        /// </summary>
        public static string? TryReadEntryIdHex(object? propertyValue)
        {
            switch (propertyValue)
            {
                case null:
                    return null;
                case byte[] bytes:
                    return ToHexOrNull(bytes);
                case string text:
                    string trimmed = text.Trim();
                    return trimmed.Length >= 8 && trimmed.Length % 2 == 0 && IsHex(trimmed)
                        ? trimmed.ToUpperInvariant()
                        : null;
                case IEnumerable values:
                    foreach (object? entry in values)
                    {
                        if (entry is byte[] entryBytes)
                        {
                            string? hex = ToHexOrNull(entryBytes);
                            if (hex != null)
                            {
                                return hex;
                            }
                        }
                    }

                    return null;
                default:
                    return null;
            }
        }

        /// <summary>
        /// READ-ONLY resolution: never creates anything. An Exchange store is resolved exactly
        /// as before (<c>GetDefaultFolder(39)</c>, then the store-object property); any other
        /// store only from its PR_IPM_ARCHIVE_ENTRYID designation. A store without one answers
        /// <see cref="NoDesignatedArchiveFolder"/>; one whose designation could not be read
        /// answers <see cref="ArchiveDesignationUnreadable"/>. The candidate is then verified
        /// (<see cref="VerifyCandidate"/>) without creating anything either.
        /// </summary>
        public static ArchiveFolderAnswer ResolveReadOnly(ISpecialFolderStore store, string storeId)
        {
            if (store == null)
            {
                throw new ArgumentNullException(nameof(store));
            }

            OutlookComSession.DefaultFolderResolution resolution =
                SpecialFolders.Resolve(store, OlFolderArchive, out object? folder, out SpecialFolderSource source);
            try
            {
                if (resolution != OutlookComSession.DefaultFolderResolution.Resolved)
                {
                    bool unreadable = resolution == OutlookComSession.DefaultFolderResolution.Unreadable
                        && !SpecialFolders.IsExchangeStore(store.ExchangeStoreType);
                    return ArchiveFolderAnswer.Failed(unreadable ? ArchiveDesignationUnreadable : NoDesignatedArchiveFolder);
                }

                return Complete(store, storeId, folder!, ViaFor(source), created: false);
            }
            finally
            {
                store.Release(folder);
            }
        }

        /// <summary>
        /// The MOVE-TO-ARCHIVE resolution: the one lookup allowed to create the Archive folder,
        /// because the caller asked for mail to be moved into it (Q84 (c)). It resolves exactly
        /// as archive_mail always did - <c>GetDefaultFolder(39)</c>, then the store-object
        /// property - and on a non-Exchange store it first lists the root folder's children, so
        /// that a folder the call CREATED is recognised and reported
        /// (<see cref="ArchiveFolderAnswer.Created"/>) on every path, including one whose
        /// verification then refuses the folder. If that list cannot be read, nothing is called
        /// and nothing is created (<see cref="ArchiveFolderStateUnreadable"/>).
        /// <para>
        /// An Exchange store is resolved like the read-only path: its Archive folder is a
        /// default folder the server keeps, so there is nothing to create or report. Its root is
        /// not listed either - a delegate store's hierarchy syncs lazily (D42), so a listing
        /// there could miss a folder that exists and report it as created.
        /// </para>
        /// </summary>
        public static ArchiveFolderAnswer ResolveForMove(ISpecialFolderStore store, string storeId)
        {
            if (store == null)
            {
                throw new ArgumentNullException(nameof(store));
            }

            if (SpecialFolders.IsExchangeStore(store.ExchangeStoreType))
            {
                return ResolveReadOnly(store, storeId);
            }

            IReadOnlyList<string>? before = store.ListRootChildEntryIds();
            if (before == null)
            {
                return ArchiveFolderAnswer.Failed(ArchiveFolderStateUnreadable);
            }

            object? folder = null;
            try
            {
                string via = ViaOutlookDefaultFolder;
                OutlookComSession.DefaultFolderResolution resolution =
                    SpecialFolders.GetDefaultFolderMayCreate(store, OlFolderArchive, out folder);
                if (resolution != OutlookComSession.DefaultFolderResolution.Resolved)
                {
                    via = ViaStoreArchiveProperty;
                    resolution = SpecialFolders.OpenStoreArchiveDesignation(store, out folder);
                    if (resolution != OutlookComSession.DefaultFolderResolution.Resolved)
                    {
                        // The creating call returned no folder. "Nothing was created" is then a
                        // claim to PROVE, not assume: the root must list exactly as it did.
                        return ArchiveFolderAnswer.Failed(
                            RootUnchanged(store, before) ? NoDesignatedArchiveFolder : ArchiveFolderCreationUnverified);
                    }
                }

                // Only the creating call can have made the folder: the store-object fallback
                // opens one that already exists, wherever it lives.
                string? entryId = store.EntryIdOf(folder!);
                bool created = via == ViaOutlookDefaultFolder
                    && entryId != null
                    && !before.Contains(entryId, StringComparer.OrdinalIgnoreCase);
                return Complete(store, storeId, folder!, via, created);
            }
            finally
            {
                store.Release(folder);
            }
        }

        /// <summary>
        /// Verification of a resolved candidate: it must live in the SAME store, be a mail
        /// folder, and not be one of the core default folders (Deleted Items/Outbox/Sent/
        /// Inbox/Drafts/Junk) - mis-designating any of those as "archive" would make
        /// archive_mail silently do something else. The core defaults are looked up with
        /// <see cref="SpecialFolders.Resolve"/>, so this CREATES NOTHING: the old comparison
        /// asked <c>GetDefaultFolder(23)</c>, which made a <c>Junk Email</c> folder in a PST
        /// that had none (measured 2026-09-24). A core default the store does not have cannot
        /// be the candidate; one that will not resolve is skipped, as a failure always was.
        /// Returns a content-free error, or null when the candidate is sound.
        /// </summary>
        public static string? VerifyCandidate(ISpecialFolderStore store, string storeId, SpecialFolderFacts candidate)
        {
            if (store == null)
            {
                throw new ArgumentNullException(nameof(store));
            }

            if (candidate == null)
            {
                throw new ArgumentNullException(nameof(candidate));
            }

            if (candidate.StoreId == null || candidate.EntryId == null)
            {
                return "ArchiveFolderVerificationFailed:probe";
            }

            if (!string.Equals(candidate.StoreId, storeId, StringComparison.OrdinalIgnoreCase))
            {
                return "ArchiveFolderVerificationFailed:store";
            }

            if (candidate.DefaultItemType == null)
            {
                return "ArchiveFolderVerificationFailed:probe";
            }

            if (candidate.DefaultItemType.Value != 0)
            {
                return "ArchiveFolderVerificationFailed:itemType";
            }

            foreach (int coreDefault in CoreDefaultFolderIds)
            {
                OutlookComSession.DefaultFolderResolution resolution =
                    SpecialFolders.Resolve(store, coreDefault, out object? coreFolder, out _);
                try
                {
                    if (resolution == OutlookComSession.DefaultFolderResolution.Resolved
                        && string.Equals(store.EntryIdOf(coreFolder!), candidate.EntryId, StringComparison.OrdinalIgnoreCase))
                    {
                        return "ArchiveFolderVerificationFailed:coreDefault";
                    }
                }
                finally
                {
                    store.Release(coreFolder);
                }
            }

            return null;
        }

        /// <summary>True only when the root lists exactly the children it listed before; an unreadable list proves nothing.</summary>
        private static bool RootUnchanged(ISpecialFolderStore store, IReadOnlyList<string> before)
        {
            IReadOnlyList<string>? after = store.ListRootChildEntryIds();
            return after != null
                && after.Count == before.Count
                && after.All(id => before.Contains(id, StringComparer.OrdinalIgnoreCase));
        }

        private static ArchiveFolderAnswer Complete(ISpecialFolderStore store, string storeId, object folder, string via, bool created)
        {
            SpecialFolderFacts facts = store.Describe(folder);
            string? verification = VerifyCandidate(store, storeId, facts);
            return verification != null
                ? ArchiveFolderAnswer.Refused(verification, created, facts.FolderPath)
                : ArchiveFolderAnswer.Resolved(facts, via, created);
        }

        private static string ViaFor(SpecialFolderSource source)
        {
            switch (source)
            {
                case SpecialFolderSource.InboxDesignation:
                    return ViaInboxArchiveProperty;
                case SpecialFolderSource.StoreDesignation:
                    return ViaStoreArchiveProperty;
                default:
                    return ViaOutlookDefaultFolder;
            }
        }

        private static string? ToHexOrNull(byte[] bytes)
        {
            // A real folder EntryID is dozens of bytes; anything shorter than 4 is a
            // marker/filler (live-probed: the Inbox slot list ends with a 4-byte
            // non-EntryID trailer), not an id.
            if (bytes.Length < 4)
            {
                return null;
            }

            StringBuilder sb = new StringBuilder(bytes.Length * 2);
            foreach (byte b in bytes)
            {
                sb.Append(b.ToString("X2", System.Globalization.CultureInfo.InvariantCulture));
            }

            return sb.ToString();
        }

        private static bool IsHex(string value)
        {
            foreach (char c in value)
            {
                bool ok = (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
                if (!ok)
                {
                    return false;
                }
            }

            return true;
        }
    }

    /// <summary>One archive-folder resolution, read by <c>OutlookComSession</c> and turned into the wire shapes.</summary>
    public sealed class ArchiveFolderAnswer
    {
        private ArchiveFolderAnswer(string? error, string? entryId, string? name, string? folderPath, string? via, bool created)
        {
            Error = error;
            EntryId = entryId;
            Name = name;
            FolderPath = folderPath;
            Via = via;
            Created = created;
        }

        /// <summary>Content-free error token, or null when the folder resolved and passed verification.</summary>
        public string? Error { get; }

        /// <summary>The folder's EntryID (resolved answers only).</summary>
        public string? EntryId { get; }

        /// <summary>The folder's localized name (resolved answers only).</summary>
        public string? Name { get; }

        /// <summary>The folder's <c>Folder.FolderPath</c> - also set on a refusal that follows a creation, so the caller can report where.</summary>
        public string? FolderPath { get; }

        /// <summary>Resolution mechanism (resolved answers only).</summary>
        public string? Via { get; }

        /// <summary>True when THIS call created the folder - reported whether the answer then resolved or was refused.</summary>
        public bool Created { get; }

        internal static ArchiveFolderAnswer Failed(string error) =>
            new ArchiveFolderAnswer(error, null, null, null, null, created: false);

        internal static ArchiveFolderAnswer Refused(string error, bool created, string? folderPath) =>
            new ArchiveFolderAnswer(error, null, null, created ? folderPath : null, null, created);

        internal static ArchiveFolderAnswer Resolved(SpecialFolderFacts facts, string via, bool created) =>
            new ArchiveFolderAnswer(null, facts.EntryId, facts.Name, facts.FolderPath, via, created);
    }
}
