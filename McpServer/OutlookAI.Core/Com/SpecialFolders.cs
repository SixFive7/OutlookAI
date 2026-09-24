using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;

namespace OutlookAI.Core.Com
{
    /// <summary>How one property read on a store or a folder came out.</summary>
    public enum PropertyReadStatus
    {
        /// <summary>The property is there and <see cref="PropertyRead.Value"/> holds it.</summary>
        Found = 0,

        /// <summary>The object does not carry the property (<c>MAPI_E_NOT_FOUND</c>).</summary>
        NotFound = 1,

        /// <summary>The read failed for any other reason, so nothing is known about the property.</summary>
        Failed = 2,
    }

    /// <summary>One property read, with the status that decides what the value may be taken to mean.</summary>
    public readonly struct PropertyRead
    {
        private PropertyRead(PropertyReadStatus status, object? value)
        {
            Status = status;
            Value = value;
        }

        /// <summary>Found, not found, or failed.</summary>
        public PropertyReadStatus Status { get; }

        /// <summary>The value, when <see cref="Status"/> is <see cref="PropertyReadStatus.Found"/>.</summary>
        public object? Value { get; }

        /// <summary>A read that returned a value.</summary>
        public static PropertyRead Found(object? value) => new PropertyRead(PropertyReadStatus.Found, value);

        /// <summary>A read the object answered with "no such property".</summary>
        public static PropertyRead Missing() => new PropertyRead(PropertyReadStatus.NotFound, null);

        /// <summary>A read that failed for any other reason.</summary>
        public static PropertyRead Failure() => new PropertyRead(PropertyReadStatus.Failed, null);
    }

    /// <summary>What the archive answer and its verification read off one folder. Any member is null when it would not read.</summary>
    public sealed class SpecialFolderFacts
    {
        /// <summary>Creates the facts of one folder.</summary>
        public SpecialFolderFacts(string? entryId, string? name, string? folderPath, string? storeId, int? defaultItemType)
        {
            EntryId = entryId;
            Name = name;
            FolderPath = folderPath;
            StoreId = storeId;
            DefaultItemType = defaultItemType;
        }

        /// <summary><c>Folder.EntryID</c>.</summary>
        public string? EntryId { get; }

        /// <summary><c>Folder.Name</c> (localized).</summary>
        public string? Name { get; }

        /// <summary><c>Folder.FolderPath</c> (<c>\\Store\A\B</c>).</summary>
        public string? FolderPath { get; }

        /// <summary><c>Folder.Store.StoreID</c>.</summary>
        public string? StoreId { get; }

        /// <summary><c>Folder.DefaultItemType</c> (0 = mail).</summary>
        public int? DefaultItemType { get; }
    }

    /// <summary>Where a resolved special folder came from.</summary>
    public enum SpecialFolderSource
    {
        /// <summary>Nothing resolved.</summary>
        None = 0,

        /// <summary><c>Store.GetDefaultFolder</c>, called on an Exchange store as before, or on another store only once the folder was proven to exist.</summary>
        DefaultFolderCall = 1,

        /// <summary>The entry id designated on the store's Inbox (MS-OXOSFLD 2.2.3 / 2.2.4), opened with <c>GetFolderFromID</c>.</summary>
        InboxDesignation = 2,

        /// <summary>The entry id designated on the store object itself, opened with <c>GetFolderFromID</c>.</summary>
        StoreDesignation = 3,
    }

    /// <summary>
    /// What a special-folder lookup needs from ONE store. It exists so the rule that decides
    /// whether <c>Store.GetDefaultFolder</c> may be called can run against a fake in T1: the COM
    /// half (<see cref="ComSpecialFolderStore"/>) only answers these questions, and every
    /// decision is made in <see cref="SpecialFolders"/>.
    /// </summary>
    public interface ISpecialFolderStore
    {
        /// <summary><c>Store.ExchangeStoreType</c> (<c>OlExchangeStoreType</c>), or null when it would not read.</summary>
        int? ExchangeStoreType { get; }

        /// <summary><c>Store.PropertyAccessor.GetProperty</c> on the STORE object.</summary>
        PropertyRead ReadStoreProperty(string schemaName);

        /// <summary><c>Folder.PropertyAccessor.GetProperty</c> on a folder this store handed out.</summary>
        PropertyRead ReadFolderProperty(object folder, string schemaName);

        /// <summary>
        /// <c>Store.GetDefaultFolder</c>. THE CREATING CALL: on a store that does not have the
        /// folder, Outlook makes it (measured 2026-09-24 for 39 = Archive and 23 = Junk Email on
        /// a POP3 PST). Throws what the COM call throws.
        /// </summary>
        object? GetDefaultFolder(int olDefaultFolderId);

        /// <summary><c>NameSpace.GetFolderFromID</c>: opens an existing folder, never makes one.</summary>
        PropertyReadStatus OpenFolder(string entryIdHex, out object? folder);

        /// <summary>
        /// The <c>EntryID</c> of every direct child of the store's root folder, or null when the
        /// list could not be read completely. Never creates anything.
        /// </summary>
        IReadOnlyList<string>? ListRootChildEntryIds();

        /// <summary><c>Folder.EntryID</c>, or null when it would not read.</summary>
        string? EntryIdOf(object folder);

        /// <summary>The facts the archive answer and its verification need, read without side effects.</summary>
        SpecialFolderFacts Describe(object folder);

        /// <summary>Releases a COM object this store handed out.</summary>
        void Release(object? comObject);
    }

    /// <summary>
    /// Looks up a store's special (default) folders WITHOUT ever asking Outlook to make one -
    /// the only way a read-only path may look one up (Q84, maintainer decision (c),
    /// 2026-09-24: read-only lookups never create folders; only the move-to-archive path may).
    /// <para>
    /// <b>Why <c>Store.GetDefaultFolder</c> cannot be the read-only answer.</b> Its reference
    /// page says that for a folder the store does not have it "returns Null". Measured on a
    /// POP3 PST (Outlook LTSC 2024 16.0.17932, 2026-09-24) it does not: <c>GetDefaultFolder(39)</c>
    /// CREATED an <c>Archive</c> folder and <c>GetDefaultFolder(23)</c> CREATED <c>Junk Email</c>.
    /// That is the behaviour MS-OXOSFLD section 3.1.4 prescribes for a client - "If the ID cannot
    /// be retrieved, or the folder cannot be opened, or the special folder does not exist within
    /// the message store, the client MUST create the special folder" - and
    /// <c>NameSpace.GetDefaultFolder</c>'s own page admits it: "depending on the type, Outlook
    /// may create and return the folder". So the call is safe exactly when the folder already
    /// exists, and the rule here is to establish that FIRST, from where MS-OXOSFLD says a
    /// special folder is identified, and never to call it otherwise.
    /// </para>
    /// <para>
    /// <b>The two branches.</b> An EXCHANGE store (<c>ExchangeStoreType</c> other than
    /// <c>olNotExchange</c>) is resolved exactly as before, with <c>GetDefaultFolder</c> -
    /// Exchange must keep working exactly as it did, and on an Exchange mailbox every folder
    /// looked up here is a server default folder: MS-OXOSFLD 2.2.2 has the server identify
    /// Deleted Items, Outbox, Sent Items and the Inbox at logon, and Microsoft Support
    /// ("Archive in Outlook for Windows") says the Archive folder is "one of Outlook's default
    /// folders" on Microsoft 365, Outlook.com and Exchange accounts and "can't be deleted".
    /// EVERY OTHER store (a PST, an IMAP or POP account, a data file) is where the creation was
    /// measured, and there nothing is called until the folder is proven to exist:
    /// </para>
    /// <list type="bullet">
    /// <item><description>Inbox, Outbox, Deleted Items, Sent Items: the store's
    /// <c>PR_VALID_FOLDER_MASK</c> (0x35DF, PT_LONG). Microsoft's "Opening a message store
    /// folder" tells a client to read it BEFORE asking for one of these folders: "If the bit is
    /// set, it indicates that the corresponding folder is supported and has a valid entry
    /// identifier." Bit values from Microsoft's MAPIDefS.h: FOLDER_IPM_INBOX_VALID 0x02,
    /// FOLDER_IPM_OUTBOX_VALID 0x04, FOLDER_IPM_WASTEBASKET_VALID 0x08,
    /// FOLDER_IPM_SENTMAIL_VALID 0x10. Bit set: <c>GetDefaultFolder</c> returns the folder that
    /// is there. Bit clear: absent, and nothing is called.</description></item>
    /// <item><description>Drafts and Archive: the binary identification properties on the Inbox
    /// (MS-OXOSFLD 2.2.3: "The implementation MUST use the Inbox folder when the mailbox is that
    /// of the owner") - <c>PR_IPM_DRAFTS_ENTRYID</c> (0x36D70102, MS-OXPROPS 2.754) and
    /// <c>PR_IPM_ARCHIVE_ENTRYID</c> (0x35FF0102, MS-OXPROPS 2.752, added to MS-OXOSFLD 2.2.3
    /// in revision 16.0 of 2024-11-12), each ONE entry id; the store object is read second,
    /// where MAPI's "Opening a message store folder" and D39 looked.</description></item>
    /// <item><description>Junk E-mail and the Sync Issues family: <c>PR_ADDITIONAL_REN_ENTRYIDS</c>
    /// (0x36D81102, PT_MV_BINARY) on the Inbox, MS-OXOSFLD 2.2.4, by index: 0 Conflicts,
    /// 1 Sync Issues, 2 Local Failures, 3 Server Failures, 4 Junk E-mail ("The implementation
    /// MUST ignore and MUST preserve data at other indexes").</description></item>
    /// </list>
    /// <para>
    /// A designated entry id is opened with <c>GetFolderFromID</c>, which cannot create
    /// anything. When the designation is not there the folder is ABSENT; when it cannot be read
    /// the answer is UNREADABLE - never absent, because reading a failure as absence would
    /// silently drop a folder whose mail really is missing from an answer
    /// (<see cref="OutlookComSession.ClassifyDefaultFolder"/> makes the same refusal).
    /// </para>
    /// </summary>
    public static class SpecialFolders
    {
        /// <summary><c>olFolderDeletedItems</c>.</summary>
        public const int OlFolderDeletedItems = 3;

        /// <summary><c>olFolderOutbox</c>.</summary>
        public const int OlFolderOutbox = 4;

        /// <summary><c>olFolderSentMail</c>.</summary>
        public const int OlFolderSentMail = 5;

        /// <summary><c>olFolderInbox</c>.</summary>
        public const int OlFolderInbox = 6;

        /// <summary><c>olFolderDrafts</c>.</summary>
        public const int OlFolderDrafts = 16;

        /// <summary><c>olFolderConflicts</c>.</summary>
        public const int OlFolderConflicts = 19;

        /// <summary><c>olFolderSyncIssues</c>.</summary>
        public const int OlFolderSyncIssues = 20;

        /// <summary><c>olFolderLocalFailures</c>.</summary>
        public const int OlFolderLocalFailures = 21;

        /// <summary><c>olFolderServerFailures</c>.</summary>
        public const int OlFolderServerFailures = 22;

        /// <summary><c>olFolderJunk</c>.</summary>
        public const int OlFolderJunk = 23;

        /// <summary><c>OlExchangeStoreType.olNotExchange</c>: "the store is not an Exchange store".</summary>
        public const int OlNotExchange = 3;

        /// <summary>PR_VALID_FOLDER_MASK (0x35DF, PT_LONG) on the store object.</summary>
        public const string ValidFolderMaskSchema = "http://schemas.microsoft.com/mapi/proptag/0x35DF0003";

        /// <summary>PR_IPM_DRAFTS_ENTRYID (PidTagIpmDraftsEntryId, 0x36D7, PT_BINARY).</summary>
        public const string DraftsEntryIdSchema = "http://schemas.microsoft.com/mapi/proptag/0x36D70102";

        /// <summary>PR_ADDITIONAL_REN_ENTRYIDS (PidTagAdditionalRenEntryIds, 0x36D8, PT_MV_BINARY).</summary>
        public const string AdditionalRenEntryIdsSchema = "http://schemas.microsoft.com/mapi/proptag/0x36D81102";

        /// <summary>FOLDER_IPM_INBOX_VALID (MAPIDefS.h).</summary>
        public const int FolderIpmInboxValid = 0x00000002;

        /// <summary>FOLDER_IPM_OUTBOX_VALID (MAPIDefS.h).</summary>
        public const int FolderIpmOutboxValid = 0x00000004;

        /// <summary>FOLDER_IPM_WASTEBASKET_VALID (MAPIDefS.h).</summary>
        public const int FolderIpmWastebasketValid = 0x00000008;

        /// <summary>FOLDER_IPM_SENTMAIL_VALID (MAPIDefS.h).</summary>
        public const int FolderIpmSentmailValid = 0x00000010;

        /// <summary>MAPI_E_NOT_FOUND - what a property read or <c>GetFolderFromID</c> answers for something that is not there.</summary>
        public const int MapiENotFound = unchecked((int)0x8004010F);

        /// <summary>
        /// True for an Exchange store. An unreadable type answers false: the non-Exchange branch
        /// never creates, so it is the safe one to fall into.
        /// </summary>
        public static bool IsExchangeStore(int? exchangeStoreType)
        {
            return exchangeStoreType.HasValue && exchangeStoreType.Value != OlNotExchange;
        }

        /// <summary>The PR_VALID_FOLDER_MASK bit that proves a folder exists, or null for a folder the mask does not cover.</summary>
        public static int? ValidFolderBit(int olDefaultFolderId)
        {
            switch (olDefaultFolderId)
            {
                case OlFolderInbox:
                    return FolderIpmInboxValid;
                case OlFolderOutbox:
                    return FolderIpmOutboxValid;
                case OlFolderDeletedItems:
                    return FolderIpmWastebasketValid;
                case OlFolderSentMail:
                    return FolderIpmSentmailValid;
                default:
                    return null;
            }
        }

        /// <summary>The PR_ADDITIONAL_REN_ENTRYIDS index of a folder (MS-OXOSFLD 2.2.4), or null for one it does not list.</summary>
        public static int? AdditionalRenIndex(int olDefaultFolderId)
        {
            switch (olDefaultFolderId)
            {
                case OlFolderConflicts:
                    return 0;
                case OlFolderSyncIssues:
                    return 1;
                case OlFolderLocalFailures:
                    return 2;
                case OlFolderServerFailures:
                    return 3;
                case OlFolderJunk:
                    return 4;
                default:
                    return null;
            }
        }

        /// <summary>The designation property a folder is identified by on the Inbox, or null for one that has none.</summary>
        public static string? DesignationSchema(int olDefaultFolderId)
        {
            if (olDefaultFolderId == OlFolderDrafts)
            {
                return DraftsEntryIdSchema;
            }

            if (olDefaultFolderId == ArchiveFolderResolution.OlFolderArchive)
            {
                return ArchiveFolderResolution.ArchiveEntryIdPropertySchema;
            }

            return AdditionalRenIndex(olDefaultFolderId).HasValue ? AdditionalRenEntryIdsSchema : null;
        }

        /// <summary>
        /// True when a failure is MAPI_E_NOT_FOUND, the answer a property read gives for a
        /// property the object does not carry (Outlook: "The property ... is unknown or cannot
        /// be found") and <c>GetFolderFromID</c> gives for an entry id nothing answers to.
        /// </summary>
        public static bool IsMapiNotFound(Exception? failure)
        {
            return failure is COMException com && com.HResult == MapiENotFound;
        }

        /// <summary>
        /// Reads a PR_VALID_FOLDER_MASK value. PT_LONG arrives as an int; any other integral
        /// shape is accepted, anything else is not a mask.
        /// </summary>
        public static bool TryReadMask(object? value, out int mask)
        {
            switch (value)
            {
                case int i:
                    mask = i;
                    return true;
                case uint u:
                    mask = unchecked((int)u);
                    return true;
                case short s:
                    mask = s;
                    return true;
                case ushort us:
                    mask = us;
                    return true;
                case byte b:
                    mask = b;
                    return true;
                case long l when l >= int.MinValue && l <= uint.MaxValue:
                    mask = unchecked((int)l);
                    return true;
                default:
                    mask = 0;
                    return false;
            }
        }

        /// <summary>
        /// Reads ONE entry id out of a PT_BINARY designation. Accepts what the old archive
        /// fallback accepted (<see cref="ArchiveFolderResolution.TryReadEntryIdHex"/>): a byte
        /// array, a plausible hex string, or the first usable byte array of an array. An empty or
        /// sub-4-byte value designates nothing; a value of any other shape is unrecognised.
        /// </summary>
        public static DesignatedEntryId ReadEntryId(object? value)
        {
            switch (value)
            {
                case null:
                    return DesignatedEntryId.None;
                case byte[] bytes:
                    return bytes.Length < 4 ? DesignatedEntryId.None : DesignatedEntryId.Of(ToHex(bytes));
                case string text:
                    string? hex = ArchiveFolderResolution.TryReadEntryIdHex(text);
                    return hex == null ? DesignatedEntryId.Unrecognised : DesignatedEntryId.Of(hex);
                case IEnumerable:
                    string? first = ArchiveFolderResolution.TryReadEntryIdHex(value);
                    return first == null ? DesignatedEntryId.None : DesignatedEntryId.Of(first);
                default:
                    return DesignatedEntryId.Unrecognised;
            }
        }

        /// <summary>
        /// Reads the entry id at <paramref name="index"/> of a PT_MV_BINARY designation
        /// (PR_ADDITIONAL_REN_ENTRYIDS). The property is a one-dimensional array of byte
        /// arrays; an index past its end, or an empty slot, designates nothing. A single byte
        /// array, or any other shape, is not this property and is unrecognised - reading it by
        /// position would pick a byte, not a folder.
        /// </summary>
        public static DesignatedEntryId ReadEntryIdAt(object? value, int index)
        {
            if (index < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(index));
            }

            if (value == null)
            {
                return DesignatedEntryId.None;
            }

            if (value is byte[] || !(value is Array array) || array.Rank != 1)
            {
                return DesignatedEntryId.Unrecognised;
            }

            if (index >= array.Length)
            {
                return DesignatedEntryId.None;
            }

            object? slot = array.GetValue(array.GetLowerBound(0) + index);
            switch (slot)
            {
                case null:
                    return DesignatedEntryId.None;
                case byte[] bytes:
                    return bytes.Length < 4 ? DesignatedEntryId.None : DesignatedEntryId.Of(ToHex(bytes));
                default:
                    return DesignatedEntryId.Unrecognised;
            }
        }

        /// <summary>
        /// Resolves ONE special folder without ever creating it. See the class remarks for the
        /// two branches and the documentation each rests on. <paramref name="folder"/> is set
        /// only for <see cref="OutlookComSession.DefaultFolderResolution.Resolved"/>, and the
        /// caller releases it.
        /// </summary>
        public static OutlookComSession.DefaultFolderResolution Resolve(
            ISpecialFolderStore store,
            int olDefaultFolderId,
            out object? folder,
            out SpecialFolderSource source)
        {
            if (store == null)
            {
                throw new ArgumentNullException(nameof(store));
            }

            folder = null;
            source = SpecialFolderSource.None;

            if (IsExchangeStore(store.ExchangeStoreType))
            {
                // Exactly as before: the folder is a server default folder of the mailbox.
                OutlookComSession.DefaultFolderResolution asBefore = CallDefaultFolder(store, olDefaultFolderId, out folder);
                if (asBefore == OutlookComSession.DefaultFolderResolution.Resolved)
                {
                    source = SpecialFolderSource.DefaultFolderCall;
                    return asBefore;
                }

                if (olDefaultFolderId == ArchiveFolderResolution.OlFolderArchive
                    && OpenDesignated(store, store.ReadStoreProperty(ArchiveFolderResolution.ArchiveEntryIdPropertySchema), null, out folder)
                        == OutlookComSession.DefaultFolderResolution.Resolved)
                {
                    // The D39 fallback, unchanged: PR_IPM_ARCHIVE_ENTRYID on the store object.
                    source = SpecialFolderSource.StoreDesignation;
                    return OutlookComSession.DefaultFolderResolution.Resolved;
                }

                return asBefore;
            }

            int? bit = ValidFolderBit(olDefaultFolderId);
            if (bit.HasValue)
            {
                OutlookComSession.DefaultFolderResolution proven = ResolveByValidMask(store, olDefaultFolderId, bit.Value, out folder);
                if (proven == OutlookComSession.DefaultFolderResolution.Resolved)
                {
                    source = SpecialFolderSource.DefaultFolderCall;
                }

                return proven;
            }

            string? schema = DesignationSchema(olDefaultFolderId);
            if (schema == null)
            {
                // Nothing documents where this folder is identified, so its existence cannot be
                // established without asking Outlook for it - which is the call that creates.
                return OutlookComSession.DefaultFolderResolution.Unreadable;
            }

            return ResolveByDesignation(store, olDefaultFolderId, schema, out folder, out source);
        }

        /// <summary>
        /// <c>Store.GetDefaultFolder</c> with NO proof that the folder exists: on a store without
        /// it, Outlook creates it. For the move-to-archive path ONLY - the caller asked for mail
        /// to be moved INTO the folder - and that path must report what it created
        /// (<see cref="ArchiveFolderResolution.ResolveForMove"/>). Everything else goes through
        /// <see cref="Resolve"/>.
        /// </summary>
        public static OutlookComSession.DefaultFolderResolution GetDefaultFolderMayCreate(
            ISpecialFolderStore store,
            int olDefaultFolderId,
            out object? folder)
        {
            if (store == null)
            {
                throw new ArgumentNullException(nameof(store));
            }

            return CallDefaultFolder(store, olDefaultFolderId, out folder);
        }

        /// <summary>
        /// Opens the folder a store-object PR_IPM_ARCHIVE_ENTRYID names - D39's fallback, which
        /// the move path keeps exactly as it was. Never creates.
        /// </summary>
        public static OutlookComSession.DefaultFolderResolution OpenStoreArchiveDesignation(
            ISpecialFolderStore store,
            out object? folder)
        {
            if (store == null)
            {
                throw new ArgumentNullException(nameof(store));
            }

            return OpenDesignated(store, store.ReadStoreProperty(ArchiveFolderResolution.ArchiveEntryIdPropertySchema), null, out folder);
        }

        /// <summary>
        /// The ONE place a read-only lookup calls <c>Store.GetDefaultFolder</c>: on an Exchange
        /// store as before, and on any other store only after <see cref="ResolveByValidMask"/>
        /// has proven the folder is there. Classified by the same pure rule the sweep has always
        /// used (<see cref="OutlookComSession.ClassifyDefaultFolder"/>).
        /// </summary>
        private static OutlookComSession.DefaultFolderResolution CallDefaultFolder(
            ISpecialFolderStore store,
            int olDefaultFolderId,
            out object? folder)
        {
            folder = null;
            object? resolved = null;
            Exception? failure = null;
            try
            {
                resolved = store.GetDefaultFolder(olDefaultFolderId);
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                failure = ex;
            }

            OutlookComSession.DefaultFolderResolution resolution = OutlookComSession.ClassifyDefaultFolder(resolved, failure);
            if (resolution == OutlookComSession.DefaultFolderResolution.Resolved)
            {
                folder = resolved;
            }

            return resolution;
        }

        private static OutlookComSession.DefaultFolderResolution ResolveByValidMask(
            ISpecialFolderStore store,
            int olDefaultFolderId,
            int bit,
            out object? folder)
        {
            folder = null;
            PropertyRead mask = store.ReadStoreProperty(ValidFolderMaskSchema);
            if (mask.Status != PropertyReadStatus.Found || !TryReadMask(mask.Value, out int bits))
            {
                // A store that will not say which of its folders exist cannot be asked for one
                // without risking the one being made. Unreadable, never absent.
                return OutlookComSession.DefaultFolderResolution.Unreadable;
            }

            if ((bits & bit) == 0)
            {
                return OutlookComSession.DefaultFolderResolution.Absent;
            }

            // Proven to exist, so this call returns the folder that is there.
            return CallDefaultFolder(store, olDefaultFolderId, out folder);
        }

        private static OutlookComSession.DefaultFolderResolution ResolveByDesignation(
            ISpecialFolderStore store,
            int olDefaultFolderId,
            string schema,
            out object? folder,
            out SpecialFolderSource source)
        {
            folder = null;
            source = SpecialFolderSource.None;
            int? index = AdditionalRenIndex(olDefaultFolderId);

            OutlookComSession.DefaultFolderResolution primary =
                ResolveByValidMask(store, OlFolderInbox, FolderIpmInboxValid, out object? inbox);
            if (primary == OutlookComSession.DefaultFolderResolution.Resolved)
            {
                try
                {
                    primary = OpenDesignated(store, store.ReadFolderProperty(inbox!, schema), index, out folder);
                }
                finally
                {
                    store.Release(inbox);
                }

                if (primary == OutlookComSession.DefaultFolderResolution.Resolved)
                {
                    source = SpecialFolderSource.InboxDesignation;
                    return primary;
                }
            }

            // A store without an Inbox has nothing designated on one (primary stays Absent), and
            // one whose Inbox would not open is Unreadable. The single-entry-id designations are
            // read off the store object too, where MAPI's "Opening a message store folder" and
            // D39 looked; a hit there resolves, anything else leaves the Inbox's answer standing.
            if (!index.HasValue
                && OpenDesignated(store, store.ReadStoreProperty(schema), null, out folder)
                    == OutlookComSession.DefaultFolderResolution.Resolved)
            {
                source = SpecialFolderSource.StoreDesignation;
                return OutlookComSession.DefaultFolderResolution.Resolved;
            }

            return primary;
        }

        /// <summary>
        /// Turns one designation read into a folder. NotFound: absent. Failed or an
        /// unrecognisable value: unreadable. An entry id that no longer opens
        /// (<c>MAPI_E_NOT_FOUND</c>) names a folder that is gone, so that is absent too; any
        /// other failure to open it is unreadable.
        /// </summary>
        private static OutlookComSession.DefaultFolderResolution OpenDesignated(
            ISpecialFolderStore store,
            PropertyRead designation,
            int? index,
            out object? folder)
        {
            folder = null;
            if (designation.Status == PropertyReadStatus.NotFound)
            {
                return OutlookComSession.DefaultFolderResolution.Absent;
            }

            if (designation.Status != PropertyReadStatus.Found)
            {
                return OutlookComSession.DefaultFolderResolution.Unreadable;
            }

            DesignatedEntryId entryId = index.HasValue
                ? ReadEntryIdAt(designation.Value, index.Value)
                : ReadEntryId(designation.Value);
            if (entryId.Kind == DesignatedEntryIdKind.None)
            {
                return OutlookComSession.DefaultFolderResolution.Absent;
            }

            if (entryId.Kind != DesignatedEntryIdKind.EntryId)
            {
                return OutlookComSession.DefaultFolderResolution.Unreadable;
            }

            PropertyReadStatus opened = store.OpenFolder(entryId.Hex!, out object? candidate);
            if (opened == PropertyReadStatus.Found && candidate != null)
            {
                folder = candidate;
                return OutlookComSession.DefaultFolderResolution.Resolved;
            }

            store.Release(candidate);
            return opened == PropertyReadStatus.Failed
                ? OutlookComSession.DefaultFolderResolution.Unreadable
                : OutlookComSession.DefaultFolderResolution.Absent;
        }

        private static string ToHex(byte[] bytes)
        {
            StringBuilder sb = new StringBuilder(bytes.Length * 2);
            foreach (byte b in bytes)
            {
                sb.Append(b.ToString("X2", CultureInfo.InvariantCulture));
            }

            return sb.ToString();
        }
    }

    /// <summary>What a designation value held.</summary>
    public enum DesignatedEntryIdKind
    {
        /// <summary>No folder is designated (no value, an empty slot, or an index past the end).</summary>
        None = 0,

        /// <summary>One entry id, in <see cref="DesignatedEntryId.Hex"/>.</summary>
        EntryId = 1,

        /// <summary>A value of a shape the property cannot have - nothing may be concluded from it.</summary>
        Unrecognised = 2,
    }

    /// <summary>One designation value, read.</summary>
    public readonly struct DesignatedEntryId
    {
        private DesignatedEntryId(DesignatedEntryIdKind kind, string? hex)
        {
            Kind = kind;
            Hex = hex;
        }

        /// <summary>No folder designated.</summary>
        public static DesignatedEntryId None => new DesignatedEntryId(DesignatedEntryIdKind.None, null);

        /// <summary>An unrecognisable value.</summary>
        public static DesignatedEntryId Unrecognised => new DesignatedEntryId(DesignatedEntryIdKind.Unrecognised, null);

        /// <summary>What the value held.</summary>
        public DesignatedEntryIdKind Kind { get; }

        /// <summary>The entry id as upper-case hex, for <see cref="DesignatedEntryIdKind.EntryId"/>.</summary>
        public string? Hex { get; }

        /// <summary>One entry id.</summary>
        public static DesignatedEntryId Of(string hex) => new DesignatedEntryId(DesignatedEntryIdKind.EntryId, hex);
    }

    /// <summary>
    /// The COM half of <see cref="ISpecialFolderStore"/>, over one OOM <c>Store</c> and the
    /// <c>NameSpace</c> it belongs to. It answers questions and decides nothing. It does not
    /// own the store or the namespace; the caller releases both. STA-only, like every other
    /// COM call in this assembly.
    /// </summary>
    public sealed class ComSpecialFolderStore : ISpecialFolderStore
    {
        private readonly object _store;
        private readonly object _session;
        private readonly string? _storeId;
        private bool _exchangeStoreTypeRead;
        private int? _exchangeStoreType;

        /// <summary>Wraps one store. <paramref name="storeId"/> scopes <c>GetFolderFromID</c>; null or empty uses the one-argument form.</summary>
        public ComSpecialFolderStore(object store, object session, string? storeId)
        {
            _store = store ?? throw new ArgumentNullException(nameof(store));
            _session = session ?? throw new ArgumentNullException(nameof(session));
            _storeId = storeId;
        }

        /// <inheritdoc />
        public int? ExchangeStoreType
        {
            get
            {
                if (!_exchangeStoreTypeRead)
                {
                    _exchangeStoreTypeRead = true;
                    try
                    {
                        _exchangeStoreType = (int)((dynamic)_store).ExchangeStoreType;
                    }
                    catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                    {
                        _exchangeStoreType = null;
                    }
                }

                return _exchangeStoreType;
            }
        }

        /// <inheritdoc />
        public PropertyRead ReadStoreProperty(string schemaName)
        {
            return ReadProperty(_store, schemaName);
        }

        /// <inheritdoc />
        public PropertyRead ReadFolderProperty(object folder, string schemaName)
        {
            return ReadProperty(folder, schemaName);
        }

        /// <inheritdoc />
        public object? GetDefaultFolder(int olDefaultFolderId)
        {
            return ((dynamic)_store).GetDefaultFolder(olDefaultFolderId);
        }

        /// <inheritdoc />
        public PropertyReadStatus OpenFolder(string entryIdHex, out object? folder)
        {
            folder = null;
            try
            {
                dynamic ns = _session;
                folder = string.IsNullOrEmpty(_storeId)
                    ? (object?)ns.GetFolderFromID(entryIdHex)
                    : (object?)ns.GetFolderFromID(entryIdHex, _storeId);
                return folder == null ? PropertyReadStatus.NotFound : PropertyReadStatus.Found;
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                return SpecialFolders.IsMapiNotFound(ex) ? PropertyReadStatus.NotFound : PropertyReadStatus.Failed;
            }
        }

        /// <inheritdoc />
        public IReadOnlyList<string>? ListRootChildEntryIds()
        {
            object? root = null;
            object? folders = null;
            try
            {
                root = ((dynamic)_store).GetRootFolder();
                folders = ((dynamic)root!).Folders;
                int count = (int)((dynamic)folders!).Count;
                List<string> ids = new List<string>(count);
                for (int i = 1; i <= count; i++)
                {
                    object? child = null;
                    try
                    {
                        child = ((dynamic)folders!)[i];
                        string? id = (string?)((dynamic)child!).EntryID;
                        if (string.IsNullOrEmpty(id))
                        {
                            // A child that will not say who it is could be the very folder the
                            // snapshot exists to recognise. An incomplete list is no list.
                            return null;
                        }

                        ids.Add(id!);
                    }
                    finally
                    {
                        OutlookComSession.Release(child);
                    }
                }

                return ids;
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                return null;
            }
            finally
            {
                OutlookComSession.Release(folders);
                OutlookComSession.Release(root);
            }
        }

        /// <inheritdoc />
        public string? EntryIdOf(object folder)
        {
            try
            {
                return (string?)((dynamic)folder).EntryID;
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                return null;
            }
        }

        /// <inheritdoc />
        public SpecialFolderFacts Describe(object folder)
        {
            dynamic f = folder;
            string? entryId = EntryIdOf(folder);
            string? name = TryRead(() => (string?)f.Name);
            string? folderPath = TryRead(() => (string?)f.FolderPath);
            string? storeId = null;
            object? owner = null;
            try
            {
                owner = f.Store;
                storeId = owner == null ? null : (string?)((dynamic)owner).StoreID;
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                storeId = null;
            }
            finally
            {
                OutlookComSession.Release(owner);
            }

            int? itemType = null;
            try
            {
                itemType = (int)f.DefaultItemType;
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                itemType = null;
            }

            return new SpecialFolderFacts(entryId, name, folderPath, storeId, itemType);
        }

        /// <inheritdoc />
        public void Release(object? comObject)
        {
            OutlookComSession.Release(comObject);
        }

        private static PropertyRead ReadProperty(object owner, string schemaName)
        {
            object? accessor = null;
            try
            {
                accessor = ((dynamic)owner).PropertyAccessor;
                object? value = ((dynamic)accessor!).GetProperty(schemaName);
                return PropertyRead.Found(value);
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                return SpecialFolders.IsMapiNotFound(ex) ? PropertyRead.Missing() : PropertyRead.Failure();
            }
            finally
            {
                OutlookComSession.Release(accessor);
            }
        }

        private static string? TryRead(Func<string?> getter)
        {
            try
            {
                return getter();
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                return null;
            }
        }
    }
}
