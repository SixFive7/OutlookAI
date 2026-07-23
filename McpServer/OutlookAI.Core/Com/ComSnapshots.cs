using System;
using System.Collections.Generic;

namespace OutlookAI.Core.Com
{
    /// <summary>One recipient of a mail item (COM-free data).</summary>
    public sealed class ComRecipientInfo
    {
        /// <summary>Creates a recipient snapshot.</summary>
        public ComRecipientInfo(string kind, string? name, string? address)
        {
            Kind = kind;
            Name = name;
            Address = address;
        }

        /// <summary>"to", "cc" or "bcc".</summary>
        public string Kind { get; }

        /// <summary>Display name.</summary>
        public string? Name { get; }

        /// <summary>SMTP address when resolvable, otherwise the raw provider address.</summary>
        public string? Address { get; }
    }

    /// <summary>One attachment of a mail item (COM-free data).</summary>
    public sealed class ComAttachmentInfo
    {
        /// <summary>Creates an attachment snapshot.</summary>
        public ComAttachmentInfo(int index, string? fileName, long? sizeBytes)
        {
            Index = index;
            FileName = fileName;
            SizeBytes = sizeBytes;
        }

        /// <summary>1-based index into the item's Attachments collection.</summary>
        public int Index { get; }

        /// <summary>Attachment file name (null for attachment types without one).</summary>
        public string? FileName { get; }

        /// <summary>Size in bytes as reported by Outlook.</summary>
        public long? SizeBytes { get; }
    }

    /// <summary>
    /// Full detail snapshot of an opened item for the MCP <c>read</c> tool (v3.MD
    /// section 8 L2). All strings are plain data; the body arrives already capped at the
    /// requested maximum with the real total length reported separately.
    /// </summary>
    public sealed class ComItemDetail
    {
        /// <summary>Creates a detail snapshot (data only).</summary>
        public ComItemDetail(
            string entryId,
            string? storeDisplayName,
            string? folderPath,
            int? itemClass,
            string? subject,
            string? senderName,
            string? senderAddress,
            DateTime? receivedTime,
            DateTime? sentTime,
            IReadOnlyList<ComRecipientInfo> recipients,
            string body,
            long bodyTotalChars,
            string bodyOrigin,
            IReadOnlyList<ComAttachmentInfo> attachments,
            long? sizeBytes,
            bool? isRead,
            string? conversationId,
            string? internetMessageId,
            string? headers)
        {
            EntryId = entryId;
            StoreDisplayName = storeDisplayName;
            FolderPath = folderPath;
            ItemClass = itemClass;
            Subject = subject;
            SenderName = senderName;
            SenderAddress = senderAddress;
            ReceivedTime = receivedTime;
            SentTime = sentTime;
            Recipients = recipients;
            Body = body;
            BodyTotalChars = bodyTotalChars;
            BodyOrigin = bodyOrigin;
            Attachments = attachments;
            SizeBytes = sizeBytes;
            IsRead = isRead;
            ConversationId = conversationId;
            InternetMessageId = internetMessageId;
            Headers = headers;
        }

        /// <summary>Real (object-model) EntryID hex string.</summary>
        public string EntryId { get; }

        /// <summary>Containing store display name.</summary>
        public string? StoreDisplayName { get; }

        /// <summary>Folder path as reported by Outlook (\\Store\Folder\...).</summary>
        public string? FolderPath { get; }

        /// <summary>OlObjectClass (43 = olMail).</summary>
        public int? ItemClass { get; }

        /// <summary>Subject.</summary>
        public string? Subject { get; }

        /// <summary>Sender display name.</summary>
        public string? SenderName { get; }

        /// <summary>Sender SMTP address when resolvable.</summary>
        public string? SenderAddress { get; }

        /// <summary>ReceivedTime (local wall time as COM reports it).</summary>
        public DateTime? ReceivedTime { get; }

        /// <summary>SentOn (local wall time as COM reports it).</summary>
        public DateTime? SentTime { get; }

        /// <summary>To/Cc/Bcc recipients.</summary>
        public IReadOnlyList<ComRecipientInfo> Recipients { get; }

        /// <summary>Plain-text body, capped at the requested maximum length.</summary>
        public string Body { get; }

        /// <summary>Real total body length in characters before capping.</summary>
        public long BodyTotalChars { get; }

        /// <summary>"text" (native plain text), "html-converted", or "none".</summary>
        public string BodyOrigin { get; }

        /// <summary>Attachment list (indexes are stable for save_attachment).</summary>
        public IReadOnlyList<ComAttachmentInfo> Attachments { get; }

        /// <summary>Total item size in bytes.</summary>
        public long? SizeBytes { get; }

        /// <summary>Read state (true = read).</summary>
        public bool? IsRead { get; }

        /// <summary>Outlook ConversationID.</summary>
        public string? ConversationId { get; }

        /// <summary>Internet Message-ID header value (durable across moves).</summary>
        public string? InternetMessageId { get; }

        /// <summary>Raw transport headers when requested (may be capped by the caller).</summary>
        public string? Headers { get; }
    }

    /// <summary>One profile account (COM-free data).</summary>
    public sealed class ComAccountInfo
    {
        /// <summary>Creates an account snapshot.</summary>
        public ComAccountInfo(string? smtpAddress, string? displayName, string? deliveryStoreDisplayName)
        {
            SmtpAddress = smtpAddress;
            DisplayName = displayName;
            DeliveryStoreDisplayName = deliveryStoreDisplayName;
        }

        /// <summary>Account SMTP address.</summary>
        public string? SmtpAddress { get; }

        /// <summary>Account display name.</summary>
        public string? DisplayName { get; }

        /// <summary>Display name of the store new mail for this account lands in.</summary>
        public string? DeliveryStoreDisplayName { get; }
    }

    /// <summary>Detailed store description for list_accounts (COM-free data).</summary>
    public sealed class ComStoreDetail
    {
        /// <summary>Creates a store snapshot.</summary>
        public ComStoreDetail(string displayName, string storeId, int? exchangeStoreType, bool? isCachedExchange)
        {
            DisplayName = displayName;
            StoreId = storeId;
            ExchangeStoreType = exchangeStoreType;
            IsCachedExchange = isCachedExchange;
        }

        /// <summary>Store display name (matches the index URL store segment, Phase-1 fact).</summary>
        public string DisplayName { get; }

        /// <summary>StoreID for GetItemFromID.</summary>
        public string StoreId { get; }

        /// <summary>Raw OlExchangeStoreType value (0 primary, 1 additional/delegate, 2 public folders, 3 not Exchange).</summary>
        public int? ExchangeStoreType { get; }

        /// <summary>Store.IsCachedExchange - false means server-only (not locally indexed, D22/D25).</summary>
        public bool? IsCachedExchange { get; }
    }

    /// <summary>One folder in a store tree (COM-free data).</summary>
    public sealed class ComFolderInfo
    {
        /// <summary>Creates a folder snapshot.</summary>
        public ComFolderInfo(string storeDisplayName, string path, string name, long? itemCount, long? unreadCount, int childFolderCount)
        {
            StoreDisplayName = storeDisplayName;
            Path = path;
            Name = name;
            ItemCount = itemCount;
            UnreadCount = unreadCount;
            ChildFolderCount = childFolderCount;
        }

        /// <summary>Containing store display name.</summary>
        public string StoreDisplayName { get; }

        /// <summary>Store-relative path, segments joined with '/'.</summary>
        public string Path { get; }

        /// <summary>Folder display name.</summary>
        public string Name { get; }

        /// <summary>Total items (PR_CONTENT_COUNT) when available.</summary>
        public long? ItemCount { get; }

        /// <summary>Unread items (PR_CONTENT_UNREAD) when available.</summary>
        public long? UnreadCount { get; }

        /// <summary>Number of direct child folders.</summary>
        public int ChildFolderCount { get; }
    }

    /// <summary>
    /// Compact mail snapshot used by the fresh-mode gap sweep and the COM conversation
    /// walk (v3.MD D19). Carries the REAL EntryID, so reads on these hits skip the
    /// locate step entirely. Public constructor: the merge logic is unit-tested with
    /// fabricated instances.
    /// </summary>
    public sealed class ComMailBrief
    {
        /// <summary>Creates a brief snapshot (data only).</summary>
        public ComMailBrief(
            string entryId,
            string storeDisplayName,
            string? storeId,
            string? folderName,
            string? folderKind,
            string? subject,
            string? senderName,
            string? senderAddress,
            DateTime? receivedTime,
            bool? isRead,
            bool? hasAttachments,
            long? sizeBytes,
            string? body)
        {
            EntryId = entryId;
            StoreDisplayName = storeDisplayName;
            StoreId = storeId;
            FolderName = folderName;
            FolderKind = folderKind;
            Subject = subject;
            SenderName = senderName;
            SenderAddress = senderAddress;
            ReceivedTime = receivedTime;
            IsRead = isRead;
            HasAttachments = hasAttachments;
            SizeBytes = sizeBytes;
            Body = body;
        }

        /// <summary>Real (object-model) EntryID hex string.</summary>
        public string EntryId { get; }

        /// <summary>Store display name.</summary>
        public string StoreDisplayName { get; }

        /// <summary>StoreID for direct GetItemFromID opens (null when unknown).</summary>
        public string? StoreId { get; }

        /// <summary>Containing folder display name (localized, matches index URL segments).</summary>
        public string? FolderName { get; }

        /// <summary>"inbox" or "sent" for sweep results; null elsewhere.</summary>
        public string? FolderKind { get; }

        /// <summary>Subject.</summary>
        public string? Subject { get; }

        /// <summary>Sender display name.</summary>
        public string? SenderName { get; }

        /// <summary>Sender SMTP address when resolvable.</summary>
        public string? SenderAddress { get; }

        /// <summary>ReceivedTime (local wall time as COM reports it).</summary>
        public DateTime? ReceivedTime { get; }

        /// <summary>Read state.</summary>
        public bool? IsRead { get; }

        /// <summary>Whether the item has attachments.</summary>
        public bool? HasAttachments { get; }

        /// <summary>Item size in bytes.</summary>
        public long? SizeBytes { get; }

        /// <summary>Plain-text body - populated only when the sweep needs term matching.</summary>
        public string? Body { get; }
    }

    /// <summary>State snapshot of the active Outlook Explorer window (COM-free data).</summary>
    public sealed class ComExplorerState
    {
        /// <summary>Creates an explorer state snapshot.</summary>
        public ComExplorerState(string? caption, string? currentFolderPath, string? currentFolderName, int? windowState)
        {
            Caption = caption;
            CurrentFolderPath = currentFolderPath;
            CurrentFolderName = currentFolderName;
            WindowState = windowState;
        }

        /// <summary>Explorer window caption (also the Win32 window title).</summary>
        public string? Caption { get; }

        /// <summary>Absolute FolderPath of the current folder (\\Store\Folder\...).</summary>
        public string? CurrentFolderPath { get; }

        /// <summary>Display name of the current folder.</summary>
        public string? CurrentFolderName { get; }

        /// <summary>OlWindowState (0 maximized, 1 minimized, 2 normal) when readable.</summary>
        public int? WindowState { get; }
    }

    /// <summary>One open Inspector window (COM-free data).</summary>
    public sealed class ComInspectorInfo
    {
        /// <summary>Creates an inspector snapshot.</summary>
        public ComInspectorInfo(string? entryId, string? subject, int? itemClass)
        {
            EntryId = entryId;
            Subject = subject;
            ItemClass = itemClass;
        }

        /// <summary>EntryID of the item shown in the Inspector (null for unsaved items).</summary>
        public string? EntryId { get; }

        /// <summary>Subject of the shown item.</summary>
        public string? Subject { get; }

        /// <summary>OlObjectClass of the shown item (43 = olMail).</summary>
        public int? ItemClass { get; }
    }

    /// <summary>Result of one exhaustive folder/date-bounded COM scan (COM-free data).</summary>
    public sealed class ComExhaustiveResult
    {
        /// <summary>Creates an exhaustive scan result.</summary>
        public ComExhaustiveResult(
            IReadOnlyList<ComMailBrief> items,
            int foldersScanned,
            int foldersSkipped,
            string engine,
            bool instantSearchEnabled,
            bool truncated,
            bool timedOut)
        {
            Items = items;
            FoldersScanned = foldersScanned;
            FoldersSkipped = foldersSkipped;
            Engine = engine;
            InstantSearchEnabled = instantSearchEnabled;
            Truncated = truncated;
            TimedOut = timedOut;
        }

        /// <summary>Matched mail items with their REAL EntryIDs.</summary>
        public IReadOnlyList<ComMailBrief> Items { get; }

        /// <summary>Folders the scan filtered (mail folders in scope).</summary>
        public int FoldersScanned { get; }

        /// <summary>Folders where both filter engines failed.</summary>
        public int FoldersSkipped { get; }

        /// <summary>"ci_phrasematch", "like", or "ci_phrasematch+like" (per-folder downgrades).</summary>
        public string Engine { get; }

        /// <summary>Store.IsInstantSearchEnabled as reported by Outlook.</summary>
        public bool InstantSearchEnabled { get; }

        /// <summary>True when the result cap cut the scan short.</summary>
        public bool Truncated { get; }

        /// <summary>True when the time budget cut the scan short.</summary>
        public bool TimedOut { get; }
    }

    /// <summary>How a derived draft is created from its source mail (v3.MD section 3: threading ONLY via these).</summary>
    public enum ComDerivedDraftKind
    {
        /// <summary>MailItem.Reply() - answer the sender.</summary>
        Reply = 0,

        /// <summary>MailItem.ReplyAll() - answer sender + all recipients.</summary>
        ReplyAll = 1,

        /// <summary>MailItem.Forward() - forward to new recipients.</summary>
        Forward = 2,
    }

    /// <summary>
    /// Identity/threading snapshot of a mail item (COM-free data): where it lives, the
    /// account it would send as, and its conversation linkage. Used for draft results
    /// and for re-open verification of drafts.
    /// </summary>
    public sealed class ComDraftInfo
    {
        /// <summary>Creates a draft-info snapshot (data only).</summary>
        public ComDraftInfo(
            string entryId,
            string? storeDisplayName,
            string? storeId,
            string? parentFolderName,
            string? parentFolderEntryId,
            string? subject,
            string? sendUsingAccountSmtp,
            string? conversationIndex,
            string? conversationId,
            IReadOnlyList<ComRecipientInfo> recipients)
        {
            EntryId = entryId;
            StoreDisplayName = storeDisplayName;
            StoreId = storeId;
            ParentFolderName = parentFolderName;
            ParentFolderEntryId = parentFolderEntryId;
            Subject = subject;
            SendUsingAccountSmtp = sendUsingAccountSmtp;
            ConversationIndex = conversationIndex;
            ConversationId = conversationId;
            Recipients = recipients;
        }

        /// <summary>Real (object-model) EntryID hex string (changes when the item is moved).</summary>
        public string EntryId { get; }

        /// <summary>Containing store display name.</summary>
        public string? StoreDisplayName { get; }

        /// <summary>Containing store's StoreID.</summary>
        public string? StoreId { get; }

        /// <summary>Containing folder display name (localized, e.g. Drafts/Concepten).</summary>
        public string? ParentFolderName { get; }

        /// <summary>Containing folder EntryID (compare against a store's default Drafts folder).</summary>
        public string? ParentFolderEntryId { get; }

        /// <summary>Subject (RE:/FW: prefixes included for derived drafts).</summary>
        public string? Subject { get; }

        /// <summary>SmtpAddress of MailItem.SendUsingAccount (null when Outlook reports none).</summary>
        public string? SendUsingAccountSmtp { get; }

        /// <summary>PR_CONVERSATION_INDEX as the hex string the object model reports; a reply's index EXTENDS its parent's.</summary>
        public string? ConversationIndex { get; }

        /// <summary>Outlook ConversationID.</summary>
        public string? ConversationId { get; }

        /// <summary>To/Cc/Bcc recipients currently on the item.</summary>
        public IReadOnlyList<ComRecipientInfo> Recipients { get; }
    }

    /// <summary>Result of one draft-creation operation (COM-free data).</summary>
    public sealed class ComDraftCreateResult
    {
        /// <summary>Creates a draft-creation result.</summary>
        public ComDraftCreateResult(
            ComDraftInfo draft,
            bool accountResolved,
            bool signatureInjected,
            long bodyTextCharsBeforeSignature,
            long bodyTextCharsAfterSignature,
            bool movedToDrafts,
            string? initialSaveFolderName,
            bool displayed)
        {
            Draft = draft;
            AccountResolved = accountResolved;
            SignatureInjected = signatureInjected;
            BodyTextCharsBeforeSignature = bodyTextCharsBeforeSignature;
            BodyTextCharsAfterSignature = bodyTextCharsAfterSignature;
            MovedToDrafts = movedToDrafts;
            InitialSaveFolderName = initialSaveFolderName;
            Displayed = displayed;
        }

        /// <summary>The created draft (EntryID is final - post-move when a move happened).</summary>
        public ComDraftInfo Draft { get; }

        /// <summary>True when SendUsingAccount was pinned from a matched Account OBJECT (v3.MD section 3).</summary>
        public bool AccountResolved { get; }

        /// <summary>
        /// True when touching GetInspector grew the body's TEXT content - i.e. Outlook
        /// injected the account's signature (text-based: template HTML expansion without
        /// a signature adds markup but no text).
        /// </summary>
        public bool SignatureInjected { get; }

        /// <summary>Non-whitespace body text chars before the GetInspector touch.</summary>
        public long BodyTextCharsBeforeSignature { get; }

        /// <summary>Non-whitespace body text chars after the GetInspector touch.</summary>
        public long BodyTextCharsAfterSignature { get; }

        /// <summary>True when Save() landed the draft elsewhere and it was moved to the target Drafts folder.</summary>
        public bool MovedToDrafts { get; }

        /// <summary>Folder name Save() initially landed in (fact-finding; equals the final folder when no move happened).</summary>
        public string? InitialSaveFolderName { get; }

        /// <summary>True when the draft was opened in an Inspector window (D4 default).</summary>
        public bool Displayed { get; }
    }

    /// <summary>
    /// Sendable-state snapshot of a mail item for the high-friction send flow (Phase 5,
    /// v3.MD D4): where the item lives, whether it is still unsent, its recipients and
    /// plain-text body (hash input - never logged), and the account whose delivery
    /// store contains it (= the identity a send would use; null when no profile account
    /// delivers into that store, e.g. delegate caches).
    /// </summary>
    public sealed class ComSendableDraftState
    {
        /// <summary>Creates the snapshot (data only).</summary>
        public ComSendableDraftState(
            string entryId,
            string? storeId,
            string? storeDisplayName,
            string? parentFolderName,
            string? subject,
            bool isSent,
            string? bodyText,
            string? resolvedAccountSmtp,
            IReadOnlyList<ComRecipientInfo> recipients)
        {
            EntryId = entryId;
            StoreId = storeId;
            StoreDisplayName = storeDisplayName;
            ParentFolderName = parentFolderName;
            Subject = subject;
            IsSent = isSent;
            BodyText = bodyText;
            ResolvedAccountSmtp = resolvedAccountSmtp;
            Recipients = recipients;
        }

        /// <summary>Real EntryID as the object model reports it.</summary>
        public string EntryId { get; }

        /// <summary>Containing store's StoreID.</summary>
        public string? StoreId { get; }

        /// <summary>Containing store display name.</summary>
        public string? StoreDisplayName { get; }

        /// <summary>Containing folder display name (localized).</summary>
        public string? ParentFolderName { get; }

        /// <summary>Subject.</summary>
        public string? Subject { get; }

        /// <summary>MailItem.Sent - true means this is NOT a sendable draft.</summary>
        public bool IsSent { get; }

        /// <summary>Plain-text body (content-hash input; kept out of every log/output).</summary>
        public string? BodyText { get; }

        /// <summary>SmtpAddress of the account delivering into this store (the send identity), or null.</summary>
        public string? ResolvedAccountSmtp { get; }

        /// <summary>To/Cc/Bcc recipients currently on the item.</summary>
        public IReadOnlyList<ComRecipientInfo> Recipients { get; }
    }

    /// <summary>
    /// Result of one EXECUTED send (Phase 5). All fields are captured BEFORE
    /// <c>Send()</c> - a sent item gets a new EntryID in Sent Items, so
    /// <see cref="EntryIdAtSend"/> stops resolving after the transport accepts it.
    /// </summary>
    public sealed class ComSendResult
    {
        /// <summary>Creates the result (data only).</summary>
        public ComSendResult(
            string entryIdAtSend,
            string? storeDisplayName,
            string accountSmtp,
            string? sentOnBehalfOfName,
            string? subject,
            IReadOnlyList<ComRecipientInfo> recipients)
        {
            EntryIdAtSend = entryIdAtSend;
            StoreDisplayName = storeDisplayName;
            AccountSmtp = accountSmtp;
            SentOnBehalfOfName = sentOnBehalfOfName;
            Subject = subject;
            Recipients = recipients;
        }

        /// <summary>Draft EntryID at the moment Send() was invoked (invalid afterwards).</summary>
        public string EntryIdAtSend { get; }

        /// <summary>Store the draft lived in.</summary>
        public string? StoreDisplayName { get; }

        /// <summary>
        /// SmtpAddress the mail went out as - pinned via the PROPERTYPUTREF accessor and
        /// getter-verified IN-SESSION immediately before Send() (Phase-4 footgun: a
        /// dynamic assignment silently no-ops and the DEFAULT account would send).
        /// </summary>
        public string AccountSmtp { get; }

        /// <summary>SentOnBehalfOfName applied to the outgoing mail (send-as), when requested.</summary>
        public string? SentOnBehalfOfName { get; }

        /// <summary>Subject of the sent mail.</summary>
        public string? Subject { get; }

        /// <summary>Recipients the mail went to.</summary>
        public IReadOnlyList<ComRecipientInfo> Recipients { get; }
    }

    /// <summary>Identity of a store's default folder (COM-free data).</summary>
    public sealed class ComDefaultFolderInfo
    {
        /// <summary>Creates a default-folder snapshot.</summary>
        public ComDefaultFolderInfo(string entryId, string name)
        {
            EntryId = entryId;
            Name = name;
        }

        /// <summary>Folder EntryID.</summary>
        public string EntryId { get; }

        /// <summary>Folder display name (localized).</summary>
        public string Name { get; }
    }

    /// <summary>Result of one gap sweep (COM-free data).</summary>
    public sealed class ComSweepResult
    {
        /// <summary>Creates a sweep result.</summary>
        public ComSweepResult(IReadOnlyList<ComMailBrief> items, int foldersSwept, int foldersSkipped)
        {
            Items = items;
            FoldersSwept = foldersSwept;
            FoldersSkipped = foldersSkipped;
        }

        /// <summary>Items received/sent at or after the sweep start.</summary>
        public IReadOnlyList<ComMailBrief> Items { get; }

        /// <summary>Default folders that were swept.</summary>
        public int FoldersSwept { get; }

        /// <summary>Default folders that could not be resolved or enumerated.</summary>
        public int FoldersSkipped { get; }
    }
}
