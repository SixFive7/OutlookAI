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
