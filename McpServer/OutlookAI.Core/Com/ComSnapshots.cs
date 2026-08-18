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
            string? headers,
            string? htmlBody = null)
        {
            HtmlBody = htmlBody;
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

        /// <summary>
        /// Outlook's stored HTMLBody, transferred only when the caller asked for it
        /// (read include_html - batch B, B2). Null otherwise: it is several times the size
        /// of the text rendering and no other consumer needs it.
        /// </summary>
        public string? HtmlBody { get; }

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
            IReadOnlyList<ComRecipientInfo> recipients,
            string? conversationTopic = null,
            int? importance = null,
            bool readReceiptRequested = false)
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
            ConversationTopic = conversationTopic;
            Importance = importance;
            ReadReceiptRequested = readReceiptRequested;
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

        /// <summary>
        /// PR_CONVERSATION_TOPIC (0x0070001F) - the grouping key Outlook falls back to
        /// when a conversation cannot be resolved from the ConversationIndex GUID. Read
        /// through the PropertyAccessor because the object model exposes it read-only.
        /// </summary>
        public string? ConversationTopic { get; }

        /// <summary>OlImportance as reported by the item: 0 low, 1 normal, 2 high (null when unreadable).</summary>
        public int? Importance { get; }

        /// <summary>MailItem.ReadReceiptRequested.</summary>
        public bool ReadReceiptRequested { get; }
    }

    /// <summary>
    /// The agent-authored body of a draft, in exactly one of the two forms the tools
    /// accept (soak fix batch B - B1, v3.MD D45). Plain text is the default and unchanged
    /// path; HTML carries an ALREADY-NORMALIZED fragment (see
    /// <c>HtmlFragmentNormalizer</c>) that is inserted as real markup into the draft
    /// region only. Keeping the two apart in one object means the COM layer never has to
    /// guess whether a string is prose or markup.
    /// </summary>
    public sealed class ComDraftBody
    {
        private ComDraftBody(bool isHtml, string text, string html)
        {
            IsHtml = isHtml;
            Text = text;
            Html = html;
        }

        /// <summary>Plain-text body (line breaks preserved, everything escaped on the way in).</summary>
        public static ComDraftBody FromText(string text)
        {
            return new ComDraftBody(false, text ?? string.Empty, string.Empty);
        }

        /// <summary>HTML body - the fragment MUST already have passed <c>HtmlFragmentNormalizer</c>.</summary>
        public static ComDraftBody FromHtml(string normalizedHtml)
        {
            return new ComDraftBody(true, string.Empty, normalizedHtml ?? string.Empty);
        }

        /// <summary>True when the body is HTML and must be inserted as markup, not as characters.</summary>
        public bool IsHtml { get; }

        /// <summary>The plain text (empty for an HTML body).</summary>
        public string Text { get; }

        /// <summary>The normalized HTML fragment (empty for a plain-text body).</summary>
        public string Html { get; }

        /// <summary>"text" or "html" - reported on the outcome and in the audit line.</summary>
        public string FormatName => IsHtml ? "html" : "text";
    }

    /// <summary>
    /// Optional cross-tool draft settings shared by all four draft creators (soak fix
    /// batch A - A2/A3/A4). Everything here is additive: Cc/Bcc are APPENDED to whatever
    /// Outlook already put on the item (a reply-all's recipient list is never replaced),
    /// the subject override replaces the derived RE:/FW: subject while the original
    /// conversation topic is carried over, and importance / read-receipt are plain item
    /// properties.
    /// </summary>
    public sealed class ComDraftOptions
    {
        /// <summary>Creates a draft-options bundle (all parts optional).</summary>
        public ComDraftOptions(
            IReadOnlyList<string>? ccRecipients = null,
            IReadOnlyList<string>? bccRecipients = null,
            string? subjectOverride = null,
            int? importance = null,
            bool? requestReadReceipt = null,
            IReadOnlyList<string>? attachmentPaths = null)
        {
            CcRecipients = ccRecipients ?? Array.Empty<string>();
            BccRecipients = bccRecipients ?? Array.Empty<string>();
            SubjectOverride = subjectOverride;
            Importance = importance;
            RequestReadReceipt = requestReadReceipt;
            AttachmentPaths = attachmentPaths ?? Array.Empty<string>();
        }

        /// <summary>
        /// Absolute paths to attach, already existence/readability-validated PRE-COM by
        /// <c>DraftAttachments.Validate</c> (D46/C3) - the STA side only adds them.
        /// </summary>
        public IReadOnlyList<string> AttachmentPaths { get; }

        /// <summary>Addresses APPENDED as Cc.</summary>
        public IReadOnlyList<string> CcRecipients { get; }

        /// <summary>Addresses APPENDED as Bcc.</summary>
        public IReadOnlyList<string> BccRecipients { get; }

        /// <summary>Replacement subject (derived drafts; null keeps Outlook's RE:/FW: subject).</summary>
        public string? SubjectOverride { get; }

        /// <summary>OlImportance: 0 low, 1 normal, 2 high (null leaves Outlook's default).</summary>
        public int? Importance { get; }

        /// <summary>Whether to request a read receipt (null leaves the account default).</summary>
        public bool? RequestReadReceipt { get; }
    }

    /// <summary>
    /// Signature-override request for the draft creators (soak fix D37, R5 steering):
    /// replace the account-default signature with the named one via the WordEditor
    /// _MailAutoSig bookmark dance (the add-in's proven pattern).
    /// </summary>
    public sealed class ComSignatureOverride
    {
        /// <summary>Creates an override request.</summary>
        public ComSignatureOverride(string name, string filePath)
        {
            Name = name;
            FilePath = filePath;
        }

        /// <summary>Signature name (for reporting/audit).</summary>
        public string Name { get; }

        /// <summary>Signature file to insert (.htm preferred - Word converts and embeds its images natively).</summary>
        public string FilePath { get; }
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
            bool displayed,
            string? signatureOverrideName = null,
            bool signatureOverrideApplied = false,
            string? signatureOverrideError = null,
            bool bodyPlacedViaWordEditor = false,
            IReadOnlyList<string>? unresolvedRecipients = null,
            bool? conversationTopicPreserved = null,
            IReadOnlyList<ComAttachmentInfo>? attachments = null,
            bool composeSurfacePromoted = false,
            string? composeSurfaceError = null)
        {
            ComposeSurfacePromoted = composeSurfacePromoted;
            ComposeSurfaceError = composeSurfaceError;
            Attachments = attachments ?? Array.Empty<ComAttachmentInfo>();
            Draft = draft;
            AccountResolved = accountResolved;
            SignatureInjected = signatureInjected;
            BodyTextCharsBeforeSignature = bodyTextCharsBeforeSignature;
            BodyTextCharsAfterSignature = bodyTextCharsAfterSignature;
            MovedToDrafts = movedToDrafts;
            InitialSaveFolderName = initialSaveFolderName;
            Displayed = displayed;
            SignatureOverrideName = signatureOverrideName;
            SignatureOverrideApplied = signatureOverrideApplied;
            SignatureOverrideError = signatureOverrideError;
            BodyPlacedViaWordEditor = bodyPlacedViaWordEditor;
            UnresolvedRecipients = unresolvedRecipients ?? Array.Empty<string>();
            ConversationTopicPreserved = conversationTopicPreserved;
        }

        /// <summary>
        /// Attachments read back from the SAVED item (D46/C3) - a real round trip, not an
        /// echo of the requested paths, so what the agent is told is what Outlook holds.
        /// </summary>
        public IReadOnlyList<ComAttachmentInfo> Attachments { get; }

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

        /// <summary>Signature name the caller asked to apply instead of the default (null = no override requested).</summary>
        public string? SignatureOverrideName { get; }

        /// <summary>True when the override was applied (Word bookmark dance completed; on false the default-signature draft stands).</summary>
        public bool SignatureOverrideApplied { get; }

        /// <summary>Content-free failure reason when a requested override could not be applied.</summary>
        public string? SignatureOverrideError { get; }

        /// <summary>
        /// True when the agent body was placed through the held Inspector's WordEditor
        /// (the _MailAutoSig anchoring the add-in has proven), false when the composition
        /// fell back to the single wholesale HTMLBody assignment.
        /// </summary>
        public bool BodyPlacedViaWordEditor { get; }

        /// <summary>
        /// D49: true when the Word editor only became available because the compose surface
        /// was promoted invisibly (Outlook was window-less). Reported so headless
        /// composition is observable rather than merely assumed to work.
        /// </summary>
        public bool ComposeSurfacePromoted { get; }

        /// <summary>
        /// D49: content-free reason the Word compose surface could not be obtained, set
        /// WHENEVER the composition fell back to the HTMLBody splice - never gated on a
        /// signature override having been requested, because the fallback degrades the
        /// draft either way and a degradation the caller cannot see is the D48 defect.
        /// </summary>
        public string? ComposeSurfaceError { get; }

        /// <summary>Addresses Outlook could not resolve against the address book (never silently dropped).</summary>
        public IReadOnlyList<string> UnresolvedRecipients { get; }

        /// <summary>
        /// Only set when the caller overrode a derived draft's subject: whether the
        /// original conversation topic could be carried over to the draft.
        /// </summary>
        public bool? ConversationTopicPreserved { get; }
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
            IReadOnlyList<ComRecipientInfo> recipients,
            IReadOnlyList<ComAttachmentInfo>? attachments = null,
            string? bodyHtmlDigest = null)
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
            Attachments = attachments ?? Array.Empty<ComAttachmentInfo>();
            BodyHtmlDigest = bodyHtmlDigest;
        }

        /// <summary>
        /// Attachments currently on the draft - a content-hash input since D46, so a file
        /// added or removed after a confirm token was issued invalidates that token.
        /// </summary>
        public IReadOnlyList<ComAttachmentInfo> Attachments { get; }

        /// <summary>
        /// SHA-256 of the stored HTML body (content-hash input since D46: a markup-only
        /// edit leaves the plain text identical, so the text alone cannot see it). Null
        /// when the item has no readable HTML body.
        /// </summary>
        public string? BodyHtmlDigest { get; }

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

    /// <summary>
    /// A store's designated Archive folder (D39) - the folder Outlook's own Archive
    /// action (Backspace), mobile swipe-archive and OWA use. COM-free data.
    /// </summary>
    public sealed class ComArchiveFolderInfo
    {
        /// <summary>Creates an archive-folder snapshot.</summary>
        public ComArchiveFolderInfo(string storeDisplayName, string storeId, string entryId, string name, string storeRelativePath, string via)
        {
            StoreDisplayName = storeDisplayName;
            StoreId = storeId;
            EntryId = entryId;
            Name = name;
            StoreRelativePath = storeRelativePath;
            Via = via;
        }

        /// <summary>Store the archive folder belongs to.</summary>
        public string StoreDisplayName { get; }

        /// <summary>StoreID for direct folder/item opens.</summary>
        public string StoreId { get; }

        /// <summary>Archive folder EntryID.</summary>
        public string EntryId { get; }

        /// <summary>Folder display name (localized - e.g. "Archive" or "Archiveren"; resolved, never guessed).</summary>
        public string Name { get; }

        /// <summary>Store-relative folder path (list_folders convention).</summary>
        public string StoreRelativePath { get; }

        /// <summary>Resolution mechanism: "outlookDefaultFolder" (GetDefaultFolder 39) or "storeArchiveProperty" (PR_IPM_ARCHIVE_ENTRYID).</summary>
        public string Via { get; }
    }

    /// <summary>
    /// Result of one item move (D39). EntryIDs CHANGE on any move: <see cref="OldEntryId"/>
    /// stops resolving, <see cref="NewEntryId"/> is the item's new identity, and
    /// <see cref="FromFolderPath"/> is the undo address (move back = move
    /// NewEntryId to FromFolderPath). COM-free data.
    /// </summary>
    public sealed class ComMoveItemResult
    {
        /// <summary>Creates a move result.</summary>
        public ComMoveItemResult(
            string oldEntryId,
            string newEntryId,
            string storeDisplayName,
            string fromFolderPath,
            string toFolderPath,
            IReadOnlyList<string> createdFolderPaths)
        {
            OldEntryId = oldEntryId;
            NewEntryId = newEntryId;
            StoreDisplayName = storeDisplayName;
            FromFolderPath = fromFolderPath;
            ToFolderPath = toFolderPath;
            CreatedFolderPaths = createdFolderPaths;
        }

        /// <summary>EntryID before the move (stale from now on).</summary>
        public string OldEntryId { get; }

        /// <summary>EntryID after the move (the item's current identity).</summary>
        public string NewEntryId { get; }

        /// <summary>Store the item lives in (moves are same-store in v1).</summary>
        public string StoreDisplayName { get; }

        /// <summary>Store-relative path of the folder the item came FROM (undo address).</summary>
        public string FromFolderPath { get; }

        /// <summary>Store-relative path of the folder the item is in now.</summary>
        public string ToFolderPath { get; }

        /// <summary>Store-relative paths of folders created for this move (create_folder=true), outermost first.</summary>
        public IReadOnlyList<string> CreatedFolderPaths { get; }
    }

    /// <summary>Result of one gap sweep (COM-free data).</summary>
    public sealed class ComSweepResult
    {
        /// <summary>Creates a sweep result.</summary>
        public ComSweepResult(
            IReadOnlyList<ComMailBrief> items,
            int foldersSwept,
            int foldersSkipped,
            IReadOnlyList<string>? sweptFolders = null,
            int foldersFailed = 0,
            IReadOnlyList<string>? itemCappedFolders = null,
            bool folderCapReached = false,
            bool depthLimitReached = false,
            bool timeBudgetExceeded = false,
            int foldersAbsent = 0)
        {
            Items = items;
            FoldersSwept = foldersSwept;
            FoldersSkipped = foldersSkipped;
            SweptFolders = sweptFolders ?? Array.Empty<string>();
            FoldersFailed = foldersFailed;
            ItemCappedFolders = itemCappedFolders ?? Array.Empty<string>();
            FolderCapReached = folderCapReached;
            DepthLimitReached = depthLimitReached;
            TimeBudgetExceeded = timeBudgetExceeded;
            FoldersAbsent = foldersAbsent;
        }

        /// <summary>Items received/sent at or after the sweep start.</summary>
        public IReadOnlyList<ComMailBrief> Items { get; }

        /// <summary>Folders that were swept.</summary>
        public int FoldersSwept { get; }

        /// <summary>
        /// Folders that could not be resolved or enumerated (or fell past the folder cap) -
        /// every one of them a coverage hole. A default folder the store simply does not
        /// HAVE is not one of these; it is counted in <see cref="FoldersAbsent"/>.
        /// </summary>
        public int FoldersSkipped { get; }

        /// <summary>
        /// Default folders the stores in scope do not have at all (a data file with no Junk
        /// Email, say). Reported so the arithmetic stays checkable -
        /// <c>FoldersSwept + FoldersSkipped + FoldersAbsent</c> covers the whole folder set
        /// the sweep set out to walk - and NOT as a coverage hole: nothing can arrive in a
        /// folder that does not exist, so absence hides no mail. Always 0 for a
        /// folder-scoped sweep, which is asked for a NAMED folder and reports one that does
        /// not resolve as skipped.
        /// </summary>
        public int FoldersAbsent { get; }

        /// <summary>
        /// The swept folders as <c>store/store-relative path</c>, so a caller can report
        /// exactly what the freshness sweep covered (soak fix 13).
        /// </summary>
        public IReadOnlyList<string> SweptFolders { get; }

        /// <summary>
        /// Folders whose item enumeration FAILED through COM. Before soak fix 15 such a
        /// folder was counted as successfully swept with zero rows, which reported a hole
        /// in freshness coverage as full coverage.
        /// </summary>
        public int FoldersFailed { get; }

        /// <summary>
        /// Folders where the per-folder item cap cut the sweep short - the cap is applied
        /// newest-first, so the OLDEST items in the freshness window are the ones that
        /// vanish. Named so the caller can say which coverage is partial (section-12
        /// no-silent-caps discipline).
        /// </summary>
        public IReadOnlyList<string> ItemCappedFolders { get; }

        /// <summary>
        /// True when the scoped sweep's FOLDER cap stopped the subtree walk, so folders
        /// below the cut-off were never visited (and are therefore not even counted in
        /// <see cref="FoldersSkipped"/> beyond the first refusal).
        /// </summary>
        public bool FolderCapReached { get; }

        /// <summary>
        /// True when the scoped sweep's subtree walk refused a folder deeper than
        /// <c>OutlookComSession.FolderWalkDepthGuard</c>. A real tree is a handful of
        /// levels deep, so this means the tree is pathological (or cyclic) - and the
        /// guard is what keeps a recursive walk from taking the process down with an
        /// uncatchable StackOverflowException.
        /// </summary>
        public bool DepthLimitReached { get; }

        /// <summary>
        /// True when the scoped sweep's subtree walk ran out of
        /// <c>OutlookComSession.ScopedSweepTimeBudgetMs</c>, so the folders it had not
        /// reached yet were never swept. Reported for the same reason as
        /// <see cref="FolderCapReached"/>: a bound that stops the walk must never be
        /// invisible (section-12 no-silent-caps discipline).
        /// </summary>
        public bool TimeBudgetExceeded { get; }
    }

    /// <summary>
    /// Outcome of <c>update_draft</c> (v3.MD D46/C1): the re-snapshotted draft plus what
    /// the revision actually changed, all read back from the SAVED item rather than
    /// echoed from the request.
    /// </summary>
    public sealed class ComDraftUpdateResult
    {
        /// <summary>Creates the update outcome.</summary>
        public ComDraftUpdateResult(
            ComDraftInfo draft,
            IReadOnlyList<string> changedFields,
            IReadOnlyList<string> unresolvedRecipients,
            IReadOnlyList<ComAttachmentInfo> attachments,
            IReadOnlyList<string> attachmentsAdded,
            IReadOnlyList<string> attachmentsRemoved,
            IReadOnlyList<string> attachmentsFailed,
            bool bodyReplaced,
            bool bodyPlacedViaWordEditor,
            bool displayed,
            string? signatureOverrideName,
            bool signatureOverrideApplied,
            string? signatureOverrideError,
            bool? conversationTopicPreserved,
            int inlineImagesDropped = 0)
        {
            InlineImagesDropped = inlineImagesDropped;
            Draft = draft;
            ChangedFields = changedFields;
            UnresolvedRecipients = unresolvedRecipients;
            Attachments = attachments;
            AttachmentsAdded = attachmentsAdded;
            AttachmentsRemoved = attachmentsRemoved;
            AttachmentsFailed = attachmentsFailed;
            BodyReplaced = bodyReplaced;
            BodyPlacedViaWordEditor = bodyPlacedViaWordEditor;
            Displayed = displayed;
            SignatureOverrideName = signatureOverrideName;
            SignatureOverrideApplied = signatureOverrideApplied;
            SignatureOverrideError = signatureOverrideError;
            ConversationTopicPreserved = conversationTopicPreserved;
        }

        /// <summary>
        /// How many <c>&lt;img&gt;</c> elements the stored body LOST across this revision
        /// (D47). Non-zero only for a draft whose inline image was still a
        /// <c>file:///</c> LINK rather than an embedded <c>cid:</c> resource - Word cannot
        /// re-serialize such a link and replaces it with a placeholder shape. Never
        /// silent: the outcome reports it and names the remedy.
        /// </summary>
        public int InlineImagesDropped { get; }

        /// <summary>The draft as it stands AFTER the update.</summary>
        public ComDraftInfo Draft { get; }

        /// <summary>Names of the fields this call actually changed.</summary>
        public IReadOnlyList<string> ChangedFields { get; }

        /// <summary>Addresses Outlook could not resolve (they stay on the draft).</summary>
        public IReadOnlyList<string> UnresolvedRecipients { get; }

        /// <summary>Attachments on the SAVED draft after the update.</summary>
        public IReadOnlyList<ComAttachmentInfo> Attachments { get; }

        /// <summary>File names added by this call.</summary>
        public IReadOnlyList<string> AttachmentsAdded { get; }

        /// <summary>File names removed by this call.</summary>
        public IReadOnlyList<string> AttachmentsRemoved { get; }

        /// <summary>
        /// File names Outlook refused at <c>Attachments.Add</c> despite passing the
        /// pre-COM checks. Reported rather than thrown: the draft already exists, so a
        /// blanket failure would tell the caller nothing happened when something did.
        /// </summary>
        public IReadOnlyList<string> AttachmentsFailed { get; }

        /// <summary>True when the draft region was rewritten (a body was supplied).</summary>
        public bool BodyReplaced { get; }

        /// <summary>True when the body went in through the held-Inspector WordEditor.</summary>
        public bool BodyPlacedViaWordEditor { get; }

        /// <summary>True when the updated draft was (re)opened for the user.</summary>
        public bool Displayed { get; }

        /// <summary>Requested signature-override name, when one was requested.</summary>
        public string? SignatureOverrideName { get; }

        /// <summary>Whether the requested signature override was applied.</summary>
        public bool SignatureOverrideApplied { get; }

        /// <summary>Content-free reason a requested override failed.</summary>
        public string? SignatureOverrideError { get; }

        /// <summary>Only set when the subject was replaced: whether threading survived it (A3).</summary>
        public bool? ConversationTopicPreserved { get; }
    }

    /// <summary>
    /// Outcome of <c>discard_draft</c> (v3.MD D46/C2, S1 v3): the identity of the draft
    /// that was SOFT-deleted, its source folder, and - best effort - where it landed, so
    /// the operation stays reversible in the same way a move is.
    /// </summary>
    public sealed class ComDraftDiscardResult
    {
        /// <summary>Creates the discard outcome.</summary>
        public ComDraftDiscardResult(
            string oldEntryId,
            string? newEntryId,
            string? storeDisplayName,
            string? fromFolder,
            string? toFolder,
            string? subject)
        {
            OldEntryId = oldEntryId;
            NewEntryId = newEntryId;
            StoreDisplayName = storeDisplayName;
            FromFolder = fromFolder;
            ToFolder = toFolder;
            Subject = subject;
        }

        /// <summary>EntryID the draft had before the soft delete.</summary>
        public string OldEntryId { get; }

        /// <summary>EntryID in Deleted Items when it could be re-located (EntryIDs change on any move).</summary>
        public string? NewEntryId { get; }

        /// <summary>Store the draft lived in.</summary>
        public string? StoreDisplayName { get; }

        /// <summary>Folder the draft was discarded FROM (the undo address).</summary>
        public string? FromFolder { get; }

        /// <summary>Deleted Items folder name (localized) it was moved to.</summary>
        public string? ToFolder { get; }

        /// <summary>The discarded draft's subject (echoed so the agent can confirm what went).</summary>
        public string? Subject { get; }
    }
}
