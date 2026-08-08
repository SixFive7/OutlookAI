using System;
using System.Collections.Generic;

namespace OutlookAI.Core.IndexSearch
{
    /// <summary>
    /// One compact SystemIndex hit (v3.MD section 8 payload discipline: subject, sender,
    /// date, folder, decoded EntryID, snippet - nothing bulkier). Full bodies are read
    /// through COM in a later layer; the index snippet (<see cref="AutoSummary"/>) is for
    /// triage only.
    /// </summary>
    public sealed class IndexHit
    {
        /// <summary>Raw mapi16:// item URL from System.ItemUrl.</summary>
        public string ItemUrl { get; internal set; } = string.Empty;

        /// <summary>System.Kind values (e.g. "email"; "document" for attachment-content entries).</summary>
        public IReadOnlyList<string> Kinds { get; internal set; } = Array.Empty<string>();

        /// <summary>System.Subject (null on attachment entries without one).</summary>
        public string? Subject { get; internal set; }

        /// <summary>System.Message.FromAddress.</summary>
        public string? FromAddress { get; internal set; }

        /// <summary>System.Message.FromName.</summary>
        public string? FromName { get; internal set; }

        /// <summary>System.Message.ToAddress values.</summary>
        public IReadOnlyList<string> ToAddresses { get; internal set; } = Array.Empty<string>();

        /// <summary>
        /// System.Message.DateReceived. Windows Search reports datetimes in UTC; the value
        /// carries DateTimeKind.Utc (verified against COM ReceivedTime in the live tests).
        /// </summary>
        public DateTime? DateReceivedUtc { get; internal set; }

        /// <summary>System.ItemPathDisplay: "/StoreDisplayName/Folder/.../ItemName".</summary>
        public string? ItemPathDisplay { get; internal set; }

        /// <summary>System.ItemNameDisplay (subject for mail; file name for attachment entries).</summary>
        public string? ItemNameDisplay { get; internal set; }

        /// <summary>System.Search.AutoSummary snippet (~1 KB), when populated.</summary>
        public string? AutoSummary { get; internal set; }

        /// <summary>System.Size in bytes.</summary>
        public long? SizeBytes { get; internal set; }

        /// <summary>System.IsRead.</summary>
        public bool? IsRead { get; internal set; }

        /// <summary>System.Message.HasAttachments.</summary>
        public bool? HasAttachments { get; internal set; }

        /// <summary>System.Message.ConversationID (string or hex of the raw value).</summary>
        public string? ConversationId { get; internal set; }

        // ---- Derived via EntryIdCodec / MapiItemUrl (v3.MD section 4) ----

        /// <summary>
        /// Decoded 48-hex-char MAPI EntryID ready for Namespace.GetItemFromID. For
        /// attachment hits this is the PARENT message's EntryID. Null when the URL tail
        /// did not decode - use <see cref="ItemPathDisplay"/> fallback mapping instead.
        /// </summary>
        public string? EntryIdHex { get; internal set; }

        /// <summary>32-hex-char store UID embedded in the EntryID (constant per store).</summary>
        public string? StoreUidHex { get; internal set; }

        /// <summary>Store display name parsed from the item URL.</summary>
        public string? StoreDisplayName { get; internal set; }

        /// <summary>Whole-store scope prefix parsed from the item URL.</summary>
        public string? StorePrefix { get; internal set; }

        /// <summary>Store type URL segment: 0 default, 1 delegate, 2 public-folder favorites.</summary>
        public int? StoreType { get; internal set; }

        /// <summary>Folder segments parsed from the item URL (store-relative).</summary>
        public IReadOnlyList<string> FolderSegments { get; internal set; } = Array.Empty<string>();

        /// <summary>True when this hit is an indexed attachment-content entry (/at= URL).</summary>
        public bool IsAttachmentHit { get; internal set; }

        /// <summary>Attachment file name, for attachment hits.</summary>
        public string? AttachmentFileName { get; internal set; }

        /// <summary>Parent message URL, for attachment hits.</summary>
        public string? ParentItemUrl { get; internal set; }

        internal IndexHit()
        {
        }
    }
}
