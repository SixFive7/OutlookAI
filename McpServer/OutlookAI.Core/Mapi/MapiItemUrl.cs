using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OutlookAI.Core.Mapi
{
    /// <summary>
    /// Parser for Windows Search MAPI item URLs (v3.MD section 4). Official grammar:
    /// <c>mapi16://{SID}/StoreDisplayName($HashNumber)/StoreType/FolderA/.../EncodedEntryID[/at=EncodedAttachID:FileName]</c>
    /// where StoreType is 0 = default-type store, 1 = delegate, 2 = public-folder favorites.
    ///
    /// Attachment contents are indexed as separate entries whose URL appends
    /// <c>/at=&lt;encodedAttachId&gt;:&lt;filename&gt;</c> to the parent message URL - an
    /// attachment hit maps to its parent message by stripping that suffix.
    ///
    /// Pure logic: no COM, no I/O.
    /// </summary>
    public sealed class MapiItemUrl
    {
        private MapiItemUrl(
            string raw,
            string scheme,
            string sidSegment,
            string storeSegment,
            string storeDisplayName,
            string? storeUrlHash,
            int? storeType,
            IReadOnlyList<string> folderSegments,
            string? encodedItemSegment,
            bool isAttachment,
            string? encodedAttachmentId,
            string? attachmentFileName,
            string? parentItemUrl)
        {
            Raw = raw;
            Scheme = scheme;
            SidSegment = sidSegment;
            StoreSegment = storeSegment;
            StoreDisplayName = storeDisplayName;
            StoreUrlHash = storeUrlHash;
            StoreType = storeType;
            FolderSegments = folderSegments;
            EncodedItemSegment = encodedItemSegment;
            IsAttachment = isAttachment;
            EncodedAttachmentId = encodedAttachmentId;
            AttachmentFileName = attachmentFileName;
            ParentItemUrl = parentItemUrl;
        }

        /// <summary>The unmodified input URL.</summary>
        public string Raw { get; }

        /// <summary>URL scheme, e.g. "mapi16".</summary>
        public string Scheme { get; }

        /// <summary>The SID path segment including braces, e.g. "{S-1-5-21-...}".</summary>
        public string SidSegment { get; }

        /// <summary>The raw store segment, e.g. "alice@example.com($deadbeef)".</summary>
        public string StoreSegment { get; }

        /// <summary>Store display name with the "($hash)" suffix stripped and trimmed.</summary>
        public string StoreDisplayName { get; }

        /// <summary>Hex hash from the "($hash)" suffix of the store segment, without "$".</summary>
        public string? StoreUrlHash { get; }

        /// <summary>Store type segment: 0 default, 1 delegate, 2 public-folder favorites.</summary>
        public int? StoreType { get; }

        /// <summary>Folder path segments between the store-type segment and the item segment.</summary>
        public IReadOnlyList<string> FolderSegments { get; }

        /// <summary>The encoded (Hangul-range) EntryID segment, when the URL addresses an item.</summary>
        public string? EncodedItemSegment { get; }

        /// <summary>True when the URL addresses an indexed attachment entry (/at= suffix).</summary>
        public bool IsAttachment { get; }

        /// <summary>Encoded attachment id (between "/at=" and ":"), for attachment entries.</summary>
        public string? EncodedAttachmentId { get; }

        /// <summary>Attachment file name (after ":"), for attachment entries.</summary>
        public string? AttachmentFileName { get; }

        /// <summary>For attachment entries: the parent message URL (the /at= segment stripped).</summary>
        public string? ParentItemUrl { get; }

        /// <summary>
        /// Attempts to decode the item segment as a 24-byte message EntryID. For attachment
        /// entries this decodes the PARENT message id (the /at= suffix addresses the
        /// attachment within it).
        /// </summary>
        public bool TryDecodeEntryId(out DecodedEntryId? decoded)
        {
            return EntryIdCodec.TryDecodeMessageEntryId(EncodedItemSegment, out decoded);
        }

        /// <summary>
        /// Parses a Windows Search MAPI item URL. Returns false for null/empty input or
        /// URLs that do not follow the mapi scheme + SID + store grammar.
        /// </summary>
        public static bool TryParse(string? url, out MapiItemUrl? parsed)
        {
            parsed = null;
            if (string.IsNullOrWhiteSpace(url))
            {
                return false;
            }

            int schemeEnd = url!.IndexOf("://", StringComparison.Ordinal);
            if (schemeEnd <= 0)
            {
                return false;
            }

            string scheme = url.Substring(0, schemeEnd);
            if (!scheme.StartsWith("mapi", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string rest = url.Substring(schemeEnd + 3);
            string[] segments = rest.Split('/');
            if (segments.Length < 2 || segments[0].Length == 0 || segments[1].Length == 0)
            {
                return false;
            }

            string sidSegment = segments[0];
            string storeSegment = segments[1];
            SplitStoreSegment(storeSegment, out string storeDisplayName, out string? storeUrlHash);

            int? storeType = null;
            int folderStart = 2;
            if (segments.Length > 2 && int.TryParse(segments[2], NumberStyles.None, CultureInfo.InvariantCulture, out int typeValue))
            {
                storeType = typeValue;
                folderStart = 3;
            }

            // Attachment suffix: a dedicated trailing segment "at=<encodedId>:<filename>".
            bool isAttachment = false;
            string? encodedAttachmentId = null;
            string? attachmentFileName = null;
            string? parentItemUrl = null;
            int end = segments.Length;
            if (end > folderStart && segments[end - 1].StartsWith("at=", StringComparison.Ordinal))
            {
                string at = segments[end - 1].Substring(3);
                int colon = at.IndexOf(':');
                if (colon >= 0)
                {
                    encodedAttachmentId = at.Substring(0, colon);
                    attachmentFileName = at.Substring(colon + 1);
                }
                else
                {
                    encodedAttachmentId = at;
                }

                isAttachment = true;
                int cut = url.LastIndexOf("/at=", StringComparison.Ordinal);
                parentItemUrl = cut > 0 ? url.Substring(0, cut) : null;
                end--;
            }

            // The last remaining segment is the encoded item id when it decodes to the
            // 24-byte message layout; otherwise the URL addresses a folder (or something
            // this parser treats as folders - store display names never use the Hangul
            // encoding range on this profile).
            string? encodedItemSegment = null;
            if (end > folderStart)
            {
                string candidate = segments[end - 1];
                if (EntryIdCodec.TryDecodeMessageEntryId(candidate, out _))
                {
                    encodedItemSegment = candidate;
                    end--;
                }
                else if (isAttachment)
                {
                    // An /at= suffix without a decodable parent segment is malformed.
                    return false;
                }
            }
            else if (isAttachment)
            {
                return false;
            }

            List<string> folders = new List<string>();
            for (int i = folderStart; i < end; i++)
            {
                folders.Add(segments[i]);
            }

            parsed = new MapiItemUrl(
                url,
                scheme,
                sidSegment,
                storeSegment,
                storeDisplayName,
                storeUrlHash,
                storeType,
                folders,
                encodedItemSegment,
                isAttachment,
                encodedAttachmentId,
                attachmentFileName,
                parentItemUrl);
            return true;
        }

        private static void SplitStoreSegment(string storeSegment, out string displayName, out string? urlHash)
        {
            displayName = storeSegment;
            urlHash = null;

            // Trailing "($hex)" - the hash Windows Search appends to the display name.
            int open = storeSegment.LastIndexOf("($", StringComparison.Ordinal);
            if (open >= 0 && storeSegment.EndsWith(")", StringComparison.Ordinal))
            {
                string candidate = storeSegment.Substring(open + 2, storeSegment.Length - open - 3);
                if (IsHex(candidate))
                {
                    displayName = storeSegment.Substring(0, open).TrimEnd();
                    urlHash = candidate;
                }
            }
        }

        private static bool IsHex(string value)
        {
            if (value.Length == 0)
            {
                return false;
            }

            foreach (char c in value)
            {
                bool hex = (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
                if (!hex)
                {
                    return false;
                }
            }

            return true;
        }

        /// <summary>
        /// Scope prefix addressing the whole store:
        /// <c>scheme://{SID}/StoreSegment</c>. Use as the SCOPE predicate value.
        /// </summary>
        public string StorePrefix => Scheme + "://" + SidSegment + "/" + StoreSegment;

        /// <summary>
        /// Derives the <c>System.ItemFolderPathDisplay</c> value of a FOLDER-level mapi
        /// URL by pure string transform (v3.MD Phase-7 probe block, rule (f), verified on
        /// 25/25 folders across all five stores): <c>'/'</c> + the store segment with its
        /// <c>($hash)</c> suffix stripped + <c>'/'</c> + every segment after the
        /// store-type segment.
        /// <para>
        /// ⚠ FOOTGUN this rule exists to avoid: for a DELEGATE store the leading name is
        /// the HOST account, never the delegate's own <c>Store.DisplayName</c> - building
        /// the path from the display name returns 0 rows, silently, on 15/15 delegate
        /// folders. Deriving it from the URL is free and was never wrong; the rejected
        /// alternative (probing the value with <c>DIRECTORY=</c>) is unsound because
        /// folders can hold attachment-only rows a shallow traversal cannot see.
        /// </para>
        /// <para>
        /// Returns false for anything that is not a mapi URL. A store-root URL yields the
        /// store's own path (<c>/account</c>), a delegate root <c>/host/delegate</c>.
        /// </para>
        /// </summary>
        public static bool TryBuildFolderPathDisplay(string? folderScopeUrl, out string? displayPath)
        {
            displayPath = null;
            if (string.IsNullOrWhiteSpace(folderScopeUrl))
            {
                return false;
            }

            string url = folderScopeUrl!.Trim().TrimEnd('/');
            int schemeEnd = url.IndexOf("://", StringComparison.Ordinal);
            if (schemeEnd <= 0 || !url.Substring(0, schemeEnd).StartsWith("mapi", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            string[] segments = url.Substring(schemeEnd + 3).Split('/');
            if (segments.Length < 2 || segments[0].Length == 0 || segments[1].Length == 0)
            {
                return false;
            }

            SplitStoreSegment(segments[1], out string storeDisplayName, out _);

            // Segment 2 is the store-type digit (0 default, 1 delegate, 2 public folders)
            // when present; folder segments follow it. A URL without it addresses the
            // store root.
            int folderStart = 2;
            if (segments.Length > 2 && int.TryParse(segments[2], NumberStyles.None, CultureInfo.InvariantCulture, out _))
            {
                folderStart = 3;
            }

            StringBuilder path = new StringBuilder("/").Append(storeDisplayName);
            for (int i = folderStart; i < segments.Length; i++)
            {
                if (segments[i].Length == 0)
                {
                    continue;
                }

                path.Append('/').Append(segments[i]);
            }

            displayPath = path.ToString();
            return true;
        }
    }
}
