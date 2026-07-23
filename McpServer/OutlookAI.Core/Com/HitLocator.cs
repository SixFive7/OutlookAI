using System;
using System.Collections.Generic;

using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;

namespace OutlookAI.Core.Com
{
    /// <summary>Which mapping path located an index hit as an openable Outlook item.</summary>
    public enum HitLocationTier
    {
        /// <summary>Not located.</summary>
        Failed = 0,

        /// <summary>Located via the item URL's store/type/folder segments (primary path).</summary>
        UrlSegments = 1,

        /// <summary>Located via the System.ItemPathDisplay-derived path (fallback path).</summary>
        ItemPathDisplay = 2,
    }

    /// <summary>Outcome of locating one index hit through COM.</summary>
    public sealed class HitLocationResult
    {
        internal HitLocationResult(HitLocationTier tier, ComOpenResult? located, string? storeDisplayName, string? error)
        {
            Tier = tier;
            Located = located;
            StoreDisplayName = storeDisplayName;
            Error = error;
        }

        /// <summary>Mapping path that succeeded (or Failed).</summary>
        public HitLocationTier Tier { get; }

        /// <summary>Snapshot of the located item; its EntryId is the REAL long-form COM EntryID.</summary>
        public ComOpenResult? Located { get; }

        /// <summary>COM store display name the item was located in.</summary>
        public string? StoreDisplayName { get; }

        /// <summary>Content-free error description when location failed.</summary>
        public string? Error { get; }
    }

    /// <summary>
    /// Maps a SystemIndex hit to an openable Outlook item. PHASE-1 DISCOVERY (this
    /// machine, all stores cached Exchange incl. delegates): the 24-byte id decoded from
    /// the index URL is the OST-internal short id, while the object model exposes 70-byte
    /// Exchange-style EntryIDs (same 16-byte store UID at bytes 4..19, then folder/message
    /// database GUIDs + counters) - Namespace.GetItemFromID rejects the short form with
    /// 0x80040107 (MAPI_E_INVALID_ENTRYID) in every store. The decode therefore
    /// identifies the STORE (UID match confirmed) but items are located via the narrow
    /// folder probe the v3.MD section-4 fallback prescribes: walk the folder path from
    /// the URL segments (or from ItemPathDisplay) and restrict on Subject plus a
    /// ReceivedTime tolerance. Returned items carry the real EntryID for direct
    /// GetItemFromID use afterwards.
    ///
    /// Delegate-store hits (URL store-type segment /1/) live in the OWNER account's URL
    /// subtree, but the item resides in the delegate's own COM store whose display name
    /// is the first folder segment.
    /// </summary>
    public static class HitLocator
    {
        /// <summary>Locates a hit; tries the URL-segment path first, then ItemPathDisplay.</summary>
        public static HitLocationResult Locate(OutlookComSession session, IndexHit hit, int toleranceSeconds = 120)
        {
            if (session == null)
            {
                throw new ArgumentNullException(nameof(session));
            }

            if (hit == null)
            {
                throw new ArgumentNullException(nameof(hit));
            }

            string lookupSubject = DeriveLookupSubject(hit);

            string? urlError = null;
            if (TryMapUrlTarget(hit, out string? storeName, out IReadOnlyList<string>? folders))
            {
                ComOpenResult? located = session.TryResolveByPath(
                    storeName!,
                    folders!,
                    lookupSubject,
                    hit.DateReceivedUtc,
                    out urlError,
                    toleranceSeconds);
                if (located != null)
                {
                    return new HitLocationResult(HitLocationTier.UrlSegments, located, storeName, null);
                }
            }
            else
            {
                urlError = "UrlNotParsable";
            }

            if (ItemPathFallback.TryDerive(hit.ItemPathDisplay, out ItemPathFallback? fallback) && fallback != null)
            {
                ComOpenResult? located = session.TryResolveByPath(
                    fallback.StoreDisplayName,
                    fallback.FolderPath,
                    lookupSubject,
                    hit.DateReceivedUtc,
                    out string? fallbackError,
                    toleranceSeconds);
                if (located != null)
                {
                    return new HitLocationResult(HitLocationTier.ItemPathDisplay, located, fallback.StoreDisplayName, null);
                }

                return new HitLocationResult(
                    HitLocationTier.Failed, null, null, $"url:{urlError} fallback:{fallbackError}");
            }

            return new HitLocationResult(HitLocationTier.Failed, null, null, $"url:{urlError} fallback:NoItemPathDisplay");
        }

        /// <summary>
        /// Derives the COM store + folder path from the URL segments, applying the
        /// delegate rule (store-type 1: first folder segment is the delegate store).
        /// </summary>
        public static bool TryMapUrlTarget(IndexHit hit, out string? storeDisplayName, out IReadOnlyList<string>? folderPath)
        {
            storeDisplayName = null;
            folderPath = null;
            if (hit.StorePrefix == null || hit.StoreDisplayName == null)
            {
                return false;
            }

            if (hit.StoreType == 1)
            {
                if (hit.FolderSegments.Count == 0)
                {
                    return false;
                }

                storeDisplayName = hit.FolderSegments[0];
                List<string> rest = new List<string>();
                for (int i = 1; i < hit.FolderSegments.Count; i++)
                {
                    rest.Add(hit.FolderSegments[i]);
                }

                folderPath = rest;
                return true;
            }

            storeDisplayName = hit.StoreDisplayName;
            folderPath = hit.FolderSegments;
            return true;
        }

        /// <summary>
        /// The subject to probe folders with. For message hits this is System.Subject.
        /// Attachment (document) entries carry COMBINED display strings (probed live):
        /// System.Subject = "&lt;filename&gt; (&lt;parent subject&gt;)" and the
        /// ItemPathDisplay tail = "&lt;parent subject&gt; : &lt;filename&gt;" - the parent
        /// subject is recovered by stripping those decorations.
        /// </summary>
        public static string DeriveLookupSubject(IndexHit hit)
        {
            if (hit == null)
            {
                throw new ArgumentNullException(nameof(hit));
            }

            if (hit.IsAttachmentHit && !string.IsNullOrEmpty(hit.AttachmentFileName))
            {
                string fileName = hit.AttachmentFileName!;

                if (ItemPathFallback.TryDerive(hit.ItemPathDisplay, out ItemPathFallback? pathFallback)
                    && pathFallback != null)
                {
                    string suffix = " : " + fileName;
                    string tail = pathFallback.ItemDisplayName;
                    if (tail.Length > suffix.Length && tail.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
                    {
                        return tail.Substring(0, tail.Length - suffix.Length);
                    }
                }

                if (hit.Subject != null)
                {
                    string prefix = fileName + " (";
                    if (hit.Subject.Length > prefix.Length + 1
                        && hit.Subject.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
                        && hit.Subject.EndsWith(")", StringComparison.Ordinal))
                    {
                        return hit.Subject.Substring(prefix.Length, hit.Subject.Length - prefix.Length - 1);
                    }
                }
            }

            if (hit.Subject != null)
            {
                return hit.Subject;
            }

            if (ItemPathFallback.TryDerive(hit.ItemPathDisplay, out ItemPathFallback? fallback) && fallback != null)
            {
                return fallback.ItemDisplayName;
            }

            return string.Empty;
        }
    }
}
