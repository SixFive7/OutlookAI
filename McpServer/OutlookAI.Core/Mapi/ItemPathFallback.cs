using System;
using System.Collections.Generic;

namespace OutlookAI.Core.Mapi
{
    /// <summary>
    /// Fallback mapping input derived from <c>System.ItemPathDisplay</c> (v3.MD section 4):
    /// the display path has the shape <c>/StoreDisplayName/FolderA/.../ItemDisplayName</c>,
    /// so the folder path is recovered by stripping the first and last segments. Used when
    /// an EntryID decode fails verification: walk
    /// <c>Store.GetRootFolder().Folders(...)</c> down <see cref="FolderPath"/> and probe the
    /// folder with a narrow subject + received-time restriction.
    ///
    /// Known limitation: item display names containing '/' make the split ambiguous - the
    /// last segment is still taken as the item name, which matches how rarely folder names
    /// contain slashes compared to subjects.
    /// </summary>
    public sealed class ItemPathFallback
    {
        private ItemPathFallback(string storeDisplayName, IReadOnlyList<string> folderPath, string itemDisplayName)
        {
            StoreDisplayName = storeDisplayName;
            FolderPath = folderPath;
            ItemDisplayName = itemDisplayName;
        }

        /// <summary>First path segment: the store display name.</summary>
        public string StoreDisplayName { get; }

        /// <summary>Middle segments: folder names from store root down to the containing folder.</summary>
        public IReadOnlyList<string> FolderPath { get; }

        /// <summary>Last segment: the item display name (the subject, for mail).</summary>
        public string ItemDisplayName { get; }

        /// <summary>
        /// Derives the fallback path from a <c>System.ItemPathDisplay</c> value. Requires a
        /// leading '/' and at least a store segment plus an item segment.
        /// </summary>
        public static bool TryDerive(string? itemPathDisplay, out ItemPathFallback? fallback)
        {
            fallback = null;
            if (string.IsNullOrEmpty(itemPathDisplay) || itemPathDisplay![0] != '/')
            {
                return false;
            }

            string[] segments = itemPathDisplay.Substring(1).Split('/');
            if (segments.Length < 2 || segments[0].Length == 0)
            {
                return false;
            }

            List<string> folders = new List<string>();
            for (int i = 1; i < segments.Length - 1; i++)
            {
                folders.Add(segments[i]);
            }

            fallback = new ItemPathFallback(segments[0], folders, segments[segments.Length - 1]);
            return true;
        }
    }
}
