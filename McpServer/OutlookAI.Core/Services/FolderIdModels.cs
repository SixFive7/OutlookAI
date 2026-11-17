using System;
using System.Collections.Generic;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// find_folder's answer (Q136 (ii)): candidates, never an action. The caller picks one - asking the
    /// user when the store and path do not make it clear - and acts on it by its folderId.
    /// </summary>
    public sealed class FindFolderOutcome
    {
        /// <summary>The name asked for, trimmed.</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>The matching folders: exact name matches first, then path matches, then partial ones.</summary>
        public IReadOnlyList<FolderCandidateView> Candidates { get; set; } = Array.Empty<FolderCandidateView>();

        /// <summary>How many folders matched in all - more than the candidates listed when <see cref="Truncated"/>.</summary>
        public int Total { get; set; }

        /// <summary>True when more folders matched than are listed: narrow with store_id or a longer name.</summary>
        public bool Truncated { get; set; }

        /// <summary>True when the folder walk hit its cap, so a match beyond it is not listed.</summary>
        public bool? WalkCapReached { get; set; }

        /// <summary>True when the walk refused folders deeper than its depth guard.</summary>
        public bool? DepthLimitReached { get; set; }

        /// <summary>What to do next, present only when something needs saying.</summary>
        public IReadOnlyList<string>? Advice { get; set; }
    }

    /// <summary>One folder find_folder offers.</summary>
    public sealed class FolderCandidateView
    {
        /// <summary>The folder's ID (Q131) - what folder_id and parent_id take.</summary>
        public string FolderId { get; set; } = string.Empty;

        /// <summary>The folder's own name.</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>Display name of the folder's store - for reading only.</summary>
        public string Store { get; set; } = string.Empty;

        /// <summary>The store's ID (Q133).</summary>
        public string? StoreId { get; set; }

        /// <summary>The folder's names from the top of its store, joined by '/' - for reading only.</summary>
        public string Folder { get; set; } = string.Empty;

        /// <summary><c>exact</c> (the name is the text), <c>path</c> (the path ends with it) or <c>partial</c> (the name contains it).</summary>
        public string Match { get; set; } = string.Empty;

        /// <summary>Items in the folder, when Outlook reports it.</summary>
        public long? Items { get; set; }

        /// <summary>Unread items in the folder, when Outlook reports it.</summary>
        public long? Unread { get; set; }
    }

    /// <summary>create_folder's answer (Q134 (i)).</summary>
    public sealed class CreateFolderOutcome
    {
        /// <summary>Always true on success.</summary>
        public bool Created { get; set; }

        /// <summary>The new folder's ID (Q131): pass it as folder_id to move mail into it.</summary>
        public string FolderId { get; set; } = string.Empty;

        /// <summary>The new folder's name, as Outlook reports it back.</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>The new folder's names from the top of its store, joined by '/' - for reading only.</summary>
        public string Folder { get; set; } = string.Empty;

        /// <summary>Display name of the store it was made in.</summary>
        public string Store { get; set; } = string.Empty;

        /// <summary>The store's ID (Q133).</summary>
        public string? StoreId { get; set; }

        /// <summary>The parent's ID, as passed.</summary>
        public string ParentId { get; set; } = string.Empty;

        /// <summary>The parent's names joined by '/'; empty for the top of the store.</summary>
        public string ParentFolder { get; set; } = string.Empty;
    }
}
