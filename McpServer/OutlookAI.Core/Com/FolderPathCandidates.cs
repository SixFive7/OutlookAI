using System;
using System.Collections.Generic;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Every folder at a path of folder NAMES, where two folders of one name can sit side by side
    /// (stream A's finding on the Exchange guest, 2026-11-17: two top-level <c>Conversation Action
    /// Settings</c>). The search index keeps ONE path for both, so an index hit whose item sits in the
    /// second of them carries a path that <c>Folders[name]</c> answers with the first - and the hit could
    /// not be opened. <see cref="HitLocator"/>'s path fallback now tries every folder this finds (decided
    /// 2026-11-18, option (a) of the TODO item: "try every folder at that path, not the first").
    /// <para>
    /// Pure and generic over the folder type, so T1 pins the walk with a tree of plain objects while
    /// <see cref="OutlookComSession"/> walks Outlook's folders with it. Names are compared the way Outlook
    /// refuses a duplicate - ordinal, case-insensitive, nothing trimmed - and read from each child's own
    /// name, never through <c>Folders[name]</c>, whose answer for a name holding <c>/</c> is not the
    /// folder's (Q99).
    /// </para>
    /// </summary>
    public static class FolderPathCandidates
    {
        /// <summary>
        /// Most folders one walk keeps at any level, and so returns. Two namesakes are what was measured;
        /// this bounds a pathological tree, never an ordinary one.
        /// </summary>
        public const int MaxCandidates = 16;

        /// <summary>
        /// The folders at <paramref name="path"/> below <paramref name="root"/>, in the order their parents
        /// list them - the one <c>Folders[name]</c> would give first among them. Empty when a name is not
        /// there, and for an empty path (the root has no namesake). <paramref name="childrenOf"/> lists a
        /// folder's children with their names (a null name never matches); every folder it hands over that
        /// is not returned is given to <paramref name="release"/> exactly once - the root never, it stays the
        /// caller's, and so do the folders returned.
        /// </summary>
        public static IReadOnlyList<TFolder> FindAll<TFolder>(
            TFolder root,
            IReadOnlyList<string> path,
            Func<TFolder, IReadOnlyList<(string? Name, TFolder Folder)>> childrenOf,
            Action<TFolder> release,
            int cap = MaxCandidates)
            where TFolder : class
        {
            if (root == null)
            {
                throw new ArgumentNullException(nameof(root));
            }

            if (path == null)
            {
                throw new ArgumentNullException(nameof(path));
            }

            if (childrenOf == null)
            {
                throw new ArgumentNullException(nameof(childrenOf));
            }

            if (release == null)
            {
                throw new ArgumentNullException(nameof(release));
            }

            if (cap < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(cap));
            }

            if (path.Count == 0)
            {
                return Array.Empty<TFolder>();
            }

            List<TFolder> level = new List<TFolder> { root };
            foreach (string segment in path)
            {
                List<TFolder> next = new List<TFolder>();
                for (int p = 0; p < level.Count; p++)
                {
                    TFolder parent = level[p];
                    IReadOnlyList<(string? Name, TFolder Folder)> children;
                    try
                    {
                        children = childrenOf(parent);
                    }
                    catch
                    {
                        // Nothing this walk holds outlives a failure: the parents not listed yet and the
                        // folders kept for the next level go back too.
                        for (int rest = p + 1; rest < level.Count; rest++)
                        {
                            ReleaseUnlessRoot(level[rest], root, release);
                        }

                        foreach (TFolder kept in next)
                        {
                            release(kept);
                        }

                        throw;
                    }
                    finally
                    {
                        // A parent is needed only to list its children: every one of them but the
                        // caller's root is released as soon as it has been.
                        ReleaseUnlessRoot(parent, root, release);
                    }

                    foreach ((string? name, TFolder child) in children)
                    {
                        if (next.Count < cap && NameMatches(name, segment))
                        {
                            next.Add(child);
                        }
                        else
                        {
                            release(child);
                        }
                    }
                }

                level = next;
                if (level.Count == 0)
                {
                    break;
                }
            }

            return level;
        }

        private static void ReleaseUnlessRoot<TFolder>(TFolder folder, TFolder root, Action<TFolder> release)
            where TFolder : class
        {
            if (!ReferenceEquals(folder, root))
            {
                release(folder);
            }
        }

        /// <summary>Whether a folder's own name is the path's name: ordinal, case-insensitive, as Outlook refuses a duplicate. Pure.</summary>
        public static bool NameMatches(string? folderName, string? pathName)
        {
            return folderName != null && pathName != null
                && string.Equals(folderName, pathName, StringComparison.OrdinalIgnoreCase);
        }
    }
}
