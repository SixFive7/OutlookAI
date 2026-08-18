using System;
using System.Collections.Generic;
using System.Globalization;

using OutlookAI.Core.Mapi;

namespace OutlookAI.Core.Services
{
    /// <summary>How a store + folder + include_subfolders request was turned into index predicates.</summary>
    public enum FolderScopeKind
    {
        /// <summary>No folder given: the whole store subtree (a bare recursive SCOPE).</summary>
        WholeStore = 0,

        /// <summary>Primary store, subfolders included: today's recursive SCOPE on the folder URL.</summary>
        PrimaryRecursive = 1,

        /// <summary>Primary store, subfolders excluded: folder SCOPE + one folder-path equality.</summary>
        PrimaryNonRecursive = 2,

        /// <summary>
        /// Delegate store: the delegate STORE ROOT scope plus folder-path equalities on
        /// FLAT leaf names - the only shape that reaches a delegate subfolder at all.
        /// </summary>
        DelegateFlat = 3,

        /// <summary>
        /// Delegate store, subfolders requested but the leaf set could not be narrowed
        /// (no folder tree available, or the set exceeded
        /// <see cref="FolderScopeResolver.DelegateFolderOrSetCap"/>): the query widened to
        /// the whole delegate store - a SUPERSET of what was asked, never a subset, and
        /// always reported.
        /// </summary>
        DelegateWidened = 4,

        /// <summary>
        /// The store exists in the PROFILE but the local index has no scope that addresses
        /// it, so there is nothing to query the index tier with. The freshness sweep covers
        /// the store on its own and the answer says so; see
        /// <see cref="FolderScopeResolver.ForUnindexedStore"/> for why this is a resolution
        /// rather than an error.
        /// </summary>
        StoreNotIndexed = 5,
    }

    /// <summary>
    /// The index predicates one store+folder request resolves to, plus everything the
    /// caller must tell the agent about (widening, leaf collisions). Data only.
    /// </summary>
    public sealed class FolderScopeResolution
    {
        internal FolderScopeResolution(
            FolderScopeKind kind,
            string? scope,
            string? storeScope,
            IReadOnlyList<string>? folderPaths,
            bool isDelegateStore,
            string? requestedFolder,
            IReadOnlyList<string>? collidingLeafNames,
            bool folderTreeUnavailable)
        {
            Kind = kind;
            Scope = scope;
            StoreScope = storeScope;
            FolderPaths = folderPaths;
            IsDelegateStore = isDelegateStore;
            RequestedFolder = requestedFolder;
            CollidingLeafNames = collidingLeafNames;
            FolderTreeUnavailable = folderTreeUnavailable;
        }

        /// <summary>
        /// The SCOPE predicate value, or null when the index cannot address this store at
        /// all (<see cref="FolderScopeKind.StoreNotIndexed"/>).
        /// <para>
        /// NULLABLE ON PURPOSE, and it is the whole point of that kind. Null here means
        /// "do not query the index tier"; it must never be read as "query it without a
        /// SCOPE", because an unscoped query answers a store-scoped request with the whole
        /// profile's mail. Callers branch on <see cref="IndexAddressable"/> rather than
        /// null-coalescing this to anything.
        /// </para>
        /// </summary>
        public string? Scope { get; }

        /// <summary>
        /// The whole-store SCOPE this folder lives under. The zero-row guard falls back to
        /// it: rows here but none under the folder bound means the FOLDER did not resolve,
        /// which is a different answer from "the folder holds no match". Null under
        /// <see cref="FolderScopeKind.StoreNotIndexed"/>, where there is no such scope and
        /// the guard does not run.
        /// </summary>
        public string? StoreScope { get; }

        /// <summary>Folder-path equalities ORed with the scope; null = recursive scope only.</summary>
        public IReadOnlyList<string>? FolderPaths { get; }

        /// <summary>Which shape was chosen.</summary>
        public FolderScopeKind Kind { get; }

        /// <summary>True when the scope addresses a delegate (<c>/1/</c>) store.</summary>
        public bool IsDelegateStore { get; }

        /// <summary>The store-relative folder path the caller asked for (null = whole store).</summary>
        public string? RequestedFolder { get; }

        /// <summary>
        /// Leaf names inside the requested delegate subtree that more than one COM folder
        /// shares. The flat delegate index namespace cannot separate them, so results may
        /// OVER-return - never silently (v3.MD constraint C3).
        /// </summary>
        public IReadOnlyList<string>? CollidingLeafNames { get; }

        /// <summary>True when the delegate folder tree could not be read (Outlook down).</summary>
        public bool FolderTreeUnavailable { get; }

        /// <summary>True when the query covers more than was asked and the caller must say so.</summary>
        public bool Widened => Kind == FolderScopeKind.DelegateWidened;

        /// <summary>
        /// Whether the index tier can be asked about this scope at all. False only for
        /// <see cref="FolderScopeKind.StoreNotIndexed"/>, where the caller must SKIP the
        /// index query rather than run it unscoped, and report that it did.
        /// </summary>
        public bool IndexAddressable => Kind != FolderScopeKind.StoreNotIndexed;
    }

    /// <summary>
    /// Turns a store scope + store-relative folder path + <c>include_subfolders</c> into
    /// the index-tier predicates. Pure logic - no COM, no index, fully unit-testable; the
    /// caller supplies the delegate store's COM folder paths when it has them.
    /// <para>
    /// ⚠ THE DELEGATE (<c>/1/</c>) INDEX NAMESPACE IS FLAT (measured 2026-07-27). A
    /// delegate folder is indexed as <c>&lt;host&gt;/1/&lt;delegate&gt;/&lt;LEAF NAME&gt;</c>
    /// with every intermediate COM folder dropped, and its
    /// <c>System.ItemFolderPathDisplay</c> is flattened the same way. Building a NESTED
    /// delegate URL - what this code used to do - addresses a folder that does not exist,
    /// so it returned 0 rows, silently, for every delegate subfolder (~3,871 items across
    /// 15 subfolders on this profile). The fix: scope to the delegate STORE ROOT and match
    /// the flat leaf path (verified exact against COM: 594/594 and 142/142).
    /// </para>
    /// <para>
    /// Primary stores keep full nesting in both the URL and the display path, so they use
    /// the folder URL as the scope and narrow with one equality when subfolders are
    /// excluded.
    /// </para>
    /// </summary>
    public static class FolderScopeResolver
    {
        /// <summary>
        /// Most folder-path equalities the delegate recursive shape will OR together
        /// before it gives up and widens to the whole delegate store.
        /// <para>
        /// MEASURED (read-only ADODB battery, 2026-07-27, delegate store root, agent-sized
        /// TOP 26 + ORDER BY, warm best-of-3): bare SCOPE 43 ms; OR-set x10 53 ms; x20
        /// 59 ms; x40 71 ms; x80 101 ms - and the provider FAILS OUTRIGHT
        /// ("Catastrophic failure", 0x8000FFFF) between 95 and 100 literals. 40 keeps the
        /// worst case around +28 ms while leaving a 2.4x margin to the hard ceiling; the
        /// real delegate mailboxes here index 11 and 23 distinct folder paths in total, so
        /// the cap is never reached in practice.
        /// </para>
        /// </summary>
        public const int DelegateFolderOrSetCap = 40;

        /// <summary>
        /// Resolves a PRIMARY (store-type 0) store scope. <paramref name="folder"/> null
        /// = whole store; otherwise the folder URL is the scope and, when
        /// <paramref name="includeSubfolders"/> is false, one folder-path equality narrows
        /// it to that folder alone (its attachment rows included - they inherit the
        /// parent's folder display path).
        /// </summary>
        public static FolderScopeResolution ForPrimaryStore(string storePrefix, string? folder, bool includeSubfolders)
        {
            if (string.IsNullOrWhiteSpace(storePrefix))
            {
                throw new ArgumentException("Store prefix is required.", nameof(storePrefix));
            }

            string? normalized = NormalizeFolder(folder);
            if (normalized == null)
            {
                return new FolderScopeResolution(
                    FolderScopeKind.WholeStore, storePrefix, storePrefix, null, false, null, null, false);
            }

            string scope = storePrefix + "/0/" + normalized;
            if (includeSubfolders)
            {
                return new FolderScopeResolution(
                    FolderScopeKind.PrimaryRecursive, scope, storePrefix, null, false, normalized, null, false);
            }

            if (!MapiItemUrl.TryBuildFolderPathDisplay(scope, out string? displayPath) || displayPath == null)
            {
                // Unreachable for a well-formed store prefix; degrade to the recursive
                // scope rather than dropping the folder bound altogether.
                return new FolderScopeResolution(
                    FolderScopeKind.PrimaryRecursive, scope, storePrefix, null, false, normalized, null, false);
            }

            return new FolderScopeResolution(
                FolderScopeKind.PrimaryNonRecursive, scope, storePrefix, new[] { displayPath },
                false, normalized, null, false);
        }

        /// <summary>
        /// Resolves a store the PROFILE has and the local index cannot address: a PST, an
        /// archive-only data file, a fresh install, a machine where Windows Search is off,
        /// excluded or still building. There is no scope and no folder bound, because the
        /// index holds nothing to bound.
        /// <para>
        /// ⚠ THIS EXISTS BECAUSE STORE SCOPE WAS RESOLVED AGAINST THE INDEX RATHER THAN
        /// AGAINST OUTLOOK (measured 2026-08-18, clean machine, single PST profile not in
        /// the Windows Search index). The store answered <c>list_folders</c>, answered an
        /// exhaustive search and was swept by an UNSCOPED search - which degraded honestly
        /// and said so - yet every non-exhaustive search naming it failed outright with
        /// "Store 'X' was not found in the local index. Known stores: ", an empty
        /// enumeration whose remedy pointed at <c>list_accounts</c>, which returns the very
        /// name that just failed. A whole feature dead on an ordinary configuration.
        /// </para>
        /// <para>
        /// It resolves rather than throwing because the question "which mail is in this
        /// store" has a true answer that one tier can still produce. The index contributes
        /// nothing, the freshness sweep covers the store, and the payload reports the hole
        /// through machinery that already existed for exactly this state
        /// (<c>no_index_frontier</c>, <c>sweep.storesWithoutIndex</c>, <c>degraded</c>) -
        /// the same state an unscoped search on the same profile has always reported.
        /// </para>
        /// <para>
        /// A store that is in NEITHER the index nor the profile still throws, from the
        /// caller: that is the case this cannot be reached for, and telling the two apart
        /// is why the caller asks Outlook before it gets here.
        /// </para>
        /// </summary>
        /// <param name="folder">
        /// The store-relative folder the caller asked for, kept so the payload can report
        /// what was requested. It bounds the SWEEP, which the caller drives from the request
        /// directly; it cannot bound an index query that is not being made, which is why no
        /// <c>include_subfolders</c> is taken here.
        /// </param>
        public static FolderScopeResolution ForUnindexedStore(string? folder)
        {
            return new FolderScopeResolution(
                FolderScopeKind.StoreNotIndexed,
                scope: null,
                storeScope: null,
                folderPaths: null,
                isDelegateStore: false,
                requestedFolder: NormalizeFolder(folder),
                collidingLeafNames: null,
                folderTreeUnavailable: false);
        }

        /// <summary>
        /// Resolves a DELEGATE (store-type 1) store scope. The scope is ALWAYS the
        /// delegate store root - nested delegate URLs do not exist in the index - and the
        /// folder bound is expressed as flat leaf-path equalities.
        /// </summary>
        /// <param name="delegateStoreScope">
        /// <c>mapi16://{SID}/&lt;host&gt;($hash)/1/&lt;delegate&gt;</c>.
        /// </param>
        /// <param name="folder">Store-relative COM folder path, or null for the whole delegate store.</param>
        /// <param name="includeSubfolders">Whether the COM subtree below <paramref name="folder"/> is in scope.</param>
        /// <param name="comFolderPaths">
        /// Every store-relative folder path of the delegate store as COM sees it (nested,
        /// '/'-separated). Null when Outlook could not be reached - the recursive request
        /// then widens rather than silently narrowing.
        /// </param>
        public static FolderScopeResolution ForDelegateStore(
            string delegateStoreScope,
            string? folder,
            bool includeSubfolders,
            IReadOnlyList<string>? comFolderPaths)
        {
            if (string.IsNullOrWhiteSpace(delegateStoreScope))
            {
                throw new ArgumentException("Delegate store scope is required.", nameof(delegateStoreScope));
            }

            string root = delegateStoreScope.TrimEnd('/');
            string? normalized = NormalizeFolder(folder);
            if (normalized == null)
            {
                // The whole delegate store: the root scope already covers every flat
                // folder, so no filter is needed and the flag cannot change the answer.
                return new FolderScopeResolution(
                    FolderScopeKind.WholeStore, root, root, null, true, null, null, comFolderPaths == null);
            }

            if (!MapiItemUrl.TryBuildFolderPathDisplay(root, out string? rootPath) || rootPath == null)
            {
                throw new ArgumentException(
                    "Delegate store scope is not a mapi URL: " + delegateStoreScope,
                    nameof(delegateStoreScope));
            }

            string leaf = LeafOf(normalized);
            if (!includeSubfolders)
            {
                return new FolderScopeResolution(
                    FolderScopeKind.DelegateFlat,
                    root,
                    root,
                    new[] { rootPath + "/" + leaf },
                    true,
                    normalized,
                    FindCollisions(comFolderPaths, new[] { leaf }),
                    comFolderPaths == null);
            }

            if (comFolderPaths == null)
            {
                // No folder tree, so the subtree's leaf names are unknown. Widening to the
                // delegate store root over-returns; narrowing to the single leaf would
                // silently drop the subfolders that were explicitly asked for.
                return new FolderScopeResolution(
                    FolderScopeKind.DelegateWidened, root, root, null, true, normalized, null, true);
            }

            List<string> leaves = SubtreeLeafNames(comFolderPaths, normalized, leaf);
            if (leaves.Count > DelegateFolderOrSetCap)
            {
                return new FolderScopeResolution(
                    FolderScopeKind.DelegateWidened, root, root, null, true, normalized, null, false);
            }

            List<string> paths = new List<string>(leaves.Count);
            foreach (string name in leaves)
            {
                paths.Add(rootPath + "/" + name);
            }

            return new FolderScopeResolution(
                FolderScopeKind.DelegateFlat,
                root,
                root,
                paths,
                true,
                normalized,
                FindCollisions(comFolderPaths, leaves),
                false);
        }

        /// <summary>
        /// Human-readable reason for a widened delegate scope, for the advice line
        /// (v3.MD constraint C2: an unhonorable narrowing is reported, never silent).
        /// </summary>
        public static string DescribeWidening(FolderScopeResolution resolution)
        {
            if (resolution == null)
            {
                throw new ArgumentNullException(nameof(resolution));
            }

            string reason = resolution.FolderTreeUnavailable
                ? "Outlook could not be reached to read the mailbox's folder tree"
                : "the subtree has more than " + DelegateFolderOrSetCap.ToString(CultureInfo.InvariantCulture)
                    + " folders, more than one index query can match at once";

            // Only a DelegateWidened resolution reaches here, and that kind always carries
            // the delegate root as its scope - so the fallback is unreachable rather than a
            // guess about what an unindexed store would be called.
            return "Delegate mailboxes are indexed WITHOUT their folder nesting, so subfolders of '"
                + resolution.RequestedFolder + "' can only be covered by listing them individually - and "
                + reason + ". The search was WIDENED to the whole '" + LeafOf(resolution.Scope ?? string.Empty)
                + "' mailbox instead of narrowing it silently: results may include mail from outside that folder. "
                + "Narrow with include_subfolders:false (that folder only), or search the subfolder directly.";
        }

        /// <summary>Advice text for a detected leaf-name collision (v3.MD constraint C3).</summary>
        public static string DescribeCollision(FolderScopeResolution resolution)
        {
            if (resolution == null)
            {
                throw new ArgumentNullException(nameof(resolution));
            }

            IReadOnlyList<string> names = resolution.CollidingLeafNames ?? Array.Empty<string>();
            return "Delegate mailboxes are indexed by folder NAME only (nesting is dropped), and this mailbox has "
                + "more than one folder named " + string.Join(", ", Quote(names))
                + " - the index cannot tell them apart, so these results can OVER-RETURN: they may include mail "
                + "from the same-named folder elsewhere in the mailbox. Use exhaustive:true to scan the exact "
                + "folder through Outlook.";
        }

        /// <summary>Advice text for the non-silent zero-row guard (v3.MD constraint C7).</summary>
        public static string DescribeUnresolvedFolder(string? folder, string store)
        {
            return "No results, and the folder path '" + folder + "' matched NOTHING in the index for store '"
                + store + "' although the store itself has indexed mail - so this looks like a folder-resolution "
                + "problem rather than an empty folder (a folder created minutes ago can also still be missing from "
                + "the index). Check the path with list_folders (delegate mailboxes are indexed by folder NAME, and "
                + "Outlook may show a localized name the index does not use), or retry with exhaustive:true to scan "
                + "the folder through Outlook directly.";
        }

        private static IEnumerable<string> Quote(IReadOnlyList<string> names)
        {
            foreach (string name in names)
            {
                yield return "'" + name + "'";
            }
        }

        /// <summary>
        /// Leaf names of the requested folder and every COM descendant, de-duplicated
        /// case-insensitively (two descendants sharing a name collapse to one equality -
        /// the index cannot separate them anyway).
        /// </summary>
        private static List<string> SubtreeLeafNames(
            IReadOnlyList<string> comFolderPaths, string requestedFolder, string requestedLeaf)
        {
            List<string> ordered = new List<string> { requestedLeaf };
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { requestedLeaf };
            string prefix = requestedFolder + "/";
            foreach (string path in comFolderPaths)
            {
                if (path == null || !path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string leaf = LeafOf(path);
                if (leaf.Length > 0 && seen.Add(leaf))
                {
                    ordered.Add(leaf);
                }
            }

            return ordered;
        }

        /// <summary>
        /// Leaf names in <paramref name="selectedLeaves"/> that more than one COM folder
        /// of the delegate store carries - the flat namespace merges them.
        /// </summary>
        private static IReadOnlyList<string>? FindCollisions(
            IReadOnlyList<string>? comFolderPaths, IReadOnlyList<string> selectedLeaves)
        {
            if (comFolderPaths == null || comFolderPaths.Count == 0)
            {
                return null;
            }

            Dictionary<string, int> counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (string path in comFolderPaths)
            {
                if (string.IsNullOrEmpty(path))
                {
                    continue;
                }

                string leaf = LeafOf(path);
                counts[leaf] = counts.TryGetValue(leaf, out int n) ? n + 1 : 1;
            }

            List<string>? collisions = null;
            HashSet<string> reported = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string leaf in selectedLeaves)
            {
                if (counts.TryGetValue(leaf, out int n) && n > 1 && reported.Add(leaf))
                {
                    (collisions ??= new List<string>()).Add(leaf);
                }
            }

            return collisions;
        }

        private static string LeafOf(string path)
        {
            int slash = path.LastIndexOf('/');
            return slash < 0 ? path : path.Substring(slash + 1);
        }

        private static string? NormalizeFolder(string? folder)
        {
            if (folder == null)
            {
                return null;
            }

            string trimmed = folder.Trim().Trim('/');
            return trimmed.Length == 0 ? null : trimmed;
        }
    }
}
