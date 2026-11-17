using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Mapi;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// Folders and stores addressed by ID (Q131-Q136, decided by the maintainer 2026-10-11): every
    /// tool that takes a folder takes its folder ID, every tool that takes a store its store ID, and
    /// names are for display only (<see cref="OutlookIds"/>). This half of the service resolves those
    /// IDs - against Outlook, or against the search index when Outlook is not running (Q132) - and
    /// holds the tools that exist only because of them: find_folder (Q136 (ii)) and create_folder
    /// (Q134 (i)), and move_mail, goto_folder and show_search_results, which now take IDs.
    /// <para>
    /// No learned maps, ever (Q132 (i)): nothing here remembers which folder an ID was at. An ID is
    /// opened with <c>GetFolderFromID</c>, or looked up in the index by the node id it carries, every
    /// time. The one list kept is the one there always was - Outlook's list of its stores, cached
    /// for <see cref="StoreDetailsCacheTtl"/> - which says what the stores ARE, not where a folder is.
    /// </para>
    /// </summary>
    public sealed partial class MailService
    {
        /// <summary>Most candidates one find_folder call returns; the total says how many matched.</summary>
        public const int FindFolderCandidatesCap = 50;

        /// <summary>Longest folder name create_folder accepts (Outlook's own limit is near it).</summary>
        public const int FolderNameCharsCap = 255;

        /// <summary>How a search found its folder in the index (<see cref="SearchScopeInfo.FolderFoundVia"/>).</summary>
        internal const string FolderFoundViaIndex = "index";

        /// <summary>See <see cref="FolderFoundViaIndex"/>.</summary>
        internal const string FolderFoundViaOutlook = "outlook";

        // ------------------------------------------------------------------ store IDs

        /// <summary>
        /// A store's ID as every tool shows it: <c>S</c> and the provider UID its root folder's id carries,
        /// or - for a store whose root folder was not read (an Exchange store Outlook does not cache) -
        /// <c>T</c> and its own entry id (<see cref="OutlookIds.StoreIdOf"/>). Null when neither reads.
        /// </summary>
        public static string? StoreIdOf(ComStoreDetail detail)
        {
            if (detail == null)
            {
                throw new ArgumentNullException(nameof(detail));
            }

            return OutlookIds.StoreIdOf(detail.RootFolderEntryId, detail.StoreId);
        }

        /// <summary>
        /// Whether a store answers to a decoded store ID: an <c>S</c> ID by the provider UID of the store's
        /// root folder, a <c>T</c> ID by the store's own entry id, byte for byte. Pure.
        /// </summary>
        public static bool StoreMatchesId(ComStoreDetail detail, StoreIdParts parts)
        {
            if (detail == null)
            {
                throw new ArgumentNullException(nameof(detail));
            }

            if (parts == null)
            {
                throw new ArgumentNullException(nameof(parts));
            }

            if (parts.ProviderUidHex != null)
            {
                return string.Equals(
                    OutlookIds.ProviderUidHexOf(detail.RootFolderEntryId), parts.ProviderUidHex, StringComparison.OrdinalIgnoreCase);
            }

            return StoreHash.TryDecodeHex(detail.StoreId, out byte[] own)
                && StoreHash.TryDecodeHex(parts.StoreEntryIdHex, out byte[] asked)
                && own.Length == asked.Length
                && own.SequenceEqual(asked);
        }

        /// <summary>
        /// The store a store ID names, among <paramref name="stores"/> - or the refusal that says why
        /// none can be picked: no store answers to it, or several do (a mailbox seen through two
        /// stores shares one provider UID). Pure.
        /// </summary>
        public static ComStoreDetail PickStore(IReadOnlyList<ComStoreDetail> stores, StoreIdParts parts, string parameterName)
        {
            if (stores == null)
            {
                throw new ArgumentNullException(nameof(stores));
            }

            List<ComStoreDetail> matches = stores.Where(d => StoreMatchesId(d, parts)).ToList();
            if (matches.Count == 1)
            {
                return matches[0];
            }

            if (matches.Count == 0)
            {
                throw new ArgumentException(
                    "No store in this Outlook profile has the ID '" + parts.Id + "' (" + parameterName + "). It may come "
                    + "from another profile or from a store removed since. " + DescribeStoreList(stores)
                    + " list_accounts lists every store with its storeId.",
                    parameterName);
            }

            throw new ArgumentException(
                matches.Count.ToString(CultureInfo.InvariantCulture) + " stores of this Outlook profile answer to the ID '"
                + parts.Id + "': " + string.Join(", ", matches.Select(m => "'" + m.DisplayName + "'"))
                + " - one mailbox seen through two stores carries one identity, so the ID cannot say which is meant and "
                + "nothing is picked at random. Address a folder of it by its folderId instead (find_folder, list_folders).",
                parameterName);
        }

        /// <summary>"The stores here are: 'A' (S...), 'B' (S...)." - capped, for a refusal.</summary>
        private static string DescribeStoreList(IReadOnlyList<ComStoreDetail> stores)
        {
            const int shown = 10;
            if (stores.Count == 0)
            {
                return "Outlook lists no store at all.";
            }

            string list = string.Join(
                ", ", stores.Take(shown).Select(d => "'" + d.DisplayName + "' (" + (StoreIdOf(d) ?? "no ID") + ")"));
            return "The stores here are: " + list
                + (stores.Count > shown ? " and " + (stores.Count - shown).ToString(CultureInfo.InvariantCulture) + " more." : ".");
        }

        /// <summary>
        /// The profile's stores - fresh (from the <see cref="StoreDetailsCacheTtl"/> cache) when Outlook
        /// answers, else the list it gave last - or null when it has not answered since this server
        /// started. A store's ID does not change while it stays attached, so the last list is a sound
        /// answer about which store an ID names; it is never used for where a FOLDER is.
        /// </summary>
        private IReadOnlyList<ComStoreDetail>? TryGetStoreDetailsForIds()
        {
            try
            {
                return _gateway.Run(s => GetStoreDetails(s));
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                return LastKnownStoreDetails();
            }
        }

        /// <summary>
        /// The store a <c>store_id</c> argument names - refused, with words an agent can act on, when the
        /// value is not a store ID, names no store of the profile, names several, or cannot be looked up
        /// because Outlook has never answered.
        /// </summary>
        internal ComStoreDetail ResolveStoreId(string storeId, string parameterName)
        {
            if (!OutlookIds.TryDecodeStoreId(storeId, parameterName, out StoreIdParts? parts, out string? problem) || parts == null)
            {
                throw new ArgumentException(problem, parameterName);
            }

            IReadOnlyList<ComStoreDetail>? stores = TryGetStoreDetailsForIds();
            if (stores == null)
            {
                throw new ArgumentException(
                    "store_id '" + parts.Id + "' cannot be looked up right now: Outlook is not answering and has not listed "
                    + "its stores since this server started. Outlook is started when it is needed - retry in a few "
                    + "seconds; outlook_health shows its state.",
                    parameterName);
            }

            return PickStore(stores, parts, parameterName);
        }

        /// <summary>
        /// The ID to show for a store known by its own entry id and/or any entry id inside it: the
        /// profile's list decides when it has the store (so every tool shows one store one way), else the
        /// provider UID the entry id carries.
        /// </summary>
        private string? StoreIdForEntries(string? storeEntryId, string? anyEntryIdInStore)
        {
            if (!string.IsNullOrEmpty(storeEntryId))
            {
                IReadOnlyList<ComStoreDetail>? known = LastKnownStoreDetails();
                ComStoreDetail? detail = known?.FirstOrDefault(d =>
                    string.Equals(d.StoreId, storeEntryId, StringComparison.OrdinalIgnoreCase));
                if (detail != null)
                {
                    return StoreIdOf(detail);
                }
            }

            return OutlookIds.StoreIdOfEntryId(anyEntryIdInStore);
        }

        /// <summary>
        /// An account's delivery store's ID: the store whose own entry id the account reports, else the one
        /// store of that display name - never a guess between two. Pure.
        /// </summary>
        internal static string? DeliveryStoreIdOf(ComAccountInfo account, IReadOnlyList<ComStoreDetail> stores)
        {
            if (account == null || stores == null)
            {
                return null;
            }

            ComStoreDetail? byId = account.DeliveryStoreEntryId == null
                ? null
                : stores.FirstOrDefault(d => string.Equals(d.StoreId, account.DeliveryStoreEntryId, StringComparison.OrdinalIgnoreCase));
            if (byId != null)
            {
                return StoreIdOf(byId);
            }

            List<ComStoreDetail> byName = stores
                .Where(d => string.Equals(d.DisplayName, account.DeliveryStoreDisplayName, StringComparison.OrdinalIgnoreCase))
                .ToList();
            return byName.Count == 1 ? StoreIdOf(byName[0]) : null;
        }

        // ------------------------------------------------------------------ folder IDs

        /// <summary>A folder ID argument as the folder's entry id (hex) - or the refusal with the pointer to find_folder (Q135 (i)).</summary>
        internal static string DecodeFolderIdOrThrow(string? folderId, string parameterName)
        {
            if (!OutlookIds.TryDecodeFolderId(folderId, parameterName, out string entryIdHex, out string? problem))
            {
                throw new ArgumentException(problem, parameterName);
            }

            return entryIdHex;
        }

        /// <summary>
        /// The folder a folder ID names, opened in Outlook (read-only) - refused when no store of the
        /// profile has it any more.
        /// </summary>
        private ComFolderDetail DescribeFolderOrThrow(string folderId, string entryIdHex, string parameterName, int? budgetMs = null)
        {
            Func<IOutlookSession, (ComFolderDetail? Detail, string? Error)> describe = s =>
            {
                ComFolderDetail? d = s.TryDescribeFolder(entryIdHex, out string? e);
                return (d, e);
            };
            (ComFolderDetail? detail, string? error) = budgetMs.HasValue
                ? _gateway.Run(describe, budgetMs.Value)
                : _gateway.Run(describe);
            if (detail != null)
            {
                return detail;
            }

            throw new ArgumentException(DescribeFolderNotOpened(parameterName, folderId, error), parameterName);
        }

        /// <summary>Why a folder ID did not open, as a refusal. Pure.</summary>
        public static string DescribeFolderNotOpened(string parameterName, string folderId, string? comError)
        {
            if (comError == null || comError == "FolderNotFound")
            {
                return "No folder in this Outlook profile has the ID '" + folderId + "' (" + parameterName + "): the folder "
                    + "was deleted for good (a folder in Deleted Items still keeps its ID), or the ID comes from another "
                    + "profile or from a data file no longer attached. Get a current folderId from find_folder or "
                    + "list_folders.";
            }

            return "The folder with ID '" + folderId + "' (" + parameterName + ") could not be read in Outlook ("
                + comError + "). Nothing was changed. Retry when Outlook is responsive (see outlook_health).";
        }

        // ------------------------------------------------------------------ search: the scope from IDs

        /// <summary>
        /// Turns a search's <see cref="SearchRequest.StoreId"/> and <see cref="SearchRequest.FolderId"/>
        /// into the scope every tier uses, and records what the rest of the search reads
        /// (<see cref="SearchRequest.ScopeStoreName"/>, <see cref="SearchRequest.ScopeFolderPath"/>,
        /// <see cref="SearchRequest.ScopeFolderEntryId"/>). Null when the search names neither.
        /// <paramref name="folderFoundVia"/> says how a folder was found in the index.
        /// </summary>
        private FolderScopeResolution? ResolveSearchScope(SearchRequest request, out string? folderFoundVia)
        {
            folderFoundVia = null;
            if (request.FolderId != null)
            {
                return ResolveSearchFolder(request, out folderFoundVia);
            }

            if (request.StoreId == null)
            {
                return null;
            }

            if (!OutlookIds.TryDecodeStoreId(request.StoreId, "store_id", out StoreIdParts? parts, out string? problem) || parts == null)
            {
                throw new ArgumentException(problem, "store_id");
            }

            IReadOnlyList<ComStoreDetail>? stores = TryGetStoreDetailsForIds();
            if (stores != null)
            {
                ComStoreDetail detail = PickStore(stores, parts, "store_id");
                request.ScopeStoreName = detail.DisplayName;
                return ResolveFolderScope(detail, detail.DisplayName, null, request.IncludeSubfolders);
            }

            // Outlook has never answered: the index alone, by the provider UID the store's items carry.
            StoreScopeInfo root = FindIndexRootOrThrow(parts.ProviderUidHex, parts.Id, "store_id");
            request.ScopeStoreName = root.StoreDisplayName;
            return FolderScopeResolver.ForPrimaryStore(root.StorePrefix, null, request.IncludeSubfolders);
        }

        /// <summary>
        /// A search's folder by its ID. A local mail file's folder (a PST-format ID) is found by its own
        /// index row first - the store's root by the provider UID its items carry, the folder's row by
        /// the node id the ID carries (Q132 (i)), with no Outlook call - so it is searched where the index
        /// files its mail now, a rename or move the index has not re-filed yet included. Any other
        /// folder, and a local one the index has no row for, is opened in Outlook, and its names make the
        /// scope. An Exchange folder while Outlook is not running is REFUSED until the index route is
        /// measured on Exchange (Q132 (i)): its <c>System.ProviderItemID</c> is unknown.
        /// </summary>
        private FolderScopeResolution ResolveSearchFolder(SearchRequest request, out string? folderFoundVia)
        {
            folderFoundVia = null;
            string folderId = request.FolderId!;
            string entryIdHex = DecodeFolderIdOrThrow(folderId, "folder_id");
            StoreHash.TryDecodeHex(entryIdHex, out byte[] bytes);
            FolderIdLayout layout = OutlookIds.Classify(bytes);
            string? folderUid = OutlookIds.ProviderUidHexOf(entryIdHex);
            CheckFolderInRequestedStore(request.StoreId, folderUid, folderId);

            request.ScopeFolderEntryId = entryIdHex;
            string? indexMiss = null;
            if (layout == FolderIdLayout.Pst && folderUid != null)
            {
                FolderScopeResolution? byRow = TryResolveFolderByIndexRow(request, bytes, folderUid, out indexMiss);
                if (byRow != null)
                {
                    folderFoundVia = FolderFoundViaIndex;
                    return byRow;
                }
            }

            ComFolderDetail? detail = null;
            string? comError = null;
            bool outlookAnswered;
            try
            {
                (detail, comError) = _gateway.Run(s =>
                {
                    ComFolderDetail? d = s.TryDescribeFolder(entryIdHex, out string? e);
                    return (d, e);
                });
                outlookAnswered = true;
            }
            catch (Exception ex) when (ex is not OutOfMemoryException && ex is not ArgumentException)
            {
                outlookAnswered = false;
            }

            if (detail == null)
            {
                if (outlookAnswered)
                {
                    throw new ArgumentException(DescribeFolderNotOpened("folder_id", folderId, comError), "folder_id");
                }

                throw new ArgumentException(DescribeFolderUnreachable(folderId, layout, indexMiss), "folder_id");
            }

            request.ScopeFolderEntryId = detail.EntryId;
            request.ScopeStoreName = detail.StoreDisplayName;
            request.ScopeFolderPath = FolderScopeResolver.JoinNames(detail.PathSegments);
            folderFoundVia = FolderFoundViaOutlook;
            ComStoreDetail? storeDetail = TryGetStoreDetailsForIds()?.FirstOrDefault(d =>
                string.Equals(d.StoreId, detail.StoreEntryId, StringComparison.OrdinalIgnoreCase));
            return ResolveFolderScope(
                storeDetail, detail.StoreDisplayName, detail.IsStoreRoot ? null : detail.PathSegments, request.IncludeSubfolders);
        }

        /// <summary>
        /// An exhaustive scan's scope from IDs: a folder ID is opened in Outlook first - the scan opens it
        /// again by its entry id, and this read refuses a stale ID before a scan of minutes starts - and a
        /// store ID is looked up in the profile's list. A store scan finds its store by display name in
        /// Outlook, so a store whose name another store shares is refused: scan a folder of it instead.
        /// </summary>
        private void ResolveExhaustiveScope(SearchRequest request)
        {
            if (request.FolderId != null)
            {
                string entryId = DecodeFolderIdOrThrow(request.FolderId, "folder_id");
                CheckFolderInRequestedStore(request.StoreId, OutlookIds.ProviderUidHexOf(entryId), request.FolderId);
                ComFolderDetail detail = DescribeFolderOrThrow(request.FolderId, entryId, "folder_id");
                request.ScopeFolderEntryId = detail.EntryId;
                request.ScopeStoreName = detail.StoreDisplayName;
                request.ScopeFolderPath = FolderScopeResolver.JoinNames(detail.PathSegments);
                return;
            }

            ComStoreDetail store = ResolveStoreId(request.StoreId!, "store_id");
            IReadOnlyList<ComStoreDetail> sameName = (TryGetStoreDetailsForIds() ?? Array.Empty<ComStoreDetail>())
                .Where(d => string.Equals(d.DisplayName, store.DisplayName, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (sameName.Count > 1)
            {
                throw new ArgumentException(
                    sameName.Count.ToString(CultureInfo.InvariantCulture) + " stores in this Outlook profile are named '"
                    + store.DisplayName + "', and a whole-store exhaustive scan finds its store by that name in Outlook. Scan "
                    + "one folder of it by folder_id instead (find_folder, list_folders), or rename one of the stores.",
                    "store_id");
            }

            request.ScopeStoreName = store.DisplayName;
        }

        /// <summary>
        /// Refuses a <c>store_id</c> that names another store than the <c>folder_id</c> beside it - a
        /// folder ID names its own store, so the two can only agree or contradict. Compared by provider
        /// UID when the store ID carries one; a store ID built from a store's own entry id is checked
        /// against the profile's list.
        /// </summary>
        private void CheckFolderInRequestedStore(string? storeId, string? folderUid, string folderId)
        {
            if (storeId == null)
            {
                return;
            }

            if (!OutlookIds.TryDecodeStoreId(storeId, "store_id", out StoreIdParts? parts, out string? problem) || parts == null)
            {
                throw new ArgumentException(problem, "store_id");
            }

            string? storeUid = parts.ProviderUidHex;
            if (storeUid == null)
            {
                IReadOnlyList<ComStoreDetail>? stores = TryGetStoreDetailsForIds();
                storeUid = stores == null ? null : OutlookIds.ProviderUidHexOf(PickStore(stores, parts, "store_id").RootFolderEntryId);
            }

            if (storeUid != null && folderUid != null && !string.Equals(storeUid, folderUid, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException(
                    "folder_id '" + folderId + "' is a folder of another store than store_id '" + parts.Id + "'. A folder ID "
                    + "names its own store, so store_id is not needed beside it - pass folder_id alone, or the storeId "
                    + "find_folder shows beside the folder.",
                    "store_id");
            }
        }

        /// <summary>
        /// Q132 (i): a PST-format folder ID's scope from the index alone. Null - with
        /// <paramref name="miss"/> saying why - when no single index store carries the ID's store UID, or
        /// no single row under it carries its node id (a folder that has never held indexed mail, a store
        /// the index does not cover). Never a guess: two roots or two rows are a miss, not a pick.
        /// </summary>
        private FolderScopeResolution? TryResolveFolderByIndexRow(SearchRequest request, byte[] entryId, string folderUid, out string? miss)
        {
            miss = null;
            string? providerItemId = OutlookIds.PstProviderItemId(entryId);
            if (providerItemId == null)
            {
                miss = "the ID is not a local mail file's";
                return null;
            }

            try
            {
                IReadOnlyList<StoreScopeInfo> roots = _index.Value.FindStoreRootsByProviderUid(folderUid, SearchIndexTimeoutSeconds);
                if (roots.Count != 1)
                {
                    miss = roots.Count == 0
                        ? "no store in the search index carries the folder's store"
                        : roots.Count.ToString(CultureInfo.InvariantCulture) + " stores in the search index carry the folder's store";
                    return null;
                }

                IReadOnlyList<string> rows = _index.Value.FindRowUrlsByProviderItemId(roots[0].StorePrefix, providerItemId, SearchIndexTimeoutSeconds);
                if (rows.Count != 1)
                {
                    miss = rows.Count == 0
                        ? "the search index has no row for the folder itself"
                        : rows.Count.ToString(CultureInfo.InvariantCulture) + " index rows carry the folder's node id";
                    return null;
                }

                FolderScopeResolution scope = FolderScopeResolver.ForIndexedFolder(roots[0].StorePrefix, rows[0], request.IncludeSubfolders);
                request.ScopeStoreName = StoreNameForUid(folderUid) ?? roots[0].StoreDisplayName;
                request.ScopeFolderPath = scope.RequestedFolder;
                return scope;
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                miss = "the search index did not answer (" + ex.GetType().Name + ")";
                return null;
            }
        }

        /// <summary>The display name Outlook gives the store whose items carry <paramref name="providerUidHex"/>, when its list has exactly one.</summary>
        private string? StoreNameForUid(string providerUidHex)
        {
            List<ComStoreDetail> named = (TryGetStoreDetailsForIds() ?? Array.Empty<ComStoreDetail>())
                .Where(d => string.Equals(OutlookIds.ProviderUidHexOf(d.RootFolderEntryId), providerUidHex, StringComparison.OrdinalIgnoreCase))
                .ToList();
            return named.Count == 1 ? named[0].DisplayName : null;
        }

        /// <summary>
        /// The one index store whose items carry <paramref name="providerUidHex"/> - the route a store ID
        /// takes while Outlook has never answered - or a refusal.
        /// </summary>
        private StoreScopeInfo FindIndexRootOrThrow(string? providerUidHex, string storeId, string parameterName)
        {
            IReadOnlyList<StoreScopeInfo> roots = providerUidHex == null
                ? Array.Empty<StoreScopeInfo>()
                : _index.Value.FindStoreRootsByProviderUid(providerUidHex, SearchIndexTimeoutSeconds);
            if (roots.Count == 1)
            {
                return roots[0];
            }

            throw new ArgumentException(
                "store_id '" + storeId + "' cannot be searched right now: Outlook is not answering and has not listed its "
                + "stores since this server started, and " + (roots.Count == 0
                    ? "the search index holds no store whose mail carries this ID"
                    : roots.Count.ToString(CultureInfo.InvariantCulture) + " stores in the search index carry it, so it cannot say which")
                + ". Outlook is started when it is needed - retry in a few seconds; outlook_health shows its state.",
                parameterName);
        }

        /// <summary>Why a folder ID could be resolved neither through the index nor through Outlook. Pure.</summary>
        public static string DescribeFolderUnreachable(string folderId, FolderIdLayout layout, string? indexMiss)
        {
            if (layout == FolderIdLayout.Pst)
            {
                return "folder_id '" + folderId + "' could not be found in the search index (" + (indexMiss ?? "no row")
                    + "), and Outlook is not answering to look it up. Outlook is started when it is needed - retry in a few "
                    + "seconds; outlook_health shows its state.";
            }

            return "folder_id '" + folderId + "' is a folder of an Exchange mailbox (or of another non-local store), and "
                + "searching one folder of such a mailbox needs Outlook: the route through the search index alone is "
                + "measured for local mail files only so far (Q132). Outlook is not answering right now - it is started "
                + "when it is needed, so retry in a few seconds - or search the whole mailbox with its store_id meanwhile.";
        }

        /// <summary>
        /// The index predicates for a store - known as one of the profile's stores (<paramref name="detail"/>)
        /// or, without one, by its display name alone - and a folder given as its NAMES, choosing the shape
        /// the store's namespace supports. The store is decided by IDENTITY when it is known
        /// (<see cref="DecideStore"/>): never by a name two stores may share.
        /// <para>
        /// ⚠ THE DELEGATE FIX (soak fix 15). A delegate folder used to be scoped by a NESTED delegate URL
        /// (<c>&lt;host&gt;/1/&lt;delegate&gt;/&lt;path&gt;</c>). The delegate index namespace is FLAT -
        /// intermediate folders are dropped from both the mapi URL and System.ItemFolderPathDisplay - so
        /// that URL addressed a folder that does not exist and every delegate SUBFOLDER search returned 0
        /// rows, silently (measured: 8/8 probed subfolders, ~3,871 items across 15 subfolders unreachable
        /// on this profile). Delegate scopes go through <see cref="FolderScopeResolver.ForDelegateStore"/>:
        /// delegate STORE ROOT scope + folder-NAME equality, verified exact against COM ground truth.
        /// </para>
        /// </summary>
        private FolderScopeResolution ResolveFolderScope(
            ComStoreDetail? detail, string store, IReadOnlyList<string>? folderSegments, bool includeSubfolders)
        {
            string? folder = folderSegments == null || folderSegments.Count == 0 ? null : FolderScopeResolver.JoinNames(folderSegments);

            // By the store's own name AND hash first (Q113 (a)): its own index store, whatever the
            // profile calls it. A store the map ties to nothing is a store the index does not hold -
            // searched with the index tier skipped, never widened, and never scoped to a same-named
            // store of another profile, which is what a name lookup would find. A store whose mail the
            // index cannot attribute is REFUSED rather than guessed. Only the name rule's own cases - no
            // map, THE ONE OPEN EXCEPTION (Exchange, Q113 (b)) and delegates - take the name rule below.
            (StoreNameDecision decision, StoreScopeInfo? matchedRoot, string? mapRefusal) = detail != null
                ? DecideStore(detail, SearchIndexTimeoutSeconds)
                : DecideStoreName(store, SearchIndexTimeoutSeconds);
            switch (decision)
            {
                case StoreNameDecision.Tied:
                    return FolderScopeResolver.ForPrimaryStoreSegments(matchedRoot!.StorePrefix, folderSegments, includeSubfolders);
                case StoreNameDecision.NotIndexed:
                    return FolderScopeResolver.ForUnindexedStore(folder);
                case StoreNameDecision.Refused:
                    throw new ArgumentException(mapRefusal, "store_id");
            }

            IReadOnlyList<StoreScopeInfo> catalog = GetCatalog(SearchIndexTimeoutSeconds);
            StoreScopeInfo? match = catalog.FirstOrDefault(s =>
                string.Equals(s.StoreDisplayName, store, StringComparison.OrdinalIgnoreCase));

            if (match == null && store.IndexOf('@') >= 0)
            {
                match = _index.Value.TryDiscoverStoreScopeByAddress(store, SearchIndexTimeoutSeconds);
                if (match != null)
                {
                    InvalidateCatalog(match);
                }
            }

            if (match != null)
            {
                // Not even by the index's own name for it does a search reach an index store whose
                // mail the index cannot attribute to one store (Q113 (a)).
                StoreIndexMatch? contested = TryGetStoreIndexMap(SearchIndexTimeoutSeconds)?.Unmatchable
                    .FirstOrDefault(m => m.Contested.Any(r =>
                        string.Equals(r.StorePrefix, match.StorePrefix, StringComparison.OrdinalIgnoreCase)));
                if (contested != null)
                {
                    throw new ArgumentException(DescribeUnmatchableStore(contested), "store_id");
                }

                return FolderScopeResolver.ForPrimaryStoreSegments(match.StorePrefix, folderSegments, includeSubfolders);
            }

            // Delegate store: scope under an owner's /1/<name> subtree.
            foreach (StoreScopeInfo owner in catalog)
            {
                string delegateScope = DelegateScope(owner, store);
                bool exists;
                try
                {
                    exists = _index.Value.ScopeHasAnyItem(delegateScope, SearchIndexTimeoutSeconds);
                }
                catch (Exception ex) when (ex is not OutOfMemoryException)
                {
                    exists = false;
                }

                if (exists)
                {
                    // The COM tree is needed only for a delegate FOLDER scope: to map the
                    // requested subtree onto the flat leaf names it contains, and to spot
                    // leaf names the flat namespace merges.
                    ComFolderPathList? tree = folder == null ? null : TryGetFolderPaths(store);
                    return FolderScopeResolver.ForDelegateStore(
                        delegateScope, folder, includeSubfolders, tree?.Paths, tree?.Incomplete ?? false);
                }
            }

            // NOTHING IN THE INDEX ADDRESSES THIS STORE - a store the profile HAS (it was named by its ID,
            // or found by it): the search proceeds with the index tier skipped - never widened, because the
            // store filters which mail may come back and a widened scope would answer with another
            // store's mail.
            return FolderScopeResolver.ForUnindexedStore(folder);
        }

        /// <summary>
        /// The store map's decision for a store known by IDENTITY (Q133): its own index store, not indexed,
        /// refused - or the name rule, where the map cannot speak for it (no map, the open exception -
        /// Exchange, Q113 (b) - or a delegate). Never decided by its name: two stores may share one.
        /// </summary>
        private (StoreNameDecision Decision, StoreScopeInfo? Root, string? Refusal) DecideStore(ComStoreDetail detail, int? commandTimeoutSeconds)
        {
            StoreIndexMap? map = TryGetStoreIndexMap(commandTimeoutSeconds);
            if (map == null)
            {
                bool listingFailed;
                lock (_catalogLock)
                {
                    listingFailed = _storeIndexListingFailed;
                }

                return listingFailed && detail.ExchangeStoreType == 3
                    ? (StoreNameDecision.Refused, null, DescribeUnreadableStoreList(detail.DisplayName))
                    : (StoreNameDecision.NameRule, null, null);
            }

            StoreIndexMatch? match = map.Stores.FirstOrDefault(m =>
                string.Equals(m.Store.StoreEntryId, detail.StoreId, StringComparison.OrdinalIgnoreCase));
            if (match == null)
            {
                return (StoreNameDecision.NameRule, null, null);
            }

            switch (match.Kind)
            {
                case StoreIndexMatchKind.NameAndHash:
                case StoreIndexMatchKind.StoreHash:
                    return (StoreNameDecision.Tied, match.Root, null);
                case StoreIndexMatchKind.None:
                    return (StoreNameDecision.NotIndexed, null, null);
                case StoreIndexMatchKind.Ambiguous:
                    return (StoreNameDecision.Refused, null, DescribeUnmatchableStore(match));
                default:
                    return (StoreNameDecision.NameRule, null, null);
            }
        }

        // ------------------------------------------------------------------ find_folder (Q136 (ii))

        /// <summary>
        /// find_folder: the folders whose NAME matches <paramref name="name"/> - every store, or the one
        /// <paramref name="storeId"/> names - as candidates with their ID, store and display path. It never
        /// acts: picking one, and acting on it by its ID, is the caller's (Q136 (ii)). Read-only; needs
        /// Outlook, like list_folders, which it walks the same way.
        /// </summary>
        public FindFolderOutcome FindFolder(string? name, string? storeId = null)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException(
                    "name is required: the folder's name, or part of it (matched case-insensitively).", nameof(name));
            }

            ComStoreDetail? detail = storeId == null ? null : ResolveStoreId(storeId, "store_id");
            ComFolderTree tree = _gateway.Run(s => s.ListFolders(detail?.StoreId, FolderWalkAbsoluteCap));
            return MatchFolders(tree, name!, TryGetStoreDetailsForIds(), FindFolderCandidatesCap);
        }

        /// <summary>
        /// The find_folder answer for a walked tree. Pure, so T1 pins the rules: a folder matches by its
        /// own name EXACTLY (case-insensitive), by its display PATH ending with the text (so
        /// <c>Projects/2026</c> finds the folder 2026 inside Projects - a name may hold <c>/</c>, and both
        /// readings are offered), or PARTIALLY (its name contains the text); exact first, then path, then
        /// partial, each in walk order; capped at <paramref name="cap"/> with the total kept.
        /// </summary>
        public static FindFolderOutcome MatchFolders(
            ComFolderTree tree, string name, IReadOnlyList<ComStoreDetail>? stores, int cap = FindFolderCandidatesCap)
        {
            if (tree == null)
            {
                throw new ArgumentNullException(nameof(tree));
            }

            string wanted = (name ?? string.Empty).Trim();
            if (wanted.Length == 0)
            {
                throw new ArgumentException("name is required.", nameof(name));
            }

            List<(int Rank, int Order, ComFolderInfo Folder)> matched = new List<(int, int, ComFolderInfo)>();
            for (int i = 0; i < tree.Folders.Count; i++)
            {
                ComFolderInfo f = tree.Folders[i];
                int rank;
                if (string.Equals(f.Name, wanted, StringComparison.OrdinalIgnoreCase))
                {
                    rank = 0;
                }
                else if (string.Equals(f.Path, wanted, StringComparison.OrdinalIgnoreCase)
                    || f.Path.EndsWith("/" + wanted, StringComparison.OrdinalIgnoreCase))
                {
                    rank = 1;
                }
                else if (f.Name.IndexOf(wanted, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    rank = 2;
                }
                else
                {
                    continue;
                }

                matched.Add((rank, i, f));
            }

            List<FolderCandidateView> candidates = matched
                .OrderBy(m => m.Rank)
                .ThenBy(m => m.Order)
                .Take(Math.Max(1, cap))
                .Select(m => ToCandidate(m.Folder, m.Rank, tree, stores))
                .ToList();

            List<string> advice = new List<string>();
            if (matched.Count == 0)
            {
                advice.Add("No folder's name matches '" + wanted + "'. Names are matched case-insensitively, as a whole name, "
                    + "as the end of a folder's path, or as part of a name. list_folders shows the whole tree with every "
                    + "folderId; create_folder makes a folder that is not there yet.");
            }
            else if (matched.Count > 1 && matched.Count(m => m.Rank == 0) != 1)
            {
                advice.Add("More than one folder matches: pick the right one by its store and path (ask the user if it "
                    + "is not clear) and pass its folderId. Never guess between folders that share a name.");
            }

            if (tree.WalkCapReached || tree.DepthLimitReached)
            {
                advice.Add("The folder walk was cut short (" + (tree.WalkCapReached ? "walk cap" : "depth limit")
                    + "), so a matching folder beyond it is not listed - narrow with store_id.");
            }

            return new FindFolderOutcome
            {
                Name = wanted,
                Candidates = candidates,
                Total = matched.Count,
                Truncated = matched.Count > candidates.Count,
                WalkCapReached = tree.WalkCapReached ? true : (bool?)null,
                DepthLimitReached = tree.DepthLimitReached ? true : (bool?)null,
                Advice = advice.Count > 0 ? advice : null,
            };
        }

        private static FolderCandidateView ToCandidate(ComFolderInfo folder, int rank, ComFolderTree tree, IReadOnlyList<ComStoreDetail>? stores)
        {
            return new FolderCandidateView
            {
                FolderId = OutlookIds.FolderIdFromEntryId(folder.EntryId) ?? string.Empty,
                Name = folder.Name,
                Store = folder.StoreDisplayName,
                StoreId = StoreIdForWalkedFolder(folder, tree, stores),
                Folder = folder.Path,
                Match = rank == 0 ? "exact" : rank == 1 ? "path" : "partial",
                Items = folder.ItemCount,
                Unread = folder.UnreadCount,
            };
        }

        /// <summary>
        /// The store ID of a folder from a walk: the profile's list decides when it has the store the folder
        /// belongs to (matched by the provider UID the folder's id carries, or by the walked store's own
        /// entry id), so list_folders, find_folder and list_accounts show one store one way. Pure.
        /// </summary>
        internal static string? StoreIdForWalkedFolder(ComFolderInfo folder, ComFolderTree tree, IReadOnlyList<ComStoreDetail>? stores)
        {
            string? uid = OutlookIds.ProviderUidHexOf(folder.EntryId);
            ComStoreRoot? root = tree.Stores.FirstOrDefault(r =>
                    uid != null && string.Equals(OutlookIds.ProviderUidHexOf(r.RootFolderEntryId), uid, StringComparison.OrdinalIgnoreCase))
                ?? tree.Stores.FirstOrDefault(r => string.Equals(r.DisplayName, folder.StoreDisplayName, StringComparison.OrdinalIgnoreCase));
            return StoreIdForRoot(root, stores) ?? OutlookIds.StoreIdOfEntryId(folder.EntryId);
        }

        /// <summary>The ID of the root folder of the store a walked folder belongs to - parent_id for a top-level create_folder. Pure.</summary>
        internal static string? RootFolderIdForWalkedStore(ComFolderInfo folder, ComFolderTree tree)
        {
            string? uid = OutlookIds.ProviderUidHexOf(folder.EntryId);
            ComStoreRoot? root = tree.Stores.FirstOrDefault(r =>
                    uid != null && string.Equals(OutlookIds.ProviderUidHexOf(r.RootFolderEntryId), uid, StringComparison.OrdinalIgnoreCase))
                ?? tree.Stores.FirstOrDefault(r => string.Equals(r.DisplayName, folder.StoreDisplayName, StringComparison.OrdinalIgnoreCase));
            return OutlookIds.FolderIdFromEntryId(root?.RootFolderEntryId);
        }

        /// <summary>The ID of a walked store: the profile's list decides when it has it, else its own ids. Pure.</summary>
        internal static string? StoreIdForRoot(ComStoreRoot? root, IReadOnlyList<ComStoreDetail>? stores)
        {
            if (root == null)
            {
                return null;
            }

            ComStoreDetail? detail = stores?.FirstOrDefault(d =>
                !string.IsNullOrEmpty(root.StoreEntryId)
                && string.Equals(d.StoreId, root.StoreEntryId, StringComparison.OrdinalIgnoreCase));
            return detail != null ? StoreIdOf(detail) : OutlookIds.StoreIdOf(root.RootFolderEntryId, root.StoreEntryId);
        }

        // ------------------------------------------------------------------ create_folder (Q134 (i))

        /// <summary>
        /// create_folder: one mail folder named <paramref name="name"/> - verbatim; a <c>/</c> is part of the
        /// name, never a separator - directly under the folder <paramref name="parentId"/>. Refused, with
        /// nothing created, under Deleted Items, the Outbox or a folder that is not a mail folder, and when
        /// the parent already has a folder of that name - whose folderId the refusal gives. Audited.
        /// </summary>
        public CreateFolderOutcome CreateFolder(string? parentId, string? name)
        {
            string parentEntryId = DecodeFolderIdOrThrow(parentId, "parent_id");
            ValidateFolderName(name);

            (ComFolderDetail? created, string? error) = _gateway.Run(s =>
            {
                ComFolderDetail? d = s.TryCreateFolder(parentEntryId, name!, out string? e);
                return (d, e);
            });

            if (created == null)
            {
                throw DescribeCreateFolderFailure(error, parentId!, name!);
            }

            string folderId = OutlookIds.FolderIdFromEntryId(created.EntryId) ?? string.Empty;
            string path = FolderScopeResolver.JoinNames(created.PathSegments);
            string parentPath = created.PathSegments.Count > 1
                ? FolderScopeResolver.JoinNames(created.PathSegments.Take(created.PathSegments.Count - 1).ToList())
                : string.Empty;
            string? storeId = StoreIdForEntries(created.StoreEntryId, created.EntryId);
            try
            {
                Audit.AuditLog.Append(
                    "create_folder",
                    ("store", created.StoreDisplayName),
                    ("storeId", storeId),
                    ("parentFolder", parentPath),
                    ("parentFolderId", OutlookIds.Canonical(parentId) ?? parentId),
                    ("folder", path),
                    ("folderId", folderId),
                    ("folderEntryId", created.EntryId));
            }
            catch (InvalidOperationException ex)
            {
                throw new OperationOutcomeException(
                    Com.MutationOutcome.Applied,
                    "The folder WAS created (folderId " + folderId + ", '" + path + "' in store '" + created.StoreDisplayName
                    + "') but the audit line could not be written: " + ex.Message + " Do not create it again.");
            }

            return new CreateFolderOutcome
            {
                Created = true,
                FolderId = folderId,
                Name = created.Name,
                Folder = path,
                Store = created.StoreDisplayName,
                StoreId = storeId,
                ParentId = OutlookIds.Canonical(parentId) ?? parentId!,
                ParentFolder = parentPath,
            };
        }

        /// <summary>The name rules create_folder holds a name to before anything is asked of Outlook. Pure.</summary>
        public static void ValidateFolderName(string? name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("name is required: the new folder's name, exactly as it should read.", nameof(name));
            }

            if (name!.Length > FolderNameCharsCap)
            {
                throw new ArgumentException(
                    "name is too long: at most " + FolderNameCharsCap.ToString(CultureInfo.InvariantCulture) + " characters.",
                    nameof(name));
            }

            foreach (char c in name)
            {
                if (char.IsControl(c))
                {
                    throw new ArgumentException("name must not contain control characters (tabs, line breaks).", nameof(name));
                }
            }
        }

        /// <summary>What a refused or failed create_folder tells the caller - the outcome decided by where it stopped. Pure.</summary>
        public static OperationOutcomeException DescribeCreateFolderFailure(string? error, string parentId, string name)
        {
            const string nothing = " Nothing was created.";
            if (error != null && error.StartsWith("FolderExists:", StringComparison.Ordinal))
            {
                string existing = OutlookIds.FolderIdFromEntryId(error.Substring("FolderExists:".Length)) ?? "unknown";
                return new OperationOutcomeException(
                    Com.MutationOutcome.Unchanged,
                    "The parent folder already has a folder named '" + name + "' (Outlook compares names case-insensitively): "
                    + "its folderId is " + existing + " - use that one." + nothing);
            }

            if (error != null && error.StartsWith("CreateFailed:", StringComparison.Ordinal))
            {
                return new OperationOutcomeException(
                    Com.MutationOutcome.Unchanged,
                    "Outlook refused to create the folder (" + error.Substring("CreateFailed:".Length) + "), and the parent "
                    + "was checked afterwards: no folder of that name is there." + nothing);
            }

            if (error != null && error.StartsWith("CreatedUnreadable:", StringComparison.Ordinal))
            {
                string made = OutlookIds.FolderIdFromEntryId(error.Substring("CreatedUnreadable:".Length)) ?? "unknown";
                return new OperationOutcomeException(
                    Com.MutationOutcome.Applied,
                    "The folder WAS created (folderId " + made + ") but could not be read back. Do not create it again; "
                    + "find_folder shows it.");
            }

            switch (error)
            {
                case "ParentFolderNotFound":
                    return new OperationOutcomeException(
                        Com.MutationOutcome.Unchanged,
                        DescribeFolderNotOpened("parent_id", parentId, null) + nothing);
                case "TargetNotAMailFolder":
                    return new OperationOutcomeException(
                        Com.MutationOutcome.Unchanged,
                        "Refused: the parent is not a mail folder (a calendar, contacts or other folder), and this tool makes "
                        + "mail folders only." + nothing);
                case "TargetIsDeletedItems":
                    return new OperationOutcomeException(
                        Com.MutationOutcome.Unchanged,
                        "Refused: folders are not made in Deleted Items or under it - that is the bin." + nothing);
                case "TargetIsOutbox":
                    return new OperationOutcomeException(
                        Com.MutationOutcome.Unchanged,
                        "Refused: folders are not made in the Outbox." + nothing);
                case SpecialFolderGuards.TargetGuardUnreadable:
                    return new OperationOutcomeException(
                        Com.MutationOutcome.Unchanged,
                        "Refused: whether the parent is the store's Deleted Items (or under it) or its Outbox could not be "
                        + "checked - Outlook would not say which folders those are." + nothing
                        + " Retry when Outlook is responsive (see outlook_health).");
                default:
                    // Every case above is decided before Folders.Add runs, or checked after it; this one is the
                    // catch-all around the whole sequence.
                    return new OperationOutcomeException(
                        Com.MutationOutcome.Unknown,
                        "Creating the folder failed (" + (error ?? "unknown") + "). WHETHER IT WAS CREATED IS UNKNOWN - look "
                        + "with find_folder before trying again.");
            }
        }

        // ------------------------------------------------------------------ move_mail

        /// <summary>
        /// Moves 1-50 items (hit ids or EntryIDs) into the folder <paramref name="folderId"/>, each within its
        /// own store (D39 v1: same-store only; an item in another store than the folder is refused per item).
        /// The folder is checked ONCE, before any item: one that will not open, the top of a store, Deleted
        /// Items or anything under it, the Outbox or a folder that is not a mail folder refuses the whole call
        /// with nothing moved. Nothing is ever created: since Q134 (i) a folder is made only by create_folder.
        /// </summary>
        public MoveMailOutcome MoveMail(IReadOnlyList<string>? ids, string? folderId)
        {
            IReadOnlyList<string> requestIds = ValidateMoveIds(ids);
            if (string.IsNullOrWhiteSpace(folderId))
            {
                throw new ArgumentException(
                    "folder_id is required: the ID of the folder to move the mail into - get it from find_folder or "
                    + "list_folders, and make a folder that is not there yet with create_folder first.",
                    "folder_id");
            }

            string targetEntryId = DecodeFolderIdOrThrow(folderId, "folder_id");
            ComFolderDetail target = DescribeFolderOrThrow(folderId!, targetEntryId, "folder_id");
            string targetEcho = target.PathSegments.Count > 0 ? FolderScopeResolver.JoinNames(target.PathSegments) : target.Name;
            string? refusal = target.IsStoreRoot ? "TargetIsStoreRoot" : target.RefusedAsTarget;
            if (refusal != null)
            {
                throw new OperationOutcomeException(
                    Com.MutationOutcome.Unchanged, DescribeMoveRefusal(refusal, targetEcho) + " Nothing was moved.");
            }

            List<MoveItemView> items = new List<MoveItemView>(requestIds.Count);
            bool auditBroken = false;
            Stopwatch batchClock = Stopwatch.StartNew();
            foreach (string id in requestIds)
            {
                if (auditBroken)
                {
                    items.Add(FailedItem(id, "Not attempted: the audit log is unavailable (every move must be audited)."));
                    continue;
                }

                int itemBudgetMs = RemainingBatchBudgetMs(batchClock);
                if (itemBudgetMs < MinimumItemBudgetMs)
                {
                    items.Add(FailedItem(id, BatchBudgetExhaustedMessage));
                    continue;
                }

                MoveItemView item = MoveOne(id, target.EntryId, targetEcho, itemBudgetMs, out bool auditFailed);
                auditBroken |= auditFailed;
                items.Add(item);
            }

            int movedCount = items.Count(i => i.Ok);
            return new MoveMailOutcome
            {
                Requested = requestIds.Count,
                Moved = movedCount,
                Failed = items.Count - movedCount,
                TargetFolder = targetEcho,
                TargetFolderId = OutlookIds.FolderIdFromEntryId(target.EntryId),
                TargetStore = target.StoreDisplayName,
                TargetStoreId = StoreIdForEntries(target.StoreEntryId, target.EntryId),
                Items = items,
                Advice = movedCount > 0 ? new[] { MoveEntryIdAdvice } : null,
            };
        }

        private MoveItemView MoveOne(string id, string targetEntryId, string targetEcho, int itemBudgetMs, out bool auditFailed)
        {
            auditFailed = false;
            string entryId;
            string? storeId;
            string? hitId;
            try
            {
                (entryId, storeId, _, _, hitId) = ResolveToEntryId(id);
            }
            catch (Exception ex) when (ex is ArgumentException || ex is InvalidOperationException)
            {
                return FailedItem(id, ex.Message);
            }

            try
            {
                (ComMoveItemResult? moved, string? comError) = _gateway.Run(
                    s =>
                {
                    ComMoveItemResult? r = s.TryMoveItemToFolder(entryId, storeId, targetEntryId, out string? e);
                    if (ShouldSearchOtherStores(storeId, r != null, e))
                    {
                        foreach (ComStoreDetail candidate in GetStoreDetails(s))
                        {
                            r = s.TryMoveItemToFolder(entryId, candidate.StoreId, targetEntryId, out e);
                            if (!KeepSearchingStores(r != null, e))
                            {
                                break;
                            }
                        }
                    }

                    return (r, e);
                },
                    itemBudgetMs,
                    allowConnectFloor: true);

                if (moved == null)
                {
                    return FailedItem(id, DescribeMoveFailure(comError, targetEcho), MoveFailureOutcome(comError));
                }

                return CompleteMove("move_mail", id, hitId, moved, out auditFailed);
            }
            catch (OutlookUnavailableException)
            {
                throw;
            }
            catch (TimeoutException ex)
            {
                return FailedItem(id, BatchItemTimedOutMessage(ex.Message), Com.MutationOutcome.Unknown);
            }
            catch (Exception ex) when (ex is InvalidOperationException || OutlookComSession.IsComCallFailure(ex))
            {
                return FailedItem(id, ex.Message, OutcomeOf(ex));
            }
        }

        // ------------------------------------------------------------------ goto_folder / show_search_results

        /// <summary>
        /// Navigates the user's Outlook window to the folder <paramref name="folderId"/>, or - when that is
        /// null - to the Inbox of the store <paramref name="storeId"/>.
        /// </summary>
        public GotoFolderOutcome GotoFolder(string? folderId = null, string? storeId = null)
        {
            NavigationTarget target = ResolveNavigationTarget(folderId, storeId, required: true)!;
            ComExplorerState state = _gateway.Run(s =>
            {
                ComExplorerState? result = s.TryGotoFolder(target.FolderEntryId, target.StoreEntryId, out string? error);
                return result ?? throw BuildNavigationError(error);
            });

            return new GotoFolderOutcome
            {
                Store = target.StoreName,
                StoreId = target.StoreId,
                Folder = target.FolderPath ?? OutlookComSession.ToStoreRelativeFolderPath(state.CurrentFolderPath, target.StoreName),
                FolderId = OutlookIds.FolderIdFromEntryId(state.CurrentFolderEntryId) ?? target.FolderId,
                ExplorerFolderPath = state.CurrentFolderPath,
                ExplorerCaption = state.Caption,
                Displayed = true,
            };
        }

        /// <summary>
        /// Shows a result list in Outlook's own search UI, after navigating to the folder
        /// <paramref name="folderId"/> or the Inbox of the store <paramref name="storeId"/> when either is given.
        /// </summary>
        public ShowSearchResultsOutcome ShowSearchResults(
            string query, string scope = "current_folder", string? folderId = null, string? storeId = null)
        {
            if (string.IsNullOrWhiteSpace(query))
            {
                throw new ArgumentException("query is required.", nameof(query));
            }

            if (query.Length > ShowSearchQueryCharsCap)
            {
                throw new ArgumentException(
                    "query is too long for the Outlook search box (max "
                    + ShowSearchQueryCharsCap.ToString(CultureInfo.InvariantCulture) + " chars).",
                    nameof(query));
            }

            foreach (char c in query)
            {
                if (char.IsControl(c))
                {
                    throw new ArgumentException("query must not contain control characters.", nameof(query));
                }
            }

            int olScope = MapSearchScope(scope);
            NavigationTarget? target = ResolveNavigationTarget(folderId, storeId, required: false);
            ComExplorerState state = _gateway.Run(s =>
            {
                ComExplorerState? result = s.TryShowSearchResults(
                    query, olScope, target?.FolderEntryId, target?.StoreEntryId, out string? error);
                return result ?? throw BuildNavigationError(error);
            });

            List<string>? advice = null;
            if (HealthReporting.ReadUiSearchBackendFromRegistry() == HealthReporting.UiSearchBackendServerAssisted)
            {
                advice = new List<string> { ServerAssistedUiSearchAdvice };
            }

            return new ShowSearchResultsOutcome
            {
                Query = query,
                Scope = NormalizeScopeName(scope),
                ExplorerFolderPath = state.CurrentFolderPath,
                ExplorerCaption = state.Caption,
                Displayed = true,
                Advice = advice,
            };
        }

        /// <summary>Where a show-me tool navigates: a folder by its entry id, or a store's Inbox by the store's own entry id.</summary>
        private sealed class NavigationTarget
        {
            internal string? FolderEntryId { get; set; }

            internal string? FolderId { get; set; }

            internal string? FolderPath { get; set; }

            internal string? StoreEntryId { get; set; }

            internal string StoreName { get; set; } = string.Empty;

            internal string? StoreId { get; set; }
        }

        /// <summary>
        /// The folder or store a show-me tool was given, checked BEFORE the window moves: a folder ID is
        /// opened read-only first, so a stale one is refused with nothing changed on screen. Null when
        /// neither is given and none is <paramref name="required"/>.
        /// </summary>
        private NavigationTarget? ResolveNavigationTarget(string? folderId, string? storeId, bool required)
        {
            if (folderId != null)
            {
                string entryId = DecodeFolderIdOrThrow(folderId, "folder_id");
                CheckFolderInRequestedStore(storeId, OutlookIds.ProviderUidHexOf(entryId), folderId);
                ComFolderDetail detail = DescribeFolderOrThrow(folderId, entryId, "folder_id");
                return new NavigationTarget
                {
                    FolderEntryId = detail.EntryId,
                    FolderId = OutlookIds.FolderIdFromEntryId(detail.EntryId),
                    FolderPath = FolderScopeResolver.JoinNames(detail.PathSegments),
                    StoreEntryId = detail.StoreEntryId,
                    StoreName = detail.StoreDisplayName,
                    StoreId = StoreIdForEntries(detail.StoreEntryId, detail.EntryId),
                };
            }

            if (storeId != null)
            {
                ComStoreDetail store = ResolveStoreId(storeId, "store_id");
                return new NavigationTarget
                {
                    StoreEntryId = store.StoreId,
                    StoreName = store.DisplayName,
                    StoreId = StoreIdOf(store),
                };
            }

            if (required)
            {
                throw new ArgumentException(
                    "Pass folder_id (the folder to show - from find_folder or list_folders) or store_id (to show that "
                    + "store's Inbox - from list_accounts).",
                    "folder_id");
            }

            return null;
        }

        private static Exception BuildNavigationError(string? error)
        {
            if (error == "StoreNotFound")
            {
                return new OperationOutcomeException(
                    Com.MutationOutcome.Unchanged,
                    "The store could not be opened in Outlook, so nothing was opened or moved on screen. list_accounts "
                    + "lists the stores with their storeId.");
            }

            if (error == "FolderNotFound")
            {
                return new OperationOutcomeException(
                    Com.MutationOutcome.Unchanged,
                    "The folder could not be opened in Outlook, so nothing was opened or moved on screen. Get a current "
                    + "folderId from find_folder or list_folders.");
            }

            return new OperationOutcomeException(
                Com.MutationOutcome.Unknown,
                "Outlook could not confirm the requested view (" + (error ?? "unknown") + "). THE WINDOW MAY HAVE MOVED "
                + "ANYWAY: an Explorer can have been created and shown, the current folder set, and a search started "
                + "before this failed. Ask the user what they can see before navigating again.");
        }

        // ------------------------------------------------------------------ list_folders

        /// <summary>
        /// list_folders: the folder tree of every store, or of the store <paramref name="storeId"/> names,
        /// paged in a stable order; every folder carries its folderId and every store its storeId and the
        /// ID of its root folder (the parent for a top-level create_folder).
        /// </summary>
        public FoldersOutcome ListFolders(string? storeId = null, int offset = 0)
        {
            if (offset < 0)
            {
                offset = 0;
            }

            ComStoreDetail? detail = storeId == null ? null : ResolveStoreId(storeId, "store_id");
            ComFolderTree tree = _gateway.Run(s => s.ListFolders(detail?.StoreId, FolderWalkAbsoluteCap));
            return PageFolders(tree, offset, TryGetStoreDetailsForIds());
        }
    }
}
