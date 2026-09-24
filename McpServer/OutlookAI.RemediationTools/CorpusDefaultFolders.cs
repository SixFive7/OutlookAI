using OutlookAI.Core.Com;

namespace OutlookAI.RemediationTools;

/// <summary>
/// How the corpus tool finds a store's default folders - Inbox, Sent Items, Deleted Items, Outbox,
/// Drafts, Junk Email, and the Calendar, Contacts and Tasks a population's undated items live in -
/// WITHOUT EVER ASKING OUTLOOK TO MAKE ONE.
/// <para>
/// <b>Why not <c>Store.GetDefaultFolder</c>.</b> On a PST it does not answer the question it is
/// asked. Measured on OAI-UNINDEXED, 2026-09-24, on PSTs attached with <c>AddStoreEx</c>: asked for
/// Drafts and Junk Email it CREATED them - the bystander gained both during its build - and asked
/// for the Inbox of a PST that has none it returned the PST's nameless non-IPM ROOT, where the
/// bystander's 172 "Inbox" items and its two subfolders then sat, invisible in Outlook's folder
/// tree. The maintainer's decision on Q84 (option (c), the same day) is that a lookup never creates
/// a folder; <see cref="SpecialFolders.Resolve"/> is the product's non-creating answer, and this is
/// the corpus tool's use of it. A folder the build needs and the store lacks is made in the open
/// instead - a stand-in named <see cref="CorpusFolderIds.StandInName"/>, recorded in the manifest so
/// teardown removes it.
/// </para>
/// <para>
/// <b>How.</b> Every mail folder, and every folder of an Exchange mailbox, is
/// <see cref="SpecialFolders.Resolve"/>'s own. The Calendar, Contacts and Tasks of any other store
/// are identified the way <see cref="SpecialFolders"/> identifies Drafts, and by the same table of
/// MS-OXOSFLD section 2.2.3 that names Drafts: a binary entry id on the Inbox -
/// <c>PR_IPM_APPOINTMENT_ENTRYID</c> (0x36D00102), <c>PR_IPM_CONTACT_ENTRYID</c> (0x36D10102) and
/// <c>PR_IPM_TASK_ENTRYID</c> (0x36D40102) - read second off the store object, and opened with
/// <c>GetFolderFromID</c>, which cannot create anything. The Inbox itself is proven present first
/// through <see cref="SpecialFolders.Resolve"/>. Not designated: ABSENT. Not readable: UNREADABLE,
/// never absent, because a census that read a failure as absence would lose items it should count.
/// </para>
/// <para>
/// The three are read here rather than added to <see cref="SpecialFolders"/>: the product has no
/// reader of them, and <c>T1/ReadOnlyFolderLookupTests</c> pins the product's resolver to answer
/// Unreadable for the Calendar, so extending it would change a pinned product answer for the sake
/// of a test fixture. Pure over <see cref="ISpecialFolderStore"/>, so T1 drives it with a fake store.
/// </para>
/// </summary>
public static class CorpusDefaultFolders
{
    /// <summary><c>PR_IPM_APPOINTMENT_ENTRYID</c> (PidTagIpmAppointmentEntryId, 0x36D0, PT_BINARY): the Calendar.</summary>
    public const string AppointmentEntryIdSchema = "http://schemas.microsoft.com/mapi/proptag/0x36D00102";

    /// <summary><c>PR_IPM_CONTACT_ENTRYID</c> (PidTagIpmContactEntryId, 0x36D1, PT_BINARY): Contacts.</summary>
    public const string ContactEntryIdSchema = "http://schemas.microsoft.com/mapi/proptag/0x36D10102";

    /// <summary><c>PR_IPM_TASK_ENTRYID</c> (PidTagIpmTaskEntryId, 0x36D4, PT_BINARY): Tasks.</summary>
    public const string TaskEntryIdSchema = "http://schemas.microsoft.com/mapi/proptag/0x36D40102";

    /// <summary>
    /// The designation property of a folder this class reads itself - the Calendar, Contacts or Tasks -
    /// or null for every other folder, which <see cref="SpecialFolders.Resolve"/> answers.
    /// </summary>
    public static string? DesignationSchema(int folderId) => folderId switch
    {
        CorpusItemKinds.CalendarFolderId => AppointmentEntryIdSchema,
        CorpusItemKinds.ContactsFolderId => ContactEntryIdSchema,
        CorpusItemKinds.TasksFolderId => TaskEntryIdSchema,
        _ => null,
    };

    /// <summary>
    /// Resolves ONE default folder without creating it. <paramref name="folder"/> is set only for
    /// <see cref="OutlookComSession.DefaultFolderResolution.Resolved"/>, and the caller releases it.
    /// </summary>
    public static OutlookComSession.DefaultFolderResolution Resolve(ISpecialFolderStore store, int folderId, out object? folder)
    {
        ArgumentNullException.ThrowIfNull(store);
        folder = null;
        string? schema = DesignationSchema(folderId);
        if (schema == null || SpecialFolders.IsExchangeStore(store.ExchangeStoreType))
        {
            return SpecialFolders.Resolve(store, folderId, out folder, out _);
        }

        OutlookComSession.DefaultFolderResolution primary =
            SpecialFolders.Resolve(store, SpecialFolders.OlFolderInbox, out object? inbox, out _);
        if (primary == OutlookComSession.DefaultFolderResolution.Resolved)
        {
            try
            {
                primary = OpenDesignated(store, store.ReadFolderProperty(inbox!, schema), out folder);
            }
            finally
            {
                store.Release(inbox);
            }

            if (primary == OutlookComSession.DefaultFolderResolution.Resolved)
            {
                return primary;
            }
        }

        // A store with no Inbox has nothing designated on one (primary stays Absent), and one whose
        // Inbox would not open is Unreadable. The store object is read second, as SpecialFolders
        // reads Drafts; a hit there resolves, anything else leaves the Inbox's answer standing.
        if (OpenDesignated(store, store.ReadStoreProperty(schema), out folder) == OutlookComSession.DefaultFolderResolution.Resolved)
        {
            return OutlookComSession.DefaultFolderResolution.Resolved;
        }

        return primary;
    }

    private static OutlookComSession.DefaultFolderResolution OpenDesignated(
        ISpecialFolderStore store, PropertyRead designation, out object? folder)
    {
        folder = null;
        if (designation.Status == PropertyReadStatus.NotFound)
        {
            return OutlookComSession.DefaultFolderResolution.Absent;
        }

        if (designation.Status != PropertyReadStatus.Found)
        {
            return OutlookComSession.DefaultFolderResolution.Unreadable;
        }

        DesignatedEntryId entryId = SpecialFolders.ReadEntryId(designation.Value);
        if (entryId.Kind == DesignatedEntryIdKind.None)
        {
            return OutlookComSession.DefaultFolderResolution.Absent;
        }

        if (entryId.Kind != DesignatedEntryIdKind.EntryId)
        {
            return OutlookComSession.DefaultFolderResolution.Unreadable;
        }

        PropertyReadStatus opened = store.OpenFolder(entryId.Hex!, out object? candidate);
        if (opened == PropertyReadStatus.Found && candidate != null)
        {
            folder = candidate;
            return OutlookComSession.DefaultFolderResolution.Resolved;
        }

        store.Release(candidate);
        return opened == PropertyReadStatus.Failed
            ? OutlookComSession.DefaultFolderResolution.Unreadable
            : OutlookComSession.DefaultFolderResolution.Absent;
    }
}

/// <summary>
/// Whether a folder is one a person can SEE: a named folder somewhere under the store's root folder
/// - the top of the tree Outlook shows. Pure; the COM half collects the facts.
/// <para>
/// It exists because an item can be placed in a folder that exists, answers its table, and is shown
/// nowhere: on OAI-UNINDEXED, 2026-09-24, <c>GetDefaultFolder(olFolderInbox)</c> on a PST attached
/// with <c>AddStoreEx</c> returned the PST's non-IPM ROOT - nameless, and not below the root folder
/// Outlook draws - and 172 bystander items were built into it while the placement probe printed
/// <c>target= landedIn=</c> and said VERIFIED. Every folder the build writes into now has to pass this.
/// </para>
/// </summary>
public static class CorpusFolderVisibility
{
    /// <summary>
    /// Whether the folder is visible.
    /// </summary>
    /// <param name="name">The folder's name. A folder with none is not one anybody can pick out.</param>
    /// <param name="ancestorEntryIds">
    /// The EntryIDs of the folder's ancestors, nearest first, as far as <c>Folder.Parent</c> could be
    /// followed. A null entry ends the chain: the parent was not a folder, or would not say.
    /// </param>
    /// <param name="rootFolderEntryId">The EntryID of <c>Store.GetRootFolder()</c> - the top of the visible tree.</param>
    public static bool IsVisible(string? name, IReadOnlyList<string?> ancestorEntryIds, string? rootFolderEntryId)
    {
        ArgumentNullException.ThrowIfNull(ancestorEntryIds);
        if (string.IsNullOrWhiteSpace(name) || string.IsNullOrEmpty(rootFolderEntryId))
        {
            return false;
        }

        foreach (string? id in ancestorEntryIds)
        {
            if (id == null)
            {
                return false;
            }

            if (string.Equals(id, rootFolderEntryId, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }
}
