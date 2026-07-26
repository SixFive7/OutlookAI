using System.Runtime.InteropServices;
using OutlookAI.Core.Com;

namespace OutlookAI.RemediationTools;

/// <summary>
/// COM mailbox operations for the incident remediation, modeled 1:1 on the tested
/// LiveOutlookTestMailer patterns (short-lived STA thread per operation, explicit
/// Release, GetTable LIKE prefilter + ordinal re-check, EntryID-addressed deletes).
/// S4 discipline: methods return counts/EntryIDs/booleans; business-store subjects
/// never leave this class except for the designated test hub (classification needs
/// the hub subjects for the deletion-log cross-check).
/// </summary>
public static class ComMailbox
{
    /// <summary>Default-folder ids: Drafts, Inbox, Sent Items, Deleted Items (3 LAST - purge order).</summary>
    public static readonly int[] SweepFolderIds = { 16, 6, 5, 3 };

    /// <summary>Hub-only sweep set including the designated Archive folder (39). Never business stores.</summary>
    public static readonly int[] HubSweepFolderIdsWithArchive = { 16, 6, 5, 39, 3 };

    /// <summary>One folder's remediation counts (audit view).</summary>
    public sealed record FolderCounts(string Store, int FolderId, string FolderName, int Total, int Tagged);

    /// <summary>One Deleted Items item snapshot for classification/dedupe.</summary>
    public sealed record DeletedItemInfo(
        string EntryId,
        string? Subject,
        bool Tagged,
        string? SenderSmtp,
        bool ReceivedByPresent,
        string? InternetMessageId);

    /// <summary>Read-only: total + exact ordinal-tagged counts for one store's folders.</summary>
    public static List<FolderCounts> CountStoreFolders(string storeDisplayName, int[] folderIds)
    {
        return RunSta(() =>
        {
            dynamic app = CreateOutlookApplication();
            dynamic? ns = null;
            dynamic? stores = null;
            dynamic? store = null;
            var results = new List<FolderCounts>();
            try
            {
                ns = app.GetNamespace("MAPI");
                stores = ns.Stores;
                store = FindStore(stores, storeDisplayName)
                    ?? throw new InvalidOperationException($"Store not found: {storeDisplayName}");
                foreach (int folderId in folderIds)
                {
                    dynamic? folder = null;
                    try
                    {
                        try
                        {
                            folder = store.GetDefaultFolder(folderId);
                        }
                        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                        {
                            continue; // store without that default folder
                        }

                        string name = (string)folder.Name;
                        dynamic? items = null;
                        int total;
                        try
                        {
                            items = folder.Items;
                            total = (int)items.Count;
                        }
                        finally
                        {
                            Release(items);
                        }

                        int tagged = CountTaggedExact(folder);
                        results.Add(new FolderCounts(storeDisplayName, folderId, name, total, tagged));
                    }
                    finally
                    {
                        Release(folder);
                    }
                }

                return results;
            }
            finally
            {
                Release(store);
                Release(stores);
                Release(ns);
                Release(app);
            }
        });
    }

    /// <summary>Read-only: snapshots every item in the store's Deleted Items folder.</summary>
    public static List<DeletedItemInfo> ListDeletedItems(string storeDisplayName, bool withMessageIds, bool withSender)
    {
        return RunSta(() =>
        {
            dynamic app = CreateOutlookApplication();
            dynamic? ns = null;
            dynamic? stores = null;
            dynamic? store = null;
            dynamic? folder = null;
            dynamic? items = null;
            var results = new List<DeletedItemInfo>();
            try
            {
                ns = app.GetNamespace("MAPI");
                stores = ns.Stores;
                store = FindStore(stores, storeDisplayName)
                    ?? throw new InvalidOperationException($"Store not found: {storeDisplayName}");
                folder = store.GetDefaultFolder(3);
                items = folder.Items;
                int count = (int)items.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic? item = null;
                    try
                    {
                        item = items[i];
                        string entryId = (string)item.EntryID;
                        string? subject = TryGetComString(item, "Subject");
                        bool tagged = RemediationRules.IsTagged(subject);
                        string? sender = null;
                        bool receivedBy = false;
                        string? messageId = null;
                        if (withSender)
                        {
                            sender = TryGetSenderSmtp(item);
                            string? receivedByName = TryGetComString(item, "ReceivedByName");
                            receivedBy = !string.IsNullOrEmpty(receivedByName);
                        }

                        if (withMessageIds)
                        {
                            messageId = TryGetMapiString(item, "http://schemas.microsoft.com/mapi/proptag/0x1035001F");
                        }

                        results.Add(new DeletedItemInfo(entryId, subject, tagged, sender, receivedBy, messageId));
                    }
                    catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                    {
                        // Unreadable row: surfaces as a count mismatch for the operator.
                    }
                    finally
                    {
                        Release(item);
                    }
                }

                return results;
            }
            finally
            {
                Release(items);
                Release(folder);
                Release(store);
                Release(stores);
                Release(ns);
                Release(app);
            }
        }, TimeSpan.FromMinutes(10));
    }

    /// <summary>
    /// Read-only: collects the trimmed PR_INTERNET_MESSAGE_ID of every item currently
    /// in the store's Inbox (the dedupe twin set), plus the count of items without one.
    /// </summary>
    public static (HashSet<string> MessageIds, int Total, int WithoutMessageId) CollectInboxMessageIds(string storeDisplayName)
    {
        return RunSta(() =>
        {
            dynamic app = CreateOutlookApplication();
            dynamic? ns = null;
            dynamic? stores = null;
            dynamic? store = null;
            dynamic? folder = null;
            dynamic? items = null;
            var set = new HashSet<string>(StringComparer.Ordinal);
            int total = 0;
            int withoutId = 0;
            try
            {
                ns = app.GetNamespace("MAPI");
                stores = ns.Stores;
                store = FindStore(stores, storeDisplayName)
                    ?? throw new InvalidOperationException($"Store not found: {storeDisplayName}");
                folder = store.GetDefaultFolder(6);
                items = folder.Items;
                int count = (int)items.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic? item = null;
                    try
                    {
                        item = items[i];
                        total++;
                        string? messageId = RemediationRules.NormalizeMessageId(
                            TryGetMapiString(item, "http://schemas.microsoft.com/mapi/proptag/0x1035001F"));
                        if (messageId == null)
                        {
                            withoutId++;
                        }
                        else
                        {
                            set.Add(messageId);
                        }
                    }
                    catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                    {
                        withoutId++;
                    }
                    finally
                    {
                        Release(item);
                    }
                }

                return (set, total, withoutId);
            }
            finally
            {
                Release(items);
                Release(folder);
                Release(store);
                Release(stores);
                Release(ns);
                Release(app);
            }
        }, TimeSpan.FromMinutes(10));
    }

    /// <summary>Per-folder purge outcome.</summary>
    public sealed record PurgeFolderResult(string Store, int FolderId, string FolderName, int Matched, int Deleted, int Failed);

    /// <summary>
    /// One purge pass over the store's folders: GetTable LIKE prefilter collects
    /// candidate EntryIDs, each item is re-opened and re-checked with the ORDINAL
    /// full-tag predicate at delete time, then Delete()d (folders 16/6/5 soft-move to
    /// Deleted Items; folder 3 hard-deletes into the dumpster). Dry-run counts only.
    /// </summary>
    public static List<PurgeFolderResult> PurgeTaggedPass(string storeDisplayName, int[] folderIds, bool execute)
    {
        return RunSta(() =>
        {
            dynamic app = CreateOutlookApplication();
            dynamic? ns = null;
            dynamic? stores = null;
            dynamic? store = null;
            var results = new List<PurgeFolderResult>();
            try
            {
                ns = app.GetNamespace("MAPI");
                stores = ns.Stores;
                store = FindStore(stores, storeDisplayName)
                    ?? throw new InvalidOperationException($"Store not found: {storeDisplayName}");
                string storeId = (string)store.StoreID;
                foreach (int folderId in folderIds)
                {
                    dynamic? folder = null;
                    try
                    {
                        try
                        {
                            folder = store.GetDefaultFolder(folderId);
                        }
                        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                        {
                            continue;
                        }

                        string name = (string)folder.Name;
                        List<string> candidates = CollectTaggedEntryIds(folder);
                        if (!execute)
                        {
                            results.Add(new PurgeFolderResult(storeDisplayName, folderId, name, candidates.Count, 0, 0));
                            continue;
                        }

                        int deleted = 0;
                        int failed = 0;
                        foreach (string entryId in candidates)
                        {
                            dynamic? item = null;
                            try
                            {
                                try
                                {
                                    item = ns.GetItemFromID(entryId, storeId);
                                }
                                catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                                {
                                    continue; // already gone (e.g. deleted by an earlier pass)
                                }

                                string? subject = TryGetComString(item, "Subject");
                                if (!RemediationRules.IsTagged(subject))
                                {
                                    continue; // ordinal re-check failed - never delete on a LIKE match alone
                                }

                                item.Delete();
                                deleted++;
                            }
                            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                            {
                                failed++;
                            }
                            finally
                            {
                                Release(item);
                            }
                        }

                        results.Add(new PurgeFolderResult(storeDisplayName, folderId, name, candidates.Count, deleted, failed));
                    }
                    finally
                    {
                        Release(folder);
                    }
                }

                return results;
            }
            finally
            {
                Release(store);
                Release(stores);
                Release(ns);
                Release(app);
            }
        }, TimeSpan.FromMinutes(10));
    }

    /// <summary>Read-only: exact ordinal-tagged count across the store's folders (stable-zero polling).</summary>
    public static int CountTaggedInFolders(string storeDisplayName, int[] folderIds)
    {
        List<FolderCounts> counts = CountStoreFolders(storeDisplayName, folderIds);
        return counts.Sum(c => c.Tagged);
    }

    /// <summary>Per-item dedupe outcome (step 3).</summary>
    public sealed record DedupeItemResult(string EntryId, RemediationRules.DedupeDecision Decision, bool Deleted, string? Error);

    /// <summary>
    /// Deletes ONE verified duplicate from the store's Deleted Items: re-opens the
    /// item by EntryID, re-reads subject + PR_INTERNET_MESSAGE_ID, re-runs
    /// <see cref="RemediationRules.DecideDuplicateDelete"/> against the CURRENT Inbox
    /// twin set, verifies the item still sits directly in Deleted Items, and only
    /// then Delete()s (hard delete into the dumpster - the 14-day natural undo).
    /// </summary>
    public static DedupeItemResult DeleteVerifiedDuplicate(
        string storeDisplayName,
        string entryId,
        IReadOnlySet<string> inboxMessageIds,
        bool execute)
    {
        return RunSta(() =>
        {
            dynamic app = CreateOutlookApplication();
            dynamic? ns = null;
            dynamic? stores = null;
            dynamic? store = null;
            dynamic? deletedItems = null;
            dynamic? item = null;
            dynamic? parent = null;
            try
            {
                ns = app.GetNamespace("MAPI");
                stores = ns.Stores;
                store = FindStore(stores, storeDisplayName)
                    ?? throw new InvalidOperationException($"Store not found: {storeDisplayName}");
                deletedItems = store.GetDefaultFolder(3);
                string deletedItemsEntryId = (string)deletedItems.EntryID;
                try
                {
                    item = ns.GetItemFromID(entryId, (string)store.StoreID);
                }
                catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                {
                    return new DedupeItemResult(entryId, RemediationRules.DedupeDecision.SkipNoInboxTwin, false,
                        "ItemNotFound: EntryID no longer resolves.");
                }

                string? subject = TryGetComString(item, "Subject");
                string? messageId = TryGetMapiString(item, "http://schemas.microsoft.com/mapi/proptag/0x1035001F");
                RemediationRules.DedupeDecision decision =
                    RemediationRules.DecideDuplicateDelete(subject, messageId, inboxMessageIds);
                if (decision != RemediationRules.DedupeDecision.Delete)
                {
                    return new DedupeItemResult(entryId, decision, false, null);
                }

                bool inDeletedItems = false;
                try
                {
                    parent = item.Parent;
                    inDeletedItems = parent != null && string.Equals(
                        (string)((dynamic)parent!).EntryID, deletedItemsEntryId, StringComparison.OrdinalIgnoreCase);
                }
                catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                {
                }

                if (!inDeletedItems)
                {
                    return new DedupeItemResult(entryId, decision, false,
                        "NotInDeletedItems: the item is no longer directly in Deleted Items - kept.");
                }

                if (!execute)
                {
                    return new DedupeItemResult(entryId, decision, false, null);
                }

                item.Delete(); // from Deleted Items = hard delete into the dumpster
                return new DedupeItemResult(entryId, decision, true, null);
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                return new DedupeItemResult(entryId, RemediationRules.DedupeDecision.Delete, false,
                    $"ComFailure: {ex.GetType().Name}");
            }
            finally
            {
                Release(parent);
                Release(item);
                Release(deletedItems);
                Release(store);
                Release(stores);
                Release(ns);
                Release(app);
            }
        });
    }

    /// <summary>Read-only: the display name of one default folder (exact move target names).</summary>
    public static string GetDefaultFolderName(string storeDisplayName, int folderId)
    {
        return RunSta(() =>
        {
            dynamic app = CreateOutlookApplication();
            dynamic? ns = null;
            dynamic? stores = null;
            dynamic? store = null;
            dynamic? folder = null;
            try
            {
                ns = app.GetNamespace("MAPI");
                stores = ns.Stores;
                store = FindStore(stores, storeDisplayName)
                    ?? throw new InvalidOperationException($"Store not found: {storeDisplayName}");
                folder = store.GetDefaultFolder(folderId);
                return (string)folder.Name;
            }
            finally
            {
                Release(folder);
                Release(store);
                Release(stores);
                Release(ns);
                Release(app);
            }
        });
    }

    /// <summary>Exact ordinal-tagged count of one folder (LIKE prefilter + ordinal re-check).</summary>
    private static int CountTaggedExact(dynamic folder)
    {
        List<(string EntryId, string? Subject)> rows = CollectTaggedRows(folder);
        return rows.Count(r => RemediationRules.IsTagged(r.Subject));
    }

    /// <summary>Candidate EntryIDs whose CURRENT subject passes the ordinal full-tag check.</summary>
    private static List<string> CollectTaggedEntryIds(dynamic folder)
    {
        List<(string EntryId, string? Subject)> rows = CollectTaggedRows(folder);
        return rows
            .Where(r => RemediationRules.IsTagged(r.Subject))
            .Select(r => r.EntryId)
            .ToList();
    }

    /// <summary>
    /// GetTable with the bracket-free LIKE fragment (fast on huge folders), returning
    /// EntryID + Subject per row; falls back to an Items.Restrict walk when GetTable
    /// fails (the proven CountTaggedArtifacts pattern).
    /// </summary>
    private static List<(string EntryId, string? Subject)> CollectTaggedRows(dynamic folder)
    {
        var rows = new List<(string, string?)>();
        string filter = "@SQL=\"urn:schemas:httpmail:subject\" LIKE '%" + RemediationRules.DaslCountFragment + "%'";
        dynamic? table = null;
        try
        {
            try
            {
                table = folder.GetTable(filter);
                int entryIdIndex = FindTableColumn(table, "EntryID");
                int subjectIndex = FindTableColumn(table, "Subject");
                if (entryIdIndex >= 0)
                {
                    while (!(bool)table.EndOfTable)
                    {
                        dynamic? row = null;
                        try
                        {
                            row = table.GetNextRow();
                            object[] values = (object[])row!.GetValues();
                            if (entryIdIndex < values.Length && values[entryIdIndex] is string entryId && entryId.Length > 0)
                            {
                                string? subject = subjectIndex >= 0 && subjectIndex < values.Length
                                    ? values[subjectIndex] as string
                                    : null;
                                rows.Add((entryId, subject));
                            }
                        }
                        finally
                        {
                            Release(row);
                        }
                    }

                    return rows;
                }
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                // Fall back to Items.Restrict below.
            }

            rows.Clear();
            dynamic? items = null;
            dynamic? restricted = null;
            try
            {
                items = folder.Items;
                restricted = items.Restrict(filter);
                int count = (int)restricted.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic? item = null;
                    try
                    {
                        item = restricted[i];
                        rows.Add(((string)item.EntryID, TryGetComString(item, "Subject")));
                    }
                    catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                    {
                    }
                    finally
                    {
                        Release(item);
                    }
                }
            }
            finally
            {
                Release(restricted);
                Release(items);
            }

            return rows;
        }
        finally
        {
            Release(table);
        }
    }

    private static string? TryGetSenderSmtp(dynamic item)
    {
        // PR_SENDER_SMTP_ADDRESS, then PR_SENT_REPRESENTING_SMTP_ADDRESS, then the
        // raw SenderEmailAddress when it is SMTP-typed (Phase-5 readback pattern).
        string? smtp = TryGetMapiString(item, "http://schemas.microsoft.com/mapi/proptag/0x5D01001F")
            ?? TryGetMapiString(item, "http://schemas.microsoft.com/mapi/proptag/0x5D02001F");
        if (!string.IsNullOrWhiteSpace(smtp))
        {
            return smtp;
        }

        string? type = TryGetComString(item, "SenderEmailType");
        if (string.Equals(type, "SMTP", StringComparison.OrdinalIgnoreCase))
        {
            return TryGetComString(item, "SenderEmailAddress");
        }

        return null;
    }

    private static string? TryGetMapiString(dynamic item, string schemaName)
    {
        dynamic? accessor = null;
        try
        {
            accessor = item.PropertyAccessor;
            object? value = accessor.GetProperty(schemaName);
            return value as string;
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
        {
            return null;
        }
        finally
        {
            Release(accessor);
        }
    }

    /// <summary>
    /// Late-bound COM string property read that never throws on COM failure (the
    /// LiveOutlookTestMailer inline-try pattern, centralized; reflection instead of a
    /// lambda because lambdas cannot be arguments of dynamic dispatch - CS1977).
    /// </summary>
    private static string? TryGetComString(object? comObject, string propertyName)
    {
        if (comObject == null)
        {
            return null;
        }

        try
        {
            object? value = comObject.GetType().InvokeMember(
                propertyName,
                System.Reflection.BindingFlags.GetProperty
                    | System.Reflection.BindingFlags.Public
                    | System.Reflection.BindingFlags.Instance,
                null,
                comObject,
                null);
            return value as string;
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(
            ex is System.Reflection.TargetInvocationException tie && tie.InnerException != null ? tie.InnerException : ex))
        {
            return null;
        }
    }

    private static int FindTableColumn(dynamic table, string columnName)
    {
        dynamic? columns = null;
        try
        {
            columns = table.Columns;
            int count = (int)columns.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic? column = null;
                try
                {
                    column = columns[i];
                    if (string.Equals((string)column.Name, columnName, StringComparison.OrdinalIgnoreCase))
                    {
                        return i - 1;
                    }
                }
                finally
                {
                    Release(column);
                }
            }

            return -1;
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
        {
            return -1;
        }
        finally
        {
            Release(columns);
        }
    }

    private static dynamic? FindStore(dynamic stores, string storeDisplayName)
    {
        int storeCount = (int)stores.Count;
        for (int i = 1; i <= storeCount; i++)
        {
            dynamic candidate = stores[i];
            string? name = TryGetComString(candidate, "DisplayName");
            if (string.Equals(name, storeDisplayName, StringComparison.OrdinalIgnoreCase))
            {
                return candidate;
            }

            Release(candidate);
        }

        return null;
    }

    private static dynamic CreateOutlookApplication()
    {
        Type progIdType = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException("Outlook.Application ProgID is not registered.");
        return Activator.CreateInstance(progIdType)
            ?? throw new InvalidOperationException("Failed to create Outlook.Application.");
    }

    private static T RunSta<T>(Func<T> work, TimeSpan? timeout = null)
    {
        T result = default!;
        Exception? failure = null;
        var thread = new Thread(() =>
        {
            try
            {
                result = work();
            }
            catch (Exception ex)
            {
                failure = ex;
            }
        })
        {
            IsBackground = true,
            Name = "OutlookAI.Remediation.Sta",
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        if (!thread.Join(timeout ?? TimeSpan.FromMinutes(3)))
        {
            throw new TimeoutException("Remediation STA operation timed out.");
        }

        if (failure != null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(failure).Throw();
        }

        return result;
    }

    private static void Release(object? comObject)
    {
        if (comObject != null && Marshal.IsComObject(comObject))
        {
            Marshal.ReleaseComObject(comObject);
        }
    }
}
