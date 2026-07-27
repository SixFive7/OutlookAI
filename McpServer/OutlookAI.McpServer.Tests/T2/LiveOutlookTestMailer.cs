using System.Runtime.InteropServices;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// TEST-ONLY Outlook COM helper for the D20 round-trip grant: sends a tagged mail from
/// the designated test hub TO ITSELF, and deletes ONLY artifacts that carry both the
/// [OutlookAI-McpTest] tag and this run's unique marker (S3 double-match rule).
/// Deliberately lives in the test project - the shipped Phase-2 server has zero
/// send/delete surface (S1). Each call runs on its own short-lived STA thread.
/// </summary>
public static class LiveOutlookTestMailer
{
    /// <summary>The S3 subject tag every test artifact must carry.</summary>
    public const string SubjectTag = "[OutlookAI-McpTest]";

    /// <summary>
    /// Name prefix every TEST FOLDER must carry (D39 move tests). Folder cleanup
    /// matches this with ordinal Contains via the tested helpers only - never shell
    /// patterns (the 7d standing rule).
    /// </summary>
    public const string TestFolderNamePrefix = "OutlookAI-McpTest-Folder";

    /// <summary>
    /// Default-folder ids swept for tagged artifacts: Drafts, Inbox, Sent Items, the
    /// SYNC ISSUES subtree, then Deleted Items LAST so the final pass purges what
    /// Delete() moved there.
    /// <para>
    /// The Sync Issues subtree (20 Sync Issues, 19 Conflicts, 21 Local Failures,
    /// 22 Server Failures) was added after a tagged test item was found stranded in
    /// the hub's Local Failures folder: a cached-Exchange store can file a copy of a
    /// test artifact there on its own, and nothing swept it. Those folders hold
    /// hundreds of items, not the ~100k an archive holds, so counting them is cheap.
    /// </para>
    /// <para>
    /// ⚠ Deliberately WITHOUT the Archive folder: business-store archives hold ~100k
    /// items and a LIKE count over them takes minutes - adding 39 here made the
    /// cross-account sweeps time out (live-bitten 2026-07-26). Hub-scoped D39
    /// cleanups pass <see cref="HubSweepFolderIdsWithArchive"/> explicitly instead.
    /// </para>
    /// </summary>
    private static readonly int[] SweepFolderIds = { 16, 6, 5, 20, 19, 21, 22, 3 };

    /// <summary>
    /// Sweep set for the TINY test hub only (D39): includes the designated Archive
    /// folder (39) so archive_mail artifacts are counted and purged. Never use
    /// against business stores (their archives are huge - see SweepFolderIds).
    /// </summary>
    public static readonly int[] HubSweepFolderIdsWithArchive = { 16, 6, 5, 39, 20, 19, 21, 22, 3 };

    /// <summary>
    /// The Sync Issues subtree ids, exposed so a targeted purge can name them
    /// (olFolderSyncIssues 20, olFolderConflicts 19, olFolderLocalFailures 21,
    /// olFolderServerFailures 22).
    /// </summary>
    public static readonly int[] SyncIssuesFolderIds = { 20, 19, 21, 22 };

    /// <summary>
    /// Sends a mail from <paramref name="smtpAddress"/> to itself (refuses to run when
    /// that account is not in the profile - the D20 grant is telefonie-to-telefonie
    /// only). Returns the UTC send timestamp.
    /// </summary>
    public static DateTime SendSelfMail(string smtpAddress, string subject, string body, string? attachmentPath)
    {
        if (!subject.Contains(SubjectTag, StringComparison.Ordinal))
        {
            throw new ArgumentException($"Test mail subject must carry the {SubjectTag} tag (S3).", nameof(subject));
        }

        return RunSta(() =>
        {
            dynamic app = CreateOutlookApplication();
            dynamic? mail = null;
            dynamic? session = null;
            dynamic? accounts = null;
            try
            {
                session = app.Session;
                accounts = session.Accounts;
                object? sendingAccount = null;
                int count = accounts.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic account = accounts[i];
                    string? accountSmtp = null;
                    try
                    {
                        accountSmtp = (string?)account.SmtpAddress;
                    }
                    catch (COMException)
                    {
                    }

                    if (string.Equals(accountSmtp, smtpAddress, StringComparison.OrdinalIgnoreCase))
                    {
                        sendingAccount = account;
                        break;
                    }

                    Release(account);
                }

                if (sendingAccount == null)
                {
                    throw new InvalidOperationException(
                        "Test-hub account not found in the profile; refusing to send from any other account (D20).");
                }

                try
                {
                    mail = app.CreateItem(0); // olMailItem
                    mail.Subject = subject;
                    mail.Body = body;
                    mail.To = smtpAddress;
                    if (attachmentPath != null)
                    {
                        dynamic attachments = mail.Attachments;
                        try
                        {
                            attachments.Add(attachmentPath);
                        }
                        finally
                        {
                            Release(attachments);
                        }
                    }

                    // D20/v3.MD section 3: SendUsingAccount takes the Account OBJECT and
                    // must be set BEFORE Send. ⚠ Phase-4 live finding: it is a
                    // PROPERTYPUTREF property - a plain dynamic assignment SILENTLY
                    // NO-OPS (the Phase-2/3 seeds actually went out from the DEFAULT
                    // account because of this). Invoke the putref accessor explicitly,
                    // then HARD-VERIFY the identity before sending: a mismatch aborts
                    // the send instead of violating the hub-only grant.
                    ((object)mail).GetType().InvokeMember(
                        "SendUsingAccount",
                        System.Reflection.BindingFlags.PutRefDispProperty
                            | System.Reflection.BindingFlags.Public
                            | System.Reflection.BindingFlags.Instance,
                        null,
                        (object)mail,
                        new[] { sendingAccount });

                    string? effectiveSender = null;
                    dynamic? sendUsing = null;
                    try
                    {
                        sendUsing = mail.SendUsingAccount;
                        effectiveSender = sendUsing != null ? (string?)sendUsing.SmtpAddress : null;
                    }
                    finally
                    {
                        Release(sendUsing);
                    }

                    if (!string.Equals(effectiveSender, smtpAddress, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new InvalidOperationException(
                            "SendUsingAccount did not stick (would send from the default account) - refusing to send (D20).");
                    }

                    mail.Send();
                }
                finally
                {
                    Release(sendingAccount);
                }

                return DateTime.UtcNow;
            }
            finally
            {
                Release(mail);
                Release(accounts);
                Release(session);
                Release(app);
            }
        });
    }

    /// <summary>
    /// Deletes every item in the store's default sweep folders (Drafts, Inbox, Sent
    /// Items, Deleted Items - hub archive coverage via
    /// <see cref="HubSweepFolderIdsWithArchive"/>) whose subject contains BOTH the tag
    /// and <paramref name="uniqueMarker"/> (S3: only artifacts this run created). Two
    /// passes: Delete() moves to Deleted Items, the final pass on folder 3 removes
    /// them for good. Returns the total deleted.
    /// </summary>
    public static int DeleteTaggedArtifacts(string storeDisplayName, string uniqueMarker, int[]? folderIds = null)
    {
        if (string.IsNullOrWhiteSpace(uniqueMarker) || uniqueMarker.Length < 12)
        {
            throw new ArgumentException("Marker too weak for a safe delete filter (S3).", nameof(uniqueMarker));
        }

        int[] folders = folderIds ?? SweepFolderIds;
        return RunSta(() =>
        {
            dynamic app = CreateOutlookApplication();
            dynamic? ns = null;
            dynamic? stores = null;
            int deleted = 0;
            try
            {
                ns = app.GetNamespace("MAPI");
                stores = ns.Stores;
                dynamic? store = FindStore(stores, storeDisplayName)
                    ?? throw new InvalidOperationException("Test-hub store not found for cleanup.");
                try
                {
                    foreach (int folderId in folders)
                    {
                        deleted += DeleteMatchingInFolder(store, folderId, uniqueMarker);
                    }
                }
                finally
                {
                    Release(store);
                }

                return deleted;
            }
            finally
            {
                Release(stores);
                Release(ns);
                Release(app);
            }
        });
    }

    /// <summary>
    /// Cleanup for artifacts of a SELF-SEND round trip (D20): the Inbox and Sent
    /// copies of a just-sent mail materialize asynchronously and can arrive AFTER a
    /// one-shot cleanup pass (live-observed: an Inbox copy appeared after delete +
    /// count both reported zero). Loops delete+count until the count stays zero for
    /// <paramref name="stableFor"/>, throwing when <paramref name="window"/> expires.
    /// Returns the total number of artifacts deleted.
    /// </summary>
    public static int DeleteTaggedArtifactsUntilStableZero(
        string storeDisplayName,
        string uniqueMarker,
        TimeSpan? window = null,
        TimeSpan? stableFor = null,
        int[]? folderIds = null)
    {
        TimeSpan totalWindow = window ?? TimeSpan.FromSeconds(120);
        TimeSpan requiredStable = stableFor ?? TimeSpan.FromSeconds(10);
        DateTime deadline = DateTime.UtcNow + totalWindow;
        DateTime? zeroSince = null;
        int totalDeleted = 0;
        while (DateTime.UtcNow < deadline)
        {
            int remaining = CountTaggedArtifacts(storeDisplayName, uniqueMarker, folderIds);
            if (remaining > 0)
            {
                totalDeleted += DeleteTaggedArtifacts(storeDisplayName, uniqueMarker, folderIds);
                zeroSince = null;
                continue;
            }

            zeroSince ??= DateTime.UtcNow;
            if (DateTime.UtcNow - zeroSince.Value >= requiredStable)
            {
                return totalDeleted;
            }

            Thread.Sleep(2000);
        }

        throw new TimeoutException(
            "Tagged artifacts kept (re)appearing for the whole cleanup window - manual check required (S3).");
    }

    /// <summary>
    /// Deletes ONE item by EntryID from the given store, refusing unless its subject
    /// carries BOTH the tag and <paramref name="uniqueMarker"/> (the S3 double-match:
    /// created-this-run id AND tag). Delete() moves it to the store's Deleted Items;
    /// call <see cref="DeleteTaggedArtifacts"/> (or rely on the caller's final cleanup)
    /// for the purge pass. Returns true when the item was found and deleted.
    /// </summary>
    public static bool DeleteItemByEntryId(string storeDisplayName, string entryIdHex, string uniqueMarker)
    {
        if (string.IsNullOrWhiteSpace(entryIdHex))
        {
            throw new ArgumentException("EntryID required.", nameof(entryIdHex));
        }

        if (string.IsNullOrWhiteSpace(uniqueMarker) || uniqueMarker.Length < 12)
        {
            throw new ArgumentException("Marker too weak for a safe delete filter (S3).", nameof(uniqueMarker));
        }

        return RunSta(() =>
        {
            dynamic app = CreateOutlookApplication();
            dynamic? ns = null;
            dynamic? stores = null;
            dynamic? store = null;
            dynamic? item = null;
            try
            {
                ns = app.GetNamespace("MAPI");
                stores = ns.Stores;
                store = FindStore(stores, storeDisplayName)
                    ?? throw new InvalidOperationException("Store not found for EntryID delete.");
                string storeId = (string)store.StoreID;
                try
                {
                    item = ns.GetItemFromID(entryIdHex, storeId);
                }
                catch (COMException)
                {
                    return false; // already gone
                }

                string? subject = null;
                try
                {
                    subject = (string?)item.Subject;
                }
                catch (COMException)
                {
                }

                if (subject == null
                    || !subject.Contains(SubjectTag, StringComparison.Ordinal)
                    || !subject.Contains(uniqueMarker, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "Refusing to delete: the item's subject does not carry both the test tag and this run's marker (S3).");
                }

                item.Delete();
                return true;
            }
            finally
            {
                Release(item);
                Release(store);
                Release(stores);
                Release(ns);
                Release(app);
            }
        });
    }

    /// <summary>
    /// Appends text to the BODY of one draft this run created (Phase-5 negative test:
    /// a modified draft must invalidate its send confirm-token). Refuses unless the
    /// item's subject carries BOTH the tag and <paramref name="uniqueMarker"/> (S3
    /// double-match - only artifacts of this run are ever touched) and the item is
    /// still unsent. Saves the draft after the change.
    /// </summary>
    public static void AppendToDraftBody(string storeDisplayName, string entryIdHex, string uniqueMarker, string textToAppend)
    {
        if (string.IsNullOrWhiteSpace(entryIdHex))
        {
            throw new ArgumentException("EntryID required.", nameof(entryIdHex));
        }

        if (string.IsNullOrWhiteSpace(uniqueMarker) || uniqueMarker.Length < 12)
        {
            throw new ArgumentException("Marker too weak for a safe modify filter (S3).", nameof(uniqueMarker));
        }

        RunSta<object?>(() =>
        {
            dynamic app = CreateOutlookApplication();
            dynamic? ns = null;
            dynamic? stores = null;
            dynamic? store = null;
            dynamic? item = null;
            try
            {
                ns = app.GetNamespace("MAPI");
                stores = ns.Stores;
                store = FindStore(stores, storeDisplayName)
                    ?? throw new InvalidOperationException("Store not found for draft modification.");
                item = ns.GetItemFromID(entryIdHex, (string)store.StoreID);

                string? subject = null;
                try
                {
                    subject = (string?)item.Subject;
                }
                catch (COMException)
                {
                }

                if (subject == null
                    || !subject.Contains(SubjectTag, StringComparison.Ordinal)
                    || !subject.Contains(uniqueMarker, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(
                        "Refusing to modify: the item's subject does not carry both the test tag and this run's marker (S3).");
                }

                if ((bool)item.Sent)
                {
                    throw new InvalidOperationException("Refusing to modify: the item is not an unsent draft.");
                }

                item.Body = (string)item.Body + textToAppend;
                item.Save();
                return null;
            }
            finally
            {
                Release(item);
                Release(store);
                Release(stores);
                Release(ns);
                Release(app);
            }
        });
    }

    /// <summary>
    /// Counts items whose subject contains <paramref name="subjectFragment"/> across
    /// the store's default folders (default set: Drafts, Inbox, Sent Items, Deleted
    /// Items; hub archive coverage via <see cref="HubSweepFolderIdsWithArchive"/>) -
    /// the post-suite artifact sweep (S3). Read-only; output is a count
    /// (content-free, S4). Uses Folder.GetTable with a DASL LIKE restriction, falling
    /// back to Items.Restrict; throws when a folder cannot be counted at all.
    /// </summary>
    public static int CountTaggedArtifacts(string storeDisplayName, string subjectFragment, int[]? folderIds = null)
    {
        if (string.IsNullOrWhiteSpace(subjectFragment) || subjectFragment.Length < 8)
        {
            throw new ArgumentException("Fragment too weak for a meaningful sweep.", nameof(subjectFragment));
        }

        if (subjectFragment.Contains('\'', StringComparison.Ordinal) || subjectFragment.Contains('%', StringComparison.Ordinal))
        {
            throw new ArgumentException("Fragment must not contain quote/wildcard characters.", nameof(subjectFragment));
        }

        int[] folders = folderIds ?? SweepFolderIds;
        return RunSta(() =>
        {
            dynamic app = CreateOutlookApplication();
            dynamic? ns = null;
            dynamic? stores = null;
            dynamic? store = null;
            int total = 0;
            try
            {
                ns = app.GetNamespace("MAPI");
                stores = ns.Stores;
                store = FindStore(stores, storeDisplayName)
                    ?? throw new InvalidOperationException("Store not found for artifact sweep.");
                foreach (int folderId in folders)
                {
                    total += CountMatchingInFolder(store, folderId, subjectFragment);
                }

                return total;
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

    private static int CountMatchingInFolder(dynamic store, int folderId, string subjectFragment)
    {
        dynamic? folder = null;
        dynamic? table = null;
        dynamic? items = null;
        dynamic? restricted = null;
        string filter = "@SQL=\"urn:schemas:httpmail:subject\" LIKE '%" + subjectFragment + "%'";
        try
        {
            try
            {
                folder = store.GetDefaultFolder(folderId);
            }
            catch (COMException)
            {
                return 0; // store without that default folder
            }

            try
            {
                table = folder.GetTable(filter);
                int count = 0;
                while (!(bool)table.EndOfTable)
                {
                    dynamic? row = null;
                    try
                    {
                        row = table.GetNextRow();
                        count++;
                    }
                    finally
                    {
                        Release(row);
                    }
                }

                return count;
            }
            catch (Exception ex) when (OutlookAI.Core.Com.OutlookComSession.IsComCallFailure(ex))
            {
                // Fall back to Items.Restrict below.
            }

            items = folder.Items;
            restricted = items.Restrict(filter);
            return (int)restricted.Count;
        }
        finally
        {
            Release(restricted);
            Release(items);
            Release(table);
            Release(folder);
        }
    }

    /// <summary>
    /// Read-only: counts folders anywhere in the store whose Name contains
    /// <see cref="TestFolderNamePrefix"/> (ordinal) - the D39 post-suite
    /// zero-test-folders assert.
    /// </summary>
    public static int CountTestFolders(string storeDisplayName)
    {
        return RunSta(() =>
        {
            dynamic app = CreateOutlookApplication();
            dynamic? ns = null;
            dynamic? stores = null;
            dynamic? store = null;
            dynamic? root = null;
            try
            {
                ns = app.GetNamespace("MAPI");
                stores = ns.Stores;
                store = FindStore(stores, storeDisplayName)
                    ?? throw new InvalidOperationException("Store not found for test-folder count.");
                root = store.GetRootFolder();
                List<(string EntryId, string Name, int Depth)> matches = new();
                CollectTestFolders(root, matches, 0);
                return matches.Count;
            }
            finally
            {
                Release(root);
                Release(store);
                Release(stores);
                Release(ns);
                Release(app);
            }
        });
    }

    /// <summary>
    /// Deletes every folder in the store whose Name contains
    /// <see cref="TestFolderNamePrefix"/> (ordinal - the S3 discipline for folders;
    /// items inside must all carry the subject tag, otherwise this REFUSES).
    ///
    /// ⚠ Sync-wedge footgun (live-probed 2026-07-26): deleting ITEMS while they sit
    /// INSIDE a folder marks that folder "synchronizing local changes" on this
    /// cached-Exchange store, and a folder in that state cannot be removed from
    /// Deleted Items for the rest of the Outlook session (Folders.Remove throws the
    /// synchronization error; only a restart clears it). Item deletions from
    /// PERMANENT folders (Inbox, Deleted Items) never wedge anything. Therefore:
    /// tagged contents are MOVED OUT to the Inbox first and deleted THERE, the then-
    /// empty folder (move-history only - probe-proven removable) is soft-deleted,
    /// and its Deleted Items copy is hard-removed via Folders.Remove. Passes repeat
    /// until a fresh walk finds ZERO matches (verified); returns deletions performed.
    /// </summary>
    public static int DeleteTestFolders(string storeDisplayName)
    {
        int total = 0;
        for (int pass = 0; pass < 6; pass++)
        {
            int actionsThisPass = RunSta(() =>
            {
                dynamic app = CreateOutlookApplication();
                dynamic? ns = null;
                dynamic? stores = null;
                dynamic? store = null;
                dynamic? root = null;
                try
                {
                    ns = app.GetNamespace("MAPI");
                    stores = ns.Stores;
                    store = FindStore(stores, storeDisplayName)
                        ?? throw new InvalidOperationException("Store not found for test-folder cleanup.");
                    string storeId = (string)store.StoreID;
                    root = store.GetRootFolder();
                    List<(string EntryId, string Name, int Depth)> matches = new();
                    CollectTestFolders(root, matches, 0);

                    int actions = 0;
                    // DEEPEST FIRST: a parent whose test child still exists cannot be
                    // removed, and a parent-first pass would keep rediscovering both
                    // until the retry window ran out (live-bitten by the first nested
                    // test folder, soak fix 15).
                    foreach ((string entryId, string name, int _) in matches.OrderByDescending(m => m.Depth))
                    {
                        if (!name.Contains(TestFolderNamePrefix, StringComparison.Ordinal))
                        {
                            continue; // double-check the guard before any delete
                        }

                        dynamic? folder = null;
                        dynamic? remaining = null;
                        try
                        {
                            folder = ns.GetFolderFromID(entryId, storeId);
                            EnsureFolderContainsOnlyTaggedItems(folder!);
                            EvictTaggedItemsViaInbox(store, folder!); // no in-place deletions (wedge)
                            remaining = folder!.Items;
                            if ((int)remaining.Count == 0)
                            {
                                actions += RemoveEmptyTestFolder(store, folder!, entryId);
                            }
                        }
                        catch (Exception ex) when (OutlookAI.Core.Com.OutlookComSession.IsComCallFailure(ex))
                        {
                            // Already gone, or a transient sync refusal - the next
                            // pass retries against fresh state.
                        }
                        finally
                        {
                            Release(remaining);
                            Release(folder);
                        }
                    }

                    return actions;
                }
                finally
                {
                    Release(root);
                    Release(store);
                    Release(stores);
                    Release(ns);
                    Release(app);
                }
            });

            total += actionsThisPass;
            if (actionsThisPass == 0 && CountTestFolders(storeDisplayName) == 0)
            {
                return total; // verified clean
            }

            Thread.Sleep(1500); // let the store register moves before the next walk
        }

        if (CountTestFolders(storeDisplayName) != 0)
        {
            throw new InvalidOperationException(
                "Test folders kept reappearing for the whole cleanup window - manual check required (S3).");
        }

        return total;
    }

    /// <summary>
    /// Moves every item of a to-be-deleted test folder (already verified all-tagged)
    /// to the Inbox and deletes it THERE - deletions recorded on a permanent folder
    /// never wedge the test folder's own sync state (see DeleteTestFolders remarks).
    /// Per-item tag re-check as the last line of defense.
    /// </summary>
    private static void EvictTaggedItemsViaInbox(dynamic store, dynamic folder)
    {
        dynamic? inbox = null;
        dynamic? items = null;
        try
        {
            inbox = store.GetDefaultFolder(6);
            items = folder.Items;
            int count = items.Count;
            for (int i = count; i >= 1; i--)
            {
                dynamic? item = null;
                dynamic? moved = null;
                try
                {
                    item = items[i];
                    string? subject = null;
                    try
                    {
                        subject = (string?)item.Subject;
                    }
                    catch (COMException)
                    {
                    }

                    if (subject != null && subject.Contains(SubjectTag, StringComparison.Ordinal))
                    {
                        moved = item.Move(inbox);
                        moved.Delete(); // recorded on the Inbox, purged by the folder-3 sweep pass
                    }
                }
                catch (COMException)
                {
                }
                finally
                {
                    Release(moved);
                    Release(item);
                }
            }
        }
        finally
        {
            Release(items);
            Release(inbox);
        }
    }

    /// <summary>
    /// Removes one EMPTY test folder: under Deleted Items it is hard-removed via
    /// Folders.Remove (index resolved by EntryID); anywhere else it is soft-deleted
    /// (the next pass hard-removes the Deleted Items copy). Returns actions performed.
    /// </summary>
    private static int RemoveEmptyTestFolder(dynamic store, dynamic folder, string folderEntryId)
    {
        bool underDeletedItems = false;
        dynamic? parent = null;
        dynamic? deletedItems = null;
        try
        {
            deletedItems = store.GetDefaultFolder(3);
            string deletedItemsEntryId = (string)deletedItems.EntryID;
            try
            {
                parent = folder.Parent;
                underDeletedItems = parent != null
                    && string.Equals((string)((dynamic)parent!).EntryID, deletedItemsEntryId, StringComparison.OrdinalIgnoreCase);
            }
            catch (COMException)
            {
            }

            if (!underDeletedItems)
            {
                folder.Delete(); // soft: moves under Deleted Items with a NEW EntryID
                return 1;
            }

            dynamic? siblings = null;
            try
            {
                siblings = ((dynamic)parent!).Folders;
                int count = siblings.Count;
                for (int i = count; i >= 1; i--)
                {
                    dynamic? candidate = null;
                    try
                    {
                        candidate = siblings[i];
                        if (string.Equals((string)candidate.EntryID, folderEntryId, StringComparison.OrdinalIgnoreCase))
                        {
                            siblings.Remove(i); // hard delete from Deleted Items
                            return 1;
                        }
                    }
                    finally
                    {
                        Release(candidate);
                    }
                }
            }
            finally
            {
                Release(siblings);
            }

            return 0;
        }
        finally
        {
            Release(deletedItems);
            Release(parent);
        }
    }

    private static void CollectTestFolders(dynamic folder, List<(string EntryId, string Name, int Depth)> matches, int depth)
    {
        if (depth > 32)
        {
            return;
        }

        dynamic? children = null;
        try
        {
            children = folder.Folders;
            int count = children.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic? child = null;
                try
                {
                    child = children[i];
                    string? name = null;
                    try
                    {
                        name = (string?)child.Name;
                    }
                    catch (COMException)
                    {
                    }

                    if (name != null && name.Contains(TestFolderNamePrefix, StringComparison.Ordinal))
                    {
                        matches.Add(((string)child.EntryID, name, depth));
                    }

                    CollectTestFolders(child, matches, depth + 1);
                }
                catch (COMException)
                {
                }
                finally
                {
                    Release(child);
                }
            }
        }
        catch (COMException)
        {
        }
        finally
        {
            Release(children);
        }
    }

    /// <summary>
    /// S3 guard for folder deletion: every item inside a to-be-deleted test folder
    /// must carry the subject tag (folders are only ever deleted with their own test
    /// contents - anything else in there aborts the cleanup loudly).
    /// </summary>
    private static void EnsureFolderContainsOnlyTaggedItems(dynamic folder)
    {
        dynamic? items = null;
        try
        {
            items = folder.Items;
            int count = items.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic? item = null;
                try
                {
                    item = items[i];
                    string? subject = null;
                    try
                    {
                        subject = (string?)item.Subject;
                    }
                    catch (COMException)
                    {
                    }

                    if (subject == null || !subject.Contains(SubjectTag, StringComparison.Ordinal))
                    {
                        throw new InvalidOperationException(
                            "Refusing to delete test folder: it contains an item without the test tag (S3).");
                    }
                }
                finally
                {
                    Release(item);
                }
            }
        }
        finally
        {
            Release(items);
        }
    }

    private static dynamic? FindStore(dynamic stores, string storeDisplayName)
    {
        int storeCount = stores.Count;
        for (int i = 1; i <= storeCount; i++)
        {
            dynamic candidate = stores[i];
            string? name = null;
            try
            {
                name = (string?)candidate.DisplayName;
            }
            catch (COMException)
            {
            }

            if (string.Equals(name, storeDisplayName, StringComparison.OrdinalIgnoreCase))
            {
                return candidate;
            }

            Release(candidate);
        }

        return null;
    }

    private static int DeleteMatchingInFolder(dynamic store, int folderId, string uniqueMarker)
    {
        dynamic? folder = null;
        dynamic? items = null;
        int deleted = 0;
        try
        {
            try
            {
                folder = store.GetDefaultFolder(folderId);
            }
            catch (COMException)
            {
                return 0;
            }

            // Cheap pre-check before the full item walk: the item-by-item pass below is
            // the TESTED delete path (double-match on tag AND marker, per S3) and must
            // stay, but walking every item of a folder holding thousands - which the
            // Sync Issues subtree does on a busy store - costs seconds for nothing when
            // the folder holds no artifact at all. GetTable with a DASL restriction
            // answers that in milliseconds.
            if (!FolderMayContain(folder, uniqueMarker))
            {
                return 0;
            }

            items = folder.Items;
            int count = items.Count;
            // Iterate backwards: Delete() reindexes the collection.
            for (int i = count; i >= 1; i--)
            {
                dynamic? item = null;
                try
                {
                    item = items[i];
                    string? subject = null;
                    try
                    {
                        subject = (string?)item.Subject;
                    }
                    catch (COMException)
                    {
                    }

                    if (subject != null
                        && subject.Contains(SubjectTag, StringComparison.Ordinal)
                        && subject.Contains(uniqueMarker, StringComparison.OrdinalIgnoreCase))
                    {
                        item.Delete();
                        deleted++;
                    }
                }
                catch (COMException)
                {
                    // Undeletable/transient row - leave it; the tag keeps it identifiable.
                }
                finally
                {
                    Release(item);
                }
            }

            return deleted;
        }
        finally
        {
            Release(items);
            Release(folder);
        }
    }

    /// <summary>
    /// True when the folder holds at least one item whose subject carries the marker.
    /// Conservative: any COM failure returns true so the caller still does the full,
    /// tested walk - a cheap optimization must never become a silent skip.
    /// </summary>
    private static bool FolderMayContain(dynamic folder, string uniqueMarker)
    {
        if (uniqueMarker.IndexOf('\'') >= 0 || uniqueMarker.IndexOf('%') >= 0)
        {
            return true; // Not expressible as a DASL literal - walk everything.
        }

        dynamic? table = null;
        try
        {
            table = folder.GetTable("@SQL=\"urn:schemas:httpmail:subject\" LIKE '%" + uniqueMarker + "%'");
            return !(bool)table.EndOfTable;
        }
        catch (COMException)
        {
            return true;
        }
        catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
        {
            return true;
        }
        finally
        {
            Release(table);
        }
    }

    private static dynamic CreateOutlookApplication()
    {
        Type progIdType = Type.GetTypeFromProgID("Outlook.Application")
            ?? throw new InvalidOperationException("Outlook.Application ProgID is not registered.");
        return Activator.CreateInstance(progIdType)
            ?? throw new InvalidOperationException("Failed to create Outlook.Application.");
    }

    private static T RunSta<T>(Func<T> work)
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
            Name = "OutlookAI.TestMailer.Sta",
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        if (!thread.Join(TimeSpan.FromMinutes(3)))
        {
            throw new TimeoutException("Test mailer STA operation timed out.");
        }

        if (failure != null)
        {
            throw new InvalidOperationException("Test mailer operation failed.", failure);
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
