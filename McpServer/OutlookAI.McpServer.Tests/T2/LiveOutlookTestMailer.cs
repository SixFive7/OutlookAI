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
    /// Deletes every item in the store's Drafts, Inbox, Sent Items and Deleted Items
    /// whose subject contains BOTH the tag and <paramref name="uniqueMarker"/> (S3:
    /// only artifacts this run created). Two passes: Delete() moves to Deleted Items,
    /// the second pass on folder 3 removes them for good. Returns the total deleted.
    /// </summary>
    public static int DeleteTaggedArtifacts(string storeDisplayName, string uniqueMarker)
    {
        if (string.IsNullOrWhiteSpace(uniqueMarker) || uniqueMarker.Length < 12)
        {
            throw new ArgumentException("Marker too weak for a safe delete filter (S3).", nameof(uniqueMarker));
        }

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
                    // 16 = Drafts (Phase 4), 6 = Inbox, 5 = Sent Items, 3 = Deleted
                    // Items (second pass).
                    foreach (int folderId in new[] { 16, 6, 5, 3 })
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
        TimeSpan? stableFor = null)
    {
        TimeSpan totalWindow = window ?? TimeSpan.FromSeconds(120);
        TimeSpan requiredStable = stableFor ?? TimeSpan.FromSeconds(10);
        DateTime deadline = DateTime.UtcNow + totalWindow;
        DateTime? zeroSince = null;
        int totalDeleted = 0;
        while (DateTime.UtcNow < deadline)
        {
            int remaining = CountTaggedArtifacts(storeDisplayName, uniqueMarker);
            if (remaining > 0)
            {
                totalDeleted += DeleteTaggedArtifacts(storeDisplayName, uniqueMarker);
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
    /// Counts items whose subject contains <paramref name="subjectFragment"/> across
    /// the store's default folders (default set: Drafts, Inbox, Sent Items, Deleted
    /// Items) - the post-suite artifact sweep (S3). Read-only; output is a count
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

        int[] folders = folderIds ?? new[] { 16, 6, 5, 3 };
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
