using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;

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
    /// OUTBOX, the SYNC ISSUES subtree, then Deleted Items LAST so the final pass
    /// purges what Delete() moved there.
    /// <para>
    /// The Sync Issues subtree (20 Sync Issues, 19 Conflicts, 21 Local Failures,
    /// 22 Server Failures) was added after a tagged test item was found stranded in
    /// the hub's Local Failures folder: a cached-Exchange store can file a copy of a
    /// test artifact there on its own, and nothing swept it. Those folders hold
    /// hundreds of items, not the ~100k an archive holds, so counting them is cheap.
    /// </para>
    /// <para>
    /// The Outbox (4) joined them for the same reason: when Outlook is running HEADLESS
    /// it may not dispatch a queued self-send, so the test mail sits in the Outbox, the
    /// arrival wait times out, and the artifact outlives its own cleanup - observed on a
    /// full-suite run whose Outlook had been started by the tests themselves. Deleting
    /// there is S3-legal (tag AND marker must both match) and cannot touch real mail.
    /// </para>
    /// <para>
    /// ⚠ Deliberately WITHOUT the Archive folder: business-store archives hold ~100k
    /// items and a LIKE count over them takes minutes - adding 39 here made the
    /// cross-account sweeps time out (live-bitten 2026-07-26). Hub-scoped D39
    /// cleanups pass <see cref="HubSweepFolderIdsWithArchive"/> explicitly instead.
    /// </para>
    /// </summary>
    private static readonly int[] SweepFolderIds = { 16, 6, 5, 4, 20, 19, 21, 22, 3 };

    /// <summary>
    /// Sweep set for the TINY test hub only (D39): includes the designated Archive
    /// folder (39) so archive_mail artifacts are counted and purged. Never use
    /// against business stores (their archives are huge - see SweepFolderIds).
    /// </summary>
    public static readonly int[] HubSweepFolderIdsWithArchive = { 16, 6, 5, 39, 4, 20, 19, 21, 22, 3 };

    /// <summary>
    /// The Sync Issues subtree ids, exposed so a targeted purge can name them
    /// (olFolderSyncIssues 20, olFolderConflicts 19, olFolderLocalFailures 21,
    /// olFolderServerFailures 22).
    /// </summary>
    public static readonly int[] SyncIssuesFolderIds = { 20, 19, 21, 22 };

    /// <summary>
    /// Rows the tripwire census asks a folder table for in one <c>Table.GetArray</c> call.
    /// <para>
    /// This is the number that turned the census from unaffordable into cheap, so it is
    /// worth saying what it trades. Upwards it saves round trips and there are barely any
    /// left to save: at 200 the whole 3,000-item per-store budget costs about fifteen calls,
    /// and the per-folder overhead already dominates. Downwards it bounds what one call
    /// materialises in this process - up to 200 rows of EntryIDs and SUBJECTS at a time,
    /// which is the only moment another mailbox's subjects exist here at all (they are
    /// projected to one boolean each and dropped; S3/S4). 200 keeps that transient small
    /// while leaving the round-trip saving essentially complete.
    /// </para>
    /// </summary>
    public const int CensusTableRowBatch = 200;

    /// <summary>
    /// How long any ordinary mailer STA operation may run before the join gives up and calls
    /// it a timeout. Unchanged at three minutes: it is a HANG detector for a wedged Outlook,
    /// not a work allowance, and every operation it covers is a handful of COM calls.
    /// </summary>
    internal static readonly TimeSpan DefaultStaBudget = TimeSpan.FromMinutes(3);

    /// <summary>
    /// The census's own STA join, and the rung directly above
    /// <see cref="CensusIdentityPlan.DefaultIdentityTimeBudgetMs"/>.
    /// <para>
    /// DERIVED rather than chosen, and the derivation is the point. The identity walk now
    /// stops itself after its own budget and lets the store be counted; that can only happen
    /// if the join outlives the budget by enough to finish the counting underneath it. Set
    /// them equal - or leave the census on the ordinary 3-minute join, which is what it had
    /// before - and the outer timer fires first, killing a census that was working perfectly
    /// well inside its own budget and refusing the whole live tier. This repository has
    /// already shipped that failure once (the inner sweep budget equal to the outer deadline,
    /// 2026-08-18), which is why the ordering is pinned in T1 rather than left to reading.
    /// </para>
    /// </summary>
    internal static readonly TimeSpan CensusStaBudget =
        TimeSpan.FromMilliseconds(CensusIdentityPlan.DefaultIdentityTimeBudgetMs) + DefaultStaBudget;

    /// <summary>
    /// Sends a mail from <paramref name="smtpAddress"/> to itself (refuses to run when
    /// that account is not in the profile - the D20 grant is telefonie-to-telefonie
    /// only). Returns the UTC send timestamp.
    /// </summary>
    public static DateTime SendSelfMail(string smtpAddress, string subject, string body, string? attachmentPath)
    {
        return SendSelfMailWithAttachments(
            smtpAddress, subject, body, attachmentPath == null ? null : new[] { attachmentPath });
    }

    /// <summary>
    /// As above with SEVERAL attachments - the attachment-recall proof needs a mail
    /// carrying more than one attachment type at once (soak fix 16).
    /// </summary>
    public static DateTime SendSelfMailWithAttachments(
        string smtpAddress, string subject, string body, IReadOnlyList<string>? attachmentPaths)
    {
        LiveStoreWriteGuard.Assert(smtpAddress, StoreWriteKind.Send, nameof(SendSelfMail));
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
                    if (attachmentPaths != null && attachmentPaths.Count > 0)
                    {
                        dynamic attachments = mail.Attachments;
                        try
                        {
                            foreach (string attachmentPath in attachmentPaths)
                            {
                                attachments.Add(attachmentPath);
                            }
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

                    // Send() only QUEUES the mail. A user-driven Outlook flushes the
                    // Outbox on its own schedule, but an Outlook the tests started
                    // themselves (D17, headless) may sit on it indefinitely - the seed
                    // then never arrives, the arrival wait times out, and the artifact
                    // outlives its cleanup in the Outbox (observed on a full-suite run
                    // whose Outlook had exited beforehand). Ask for delivery explicitly;
                    // best-effort, because a profile can refuse it and the wait loop
                    // remains the authority either way.
                    try
                    {
                        session!.SendAndReceive(false);
                    }
                    catch (COMException)
                    {
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    {
                    }
                }
                finally
                {
                    Release(sendingAccount);
                }

                // Wall clock ON PURPOSE: every caller turns this into a sweep window base
                // compared against the DateReceived that Outlook stamps on the arriving
                // copy, so it has to be real calendar time rather than a monotonic reading.
                // The durations the live tier MEASURES are on LiveWaitBudget instead.
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
        LiveStoreWriteGuard.Assert(storeDisplayName, StoreWriteKind.Delete, nameof(DeleteTaggedArtifacts));
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
        // Monotonic on both counters: this loop decides when a mailbox is CLEAN, and a
        // forwards clock jump would end it with tagged items still in a real mailbox while
        // a backwards one would keep deleting for as long as the jump lasted.
        LiveWaitBudget wait = LiveWaitBudget.Of(totalWindow);
        TimeSpan? zeroSince = null;
        int totalDeleted = 0;
        while (wait.HasTimeLeft)
        {
            int remaining = CountTaggedArtifacts(storeDisplayName, uniqueMarker, folderIds);
            if (remaining > 0)
            {
                totalDeleted += DeleteTaggedArtifacts(storeDisplayName, uniqueMarker, folderIds);
                zeroSince = null;
                continue;
            }

            zeroSince ??= wait.Elapsed;
            if (wait.Elapsed - zeroSince.Value >= requiredStable)
            {
                return totalDeleted;
            }

            Thread.Sleep(2000);
        }

        throw new TimeoutException(
            "Tagged artifacts kept (re)appearing for the whole cleanup window - manual check required (S3).");
    }

    /// <summary>
    /// Final post-test zero check for suites that SELF-SEND. <see
    /// cref="DeleteTaggedArtifactsUntilStableZero"/> returns once the count has held at
    /// zero for its stability window, but a Sent-Items or Inbox copy can still surface
    /// LATER than that under load (Phase-4 fact 6 / soak fix 15 - live-observed twice
    /// again during soak-fix batch A, where two suites of the same collection each found
    /// exactly one straggler seconds after a stable zero). A straggler is therefore
    /// PURGED once more - S3-legal, it matches tag AND this run's marker - and only what
    /// survives that second purge is reported. Returns the final count (assert it is 0)
    /// and sets <paramref name="stragglersPurged"/> for the test's log line, so a genuine
    /// leak still fails loudly while the documented lag does not.
    /// </summary>
    public static int CountTaggedArtifactsAfterPurgingStragglers(
        string storeDisplayName,
        string uniqueMarker,
        int[]? folderIds,
        out int stragglersPurged)
    {
        stragglersPurged = 0;
        int count = CountTaggedArtifacts(storeDisplayName, uniqueMarker, folderIds);
        if (count == 0)
        {
            return 0;
        }

        stragglersPurged = count;
        DeleteTaggedArtifactsUntilStableZero(storeDisplayName, uniqueMarker, folderIds: folderIds);
        return CountTaggedArtifacts(storeDisplayName, uniqueMarker, folderIds);
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
        LiveStoreWriteGuard.Assert(storeDisplayName, StoreWriteKind.Delete, nameof(DeleteItemByEntryId));
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
        LiveStoreWriteGuard.Assert(storeDisplayName, StoreWriteKind.Draft, nameof(AppendToDraftBody));
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
    /// Replaces the raw <c>HTMLBody</c> of one draft this run created. Exists for exactly
    /// one purpose (D47): seeding a draft in the LEGACY shape - a signature image left as
    /// a <c>file:///</c> LINK rather than an embedded <c>cid:</c> resource - so the
    /// update path's rescue of such a draft can be proven rather than assumed. Same S3
    /// double-match guard as <see cref="AppendToDraftBody"/>: tag AND this run's marker,
    /// and the item must still be unsent.
    /// </summary>
    public static void SetDraftHtmlBody(string storeDisplayName, string entryIdHex, string uniqueMarker, string html)
    {
        LiveStoreWriteGuard.Assert(storeDisplayName, StoreWriteKind.Draft, nameof(SetDraftHtmlBody));
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

                item.HTMLBody = html;
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
    /// <summary>
    /// Saves a tagged DRAFT carrying <paramref name="attachmentPaths"/> into the given
    /// store's Drafts folder and returns its EntryID.
    /// <para>
    /// Deliberately a draft rather than a self-send: the attachment-recall proof only
    /// needs an indexed item that HOLDS the attachments, and transport adds a dependency
    /// on Outlook actually flushing the Outbox (which a headless instance may not do -
    /// soak fix 15). Drafts are indexed exactly like received mail.
    /// </para>
    /// </summary>
    public static string SaveTaggedDraftWithAttachments(
        string storeDisplayName, string subject, string body, IReadOnlyList<string> attachmentPaths)
    {
        LiveStoreWriteGuard.Assert(storeDisplayName, StoreWriteKind.Draft, nameof(SaveTaggedDraftWithAttachments));
        if (!subject.Contains(SubjectTag, StringComparison.Ordinal))
        {
            throw new ArgumentException($"Test draft subject must carry the {SubjectTag} tag (S3).", nameof(subject));
        }

        return RunSta(() =>
        {
            dynamic app = CreateOutlookApplication();
            dynamic? ns = null;
            dynamic? stores = null;
            dynamic? store = null;
            dynamic? drafts = null;
            dynamic? items = null;
            dynamic? mail = null;
            try
            {
                ns = app.GetNamespace("MAPI");
                stores = ns.Stores;
                store = FindStore(stores, storeDisplayName)
                    ?? throw new InvalidOperationException("Store not found for the attachment draft.");

                // Per-account filing (Phase-4 footgun): a plain CreateItem draft lands in
                // the DEFAULT store's Drafts.
                drafts = store.GetDefaultFolder(16);
                items = drafts.Items;
                mail = items.Add(0); // olMailItem
                mail.Subject = subject;
                mail.Body = body;

                dynamic attachments = mail.Attachments;
                try
                {
                    foreach (string path in attachmentPaths)
                    {
                        attachments.Add(path);
                    }
                }
                finally
                {
                    Release(attachments);
                }

                mail.Save();
                return (string)mail.EntryID;
            }
            finally
            {
                Release(mail);
                Release(items);
                Release(drafts);
                Release(store);
                Release(stores);
                Release(ns);
                Release(app);
            }
        });
    }

    /// <summary>
    /// READ-ONLY census for the per-store tripwire: store-relative path -&gt; what was in
    /// every MAIL folder of <paramref name="storeDisplayName"/> (mail-typed folders plus
    /// Deleted Items, Outbox and the Sync Issues subtree, which are all mail-typed).
    /// Calendar/Contacts/Tasks are skipped deliberately - their churn is not mail loss and
    /// would only add false positives.
    /// <para>
    /// Every folder is COUNTED. Folders within <paramref name="plan"/>'s budget are also
    /// WALKED, so the tripwire can say which items left rather than only how many: a count
    /// cannot tell a deletion from a filing, and cannot see an item removed while another
    /// arrives. The walk is a bulk <c>Folder.GetTable</c> read of four columns - EntryID,
    /// ReceivedTime, Size and Subject - and the subject is projected to a boolean saying
    /// whether it carried <see cref="SubjectTag"/> and then dropped, so no other mailbox's
    /// content is stored or logged (S3/S4).
    /// </para>
    /// <para>
    /// Throws if the store cannot be enumerated - the tripwire is fail-closed. A folder
    /// whose WALK fails or comes back inconsistent degrades to a count, never to nothing, and
    /// so does a folder the identity TIME budget can no longer afford: this call runs under
    /// <see cref="CensusStaBudget"/> precisely so that budget can expire inside it and leave
    /// the counting to finish, instead of the join killing the store's census outright.
    /// </para>
    /// </summary>
    public static IReadOnlyDictionary<string, FolderCensus> CaptureMailFolderCensus(
        string storeDisplayName, CensusIdentityPlan plan)
    {
        ArgumentNullException.ThrowIfNull(plan);
        return RunSta<IReadOnlyDictionary<string, FolderCensus>>(
            () =>
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
                    ?? throw new InvalidOperationException(
                        "Store not found for the count tripwire: '" + storeDisplayName + "'.");
                root = store.GetRootFolder();

                // Folders the SYSTEM prunes on its own: Deleted Items ages out, JUNK MAIL
                // expires on a server policy nobody here controls, and Outlook writes and
                // removes sync-issue reports whenever it feels like it. A shrink there is
                // not evidence of anything, so the census marks them and the tripwire notes
                // rather than fails them. Junk (23) was missing from this list and cost a
                // false alarm on 2026-08-18: 'Ongewenste e-mail' going 1 -> 0 during a run
                // is junk expiry, not mail loss.
                HashSet<string> volatileIds = new(StringComparer.OrdinalIgnoreCase);
                foreach (int volatileFolderId in new[] { 3, 20, 19, 21, 22, 23 })
                {
                    try
                    {
                        dynamic volatileFolder = store.GetDefaultFolder(volatileFolderId);
                        try
                        {
                            volatileIds.Add((string)volatileFolder.EntryID);
                        }
                        finally
                        {
                            Release(volatileFolder);
                        }
                    }
                    catch (Exception ex) when (ex is COMException or RuntimeBinderException)
                    {
                    }
                }

                Dictionary<string, FolderCensus> census = new(StringComparer.OrdinalIgnoreCase);
                CollectMailFolderCensus(root, string.Empty, census, 0, volatileIds, false, plan);
                if (census.Count == 0)
                {
                    throw new InvalidOperationException(
                        "The count tripwire found no mail folder in '" + storeDisplayName + "'.");
                }

                return census;
            }
            finally
            {
                Release(root);
                Release(store);
                Release(stores);
                Release(ns);
                Release(app);
            }
        },
            CensusStaBudget);
    }

    private static void CollectMailFolderCensus(
        dynamic folder,
        string parentPath,
        Dictionary<string, FolderCensus> census,
        int depth,
        HashSet<string> volatileFolderIds,
        bool parentIsVolatile,
        CensusIdentityPlan plan)
    {
        if (depth > 32)
        {
            return;
        }

        dynamic? children = null;
        try
        {
            string name = (string)folder.Name;
            string path = parentPath.Length == 0 ? name : parentPath + "/" + name;
            bool isVolatile = parentIsVolatile;
            if (!isVolatile)
            {
                try
                {
                    isVolatile = volatileFolderIds.Contains((string)folder.EntryID);
                }
                catch (Exception ex) when (ex is COMException or RuntimeBinderException)
                {
                }
            }

            if (depth > 0)
            {
                bool isMail;
                try
                {
                    isMail = (int)folder.DefaultItemType == 0; // olMailItem
                }
                catch (Exception ex) when (ex is COMException or RuntimeBinderException)
                {
                    isMail = false;
                }

                if (isMail)
                {
                    string key = (isVolatile ? StoreCountTripwire.VolatilePrefix : string.Empty) + path;
                    census[key] = CaptureFolder(folder, key, isVolatile, plan);
                }
            }

            children = folder.Folders;
            int childCount = children.Count;
            for (int i = 1; i <= childCount; i++)
            {
                dynamic child = children[i];
                try
                {
                    CollectMailFolderCensus(
                        child, depth == 0 ? string.Empty : path, census, depth + 1, volatileFolderIds, isVolatile,
                        plan);
                }
                finally
                {
                    Release(child);
                }
            }
        }
        finally
        {
            Release(children);
        }
    }

    /// <summary>
    /// One folder's entry: always the count, plus the item list when <paramref name="plan"/>
    /// can afford it. Degrades to a count on any failure - a folder the tripwire can only
    /// count is still guarded, while a folder it cannot measure at all would be a hole.
    /// </summary>
    private static FolderCensus CaptureFolder(dynamic folder, string key, bool isVolatile, CensusIdentityPlan plan)
    {
        dynamic? items = null;
        try
        {
            plan.NoteFolderMeasured();
            items = folder.Items;
            int count = (int)items.Count;
            if (!plan.ShouldIdentify(key, isVolatile, count))
            {
                return FolderCensus.CountOnly(count);
            }

            List<CensusItem>? walked = WalkFolderItems(folder, items, count);
            if (walked == null)
            {
                plan.NoteDegradedToCount();
                return FolderCensus.CountOnly(count);
            }

            plan.Spend(walked.Count);
            return FolderCensus.WithItems(walked);
        }
        finally
        {
            Release(items);
        }
    }

    /// <summary>
    /// Reads one folder in BULK from its <c>Table</c>, or returns null when the folder
    /// moved under the read (mail arriving mid-census) or the table could not answer what
    /// the census needs. Null means "count this folder instead", never "assume nothing
    /// changed".
    /// <para>
    /// WHY A TABLE. Until 2026-08-20 this walked <c>Items[i]</c> and read four properties
    /// off each item: five cross-process calls per item, up to the 3,000-item store budget,
    /// so 15,000 round trips for one store of one census pass. That is affordable against a
    /// local PST and it is not affordable against an Exchange mailbox, where a shared or
    /// delegate store may not be cached at all and every one of those calls is a server
    /// round trip - it needs only 12 ms per call to exceed the 3-minute STA budget, and on
    /// 2026-08-20 one store did exactly that and refused the whole live tier.
    /// <c>Table.GetArray</c> returns <see cref="CensusTableRowBatch"/> rows of exactly the
    /// columns asked for in ONE call, so the same 3,000 items cost about fifteen.
    /// </para>
    /// <para>
    /// WHY THIS IS NOT THE <c>Items.SetColumns</c> TRAP THIS CODE USED TO AVOID.
    /// <c>SetColumns</c> leaves the ITEM in place and documents that some of its properties
    /// come back empty - so a blank EntryID would read as an item that had vanished, which
    /// is the one wrong answer this guard must never give. A table row is not an item: its
    /// EntryID is a column read straight out of the contents table, and a row that does not
    /// produce one is DETECTED here and abandons the folder to a count. Every other shape
    /// failure is treated the same way, because the alternative to a weaker reading is a
    /// false one.
    /// </para>
    /// <para>
    /// Three independent cross-checks have to agree before a walk is accepted: the table
    /// must hand back exactly as many rows as <c>Items.Count</c> promised, no EntryID may
    /// repeat, and the count must still be the same afterwards. Any disagreement means the
    /// reading spans two moments and is discarded.
    /// </para>
    /// </summary>
    private static List<CensusItem>? WalkFolderItems(dynamic folder, dynamic items, int expectedCount)
    {
        if (expectedCount < 0)
        {
            return null;
        }

        // An empty folder needs no table at all, and there are a lot of them (Outbox, the
        // sync-issue folders, unused subfolders). It still gets the confirmation read: an
        // item arriving into it mid-census makes even "it was empty" a two-moment reading.
        if (expectedCount == 0)
        {
            return ConfirmUnchanged(items, 0) ? new List<CensusItem>() : null;
        }

        dynamic? table = null;
        try
        {
            try
            {
                table = folder.GetTable();
            }
            catch (Exception ex) when (OutlookAI.Core.Com.OutlookComSession.IsComCallFailure(ex))
            {
                return null;
            }

            dynamic t = table!;
            AddCensusColumn(t, CensusTableRow.ReceivedColumnNames);
            AddCensusColumn(t, CensusTableRow.SizeColumnNames);

            CensusColumnMap columns = MapCensusColumns(t);
            if (!columns.IsUsable)
            {
                // The table cannot say when an item arrived, how big it is, or whether it is
                // the suite's own. Identity without those is identity that cannot prove a
                // filing or name an actor, so this folder is recorded as a count instead -
                // the same treatment a folder above the budget gets.
                return null;
            }

            List<CensusItem> walked = new(expectedCount);
            HashSet<string> seen = new(StringComparer.Ordinal);
            while (walked.Count < expectedCount)
            {
                int wanted = Math.Min(CensusTableRowBatch, expectedCount - walked.Count);
                object? batch;
                try
                {
                    batch = t.GetArray(wanted);
                }
                catch (Exception ex) when (OutlookAI.Core.Com.OutlookComSession.IsComCallFailure(ex))
                {
                    return null;
                }

                // Shape is checked rather than assumed, and the check lives in the pure
                // half so a CI test can reach it: no non-live test can execute a single
                // line of this method.
                if (!CensusTableRow.TryReadBlock(batch, wanted, columns.ColumnCount, out Array? rows)
                    || !CensusTableRow.ProjectRows(rows!, columns, walked, seen))
                {
                    return null;
                }
            }

            // The folder promised expectedCount rows and delivered them; if the table still
            // has more, the folder grew while it was being read.
            try
            {
                if (!(bool)t.EndOfTable)
                {
                    return null;
                }
            }
            catch (Exception ex) when (OutlookAI.Core.Com.OutlookComSession.IsComCallFailure(ex))
            {
                return null;
            }

            return ConfirmUnchanged(items, expectedCount) ? walked : null;
        }
        finally
        {
            Release(table);
        }
    }

    /// <summary>
    /// Re-reads the folder's count and says whether it is still what the walk was based on.
    /// A count that moved means the item list is a mix of two moments.
    /// </summary>
    private static bool ConfirmUnchanged(dynamic items, int expectedCount)
    {
        try
        {
            return (int)items.Count == expectedCount;
        }
        catch (Exception ex) when (ex is COMException or RuntimeBinderException or InvalidCastException)
        {
            return false;
        }
    }

    /// <summary>
    /// Adds one census column to a table, trying each accepted spelling in turn. Silent on
    /// failure by design: whether the column actually landed is decided afterwards by
    /// <see cref="MapCensusColumns"/> reading the table back, which is the only answer that
    /// cannot disagree with what the rows contain.
    /// </summary>
    private static void AddCensusColumn(dynamic table, IReadOnlyList<string> spellings)
    {
        foreach (string spelling in spellings)
        {
            dynamic? columns = null;
            try
            {
                columns = table.Columns;
                columns!.Add(spelling);
                return;
            }
            catch (Exception ex) when (OutlookAI.Core.Com.OutlookComSession.IsComCallFailure(ex))
            {
            }
            finally
            {
                Release(columns);
            }
        }
    }

    /// <summary>
    /// Reads the table's column names ONCE and hands them to the pure mapper. One pass
    /// rather than one lookup per column, because <c>Columns[i].Name</c> is a round trip
    /// and this method is the reason a folder costs a fixed handful of them instead of a
    /// handful per column per row.
    /// </summary>
    private static CensusColumnMap MapCensusColumns(dynamic table)
    {
        dynamic? columns = null;
        try
        {
            columns = table.Columns;
            int count = (int)columns!.Count;
            List<string> names = new(count);
            for (int i = 1; i <= count; i++)
            {
                dynamic? column = null;
                try
                {
                    column = columns[i];
                    names.Add((string)column!.Name);
                }
                finally
                {
                    Release(column);
                }
            }

            return CensusTableRow.MapColumns(names);
        }
        catch (Exception ex) when (OutlookAI.Core.Com.OutlookComSession.IsComCallFailure(ex))
        {
            // An unusable map, which the caller turns into a count-only folder.
            return new CensusColumnMap(-1, -1, -1, -1, 0);
        }
        finally
        {
            Release(columns);
        }
    }

    /// <summary>
    /// Test folders a cleanup guard may legitimately FAIL over: any that still holds
    /// something, or still lives outside Deleted Items. <paramref name="wedgedEmpty"/>
    /// receives the rest - EMPTY folders stranded in Deleted Items by the documented
    /// same-session <c>Folders.Remove</c> limitation, which <see cref="DeleteTestFolders"/>
    /// deliberately tolerates and reports rather than failing on. Read-only.
    /// <para>
    /// Exists because asserting the RAW test-folder count fails the suite over exactly
    /// the remnant the cleanup helper just decided was acceptable - a test pinning an
    /// Outlook limitation instead of a contract, which is a flake by construction (it
    /// depends on whether Outlook has restarted since the folder was created).
    /// </para>
    /// </summary>
    public static int CountLiveTestFolders(string storeDisplayName, out int wedgedEmpty)
    {
        wedgedEmpty = CountWedgedEmptyTestFolders(storeDisplayName, out int live);
        return live;
    }

    /// <summary>
    /// Splits the remaining test folders into the ones that still matter (anywhere but
    /// Deleted Items, or holding items) and the empty ones wedged in Deleted Items by the
    /// documented same-session Folders.Remove limitation. Read-only.
    /// </summary>
    private static int CountWedgedEmptyTestFolders(string storeDisplayName, out int liveTestFolders)
    {
        (int Wedged, int Live) counts = RunSta(() =>
        {
            dynamic app = CreateOutlookApplication();
            dynamic? ns = null;
            dynamic? stores = null;
            dynamic? store = null;
            dynamic? root = null;
            dynamic? deleted = null;
            try
            {
                ns = app.GetNamespace("MAPI");
                stores = ns.Stores;
                store = FindStore(stores, storeDisplayName)
                    ?? throw new InvalidOperationException("Store not found for the test-folder check.");
                root = store.GetRootFolder();
                deleted = store.GetDefaultFolder(3);
                string deletedId = (string)deleted.EntryID;

                List<(string EntryId, string Name, int Depth)> inDeleted = new();
                CollectTestFolders(deleted, inDeleted, 0);
                List<(string EntryId, string Name, int Depth)> all = new();
                CollectTestFolders(root, all, 0);

                HashSet<string> deletedIds = new(inDeleted.Select(m => m.EntryId), StringComparer.OrdinalIgnoreCase);
                int wedged = 0;
                int live = 0;
                foreach ((string entryId, string _, int _) in all)
                {
                    dynamic? folder = null;
                    try
                    {
                        folder = ns.GetFolderFromID(entryId);
                        bool empty = (int)folder.Items.Count == 0 && (int)folder.Folders.Count == 0;
                        if (deletedIds.Contains(entryId) && empty)
                        {
                            wedged++;
                        }
                        else
                        {
                            live++;
                        }
                    }
                    catch (Exception ex) when (ex is COMException or RuntimeBinderException)
                    {
                        live++; // Cannot prove it is harmless - treat it as live.
                    }
                    finally
                    {
                        Release(folder);
                    }
                }

                return (wedged, live);
            }
            finally
            {
                Release(deleted);
                Release(root);
                Release(store);
                Release(stores);
                Release(ns);
                Release(app);
            }
        });

        liveTestFolders = counts.Live;
        return counts.Wedged;
    }

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
        LiveStoreWriteGuard.Assert(storeDisplayName, StoreWriteKind.Folder, nameof(DeleteTestFolders));
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

        int wedged = CountWedgedEmptyTestFolders(storeDisplayName, out int liveTestFolders);
        if (liveTestFolders != 0)
        {
            throw new InvalidOperationException(
                "Test folders kept reappearing for the whole cleanup window - manual check required (S3).");
        }

        if (wedged > 0)
        {
            // DOCUMENTED OUTLOOK LIMITATION (see this method's remarks): a folder created
            // and removed inside ONE Outlook session can wedge in Deleted Items -
            // Folders.Remove keeps failing with the synchronization error until Outlook
            // restarts. What is left is EMPTY and sits in Deleted Items, so it holds no
            // test artifact and no real mail; failing the suite over it would only teach
            // everyone to ignore the cleanup guard. Reported, and swept by the next run.
            Console.WriteLine(
                $"[cleanup] {wedged} empty test folder(s) wedged in Deleted Items until Outlook restarts "
                + "(documented same-session limitation) - no items involved.");
            return total;
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

    /// <summary>
    /// Asks Outlook to flush the Outbox and fetch waiting mail, and swallows every refusal.
    /// <para>
    /// Exists so an arrival wait can keep asking. <c>Send()</c> only QUEUES; the single
    /// best-effort call this file already makes right after it is not enough, because
    /// <c>SendAndReceive</c> is documented asynchronous and reports nothing, so it can
    /// complete its fetch before the submission it triggered has been handed over. On a
    /// machine whose transport is a local sink, missing that window means waiting for
    /// Outlook's own send/receive schedule, which is half an hour by default.
    /// </para>
    /// <para>
    /// It creates and releases its own short-lived session, which is what every other
    /// operation in this file does, rather than holding one open across a wait.
    /// </para>
    /// </summary>
    internal static void RequestDelivery()
    {
        RunSta<object?>(() =>
        {
            dynamic? app = null;
            dynamic? session = null;
            try
            {
                app = CreateOutlookApplication();
                session = app.Session;
                session.SendAndReceive(false);
            }
            catch (COMException)
            {
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
            }
            finally
            {
                Release(session);
                Release(app);
            }

            return null;
        });
    }

    private static T RunSta<T>(Func<T> work, TimeSpan? budget = null)
    {
        TimeSpan join = budget ?? DefaultStaBudget;
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
        if (!thread.Join(join))
        {
            throw new TimeoutException(
                "Test mailer STA operation timed out after "
                + join.TotalSeconds.ToString("0", CultureInfo.InvariantCulture) + " s.");
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
