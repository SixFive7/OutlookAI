using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.CSharp.RuntimeBinder;
using OutlookAI.Core.Com;

namespace OutlookAI.RemediationTools;

/// <summary>
/// The COM half of the corpus generator, and deliberately the ONLY half that touches a
/// mailbox. Every decision it acts on - which store is acceptable, what item N looks like,
/// which items may be deleted, whether the dates it wrote are real - is made by the pure
/// types beside it (<see cref="CorpusSafety"/>, <see cref="CorpusPlan"/>,
/// <see cref="CorpusDateFidelity"/>, <see cref="CorpusManifest"/>), which the T1 tier pins
/// without Outlook. What is left here is the mechanics: open, create, save, read back,
/// delete, release.
/// <para>
/// The COM patterns are the project's proven ones, taken from
/// <c>LiveOutlookTestMailer</c> by way of <see cref="ComMailbox"/>: a dedicated STA thread,
/// explicit <see cref="Marshal.ReleaseComObject"/> on every reference, a
/// <c>GetTable</c> LIKE prefilter with an ordinal re-check for anything selective, and
/// EntryID-addressed writes. One difference from both, forced by the workload: a build of
/// tens of thousands of items runs inside ONE STA session with the folders resolved once,
/// because a session per item would spend more time creating Outlook Application objects
/// than writing mail.
/// </para>
/// </summary>
public static class ComCorpusMailbox
{
    /// <summary>PR_MESSAGE_DELIVERY_TIME - what MailItem.ReceivedTime and urn:schemas:httpmail:datereceived read.</summary>
    private const string PrMessageDeliveryTime = "http://schemas.microsoft.com/mapi/proptag/0x0E060040";

    /// <summary>PR_CLIENT_SUBMIT_TIME - what MailItem.SentOn and urn:schemas:httpmail:date read.</summary>
    private const string PrClientSubmitTime = "http://schemas.microsoft.com/mapi/proptag/0x00390040";

    /// <summary>PR_MESSAGE_FLAGS - MSGFLAG_READ 0x1, MSGFLAG_UNSENT 0x8.</summary>
    private const string PrMessageFlags = "http://schemas.microsoft.com/mapi/proptag/0x0E070003";

    private const int MsgFlagRead = 0x1;
    private const int MsgFlagUnsent = 0x8;

    /// <summary>Default-folder id for Junk Email; a PST often has no such default folder.</summary>
    private const int JunkFolderId = 23;

    /// <summary>How many rows a probe's verification table may walk before giving up.</summary>
    private const int ProbeTableRowCap = 2_000;

    /// <summary>How many teardown passes may run before it reports what is left rather than looping.</summary>
    private const int TeardownMaxPasses = 6;

    /// <summary>Progress during a build. Elapsed comes from a Stopwatch, never the wall clock.</summary>
    /// <param name="Created">Items created so far.</param>
    /// <param name="Skipped">Ordinals already present in the manifest.</param>
    /// <param name="Failed">Items whose creation threw.</param>
    /// <param name="Remaining">Ordinals still to do.</param>
    /// <param name="BodyBytesWritten">Sum of body lengths written so far.</param>
    /// <param name="Elapsed">Monotonic elapsed time since the build started.</param>
    public sealed record BuildProgress(int Created, int Skipped, int Failed, int Remaining, long BodyBytesWritten, TimeSpan Elapsed);

    /// <summary>What a build did.</summary>
    /// <param name="Created">Items created.</param>
    /// <param name="Skipped">Ordinals already present.</param>
    /// <param name="Failed">Items whose creation threw.</param>
    /// <param name="BodyBytesWritten">Sum of body lengths written.</param>
    /// <param name="Elapsed">Monotonic elapsed time.</param>
    /// <param name="FirstError">The first failure's message, so a run that failed everywhere says why once.</param>
    public sealed record BuildOutcome(int Created, int Skipped, int Failed, long BodyBytesWritten, TimeSpan Elapsed, string? FirstError);

    /// <summary>What a teardown did.</summary>
    /// <param name="Considered">Items examined.</param>
    /// <param name="Deleted">Items deleted.</param>
    /// <param name="RefusedByRule">Items the two-key rule declined to delete.</param>
    /// <param name="AlreadyGone">Manifest entries whose item no longer exists.</param>
    /// <param name="Failed">Deletes that threw.</param>
    /// <param name="FoldersRemoved">Builder-created folders removed.</param>
    /// <param name="RemainingInStore">Corpus items a final read-only scan still found.</param>
    public sealed record TeardownOutcome(
        int Considered, int Deleted, int RefusedByRule, int AlreadyGone, int Failed, int FoldersRemoved, int RemainingInStore);

    /// <summary>One corpus item found by a read-only scan.</summary>
    /// <param name="Ordinal">Ordinal parsed out of the subject.</param>
    /// <param name="EntryId">Its EntryID right now.</param>
    /// <param name="FolderId">The default-folder id it was found under, or 0 for a builder-created folder.</param>
    public sealed record ScanRow(int Ordinal, string EntryId, int FolderId);

    /// <summary>
    /// Reads the four facts <see cref="CorpusSafety.EvaluateStore"/> judges a store on.
    /// Read-only. Each fact is read in its own try block so one unreadable property leaves
    /// the others intact and the verdict is "unprovable" rather than an exception.
    /// </summary>
    public static CorpusStoreFacts ReadStoreFacts(string storeDisplayName)
    {
        return RunSta(
            () =>
            {
                dynamic app = CreateOutlookApplication();
                dynamic? ns = null;
                dynamic? stores = null;
                dynamic? store = null;
                try
                {
                    ns = app.GetNamespace("MAPI");
                    stores = ns.Stores;
                    store = FindStore(stores, storeDisplayName);
                    if (store == null)
                    {
                        return new CorpusStoreFacts(null, null, null, null);
                    }

                    return new CorpusStoreFacts(
                        TryRead<string>(() => (string)store!.DisplayName),
                        TryReadStruct(() => (bool)store!.IsDataFileStore),
                        TryReadStruct(() => (int)store!.ExchangeStoreType),
                        TryRead<string>(() => (string)store!.FilePath));
                }
                finally
                {
                    Release(store);
                    Release(stores);
                    Release(ns);
                    Release(app);
                }
            },
            TimeSpan.FromMinutes(3));
    }

    /// <summary>
    /// Walks <see cref="CorpusDateFidelity.Ladder"/> against the store's Inbox, one
    /// throwaway item per rung, and reports what each rung actually achieved. Every probe
    /// item is deleted before this returns, by the same two-key rule as the teardown.
    /// <para>
    /// A rung that reads back displaced by exactly the machine's UTC offset is re-tried
    /// once with a pre-compensated write, and it is the RE-TRY that is reported - so the
    /// caller never has to reason about the PropertyAccessor's local-time conversion, it
    /// only sees whether the corrected write landed.
    /// </para>
    /// </summary>
    public static IReadOnlyList<CorpusDateProbe> ProbeDateFidelity(string storeDisplayName, string corpusId, DateTime requestedUtc)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(corpusId);
        DateTime requested = DateTime.SpecifyKind(requestedUtc, DateTimeKind.Utc);
        TimeSpan localOffset = TimeZoneInfo.Local.GetUtcOffset(requested);

        return RunSta(
            () =>
            {
                dynamic app = CreateOutlookApplication();
                dynamic? ns = null;
                dynamic? stores = null;
                dynamic? store = null;
                dynamic? folder = null;
                var probes = new List<CorpusDateProbe>();
                try
                {
                    ns = app.GetNamespace("MAPI");
                    stores = ns.Stores;
                    store = FindStore(stores, storeDisplayName)
                        ?? throw new InvalidOperationException("Store not found for the date probe.");
                    string storeId = (string)store.StoreID;
                    folder = store.GetDefaultFolder(6); // Inbox - always present in a PST

                    foreach (CorpusDateWriteMethod method in CorpusDateFidelity.Ladder)
                    {
                        CorpusDateProbe first = RunOneProbe(ns!, folder!, storeId, corpusId, method, requested, requested);
                        CorpusDateOffsetVerdict verdict =
                            CorpusDateFidelity.ClassifyOffset(requested, first.ReadBackReceivedUtc, localOffset);
                        if (verdict != CorpusDateOffsetVerdict.LocalOffsetApplied)
                        {
                            probes.Add(first);
                            continue;
                        }

                        DateTime compensated = CorpusDateFidelity.CompensatedWriteValue(
                            requested, verdict, localOffset, first.ReadBackReceivedUtc!.Value);
                        probes.Add(RunOneProbe(ns!, folder!, storeId, corpusId, method, requested, compensated));
                    }

                    return (IReadOnlyList<CorpusDateProbe>)probes;
                }
                finally
                {
                    Release(folder);
                    Release(store);
                    Release(stores);
                    Release(ns);
                    Release(app);
                }
            },
            TimeSpan.FromMinutes(10));
    }

    /// <summary>
    /// Creates the missing items of <paramref name="plan"/> up to <paramref name="itemCount"/>,
    /// in one STA session. <paramref name="record"/> is called with each created item BEFORE
    /// the next one is started, so an interrupted build leaves a manifest that describes
    /// everything it managed to create bar at most one item.
    /// </summary>
    /// <param name="plan">The corpus shape.</param>
    /// <param name="storeDisplayName">Target store; already vetted by <see cref="CorpusSafety.EvaluateStore"/>.</param>
    /// <param name="itemCount">Build ordinals 1..itemCount, skipping those the manifest already holds.</param>
    /// <param name="method">The date-write rung the probe verified.</param>
    /// <param name="writeShift">Pre-compensation added to every date written; <see cref="TimeSpan.Zero"/> unless the probe found an offset.</param>
    /// <param name="manifest">What already exists; also receives created folders.</param>
    /// <param name="record">Called with every created item; must persist it before returning.</param>
    /// <param name="recordFolder">Called with every folder the builder had to create.</param>
    /// <param name="progress">Called every <paramref name="progressEvery"/> items.</param>
    /// <param name="progressEvery">Progress interval in items.</param>
    public static BuildOutcome Build(
        CorpusPlan plan,
        string storeDisplayName,
        int itemCount,
        CorpusDateWriteMethod method,
        TimeSpan writeShift,
        CorpusManifest manifest,
        Action<CorpusManifestItem> record,
        Action<CorpusManifestFolder> recordFolder,
        Action<BuildProgress> progress,
        int progressEvery)
    {
        ArgumentNullException.ThrowIfNull(plan);
        ArgumentNullException.ThrowIfNull(manifest);
        ArgumentNullException.ThrowIfNull(record);
        ArgumentNullException.ThrowIfNull(recordFolder);
        ArgumentNullException.ThrowIfNull(progress);

        List<int> todo = manifest.MissingOrdinals(itemCount).ToList();
        int skipped = itemCount - todo.Count;

        // Monotonic: this number is reported as build throughput and compared against later
        // runs, so it must come from a clock that only moves forward.
        Stopwatch elapsed = Stopwatch.StartNew();

        return RunSta<BuildOutcome>(
            () =>
            {
                dynamic app = CreateOutlookApplication();
                dynamic? ns = null;
                dynamic? stores = null;
                dynamic? store = null;
                var folderItems = new Dictionary<int, dynamic>();
                var folders = new Dictionary<int, dynamic>();
                int created = 0;
                int failed = 0;
                long bytes = 0;
                string? firstError = null;
                try
                {
                    ns = app.GetNamespace("MAPI");
                    stores = ns.Stores;
                    store = FindStore(stores, storeDisplayName)
                        ?? throw new InvalidOperationException("Store not found for the corpus build.");
                    // Read before the first item is created: a store that cannot answer
                    // this is a store nothing should be written into, and finding that out
                    // after 40 000 items would be finding it out too late.
                    _ = (string)store.StoreID;

                    foreach (int ordinal in todo)
                    {
                        CorpusItemSpec spec = plan.Describe(ordinal);
                        if (!folderItems.TryGetValue(spec.FolderId, out dynamic? items))
                        {
                            dynamic resolved = ResolveFolder(store!, spec.FolderId, manifest, recordFolder);
                            folders[spec.FolderId] = resolved;
                            items = resolved.Items;
                            folderItems[spec.FolderId] = items!;
                        }

                        dynamic? mail = null;
                        try
                        {
                            mail = items!.Add(0); // olMailItem
                            mail.Subject = spec.Subject;
                            mail.Body = plan.BuildBody(spec);
                            mail.UnRead = !spec.IsRead;
                            mail.Save();

                            DateTime? readBack = ApplyDates(
                                mail!, method, spec.ReceivedUtc + writeShift, spec.SentUtc + writeShift, spec.IsRead);
                            string entryId = (string)mail.EntryID;
                            var line = new CorpusManifestItem(
                                ordinal,
                                entryId,
                                spec.FolderId,
                                spec.BodyBytes,
                                readBack == null ? null : CorpusManifest.FormatUtc(readBack.Value));
                            manifest.Add(line);
                            record(line);
                            created++;
                            bytes += spec.BodyBytes;
                        }
                        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                        {
                            failed++;
                            firstError ??= ex.Message;
                        }
                        finally
                        {
                            Release(mail);
                        }

                        if ((created + failed) % progressEvery == 0)
                        {
                            progress(new BuildProgress(
                                created, skipped, failed, todo.Count - created - failed, bytes, elapsed.Elapsed));
                        }
                    }

                    return new BuildOutcome(created, skipped, failed, bytes, elapsed.Elapsed, firstError);
                }
                finally
                {
                    foreach (dynamic items in folderItems.Values)
                    {
                        Release(items);
                    }

                    foreach (dynamic folder in folders.Values)
                    {
                        Release(folder);
                    }

                    Release(store);
                    Release(stores);
                    Release(ns);
                    Release(app);
                }
            },
            timeout: null);
    }

    /// <summary>
    /// Removes exactly what the manifest records, and then what those removals turned into.
    /// <para>
    /// TWO PHASES, and the second is not optional. <c>MailItem.Delete()</c> on an item in
    /// the Inbox, Sent Items or Junk Email SOFT-DELETES it: the item moves to Deleted Items
    /// and is issued a NEW EntryID. So a purely manifest-driven teardown would leave a copy
    /// of every item it "deleted" sitting in Deleted Items under an id no manifest has ever
    /// seen. Phase 2 therefore READS the store, collects the current EntryIDs of items whose
    /// subject parses as this corpus's, and uses that freshly-enumerated set as the
    /// allowlist for a second delete pass - still two keys, still no pattern matching, and
    /// still confined to the one store the caller allowlisted. The pair repeats until a
    /// scan comes back empty or <see cref="TeardownMaxPasses"/> is spent.
    /// </para>
    /// </summary>
    public static TeardownOutcome Teardown(string storeDisplayName, string corpusId, CorpusManifest manifest)
    {
        ArgumentNullException.ThrowIfNull(manifest);
        HashSet<string> manifestIds = CorpusSafety.BuildEntryIdAllowlist(manifest.EntryIds);

        return RunSta(
            () =>
            {
                dynamic app = CreateOutlookApplication();
                dynamic? ns = null;
                dynamic? stores = null;
                dynamic? store = null;
                int considered = 0;
                int deleted = 0;
                int refused = 0;
                int gone = 0;
                int failed = 0;
                int foldersRemoved = 0;
                int remaining;
                try
                {
                    ns = app.GetNamespace("MAPI");
                    stores = ns.Stores;
                    store = FindStore(stores, storeDisplayName)
                        ?? throw new InvalidOperationException("Store not found for the corpus teardown.");
                    string storeId = (string)store.StoreID;

                    // Phase 1: the manifest's own ids.
                    foreach (string entryId in manifestIds)
                    {
                        considered++;
                        DeleteVerdict verdict = DeleteOne(ns!, storeId, entryId, manifestIds, corpusId);
                        switch (verdict)
                        {
                            case DeleteVerdict.Deleted: deleted++; break;
                            case DeleteVerdict.Refused: refused++; break;
                            case DeleteVerdict.Gone: gone++; break;
                            default: failed++; break;
                        }
                    }

                    // Phase 2: whatever those deletes turned into, plus anything a lost
                    // manifest line left behind.
                    for (int pass = 0; pass < TeardownMaxPasses; pass++)
                    {
                        List<ScanRow> found = ScanStore(store!, manifest, corpusId);
                        if (found.Count == 0)
                        {
                            break;
                        }

                        HashSet<string> passIds = CorpusSafety.BuildEntryIdAllowlist(found.Select(r => r.EntryId));
                        foreach (string entryId in passIds)
                        {
                            considered++;
                            DeleteVerdict verdict = DeleteOne(ns!, storeId, entryId, passIds, corpusId);
                            switch (verdict)
                            {
                                case DeleteVerdict.Deleted: deleted++; break;
                                case DeleteVerdict.Refused: refused++; break;
                                case DeleteVerdict.Gone: gone++; break;
                                default: failed++; break;
                            }
                        }
                    }

                    foldersRemoved = RemoveCreatedFolders(ns!, storeId, manifest);
                    remaining = ScanStore(store!, manifest, corpusId).Count;
                }
                finally
                {
                    Release(store);
                    Release(stores);
                    Release(ns);
                    Release(app);
                }

                return new TeardownOutcome(considered, deleted, refused, gone, failed, foldersRemoved, remaining);
            },
            timeout: null);
    }

    /// <summary>
    /// READ-ONLY: every item in the store whose subject parses as belonging to
    /// <paramref name="corpusId"/>. This is what <c>corpus-reindex</c> reports and what the
    /// teardown's second phase consumes; it never deletes anything itself.
    /// </summary>
    public static IReadOnlyList<ScanRow> Scan(string storeDisplayName, string corpusId, CorpusManifest? manifest)
    {
        return RunSta(
            () =>
            {
                dynamic app = CreateOutlookApplication();
                dynamic? ns = null;
                dynamic? stores = null;
                dynamic? store = null;
                try
                {
                    ns = app.GetNamespace("MAPI");
                    stores = ns.Stores;
                    store = FindStore(stores, storeDisplayName)
                        ?? throw new InvalidOperationException("Store not found for the corpus scan.");
                    return (IReadOnlyList<ScanRow>)ScanStore(store!, manifest, corpusId);
                }
                finally
                {
                    Release(store);
                    Release(stores);
                    Release(ns);
                    Release(app);
                }
            },
            timeout: null);
    }

    private enum DeleteVerdict
    {
        Deleted,
        Refused,
        Gone,
        Failed,
    }

    private static DeleteVerdict DeleteOne(dynamic ns, string storeId, string entryId, ISet<string> allowlist, string corpusId)
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
                return DeleteVerdict.Gone;
            }

            string? subject = TryRead<string>(() => (string)item!.Subject);

            // The two-key rule, re-evaluated against the subject as it reads RIGHT NOW -
            // not against what the manifest thought it was when the item was created.
            if (!CorpusSafety.MayDelete(entryId, subject, allowlist, corpusId))
            {
                return DeleteVerdict.Refused;
            }

            item!.Delete();
            return DeleteVerdict.Deleted;
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
        {
            return DeleteVerdict.Failed;
        }
        finally
        {
            Release(item);
        }
    }

    private static List<ScanRow> ScanStore(dynamic store, CorpusManifest? manifest, string corpusId)
    {
        var rows = new List<ScanRow>();
        var folderIds = new List<int> { 6, 5, 23, 3 }; // Deleted Items LAST: everything else drains into it
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

                CollectCorpusItems(folder!, folderId, corpusId, rows);
            }
            finally
            {
                Release(folder);
            }
        }

        foreach (CorpusManifestFolder created in manifest?.Folders ?? (IReadOnlyList<CorpusManifestFolder>)Array.Empty<CorpusManifestFolder>())
        {
            dynamic? folder = null;
            try
            {
                folder = store.Session.GetFolderFromID(created.EntryId);
                CollectCorpusItems(folder!, 0, corpusId, rows);
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                // Folder already gone - nothing to collect.
            }
            finally
            {
                Release(folder);
            }
        }

        return rows;
    }

    /// <summary>
    /// Adds every item of one folder whose subject parses as this corpus's. A
    /// <c>GetTable</c> LIKE prefilter narrows the walk - the bracket-free fragment is a
    /// SUPERSET of real matches by construction - and the authoritative decision is always
    /// the ordinal parse on the row's own subject.
    /// </summary>
    private static void CollectCorpusItems(dynamic folder, int folderId, string corpusId, List<ScanRow> rows)
    {
        dynamic? table = null;
        try
        {
            table = folder.GetTable(
                "@SQL=\"urn:schemas:httpmail:subject\" LIKE '%" + CorpusPlan.DaslCountFragment + "%'");
            while (!(bool)table.EndOfTable)
            {
                dynamic? row = null;
                try
                {
                    row = table.GetNextRow();
                    object[] values = (object[])row.GetValues();
                    string? entryId = values.Length > 0 ? values[0] as string : null;
                    string? subject = values.Length > 1 ? values[1] as string : null;
                    if (entryId != null && CorpusPlan.TryParseOrdinal(subject, corpusId, out int ordinal))
                    {
                        rows.Add(new ScanRow(ordinal, entryId, folderId));
                    }
                }
                finally
                {
                    Release(row);
                }
            }
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
        {
            // A folder that cannot be tabled is reported by its absence from the scan, and
            // the teardown's final remaining-count is what fails the run loudly.
        }
        finally
        {
            Release(table);
        }
    }

    /// <summary>
    /// Removes folders the builder created, deepest names last, and only when BOTH keys
    /// agree: the folder's EntryID is one the manifest recorded creating AND its name
    /// ordinal-contains <see cref="CorpusManifest.CreatedFolderPrefix"/>.
    /// </summary>
    private static int RemoveCreatedFolders(dynamic ns, string storeId, CorpusManifest manifest)
    {
        int removed = 0;
        foreach (CorpusManifestFolder record in manifest.Folders)
        {
            dynamic? folder = null;
            try
            {
                folder = ns.GetFolderFromID(record.EntryId, storeId);
                string name = (string)folder!.Name;
                if (!name.Contains(CorpusManifest.CreatedFolderPrefix, StringComparison.Ordinal))
                {
                    continue;
                }

                if ((int)folder.Items.Count != 0)
                {
                    continue; // still holds something - leave it and let the count report it
                }

                folder.Delete();
                removed++;
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                // Already gone, or wedged by the documented same-session Folders.Remove
                // limitation; either way it is reported by not being counted.
            }
            finally
            {
                Release(folder);
            }
        }

        return removed;
    }

    /// <summary>
    /// The store's default folder for <paramref name="folderId"/>, or - when the store has
    /// no such default folder, which is the normal case for Junk Email in a PST - a folder
    /// under the store root created for the purpose and recorded in the manifest so the
    /// teardown can remove it.
    /// </summary>
    private static dynamic ResolveFolder(dynamic store, int folderId, CorpusManifest manifest, Action<CorpusManifestFolder> recordFolder)
    {
        try
        {
            return store.GetDefaultFolder(folderId);
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
        {
            // Fall through to the substitute folder.
        }

        string name = CorpusManifest.CreatedFolderPrefix + "-" + (folderId == JunkFolderId ? "Junk" : folderId.ToString());
        foreach (CorpusManifestFolder known in manifest.Folders)
        {
            if (known.FolderId == folderId)
            {
                try
                {
                    return store.Session.GetFolderFromID(known.EntryId);
                }
                catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                {
                    break; // recorded folder is gone - make a fresh one
                }
            }
        }

        dynamic? root = null;
        try
        {
            root = store.GetRootFolder();
            dynamic created = root.Folders.Add(name);
            var record = new CorpusManifestFolder((string)created.EntryID, name, folderId);
            manifest.Add(record);
            recordFolder(record);
            return created;
        }
        finally
        {
            Release(root);
        }
    }

    /// <summary>
    /// Writes the dates for one already-saved item using <paramref name="method"/> and
    /// returns the received time the item reports afterwards (null when it cannot be read).
    /// Throws when the rung fails, so a build cannot silently degrade to undated items
    /// halfway through - the probe already established that this rung works on this store.
    /// </summary>
    private static DateTime? ApplyDates(
        dynamic mail, CorpusDateWriteMethod method, DateTime receivedUtc, DateTime sentUtc, bool isRead)
    {
        switch (method)
        {
            case CorpusDateWriteMethod.None:
                break;

            case CorpusDateWriteMethod.ObjectModel:
                // SentOn is read-only in the object model, so this rung carries only the
                // delivery half. Deliberately not wrapped in a try: the probe proved this
                // works here, so a failure now is a real change of state, not noise.
                mail.ReceivedTime = receivedUtc;
                mail.Save();
                break;

            default:
                dynamic? accessor = null;
                try
                {
                    accessor = mail.PropertyAccessor;
                    accessor.SetProperty(PrMessageDeliveryTime, receivedUtc);
                    accessor.SetProperty(PrClientSubmitTime, sentUtc);
                    if (method == CorpusDateWriteMethod.PropertyAccessorWithFlags)
                    {
                        // The whole value is written, which is what clears MSGFLAG_UNSENT
                        // (0x8) and makes the item read as delivered mail rather than as a
                        // draft. MSGFLAG_READ carries the state the PLAN asked for: forcing
                        // it on would quietly destroy the unread population the corpus is
                        // supposed to contain, and an unread-only filter would then select
                        // nothing at all.
                        accessor.SetProperty(PrMessageFlags, isRead ? MsgFlagRead : 0);
                    }
                }
                finally
                {
                    Release(accessor);
                }

                mail.Save();
                break;
        }

        return TryReadStruct(() => ((DateTime)mail.ReceivedTime).ToUniversalTime());
    }

    /// <summary>
    /// One rung, one throwaway item: create, date, save, re-open by EntryID, read the date
    /// back, then ask a DASL restriction on either side of the instant whether it selects
    /// the item. Deletes the probe before returning, by the two-key rule.
    /// </summary>
    private static CorpusDateProbe RunOneProbe(
        dynamic ns,
        dynamic folder,
        string storeId,
        string corpusId,
        CorpusDateWriteMethod method,
        DateTime requestedUtc,
        DateTime writeUtc)
    {
        dynamic? items = null;
        dynamic? mail = null;
        string? entryId = null;
        try
        {
            items = folder.Items;
            mail = items.Add(0);

            // A probe is an ordinary corpus item as far as every guard is concerned: it
            // carries both tags and an ordinal, so it is deletable by the same rule and
            // findable by the same scan if this process dies mid-probe.
            string subject = CorpusPlan.SubjectTag + CorpusPlan.CorpusTagOpen + corpusId + "#"
                + int.MaxValue.ToString(System.Globalization.CultureInfo.InvariantCulture) + "] date probe "
                + method;
            mail.Subject = subject;
            mail.Body = "date fidelity probe";
            mail.Save();
            entryId = (string)mail.EntryID;

            DateTime? readBack;
            try
            {
                readBack = ApplyDates(mail!, method, writeUtc, writeUtc, isRead: true);
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                return new CorpusDateProbe(method, requestedUtc, writeUtc, null, false, false, ex.Message);
            }

            Release(mail);
            mail = null;

            // Re-open by EntryID rather than trusting the handle we just wrote through: the
            // question is what the STORE holds, not what the in-memory item says.
            dynamic? reopened = null;
            try
            {
                reopened = ns.GetItemFromID(entryId, storeId);
                readBack = TryReadStruct(() => ((DateTime)reopened!.ReceivedTime).ToUniversalTime());
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                return new CorpusDateProbe(method, requestedUtc, writeUtc, null, false, false, ex.Message);
            }
            finally
            {
                Release(reopened);
            }

            bool selected = TableContains(folder, DateWindowFilter(requestedUtc.AddDays(-1), requestedUtc.AddDays(1)), entryId!);
            bool excluded = !TableContains(folder, DateWindowFilter(requestedUtc.AddDays(2), null), entryId!);
            return new CorpusDateProbe(method, requestedUtc, writeUtc, readBack, selected, excluded, null);
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
        {
            return new CorpusDateProbe(method, requestedUtc, writeUtc, null, false, false, ex.Message);
        }
        finally
        {
            Release(mail);
            Release(items);
            if (entryId != null)
            {
                HashSet<string> allowlist = CorpusSafety.BuildEntryIdAllowlist(new[] { entryId });
                DeleteOne(ns, storeId, entryId, allowlist, corpusId);
            }
        }
    }

    /// <summary>
    /// A DASL restriction on the received date, narrowed to corpus subjects so a probe
    /// never has to walk a whole folder. Date literals come from
    /// <see cref="DaslDateLiteral"/> - the year-first form, because Outlook parses these in
    /// the MACHINE locale and a day-first literal silently returns the wrong rows.
    /// </summary>
    private static string DateWindowFilter(DateTime fromUtc, DateTime? toUtc)
    {
        string filter = "@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%" + CorpusPlan.DaslCountFragment + "%')"
            + " AND (\"urn:schemas:httpmail:datereceived\" >= '" + DaslDateLiteral.FormatUtc(fromUtc) + "')";
        if (toUtc != null)
        {
            filter += " AND (\"urn:schemas:httpmail:datereceived\" < '" + DaslDateLiteral.FormatUtc(toUtc.Value) + "')";
        }

        return filter;
    }

    private static bool TableContains(dynamic folder, string filter, string entryId)
    {
        dynamic? table = null;
        try
        {
            table = folder.GetTable(filter);
            int walked = 0;
            while (!(bool)table.EndOfTable && walked < ProbeTableRowCap)
            {
                dynamic? row = null;
                try
                {
                    row = table.GetNextRow();
                    walked++;
                    object[] values = (object[])row.GetValues();
                    if (values.Length > 0 && values[0] is string id
                        && string.Equals(id, entryId, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
                finally
                {
                    Release(row);
                }
            }

            return false;
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
        {
            return false;
        }
        finally
        {
            Release(table);
        }
    }

    private static T? TryRead<T>(Func<T> read)
        where T : class
    {
        try
        {
            return read();
        }
        catch (Exception ex) when (ex is COMException or RuntimeBinderException or InvalidCastException)
        {
            return null;
        }
    }

    private static T? TryReadStruct<T>(Func<T> read)
        where T : struct
    {
        try
        {
            return read();
        }
        catch (Exception ex) when (ex is COMException or RuntimeBinderException or InvalidCastException)
        {
            return null;
        }
    }

    private static dynamic? FindStore(dynamic stores, string storeDisplayName)
    {
        int storeCount = stores.Count;
        for (int i = 1; i <= storeCount; i++)
        {
            dynamic candidate = stores[i];
            string? name = TryRead<string>(() => (string)candidate.DisplayName);
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

    /// <summary>
    /// Runs <paramref name="work"/> on a dedicated STA thread. <paramref name="timeout"/> is
    /// nullable and the BUILD passes null on purpose: a corpus of tens of thousands of items
    /// legitimately runs for hours, and a timeout that fires mid-build would abandon a
    /// half-written PST with no manifest line for the item in flight.
    /// </summary>
    private static T RunSta<T>(Func<T> work, TimeSpan? timeout)
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
            Name = "OutlookAI.Corpus.Sta",
        };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        if (timeout == null)
        {
            thread.Join();
        }
        else if (!thread.Join(timeout.Value))
        {
            throw new TimeoutException("Corpus STA operation timed out.");
        }

        if (failure != null)
        {
            throw new InvalidOperationException("Corpus operation failed.", failure);
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
