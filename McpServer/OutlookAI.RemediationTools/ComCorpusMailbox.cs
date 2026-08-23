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

    /// <summary>MSGFLAG_SUBMIT - "this message is queued for delivery". The bit that puts an item in the Outbox.</summary>
    private const int MsgFlagSubmit = 0x4;

    private const int MsgFlagUnsent = 0x8;

    /// <summary>Default-folder id for Junk Email; a PST often has no such default folder.</summary>
    private const int JunkFolderId = 23;

    /// <summary>Default-folder id for Drafts - where Outlook files an unsent item whatever folder it was added to.</summary>
    private const int DraftsFolderId = 16;

    /// <summary>
    /// Folders a corpus scan walks, Deleted Items LAST because every other folder drains
    /// into it when an item is soft-deleted.
    /// <para>
    /// DRAFTS (16) and the OUTBOX (4) are in this set and were missing from the first
    /// version, which was a real hole rather than an oversight to note quietly: items
    /// created by <c>Items.Add</c> + <c>Save</c> are UNSENT and Outlook files them in
    /// Drafts, so the 40 000 items of the first real build lived in exactly the two folders
    /// this scan did not look at. <c>corpus-reindex</c> - the recovery path for a lost
    /// manifest - would have reported ZERO items with 40 000 in the store, and the
    /// post-teardown "remaining" count would have said 0 for the same reason. It is the
    /// same lesson <see cref="ComMailbox.SweepFolderIds"/> already records about the Outbox:
    /// a folder nothing sweeps is a folder items can be stranded in indefinitely.
    /// </para>
    /// </summary>
    private static readonly int[] ScanFolderIds = { DraftsFolderId, 6, 5, JunkFolderId, 4, 3 };

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
    /// <param name="FolderId">
    /// The default-folder id it was found under. For an item in a folder the builder created
    /// because the store has no such default folder, this is the id that folder STANDS IN
    /// FOR - so a census can compare it against the plan, which only ever speaks in default
    /// folder ids.
    /// </param>
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
    /// Reads what <see cref="CorpusSafety.EvaluateProfile"/> judges the profile on: how many
    /// accounts exist, and how many of them deliver into the target store.
    /// <para>
    /// Delivery stores are compared by <c>StoreID</c>, never by display name - names are
    /// user-editable and a profile may mount two stores with the same one, so a name
    /// comparison could clear an account that does deliver into the target.
    /// </para>
    /// <para>
    /// An account whose <c>DeliveryStore</c> cannot be read is COUNTED rather than skipped.
    /// It is the difference between "no account delivers here" and "no account I could
    /// examine delivers here", and only the first is a proof.
    /// </para>
    /// </summary>
    public static CorpusProfileFacts ReadProfileFacts(string storeDisplayName)
    {
        return RunSta(
            () =>
            {
                dynamic app = CreateOutlookApplication();
                dynamic? ns = null;
                dynamic? stores = null;
                dynamic? store = null;
                dynamic? accounts = null;
                try
                {
                    ns = app.GetNamespace("MAPI");
                    stores = ns.Stores;
                    store = FindStore(stores, storeDisplayName);
                    string? targetStoreId = store == null ? null : TryRead<string>(() => (string)store!.StoreID);

                    accounts = ns.Accounts;
                    int? count = TryReadStruct(() => (int)accounts!.Count);
                    if (count == null)
                    {
                        return new CorpusProfileFacts(null, 0, 0);
                    }

                    int delivering = 0;
                    int unreadable = 0;
                    for (int i = 1; i <= count.Value; i++)
                    {
                        dynamic? account = null;
                        dynamic? deliveryStore = null;
                        try
                        {
                            account = accounts![i];
                            deliveryStore = account!.DeliveryStore;
                            string? deliveryStoreId = deliveryStore == null
                                ? null
                                : TryRead<string>(() => (string)deliveryStore!.StoreID);
                            if (deliveryStoreId == null || targetStoreId == null)
                            {
                                unreadable++;
                            }
                            else if (string.Equals(deliveryStoreId, targetStoreId, StringComparison.OrdinalIgnoreCase))
                            {
                                delivering++;
                            }
                        }
                        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                        {
                            unreadable++;
                        }
                        finally
                        {
                            Release(deliveryStore);
                            Release(account);
                        }
                    }

                    return new CorpusProfileFacts(count, delivering, unreadable);
                }
                finally
                {
                    Release(accounts);
                    Release(store);
                    Release(stores);
                    Release(ns);
                    Release(app);
                }
            },
            TimeSpan.FromMinutes(3));
    }

    /// <summary>
    /// Walks <see cref="CorpusPlacement.Ladder"/> against the store's Inbox, one throwaway
    /// item per rung, and reports where each one actually ended up. Every probe item is
    /// deleted before this returns, by the same two-key rule as the teardown.
    /// <para>
    /// The Inbox is the right folder to probe: it is the folder a PST always has, and it is
    /// where the plan puts most of the corpus. A rung that can place an item in the Inbox
    /// can place one in Sent Items or Junk Email, because the obstacle is the item's unsent
    /// state and not the destination.
    /// </para>
    /// </summary>
    public static IReadOnlyList<CorpusPlacementProbe> ProbePlacement(string storeDisplayName, string corpusId)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(corpusId);
        return RunSta(
            () =>
            {
                dynamic app = CreateOutlookApplication();
                dynamic? ns = null;
                dynamic? stores = null;
                dynamic? store = null;
                dynamic? target = null;
                dynamic? drafts = null;
                var probes = new List<CorpusPlacementProbe>();
                try
                {
                    ns = app.GetNamespace("MAPI");
                    stores = ns.Stores;
                    store = FindStore(stores, storeDisplayName)
                        ?? throw new InvalidOperationException("Store not found for the placement probe.");
                    string storeId = (string)store.StoreID;
                    target = store.GetDefaultFolder(6);
                    drafts = store.GetDefaultFolder(DraftsFolderId);
                    string targetName = (string)target!.Name;
                    string targetFolderId = (string)target!.EntryID;

                    foreach (CorpusPlacementMethod method in CorpusPlacement.Ladder)
                    {
                        probes.Add(RunOnePlacementProbe(
                            ns!, target!, drafts!, storeId, targetFolderId, targetName, corpusId, method));
                    }

                    return (IReadOnlyList<CorpusPlacementProbe>)probes;
                }
                finally
                {
                    Release(drafts);
                    Release(target);
                    Release(store);
                    Release(stores);
                    Release(ns);
                    Release(app);
                }
            },
            TimeSpan.FromMinutes(10));
    }

    private static CorpusPlacementProbe RunOnePlacementProbe(
        dynamic ns,
        dynamic target,
        dynamic drafts,
        string storeId,
        string targetFolderId,
        string targetFolderName,
        string corpusId,
        CorpusPlacementMethod method)
    {
        dynamic? items = null;
        dynamic? mail = null;
        string? entryId = null;
        try
        {
            dynamic source = CorpusPlacement.CreatesInDrafts(method) ? drafts : target;
            items = source.Items;
            mail = items.Add(0);
            mail.Subject = ProbeSubject(corpusId, "placement " + method);
            mail.Body = "placement probe";
            mail.Save();

            ApplyMessageFlags(mail!, isRead: true, clearUnsent: CorpusPlacement.WritesSentFlag(method));

            if (CorpusPlacement.RequiresMove(method))
            {
                // Move issues a NEW EntryID, so everything after this point - including the
                // delete in the finally block - must use the moved item's id.
                dynamic moved = mail!.Move(target);
                Release(mail);
                mail = moved;
            }

            entryId = (string)mail!.EntryID;
            bool sentFlag = TryReadStruct(() => (bool)mail!.Sent) ?? false;
            Release(mail);
            mail = null;

            // Re-open by EntryID: the question is where the STORE put it, not what the
            // handle we just wrote through believes.
            string? parentId = null;
            string? parentName = null;
            dynamic? reopened = null;
            try
            {
                reopened = ns.GetItemFromID(entryId, storeId);
                dynamic? parent = null;
                try
                {
                    parent = reopened!.Parent;
                    parentId = TryRead<string>(() => (string)parent!.EntryID);
                    parentName = TryRead<string>(() => (string)parent!.Name);
                }
                finally
                {
                    Release(parent);
                }
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                return new CorpusPlacementProbe(
                    method, targetFolderName, false, false, sentFlag, null, ex.Message, true);
            }
            finally
            {
                Release(reopened);
            }

            bool parentMatches = parentId != null
                && string.Equals(parentId, targetFolderId, StringComparison.OrdinalIgnoreCase);

            // The decisive check. The freshness sweep enumerates a folder through its TABLE,
            // so an item the table does not carry does not exist as far as the measurement
            // is concerned - however correct its Parent looks.
            TableLookup lookup = TableFind(target, ProbeSubjectFilter(corpusId), entryId!);

            return new CorpusPlacementProbe(
                method,
                targetFolderName,
                parentMatches,
                lookup == TableLookup.Found,
                sentFlag,
                parentName,
                null,
                lookup != TableLookup.Inconclusive);
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
        {
            return new CorpusPlacementProbe(
                method, targetFolderName, false, false, false, null, ex.Message, true);
        }
        finally
        {
            Release(mail);
            Release(items);
            if (entryId != null)
            {
                DeleteOne(ns, storeId, entryId, CorpusSafety.BuildEntryIdAllowlist(new[] { entryId }), corpusId);
            }
        }
    }

    /// <summary>
    /// A bracket-free DASL LIKE restriction selecting ONE probe item: the corpus tag plus
    /// the reserved probe ordinal.
    /// <para>
    /// It used to select the whole corpus, and that is what broke the probe against a large
    /// folder: with ~22 000 corpus items already in the Inbox, the walk hit its row cap long
    /// before it reached the item it had just created, and "I stopped looking" was reported
    /// as "it is not there". A filter that names the item cannot have that failure mode, and
    /// the row cap goes back to being a bound on a runaway rather than a search budget.
    /// </para>
    /// </summary>
    private static string ProbeSubjectFilter(string corpusId)
        => "@SQL=" + "\"" + "urn:schemas:httpmail:subject" + "\"" + " LIKE '%"
            + CorpusPlan.DaslSubjectFragment(corpusId, CorpusPlan.ProbeOrdinal) + "%'";

    /// <summary>
    /// The subject every throwaway probe item carries: both tags plus the reserved probe
    /// ordinal, so a probe is deletable by exactly the same two-key rule as a corpus item,
    /// findable by the same scan if this process dies between creating one and deleting it,
    /// and selectable on its own by <see cref="ProbeSubjectFilter"/>.
    /// </summary>
    private static string ProbeSubject(string corpusId, string what)
        => CorpusPlan.SubjectTag + CorpusPlan.CorpusTagOpen + corpusId + "#"
            + CorpusPlan.ProbeOrdinal.ToString("D7", System.Globalization.CultureInfo.InvariantCulture) + "] " + what;

    /// <summary>
    /// Writes the read state, and - when the placement rung says so - clears MSGFLAG_UNSENT
    /// so the store stops filing the item as a draft. MSGFLAG_SUBMIT is ALWAYS cleared.
    /// <para>
    /// <b>Read-modify-write, and that is the fix.</b> This used to write the whole value:
    /// <c>MSGFLAG_READ</c> for a read item and <c>0</c> for an unread one. The first real
    /// build queued 5 532 items for delivery into the Outbox, and 5 532 is EXACTLY the
    /// number of items the plan for that shape marks unread - so whatever the store did with
    /// them, the unread ones are the population it did it to, and the read state is the only
    /// thing that distinguished them. Two things follow, and both are done here rather than
    /// one: the value is now derived from what the item already carries instead of replacing
    /// it wholesale (a blind write destroys MSGFLAG_HASATTACH, MSGFLAG_FROMME and anything
    /// else the store set), and MSGFLAG_SUBMIT - the bit that MEANS "queued for delivery" -
    /// is cleared explicitly on every item, whatever set it.
    /// </para>
    /// <para>
    /// <b>And the read state no longer goes through <c>MailItem.UnRead</c>.</b> The build
    /// used to set that property as well, on every item, before this ran; it is the object
    /// model's view of the same MSGFLAG_READ bit, so it was redundant, and it was applied to
    /// exactly the affected population. One property, one writer.
    /// </para>
    /// <para>
    /// MSGFLAG_UNSENT is cleared only for the rungs that say so, because it is the bit that
    /// decides where the item LIVES: clearing it on a control rung would make that rung stop
    /// being a control. Writing the read bit does not move an item, so it is safe on all
    /// four - and the corpus needs its unread population whichever rung places it, or an
    /// unread-only filter selects nothing.
    /// </para>
    /// </summary>
    private static void ApplyMessageFlags(dynamic mail, bool isRead, bool clearUnsent)
    {
        dynamic? accessor = null;
        try
        {
            accessor = mail.PropertyAccessor;
            int current = TryReadStruct(() => (int)accessor!.GetProperty(PrMessageFlags)) ?? MsgFlagUnsent;
            int wanted = current & ~MsgFlagSubmit;
            if (clearUnsent)
            {
                wanted &= ~MsgFlagUnsent;
            }

            wanted = isRead ? (wanted | MsgFlagRead) : (wanted & ~MsgFlagRead);
            if (wanted != current)
            {
                accessor.SetProperty(PrMessageFlags, wanted);
            }
        }
        finally
        {
            Release(accessor);
        }

        mail.Save();
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
    public static IReadOnlyList<CorpusDateProbe> ProbeDateFidelity(
        string storeDisplayName, string corpusId, DateTime requestedUtc, CorpusPlacementMethod placement)
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
                dynamic? drafts = null;
                var probes = new List<CorpusDateProbe>();
                try
                {
                    ns = app.GetNamespace("MAPI");
                    stores = ns.Stores;
                    store = FindStore(stores, storeDisplayName)
                        ?? throw new InvalidOperationException("Store not found for the date probe.");
                    string storeId = (string)store.StoreID;
                    folder = store.GetDefaultFolder(6); // Inbox - always present in a PST
                    drafts = store.GetDefaultFolder(DraftsFolderId);

                    foreach (CorpusDateWriteMethod method in CorpusDateFidelity.Ladder)
                    {
                        CorpusDateProbe first = RunOneProbe(
                            ns!, folder!, drafts!, storeId, corpusId, method, placement, requested, requested);
                        CorpusDateOffsetVerdict verdict =
                            CorpusDateFidelity.ClassifyOffset(requested, first.ReadBackReceivedUtc, localOffset);
                        if (verdict != CorpusDateOffsetVerdict.LocalOffsetApplied)
                        {
                            probes.Add(first);
                            continue;
                        }

                        DateTime compensated = CorpusDateFidelity.CompensatedWriteValue(
                            requested, verdict, localOffset, first.ReadBackReceivedUtc!.Value);
                        probes.Add(RunOneProbe(
                            ns!, folder!, drafts!, storeId, corpusId, method, placement, requested, compensated));
                    }

                    return (IReadOnlyList<CorpusDateProbe>)probes;
                }
                finally
                {
                    Release(drafts);
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
    /// <param name="dateMethod">The date-write rung the probe verified.</param>
    /// <param name="placement">The placement rung the probe verified - what makes items live where the plan says.</param>
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
        CorpusDateWriteMethod dateMethod,
        CorpusPlacementMethod placement,
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
                dynamic? draftsItems = null;
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

                    if (CorpusPlacement.CreatesInDrafts(placement))
                    {
                        dynamic draftsFolder = ResolveFolder(store!, DraftsFolderId, manifest, recordFolder);
                        folders[DraftsFolderId] = draftsFolder;
                        draftsItems = draftsFolder.Items;
                    }

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
                            // The order is load-bearing: flags and dates are written BEFORE
                            // any move, so the item that arrives in the target folder is
                            // already in its final state and is not re-filed as a draft on
                            // the way in.
                            mail = (CorpusPlacement.CreatesInDrafts(placement) ? draftsItems! : items!).Add(0);
                            mail.Subject = spec.Subject;
                            mail.Body = plan.BuildBody(spec);
                            mail.Save();

                            // The read state is written HERE and nowhere else. It used to be
                            // set through MailItem.UnRead as well, one line above the Save,
                            // and the population that got UnRead = true is exactly the
                            // population that ended up queued in the Outbox on the first
                            // real build - see ApplyMessageFlags.
                            ApplyMessageFlags(mail!, spec.IsRead, CorpusPlacement.WritesSentFlag(placement));

                            DateTime? readBack = ApplyDates(
                                mail!, dateMethod, spec.ReceivedUtc + writeShift, spec.SentUtc + writeShift);

                            if (CorpusPlacement.RequiresMove(placement))
                            {
                                // A move issues a NEW EntryID. The manifest is the teardown
                                // allowlist, so recording the pre-move id would name nothing
                                // and leave every item in this corpus undeletable.
                                dynamic moved = mail!.Move(folders[spec.FolderId]);
                                Release(mail);
                                mail = moved;
                            }

                            string entryId = (string)mail!.EntryID;
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
                    Release(draftsItems);
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

    /// <summary>What a re-anchor did.</summary>
    /// <param name="Rewritten">Items whose dates were written.</param>
    /// <param name="AlreadyCorrect">Items the plan said were already carrying the target instant.</param>
    /// <param name="Refused">Items the two-key-plus-ordinal rule declined to touch.</param>
    /// <param name="Gone">Manifest entries whose item no longer exists.</param>
    /// <param name="Failed">Writes that threw.</param>
    /// <param name="Elapsed">Monotonic elapsed time.</param>
    /// <param name="FirstError">The first failure's message, so a run that failed everywhere says why once.</param>
    public sealed record ReanchorOutcome(
        int Rewritten, int AlreadyCorrect, int Refused, int Gone, int Failed, TimeSpan Elapsed, string? FirstError);

    /// <summary>
    /// Writes the target instants onto the items a <see cref="CorpusReanchorPlan"/> names, in
    /// one STA session, recording each one as it goes.
    /// <para>
    /// It NEVER creates, moves or removes anything - it opens an item by EntryID and writes
    /// two date properties - and it touches an item only when the EntryID came from this
    /// corpus's manifest, the subject re-read from the item still carries both tags, and the
    /// ordinal in that subject is the ordinal being addressed
    /// (<see cref="CorpusSafety.MayRewrite"/>).
    /// </para>
    /// <para>
    /// <paramref name="record"/> is called with a replacement manifest line per item, before
    /// the next item is opened. Manifest item lines are last-writer-wins by ordinal, so
    /// appending is all that is needed: an interrupted re-anchor leaves a manifest that
    /// describes exactly what the store now holds, and running it again finishes the job,
    /// because the remaining work is derived from that manifest rather than from a cursor.
    /// </para>
    /// </summary>
    /// <param name="reanchor">The work sheet, already decided.</param>
    /// <param name="storeDisplayName">Target store; already vetted by <see cref="CorpusSafety.EvaluateStore"/>.</param>
    /// <param name="corpusId">The corpus id every touched subject must carry.</param>
    /// <param name="dateMethod">The date-write rung the probe verified on this store.</param>
    /// <param name="writeShift">Pre-compensation added to every instant written; zero unless the probe found an offset.</param>
    /// <param name="allowlist">EntryIDs this corpus's manifest records. Nothing outside it is opened.</param>
    /// <param name="record">Called with every rewritten item; must persist it before returning.</param>
    /// <param name="progress">Called every <paramref name="progressEvery"/> items.</param>
    /// <param name="progressEvery">Progress interval in items.</param>
    public static ReanchorOutcome Reanchor(
        CorpusReanchorPlan reanchor,
        string storeDisplayName,
        string corpusId,
        CorpusDateWriteMethod dateMethod,
        TimeSpan writeShift,
        ISet<string> allowlist,
        Action<CorpusManifestItem> record,
        Action<BuildProgress> progress,
        int progressEvery)
    {
        ArgumentNullException.ThrowIfNull(reanchor);
        ArgumentNullException.ThrowIfNull(allowlist);
        ArgumentNullException.ThrowIfNull(record);
        ArgumentNullException.ThrowIfNull(progress);
        ArgumentException.ThrowIfNullOrWhiteSpace(corpusId);

        Stopwatch elapsed = Stopwatch.StartNew();
        return RunSta<ReanchorOutcome>(
            () =>
            {
                dynamic app = CreateOutlookApplication();
                dynamic? ns = null;
                dynamic? stores = null;
                dynamic? store = null;
                int rewritten = 0;
                int refused = 0;
                int gone = 0;
                int failed = 0;
                string? firstError = null;
                try
                {
                    ns = app.GetNamespace("MAPI");
                    stores = ns.Stores;
                    store = FindStore(stores, storeDisplayName)
                        ?? throw new InvalidOperationException("Store not found for the corpus re-anchor.");
                    string storeId = (string)store.StoreID;

                    foreach (CorpusReanchorItem item in reanchor.Todo)
                    {
                        dynamic? mail = null;
                        try
                        {
                            try
                            {
                                mail = ns!.GetItemFromID(item.EntryId, storeId);
                            }
                            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
                            {
                                gone++;
                                continue;
                            }

                            string? subject = TryRead<string>(() => (string)mail!.Subject);
                            if (!CorpusSafety.MayRewrite(item.EntryId, subject, allowlist, corpusId, item.Ordinal))
                            {
                                refused++;
                                continue;
                            }

                            DateTime? readBack = ApplyDates(
                                mail!, dateMethod, item.ReceivedUtc + writeShift, item.SentUtc + writeShift);

                            // FolderId and BodyBytes are recorded as 0: a re-anchor knows
                            // neither and must not claim to. The manifest's own reader takes
                            // the LAST line for an ordinal, and the only field a re-anchor is
                            // entitled to restate is the instant it just wrote.
                            var line = new CorpusManifestItem(
                                item.Ordinal,
                                item.EntryId,
                                0,
                                0,
                                CorpusManifest.FormatUtc(readBack ?? item.ReceivedUtc));
                            record(line);
                            rewritten++;
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

                        if ((rewritten + failed + refused + gone) % progressEvery == 0)
                        {
                            progress(new BuildProgress(
                                rewritten,
                                reanchor.AlreadyCorrect,
                                failed,
                                reanchor.Todo.Count - rewritten - failed - refused - gone,
                                0,
                                elapsed.Elapsed));
                        }
                    }

                    return new ReanchorOutcome(
                        rewritten, reanchor.AlreadyCorrect, refused, gone, failed, elapsed.Elapsed, firstError);
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
        foreach (int folderId in ScanFolderIds)
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
                // The folder id the substitute STANDS IN FOR, not 0. A census compares where
                // an item is against where the plan puts it, and a Junk item found in the
                // folder created because the PST has no Junk Email is where it belongs.
                CollectCorpusItems(folder!, created.FolderId, corpusId, rows);
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
        dynamic mail, CorpusDateWriteMethod method, DateTime receivedUtc, DateTime sentUtc)
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
                // Dates only. The MSGFLAG_UNSENT write that used to live here has moved to
                // ApplySentFlag, because it decides where an item LIVES rather than what it
                // is dated - and while the two shared a rung, overriding the date refusal
                // silently disabled the placement fix as well.
                dynamic? accessor = null;
                try
                {
                    accessor = mail.PropertyAccessor;
                    accessor.SetProperty(PrMessageDeliveryTime, receivedUtc);
                    accessor.SetProperty(PrClientSubmitTime, sentUtc);
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
        dynamic drafts,
        string storeId,
        string corpusId,
        CorpusDateWriteMethod method,
        CorpusPlacementMethod placement,
        DateTime requestedUtc,
        DateTime writeUtc)
    {
        dynamic? items = null;
        dynamic? mail = null;
        string? entryId = null;
        try
        {
            // Built with the SAME placement the build will use, and that is not a detail.
            // The first version created the probe straight into the Inbox and then asked the
            // Inbox's table about it; because an unsent item is filed into Drafts, the table
            // never held it, and the probe reported daslIn=False - which reads as "the date
            // does not drive selection" when the real cause was "the item is in another
            // folder". The two failures were indistinguishable in the output, and the date
            // verdict taken from that run cannot be trusted. Placement is settled first now,
            // and the date probe inherits it.
            dynamic source = CorpusPlacement.CreatesInDrafts(placement) ? drafts : folder;
            items = source.Items;
            mail = items.Add(0);
            mail.Subject = ProbeSubject(corpusId, "date " + method);
            mail.Body = "date fidelity probe";
            mail.Save();

            ApplyMessageFlags(mail!, isRead: true, clearUnsent: CorpusPlacement.WritesSentFlag(placement));

            DateTime? readBack;
            try
            {
                readBack = ApplyDates(mail!, method, writeUtc, writeUtc);
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                entryId = TryRead<string>(() => (string)mail!.EntryID);
                return new CorpusDateProbe(method, requestedUtc, writeUtc, null, false, false, ex.Message);
            }

            if (CorpusPlacement.RequiresMove(placement))
            {
                dynamic moved = mail!.Move(folder);
                Release(mail);
                mail = moved;
            }

            entryId = (string)mail!.EntryID;
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

            TableLookup inside = TableFind(
                folder, DateWindowFilter(corpusId, requestedUtc.AddDays(-1), requestedUtc.AddDays(1)), entryId!);
            TableLookup outside = TableFind(
                folder, DateWindowFilter(corpusId, requestedUtc.AddDays(2), null), entryId!);
            if (inside == TableLookup.Inconclusive || outside == TableLookup.Inconclusive)
            {
                // Reported as an ERROR rather than as a negative result. A date rung that
                // could not be checked is not a rung that failed, and treating it as one
                // sends the ladder down to a worse method for no reason.
                return new CorpusDateProbe(
                    method, requestedUtc, writeUtc, readBack, false, false,
                    "the folder table could not answer whether the item is in the window "
                    + "(row cap reached, or the table could not be read)");
            }

            return new CorpusDateProbe(
                method, requestedUtc, writeUtc, readBack,
                inside == TableLookup.Found, outside == TableLookup.NotFound, null);
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
    private static string DateWindowFilter(string corpusId, DateTime fromUtc, DateTime? toUtc)
    {
        // Narrowed to the ONE probe ordinal, not to the whole corpus. The wide form had the
        // same defect as the placement check: against a populated store the open-ended
        // "everything newer than X" restriction returns thousands of rows, the walk stops at
        // its cap, and "not found" is then indistinguishable from "not reached" - which is
        // how the exclusion half of this probe could report a pass it had not proved.
        string received = "\"urn:schemas:httpmail:datereceived\"";
        string filter = "@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%"
            + CorpusPlan.DaslSubjectFragment(corpusId, CorpusPlan.ProbeOrdinal) + "%')"
            + " AND (" + received + " >= '" + DaslDateLiteral.FormatUtc(fromUtc) + "')";
        if (toUtc != null)
        {
            filter += " AND (" + received + " < '" + DaslDateLiteral.FormatUtc(toUtc.Value) + "')";
        }

        return filter;
    }

    /// <summary>What a table lookup for one item concluded.</summary>
    private enum TableLookup
    {
        /// <summary>The table returned the item.</summary>
        Found,

        /// <summary>The table was walked to its end and the item was not in it.</summary>
        NotFound,

        /// <summary>
        /// The question was not answered: the row cap was reached with rows still to come, or
        /// the table could not be read. Kept distinct from <see cref="NotFound"/> because
        /// collapsing the two is what made a large folder look like a placement failure - a
        /// probe gave up after 2 000 rows of a 22 000-row table, reported the item ABSENT,
        /// and the build refused a placement that worked.
        /// </summary>
        Inconclusive,
    }

    /// <summary>
    /// Asks a folder's table whether it carries one specific item. The FILTER is expected to
    /// be selective enough to return roughly that one item - <see cref="ProbeTableRowCap"/>
    /// is a bound on a runaway walk, not a search budget - and reaching it is reported as
    /// <see cref="TableLookup.Inconclusive"/> rather than as an answer.
    /// </summary>
    private static TableLookup TableFind(dynamic folder, string filter, string entryId)
    {
        dynamic? table = null;
        try
        {
            table = folder.GetTable(filter);
            int walked = 0;
            while (walked < ProbeTableRowCap)
            {
                if ((bool)table.EndOfTable)
                {
                    return TableLookup.NotFound;
                }

                dynamic? row = null;
                try
                {
                    row = table.GetNextRow();
                    walked++;
                    object[] values = (object[])row.GetValues();
                    if (values.Length > 0 && values[0] is string id
                        && string.Equals(id, entryId, StringComparison.OrdinalIgnoreCase))
                    {
                        return TableLookup.Found;
                    }
                }
                finally
                {
                    Release(row);
                }
            }

            return TableLookup.Inconclusive;
        }
        catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
        {
            return TableLookup.Inconclusive;
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
