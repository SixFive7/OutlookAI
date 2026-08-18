using System.Globalization;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// One item as the census sees it. Deliberately content-free (S3): an opaque
/// <see cref="Id"/>, a move-stable <see cref="Fingerprint"/> built from metadata, and a
/// flag saying whether the SUBJECT carried the live-tier tag. The subject itself is read
/// in-process to compute that flag and is never stored, never logged and never printed -
/// only the boolean leaves the census.
/// </summary>
public readonly struct CensusItem
{
    /// <summary>Builds one census entry.</summary>
    /// <param name="id">The item's EntryID at census time. Opaque, and unique within a folder.</param>
    /// <param name="fingerprint">
    /// A key that survives a move between folders, or null when the item had no usable one.
    /// EntryIDs do NOT survive a move - Outlook reissues them - so a moved item can only be
    /// recognised at its destination by something derived from the message itself.
    /// </param>
    /// <param name="tagged">True when the subject carried <see cref="LiveOutlookTestMailer.SubjectTag"/>.</param>
    public CensusItem(string id, string? fingerprint, bool tagged)
    {
        Id = id;
        Fingerprint = fingerprint;
        Tagged = tagged;
    }

    /// <summary>The item's EntryID when the census ran. Opaque; safe to print.</summary>
    public string Id { get; }

    /// <summary>Move-stable key, or null when none could be read.</summary>
    public string? Fingerprint { get; }

    /// <summary>True when this item was created by the live tier.</summary>
    public bool Tagged { get; }
}

/// <summary>
/// One folder's census: how many items it held, and - when the identity budget stretched
/// that far - WHICH items. A count alone is blind in both directions: it cannot tell a
/// deletion from a move, and it cannot see one item removed while another arrives. The
/// item list is what removes both blind spots, and it is optional because capturing it
/// over a real profile's Archive and Deleted Items would cost more than the run.
/// </summary>
public sealed class FolderCensus
{
    private FolderCensus(int count, IReadOnlyList<CensusItem>? items)
    {
        Count = count;
        Items = items;
    }

    /// <summary>Items the folder held. Equals <c>Items.Count</c> whenever identities were captured.</summary>
    public int Count { get; }

    /// <summary>The items, or null when only the count was affordable.</summary>
    public IReadOnlyList<CensusItem>? Items { get; }

    /// <summary>True when this folder can be compared item by item rather than by count.</summary>
    public bool HasIdentities => Items != null;

    /// <summary>A folder measured by <c>Folder.Items.Count</c> alone.</summary>
    public static FolderCensus CountOnly(int count)
    {
        return new FolderCensus(count, null);
    }

    /// <summary>A folder walked item by item. The count IS the walk, so the two cannot disagree.</summary>
    public static FolderCensus WithItems(IReadOnlyList<CensusItem> items)
    {
        ArgumentNullException.ThrowIfNull(items);
        return new FolderCensus(items.Count, items);
    }
}

/// <summary>Verdict of one before/after comparison. Failures fail the suite; notes are informational.</summary>
public sealed class TripwireVerdict
{
    internal TripwireVerdict(IReadOnlyList<string> failures, IReadOnlyList<string> notes, string? attribution = null)
    {
        Failures = failures;
        Notes = notes;
        Attribution = attribution;
    }

    /// <summary>Losses: items or folders that left a mailbox the suite may not write to.</summary>
    public IReadOnlyList<string> Failures { get; }

    /// <summary>Benign observations (mail arriving elsewhere during the run, filing, hub churn).</summary>
    public IReadOnlyList<string> Notes { get; }

    /// <summary>
    /// What the evidence says about WHO did it, or null when nothing failed. Never guesses:
    /// it says "the suite" only when a departed item carried the live-tier tag, and otherwise
    /// says the question is undecidable from a before/after census - which it is.
    /// </summary>
    public string? Attribution { get; }

    /// <summary>True when anything must fail the suite.</summary>
    public bool Failed => Failures.Count > 0;

    /// <summary>One multi-line message naming every store/folder/item, ending in the attribution.</summary>
    public string Describe()
    {
        string body = "STORE COUNT TRIPWIRE: the live tier changed mailboxes it may not touch."
            + Environment.NewLine + string.Join(Environment.NewLine, Failures);
        return Attribution == null ? body : body + Environment.NewLine + Attribution;
    }
}

/// <summary>
/// Pure comparison behind the per-store tripwire: a live run may add nothing and remove
/// nothing outside the designated test mailbox.
/// <para>
/// <b>What it can and cannot decide.</b> No before/after reading of a mailbox can name the
/// actor - Outlook records that an item is gone, never who removed it. So the guard stays
/// fail-closed on every removal it cannot explain, and spends its effort on making a firing
/// CHECKABLE instead: which items left, where they went, and whether any of them was the
/// suite's own. A guard nobody can check is a guard that gets waved through, and this one
/// stands between the suite and the incident that once destroyed real mail.
/// </para>
/// <para>
/// Tuned for low false positives, because a 27-minute run happens while real mail arrives
/// and a real person works:
/// <list type="bullet">
/// <item>item ARRIVALS outside the hub are normal (mail arriving, a rule filing it)
/// - noted, never failed;</item>
/// <item>an item that left one folder and turned up in ANOTHER ORDINARY FOLDER of the same
/// store was FILED, not lost - noted, never failed. This is the one exoneration the census
/// can actually prove, and it needs the identity capture: a count cannot see it;</item>
/// <item>item DEPARTURES that are not accounted for that way FAIL - including a departure
/// masked by an arrival, which a count tripwire cannot see at all;</item>
/// <item>a folder ADDED or REMOVED outside the hub FAILS - the suite creates folders only
/// in the hub, and removes only its own;</item>
/// <item>the hub is exempt here: its churn is tagged, and the zero-tagged-artifact sweep
/// plus the move/archive reconciliation already police it.</item>
/// </list>
/// </para>
/// </summary>
public static class StoreCountTripwire
{
    /// <summary>
    /// Marks a census entry the SYSTEM prunes on its own (Deleted Items ages out, junk mail
    /// expires, and Outlook writes and removes sync-issue reports unprompted). Shrinking
    /// there is not evidence of anything, so it is noted rather than failed - a tripwire
    /// that cries wolf gets ignored, which is the one outcome that must not happen.
    /// <para>
    /// It is NOT a benign DESTINATION, though: an item that left an ordinary folder and
    /// turned up in one of these was deleted or junked, and that still fails.
    /// </para>
    /// </summary>
    public const string VolatilePrefix = "~";

    /// <summary>
    /// How many departed items one failure line names before it summarises the rest. Five
    /// is enough to recognise an afternoon's mail and short enough that a mass deletion does
    /// not bury its own headline under a thousand EntryIDs.
    /// </summary>
    public const int MaxReportedDepartures = 5;

    /// <summary>True when a census key names a self-pruning folder.</summary>
    public static bool IsVolatile(string folderKey)
    {
        return folderKey != null && folderKey.StartsWith(VolatilePrefix, StringComparison.Ordinal);
    }

    private static string Display(string folderKey)
    {
        return IsVolatile(folderKey) ? folderKey[VolatilePrefix.Length..] + " (self-pruning)" : folderKey;
    }

    /// <summary>Compares two per-store, per-folder censuses.</summary>
    public static TripwireVerdict Evaluate(
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, FolderCensus>> before,
        IReadOnlyDictionary<string, IReadOnlyDictionary<string, FolderCensus>> after,
        string hubStoreDisplayName,
        IEnumerable<string>? lazyHierarchyStores = null)
    {
        ArgumentNullException.ThrowIfNull(before);
        ArgumentNullException.ThrowIfNull(after);

        // MEASURED on this profile: a delegate/shared store's folder HIERARCHY is synced
        // lazily - the same mailbox enumerated 165 folders in one census and 159 minutes
        // later, with a real 450-item subfolder simply absent from the second walk (and
        // present again afterwards). A folder appearing or disappearing there is therefore
        // evidence about the hierarchy cache, not about deletion. Item loss inside a folder
        // seen in BOTH censuses stays a hard failure everywhere - that is the shape mass
        // deletion actually has.
        HashSet<string> lazyStores = new(lazyHierarchyStores ?? [], StringComparer.OrdinalIgnoreCase);

        List<string> failures = new();
        List<string> notes = new();
        int taggedDepartures = 0;
        bool sawDepartures = false;

        foreach (KeyValuePair<string, IReadOnlyDictionary<string, FolderCensus>> store in before)
        {
            bool isHub = string.Equals(store.Key, hubStoreDisplayName, StringComparison.OrdinalIgnoreCase);
            bool lazyHierarchy = lazyStores.Contains(store.Key);
            if (!after.TryGetValue(store.Key, out IReadOnlyDictionary<string, FolderCensus>? now))
            {
                // A store present at the start and gone at the end is never benign.
                failures.Add("  store '" + store.Key + "' could not be re-counted after the run (store missing).");
                continue;
            }

            // Everything that ARRIVED anywhere in this store, keyed so a departure can be
            // matched against it. Built from arrivals only - an item that was already
            // sitting in Deleted Items must never be mistaken for the one that just left
            // the Inbox.
            RelocationIndex relocations = isHub ? RelocationIndex.Empty : RelocationIndex.Build(store.Value, now);

            foreach (KeyValuePair<string, FolderCensus> folder in store.Value)
            {
                if (!now.TryGetValue(folder.Key, out FolderCensus? current))
                {
                    if (isHub || lazyHierarchy || IsVolatile(folder.Key))
                    {
                        notes.Add("  folder not enumerated after the run: store '" + store.Key + "' folder '"
                            + Display(folder.Key) + "'.");
                    }
                    else
                    {
                        failures.Add("  FOLDER REMOVED: store '" + store.Key + "' folder '" + Display(folder.Key)
                            + "' existed before the run and is gone (" + Count(folder.Value.Count) + " before).");
                    }

                    continue;
                }

                bool exempt = isHub || IsVolatile(folder.Key);
                if (folder.Value.HasIdentities && current.HasIdentities)
                {
                    EvaluateByIdentity(
                        store.Key, folder.Key, folder.Value, current, exempt, relocations,
                        failures, notes, ref taggedDepartures, ref sawDepartures);
                }
                else
                {
                    EvaluateByCount(
                        store.Key, folder.Key, folder.Value.Count, current.Count, exempt, isHub, failures, notes);
                }
            }

            foreach (KeyValuePair<string, FolderCensus> folder in now)
            {
                if (store.Value.ContainsKey(folder.Key))
                {
                    continue;
                }

                if (isHub || lazyHierarchy || IsVolatile(folder.Key))
                {
                    notes.Add("  folder newly enumerated: store '" + store.Key + "' folder '"
                        + Display(folder.Key) + "'.");
                }
                else
                {
                    failures.Add("  FOLDER ADDED: store '" + store.Key + "' folder '" + Display(folder.Key)
                        + "' did not exist before the run (" + Count(folder.Value.Count) + " items).");
                }
            }
        }

        foreach (string storeName in after.Keys)
        {
            if (!before.ContainsKey(storeName))
            {
                failures.Add("  store '" + storeName + "' appeared during the run and was never baselined.");
            }
        }

        return new TripwireVerdict(failures, notes, Attribute(failures, taggedDepartures, sawDepartures));
    }

    /// <summary>
    /// The count rule, for folders the identity budget did not reach: a DECREASE outside the
    /// hub fails, an increase is noted. It cannot see a departure masked by an arrival, which
    /// is exactly why the identity path exists and why this one says so when it fires.
    /// </summary>
    private static void EvaluateByCount(
        string store, string folderKey, int wasCount, int nowCount, bool exempt, bool isHub,
        List<string> failures, List<string> notes)
    {
        int delta = nowCount - wasCount;
        if (delta == 0)
        {
            return;
        }

        if (delta < 0 && !exempt)
        {
            failures.Add("  ITEMS LOST: store '" + store + "' folder '" + Display(folderKey) + "' "
                + Count(wasCount) + " -> " + Count(nowCount) + " (" + Count(delta)
                + "); counted only, so WHICH items left is not known (folder above the identity budget).");
        }
        else if (!isHub)
        {
            notes.Add("  churn: store '" + store + "' folder '" + Display(folderKey) + "' "
                + Count(wasCount) + " -> " + Count(nowCount)
                + " (" + (delta > 0 ? "+" : string.Empty) + Count(delta) + ").");
        }
    }

    /// <summary>
    /// The identity rule: every item that was here at the baseline must still be here, OR be
    /// somewhere else in the same store. Anything else is a removal and fails, whether or not
    /// the folder's COUNT moved - a run that deletes one item while one arrives leaves the
    /// count untouched and is invisible to the count rule.
    /// </summary>
    private static void EvaluateByIdentity(
        string store, string folderKey, FolderCensus was, FolderCensus now, bool exempt,
        RelocationIndex relocations, List<string> failures, List<string> notes,
        ref int taggedDepartures, ref bool sawDepartures)
    {
        HashSet<string> present = new(now.Items!.Select(i => i.Id), StringComparer.Ordinal);
        List<CensusItem> departed = was.Items!.Where(i => !present.Contains(i.Id)).ToList();
        int arrived = now.Count - (was.Count - departed.Count);

        if (arrived > 0 && !exempt)
        {
            notes.Add("  churn: store '" + store + "' folder '" + Display(folderKey) + "' gained "
                + Count(arrived) + " item(s) (" + Count(was.Count) + " -> " + Count(now.Count) + ").");
        }

        if (departed.Count == 0 || exempt)
        {
            return;
        }

        sawDepartures = true;
        List<string> filed = new();
        List<string> removed = new();
        foreach (CensusItem item in departed)
        {
            if (item.Tagged)
            {
                taggedDepartures++;
            }

            if (relocations.TryConsume(item, out string? destination) && !IsVolatile(destination!))
            {
                filed.Add(Display(destination!));
                continue;
            }

            removed.Add(DescribeDeparture(item, destination));
        }

        if (filed.Count > 0)
        {
            notes.Add("  filed (not loss): store '" + store + "' folder '" + Display(folderKey) + "' "
                + Count(filed.Count) + " item(s) moved to " + string.Join(", ", filed.Distinct(StringComparer.Ordinal))
                + " in the same store.");
        }

        if (removed.Count == 0)
        {
            return;
        }

        string listed = string.Join("; ", removed.Take(MaxReportedDepartures));
        string rest = removed.Count > MaxReportedDepartures
            ? " and " + Count(removed.Count - MaxReportedDepartures) + " more"
            : string.Empty;
        failures.Add("  ITEMS REMOVED: store '" + store + "' folder '" + Display(folderKey) + "' lost "
            + Count(removed.Count) + " item(s) (" + Count(was.Count) + " -> " + Count(now.Count)
            + ", " + Count(arrived) + " arriving): " + listed + rest + ".");
    }

    /// <summary>
    /// One departed item, in terms nobody's mailbox leaks through: its baseline EntryID, its
    /// move-stable fingerprint (a received instant and a size, never a subject or a body),
    /// and where it ended up if the census could find it.
    /// </summary>
    private static string DescribeDeparture(CensusItem item, string? volatileDestination)
    {
        string where = volatileDestination == null
            ? "not found in any folder this census identified"
            : "now in '" + Display(volatileDestination) + "'";
        string tag = item.Tagged ? ", TEST-TAGGED" : string.Empty;
        return "[" + item.Id + (item.Fingerprint == null ? string.Empty : " " + item.Fingerprint) + "] " + where + tag;
    }

    /// <summary>
    /// Who did it, said only as far as the evidence goes. A tagged item leaving a store the
    /// suite may not write to is the suite and nothing else. Everything else is undecidable
    /// from a before/after census, and saying so plainly is the point: the maintainer
    /// deleting his own mail during a 27-minute run produces the same reading as a runaway
    /// test doing it, so the guard fails either way and hands over what it saw.
    /// </summary>
    private static string? Attribute(IReadOnlyList<string> failures, int taggedDepartures, bool sawDepartures)
    {
        if (failures.Count == 0)
        {
            return null;
        }

        if (taggedDepartures > 0)
        {
            return "ATTRIBUTION: THE SUITE. " + Count(taggedDepartures) + " departed item(s) carried "
                + LiveOutlookTestMailer.SubjectTag + " in a mailbox the suite may not write to. "
                + "This is a real incident, not background activity - stop and investigate.";
        }

        return "ATTRIBUTION: undecidable. No departed item carried " + LiveOutlookTestMailer.SubjectTag
            + ", and the write allowlist confines the suite to the hub store"
            + (sawDepartures ? ", so this reads the same as a person using the mailbox during the run" : string.Empty)
            + " - a before/after census cannot name the actor, so it fails and shows its working. "
            + "Check the EntryIDs and destinations above before dismissing it.";
    }

    private static string Count(int value)
    {
        return value.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// Where an item that left one folder could have gone: everything that ARRIVED elsewhere
    /// in the same store, matched by the move-stable fingerprint first and by EntryID second.
    /// <para>
    /// Built from arrivals rather than from the whole after-census, so an item that was
    /// already sitting in Deleted Items cannot be mistaken for the one that just left the
    /// Inbox; and each arrival is CONSUMED by at most one departure, so two deletions cannot
    /// both be explained away by one unrelated move.
    /// </para>
    /// </summary>
    private sealed class RelocationIndex
    {
        private readonly Dictionary<string, List<string>> _byFingerprint = new(StringComparer.Ordinal);
        private readonly Dictionary<string, List<string>> _byId = new(StringComparer.Ordinal);

        /// <summary>An index that explains nothing - used for the hub, which this guard exempts.</summary>
        public static RelocationIndex Empty { get; } = new RelocationIndex();

        /// <summary>Indexes every item present after the run that was not in the same folder before it.</summary>
        public static RelocationIndex Build(
            IReadOnlyDictionary<string, FolderCensus> before,
            IReadOnlyDictionary<string, FolderCensus> after)
        {
            RelocationIndex index = new();
            foreach (KeyValuePair<string, FolderCensus> folder in after)
            {
                if (folder.Value.Items == null)
                {
                    continue;
                }

                HashSet<string> baseline = before.TryGetValue(folder.Key, out FolderCensus? was) && was.Items != null
                    ? new HashSet<string>(was.Items.Select(i => i.Id), StringComparer.Ordinal)
                    : new HashSet<string>(StringComparer.Ordinal);

                foreach (CensusItem item in folder.Value.Items)
                {
                    if (baseline.Contains(item.Id))
                    {
                        continue;
                    }

                    Add(index._byId, item.Id, folder.Key);
                    if (item.Fingerprint != null)
                    {
                        Add(index._byFingerprint, item.Fingerprint, folder.Key);
                    }
                }
            }

            return index;
        }

        /// <summary>
        /// Claims one arrival for <paramref name="departed"/> and returns the folder it landed
        /// in, or false when nothing in this store can account for it.
        /// </summary>
        public bool TryConsume(CensusItem departed, out string? folderKey)
        {
            if (departed.Fingerprint != null && Take(_byFingerprint, departed.Fingerprint, out folderKey))
            {
                return true;
            }

            // Second chance: some stores DO keep an EntryID across a move. Cheap to check,
            // and its unreliability is what makes the fingerprint necessary rather than optional.
            return Take(_byId, departed.Id, out folderKey);
        }

        private static void Add(Dictionary<string, List<string>> index, string key, string folderKey)
        {
            if (!index.TryGetValue(key, out List<string>? folders))
            {
                folders = new List<string>();
                index[key] = folders;
            }

            folders.Add(folderKey);
        }

        private static bool Take(Dictionary<string, List<string>> index, string key, out string? folderKey)
        {
            folderKey = null;
            if (!index.TryGetValue(key, out List<string>? folders) || folders.Count == 0)
            {
                return false;
            }

            folderKey = folders[0];
            folders.RemoveAt(0);
            return true;
        }
    }
}
