using System.Globalization;

namespace OutlookAI.RemediationTools;

/// <summary>One recipient row as Outlook reports it back.</summary>
/// <param name="Address">Recipient.Address - the SMTP address for a resolved one-off.</param>
/// <param name="Kind">Recipient.Type: 1 To, 2 Cc, 3 Bcc.</param>
public sealed record CorpusObservedRecipient(string? Address, int Kind);

/// <summary>
/// What one built item actually carries, read back from the store by EntryID. Every field is
/// nullable for the same reason <see cref="CorpusStoreFacts"/> is: a late-bound read that fails
/// must read as "not established", never as "empty".
/// </summary>
/// <param name="Ordinal">The ordinal the manifest records for this EntryID.</param>
/// <param name="SenderAddress">The sender's SMTP address.</param>
/// <param name="SenderName">The sender's display name.</param>
/// <param name="Recipients">Every recipient row, or null when the collection could not be read.</param>
/// <param name="AttachmentNames">Every attachment's file name, or null when the collection could not be read.</param>
/// <param name="ConversationIndexHex">PR_CONVERSATION_INDEX as hex, or null.</param>
/// <param name="ConversationId">MailItem.ConversationID - what the store COMPUTED the conversation to be.</param>
/// <param name="Error">Why the item could not be read at all, or null.</param>
/// <param name="MessageClass">The item's MessageClass, for an UNDATED item's read-back; null when not read.</param>
/// <param name="HasDeliveryTime">
/// Whether the item carries PR_MESSAGE_DELIVERY_TIME at all - for an UNDATED item's read-back, where
/// the answer must be false; null when not read.
/// </param>
public sealed record CorpusEnrichmentObservation(
    int Ordinal,
    string? SenderAddress,
    string? SenderName,
    IReadOnlyList<CorpusObservedRecipient>? Recipients,
    IReadOnlyList<string>? AttachmentNames,
    string? ConversationIndexHex,
    string? ConversationId,
    string? Error = null,
    string? MessageClass = null,
    bool? HasDeliveryTime = null);

/// <summary>What an undated item's message class is judged on. Pure.</summary>
public static class CorpusMessageFlags
{
    /// <summary>
    /// Whether <paramref name="actual"/> is the message class <paramref name="expected"/> or a subclass
    /// of it (<c>IPM.Note.Something</c> for <c>IPM.Note</c>). Case-insensitive, as MAPI treats classes.
    /// </summary>
    public static bool ClassMatches(string? actual, string expected)
        => actual != null
            && (string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase)
                || actual.StartsWith(expected + ".", StringComparison.OrdinalIgnoreCase));
}

/// <summary>What a read-back of a population found, counted.</summary>
/// <param name="Planned">Items the plan describes.</param>
/// <param name="Observed">Items that could be read back.</param>
/// <param name="Unreadable">Items the read-back could not open at all.</param>
/// <param name="SenderMismatches">Items whose sender address is not the planned one.</param>
/// <param name="RecipientMismatches">Items whose recipient rows are not the planned ones.</param>
/// <param name="AttachmentMismatches">Items whose attachment file names are not the planned ones.</param>
/// <param name="ConversationIndexMismatches">Conversation members whose index is not the one written.</param>
/// <param name="Threads">Conversations the plan describes.</param>
/// <param name="ThreadsGroupedByStore">Conversations whose members the store gave ONE shared, non-empty conversation id.</param>
/// <param name="ThreadsUnestablished">Conversations whose conversation id could not be read for every member.</param>
/// <param name="ThreadsSharingAnId">Pairs of different conversations the store gave the same id.</param>
public sealed record CorpusEnrichmentReport(
    int Planned,
    int Observed,
    int Unreadable,
    int SenderMismatches,
    int RecipientMismatches,
    int AttachmentMismatches,
    int ConversationIndexMismatches,
    int Threads,
    int ThreadsGroupedByStore,
    int ThreadsUnestablished,
    int ThreadsSharingAnId)
{
    /// <summary>UNDATED items the plan describes - appointments, contacts, tasks.</summary>
    public int UndatedPlanned { get; init; }

    /// <summary>Undated items that read back as a different kind of item than the plan made.</summary>
    public int UndatedClassMismatches { get; init; }

    /// <summary>
    /// Undated items that read back CARRYING a delivery time - the one thing they exist not to carry.
    /// Non-zero means the index will date them, and the order-key tests measure nothing again.
    /// </summary>
    public int UndatedCarryingADate { get; init; }

    /// <summary>Undated items whose delivery time could not be read either way.</summary>
    public int UndatedDateUnestablished { get; init; }
}

/// <summary>
/// Says whether a built population carries what the plan says it carries. The census proves the
/// ITEMS are there, in the right folders, once each; this proves they are the right items -
/// addressed, attached and threaded as the live tests that read them assume.
/// <para>
/// <b>Why a population needs this and the measurement corpus never did.</b> A corpus item is a
/// subject, a body and two dates, and the census and the date probe already cover all four. A
/// population item also carries a sender and recipients written through the PropertyAccessor and
/// the recipient table, attachments, and a conversation index - and a write that Outlook quietly
/// declines leaves an item that LOOKS built. The tests that read those shapes would then fail far
/// away, or worse pass having read nothing (a sender filter over no senders). So the build reads
/// every item back and fails loudly instead.
/// </para>
/// <para>
/// The conversation half is judged on what the STORE computed, not on what was written: Outlook
/// groups a conversation by an id it derives from the index, and <c>thread</c> and the index tier
/// both read that id. A thread whose members come back with different ids, or none, was written
/// correctly and is still not a conversation - and that is the fact the build reports.
/// </para>
/// </summary>
public static class CorpusEnrichmentCheck
{
    /// <summary>Compares read-back observations against the plan. Pure.</summary>
    public static CorpusEnrichmentReport Compare(
        CorpusPlan plan, int itemCount, IEnumerable<CorpusEnrichmentObservation> observations)
    {
        ArgumentNullException.ThrowIfNull(plan);
        ArgumentNullException.ThrowIfNull(observations);
        if (plan.Population == null)
        {
            throw new ArgumentException("Only a population plan carries enrichment to check.", nameof(plan));
        }

        var byOrdinal = new Dictionary<int, CorpusEnrichmentObservation>();
        foreach (CorpusEnrichmentObservation o in observations)
        {
            byOrdinal[o.Ordinal] = o;
        }

        int observed = 0;
        int unreadable = 0;
        int senders = 0;
        int recipients = 0;
        int attachments = 0;
        int indexes = 0;
        int undatedPlanned = 0;
        int undatedClass = 0;
        int undatedDated = 0;
        int undatedUnknownDate = 0;
        var idsByThread = new SortedDictionary<int, List<string?>>();
        for (int ordinal = 1; ordinal <= itemCount; ordinal++)
        {
            CorpusItemEnrichment? want = plan.Enrich(ordinal);
            if (want == null)
            {
                // An UNDATED item: no sender, recipients, attachments or conversation to compare.
                // What it must be instead is the right KIND of item, with NO delivery time.
                CorpusItemSpec spec = plan.Describe(ordinal);
                undatedPlanned++;
                if (!byOrdinal.TryGetValue(ordinal, out CorpusEnrichmentObservation? seen))
                {
                    continue;
                }

                if (seen.Error != null)
                {
                    unreadable++;
                    continue;
                }

                observed++;
                if (!CorpusMessageFlags.ClassMatches(seen.MessageClass, CorpusItemKinds.MessageClassOf(spec.Kind)))
                {
                    undatedClass++;
                }

                if (seen.HasDeliveryTime == true)
                {
                    undatedDated++;
                }
                else if (seen.HasDeliveryTime == null)
                {
                    undatedUnknownDate++;
                }

                continue;
            }

            if (want.ThreadKey != null && !idsByThread.ContainsKey(want.ThreadKey.Value))
            {
                idsByThread[want.ThreadKey.Value] = new List<string?>();
            }

            if (!byOrdinal.TryGetValue(ordinal, out CorpusEnrichmentObservation? got))
            {
                continue;
            }

            if (got.Error != null)
            {
                unreadable++;
                continue;
            }

            observed++;
            if (!string.Equals(got.SenderAddress, want.Sender.Address, StringComparison.OrdinalIgnoreCase))
            {
                senders++;
            }

            if (!RecipientsMatch(want.Recipients, got.Recipients))
            {
                recipients++;
            }

            if (!AttachmentsMatch(want.Attachments, got.AttachmentNames))
            {
                attachments++;
            }

            if (want.ThreadKey != null)
            {
                if (!string.Equals(got.ConversationIndexHex, want.ConversationIndexHex, StringComparison.OrdinalIgnoreCase))
                {
                    indexes++;
                }

                idsByThread[want.ThreadKey.Value].Add(string.IsNullOrWhiteSpace(got.ConversationId) ? null : got.ConversationId);
            }
        }

        int grouped = 0;
        int unestablished = 0;
        var threadIds = new List<string>();
        foreach (List<string?> ids in idsByThread.Values)
        {
            if (ids.Count == 0 || ids.Any(i => i == null))
            {
                unestablished++;
                continue;
            }

            if (ids.Distinct(StringComparer.OrdinalIgnoreCase).Count() == 1)
            {
                grouped++;
                threadIds.Add(ids[0]!);
            }
        }

        int sharing = threadIds.Count - threadIds.Distinct(StringComparer.OrdinalIgnoreCase).Count();
        return new CorpusEnrichmentReport(
            itemCount, observed, unreadable, senders, recipients, attachments, indexes,
            idsByThread.Count, grouped, unestablished, sharing)
        {
            UndatedPlanned = undatedPlanned,
            UndatedClassMismatches = undatedClass,
            UndatedCarryingADate = undatedDated,
            UndatedDateUnestablished = undatedUnknownDate,
        };
    }

    /// <summary>Whether the population is what the plan says, and what to print either way.</summary>
    public static (bool Clean, string Message) Decide(CorpusEnrichmentReport report)
    {
        ArgumentNullException.ThrowIfNull(report);
        CultureInfo invariant = CultureInfo.InvariantCulture;
        var faults = new List<string>();
        int notRead = report.Planned - report.Observed - report.Unreadable;
        if (notRead > 0)
        {
            faults.Add($"{notRead.ToString(invariant)} planned item(s) had no manifest line to read back from");
        }

        if (report.Unreadable > 0)
        {
            faults.Add($"{report.Unreadable.ToString(invariant)} item(s) could not be opened by EntryID");
        }

        if (report.SenderMismatches > 0)
        {
            faults.Add($"{report.SenderMismatches.ToString(invariant)} item(s) do not carry the planned SENDER - the sender filter "
                + "and the sender half of the store discovery read exactly this");
        }

        if (report.RecipientMismatches > 0)
        {
            faults.Add($"{report.RecipientMismatches.ToString(invariant)} item(s) do not carry the planned RECIPIENTS - the "
                + "store discovery finds a small store by mail addressed to it");
        }

        if (report.AttachmentMismatches > 0)
        {
            faults.Add($"{report.AttachmentMismatches.ToString(invariant)} item(s) do not carry the planned ATTACHMENTS");
        }

        if (report.ConversationIndexMismatches > 0)
        {
            faults.Add($"{report.ConversationIndexMismatches.ToString(invariant)} conversation member(s) do not carry the "
                + "conversation index that was written");
        }

        if (report.ThreadsSharingAnId > 0)
        {
            faults.Add($"{report.ThreadsSharingAnId.ToString(invariant)} different conversation(s) came back sharing one id");
        }

        int split = report.Threads - report.ThreadsGroupedByStore - report.ThreadsUnestablished;
        if (split > 0)
        {
            faults.Add($"{split.ToString(invariant)} of {report.Threads.ToString(invariant)} conversation(s) came back with "
                + "DIFFERENT conversation ids across their members, so the store does not see them as one conversation "
                + "and the thread walk has nothing to walk");
        }

        if (report.UndatedClassMismatches > 0)
        {
            faults.Add($"{report.UndatedClassMismatches.ToString(invariant)} UNDATED item(s) read back as a different kind of "
                + "item than the plan made (message class)");
        }

        if (report.UndatedCarryingADate > 0)
        {
            faults.Add($"{report.UndatedCarryingADate.ToString(invariant)} UNDATED item(s) came back CARRYING a delivery time, "
                + "so the index will date them and LiveOrderKeyCollationTests measure no undated row at all");
        }

        if (report.UndatedDateUnestablished > 0)
        {
            faults.Add($"{report.UndatedDateUnestablished.ToString(invariant)} UNDATED item(s) could not be asked whether they "
                + "carry a delivery time, so their being undated is not established");
        }

        string head = $"Population read-back: {report.Observed.ToString(invariant)} of {report.Planned.ToString(invariant)} "
            + $"item(s) read; {report.ThreadsGroupedByStore.ToString(invariant)} of {report.Threads.ToString(invariant)} "
            + "conversation(s) grouped by the store under one id.";
        string unestablished = report.ThreadsUnestablished > 0
            ? $" NOT ESTABLISHED: {report.ThreadsUnestablished.ToString(invariant)} conversation(s) have a member whose "
                + "conversation id could not be read, so whether the store groups them is unknown - the thread tests "
                + "will say."
            : string.Empty;

        string undated = report.UndatedPlanned > 0
            ? $" {report.UndatedPlanned.ToString(invariant)} of them UNDATED - the right kind of item, with no delivery time."
            : string.Empty;

        return faults.Count == 0
            ? (true, head + " Every item carries the sender, recipients, attachments and conversation the plan names."
                + undated + unestablished)
            : (false, head + " FAULTS: " + string.Join("; ", faults) + "." + unestablished);
    }

    private static bool RecipientsMatch(IReadOnlyList<CorpusRecipient> want, IReadOnlyList<CorpusObservedRecipient>? got)
    {
        if (got == null || got.Count != want.Count)
        {
            return false;
        }

        var wanted = want
            .Select(r => (r.Person.Address.ToLowerInvariant(), (int)r.Kind))
            .OrderBy(r => r.Item1, StringComparer.Ordinal).ThenBy(r => r.Item2)
            .ToList();
        var have = got
            .Select(r => ((r.Address ?? string.Empty).ToLowerInvariant(), r.Kind))
            .OrderBy(r => r.Item1, StringComparer.Ordinal).ThenBy(r => r.Item2)
            .ToList();
        return wanted.SequenceEqual(have);
    }

    private static bool AttachmentsMatch(IReadOnlyList<CorpusAttachment> want, IReadOnlyList<string>? got)
    {
        if (got == null || got.Count != want.Count)
        {
            return false;
        }

        return want.Select(a => a.FileName).OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
            .SequenceEqual(got.OrderBy(n => n, StringComparer.OrdinalIgnoreCase), StringComparer.OrdinalIgnoreCase);
    }
}

/// <summary>
/// What the enrichment probe achieved on one throwaway item: every write a population build makes,
/// made once, placed the way the build places, and read back by EntryID.
/// </summary>
/// <param name="SenderWritten">The sender address read back is the one written.</param>
/// <param name="RecipientsWritten">
/// Both recipient rows came back RESOLVED, with the planned addresses and types - the To row the
/// store's OWNER, added exactly as every received item adds it. The probe addressed only
/// correspondents until 2026-09-24, and so passed on OAI-UNINDEXED while every received item of three
/// populations came out with an unresolved owner row (<see cref="CorpusCorrespondent.ToRecipientSpec"/>).
/// </param>
/// <param name="AttachmentWritten">The attachment came back under its file name, with content.</param>
/// <param name="ConversationIndexWritten">PR_CONVERSATION_INDEX came back as written.</param>
/// <param name="ConversationId">What the store computed as the conversation id, or null.</param>
/// <param name="Error">The COM failure that stopped the probe, or null.</param>
public sealed record CorpusEnrichmentProbe(
    bool SenderWritten,
    bool RecipientsWritten,
    bool AttachmentWritten,
    bool ConversationIndexWritten,
    string? ConversationId,
    string? Error);

/// <summary>
/// The go/no-go on an enrichment probe, as a pure function - the same discipline as the placement
/// and date probes: a population is only built into a store where one throwaway item proved every
/// write lands.
/// </summary>
public static class CorpusEnrichmentFidelity
{
    /// <summary>Whether a build may proceed, and the sentence that says why either way.</summary>
    public static (bool Proceed, string Message) Decide(CorpusEnrichmentProbe probe)
    {
        ArgumentNullException.ThrowIfNull(probe);
        if (probe.Error != null)
        {
            return (false, "REFUSING to build a population: the enrichment probe failed before it could read anything back ("
                + probe.Error + "). Nothing about sender, recipient, attachment or conversation writes is established on "
                + "this store.");
        }

        var failed = new List<string>();
        if (!probe.SenderWritten)
        {
            failed.Add("the SENDER did not read back as written");
        }

        if (!probe.RecipientsWritten)
        {
            failed.Add("the RECIPIENTS did not read back as written - the owner and the Cc, each resolved to its address");
        }

        if (!probe.AttachmentWritten)
        {
            failed.Add("the ATTACHMENT did not read back");
        }

        if (!probe.ConversationIndexWritten)
        {
            failed.Add("the CONVERSATION INDEX did not read back as written");
        }

        if (failed.Count > 0)
        {
            return (false, "REFUSING to build a population: " + string.Join("; ", failed) + ". A population missing any of "
                + "these is one the live tests that read it would misread - a sender filter over no senders passes having "
                + "checked nothing. There is no flag for this; the write path has to be fixed.");
        }

        string grouping = string.IsNullOrWhiteSpace(probe.ConversationId)
            ? " The store reported NO conversation id for the probe, so whether it will group the conversations is not "
                + "established here; the post-build read-back says for every one of them."
            : " The store computed a conversation id from the index.";
        return (true, "Enrichment probe verified: sender, recipients, attachment and conversation index all read back as "
            + "written." + grouping);
    }
}

/// <summary>
/// What the undated probe achieved for ONE kind: one throwaway item created in that kind's folder -
/// the store's own default folder, or the stand-in the build files the kind in when the store has
/// none - the way the build creates it, saved, re-opened by EntryID and read back.
/// </summary>
/// <param name="Kind">The undated kind probed.</param>
/// <param name="FolderReachable">The kind's folder was found or made, and <c>Items.Add</c> on it answered.</param>
/// <param name="InTheFolder">The saved item's parent IS that folder, compared by folder EntryID.</param>
/// <param name="SubjectTagParses">
/// The subject read back still carries the probe's corpus tags and parses to the probe ordinal - so the
/// two-key rule can delete it. For a contact this is the question that matters: which of its name fields
/// Outlook derives the subject from is not documented, and a contact whose subject lost its tag is one
/// nothing here could ever remove.
/// </param>
/// <param name="HasNoDeliveryTime">PR_MESSAGE_DELIVERY_TIME is ABSENT - the item is undated, which is its whole purpose.</param>
/// <param name="ClassMatches">The message class is the kind's own.</param>
/// <param name="InTheTargetStore">
/// The saved item is in the TARGET store - its parent folder's StoreID is the target's. A save that
/// Outlook files somewhere else is a write into a store no allowlist named, which is what an unsent
/// mail item's first save does (see <see cref="CorpusItemKind"/>); this is the check that would catch
/// any other kind doing the same.
/// </param>
/// <param name="Error">The COM failure that stopped this kind's probe, or null.</param>
public sealed record CorpusUndatedProbe(
    CorpusItemKind Kind,
    bool FolderReachable,
    bool InTheFolder,
    bool SubjectTagParses,
    bool HasNoDeliveryTime,
    bool ClassMatches,
    bool InTheTargetStore,
    string? Error);

/// <summary>
/// The go/no-go on the undated probe, as a pure function: a population carrying undated items is only
/// built into a store where one throwaway item OF EACH KIND proved it lands in its folder, in the
/// target store, keeps its tag and carries no delivery time. No override, like every other probe.
/// </summary>
public static class CorpusUndatedFidelity
{
    /// <summary>Whether a build may proceed, and the sentence that says why either way.</summary>
    /// <param name="required">The undated kinds the population carries.</param>
    /// <param name="probes">What the probe found, one entry per kind it probed.</param>
    public static (bool Proceed, string Message) Decide(
        IReadOnlyList<CorpusItemKind> required, IReadOnlyList<CorpusUndatedProbe> probes)
    {
        ArgumentNullException.ThrowIfNull(required);
        ArgumentNullException.ThrowIfNull(probes);
        if (required.Count == 0)
        {
            return (true, "Undated probe: this population carries no undated item; nothing to probe.");
        }

        var failed = new List<string>();
        foreach (CorpusItemKind kind in required)
        {
            CorpusUndatedProbe? probe = probes.FirstOrDefault(p => p.Kind == kind);
            string name = kind.ToString().ToUpperInvariant();
            if (probe == null)
            {
                failed.Add(name + ": not probed at all");
                continue;
            }

            if (probe.Error != null)
            {
                failed.Add(name + ": the probe failed (" + probe.Error + ")");
                continue;
            }

            var why = new List<string>();
            if (!probe.FolderReachable)
            {
                why.Add("no folder for it could be found or made - the store could not say whether it has one");
            }

            if (!probe.InTheTargetStore)
            {
                why.Add("its save landed OUTSIDE the target store, a write no allowlist named");
            }

            if (!probe.InTheFolder)
            {
                why.Add("the saved item is not in that folder");
            }

            if (!probe.SubjectTagParses)
            {
                why.Add("its subject lost the corpus tag, so the two-key rule could never delete it");
            }

            if (!probe.HasNoDeliveryTime)
            {
                why.Add("it CARRIES a delivery time, so the index would date it");
            }

            if (!probe.ClassMatches)
            {
                why.Add("it read back as another kind of item");
            }

            if (why.Count > 0)
            {
                failed.Add(name + ": " + string.Join(", ", why));
            }
        }

        if (failed.Count > 0)
        {
            return (false, "REFUSING to build a population with undated items: " + string.Join("; ", failed) + ". "
                + "An undated item that is dated, misfiled or untagged is exactly what LiveOrderKeyCollationTests would "
                + "mis-measure, and an untagged one is an item teardown cannot remove. There is no flag for this.");
        }

        return (true, "Undated probe verified for " + string.Join(", ", required.Select(k => k.ToString().ToLowerInvariant()))
            + ": each lands in its own folder of the target store, keeps its tag, and carries no delivery time.");
    }
}
