using System.Globalization;

namespace OutlookAI.RemediationTools;

/// <summary>One item a re-anchor has to write, with the instants it must end up carrying.</summary>
/// <param name="Ordinal">The item's ordinal in the plan.</param>
/// <param name="EntryId">Its EntryID as the manifest records it.</param>
/// <param name="ReceivedUtc">The received instant to write.</param>
/// <param name="SentUtc">The submit instant to write; it keeps its original distance from the received one.</param>
/// <param name="FolderId">The folder id the manifest already records, carried so a replacement line can restate it.</param>
/// <param name="BodyBytes">The body size the manifest already records, carried for the same reason.</param>
/// <remarks>
/// <para>
/// <b>Why the folder and body size travel with an item that is only having its dates
/// rewritten.</b> A re-anchor appends a replacement manifest line per item, and the manifest
/// reader is last-writer-wins <i>per ordinal, wholesale</i> - it does
/// <c>_items[ordinal] = item</c>, not a per-field merge. So a replacement line that leaves
/// these at zero does not "decline to restate what it does not know"; it DELETES what the
/// build recorded. Measured on the test VM 2026-08-24: after a re-anchor of 20,000 items,
/// every entry read back with <c>FolderId 0, BodyBytes 0</c>.
/// </para>
/// </remarks>
public sealed record CorpusReanchorItem(
    int Ordinal,
    string EntryId,
    DateTime ReceivedUtc,
    DateTime SentUtc,
    int FolderId,
    int BodyBytes);

/// <summary>What a re-anchor is going to do, decided before a single item is opened.</summary>
/// <param name="TargetShiftSeconds">The shift from the PLAN's instants that the store must end up carrying.</param>
/// <param name="TargetAnchorUtc">Where the corpus's newest edge will sit afterwards.</param>
/// <param name="Todo">Items whose recorded instant is not already the target one.</param>
/// <param name="AlreadyCorrect">Items already carrying the target instant - skipped, which is what makes this resumable.</param>
/// <param name="Unrecorded">Ordinals in range that the manifest does not record at all, so nothing here can address them.</param>
/// <param name="Undated">Manifest entries carrying no received instant, so nothing can say what they hold now.</param>
public sealed record CorpusReanchorPlan(
    long TargetShiftSeconds,
    DateTime TargetAnchorUtc,
    IReadOnlyList<CorpusReanchorItem> Todo,
    int AlreadyCorrect,
    int Unrecorded,
    int Undated)
{
    /// <summary>The target shift as a span.</summary>
    public TimeSpan TargetShift => TimeSpan.FromSeconds(TargetShiftSeconds);
}

/// <summary>
/// Moves a whole corpus forward in time without regenerating it.
/// <para>
/// <b>Why an ABSOLUTE target rather than an increment.</b> The shift is always expressed as
/// the offset from the PLAN's own instants - never as "add another six weeks". So running a
/// re-anchor twice is a no-op rather than a double shift, an interrupted one is finished by
/// running it again, and the state of the store is a function of the target alone rather
/// than of how many times anything has been run. That is the same discipline
/// <see cref="CorpusManifest.MissingOrdinals"/> uses for the build: derive the work from the
/// data, never from a cursor.
/// </para>
/// <para>
/// <b>Why the manifest header is not rewritten.</b> The header's anchor is half of the
/// corpus's identity - it feeds <see cref="CorpusPlanOptions.ShapeKey"/>, which is what lets
/// a resumed build prove it is adding to the same corpus and lets a teardown prove it is
/// deleting from one. Rewriting it would make every later <c>--anchor</c> argument wrong and
/// would quietly sever the corpus from the seed that describes it. So the anchor stays put
/// and the shift is DERIVED, by comparing what the manifest records each item as carrying
/// against what the plan says it should have been given. The manifest is append-only and its
/// item lines are last-writer-wins by ordinal, so a re-anchor records its work by appending
/// a fresh line per item - which is crash-safe, costs one line per item, and leaves the file
/// showing both the old value and the new one in the order they happened.
/// </para>
/// </summary>
public static class CorpusReanchor
{
    /// <summary>
    /// A recorded instant is treated as matching a computed one when they are this close.
    /// Outlook stores delivery times to the second and the manifest renders them to the
    /// second, so anything finer would make every item look wrong and re-write the whole
    /// corpus on every run.
    /// </summary>
    public static readonly TimeSpan MatchTolerance = TimeSpan.FromSeconds(1);

    /// <summary>
    /// Whether an item's dates actually landed: does what the store reports after the write
    /// match what the write intended, to <see cref="MatchTolerance"/>?
    /// <para>
    /// <b>This exists because a re-anchor once destroyed a corpus while reporting success.</b>
    /// Measured on the test VM 2026-08-24: a run over 20,000 items finished
    /// "rewritten 20,000, refused 0, failed 0" and left every item dated inside the six
    /// minutes the tool had been running - the whole age-band structure the corpus exists for
    /// replaced by "everything arrived while the tool ran". The write path already READ the
    /// value back afterwards; it simply never compared it, and then recorded the read-back
    /// into the manifest as though it were the intention, destroying the manifest too.
    /// </para>
    /// <para>
    /// The date-write method is chosen by a probe that creates throwaway items, and the
    /// re-anchor's own dry run says plainly that "the date probe was not run (it creates
    /// items)" - so the method is proven for NEW items and reused unverified on EXISTING
    /// ones. That gap is why a per-item check, rather than a per-run one, is the right shape:
    /// the very first item answers it.
    /// </para>
    /// </summary>
    /// <param name="intendedUtc">The instant the write asked for.</param>
    /// <param name="readBackUtc">What the item reports afterwards; null when it could not be read.</param>
    /// <returns>True only when the store demonstrably carries the intended instant.</returns>
    public static bool WriteLanded(DateTime intendedUtc, DateTime? readBackUtc) =>
        readBackUtc.HasValue
        && Math.Abs((readBackUtc.Value - intendedUtc).TotalSeconds) <= MatchTolerance.TotalSeconds;

    /// <summary>
    /// The replacement manifest line for an item whose dates have just been rewritten.
    /// <para>
    /// Pure, and separate from the write, because the write lives on a COM path no CI test can
    /// enter - and the bug this closes was exactly there. Zeroing <c>FolderId</c> and
    /// <c>BodyBytes</c> here does not decline to restate what a re-anchor does not know: the
    /// manifest reader is last-writer-wins per ordinal WHOLESALE, so a zeroed replacement
    /// DELETES what the build recorded. Both are carried on <see cref="CorpusReanchorItem"/>
    /// for this one reason.
    /// </para>
    /// </summary>
    /// <param name="item">The item as planned, carrying the fields the manifest already holds.</param>
    /// <param name="intendedUtc">The instant the write asked for.</param>
    /// <param name="readBackUtc">What the store reports; the recorded value when it is readable.</param>
    public static CorpusManifestItem ReplacementLine(
        CorpusReanchorItem item, DateTime intendedUtc, DateTime? readBackUtc)
    {
        ArgumentNullException.ThrowIfNull(item);
        return new CorpusManifestItem(
            item.Ordinal,
            item.EntryId,
            item.FolderId,
            item.BodyBytes,
            CorpusManifest.FormatUtc(readBackUtc ?? intendedUtc));
    }

    /// <summary>
    /// What to say when a write did not land. It names both instants, because the difference
    /// between them is the diagnosis: a value near "now" means the store re-stamped the item,
    /// and an unreadable one means the write threw somewhere that swallowed it.
    /// </summary>
    public static string DescribeWriteRefusal(int ordinal, DateTime intendedUtc, DateTime? readBackUtc)
    {
        string got = readBackUtc.HasValue
            ? readBackUtc.Value.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture)
            : "unreadable";
        return $"Item {ordinal}: the date write did not land. Asked for "
            + intendedUtc.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture)
            + $", the store reports {got}. Stopping before the rest of the corpus is touched - "
            + "a re-anchor that cannot move one item's dates cannot move twenty thousand, and "
            + "continuing would rewrite every remaining item with a value nothing has checked.";
    }

    /// <summary>
    /// The share of dated manifest entries that must agree on one shift before the answer is
    /// trusted. Below it the corpus is not one corpus with one shift - it is a store whose
    /// dates nobody can account for, and guessing a shift there would produce a freshness
    /// verdict about a corpus that does not exist.
    /// </summary>
    public const double ShiftAgreementFloor = 0.90;

    /// <summary>
    /// What shift the store is already carrying, derived from the manifest rather than
    /// recorded anywhere. For each dated entry it compares the instant Outlook reported
    /// after the write against the instant the plan asked for; the MODE of those differences
    /// is the answer, because a handful of items whose date write failed must not drag it.
    /// </summary>
    /// <param name="plan">The corpus shape.</param>
    /// <param name="manifest">What the build (and any earlier re-anchor) recorded.</param>
    /// <param name="agreeing">How many dated entries carry the returned shift.</param>
    /// <param name="dated">How many entries carried a received instant at all.</param>
    /// <returns>
    /// The shift, and whether it is provable - false when nothing is dated, or when fewer
    /// than <see cref="ShiftAgreementFloor"/> of the dated entries agree.
    /// </returns>
    public static (TimeSpan Shift, bool Provable) DeriveAppliedShift(
        CorpusPlan plan, CorpusManifest manifest, out int agreeing, out int dated)
    {
        ArgumentNullException.ThrowIfNull(plan);
        ArgumentNullException.ThrowIfNull(manifest);

        var tally = new Dictionary<long, int>();
        dated = 0;
        foreach (CorpusManifestItem item in manifest.Items.Values)
        {
            DateTime? recorded = CorpusManifest.ParseUtc(item.ReceivedUtc);
            if (recorded == null || item.Ordinal < 1)
            {
                continue;
            }

            dated++;
            long seconds = (long)Math.Round((recorded.Value - plan.Describe(item.Ordinal).ReceivedUtc).TotalSeconds);
            tally[seconds] = tally.TryGetValue(seconds, out int n) ? n + 1 : 1;
        }

        if (dated == 0)
        {
            agreeing = 0;
            return (TimeSpan.Zero, false);
        }

        long mode = 0;
        int best = 0;
        foreach (KeyValuePair<long, int> pair in tally)
        {
            // Ties break toward the smaller shift, so a corpus split exactly in half by an
            // interrupted re-anchor reports the state it is being moved AWAY from, and the
            // re-anchor that follows still computes the same absolute target.
            if (pair.Value > best || (pair.Value == best && pair.Key < mode))
            {
                mode = pair.Key;
                best = pair.Value;
            }
        }

        agreeing = best;
        return (TimeSpan.FromSeconds(mode), (double)best / dated >= ShiftAgreementFloor);
    }

    /// <summary>
    /// The work a re-anchor to <paramref name="targetAnchorUtc"/> has to do. Nothing is
    /// opened and nothing is written here: this is the sheet the COM half executes and the
    /// operator reads first.
    /// </summary>
    /// <param name="plan">The corpus shape.</param>
    /// <param name="manifest">What exists, and what each item currently carries.</param>
    /// <param name="itemCount">The corpus size - ordinals 1..itemCount are considered.</param>
    /// <param name="targetAnchorUtc">Where the corpus's newest edge should end up.</param>
    public static CorpusReanchorPlan Build(
        CorpusPlan plan, CorpusManifest manifest, int itemCount, DateTime targetAnchorUtc)
    {
        ArgumentNullException.ThrowIfNull(plan);
        ArgumentNullException.ThrowIfNull(manifest);
        if (itemCount < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(itemCount), "A corpus with no items cannot be re-anchored.");
        }

        DateTime target = DateTime.SpecifyKind(targetAnchorUtc, DateTimeKind.Utc);
        TimeSpan shift = target - DateTime.SpecifyKind(plan.Options.AnchorUtc, DateTimeKind.Utc);

        var todo = new List<CorpusReanchorItem>();
        int alreadyCorrect = 0;
        int unrecorded = 0;
        int undated = 0;
        for (int ordinal = 1; ordinal <= itemCount; ordinal++)
        {
            if (!manifest.Items.TryGetValue(ordinal, out CorpusManifestItem? recorded))
            {
                unrecorded++;
                continue;
            }

            CorpusItemSpec spec = plan.Describe(ordinal);
            DateTime wantReceived = spec.ReceivedUtc + shift;
            DateTime wantSent = spec.SentUtc + shift;
            DateTime? has = CorpusManifest.ParseUtc(recorded.ReceivedUtc);
            if (has == null)
            {
                // Nothing recorded means nothing known, and an item whose current instant is
                // unknown is written rather than assumed correct.
                undated++;
                todo.Add(new CorpusReanchorItem(
                    ordinal, recorded.EntryId, wantReceived, wantSent, recorded.FolderId, recorded.BodyBytes));
                continue;
            }

            if (Math.Abs((has.Value - wantReceived).TotalSeconds) <= MatchTolerance.TotalSeconds)
            {
                alreadyCorrect++;
                continue;
            }

            todo.Add(new CorpusReanchorItem(
                ordinal, recorded.EntryId, wantReceived, wantSent, recorded.FolderId, recorded.BodyBytes));
        }

        return new CorpusReanchorPlan(
            (long)Math.Round(shift.TotalSeconds), target, todo, alreadyCorrect, unrecorded, undated);
    }

    /// <summary>
    /// Whether the re-anchor may run, and what to print either way. It refuses to move a
    /// corpus BACKWARDS by default: the only reason to re-anchor is that the clock has moved
    /// on, and a backwards shift is far more likely to be a mistyped date than an intention.
    /// </summary>
    /// <param name="plan">The re-anchor work sheet.</param>
    /// <param name="appliedShift">What the store carries now.</param>
    /// <param name="allowBackwards">Whether a target older than the current one is permitted.</param>
    public static (bool Proceed, string Message) Decide(
        CorpusReanchorPlan plan, TimeSpan appliedShift, bool allowBackwards)
    {
        ArgumentNullException.ThrowIfNull(plan);
        CultureInfo invariant = CultureInfo.InvariantCulture;
        TimeSpan delta = plan.TargetShift - appliedShift;
        string what = $"Re-anchor: target {CorpusManifest.FormatUtc(plan.TargetAnchorUtc)}, "
            + $"a net move of {FormatDelta(delta)} from where the store sits now. "
            + $"{plan.Todo.Count.ToString("N0", invariant)} item(s) to write, "
            + $"{plan.AlreadyCorrect.ToString("N0", invariant)} already correct"
            + (plan.Undated > 0 ? $", {plan.Undated.ToString("N0", invariant)} with no recorded instant" : string.Empty)
            + (plan.Unrecorded > 0 ? $", {plan.Unrecorded.ToString("N0", invariant)} not in the manifest at all" : string.Empty)
            + ".";

        if (delta < TimeSpan.Zero && !allowBackwards)
        {
            return (false, what + " REFUSING: that moves the corpus BACKWARDS. Re-anchoring exists because the clock "
                + "moved on, so a backwards target is far more likely to be a mistyped date than an intention. Pass "
                + "--allow-backwards if it really is one.");
        }

        if (plan.Unrecorded > 0)
        {
            return (false, what + " REFUSING: items the manifest does not record cannot be addressed by EntryID, so "
                + "this run would shift part of the corpus and leave the rest where it is - which is worse than "
                + "either end state, because the derived shift would then agree with nothing. Rebuild the manifest "
                + "with corpus-reindex first.");
        }

        return (true, what);
    }

    private static string FormatDelta(TimeSpan span)
    {
        CultureInfo invariant = CultureInfo.InvariantCulture;
        TimeSpan abs = span < TimeSpan.Zero ? span.Negate() : span;
        string sign = span < TimeSpan.Zero ? "-" : "+";
        return sign + ((int)abs.TotalDays).ToString(invariant) + "d " + abs.Hours.ToString(invariant) + "h "
            + abs.Minutes.ToString(invariant) + "m";
    }
}
