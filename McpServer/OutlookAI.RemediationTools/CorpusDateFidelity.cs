using System.Globalization;

namespace OutlookAI.RemediationTools;

/// <summary>
/// How a corpus item's dates were written, in the order the builder tries them.
/// </summary>
public enum CorpusDateWriteMethod
{
    /// <summary>Nothing worked; the item's dates are whatever Outlook stamped on it.</summary>
    None = 0,

    /// <summary>
    /// PropertyAccessor writes of PR_MESSAGE_DELIVERY_TIME (0x0E060040) and
    /// PR_CLIENT_SUBMIT_TIME (0x00390040) - the two properties the freshness sweep's DASL
    /// filter reads, as <c>urn:schemas:httpmail:datereceived</c> and
    /// <c>urn:schemas:httpmail:date</c>.
    /// <para>
    /// This rung USED to also write PR_MESSAGE_FLAGS to clear MSGFLAG_UNSENT. That write has
    /// moved to <see cref="CorpusPlacement"/>, where it belongs: it decides where an item
    /// LIVES, not what it is dated. Bundling them meant that overriding the date refusal
    /// silently disabled the placement fix too, which is exactly what happened on the first
    /// real build - 40 000 items filed as drafts, invisible to the sweep.
    /// </para>
    /// </summary>
    PropertyAccessorDates = 1,

    /// <summary>
    /// Plain object-model assignment of MailItem.ReceivedTime. Weakest rung, tried last,
    /// and it can only ever carry the received date: MailItem.SentOn is read-only in the
    /// object model, so the submit half of the sweep filter goes unset on this rung.
    /// </summary>
    ObjectModel = 3,
}

/// <summary>
/// How the instant that came back compares to the instant that was asked for.
/// </summary>
public enum CorpusDateOffsetVerdict
{
    /// <summary>Read back what was written, within tolerance.</summary>
    Exact = 0,

    /// <summary>
    /// Read back displaced by exactly the machine's UTC offset. The PropertyAccessor is
    /// documented to convert a PT_SYSTIME write from LOCAL time and to return UTC, so a UTC
    /// value handed to it lands one offset away. Detectable, and correctable by
    /// pre-compensating the write - which the builder then re-probes rather than assumes.
    /// </summary>
    LocalOffsetApplied = 1,

    /// <summary>Read back something else entirely, or nothing. Unusable.</summary>
    Unusable = 2,
}

/// <summary>The outcome of one attempt to write dates on a throwaway probe item.</summary>
/// <param name="Method">Which rung this probe exercised.</param>
/// <param name="RequestedUtc">The instant the probe asked for.</param>
/// <param name="WrittenUtc">The instant actually handed to Outlook (may be pre-compensated).</param>
/// <param name="ReadBackReceivedUtc">MailItem.ReceivedTime after re-opening the saved item by EntryID; null when unreadable.</param>
/// <param name="DaslSelectedInWindow">True when a DASL restriction whose lower bound sits just before the requested instant returned the probe.</param>
/// <param name="DaslExcludedOutsideWindow">True when a DASL restriction whose lower bound sits AFTER the requested instant did NOT return the probe.</param>
/// <param name="Error">Why the rung failed, when it did.</param>
public sealed record CorpusDateProbe(
    CorpusDateWriteMethod Method,
    DateTime RequestedUtc,
    DateTime WrittenUtc,
    DateTime? ReadBackReceivedUtc,
    bool DaslSelectedInWindow,
    bool DaslExcludedOutsideWindow,
    string? Error);

/// <summary>
/// Decides, from probe results alone, whether a corpus can carry believable dates - and
/// says so out loud when it cannot.
/// <para>
/// <b>Why this exists at all.</b> A mail item's received time is not an ordinary property.
/// <c>MailItem.SentOn</c> is read-only in the Outlook object model, and a message created
/// directly in a folder is an UNSENT item, which some stores date themselves. The
/// documented way to place a message in the past is a PropertyAccessor write of the
/// underlying MAPI properties, and whether a given Outlook build and a given store honour
/// it cannot be settled by reading documentation - only by writing one item and reading it
/// back.
/// </para>
/// <para>
/// So the builder probes before it builds: one throwaway item per rung, saved, re-opened by
/// EntryID, its ReceivedTime read back, and - the part that actually matters - selected and
/// then NOT selected by the same kind of DASL date restriction the freshness sweep uses.
/// The probe items are deleted by the ordinary two-key rule. If no rung passes, the build
/// REFUSES unless the caller explicitly accepts an undated corpus, because a corpus whose
/// items are all dated "now" would make every window measurement meaningless while looking
/// exactly like a good one.
/// </para>
/// </summary>
public static class CorpusDateFidelity
{
    /// <summary>
    /// How far the read-back may sit from the request and still count as exact. MAPI stores
    /// PT_SYSTIME at 100 ns resolution, but round trips through the object model and the
    /// DASL layer are second-grained, so a sub-second difference proves nothing.
    /// </summary>
    public static readonly TimeSpan Tolerance = TimeSpan.FromSeconds(2);

    /// <summary>The rungs, strongest first. The builder takes the first that fully verifies.</summary>
    public static readonly CorpusDateWriteMethod[] Ladder =
    {
        CorpusDateWriteMethod.PropertyAccessorDates,
        CorpusDateWriteMethod.ObjectModel,
    };

    /// <summary>
    /// Classifies the read-back. <paramref name="localOffset"/> is the machine's UTC offset
    /// at the requested instant; pass <see cref="TimeSpan.Zero"/> on a UTC machine, where
    /// the offset verdict is indistinguishable from exact and correctly never fires.
    /// </summary>
    public static CorpusDateOffsetVerdict ClassifyOffset(DateTime requestedUtc, DateTime? readBackUtc, TimeSpan localOffset)
    {
        if (readBackUtc == null)
        {
            return CorpusDateOffsetVerdict.Unusable;
        }

        TimeSpan delta = readBackUtc.Value - requestedUtc;
        if (Abs(delta) <= Tolerance)
        {
            return CorpusDateOffsetVerdict.Exact;
        }

        if (localOffset != TimeSpan.Zero
            && (Abs(delta - localOffset) <= Tolerance || Abs(delta + localOffset) <= Tolerance))
        {
            return CorpusDateOffsetVerdict.LocalOffsetApplied;
        }

        return CorpusDateOffsetVerdict.Unusable;
    }

    /// <summary>
    /// What to hand Outlook so the item lands on <paramref name="requestedUtc"/>, given what
    /// a first probe showed. Only the offset verdict changes the value - a rung that reads
    /// back exactly is written exactly.
    /// </summary>
    public static DateTime CompensatedWriteValue(DateTime requestedUtc, CorpusDateOffsetVerdict verdict, TimeSpan localOffset, DateTime readBackUtc)
    {
        if (verdict != CorpusDateOffsetVerdict.LocalOffsetApplied)
        {
            return requestedUtc;
        }

        // Move the write by the same distance the read-back moved, in the opposite
        // direction. Derived from the observation rather than from an assumed sign, so it
        // is right on either side of the meridian and through a DST change.
        TimeSpan observed = readBackUtc - requestedUtc;
        return requestedUtc - observed;
    }

    /// <summary>
    /// A probe counts as usable only when all three signals agree: the read-back matched,
    /// a DASL restriction covering the instant returned the item, and a DASL restriction
    /// starting after the instant did not. The third is what separates a real date from a
    /// property that reads back correctly but does not drive selection - and selection is
    /// the only thing the window measurement depends on.
    /// <para>
    /// The read-back must be EXACT against what was REQUESTED, with no offset allowance.
    /// That is not an oversight: a rung whose first attempt landed one UTC offset away is
    /// re-run with a pre-compensated write, and it is that second probe which is judged
    /// here - so "exact" means the correction worked, and a rung still one offset out after
    /// correction is genuinely unusable.
    /// </para>
    /// </summary>
    public static bool IsUsable(CorpusDateProbe probe)
    {
        ArgumentNullException.ThrowIfNull(probe);
        return probe.Error == null
            && probe.DaslSelectedInWindow
            && probe.DaslExcludedOutsideWindow
            && ClassifyOffset(probe.RequestedUtc, probe.ReadBackReceivedUtc, localOffset: TimeSpan.Zero)
                == CorpusDateOffsetVerdict.Exact;
    }

    /// <summary>
    /// The rung to build with: the first in <see cref="Ladder"/> order that fully verified.
    /// <see cref="CorpusDateWriteMethod.None"/> when none did.
    /// </summary>
    public static CorpusDateWriteMethod Choose(IReadOnlyCollection<CorpusDateProbe> probes)
    {
        ArgumentNullException.ThrowIfNull(probes);
        foreach (CorpusDateWriteMethod method in Ladder)
        {
            foreach (CorpusDateProbe probe in probes)
            {
                if (probe.Method == method && IsUsable(probe))
                {
                    return method;
                }
            }
        }

        return CorpusDateWriteMethod.None;
    }

    /// <summary>
    /// Whether the build may proceed, and what to print either way. An undated corpus is
    /// allowed only when the caller asked for it in so many words, and even then the message
    /// says plainly what such a corpus can and cannot be used to measure.
    /// </summary>
    public static (bool Proceed, string Message) Decide(CorpusDateWriteMethod chosen, bool allowUndated, int itemCount)
    {
        if (chosen != CorpusDateWriteMethod.None)
        {
            string caveat = chosen == CorpusDateWriteMethod.ObjectModel
                ? " Submit time (PR_CLIENT_SUBMIT_TIME) is NOT set on this rung, so a filter reading "
                    + "urn:schemas:httpmail:date alone will not see these items; the sweep reads both, so it will."
                : string.Empty;
            return (true, $"Date fidelity: VERIFIED via {chosen}. Received dates drive DASL selection." + caveat);
        }

        // CORRECTED 2026-08-19. This message used to say the items "would carry a received
        // time of roughly 'now'", from which an operator reasonably concluded that an
        // all-recent corpus is still the sweep's worst case, overrode the guard, and lost a
        // 12-minute build. That conclusion was wrong because the premise was: an item whose
        // delivery time is not readable through the folder table is not SELECTED by a date
        // restriction at all. The sweep filters on
        // (datereceived >= X) OR (date >= X), so such items are invisible to it - not
        // recent, not old, absent. The consequence is now stated as a count, because a
        // number cannot be reasoned around the way a description can.
        string what = "Date fidelity: NOT ACHIEVABLE on this store. No write method placed an item in the past "
            + "and had a DASL date restriction select it. The freshness sweep selects with "
            + "(datereceived >= X) OR (date >= X), so items whose delivery time the folder table does not carry "
            + "are not selected by ANY window - they are invisible to the sweep, not merely mis-dated. This is "
            + "NOT 'an all-recent corpus', which would at least be the sweep's worst case: a sweep would select "
            + $"0 of {itemCount.ToString("N0", CultureInfo.InvariantCulture)} items, and both a 7-day and a "
            + "60-day window would return nothing.";

        return allowUndated
            ? (true, what + " Proceeding because --allow-undated was given. This corpus can still be used for "
                + "out-of-band per-item and per-folder timing (measurement plan step 2), body-cap behaviour and "
                + "frame size at a known item count and known body sizes; it CANNOT be used to measure the "
                + "freshness sweep or any date window.")
            : (false, what + " Refusing to build.");
    }

    private static TimeSpan Abs(TimeSpan value) => value < TimeSpan.Zero ? -value : value;
}
