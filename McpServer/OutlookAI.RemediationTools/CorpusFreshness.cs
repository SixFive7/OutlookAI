using System.Globalization;

namespace OutlookAI.RemediationTools;

/// <summary>What a freshness check concluded about a corpus.</summary>
public enum CorpusFreshnessVerdict
{
    /// <summary>Every window under test still selects at least one item.</summary>
    Fresh = 0,

    /// <summary>
    /// At least one window under test selects NOTHING, while the widest still selects
    /// something. This is the state the check exists for: the narrow windows have gone
    /// quiet while every test asking about them still passes.
    /// </summary>
    WindowsEmptied = 1,

    /// <summary>
    /// The newest item is older than the WIDEST window under test, so no window selects
    /// anything. The corpus is a store full of mail that no measurement can see.
    /// </summary>
    Dead = 2,

    /// <summary>
    /// The check could not be made: the manifest records no received instants, or too few
    /// of them agree on one shift to call it. Deliberately NOT treated as fresh - an
    /// unprovable freshness claim is exactly the silence this check replaces.
    /// </summary>
    Unprovable = 3,
}

/// <summary>One measurement window, and what the corpus can still offer it.</summary>
/// <param name="Days">Window width in days, measured back from the moment of the check.</param>
/// <param name="PlannedCount">How many items the plan puts in a window this wide, measured back from the anchor.</param>
/// <param name="StillInWindow">How many items are still inside it today, given the shift already applied.</param>
public sealed record CorpusWindowFreshness(int Days, int PlannedCount, int StillInWindow);

/// <summary>Everything the freshness check found, so a failure names its own remedy.</summary>
/// <param name="PlanAnchorUtc">The anchor the corpus was generated against. Never changes; it is half of the seed's identity.</param>
/// <param name="AppliedShiftSeconds">The shift already applied to the store by earlier re-anchors, in seconds.</param>
/// <param name="EffectiveAnchorUtc">Where the corpus's newest edge actually sits now: the plan anchor plus the applied shift.</param>
/// <param name="AsOfUtc">The instant the check was made against.</param>
/// <param name="AgeSeconds">How far the effective anchor has fallen behind <paramref name="AsOfUtc"/>.</param>
/// <param name="Windows">Per-window findings, narrowest first.</param>
/// <param name="Verdict">The conclusion.</param>
public sealed record CorpusFreshnessReport(
    DateTime PlanAnchorUtc,
    long AppliedShiftSeconds,
    DateTime EffectiveAnchorUtc,
    DateTime AsOfUtc,
    long AgeSeconds,
    IReadOnlyList<CorpusWindowFreshness> Windows,
    CorpusFreshnessVerdict Verdict)
{
    /// <summary>The shift already applied, as a span.</summary>
    public TimeSpan AppliedShift => TimeSpan.FromSeconds(AppliedShiftSeconds);

    /// <summary>How far the corpus has fallen behind, as a span.</summary>
    public TimeSpan Age => TimeSpan.FromSeconds(AgeSeconds);

    /// <summary>Windows that now select nothing, narrowest first. Empty when the corpus is fresh.</summary>
    public IReadOnlyList<int> EmptiedWindowDays
        => Windows.Where(w => w.StillInWindow == 0).Select(w => w.Days).ToList();
}

/// <summary>
/// Whether a generated corpus can still answer the questions it was built to answer.
/// <para>
/// <b>Why this exists.</b> The corpus is generated against a FIXED anchor, and every test
/// asking about "the last N days" selects against the CLOCK. The two diverge from the
/// moment the corpus is written. About six weeks after generation a seven-day window
/// selects nothing at all - and every test asking about that window still PASSES, because
/// "no mail in the last seven days" is a valid answer about an empty window. The corpus
/// does not break loudly; it quietly stops being a corpus, and the suite keeps reporting
/// green over a measurement that is no longer being taken.
/// </para>
/// <para>
/// <b>What it does about it.</b> Nothing here repairs anything - it converts the silence
/// into a failure that names the remedy. The repair is <see cref="CorpusReanchor"/>, which
/// shifts every item forward by the elapsed offset. Regenerating instead was considered and
/// rejected: the corpus is a snapshot that measurements are held against, and a regenerated
/// one would be a different population wearing the same numbers.
/// </para>
/// <para>
/// <b>Why it is stricter than "older than the widest window".</b> A corpus whose newest item
/// has fallen past the widest window is DEAD - no window selects anything - and that is the
/// floor. But the failure being fixed happens long before the floor is reached: the 7-day
/// window empties at six weeks while the 365-day window still selects 22 000 items, and a
/// test asserting on the 7-day window is a lie from the moment it empties. So any window
/// under test that selects nothing fails the check, and the report names which ones.
/// </para>
/// </summary>
public static class CorpusFreshness
{
    /// <summary>
    /// Evaluates a corpus against the clock.
    /// </summary>
    /// <param name="plan">The corpus shape - the source of every item's intended received instant.</param>
    /// <param name="itemCount">How many ordinals the corpus holds.</param>
    /// <param name="appliedShift">
    /// The shift already applied to the store by earlier re-anchors, as derived from the
    /// manifest by <see cref="CorpusReanchor.DeriveAppliedShift"/>. <see cref="TimeSpan.Zero"/>
    /// for a corpus that has never been re-anchored.
    /// </param>
    /// <param name="asOfUtc">The instant to measure against - the caller's clock, never taken here.</param>
    /// <param name="windowDays">
    /// The windows under test. Defaults to <see cref="CorpusPlan.MeasurementWindowDays"/>.
    /// The caller passes its own set when it knows it tests fewer: a suite that never asks
    /// about a one-day window should not be stopped by a one-day window having emptied.
    /// </param>
    /// <param name="shiftProvable">
    /// False when the manifest could not tell us what shift is already applied. The verdict
    /// is then <see cref="CorpusFreshnessVerdict.Unprovable"/> whatever the counts say.
    /// </param>
    public static CorpusFreshnessReport Evaluate(
        CorpusPlan plan,
        int itemCount,
        TimeSpan appliedShift,
        DateTime asOfUtc,
        IReadOnlyList<int>? windowDays = null,
        bool shiftProvable = true)
    {
        ArgumentNullException.ThrowIfNull(plan);
        if (itemCount < 1)
        {
            throw new ArgumentOutOfRangeException(nameof(itemCount), "A corpus with no items has no freshness.");
        }

        DateTime asOf = DateTime.SpecifyKind(asOfUtc, DateTimeKind.Utc);
        IReadOnlyList<int> marks = (windowDays ?? CorpusPlan.MeasurementWindowDays)
            .Where(d => d > 0).Distinct().OrderBy(d => d).ToList();
        if (marks.Count == 0)
        {
            throw new ArgumentException("At least one window is needed.", nameof(windowDays));
        }

        DateTime planAnchor = DateTime.SpecifyKind(plan.Options.AnchorUtc, DateTimeKind.Utc);
        DateTime effectiveAnchor = planAnchor + appliedShift;

        var planned = new int[marks.Count];
        var live = new int[marks.Count];
        for (int ordinal = 1; ordinal <= itemCount; ordinal++)
        {
            DateTime intended = plan.Describe(ordinal).ReceivedUtc;
            DateTime actual = intended + appliedShift;
            for (int i = 0; i < marks.Count; i++)
            {
                // "Planned" is the sheet: what a window this wide selected on the day the
                // corpus was anchored. "Live" is what it selects now. Reporting both is what
                // turns "the 7-day window is empty" into "the 7-day window held 3 180 items
                // and now holds none", which is the difference between a number and a cause.
                if (intended > planAnchor.AddDays(-marks[i]))
                {
                    planned[i]++;
                }

                if (actual > asOf.AddDays(-marks[i]))
                {
                    live[i]++;
                }
            }
        }

        var windows = new List<CorpusWindowFreshness>(marks.Count);
        for (int i = 0; i < marks.Count; i++)
        {
            windows.Add(new CorpusWindowFreshness(marks[i], planned[i], live[i]));
        }

        CorpusFreshnessVerdict verdict;
        if (!shiftProvable)
        {
            verdict = CorpusFreshnessVerdict.Unprovable;
        }
        else if (windows[^1].StillInWindow == 0)
        {
            verdict = CorpusFreshnessVerdict.Dead;
        }
        else if (windows.Any(w => w.StillInWindow == 0))
        {
            verdict = CorpusFreshnessVerdict.WindowsEmptied;
        }
        else
        {
            verdict = CorpusFreshnessVerdict.Fresh;
        }

        return new CorpusFreshnessReport(
            planAnchor,
            (long)Math.Round(appliedShift.TotalSeconds),
            effectiveAnchor,
            asOf,
            (long)Math.Round((asOf - effectiveAnchor).TotalSeconds),
            windows,
            verdict);
    }

    /// <summary>
    /// Whether a run may proceed on this corpus, and what to print either way. The message
    /// states the consequence as a COUNT and names the command that repairs it, for the same
    /// reason the placement guard does: the date guard's original prose refusal invited a
    /// reasonable inference that happened to be wrong, and an operator overrode it.
    /// </summary>
    public static (bool Proceed, string Message) Decide(CorpusFreshnessReport report)
    {
        ArgumentNullException.ThrowIfNull(report);
        CultureInfo invariant = CultureInfo.InvariantCulture;
        string anchors = $"anchor {CorpusManifest.FormatUtc(report.PlanAnchorUtc)}"
            + (report.AppliedShiftSeconds == 0
                ? " (never re-anchored)"
                : $" shifted by {FormatAge(report.AppliedShift)} to {CorpusManifest.FormatUtc(report.EffectiveAnchorUtc)}")
            + $", {FormatAge(report.Age)} behind the clock";

        if (report.Verdict == CorpusFreshnessVerdict.Unprovable)
        {
            return (false, "Freshness: UNPROVABLE. The manifest does not record enough received instants to say what "
                + "shift the store already carries, so no window count here would mean anything. Refusing. Rebuild the "
                + "manifest with corpus-reindex, or rebuild the corpus. " + anchors + ".");
        }

        string counts = string.Join(", ", report.Windows.Select(w =>
            w.Days.ToString(invariant) + "d=" + w.StillInWindow.ToString("N0", invariant)
            + "/" + w.PlannedCount.ToString("N0", invariant)));

        if (report.Verdict == CorpusFreshnessVerdict.Fresh)
        {
            return (true, $"Freshness: OK - {anchors}. Windows now/at-anchor: {counts}.");
        }

        string what = report.Verdict == CorpusFreshnessVerdict.Dead
            ? "Freshness: DEAD. The newest item is older than the widest window under test, so EVERY window selects "
                + "0 items."
            : "Freshness: STALE. These windows now select 0 items: "
                + string.Join(", ", report.EmptiedWindowDays.Select(d => d.ToString(invariant) + "d")) + ".";

        return (false, what + $" Counts are now/at-anchor: {counts}. {anchors}. "
            + "A test asserting on an emptied window still PASSES - selecting nothing is a valid answer about an "
            + "empty window - which is why this is a refusal and not a warning. REBUILD the corpus: "
            + "'corpus-teardown --execute' (or delete the .pst), then 'corpus-build' with the same seed and count "
            + "and an --anchor at or near today. That is deterministic, and the recorded build was 20,000 items in "
            + "13m25s. A rebuild from the same seed is the same population - the plan is a pure function of the "
            + "seed and the ordinal, with no clock in it, so only the anchor moves. Do NOT use 'corpus-reanchor': "
            + "it is retired, its date writes do not land on already-delivered items, and it once reported "
            + "rewriting 20,000 items successfully while dating every one of them inside the six minutes it had "
            + "been running.");
    }

    /// <summary>Renders a span as days and hours, which is the resolution this decision is made at.</summary>
    private static string FormatAge(TimeSpan span)
    {
        CultureInfo invariant = CultureInfo.InvariantCulture;
        TimeSpan abs = span < TimeSpan.Zero ? span.Negate() : span;
        string sign = span < TimeSpan.Zero ? "-" : string.Empty;
        return abs.TotalDays >= 1
            ? sign + ((int)abs.TotalDays).ToString(invariant) + "d " + abs.Hours.ToString(invariant) + "h"
            : sign + ((int)abs.TotalHours).ToString(invariant) + "h " + abs.Minutes.ToString(invariant) + "m";
    }
}
