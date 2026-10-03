using System.Globalization;
using OutlookAI.RemediationTools;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>A received-date window a live test asks the index for, and what it was anchored to.</summary>
/// <param name="OnOrAfterUtc">The window's start, inclusive.</param>
/// <param name="BeforeUtc">Its end, exclusive - or null for "up to now", which only the clock-anchored window uses.</param>
/// <param name="Basis">What the window was anchored to, for the test's output.</param>
public sealed record LiveDateWindow(DateTime OnOrAfterUtc, DateTime? BeforeUtc, string Basis);

/// <summary>
/// Anchors a live test's date window to the DATA when the machine declares data with a fixed anchor,
/// and to the clock only when it does not - decided by the maintainer 2026-10-03 (Q130 (b) of
/// <c>Docs/overnight-review-2026-10-03.md</c>), together with freezing the Outlook guests' clocks (Q130 (a)).
/// <para>
/// <b>Why.</b> <c>LiveIndexSearchTests.ProbeParity_DateRangeQuery_HitsUnder2s</c> times an unscoped
/// date-range query. It used to ask for "the last 30 days" of the clock, so on the indexed guest it
/// timed a predicate that matched the 160,000-item Corpus A's rows only while that corpus was younger
/// than 30 days - after 2026-11-01 it would still pass, on the hub's few dozen rows, while no longer
/// timing what it is for. Anchored to the corpus it asks the same question on every run, whatever the
/// clock says: the 30 days BEFORE the corpus's anchor, which select the same 24,596 of Corpus A's items
/// for ever (<c>Testbed/testbed.json</c>, the vm-indexed record's <c>selectedByWindowDays</c>).
/// </para>
/// <para>
/// <b>Where nothing has a fixed anchor</b> - a machine whose settings declare no corpus, such as the
/// maintainer's workstation, whose mail is always current - the window stays what it was: the last
/// N days of the clock, open-ended. The clock is still the only thing that makes that one meaningful.
/// </para>
/// <para>Pure: the settings block and a clock in, a window out. <c>T1/LiveDataWindowTests</c> pins every branch.</para>
/// </summary>
public static class LiveDataWindow
{
    /// <summary>
    /// The <paramref name="days"/> days before the declared corpus's anchor, or - with no corpus - the
    /// last <paramref name="days"/> days before <paramref name="nowUtc"/>.
    /// </summary>
    public static LiveDateWindow Before(CorpusSettings? corpus, int days, DateTime nowUtc)
    {
        if (days <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(days), days, "A date window is at least one day wide.");
        }

        CultureInfo invariant = CultureInfo.InvariantCulture;
        if (corpus != null && !string.IsNullOrWhiteSpace(corpus.AnchorUtc))
        {
            // An anchor that does not parse is refused, never replaced by the clock: falling back would
            // put back exactly the clock-relative window this exists to remove, and say nothing.
            DateTime anchor = CorpusManifest.ParseUtc(corpus.AnchorUtc)
                ?? throw new InvalidOperationException(
                    $"The corpus anchor '{corpus.AnchorUtc}' is not an instant, so no window can be anchored to it. "
                    + "Use yyyy-MM-dd or yyyy-MM-ddTHH:mm:ssZ - the corpus manifest's header line says which.");
            return new LiveDateWindow(
                anchor.AddDays(-days),
                anchor,
                string.Format(
                    invariant,
                    "the {0} days before corpus '{1}''s anchor {2} - the data's, not the clock's",
                    days,
                    corpus.CorpusId,
                    CorpusManifest.FormatUtc(anchor)));
        }

        DateTime now = DateTime.SpecifyKind(nowUtc, DateTimeKind.Utc);
        return new LiveDateWindow(
            now.AddDays(-days),
            null,
            string.Format(invariant, "the last {0} days of this machine's clock - it declares no corpus", days));
    }
}
