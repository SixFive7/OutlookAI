using System.Text;

using OutlookAI.Core.Com;

using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// READ-ONLY probe that settles one question: <b>does the freshness sweep's sort call actually
/// sort?</b>
/// <para>
/// The sweep orders each folder's table newest-first with
/// <c>Table.Sort("urn:schemas:httpmail:datereceived", true)</c>, and Microsoft's
/// <c>Table.Sort</c> reference states that sort properties may be referenced "by their
/// explicit string names only; cannot reference properties by their namespaces". That call
/// passes a namespace. If the documentation holds for it, then the sweep has NEVER sorted, on
/// any store, for any user - its 200-item-per-folder cap has always cut an arbitrary slice
/// rather than dropping the oldest of the window, the tier whose entire purpose is recent mail
/// has been returning an arbitrary 200 instead of the newest 200, and this repository's
/// reading of <c>item_cap_unsorted</c> ("the sort genuinely does not apply on that store") is
/// wrong in a way that matters: it would not apply anywhere.
/// </para>
/// <para>
/// <b>Why this is a test and not a script.</b> Four read-only PowerShell probes failed
/// identically on every property form AND on the no-argument form, with
/// <c>DISP_E_PARAMNOTOPTIONAL</c>. That uniformity is the tell: Outlook rejecting a
/// namespace-qualified property would not also reject <c>ReceivedTime</c> called with no
/// order argument. The failure was PowerShell's late binding against the <c>Table</c> COM
/// object, not Outlook's verdict. Early-bound C# has no such problem.
/// </para>
/// <para>
/// <b>What it touches.</b> It opens one table per store, adds a column, calls <c>Sort</c>, and
/// reads the FIRST ROW. It opens no item, creates nothing, moves nothing, deletes nothing, and
/// never calls <c>Application.Quit</c>. Output is dates and EntryID prefixes only - never a
/// subject or a body, from this store or any other.
/// </para>
/// <para>
/// <b>What makes a run conclusive.</b> Every outcome except "the probe could not run" answers
/// the question, so the assertion is only that it ran: a table opened with at least one row.
/// The verdict is printed, per store, in the summary block at the end.
/// </para>
/// <para>
/// <b>A second question rides along</b> (<c>ATableDate_IsEitherUtcOrLocal_AndTheRunSaysWhich</c>),
/// because it is answerable from the same read and one live run is what is scarce here: whether
/// an Outlook <c>Table</c> reports date-time values in UTC or in local time. That one decides
/// whether a resumed exhaustive scan's date bound lands one local offset early, which would skip
/// the mail in that window and report the scan complete.
/// </para>
/// </summary>
[Collection(LiveCollections.Phase3)]
[Trait("Category", "Live")]
public sealed class LiveTableSortProbeTests
{
    /// <summary>The explicit built-in name Microsoft's own <c>Table.Sort</c> example uses.</summary>
    private const string ExplicitName = "ReceivedTime";

    /// <summary>The namespace reference the shipped sweep passes, and the suspect.</summary>
    private const string NamespaceReference = "urn:schemas:httpmail:datereceived";

    private readonly LivePhase3Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveTableSortProbeTests(LivePhase3Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    [Fact]
    [Trait("Requires", "OutlookInstance")]
    [Trait("Writes", "Nothing")]
    public void TableSort_AcceptsAnExplicitPropertyName_OrRefusesBoth_AndTheRunSaysWhich()
    {
        IReadOnlyList<ComStoreDetail> stores = _fixture.VerifySession.GetStoreDetails();
        Assert.True(stores.Count > 0, "the profile mounts no stores, so nothing can be probed");

        StringBuilder summary = new StringBuilder();
        _ = summary.AppendLine();
        _ = summary.AppendLine("=================================================================");
        _ = summary.AppendLine(" Table.Sort probe - does the freshness sweep's sort ever apply?");
        _ = summary.AppendLine("=================================================================");

        int probed = 0;
        int explicitWorked = 0;
        int namespaceWorked = 0;
        int bothRefused = 0;

        foreach (ComStoreDetail store in stores)
        {
            ComTableSortProbe probe;
            try
            {
                probe = _fixture.VerifySession.ProbeTableSort(
                    store.DisplayName, null, new[] { ExplicitName, NamespaceReference });
            }
            catch (Exception ex)
            {
                _ = summary.AppendLine();
                _ = summary.AppendLine("STORE " + store.DisplayName + " - NOT PROBED (" + ex.GetType().Name + "): "
                    + ex.Message);
                continue;
            }

            _ = summary.AppendLine();
            _ = summary.AppendLine("STORE " + probe.StoreDisplayName + "  folder=" + probe.FolderName
                + "  rows=" + probe.RowCount.ToString(System.Globalization.CultureInfo.InvariantCulture));
            _ = summary.AppendLine("  " + Describe("(no sort asked for)", probe.Baseline));

            bool anyRow = probe.Baseline.FirstRowEntryId != null;
            ComTableSortAttempt? explicitAttempt = null;
            ComTableSortAttempt? namespaceAttempt = null;
            foreach (ComTableSortAttempt attempt in probe.Attempts)
            {
                _ = summary.AppendLine("  " + Describe(attempt.Property ?? "?", attempt));
                anyRow = anyRow || attempt.FirstRowEntryId != null;
                if (string.Equals(attempt.Property, ExplicitName, StringComparison.Ordinal))
                {
                    explicitAttempt = attempt;
                }
                else if (string.Equals(attempt.Property, NamespaceReference, StringComparison.Ordinal))
                {
                    namespaceAttempt = attempt;
                }
            }

            if (!anyRow)
            {
                _ = summary.AppendLine("  VERDICT: INCONCLUSIVE - no row came back, so no ordering can be observed.");
                continue;
            }

            probed++;
            bool explicitOk = explicitAttempt != null && explicitAttempt.SortApplied;
            bool namespaceOk = namespaceAttempt != null && namespaceAttempt.SortApplied;
            explicitWorked += explicitOk ? 1 : 0;
            namespaceWorked += namespaceOk ? 1 : 0;
            bothRefused += !explicitOk && !namespaceOk ? 1 : 0;

            _ = summary.AppendLine("  VERDICT: " + Verdict(explicitOk, namespaceOk));
        }

        _ = summary.AppendLine();
        _ = summary.AppendLine("-----------------------------------------------------------------");
        _ = summary.AppendLine(" stores probed .................. "
            + probed.ToString(System.Globalization.CultureInfo.InvariantCulture));
        _ = summary.AppendLine(" Sort(\"" + ExplicitName + "\") accepted ...... "
            + explicitWorked.ToString(System.Globalization.CultureInfo.InvariantCulture));
        _ = summary.AppendLine(" Sort(\"" + NamespaceReference + "\") accepted ... "
            + namespaceWorked.ToString(System.Globalization.CultureInfo.InvariantCulture));
        _ = summary.AppendLine(" both refused ................... "
            + bothRefused.ToString(System.Globalization.CultureInfo.InvariantCulture));
        _ = summary.AppendLine();
        _ = summary.AppendLine(" ANSWER: " + Answer(probed, explicitWorked, namespaceWorked, bothRefused));
        _ = summary.AppendLine("-----------------------------------------------------------------");

        _output.WriteLine(summary.ToString());

        Assert.True(
            probed > 0,
            "the probe could not run anywhere: no store returned a table with a readable first row, so this run "
            + "answers nothing. Check that Outlook is running and that at least one store has mail.");
    }

    /// <summary>
    /// The SECOND question one read-only run can settle: does an Outlook <c>Table</c> report
    /// its date-time values in UTC or in local time?
    /// <para>
    /// A COM-marshalled date always arrives with <c>DateTimeKind.Unspecified</c>, so nothing
    /// in the value says which. Until 2026-08-23 this solution held both answers at once: the
    /// tripwire census took an unspecified kind as already-UTC, and
    /// <c>OutlookComSession.ReadRowDate</c> called <c>ToUniversalTime</c> on it. Both now
    /// call one helper, so this run does not decide whether they AGREE - they do - it decides
    /// whether the reading they share is the right one.
    /// </para>
    /// <para>
    /// <b>Why it is not cosmetic.</b> That value becomes a resumed exhaustive scan's
    /// inclusive "at or before" bound. A bound one local offset too EARLY skips the mail
    /// received in that window and reports the scan complete, in the one search mode a caller
    /// picks because completeness matters. Too LATE only re-reads rows the chain already
    /// suppresses by EntryID, which is why the shipped helper takes the later of the two
    /// readings while this run is outstanding.
    /// </para>
    /// <para>
    /// <b>What makes it conclusive.</b> All three readings are of ONE item, found by the
    /// row's own EntryID. An earlier eyeballed two-hour gap between a sorted and an unsorted
    /// first row on a two-row store looked exactly like this bug and could equally have been
    /// two different items; nothing here can make that mistake.
    /// </para>
    /// </summary>
    [Fact]
    [Trait("Requires", "OutlookInstance")]
    [Trait("Writes", "Nothing")]
    public void ATableDate_IsEitherUtcOrLocal_AndTheRunSaysWhich()
    {
        IReadOnlyList<ComStoreDetail> stores = _fixture.VerifySession.GetStoreDetails();
        Assert.True(stores.Count > 0, "the profile mounts no stores, so nothing can be probed");

        StringBuilder summary = new StringBuilder();
        _ = summary.AppendLine();
        _ = summary.AppendLine("=================================================================");
        _ = summary.AppendLine(" Table date-kind probe - does a Table report UTC or local time?");
        _ = summary.AppendLine("=================================================================");

        int read = 0;
        int saysUtc = 0;
        int saysLocal = 0;
        int saysNeither = 0;
        int inconclusive = 0;
        int helperWrong = 0;

        foreach (ComStoreDetail store in stores)
        {
            ComTableDateKindProbe probe;
            try
            {
                probe = _fixture.VerifySession.ProbeTableDateKind(store.DisplayName);
            }
            catch (Exception ex)
            {
                _ = summary.AppendLine();
                _ = summary.AppendLine("STORE " + store.DisplayName + " - NOT PROBED (" + ex.GetType().Name + "): "
                    + ex.Message);
                continue;
            }

            _ = summary.AppendLine();
            _ = summary.AppendLine("STORE " + probe.StoreDisplayName + "  folder=" + probe.FolderName
                + "  rowsExamined=" + probe.RowsExamined.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + "  id=" + Prefix(probe.EntryId));

            if (probe.Error != null || !probe.TableRawValue.HasValue || !probe.ItemRawValue.HasValue)
            {
                _ = summary.AppendLine("  NO READING: " + (probe.Error ?? "the probe returned no pair of values."));
                continue;
            }

            read++;
            _ = summary.AppendLine("  table raw ................ " + FormatDate(probe.TableRawValue)
                + "  kind=" + probe.TableRawKind);
            _ = summary.AppendLine("  table via ReadRowDate .... " + FormatDate(probe.TableThroughSharedHelper));
            _ = summary.AppendLine("  item  MailItem.ReceivedTime " + FormatDate(probe.ItemRawValue)
                + "  kind=" + probe.ItemRawKind);
            _ = summary.AppendLine("  item  via shared helper .. " + FormatDate(probe.ItemThroughSharedHelper));
            _ = summary.AppendLine("  machine UTC offset ....... "
                + probe.LocalOffsetMinutes.ToString(System.Globalization.CultureInfo.InvariantCulture) + " min");

            if (probe.LocalOffsetMinutes == 0)
            {
                inconclusive++;
                _ = summary.AppendLine("  VERDICT: INCONCLUSIVE - this machine is at UTC, so the two readings are the "
                    + "same number and neither is tested. Re-run somewhere with a non-zero offset.");
                continue;
            }

            // The object model is documented to return LOCAL wall time, so the item's own
            // value is the reference and the RAW table value is what is being classified - and,
            // separately, whether the shipped helper turns that raw value into the item's UTC
            // instant. The two used to be one test: "the helper agrees" stood in for "the table
            // is UTC", which was only ever true while the helper relabelled. Since Q11 was settled
            // (2026-10-03: the explicit name this probe adds reports LOCAL) the helper converts by
            // spelling, so a LOCAL table and an agreeing helper are now the expected pair.
            bool tableIsLocal = probe.TableRawValue!.Value == probe.ItemRawValue!.Value;
            bool tableIsUtc = probe.ItemThroughSharedHelper.HasValue
                && probe.TableRawValue.Value == probe.ItemThroughSharedHelper.Value;
            bool helperAgrees = probe.TableThroughSharedHelper.HasValue
                && probe.ItemThroughSharedHelper.HasValue
                && probe.TableThroughSharedHelper.Value == probe.ItemThroughSharedHelper.Value;
            if (!helperAgrees)
            {
                helperWrong++;
            }

            string helperLine = helperAgrees
                ? " The shipped helper reads it as the item's own UTC instant (ComDateValue.FromTableValue by column spelling)."
                : " The shipped helper does NOT read it as the item's own UTC instant - it is wrong by the difference above.";
            if (tableIsLocal && !tableIsUtc)
            {
                saysLocal++;
                _ = summary.AppendLine("  VERDICT: TABLE REPORTS LOCAL TIME on this store for the column this probe added. "
                    + "The raw table value equals the opened item's own ReceivedTime, which the object model documents "
                    + "as local." + helperLine);
            }
            else if (tableIsUtc && !tableIsLocal)
            {
                saysUtc++;
                _ = summary.AppendLine("  VERDICT: TABLE REPORTS UTC on this store for the column this probe added. "
                    + "The raw table value equals the item's ReceivedTime converted to UTC." + helperLine);
            }
            else
            {
                saysNeither++;
                _ = summary.AppendLine("  VERDICT: NEITHER - the two values differ by something other than this "
                    + "machine's UTC offset, so the table is not simply one zone or the other. Read the four numbers "
                    + "above before changing anything: a daylight-saving boundary, a provider that rounds, or an item "
                    + "whose delivery time was rewritten are all live possibilities." + helperLine);
            }
        }

        _ = summary.AppendLine();
        _ = summary.AppendLine("-----------------------------------------------------------------");
        _ = summary.AppendLine(" stores read .................... "
            + read.ToString(System.Globalization.CultureInfo.InvariantCulture));
        _ = summary.AppendLine(" table reports UTC .............. "
            + saysUtc.ToString(System.Globalization.CultureInfo.InvariantCulture));
        _ = summary.AppendLine(" table reports LOCAL ............ "
            + saysLocal.ToString(System.Globalization.CultureInfo.InvariantCulture));
        _ = summary.AppendLine(" neither ........................ "
            + saysNeither.ToString(System.Globalization.CultureInfo.InvariantCulture));
        _ = summary.AppendLine(" inconclusive (machine at UTC) .. "
            + inconclusive.ToString(System.Globalization.CultureInfo.InvariantCulture));
        _ = summary.AppendLine(" helper reads it wrongly ........ "
            + helperWrong.ToString(System.Globalization.CultureInfo.InvariantCulture));
        _ = summary.AppendLine();
        _ = summary.AppendLine(" ANSWER: " + DateKindAnswer(read, saysUtc, saysLocal, saysNeither, inconclusive, helperWrong));
        _ = summary.AppendLine("-----------------------------------------------------------------");

        _output.WriteLine(summary.ToString());

        Assert.True(
            read > 0,
            "the date-kind probe could not read a single item two ways, so this run answers nothing. Check that "
            + "Outlook is running and that at least one store has a dated item in its Inbox.");

        // Since Q11 was settled (2026-10-03) the question has an answer the product relies on: the
        // shipped helper must turn the table's value into the item's own UTC instant on every
        // store that could be read with a non-zero offset. A wrong reading puts a resumed
        // exhaustive scan's bound one offset off - a duplicate east of UTC, SKIPPED mail west of it.
        int conclusive = read - inconclusive;
        Assert.True(
            conclusive == 0 || helperWrong == 0,
            helperWrong + " of " + conclusive + " store(s) read the table date wrongly through the shipped helper - "
            + "see the per-store verdicts above");
    }

    private static string DateKindAnswer(int read, int saysUtc, int saysLocal, int saysNeither, int inconclusive, int helperWrong)
    {
        if (read == 0)
        {
            return "none - the probe did not read anything.";
        }

        if (inconclusive == read)
        {
            return "none - this machine runs at UTC, where both readings produce the same number. The question can "
                + "only be settled where the offset is non-zero.";
        }

        string helper = helperWrong == 0
            ? " The shipped helper (ComDateValue.FromTableValue, by column spelling) reads every one correctly."
            : " The shipped helper reads " + helperWrong.ToString(System.Globalization.CultureInfo.InvariantCulture)
                + " of them WRONGLY - a resumed exhaustive scan's bound is one offset off there.";
        if (saysUtc > 0 && saysLocal == 0 && saysNeither == 0)
        {
            return "a Table reports UTC for the column this probe added (the explicit name ReceivedTime) - which "
                + "contradicts the 2026-10-03 measurement that settled QUESTIONS.md Q11 the other way; find out "
                + "what differs on this machine." + helper;
        }

        if (saysLocal > 0 && saysUtc == 0 && saysNeither == 0)
        {
            return "a Table reports LOCAL time for the explicit name ReceivedTime, as measured when QUESTIONS.md Q11 "
                + "was settled (2026-10-03)." + helper;
        }

        return "MIXED or unexplained - read the per-store verdicts above. A table's zone is a property of the "
            + "provider and the column spelling, so a per-store difference means something other than the zone is "
            + "in play." + helper;
    }

    private static string Describe(string label, ComTableSortAttempt attempt)
    {
        StringBuilder line = new StringBuilder();
        _ = line.Append(label.PadRight(38)).Append(" | ");
        _ = line.Append("column=").Append(attempt.Property == null ? "n/a  " : attempt.ColumnAdded ? "ok   " : "FAILED");
        _ = line.Append(" | sort=")
            .Append(attempt.Property == null ? "n/a    " : attempt.SortApplied ? "APPLIED" : "REFUSED");
        _ = line.Append(" | firstRow=").Append(FormatDate(attempt.FirstRowReceivedUtc));
        _ = line.Append(" id=").Append(Prefix(attempt.FirstRowEntryId));
        if (attempt.ColumnError != null)
        {
            _ = line.Append(" | columnError=").Append(attempt.ColumnError);
        }

        if (attempt.SortError != null)
        {
            _ = line.Append(" | sortError=").Append(attempt.SortError);
        }

        return line.ToString();
    }

    private static string Verdict(bool explicitOk, bool namespaceOk)
    {
        if (explicitOk && !namespaceOk)
        {
            return "HYPOTHESIS CONFIRMED on this store - the explicit name sorts and the namespace reference does "
                + "not, exactly as Table.Sort's documentation says. The shipped sweep passes the namespace form, so "
                + "its sort has never applied here and its 200-item cap has been cutting arbitrarily.";
        }

        if (explicitOk && namespaceOk)
        {
            return "HYPOTHESIS REFUTED on this store - BOTH forms sort, so the namespace reference is not what has "
                + "been stopping the sweep's sort. Look for a per-store or per-folder cause instead.";
        }

        if (!explicitOk && namespaceOk)
        {
            return "UNEXPECTED on this store - the namespace form sorts and the explicit name does not. The sweep's "
                + "call is fine here; something about the explicit name is not.";
        }

        return "SORT REFUSED ENTIRELY on this store - neither spelling is accepted, so the property name is not the "
            + "cause and the table simply will not order by received date. The sweep's cap cuts arbitrarily here "
            + "whatever it passes, and the resumable scan's date rung is unavailable on this store.";
    }

    private static string Answer(int probed, int explicitWorked, int namespaceWorked, int bothRefused)
    {
        if (probed == 0)
        {
            return "none - the probe did not run.";
        }

        if (explicitWorked == probed && namespaceWorked == 0)
        {
            return "the namespace reference is the whole story. Change the sweep's Sort call to the explicit "
                + "property name; the resumable scan's date rung becomes the normal path rather than the lucky one, "
                + "and gap H2's advice sentence becomes true for the first time.";
        }

        if (bothRefused == probed)
        {
            return "the property name is NOT the cause - no spelling sorts anywhere on this profile. The sweep's "
                + "cap really does cut arbitrarily, item_cap_unsorted is correct as it stands, and the resumable "
                + "scan will live on its ordinal and restart rungs here.";
        }

        if (namespaceWorked == probed)
        {
            return "the shipped call works on this profile. Whatever has made sortApplied false must be looked for "
                + "elsewhere - a folder without the column, or a provider that refuses per folder rather than per "
                + "store.";
        }

        return "MIXED across stores - read the per-store verdicts above. A mixed result rules the documentation "
            + "restriction OUT as a complete explanation, because it would apply identically everywhere.";
    }

    private static string FormatDate(DateTime? value)
    {
        return value.HasValue
            ? value.Value.ToString("yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture)
            : "(none)             ";
    }

    /// <summary>
    /// The first eight characters of an EntryID - enough to see whether the first row CHANGED
    /// between attempts, which is the only thing the id is here for, and short enough that no
    /// other store's item is identified in a log.
    /// </summary>
    private static string Prefix(string? entryId)
    {
        if (string.IsNullOrEmpty(entryId))
        {
            return "--------";
        }

        return entryId!.Length <= 8 ? entryId : entryId.Substring(0, 8);
    }
}
