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
/// </summary>
[Collection("LivePhase3")]
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
