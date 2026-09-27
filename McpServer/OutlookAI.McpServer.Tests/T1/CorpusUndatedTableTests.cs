using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the undated probe's two additions after OAI-UNINDEXED, 2026-09-27: the SECOND read of an
/// undated item's received date - the folder's own table, restricted by the store - and the report of
/// a delivery time Outlook refused to remove.
/// <para>
/// <b>Why.</b> That day the PropertyAccessor read a delivery time on every appointment, contact and task
/// saved into a PST, and refused to delete it for an appointment (<c>UnauthorizedAccessException</c>,
/// "does not support this operation") - a refusal that escaped the probe and ended the run. Whether the
/// store holds the value, or only the object model reports one, is the question the maintainer's
/// decision turns on; the table answers it without writing, and the refusal is now a column rather than
/// a crash.
/// </para>
/// </summary>
public sealed class CorpusUndatedTableTests
{
    private const string Id = "hub-unindexed";

    [Fact]
    public void TheUndatedFilter_SelectsTheProbeItem_WhereTheReceivedDateIsNull()
    {
        string fragment = CorpusPlan.DaslSubjectFragment(Id, CorpusPlan.ProbeOrdinal);
        Assert.Equal(
            "@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%" + fragment + "%') AND (\"urn:schemas:httpmail:datereceived\" IS NULL)",
            CorpusUndatedTable.Filter(Id, withReceivedDate: false));
    }

    [Fact]
    public void TheDatedFilter_IsTheSameRestriction_Negated()
    {
        string fragment = CorpusPlan.DaslSubjectFragment(Id, CorpusPlan.ProbeOrdinal);
        Assert.Equal(
            "@SQL=(\"urn:schemas:httpmail:subject\" LIKE '%" + fragment + "%') AND (NOT (\"urn:schemas:httpmail:datereceived\" IS NULL))",
            CorpusUndatedTable.Filter(Id, withReceivedDate: true));
    }

    [Fact]
    public void TheFilters_NameTheProbeOrdinal_NotTheWholeCorpus_TheControl()
    {
        // The control: a filter that selected the whole corpus is what once made a populated folder look
        // like a placement failure (the walk hit its row cap first). Both name the ONE probe ordinal.
        foreach (bool dated in new[] { false, true })
        {
            string filter = CorpusUndatedTable.Filter(Id, dated);
            Assert.Contains(CorpusPlan.DaslSubjectFragment(Id, CorpusPlan.ProbeOrdinal), filter, StringComparison.Ordinal);
            Assert.DoesNotContain("[", filter, StringComparison.Ordinal);
        }

        Assert.Throws<ArgumentException>(() => CorpusUndatedTable.Filter(" ", false));
    }

    [Theory]
    [InlineData(true, false, true)]
    [InlineData(false, true, false)]
    [InlineData(true, true, null)]
    [InlineData(false, false, null)]
    [InlineData(null, false, null)]
    [InlineData(true, null, null)]
    [InlineData(null, null, null)]
    public void TheVerdict_AnswersOnlyWhenTheTwoLookupsAgree(bool? foundWhereUndated, bool? foundWhereDated, bool? expected)
        => Assert.Equal(expected, CorpusUndatedTable.Verdict(foundWhereUndated, foundWhereDated));

    [Fact]
    public void TheLine_CarriesTheTableAnswer_AndARefusedRemovalOnlyWhenThereIsOne()
    {
        var measured = new CorpusUndatedProbe(
            CorpusItemKind.Appointment, true, true, true, false, true, true, null,
            "UnauthorizedAccessException: does not support this operation.", false);
        string line = CorpusUndatedFidelity.Line(measured);
        Assert.Contains("appointment", line, StringComparison.Ordinal);
        Assert.Contains(" undated=False tableUndated=False ", line, StringComparison.Ordinal);
        Assert.Contains(" removalRefused=UnauthorizedAccessException: does not support this operation.", line, StringComparison.Ordinal);
        Assert.DoesNotContain("error=", line, StringComparison.Ordinal);

        // The control: an item with nothing refused and no table answer says so, and prints no refusal.
        string plain = CorpusUndatedFidelity.Line(measured with { DeliveryTimeRemovalRefused = null, StoreTableSaysUndated = null });
        Assert.Contains(" tableUndated=(unknown) ", plain, StringComparison.Ordinal);
        Assert.DoesNotContain("removalRefused", plain, StringComparison.Ordinal);
    }

    [Fact]
    public void ADatedItem_WhoseRemovalWasRefused_IsRefused_NamingTheRefusalAndTheTable()
    {
        IReadOnlyList<CorpusItemKind> kinds = new[] { CorpusItemKind.Appointment };
        var probe = new CorpusUndatedProbe(
            CorpusItemKind.Appointment, true, true, true, false, true, true, null, "UnauthorizedAccessException: no", false);
        (bool proceed, string message) = CorpusUndatedFidelity.Decide(kinds, new[] { probe });
        Assert.False(proceed);
        Assert.Contains("CARRIES a delivery time", message, StringComparison.Ordinal);
        Assert.Contains("(the folder's own table agrees)", message, StringComparison.Ordinal);
        Assert.Contains("Outlook refused to remove it (UnauthorizedAccessException: no)", message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnItemThatReadsUndated_ButTheTableDates_IsRefused()
    {
        IReadOnlyList<CorpusItemKind> kinds = new[] { CorpusItemKind.Task };
        var contradicted = new CorpusUndatedProbe(CorpusItemKind.Task, true, true, true, true, true, true, null, null, false);
        (bool proceed, string message) = CorpusUndatedFidelity.Decide(kinds, new[] { contradicted });
        Assert.False(proceed);
        Assert.Contains("TASK", message, StringComparison.Ordinal);
        Assert.Contains("the folder's own table finds it DATED", message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnUndatedItem_TheTableAgreesWith_OrCouldNotAnswerFor_Proceeds_TheControl()
    {
        // The control: the second read refuses only a CONTRADICTION. Agreement passes, and so does a table
        // that could not answer - the PropertyAccessor's read stays the criterion it always was.
        IReadOnlyList<CorpusItemKind> kinds = new[] { CorpusItemKind.Contact };
        foreach (bool? table in new bool?[] { true, null })
        {
            var probe = new CorpusUndatedProbe(CorpusItemKind.Contact, true, true, true, true, true, true, null, null, table);
            (bool proceed, string message) = CorpusUndatedFidelity.Decide(kinds, new[] { probe });
            Assert.True(proceed, message);
        }
    }
}
