using System.Globalization;
using OutlookAI.Core.IndexSearch;
using OutlookAI.RemediationTools;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins <see cref="CorpusUndatedIndex"/> - the read-only observer behind
/// <c>corpus-probe --undated-index-wait N</c>, added 2026-10-03 for Q98 (f): what the Windows Search
/// index makes of an appointment, a contact and a task the undated probe holds in a PST. The COM half
/// (create, hold, delete) is the undated probe's own; everything this class decides is here, with a
/// fake index client and an injected clock, so nothing waits and nothing touches an index.
/// </summary>
public sealed class CorpusUndatedIndexTests
{
    private const string Tag = "[OutlookAI-Corpus:q98scratch#2147483647]";
    private const string AppointmentSubject = "[OutlookAI-Corpus]" + Tag + " undated appointment";
    private const string ContactSubject = "[OutlookAI-Corpus]" + Tag + " undated contact";

    private static Dictionary<string, object?> Row(string url, string? subject = null, string? name = null, string? fullName = null)
        => new(StringComparer.OrdinalIgnoreCase)
        {
            ["System.ItemUrl"] = url,
            ["System.Subject"] = subject,
            ["System.ItemNameDisplay"] = name,
            ["System.Contact.FullName"] = fullName,
        };

    private const string MapiUrl = "mapi16://{S-1-5-21-1-2-3-1000}/q98scratch@vm.invalid($65d10200)/0/Calendar/\uAC00\uAC01";

    [Fact]
    public void TheColumns_StartWithTheUrl_AndCarryBothOrderByKeysOfTheProduct()
    {
        Assert.Equal("System.ItemUrl", CorpusUndatedIndex.Columns[0]);
        Assert.Contains("System.Message.DateReceived", CorpusUndatedIndex.Columns);
        Assert.Contains("System.Size", CorpusUndatedIndex.Columns);
        Assert.Equal(CorpusUndatedIndex.Columns.Count, CorpusUndatedIndex.Columns.Distinct(StringComparer.Ordinal).Count());
        Assert.All(CorpusUndatedIndex.Statements(), sql =>
        {
            Assert.StartsWith("SELECT TOP 100 System.ItemUrl, ", sql, StringComparison.Ordinal);
            Assert.Contains("'\"undated\"'", sql, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void TheProbeTag_IsTheBracketedCorpusIdAndOrdinal_AndAnUntaggedSubjectHasNone()
    {
        Assert.Equal(Tag, CorpusUndatedIndex.ProbeTag(AppointmentSubject));
        Assert.Null(CorpusUndatedIndex.ProbeTag("[OutlookAI-Corpus] undated appointment"));
        Assert.Null(CorpusUndatedIndex.ProbeTag(null));
        Assert.Null(CorpusUndatedIndex.ProbeTag("[OutlookAI-Corpus:unclosed"));
    }

    [Fact]
    public void ARow_MatchesOnlyItsOwnKindAndProbeTag_AndOnlyAnOutlookUrl()
    {
        Assert.True(CorpusUndatedIndex.IsRowFor(Row(MapiUrl, subject: AppointmentSubject), AppointmentSubject, CorpusItemKind.Appointment));

        // Another kind's row - an earlier probe item the index has not dropped yet - never matches.
        Assert.False(CorpusUndatedIndex.IsRowFor(Row(MapiUrl, subject: AppointmentSubject), AppointmentSubject, CorpusItemKind.Task));

        // Another corpus's probe item never matches.
        string other = AppointmentSubject.Replace("q98scratch", "someoneelse", StringComparison.Ordinal);
        Assert.False(CorpusUndatedIndex.IsRowFor(Row(MapiUrl, subject: other), AppointmentSubject, CorpusItemKind.Appointment));

        // A file-system row carrying the same text is not an Outlook item.
        Assert.False(CorpusUndatedIndex.IsRowFor(Row("file:///C:/notes.txt", subject: AppointmentSubject), AppointmentSubject, CorpusItemKind.Appointment));
        Assert.False(CorpusUndatedIndex.IsRowFor(Row(MapiUrl, subject: AppointmentSubject), "untagged subject", CorpusItemKind.Appointment));
    }

    [Fact]
    public void AContact_MatchesOnItsReorderedName_WhenTheIndexHasNoSubjectForIt()
    {
        // File-as order: the words move, the tag - which has no space in it - survives whole.
        string fileAs = "contact, [OutlookAI-Corpus]" + Tag + " undated";
        Assert.True(CorpusUndatedIndex.IsRowFor(Row(MapiUrl, subject: null, name: fileAs), ContactSubject, CorpusItemKind.Contact));
        Assert.True(CorpusUndatedIndex.IsRowFor(Row(MapiUrl, fullName: ContactSubject), ContactSubject, CorpusItemKind.Contact));
    }

    [Fact]
    public void ADescription_SaysNullForAMissingColumn_UsesInvariantDates_AndHidesTheEncodedId()
    {
        Dictionary<string, object?> row = Row(MapiUrl, subject: AppointmentSubject);
        row["System.Message.DateReceived"] = new DateTime(2026, 10, 3, 1, 2, 3);
        row["System.Size"] = 4096L;
        row["System.Kind"] = new[] { "calendar", "communication" };
        row["System.DueDate"] = DBNull.Value;

        string text = CorpusUndatedIndex.Describe(CorpusItemKind.Appointment, row, TimeSpan.FromSeconds(12.7));

        Assert.StartsWith("  appointment: in the index after 12 s", text, StringComparison.Ordinal);
        Assert.Contains("System.Message.DateReceived = 2026-10-03 01:02:03", text, StringComparison.Ordinal);
        Assert.Contains("System.Size = 4096", text, StringComparison.Ordinal);
        Assert.Contains("System.Kind = calendar|communication", text, StringComparison.Ordinal);
        Assert.Contains("System.DueDate = <null>", text, StringComparison.Ordinal);
        Assert.Contains("System.StartDate = <null>", text, StringComparison.Ordinal);
        Assert.Contains("/0/Calendar/<id>", text, StringComparison.Ordinal);
        Assert.DoesNotContain("\uAC00", text, StringComparison.Ordinal);
        Assert.Equal(CorpusUndatedIndex.Columns.Count + 1, text.Split('\n').Length);
    }

    [Fact]
    public void TheReader_AsksAgainUntilTheRowArrives_AndWritesIt()
    {
        var client = new ScriptedClient(
            _ => Array.Empty<IReadOnlyDictionary<string, object?>>(),
            _ => Array.Empty<IReadOnlyDictionary<string, object?>>(),
            _ => Array.Empty<IReadOnlyDictionary<string, object?>>(),
            _ => new IReadOnlyDictionary<string, object?>[] { Row(MapiUrl, subject: AppointmentSubject) });
        var clock = new FakeClock();
        using var output = new StringWriter(CultureInfo.InvariantCulture);

        CorpusUndatedIndex.CreateReader(output, TimeSpan.FromSeconds(60), client, clock.Start, clock.Sleep)(CorpusItemKind.Appointment, AppointmentSubject);

        string text = output.ToString();
        Assert.Contains("appointment: in the index after 3 s", text, StringComparison.Ordinal);
        Assert.Equal(4, client.Calls);
        Assert.Equal(new[] { CorpusUndatedIndex.PollInterval }, clock.Slept);
    }

    [Fact]
    public void TheReader_GivenAWrittenInstant_WaitsPastTheFirstSaveUntilTheIndexDatesItAsWritten()
    {
        // D62 (b): the index takes an item's first save first - dated at its creation - and the probe's written
        // delivery time only when it re-reads the item. So a written instant is waited for, not just a row.
        DateTime written = new(2026, 8, 2, 9, 32, 50, DateTimeKind.Utc);
        DateTime created = new(2026, 10, 3, 16, 40, 0, DateTimeKind.Utc);
        Dictionary<string, object?> Dated(DateTime at)
        {
            Dictionary<string, object?> row = Row(MapiUrl, subject: AppointmentSubject);
            row["System.Message.DateReceived"] = at;
            return row;
        }

        var client = new ScriptedClient(
            _ => new IReadOnlyDictionary<string, object?>[] { Dated(created) },
            _ => new IReadOnlyDictionary<string, object?>[] { Dated(created) },
            _ => new IReadOnlyDictionary<string, object?>[] { Dated(written) });
        var clock = new FakeClock();
        using var output = new StringWriter(CultureInfo.InvariantCulture);

        CorpusUndatedIndex.CreateReader(output, TimeSpan.FromSeconds(60), client, clock.Start, clock.Sleep, _ => written)(CorpusItemKind.Appointment, AppointmentSubject);

        string text = output.ToString();
        Assert.Equal(1, CountOf(text, ": in the index after"));
        Assert.Contains("appointment: DATED AS WRITTEN after 6 s - System.Message.DateReceived 2026-08-02T09:32:50Z", text, StringComparison.Ordinal);

        // And one the index keeps at its creation is said to be so when the wait runs out.
        var stuck = new ScriptedClient(_ => new IReadOnlyDictionary<string, object?>[] { Dated(created) });
        using var stuckOutput = new StringWriter(CultureInfo.InvariantCulture);
        var stuckClock = new FakeClock();
        CorpusUndatedIndex.CreateReader(stuckOutput, TimeSpan.FromSeconds(9), stuck, stuckClock.Start, stuckClock.Sleep, _ => written)(CorpusItemKind.Appointment, AppointmentSubject);
        Assert.Contains("NOT DATED AS WRITTEN after 9 s - the probe wrote 2026-08-02T09:32:50Z and the index still says 2026-10-03T16:40:00Z", stuckOutput.ToString(), StringComparison.Ordinal);
    }

    [Fact]
    public void TheReader_GivesUpAtItsWait_AndSaysWhatToCheck()
    {
        var client = new ScriptedClient(_ => Array.Empty<IReadOnlyDictionary<string, object?>>());
        var clock = new FakeClock();
        using var output = new StringWriter(CultureInfo.InvariantCulture);

        CorpusUndatedIndex.CreateReader(output, TimeSpan.FromSeconds(10), client, clock.Start, clock.Sleep)(CorpusItemKind.Task, "[OutlookAI-Corpus]" + Tag + " undated task");

        string text = output.ToString();
        Assert.Contains("task: NOT in the index after 10 s", text, StringComparison.Ordinal);
        Assert.Contains("NOT elevated Outlook", text, StringComparison.Ordinal);
        Assert.True(clock.Now >= TimeSpan.FromSeconds(10));
        Assert.All(clock.Slept, s => Assert.True(s <= CorpusUndatedIndex.PollInterval));
    }

    [Fact]
    public void TheReader_ReportsARefusedStatementOnce_UsesTheOthers_AndNeverThrows()
    {
        var client = new ScriptedClient(sql => sql.Contains("System.Subject,", StringComparison.Ordinal) && sql.Contains("CONTAINS(System.Subject", StringComparison.Ordinal)
            ? throw new InvalidOperationException("refused")
            : sql.Contains("CONTAINS(System.ItemNameDisplay", StringComparison.Ordinal)
                ? new IReadOnlyDictionary<string, object?>[] { Row(MapiUrl, name: ContactSubject) }
                : Array.Empty<IReadOnlyDictionary<string, object?>>());
        var clock = new FakeClock();
        using var output = new StringWriter(CultureInfo.InvariantCulture);
        Action<CorpusItemKind, string> reader = CorpusUndatedIndex.CreateReader(output, TimeSpan.FromSeconds(10), client, clock.Start, clock.Sleep);

        reader(CorpusItemKind.Contact, ContactSubject);
        reader(CorpusItemKind.Contact, ContactSubject);

        string text = output.ToString();
        Assert.Equal(1, CountOf(text, "index refused a statement"));
        Assert.Equal(2, CountOf(text, "contact: in the index after"));

        var broken = new ScriptedClient(_ => throw new OutOfMemoryException("not this one"));
        Assert.Throws<OutOfMemoryException>(() =>
            CorpusUndatedIndex.CreateReader(output, TimeSpan.FromSeconds(1), broken, clock.Start, clock.Sleep)(CorpusItemKind.Task, "x"));
    }

    [Fact]
    public void TheOption_ParsesAsSeconds_AndANegativeWaitHoldsNothing()
    {
        Assert.Equal(0, CorpusOptions.Parse(Array.Empty<string>()).UndatedIndexWaitSeconds);
        Assert.Equal(90, CorpusOptions.Parse(new[] { "--undated-index-wait", "90" }).UndatedIndexWaitSeconds);
        Assert.Equal(0, CorpusOptions.Parse(new[] { "--undated-index-wait", "-5" }).UndatedIndexWaitSeconds);
    }

    private static int CountOf(string text, string what)
    {
        int count = 0;
        for (int at = text.IndexOf(what, StringComparison.Ordinal); at >= 0; at = text.IndexOf(what, at + what.Length, StringComparison.Ordinal))
        {
            count++;
        }

        return count;
    }

    /// <summary>Answers the n-th statement with the n-th script; the last script repeats.</summary>
    private sealed class ScriptedClient : IIndexClient
    {
        private readonly Func<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>[] _scripts;

        public ScriptedClient(params Func<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>>[] scripts) => _scripts = scripts;

        public int Calls { get; private set; }

        public IndexProviderKind Provider => IndexProviderKind.OleDb;

        public IReadOnlyList<IReadOnlyDictionary<string, object?>> ExecuteRows(string sql, int maxRows, int? commandTimeoutSeconds = null)
        {
            Func<string, IReadOnlyList<IReadOnlyDictionary<string, object?>>> script = _scripts[Math.Min(Calls, _scripts.Length - 1)];
            Calls++;
            return script(sql);
        }
    }

    /// <summary>A clock that only moves when the reader sleeps.</summary>
    private sealed class FakeClock
    {
        public TimeSpan Now { get; private set; }

        public List<TimeSpan> Slept { get; } = new();

        public Func<TimeSpan> Start()
        {
            TimeSpan origin = Now;
            return () => Now - origin;
        }

        public void Sleep(TimeSpan span)
        {
            Slept.Add(span);
            Now += span;
        }
    }
}
