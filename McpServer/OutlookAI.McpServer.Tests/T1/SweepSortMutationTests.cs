using System.Reflection;

using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The holes a mutation pass through the freshness-sweep sort fix (commit <c>bea7fc9</c>)
/// found in the suite that shipped with it.
/// <para>
/// <see cref="SweepSortPropertyTests"/> pins the LADDER - which spellings, in which order -
/// and it kills every mutation of the two arrays that matters. What no test reached was the
/// code that WALKS the ladder: <c>TryOrderSweptTable</c> could be truncated to its first
/// rung, made to sort a column it had just failed to add, made to report a column it never
/// added, or made to claim a sort it never applied, and 1,936 tests still passed. That
/// method is the whole fix - a ladder nothing walks correctly is the shipped defect with
/// more spellings in it - so it is exercised here against a stand-in table.
/// </para>
/// <para>
/// TIER 1: pure. The stand-in is an ordinary object; no COM, no Outlook, no mailbox. It works
/// because <c>TryOrderSweptTable</c> reaches its table through <c>dynamic</c>, which binds
/// against the runtime type, and because <c>Release</c> is guarded by
/// <c>Marshal.IsComObject</c> and so no-ops on a plain object. The refusals are thrown as
/// <see cref="ArgumentException"/> because that is the shape a real refusal arrived in -
/// late-bound COM maps <c>E_INVALIDARG</c> to it, measured on 5 of 5 stores - and it is one
/// of the types <c>OutlookComSession.IsComCallFailure</c> admits.
/// </para>
/// </summary>
public sealed class SweepSortMutationTests
{
    private const string Arrival = "ReceivedTime";

    private const string ArrivalNamespace = "urn:schemas:httpmail:datereceived";

    private const string Submit = "SentOn";

    private const string SubmitNamespace = "urn:schemas:httpmail:date";

    // ================================================================== the ladder itself

    /// <summary>
    /// The namespace spelling is kept as the LAST rung of both ladders rather than deleted,
    /// which is a decision the commit states and argues ("a store that accepts them is not
    /// impossible and falling back costs nothing") and which nothing pinned: truncating
    /// either array to its first element left all 1,936 tests passing. The refusal is
    /// measured on ONE profile and ONE Outlook build, so deleting the fallback would leave a
    /// provider that accepts only the namespace form with no sort at all - silently, because
    /// the sweep still returns mail, just an arbitrary 200 of it.
    /// </summary>
    [Theory]
    [InlineData(null, ArrivalNamespace)]
    [InlineData("inbox", ArrivalNamespace)]
    [InlineData("sent", ArrivalNamespace)]
    public void EachLadder_KeepsANamespaceSpellingAsItsLastResort(string? folderKind, string expectedLast)
    {
        IReadOnlyList<string> ladder = OutlookComSession.SweepSortProperties(folderKind);

        Assert.True(
            ladder.Count >= 2,
            "a ladder with one rung cannot fall back at all, so a store that refuses that rung gets no sort and "
            + "the 200-item cap cuts arbitrarily - which is the defect this commit fixed. Kind '"
            + (folderKind ?? "(null)") + "' has: " + string.Join(", ", ladder));
        Assert.Equal(expectedLast, ladder[ladder.Count - 1]);
    }

    // ================================================================== walking the ladder

    /// <summary>The first spelling that both goes on the table and orders it wins, and nothing below it is tried.</summary>
    [Fact]
    public void TheFirstAcceptedSpelling_EndsTheLadder()
    {
        StandInTable table = new StandInTable();

        (bool columnAdded, bool sortApplied) = Order(table, Arrival, ArrivalNamespace);

        Assert.True(columnAdded);
        Assert.True(sortApplied);
        Assert.Equal(new[] { Arrival }, table.ColumnsAdded);
        Assert.Equal(new[] { Arrival }, table.Sorted);
    }

    /// <summary>
    /// A spelling the table will not CARRY is never handed to <c>Table.Sort</c> - there is
    /// no column to order by, so the call could only produce a second failure and a false
    /// refusal count - and the next spelling is tried instead of the ladder ending.
    /// </summary>
    [Fact]
    public void ASpellingTheTableWillNotCarry_IsNeverSorted_AndTheNextRungIsTried()
    {
        StandInTable table = new StandInTable(refuseColumn: new[] { Arrival });

        (bool columnAdded, bool sortApplied) = Order(table, Arrival, ArrivalNamespace);

        Assert.True(columnAdded);
        Assert.True(sortApplied);
        Assert.Equal(new[] { Arrival, ArrivalNamespace }, table.ColumnsAdded);
        Assert.Equal(new[] { ArrivalNamespace }, table.Sorted);
    }

    /// <summary>
    /// A table that carries NO spelling reports exactly that, and asks for no sort at all.
    /// This is the half of the split the commit exists to keep separate: it must not be
    /// counted into <c>sweep.sortRefusedFolders</c>, which is the field a healthy profile is
    /// supposed to read zero on.
    /// </summary>
    [Fact]
    public void ATableThatCarriesNoSpelling_ReportsNoColumnAndAsksForNoSort()
    {
        StandInTable table = new StandInTable(refuseColumn: new[] { Arrival, ArrivalNamespace });

        (bool columnAdded, bool sortApplied) = Order(table, Arrival, ArrivalNamespace);

        Assert.False(columnAdded);
        Assert.False(sortApplied);
        Assert.Equal(new[] { Arrival, ArrivalNamespace }, table.ColumnsAdded);
        Assert.Empty(table.Sorted);
    }

    /// <summary>
    /// A REFUSED sort falls through to the next spelling rather than ending the ladder. This
    /// is the rung the whole fix is built on: the profile that shipped broken refuses the
    /// namespace form and accepts the explicit one, so a ladder that stops at its first
    /// refusal is the defect in a new spelling.
    /// </summary>
    [Fact]
    public void ARefusedSort_FallsThroughToTheNextSpelling()
    {
        StandInTable table = new StandInTable(refuseSort: new[] { Arrival });

        (bool columnAdded, bool sortApplied) = Order(table, Arrival, ArrivalNamespace);

        Assert.True(columnAdded);
        Assert.True(sortApplied);
        Assert.Equal(new[] { Arrival, ArrivalNamespace }, table.Sorted);
    }

    /// <summary>
    /// The other half of the split: every spelling went ON the table and Outlook refused to
    /// order by all of them. That is a refusal, it is reported rather than swallowed, and it
    /// is the only thing <c>sweep.sortRefusedFolders</c> counts.
    /// </summary>
    [Fact]
    public void ASortRefusedOnEverySpelling_ReportsTheColumnWentOnAndTheSortDidNot()
    {
        StandInTable table = new StandInTable(refuseSort: new[] { Arrival, ArrivalNamespace });

        (bool columnAdded, bool sortApplied) = Order(table, Arrival, ArrivalNamespace);

        Assert.True(columnAdded);
        Assert.False(sortApplied);
        Assert.Equal(new[] { Arrival, ArrivalNamespace }, table.Sorted);
    }

    /// <summary>
    /// The SENT ladder is four rungs deep and every one of them is offered, in order, before
    /// the folder is given up on. Walked against the shipped array rather than a copy of it,
    /// so shortening the array shortens this test's expectation too and only the WALK is
    /// under test here.
    /// </summary>
    [Fact]
    public void TheSentLadder_IsWalkedToItsLastRung()
    {
        IReadOnlyList<string> ladder = OutlookComSession.SweepSortProperties("sent");
        StandInTable table = new StandInTable(refuseSort: ladder);

        (bool columnAdded, bool sortApplied) = Order(table, ladder);

        Assert.True(columnAdded);
        Assert.False(sortApplied);
        Assert.Equal(new[] { Submit, Arrival, SubmitNamespace, ArrivalNamespace }, table.Sorted);
    }

    // ================================================================== the reported counter

    /// <summary>
    /// A refusal that happened has to survive the trip into the report, because the counter
    /// is the only thing that can settle from the field whether the sort works at all.
    /// </summary>
    [Fact]
    public void ARefusalCount_ReachesTheSweepReport()
    {
        SweepInfo info = new SweepInfo();

        MailService.ApplySweepCounters(info, SweepResult(sortRefusedFolders: 3), null);

        Assert.Equal(3, info.SortRefusedFolders);
    }

    /// <summary>
    /// And a healthy profile - the one the fix predicts - omits the field rather than
    /// reporting a zero, which is the boundary the mapping is written on.
    /// </summary>
    [Fact]
    public void AProfileThatRefusedNothing_OmitsTheCounterEntirely()
    {
        SweepInfo info = new SweepInfo();

        MailService.ApplySweepCounters(info, SweepResult(sortRefusedFolders: 0), null);

        Assert.Null(info.SortRefusedFolders);
    }

    // ================================================================== reading a row's date

    /// <summary>
    /// A date column index outside the row reads as NO date rather than throwing. The commit
    /// moved the time-zone reading out of <c>ReadRowDate</c> and left this bounds check as
    /// "the only decision left here" - and nothing pinned it: removing the upper half left
    /// all 1,936 tests passing, while a real short row would take an
    /// <see cref="IndexOutOfRangeException"/> straight out of a swept folder's row loop.
    /// </summary>
    [Fact]
    public void ADateColumnOutsideTheRow_ReadsAsNoDate()
    {
        Assert.Null(ReadRowDate(new object[] { "AAA" }, 1));
        Assert.Null(ReadRowDate(new object[] { "AAA" }, 5));
        Assert.Null(ReadRowDate(Array.Empty<object>(), 0));
        Assert.Null(ReadRowDate(new object[] { "AAA" }, -1));
    }

    /// <summary>
    /// And a date that IS there is read through the ONE shared table helper, which takes an
    /// unspecified kind as already-UTC. Reading it as local instead moves a resumed
    /// exhaustive scan's inclusive "at or before" bound EARLIER by the machine's offset,
    /// which skips the mail in that window and still reports the scan complete - in the one
    /// mode a caller picks because completeness matters.
    /// </summary>
    [Fact]
    public void ARowDate_IsReadThroughTheSharedTableHelper()
    {
        DateTime unspecified = new DateTime(2026, 8, 23, 18, 5, 0, DateTimeKind.Unspecified);

        DateTime? read = ReadRowDate(new object[] { "AAA", unspecified }, 1);

        Assert.NotNull(read);
        Assert.Equal(ComDateValue.FromTableValue(unspecified), read);
        Assert.Equal(DateTimeKind.Utc, read!.Value.Kind);
        Assert.Equal(unspecified.Ticks, read.Value.Ticks);
    }

    // ================================================================== harness

    private static DateTime? ReadRowDate(object[] values, int dateIndex)
    {
        MethodInfo method =
            typeof(OutlookComSession).GetMethod("ReadRowDate", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException(
                "OutlookComSession.ReadRowDate is gone or renamed. It turns a swept or scanned row's date column "
                + "into the UTC instant a resumed scan bounds itself by; if it moved, move this test with it.");

        return (DateTime?)method.Invoke(null, new object?[] { values, dateIndex });
    }

    private static ComSweepResult SweepResult(int sortRefusedFolders)
    {
        return new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 4,
            foldersSkipped: 0,
            sweptFolders: new[] { "alice@example.com/Inbox" },
            perStore: new[] { new ComStoreSweepCounters("alice@example.com", 4, 0, 0, 0) },
            sortRefusedFolders: sortRefusedFolders);
    }

    /// <summary>
    /// Runs the shipped <c>TryOrderSweptTable</c>. Private, so reflection - as
    /// <c>UnresolvedRecipientReportingTests</c> and <c>BudgetCompositionTests</c> already do
    /// for the same reason: the decision is worth pinning and does not belong on a public
    /// surface. A rename fails here loudly rather than silently skipping the checks.
    /// </summary>
    private static (bool ColumnAdded, bool SortApplied) Order(StandInTable table, params string[] ladder)
    {
        return Order(table, (IReadOnlyList<string>)ladder);
    }

    private static (bool ColumnAdded, bool SortApplied) Order(StandInTable table, IReadOnlyList<string> ladder)
    {
        MethodInfo method =
            typeof(OutlookComSession).GetMethod("TryOrderSweptTable", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException(
                "OutlookComSession.TryOrderSweptTable is gone or renamed. It is the code that walks the sweep's "
                + "sort ladder; if it moved, move these tests with it rather than deleting them.");

        object?[] args = { table, ladder, false, false };
        method.Invoke(null, args);
        return ((bool)args[2]!, (bool)args[3]!);
    }

    /// <summary>
    /// An Outlook <c>Table</c> shaped just enough for the ladder walk, and PUBLIC on purpose:
    /// the dynamic call sites live in OutlookAI.Core, and the C# runtime binder resolves
    /// members from the CALL SITE's accessibility context, so a private stand-in would fail
    /// to bind - and a binder failure is a <c>RuntimeBinderException</c>, which
    /// <c>IsComCallFailure</c> admits, so every rung would silently look "refused" and the
    /// tests would pass while measuring nothing.
    /// </summary>
    public sealed class StandInTable
    {
        private readonly HashSet<string> _refuseColumn;

        private readonly HashSet<string> _refuseSort;

        /// <summary>Creates a table that accepts, or refuses, the named spellings.</summary>
        public StandInTable(
            IReadOnlyCollection<string>? refuseColumn = null,
            IReadOnlyCollection<string>? refuseSort = null)
        {
            _refuseColumn = new HashSet<string>(refuseColumn ?? Array.Empty<string>(), StringComparer.Ordinal);
            _refuseSort = new HashSet<string>(refuseSort ?? Array.Empty<string>(), StringComparer.Ordinal);
        }

        /// <summary>Every spelling <c>Columns.Add</c> was called with, in order.</summary>
        public List<string> ColumnsAdded { get; } = new List<string>();

        /// <summary>Every spelling <c>Table.Sort</c> was called with, in order.</summary>
        public List<string> Sorted { get; } = new List<string>();

        /// <summary>The column collection, fetched fresh per rung exactly as Outlook's is.</summary>
        public object Columns => new StandInColumns(this);

        /// <summary>Orders the table, or refuses the way Outlook refuses.</summary>
        public void Sort(string property, bool descending)
        {
            Assert.True(descending, "the sweep must always ask for NEWEST first");
            Sorted.Add(property);
            if (_refuseSort.Contains(property))
            {
                throw new ArgumentException("Table.Sort refused '" + property + "'.", nameof(property));
            }
        }

        internal void Add(string property)
        {
            ColumnsAdded.Add(property);
            if (_refuseColumn.Contains(property))
            {
                throw new ArgumentException("Columns.Add refused '" + property + "'.", nameof(property));
            }
        }
    }

    /// <summary>The <c>Columns</c> half of the stand-in. Public for the reason given on <see cref="StandInTable"/>.</summary>
    public sealed class StandInColumns
    {
        private readonly StandInTable _table;

        /// <summary>Binds this collection to the table recording the calls.</summary>
        public StandInColumns(StandInTable table) => _table = table;

        /// <summary>Puts a spelling on the table, or refuses it.</summary>
        public void Add(string property) => _table.Add(property);
    }
}
