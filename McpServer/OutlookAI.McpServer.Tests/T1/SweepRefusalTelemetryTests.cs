using System.Reflection;

using OutlookAI.Core.Com;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The freshness sweep's sort-refusal telemetry, pinned where CI can reach it.
/// <para>
/// <c>sweep.sortRefusedFolders</c> is the counter that answers, from an ordinary run against a
/// real profile, whether <c>Table.Sort</c> works at all - the defect fixed in <c>bea7fc9</c>
/// had been live for the whole life of the feature and this number is how anyone would learn
/// it had come back. It is read by comparing against zero, which is exactly the reading a
/// counter nobody wired up also produces.
/// </para>
/// <para>
/// A mutation pass over the fix left six survivors and five of them were this counter's own
/// wiring: the refusal decision could be collapsed to <c>!sortApplied</c> or hard-coded false,
/// the guard that increments the tally could be inverted at both sweep call sites, and either
/// of the two hand-copied <see cref="ComSweepResult"/> argument lists could report 0 - all
/// with the suite green, because every one of them sat inside a COM-driven method no
/// mailbox-free test can enter. The decisions are now small pure methods and the two argument
/// lists are one builder, so this file can hold them.
/// </para>
/// <para>
/// TIER 1: pure. No COM, no Outlook, no mailbox. The builder and the sweep's own tally are
/// private, so they are reached by reflection exactly as <c>TryOrderSweptTable</c> already is
/// in <see cref="SweepSortMutationTests"/>, and a rename fails here loudly rather than
/// silently skipping the checks.
/// </para>
/// </summary>
public sealed class SweepRefusalTelemetryTests
{
    // ============================================================ the refusal decision itself

    /// <summary>
    /// The one distinction the counter is built on: Outlook would not CARRY the property is
    /// not the same event as Outlook carried it and then refused to ORDER by it. Only the
    /// second is a refusal, because only the second means Outlook was actually asked.
    /// <para>
    /// Both mutations that survived live here. Collapsing this to <c>!sortApplied</c> counts a
    /// provider whose table has no such column at all - which is a capability gap, not a
    /// refusal - and would make the number non-zero on healthy profiles, retiring the only
    /// signal there is. Hard-coding it false empties the counter, which reads as the good news
    /// it is supposed to be evidence of.
    /// </para>
    /// </summary>
    [Theory]
    [InlineData(true, false, true)]   // the column went on and Sort still threw: a refusal.
    [InlineData(true, true, false)]   // sorted: nothing to report.
    [InlineData(false, false, false)] // no column on any rung, so no sort was ever asked for.
    [InlineData(false, true, false)]  // cannot happen - a sort needs its column - and is still not a refusal.
    public void ARefusal_IsAColumnThatWentOnAndASortThatDidNot(bool columnAdded, bool sortApplied, bool expected)
    {
        Assert.Equal(expected, OutlookComSession.SweepSortWasRefused(columnAdded, sortApplied));
    }

    // ================================================================== counting the refusals

    /// <summary>
    /// The counter rises on a refusal and on nothing else. Inverting this turns a healthy
    /// profile's zero into "every folder refused" and a broken profile's evidence into a zero,
    /// and both would be believed.
    /// </summary>
    [Theory]
    [InlineData(0, false, 0)]
    [InlineData(0, true, 1)]
    [InlineData(4, false, 4)]
    [InlineData(4, true, 5)]
    public void TheCount_RisesByOne_OnlyOnARefusal(int before, bool sortRefused, int expected)
    {
        Assert.Equal(expected, OutlookComSession.AddSortRefusal(before, sortRefused));
    }

    /// <summary>
    /// And it counts FOLDERS, one per call, rather than latching or saturating: a sweep that
    /// meets the refusal on every folder of every store is what tells the namespace-property
    /// cause apart from a single awkward provider, and a flag could not say that.
    /// </summary>
    [Fact]
    public void TheCount_AccumulatesOverFolders()
    {
        int count = 0;
        for (int i = 0; i < 7; i++)
        {
            count = OutlookComSession.AddSortRefusal(count, sortRefused: true);
        }

        Assert.Equal(7, count);
    }

    /// <summary>
    /// The two decisions composed the way both sweep walks compose them, so a folder is
    /// counted for exactly the outcome the counter is documented to mean.
    /// </summary>
    [Theory]
    [InlineData(true, false, 1)]
    [InlineData(true, true, 0)]
    [InlineData(false, false, 0)]
    public void OneFolder_IsCountedForARefusalAndForNothingElse(bool columnAdded, bool sortApplied, int expected)
    {
        Assert.Equal(
            expected,
            OutlookComSession.AddSortRefusal(0, OutlookComSession.SweepSortWasRefused(columnAdded, sortApplied)));
    }

    // ============================================================== the tally reaches the result

    /// <summary>
    /// Every counter a sweep collects arrives on the result under its own name. Distinct
    /// values throughout, so a mapping that crossed two wires fails rather than agreeing with
    /// itself - which is what a table of ones would do.
    /// </summary>
    [Fact]
    public void TheBuilder_CarriesEveryCounterOffTheTally()
    {
        object tally = NewTally();
        Set(tally, "Failed", 3);
        Set(tally, "RowsUnreadable", 5);
        Set(tally, "SortRefused", 7);
        Set(tally, "BodiesTruncated", 11);
        Capped(tally, "ItemCapped").Add("alice@example.com/Inbox");
        Capped(tally, "ItemCappedUnsorted").Add("alice@example.com/Sent Items");

        ComSweepResult result = Build(
            tally,
            items: new[] { Brief("AAA"), Brief("BBB") },
            sweptFolders: new[] { "alice@example.com/Inbox", "alice@example.com/Sent Items" },
            foldersSkipped: 13,
            foldersAbsent: 17,
            perStore: new[] { new ComStoreSweepCounters("alice@example.com", 2, 13, 3, 17, 5) },
            storesUnnamed: 19);

        Assert.Equal(3, result.FoldersFailed);
        Assert.Equal(5, result.RowsUnreadable);
        Assert.Equal(7, result.SortRefusedFolders);
        Assert.Equal(11, result.BodiesTruncated);
        Assert.Equal(13, result.FoldersSkipped);
        Assert.Equal(17, result.FoldersAbsent);
        Assert.Equal(19, result.StoresUnnamed);
        Assert.Equal(2, result.Items.Count);
        Assert.Equal(new[] { "alice@example.com/Inbox" }, result.ItemCappedFolders);
        Assert.Equal(new[] { "alice@example.com/Sent Items" }, result.ItemCappedFoldersUnsorted);
        Assert.Equal("alice@example.com", Assert.Single(result.PerStore).StoreDisplayName);
    }

    /// <summary>
    /// A profile that refused nothing reports nothing, which is the reading the whole counter
    /// rests on: <c>sweep.sortRefusedFolders</c> is omitted at zero
    /// (<c>MailService.ApplySweepCounters</c>), so a builder that always reported zero would be
    /// indistinguishable from the fix working.
    /// </summary>
    [Fact]
    public void TheBuilder_ReportsZeroOnlyWhenTheTallyCountedNothing()
    {
        Assert.Equal(0, Build(NewTally()).SortRefusedFolders);
    }

    /// <summary>
    /// The folder count is the list, not a number beside it. Both call sites passed
    /// <c>sweptFolders.Count</c> by hand, and a count that can disagree with the labels the
    /// advice sentences are built from is the same silent-cap defect one size smaller.
    /// </summary>
    [Fact]
    public void TheBuilder_CountsTheFolderLabelsItWasGiven()
    {
        ComSweepResult result = Build(
            NewTally(),
            sweptFolders: new[] { "a/Inbox", "a/Sent Items", "b/Inbox" },
            foldersSkipped: 2,
            foldersAbsent: 1);

        Assert.Equal(3, result.FoldersSwept);
        Assert.Equal(3, result.SweptFolders.Count);
    }

    /// <summary>
    /// The five latched bounds stay apart. One at a time, because five booleans set together
    /// cannot tell a crossed pair from a correct one - and they lead to five different
    /// remedies in the advice a caller is given.
    /// </summary>
    [Theory]
    [InlineData("FolderCapReached")]
    [InlineData("DepthLimitReached")]
    [InlineData("TimeBudgetExceeded")]
    [InlineData("BodyBudgetExhausted")]
    [InlineData("SweepBudgetExpired")]
    public void TheBuilder_KeepsTheLatchedBoundsApart(string latched)
    {
        object tally = NewTally();
        Set(tally, latched, true);

        ComSweepResult result = Build(tally);

        Dictionary<string, bool> reported = new Dictionary<string, bool>(StringComparer.Ordinal)
        {
            ["FolderCapReached"] = result.FolderCapReached,
            ["DepthLimitReached"] = result.DepthLimitReached,
            ["TimeBudgetExceeded"] = result.TimeBudgetExceeded,
            ["BodyBudgetExhausted"] = result.BodyBudgetExhausted,
            ["SweepBudgetExpired"] = result.SweepBudgetExpired,
        };

        foreach (KeyValuePair<string, bool> entry in reported)
        {
            Assert.Equal(string.Equals(entry.Key, latched, StringComparison.Ordinal), entry.Value);
        }
    }

    // ================================================================================ harness

    private static ComMailBrief Brief(string entryId)
    {
        return new ComMailBrief(
            entryId: entryId,
            storeDisplayName: "alice@example.com",
            storeId: "store-alice",
            folderName: "Inbox",
            folderKind: "inbox",
            subject: "a swept item",
            senderName: "Bob",
            senderAddress: "bob@example.com",
            receivedTime: new DateTime(2026, 8, 24, 9, 0, 0, DateTimeKind.Utc),
            isRead: true,
            hasAttachments: false,
            sizeBytes: 2048,
            body: null);
    }

    private static Type TallyType()
    {
        return typeof(OutlookComSession).GetNestedType("SweepTally", BindingFlags.NonPublic)
            ?? throw new InvalidOperationException(
                "OutlookComSession.SweepTally is gone or renamed. It is the sweep's counter bag, and the builder "
                + "these tests pin maps it onto ComSweepResult; if it moved, move these tests with it.");
    }

    private static object NewTally()
    {
        return Activator.CreateInstance(
            TallyType(),
            BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public,
            null,
            new object?[] { 0 },
            null)
            ?? throw new InvalidOperationException("SweepTally could not be constructed.");
    }

    private static PropertyInfo Property(object tally, string name)
    {
        return tally.GetType().GetProperty(name, BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public)
            ?? throw new InvalidOperationException(
                "SweepTally." + name + " is gone or renamed - so the counter it fed no longer reaches the result, "
                + "and nothing else in this suite would have said so.");
    }

    /// <summary>
    /// Writes one counter. <c>SweepBudgetExpired</c> latches through a private setter, which
    /// reflection can drive; the field fallback is there so that turning it into a plain field
    /// fails the assertion rather than the harness.
    /// </summary>
    private static void Set(object tally, string name, object value)
    {
        PropertyInfo property = Property(tally, name);
        if (property.GetSetMethod(nonPublic: true) != null)
        {
            property.SetValue(tally, value);
            return;
        }

        FieldInfo field =
            tally.GetType().GetFields(BindingFlags.Instance | BindingFlags.NonPublic)
                .FirstOrDefault(f => f.Name.Contains(name, StringComparison.Ordinal))
            ?? throw new InvalidOperationException("SweepTally." + name + " has neither a setter nor a backing field.");

        field.SetValue(tally, value);
    }

    private static List<string> Capped(object tally, string name)
    {
        return (List<string>?)Property(tally, name).GetValue(tally)
            ?? throw new InvalidOperationException("SweepTally." + name + " was null.");
    }

    private static ComSweepResult Build(
        object tally,
        IReadOnlyList<ComMailBrief>? items = null,
        IReadOnlyList<string>? sweptFolders = null,
        int foldersSkipped = 0,
        int foldersAbsent = 0,
        IReadOnlyList<ComStoreSweepCounters>? perStore = null,
        int storesUnnamed = 0)
    {
        MethodInfo method =
            typeof(OutlookComSession).GetMethod("BuildSweepResult", BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException(
                "OutlookComSession.BuildSweepResult is gone or renamed. It is the ONE mapping from a finished "
                + "sweep's counters to the result a caller reads; if it moved, move these tests with it rather "
                + "than deleting them - two copies of that mapping is what they exist to prevent.");

        object?[] args =
        {
            tally,
            items ?? Array.Empty<ComMailBrief>(),
            sweptFolders ?? Array.Empty<string>(),
            foldersSkipped,
            foldersAbsent,
            perStore ?? Array.Empty<ComStoreSweepCounters>(),
            storesUnnamed,
        };

        return (ComSweepResult?)method.Invoke(null, args)
            ?? throw new InvalidOperationException("BuildSweepResult returned null.");
    }
}
