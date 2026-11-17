using System.Globalization;
using System.Reflection;
using System.Text;
using OutlookAI.Core.Com;
using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The check that every test which can read the machine's time zone carries
/// <c>[Trait("TimeZone", "Machine")]</c> (Q144, decided by the maintainer 2026-10-11). The build VM
/// runs the suite in UTC (Q126 (a)); its runner then switches the VM to W. Europe Standard Time and
/// runs the marked tests again, so a test that compares a local value with a UTC one meets a zone
/// where the two differ. A test this check finds and nobody marked would never meet one.
/// <para>
/// <see cref="MachineZoneReads"/> finds them: its remarks give the method and what it cannot see.
/// The controls below prove each way the walk is said to follow, and each it is said not to count, on
/// methods of <see cref="MachineZoneReadControls"/> that nothing ever runs.
/// </para>
/// <para>
/// <b>The rule for a test that reads the zone</b> (AGENTS.md, "Tests run on the build VM"): mark
/// it, and when you next touch it, rewrite it to set its own zone - as
/// <see cref="ComDateValueTests"/> does with <see cref="ComDateValueTests.OddFixedZone"/> - and drop
/// the marker once this check no longer finds it.
/// </para>
/// </summary>
public sealed class TimeZoneMarkerTests
{
    private readonly ITestOutputHelper _output;

    public TimeZoneMarkerTests(ITestOutputHelper output)
    {
        _output = output;
    }

    /// <summary>Each control, and whether the walk must find a read of the machine's zone in it.</summary>
    public static TheoryData<string, bool> Controls => new()
    {
        { nameof(MachineZoneReadControls.ReadsTheClock), true },
        { nameof(MachineZoneReadControls.ReadsTheZone), true },
        { nameof(MachineZoneReadControls.ConvertsToUtc), true },
        { nameof(MachineZoneReadControls.InALambda), true },
        { nameof(MachineZoneReadControls.InAnAsyncMethod), true },
        { nameof(MachineZoneReadControls.InALocalFunction), true },
        { nameof(MachineZoneReadControls.ThroughAStaticField), true },
        { nameof(MachineZoneReadControls.ThroughAnInterfaceImplementedHere), true },
        { nameof(MachineZoneReadControls.ThroughTheProduct), true },
        { nameof(MachineZoneReadControls.FormatsTheOffset), true },
        { nameof(MachineZoneReadControls.InterpolatesTheOffset), true },
        { nameof(MachineZoneReadControls.ParsesWithoutStyles), true },
        { nameof(MachineZoneReadControls.ReadsTheUtcClock), false },
        { nameof(MachineZoneReadControls.ConvertsInAZoneOfItsOwn), false },
        { nameof(MachineZoneReadControls.ThroughTheProductInAZoneOfItsOwn), false },
        { nameof(MachineZoneReadControls.FormatsRoundTrip), false },
        { nameof(MachineZoneReadControls.FormatsAQuotedZ), false },
        { nameof(MachineZoneReadControls.ParsesWithStyles), false },
    };

    /// <summary>
    /// THE CHECK: every non-live test that can reach a read of the machine's time zone carries the
    /// marker, on itself or on its class. A test that fails here names the chain that reaches the
    /// read, so the fix is visible from the message: mark it, or give it a zone of its own.
    /// </summary>
    [Fact]
    public void EveryTestThatCanReadTheMachinesZone_CarriesTheMarker()
    {
        MachineZoneReads.ZoneSurvey survey = MachineZoneReads.Survey();
        List<MachineZoneReads.ZoneReadingTest> reading = survey.Tests.Where(t => t.Chain != null).ToList();
        List<MachineZoneReads.ZoneReadingTest> unmarked = reading.Where(t => !t.IsMarked).ToList();
        List<MachineZoneReads.ZoneReadingTest> markedOnly = survey.Tests.Where(t => t.Chain == null && t.IsMarked).ToList();

        _output.WriteLine(
            $"{survey.Tests.Count} non-live test methods, {survey.MethodsWalked} methods walked; {reading.Count} can read the "
            + $"machine's zone, in {reading.Select(t => t.TestClass).Distinct().Count()} classes; {unmarked.Count} of them unmarked; "
            + $"{markedOnly.Count} marked tests read it in no way this walk sees.");

        if (unmarked.Count > 0)
        {
            StringBuilder message = new();
            message.AppendLine(
                $"{unmarked.Count} test(s) can read the machine's time zone but do not carry [Trait(\"{MachineZoneReads.TraitName}\", "
                + $"\"{MachineZoneReads.TraitValue}\")]. The build VM runs the suite in UTC, where that zone and UTC are one zone; "
                + "the runner runs the marked tests again in W. Europe Standard Time (Testbed/README.md section 1c). Mark each one - "
                + "on the method, or on its class when every test in the class reads the zone - or rewrite it to use a zone of its "
                + "own, as ComDateValueTests does with OddFixedZone:");
            foreach (MachineZoneReads.ZoneReadingTest test in unmarked)
            {
                message.AppendLine("  " + test.Name);
                message.AppendLine("      " + string.Join(" -> ", test.Chain!));
            }

            Assert.Fail(message.ToString());
        }
    }

    /// <summary>
    /// The survey is not green by finding nothing: it finds a test that reads
    /// <see cref="TimeZoneInfo.Local"/> in its own body, and passes over one that names a zone of its
    /// own - and it walked well beyond the tests themselves, into the product.
    /// </summary>
    [Fact]
    public void TheSurvey_FindsAKnownReader_AndPassesOverAZoneExplicitTest()
    {
        MachineZoneReads.ZoneSurvey survey = MachineZoneReads.Survey();

        MachineZoneReads.ZoneReadingTest reader = Single(survey, nameof(ComDateValueTests.TheTwoReadings_DifferByExactlyTheLocalOffset));
        Assert.NotNull(reader.Chain);
        Assert.Equal("reads TimeZoneInfo.Local", reader.Chain![^1]);

        Assert.Null(Single(survey, nameof(ComDateValueTests.AnItemValue_IsReadInTheZoneItIsGiven_NotInTheMachines)).Chain);

        Assert.True(survey.Tests.Count > 1000, $"only {survey.Tests.Count} non-live tests were surveyed");
        Assert.True(survey.MethodsWalked > 5000, $"the walk read only {survey.MethodsWalked} method bodies");

        static MachineZoneReads.ZoneReadingTest Single(MachineZoneReads.ZoneSurvey survey, string name)
        {
            return Assert.Single(survey.Tests, t => t.TestClass == typeof(ComDateValueTests) && t.Test.Name == name);
        }
    }

    /// <summary>Each way the walk follows, and each read it is said not to count, on a method nothing runs.</summary>
    [Theory]
    [MemberData(nameof(Controls))]
    public void TheWalk_SeesEachWayAControlReadsTheZone_AndNoMore(string control, bool reads)
    {
        MethodInfo method = typeof(MachineZoneReadControls).GetMethod(control, BindingFlags.NonPublic | BindingFlags.Static)
            ?? throw new InvalidOperationException("no control named " + control);

        IReadOnlyList<string>? chain = MachineZoneReads.ChainFrom(new[] { ((MethodBase)method, control) });

        Assert.True(
            (chain != null) == reads,
            control + (reads ? " reads the machine's zone, and the walk did not see it" : " was taken for a read: " + string.Join(" -> ", chain ?? Array.Empty<string>())));
    }

    /// <summary>
    /// What xunit runs around a test is walked too: a static constructor of the test's class - a field
    /// initializer every test in the class shares - and the member a <c>[MemberData]</c> names, which
    /// builds the test's arguments.
    /// </summary>
    [Fact]
    public void WhatXunitRunsAroundATest_IsWalkedToo()
    {
        MethodInfo plain = typeof(ZoneReadingClassControl).GetMethod(nameof(ZoneReadingClassControl.ReadsNothingItself))!;
        Assert.Null(MachineZoneReads.ChainFrom(new[] { ((MethodBase)plain, "the method alone") }));
        Assert.NotNull(MachineZoneReads.ChainFrom(MachineZoneReads.RootsOf(typeof(ZoneReadingClassControl), plain)));

        MethodInfo theory = typeof(TimeZoneMarkerTests).GetMethod(nameof(TheWalk_SeesEachWayAControlReadsTheZone_AndNoMore))!;
        Assert.Contains(
            MachineZoneReads.RootsOf(typeof(TimeZoneMarkerTests), theory),
            r => r.Root == typeof(TimeZoneMarkerTests).GetProperty(nameof(Controls))!.GetMethod);
    }

    /// <summary>
    /// Local-time formatting is a read when the format writes the machine's offset (<c>z</c>, quoted
    /// or escaped excepted) or converts with it (<c>U</c>); <c>o</c> and <c>K</c> write it only for a
    /// Local value and are not counted.
    /// </summary>
    [Theory]
    [InlineData("zzz", true)]
    [InlineData("yyyy-MM-dd'T'HH:mm:ss.fffzzz", true)]
    [InlineData("%z", true)]
    [InlineData("U", true)]
    [InlineData("HH:mm'z'", false)]
    [InlineData("HH:mm\"zone\"", false)]
    [InlineData("HH:mm\\z", false)]
    [InlineData("o", false)]
    [InlineData("u", false)]
    [InlineData("yyyy-MM-ddTHH:mm:ssK", false)]
    [InlineData("yyyy-MM-dd HH:mm:ss", false)]
    public void AFormat_ReadsTheZone_OnlyWhenItWritesOrConvertsByTheMachinesOffset(string format, bool reads)
    {
        Assert.Equal(reads, MachineZoneReads.FormatWritesTheMachinesOffset(format));
    }

    /// <summary>
    /// The build-VM runner's second pass selects the tests by exactly this marker - the filter is in
    /// the guest script, and a marker spelt differently from it would leave every marked test out of
    /// the second pass while this check stayed green.
    /// </summary>
    [Fact]
    public void TheMarker_IsWhatTheBuildVmRunnersSecondPassSelects()
    {
        string guestScript = File.ReadAllText(Path.Combine(RepoPaths.RepoRoot(), "Testbed", "guest", "Invoke-BuildVmRun.ps1"));

        Assert.Contains("$ZonePassMarker = '" + MachineZoneReads.Filter + "'", guestScript, StringComparison.Ordinal);
    }
}

/// <summary>
/// A class whose static field reads the machine's zone, so every test in a class like it would: the
/// walk reaches the read only through the class's static constructor. Never constructed.
/// </summary>
internal sealed class ZoneReadingClassControl
{
    private static readonly DateTime Today = DateTime.Today;

    public int ReadsNothingItself() => 1;

    public override string ToString() => Today.Ticks.ToString(CultureInfo.InvariantCulture);
}

/// <summary>An interface of this assembly, so a read behind it is reached only through its implementation.</summary>
internal interface IZoneClockControl
{
    DateTime Read();
}

/// <summary>The implementation the walk must follow the interface call into.</summary>
internal sealed class LocalZoneClockControl : IZoneClockControl
{
    public DateTime Read() => DateTime.Now;
}

/// <summary>A static field set from the machine's zone, read by a method that calls nothing.</summary>
internal static class ZoneReadingStaticControl
{
    internal static readonly DateTime Today = DateTime.Today;
}

/// <summary>
/// The controls <see cref="TimeZoneMarkerTests"/> walks. NOTHING CALLS THEM: the tests find them by
/// name through reflection, so no test reaches them through its own code - a test that called one
/// would read the machine's zone itself, and need the marker.
/// </summary>
internal static class MachineZoneReadControls
{
    internal static DateTime ReadsTheClock() => DateTime.Now;

    internal static TimeSpan ReadsTheZone(DateTime value) => TimeZoneInfo.Local.GetUtcOffset(value);

    internal static DateTime ConvertsToUtc(DateTime value) => value.ToUniversalTime();

    internal static Func<DateTime> InALambda() => () => DateTime.Now;

    internal static async Task<DateTime> InAnAsyncMethod()
    {
        await Task.Yield();
        return DateTime.Now;
    }

    internal static DateTime InALocalFunction()
    {
        return Read();

        static DateTime Read() => DateTime.Today;
    }

    internal static DateTime ThroughAStaticField() => ZoneReadingStaticControl.Today;

    internal static DateTime ThroughAnInterfaceImplementedHere(IZoneClockControl clock) => clock.Read();

    /// <summary>The product overload that converts in the machine's zone (Q95).</summary>
    internal static DateTime? ThroughTheProduct(DateTime? value) => ComDateValue.FromItemValue(value);

    internal static string FormatsTheOffset(DateTime value) => value.ToString("yyyy-MM-dd HH:mm zzz", CultureInfo.InvariantCulture);

    internal static string InterpolatesTheOffset(DateTime value) => string.Create(CultureInfo.InvariantCulture, $"{value:HH:mm zzz}");

    internal static DateTime ParsesWithoutStyles(string text) => DateTime.Parse(text, CultureInfo.InvariantCulture);

    internal static DateTime ReadsTheUtcClock() => DateTime.UtcNow;

    internal static DateTime ConvertsInAZoneOfItsOwn(DateTime utc) => TimeZoneInfo.ConvertTimeFromUtc(utc, ComDateValueTests.OddFixedZone);

    /// <summary>The same product conversion with a zone of the caller's own (Q95): no read of the machine's.</summary>
    internal static DateTime? ThroughTheProductInAZoneOfItsOwn(DateTime? value) => ComDateValue.FromItemValue(value, ComDateValueTests.OddFixedZone);

    internal static string FormatsRoundTrip(DateTime value) => value.ToString("o", CultureInfo.InvariantCulture);

    internal static string FormatsAQuotedZ(DateTime value) => value.ToString("HH:mm'z'", CultureInfo.InvariantCulture);

    internal static DateTime ParsesWithStyles(string text) => DateTime.Parse(
        text, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal);
}
