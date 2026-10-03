using System.Globalization;
using System.Reflection;

using OutlookAI.Core.Com;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the two readings of an Outlook date, and the fact that they are DIFFERENT on
/// purpose.
/// <para>
/// A COM-marshalled date always arrives with <see cref="DateTimeKind.Unspecified"/>, so the
/// value itself says nothing about which zone it is in. Until 2026-08-23 this solution held
/// both answers at once: the live tripwire census read a table value as already-UTC while
/// <c>OutlookComSession.ReadRowDate</c> called <c>ToUniversalTime</c> on it. Each side was
/// internally consistent, so neither could see the other, and the difference was exactly the
/// machine's UTC offset - the least visible size an error of this kind can have.
/// </para>
/// <para>
/// It matters in one place and not in the other. In the census, both ends of every
/// comparison come through one method, so items still match. In <c>ReadRowDate</c> the value
/// becomes a resumed exhaustive scan's inclusive "at or before" bound: a bound one offset too
/// EARLY skips the mail received in that window and reports the scan complete, in the one
/// mode a caller chooses because completeness matters.
/// </para>
/// <para>
/// <b>The item reading is pinned in a zone of the test's own, and on the machine's (Q95,
/// 2026-10-03).</b> On a machine whose zone is UTC the item conversion is the identity, so a
/// test that only ever used the machine's zone could not see whether it happened - CI runs in
/// UTC, and failed on exactly that while the maintainer's UTC+2 machine passed. The tests that
/// pin WHAT the conversion does name <see cref="OddFixedZone"/> or <see cref="OddDaylightZone"/>
/// and state the expected instant as a literal; <see cref="TheMachineOverload_GivesTheAnswerTheProductAlwaysGave"/>
/// holds the overload the product calls to the expression it replaced, on whatever zone the
/// machine has.
/// </para>
/// </summary>
public sealed class ComDateValueTests
{
    /// <summary>A wall-clock instant with no zone attached, which is all COM ever hands over.</summary>
    private static readonly DateTime Unspecified =
        new DateTime(2026, 6, 22, 14, 4, 23, DateTimeKind.Unspecified);

    /// <summary>
    /// A zone of the tests' own: UTC+07:13, all year. No real zone has that offset - every one is
    /// a whole quarter hour - so a conversion that quietly used the MACHINE's zone instead (UTC
    /// on CI, UTC+1 or +2 on the maintainer's machine, whatever it is anywhere else) cannot land
    /// on the same instant by coincidence. Built in code rather than looked up, so it is the same
    /// zone on a machine whose time-zone database differs or is missing.
    /// </summary>
    internal static readonly TimeZoneInfo OddFixedZone = TimeZoneInfo.CreateCustomTimeZone(
        "OutlookAI-Test-UTC+07:13", new TimeSpan(7, 13, 0), "OutlookAI test zone (UTC+07:13)", "OutlookAI test zone");

    /// <summary>
    /// The same base offset with an hour of daylight saving from 1 April 02:00 to 1 October 03:00
    /// (wall time), every year: UTC+08:13 in summer, UTC+07:13 in winter. It is how a conversion
    /// that used the zone's BASE offset rather than the offset the zone holds at that instant is
    /// told apart, which no fixed zone - and no UTC machine - can do.
    /// </summary>
    internal static readonly TimeZoneInfo OddDaylightZone = TimeZoneInfo.CreateCustomTimeZone(
        "OutlookAI-Test-UTC+07:13-DST",
        new TimeSpan(7, 13, 0),
        "OutlookAI test zone (UTC+07:13, daylight saving)",
        "OutlookAI test zone",
        "OutlookAI test zone (daylight)",
        new[]
        {
            TimeZoneInfo.AdjustmentRule.CreateAdjustmentRule(
                DateTime.MinValue.Date,
                DateTime.MaxValue.Date,
                TimeSpan.FromHours(1),
                TimeZoneInfo.TransitionTime.CreateFixedDateRule(new DateTime(1, 1, 1, 2, 0, 0), 4, 1),
                TimeZoneInfo.TransitionTime.CreateFixedDateRule(new DateTime(1, 1, 1, 3, 0, 0), 10, 1)),
        });

    /// <summary>
    /// The table reading: an unspecified kind is the instant it says it is. Asserted as
    /// "the wall clock did not move", not merely as "Kind is Utc" - relabelling and shifting
    /// are the two things that could happen here and only one of them is right.
    /// </summary>
    [Fact]
    public void ATableValue_WithNoKind_IsTakenAsAlreadyUtc()
    {
        DateTime? read = ComDateValue.FromTableValue(Unspecified);

        Assert.NotNull(read);
        Assert.Equal(DateTimeKind.Utc, read!.Value.Kind);
        Assert.Equal(Unspecified.Ticks, read.Value.Ticks);
    }

    /// <summary>A kind that IS stated is believed, in both directions.</summary>
    [Fact]
    public void ATableValue_ThatStatesItsKind_IsBelieved()
    {
        DateTime utc = DateTime.SpecifyKind(Unspecified, DateTimeKind.Utc);
        DateTime local = DateTime.SpecifyKind(Unspecified, DateTimeKind.Local);

        Assert.Equal(utc, ComDateValue.FromTableValue(utc));
        Assert.Equal(local.ToUniversalTime(), ComDateValue.FromTableValue(local));
    }

    /// <summary>
    /// A row whose date property was never set, and a column that came back as something
    /// else entirely, are the same answer: no date. Never a default instant, which would
    /// enter a scan's cursor as a real bound.
    /// </summary>
    [Theory]
    [InlineData(null)]
    [InlineData("2026-06-22T14:04:23Z")]
    [InlineData(0)]
    public void ATableValue_ThatIsNotADate_IsNoDate(object? value)
    {
        Assert.Null(ComDateValue.FromTableValue(value));
    }

    /// <summary>
    /// Q11, settled by measurement 2026-10-03 (first live run on a test guest, Office LTSC 2024,
    /// UTC+2): the zone a table reports a date in follows the SPELLING its column was added
    /// under. <c>ReceivedTime</c> read 13:44:18 for an item received 11:44:18Z - the item's own
    /// local wall time - and <c>urn:schemas:httpmail:datereceived</c> read the same folder in
    /// UTC. So an explicit name converts as local wall time and a namespace reference does not.
    /// The zone is held still (<see cref="OddFixedZone"/>) so the answer is the same on a UTC
    /// runner as anywhere else.
    /// </summary>
    [Fact]
    public void ATableValue_IsReadInTheZoneItsColumnSpellingReportsIn()
    {
        DateTime expectedFromLocal = new DateTime(2026, 6, 22, 6, 51, 23, DateTimeKind.Utc); // 14:04:23 - 07:13

        Assert.Equal(expectedFromLocal, ComDateValue.FromTableValue(Unspecified, "ReceivedTime", OddFixedZone));
        Assert.Equal(expectedFromLocal, ComDateValue.FromTableValue(Unspecified, "SentOn", OddFixedZone));

        DateTime asUtc = DateTime.SpecifyKind(Unspecified, DateTimeKind.Utc);
        Assert.Equal(asUtc, ComDateValue.FromTableValue(Unspecified, "urn:schemas:httpmail:datereceived", OddFixedZone));
        Assert.Equal(
            asUtc,
            ComDateValue.FromTableValue(Unspecified, "http://schemas.microsoft.com/mapi/proptag/0x0E060040", OddFixedZone));

        // Exactly the item reading for an explicit name, exactly the old table reading otherwise.
        Assert.Equal(ComDateValue.FromItemValue(Unspecified, OddFixedZone), ComDateValue.FromTableValue(Unspecified, "ReceivedTime", OddFixedZone));
        Assert.Equal(ComDateValue.FromTableValue(Unspecified), ComDateValue.FromTableValue(Unspecified, "urn:schemas:httpmail:datereceived", OddFixedZone));
    }

    /// <summary>
    /// The measured row itself, as a control that fails against the code it replaced: the
    /// duplicate the paged-scan acceptance caught on 2026-10-03 came from reading the explicit
    /// column's 15:29:46 local as 15:29:46Z, a bound two hours late, which re-admitted an item
    /// received 13:29:46Z. Read by spelling, the bound is the item's own instant.
    /// </summary>
    [Fact]
    public void TheMeasuredDuplicate_IsGoneWhenTheExplicitColumnIsReadAsLocal()
    {
        TimeZoneInfo utcPlusTwo = TimeZoneInfo.CreateCustomTimeZone("OutlookAI-Test-UTC+2", TimeSpan.FromHours(2), "UTC+2", "UTC+2");
        DateTime tableValue = new DateTime(2026, 9, 30, 15, 29, 46, DateTimeKind.Unspecified);
        DateTime received = new DateTime(2026, 9, 30, 13, 29, 46, DateTimeKind.Utc);

        Assert.Equal(received, ComDateValue.FromTableValue(tableValue, "ReceivedTime", utcPlusTwo));
        Assert.NotEqual(received, ComDateValue.FromTableValue(tableValue)); // the old reading: two hours late
    }

    [Theory]
    [InlineData("ReceivedTime", false)]
    [InlineData("SentOn", false)]
    [InlineData("urn:schemas:httpmail:datereceived", true)]
    [InlineData("URN:schemas:httpmail:date", true)]
    [InlineData("http://schemas.microsoft.com/mapi/proptag/0x0E060040", true)]
    [InlineData("https://schemas.example/whatever", true)]
    public void AColumnSpelling_IsANamespaceReference_OnlyWhenItLooksLikeOne(string spelling, bool expected)
    {
        Assert.Equal(expected, ComDateValue.IsNamespaceReference(spelling));
    }

    /// <summary>
    /// The item reading is the OPPOSITE default, because the object model returns local wall
    /// time. Same input, same absent kind, different answer - which is why these are two
    /// named methods and not one call that inspects the kind.
    /// </summary>
    [Fact]
    public void AnItemValue_WithNoKind_IsTakenAsLocalWallTime()
    {
        DateTime? read = ComDateValue.FromItemValue(Unspecified);

        Assert.NotNull(read);
        Assert.Equal(DateTimeKind.Utc, read!.Value.Kind);
        Assert.Equal(DateTime.SpecifyKind(Unspecified, DateTimeKind.Local).ToUniversalTime(), read.Value);
    }

    /// <summary>An item value that already knows it is UTC is not shifted a second time.</summary>
    [Fact]
    public void AnItemValue_AlreadyUtc_IsNotShiftedAgain()
    {
        DateTime utc = DateTime.SpecifyKind(Unspecified, DateTimeKind.Utc);

        Assert.Equal(utc, ComDateValue.FromItemValue(utc));
    }

    /// <summary>Null in, null out: an item with no received time is not an item at the epoch.</summary>
    [Fact]
    public void AnItemValue_ThatIsAbsent_StaysAbsent()
    {
        Assert.Null(ComDateValue.FromItemValue(null));
    }

    /// <summary>
    /// The invariant that says the two readings are deliberately opposed: for the SAME
    /// unspecified input they differ by exactly this machine's UTC offset at that instant.
    /// Written as arithmetic rather than as a fixed number of hours so it holds on a UTC
    /// machine and across a daylight-saving boundary, where a hard-coded offset would pin
    /// the developer's own summer instead of the rule.
    /// </summary>
    [Fact]
    public void TheTwoReadings_DifferByExactlyTheLocalOffset()
    {
        TimeSpan offset = TimeZoneInfo.Local.GetUtcOffset(Unspecified);

        DateTime table = ComDateValue.FromTableValue(Unspecified)!.Value;
        DateTime item = ComDateValue.FromItemValue(Unspecified)!.Value;

        Assert.Equal(table - offset, item);
    }

    /// <summary>
    /// The consistency claim the table reading rests on, checked rather than asserted in
    /// prose: the DASL literal that SELECTS a row and the value read back OUT of that row
    /// describe the same instant. If these two ever disagreed, a scan's resume bound would
    /// be expressed in one zone and evaluated in another.
    /// </summary>
    [Fact]
    public void TheTableReading_AgreesWithTheDaslLiteralThatSelectedTheRow()
    {
        string literal = DaslDateLiteral.FormatUtc(Unspecified);
        DateTime read = ComDateValue.FromTableValue(Unspecified)!.Value;

        Assert.Equal(literal, read.ToString(DaslDateLiteral.Format, CultureInfo.InvariantCulture));
    }

    // ------------------------------------------- the item reading, in a zone of the test's own (Q95)

    /// <summary>
    /// What the item conversion DOES, on every machine: 14:04:23 of wall time at UTC+07:13 is
    /// 06:51:23 UTC. Stated as a literal rather than recomputed through any conversion, so the
    /// expected value cannot share a defect with the code under test.
    /// </summary>
    [Fact]
    public void AnItemValue_IsReadInTheZoneItIsGiven_NotInTheMachines()
    {
        DateTime? read = ComDateValue.FromItemValue(Unspecified, OddFixedZone);

        Assert.NotNull(read);
        Assert.Equal(DateTimeKind.Utc, read!.Value.Kind);
        Assert.Equal(new DateTime(2026, 6, 22, 6, 51, 23, DateTimeKind.Utc), read.Value);
    }

    /// <summary>
    /// The offset is the one the zone holds AT THAT WALL TIME, not its base offset: an hour apart
    /// across the zone's daylight-saving boundary, and the summer one is an hour earlier in UTC.
    /// </summary>
    [Fact]
    public void AnItemValue_UsesTheOffsetItsZoneHoldsAtThatInstant()
    {
        DateTime summer = new(2026, 6, 22, 14, 4, 23, DateTimeKind.Unspecified);
        DateTime winter = new(2026, 1, 22, 14, 4, 23, DateTimeKind.Unspecified);

        Assert.Equal(new DateTime(2026, 6, 22, 5, 51, 23, DateTimeKind.Utc), ComDateValue.FromItemValue(summer, OddDaylightZone));
        Assert.Equal(new DateTime(2026, 1, 22, 6, 51, 23, DateTimeKind.Utc), ComDateValue.FromItemValue(winter, OddDaylightZone));
    }

    /// <summary>
    /// <see cref="TheTwoReadings_DifferByExactlyTheLocalOffset"/> in a zone of the test's own,
    /// where the difference is never zero: on a UTC machine the machine-zone version is
    /// <c>table == item</c>, which an item reading that forgot to convert satisfies as well.
    /// </summary>
    [Fact]
    public void TheTwoReadings_DifferByExactlyTheZonesOffset_OnEveryMachine()
    {
        DateTime table = ComDateValue.FromTableValue(Unspecified)!.Value;
        DateTime item = ComDateValue.FromItemValue(Unspecified, OddFixedZone)!.Value;

        Assert.Equal(new TimeSpan(7, 13, 0), table - item);
    }

    /// <summary>An item value that already says it is UTC is not shifted, whatever zone is given.</summary>
    [Fact]
    public void AnItemValue_AlreadyUtc_IsNotShiftedByTheZoneEither()
    {
        DateTime utc = DateTime.SpecifyKind(Unspecified, DateTimeKind.Utc);

        Assert.Equal(utc, ComDateValue.FromItemValue(utc, OddFixedZone));
        Assert.Null(ComDateValue.FromItemValue(null, OddFixedZone));
        Assert.Throws<ArgumentNullException>(() => ComDateValue.FromItemValue(Unspecified, null!));
    }

    /// <summary>
    /// The overload the product calls gives exactly what it gave before the zone became a
    /// parameter - <c>DateTime.SpecifyKind(value, DateTimeKind.Local).ToUniversalTime()</c> - in
    /// value AND kind, over every quarter hour of two whole years of THIS machine's zone,
    /// through each of its daylight-saving gaps and overlaps, at both ends of the
    /// <see cref="DateTime"/> range, and for a value that arrives already marked Local.
    /// <para>
    /// As strong as the machine's zone: on a UTC machine both sides are the identity and this
    /// proves only that, which is why the conversion itself is pinned above in a zone of the
    /// test's own. On a machine with daylight saving it is the proof that a gap or an overlap
    /// is still read the way it always was.
    /// </para>
    /// </summary>
    [Fact]
    public void TheMachineOverload_GivesTheAnswerTheProductAlwaysGave()
    {
        List<DateTime> wallTimes = new() { DateTime.MinValue, DateTime.MaxValue };
        for (DateTime t = new(2025, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
             t < new DateTime(2027, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
             t = t.AddMinutes(15))
        {
            wallTimes.Add(t);
        }

        int compared = 0;
        foreach (DateTime wall in wallTimes)
        {
            foreach (DateTime input in new[] { wall, DateTime.SpecifyKind(wall, DateTimeKind.Local) })
            {
                DateTime before = DateTime.SpecifyKind(input, DateTimeKind.Local).ToUniversalTime();
                DateTime? now = ComDateValue.FromItemValue(input);
                DateTime? explicitLocal = ComDateValue.FromItemValue(input, TimeZoneInfo.Local);

                bool same = now.HasValue && explicitLocal.HasValue
                    && now.Value.Ticks == before.Ticks && now.Value.Kind == before.Kind
                    && explicitLocal.Value.Ticks == before.Ticks && explicitLocal.Value.Kind == before.Kind;
                if (!same)
                {
                    Assert.Fail(
                        $"{input:o} ({input.Kind}) in {TimeZoneInfo.Local.Id}: was {before:o} ({before.Kind}), is now "
                        + $"{now:o} ({now?.Kind}) through the machine overload and {explicitLocal:o} ({explicitLocal?.Kind}) "
                        + "with the machine's zone passed in.");
                }

                compared++;
            }
        }

        // Two years of quarter hours, both kinds, plus the two ends - and not a quietly empty loop.
        Assert.Equal(2 * (((365 + 365) * 24 * 4) + 2), compared);
    }

    /// <summary>
    /// The overload the product calls hands the conversion THIS MACHINE's zone and nothing else,
    /// read out of its compiled body. The behavioural test above cannot see this on a UTC machine,
    /// where <see cref="TimeZoneInfo.Local"/> and <see cref="TimeZoneInfo.Utc"/> give the same
    /// answer - so a zone-less overload that passed UTC would be green on every CI run and wrong
    /// on every user's machine east or west of Greenwich.
    /// </summary>
    [Fact]
    public void TheMachineOverload_PassesTheMachinesZone()
    {
        MethodInfo zoneLess = typeof(ComDateValue).GetMethod(
            nameof(ComDateValue.FromItemValue), new[] { typeof(DateTime?) })!;

        List<MethodBase> called = CalleesOf(zoneLess);

        Assert.Contains(called, m => IsMethod(m, typeof(TimeZoneInfo), "get_Local"));
        Assert.DoesNotContain(called, m => IsMethod(m, typeof(TimeZoneInfo), "get_Utc"));
        Assert.Contains(called, m => IsMethod(
            m, typeof(ComDateValue), nameof(ComDateValue.FromItemValue), typeof(DateTime?), typeof(TimeZoneInfo)));
    }

    /// <summary>
    /// True when <paramref name="method"/> is <paramref name="name"/> on <paramref name="declaringType"/>
    /// taking exactly <paramref name="parameters"/> - compared by name and signature rather than by
    /// <see cref="MethodInfo"/> identity, so it cannot depend on which reflection path produced each.
    /// </summary>
    internal static bool IsMethod(MethodBase method, Type declaringType, string name, params Type[] parameters)
    {
        return method.DeclaringType == declaringType
            && string.Equals(method.Name, name, StringComparison.Ordinal)
            && method.GetParameters().Select(p => p.ParameterType).SequenceEqual(parameters);
    }

    /// <summary>
    /// Every method one method calls, resolved from its IL (<c>call</c> 0x28, <c>callvirt</c>
    /// 0x6F, <c>newobj</c> 0x73, each with a 4-byte token) - the reader
    /// <c>TripwireReRunDriverTests.CalleesOf</c> uses. The token is resolved, so a stray operand
    /// byte cannot pass for an instruction.
    /// </summary>
    internal static List<MethodBase> CalleesOf(MethodInfo method)
    {
        byte[] il = method.GetMethodBody()!.GetILAsByteArray()!;
        List<MethodBase> called = new();
        for (int i = 0; i + 4 < il.Length; i++)
        {
            if (il[i] != 0x28 && il[i] != 0x6F && il[i] != 0x73)
            {
                continue;
            }

            int token = il[i + 1] | (il[i + 2] << 8) | (il[i + 3] << 16) | (il[i + 4] << 24);
            try
            {
                MethodBase? resolved = method.Module.ResolveMethod(token);
                if (resolved != null)
                {
                    called.Add(resolved);
                }
            }
            catch (ArgumentException)
            {
                // Not a real call - the bytes happened to look like one.
            }
        }

        return called;
    }
}
