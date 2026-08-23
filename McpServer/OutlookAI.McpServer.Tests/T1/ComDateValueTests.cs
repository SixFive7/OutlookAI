using System.Globalization;

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
/// </summary>
public sealed class ComDateValueTests
{
    /// <summary>A wall-clock instant with no zone attached, which is all COM ever hands over.</summary>
    private static readonly DateTime Unspecified =
        new DateTime(2026, 6, 22, 14, 4, 23, DateTimeKind.Unspecified);

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
}
