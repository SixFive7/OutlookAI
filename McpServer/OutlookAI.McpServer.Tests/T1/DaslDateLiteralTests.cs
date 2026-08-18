using System.Globalization;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the DASL date literal (measured defect, 2026-08-18). Outlook parses a DASL
/// date literal in the MACHINE locale, and both DASL formatters used to emit the
/// invariant US "MM/dd/yyyy". On a day-first machine that transposes day and month for
/// every date whose DAY is 12 or lower - roughly 40% of days - so an exhaustive search
/// bounded to 1-5 August 2026 answered about 8 January to 8 June instead, and a
/// freshness sweep whose window opened on a low day either over-selected months of mail
/// or (when the transposed date landed in the future) selected NOTHING while still
/// reporting foldersSwept and freshness "live".
/// <para>
/// The pins below therefore do two things a single "expected string" assertion cannot.
/// They fix the exact literal for a date whose day is 12 or lower - the case that was
/// silently wrong, and the case a month-first bug cannot survive - and they assert the
/// output does not move when the THREAD culture does. A test that only passed under the
/// developer's own locale would be the same defect wearing a different hat.
/// </para>
/// </summary>
public sealed class DaslDateLiteralTests
{
    /// <summary>
    /// Cultures the literal must be identical under: month-first, day-first (this
    /// machine), year-first, the invariant culture, and a culture whose default calendar
    /// is not Gregorian (th-TH renders 2026 as the Buddhist year 2569 - a current-culture
    /// formatter would emit that into the filter).
    /// </summary>
    private static readonly string[] CultureNames = ["en-US", "nl-NL", "de-DE", "lt-LT", "th-TH", ""];

    [Fact]
    public void DayTwelveOrLower_PinsTheExactYearFirstLiteral()
    {
        // Day 5 <= 12: under the old month-first literal Outlook read "08/05/2026" as
        // 8 May on this machine. Year-first, no locale can take 2026 for a day or a month.
        DateTime utc = new(2026, 8, 5, 7, 9, 3, DateTimeKind.Utc);

        string literal = DaslDateLiteral.FormatUtc(utc);

        Assert.Equal("2026-08-05 07:09:03", literal);
        Assert.DoesNotContain("/", literal, StringComparison.Ordinal);
        Assert.StartsWith("2026-", literal, StringComparison.Ordinal);
    }

    [Fact]
    public void EveryMonth_FirstTwelveDays_StayYearMonthDay()
    {
        // The whole ambiguous band in one pin: days 1-12 of every month are exactly the
        // dates a month-first literal can transpose.
        for (int month = 1; month <= 12; month++)
        {
            for (int day = 1; day <= 12; day++)
            {
                DateTime utc = new(2026, month, day, 0, 0, 0, DateTimeKind.Utc);

                string literal = DaslDateLiteral.FormatUtc(utc);

                Assert.Equal(
                    string.Format(CultureInfo.InvariantCulture, "2026-{0:D2}-{1:D2} 00:00:00", month, day),
                    literal);
            }
        }
    }

    [Fact]
    public void Literal_IsIdenticalUnderEveryThreadCulture()
    {
        DateTime utc = new(2026, 8, 5, 7, 9, 3, DateTimeKind.Utc);
        const string Expected = "2026-08-05 07:09:03";

        foreach (string name in CultureNames)
        {
            RunUnderCulture(name, () => Assert.Equal(Expected, DaslDateLiteral.FormatUtc(utc)));
        }
    }

    [Fact]
    public void ExhaustiveFilter_DateClauses_AreIdenticalUnderEveryThreadCulture()
    {
        // The formatter is pure, but the defect shipped inside the FILTER - so the filter
        // is pinned under the same cultures, not just the helper it now calls.
        DateTime since = new(2026, 8, 1, 8, 30, 0, DateTimeKind.Utc);
        DateTime before = new(2026, 8, 6, 0, 0, 0, DateTimeKind.Utc);
        string expected = ExhaustiveDaslFilter.Build(null, since, before, ExhaustiveEngine.Like);

        Assert.Contains("'2026-08-01 08:30:00'", expected, StringComparison.Ordinal);
        Assert.Contains("'2026-08-06 00:00:00'", expected, StringComparison.Ordinal);

        foreach (string name in CultureNames)
        {
            RunUnderCulture(
                name,
                () => Assert.Equal(expected, ExhaustiveDaslFilter.Build(null, since, before, ExhaustiveEngine.Like)));
        }
    }

    [Fact]
    public void Format_IsYearFirstWithAFourDigitYear()
    {
        Assert.Equal("yyyy-MM-dd HH:mm:ss", DaslDateLiteral.Format);
        Assert.StartsWith("yyyy", DaslDateLiteral.Format, StringComparison.Ordinal);
        Assert.DoesNotContain("MM/dd", DaslDateLiteral.Format, StringComparison.Ordinal);
        Assert.DoesNotContain("dd/MM", DaslDateLiteral.Format, StringComparison.Ordinal);
    }

    [Fact]
    public void LocalKind_IsConvertedToUtc_UnspecifiedIsTakenAsUtc()
    {
        DateTime utc = new(2026, 8, 5, 7, 9, 3, DateTimeKind.Utc);

        // Time-zone independent: the local rendering of that instant must produce the
        // same literal, because DASL compares in UTC.
        Assert.Equal(DaslDateLiteral.FormatUtc(utc), DaslDateLiteral.FormatUtc(utc.ToLocalTime()));

        // Unspecified is the contract every caller in Core relies on: already UTC.
        DateTime unspecified = new(2026, 8, 5, 7, 9, 3, DateTimeKind.Unspecified);
        Assert.Equal("2026-08-05 07:09:03", DaslDateLiteral.FormatUtc(unspecified));
    }

    [Fact]
    public void DaslAndWsSqlTiers_EmitTheSameDateLiteral()
    {
        // Both query languages now carry one shape, so a future change to either tier's
        // date handling has to break this pin first.
        DateTime since = new(2026, 8, 5, 7, 9, 3, DateTimeKind.Utc);
        string sql = WsSqlBuilder.Build(new IndexQuery
        {
            Scope = "mapi16://{S-1-5-21-1111111111-2222222222-3333333333-1001}/alice@example.com($deadbeef)",
            ReceivedOnOrAfterUtc = since,
        });

        Assert.Contains("'" + DaslDateLiteral.FormatUtc(since) + "'", sql, StringComparison.Ordinal);
    }

    private static void RunUnderCulture(string name, Action assert)
    {
        CultureInfo previousCulture = CultureInfo.CurrentCulture;
        CultureInfo previousUiCulture = CultureInfo.CurrentUICulture;
        try
        {
            CultureInfo culture = name.Length == 0 ? CultureInfo.InvariantCulture : new CultureInfo(name);
            CultureInfo.CurrentCulture = culture;
            CultureInfo.CurrentUICulture = culture;
            assert();
        }
        finally
        {
            CultureInfo.CurrentCulture = previousCulture;
            CultureInfo.CurrentUICulture = previousUiCulture;
        }
    }
}
