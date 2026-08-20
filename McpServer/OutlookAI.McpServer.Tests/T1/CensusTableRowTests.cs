using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the projection the tripwire census reads a mailbox through since 2026-08-20, when
/// the item-by-item walk was replaced by a bulk table read because it could not finish
/// inside the STA budget on an Exchange profile.
/// <para>
/// The projection is the whole of what the guard compares, so these cases are about the
/// three questions a firing turns on: can this row NAME an item, can it be recognised
/// somewhere else in the store afterwards, and was it the suite's own. The fourth question
/// - what the census must NOT keep - is pinned here too: a subject goes in and only a
/// boolean comes out.
/// </para>
/// </summary>
public sealed class CensusTableRowTests
{
    private const int IdIndex = 0;
    private const int SubjectIndex = 1;
    private const int ReceivedIndex = 2;
    private const int SizeIndex = 3;

    private static CensusColumnMap Map()
    {
        return new CensusColumnMap(IdIndex, SubjectIndex, ReceivedIndex, SizeIndex, 4);
    }

    private static object?[] Row(object? id, object? subject, object? received, object? size)
    {
        return new[] { id, subject, received, size };
    }

    [Fact]
    public void ACompleteRow_BecomesAnIdentifiableItem()
    {
        CensusItem? item = CensusTableRow.Project(
            Row("000102EID", "Quarterly figures", new DateTime(2026, 8, 20, 9, 30, 0, DateTimeKind.Utc), 24_576),
            Map());

        Assert.NotNull(item);
        Assert.Equal("000102EID", item!.Value.Id);
        Assert.Equal("2026-08-20T09:30:00Z/24576", item.Value.Fingerprint);
        Assert.False(item.Value.Tagged);
    }

    [Fact]
    public void ARowThatNamesNoItem_IsRejectedRatherThanSkipped()
    {
        // A skipped row would read exactly like an item that had been deleted, which is the
        // one wrong answer this guard must never give. The caller turns null into a
        // count-only folder instead.
        CensusColumnMap map = Map();

        Assert.Null(CensusTableRow.Project(Row(null, "s", DateTime.UtcNow, 1), map));
        Assert.Null(CensusTableRow.Project(Row(string.Empty, "s", DateTime.UtcNow, 1), map));
        Assert.Null(CensusTableRow.Project(Row(42, "s", DateTime.UtcNow, 1), map));
    }

    [Fact]
    public void AnUnusableColumnMap_ProjectsNothing()
    {
        // A table that cannot say when an item arrived, how big it is or whether it is the
        // suite's own cannot support the identity reading at all.
        object?[] row = Row("ID", "s", DateTime.UtcNow, 1);

        Assert.Null(CensusTableRow.Project(row, new CensusColumnMap(-1, 1, 2, 3, 4)));
        Assert.Null(CensusTableRow.Project(row, new CensusColumnMap(0, -1, 2, 3, 4)));
        Assert.Null(CensusTableRow.Project(row, new CensusColumnMap(0, 1, -1, 3, 4)));
        Assert.Null(CensusTableRow.Project(row, new CensusColumnMap(0, 1, 2, -1, 4)));
        Assert.False(new CensusColumnMap(0, 1, 2, -1, 4).IsUsable);
        Assert.True(Map().IsUsable);
    }

    [Fact]
    public void ARowShorterThanTheMapPromised_IsRejected()
    {
        Assert.Null(CensusTableRow.Project(new object?[] { "ID" }, Map()));
    }

    [Fact]
    public void TheSuitesOwnMailIsRecognisedByItsTag_CaseInsensitively()
    {
        // Attribution turns on this boolean: a tagged item leaving a mailbox the suite may
        // not write to is the suite and nothing else.
        Assert.True(Tagged("Re: " + LiveOutlookTestMailer.SubjectTag + " round trip"));
        Assert.True(Tagged(LiveOutlookTestMailer.SubjectTag.ToUpperInvariant()));
        Assert.False(Tagged("An ordinary subject"));
        Assert.False(Tagged(null));

        static bool Tagged(string? subject)
        {
            CensusItem? item = CensusTableRow.Project(
                Row("ID", subject, new DateTime(2026, 8, 20, 0, 0, 0, DateTimeKind.Utc), 1), Map());
            Assert.NotNull(item);
            return item!.Value.Tagged;
        }
    }

    [Fact]
    public void ASubjectNeverSurvivesTheProjection()
    {
        // S3/S4: the census may not carry another mailbox's content anywhere a log or a
        // failure message could reach. Only the flag leaves this method.
        const string Secret = "Salary review for A. Person";
        CensusItem? item = CensusTableRow.Project(
            Row("ID", Secret, new DateTime(2026, 8, 20, 0, 0, 0, DateTimeKind.Utc), 7), Map());

        Assert.NotNull(item);
        Assert.DoesNotContain("Salary", item!.Value.Id, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Salary", item.Value.Fingerprint ?? string.Empty, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AnItemWithNoReceivedTimeOrNoSize_IsStillCensusedButCannotBeTracedAcrossAMove()
    {
        // Drafts and report items have no received time. They keep their EntryID, so a
        // removal is still detected; what is lost is the exoneration that proves a FILING.
        CensusItem? draft = CensusTableRow.Project(Row("ID", "Draft", null, 1_024), Map());
        Assert.NotNull(draft);
        Assert.Null(draft!.Value.Fingerprint);

        CensusItem? unsized = CensusTableRow.Project(
            Row("ID", "Mail", new DateTime(2026, 8, 20, 0, 0, 0, DateTimeKind.Utc), null), Map());
        Assert.NotNull(unsized);
        Assert.Null(unsized!.Value.Fingerprint);
    }

    [Fact]
    public void AnUnspecifiedKindIsReadAsUtc_SoTwoCensusesAgree()
    {
        // The Table object reports date-time values in UTC, and this solution already takes
        // an unspecified kind as UTC everywhere else (DaslDateLiteral). What matters most is
        // that the reading is the SAME at both ends of a comparison: an item that moved
        // between two folders is recognised by this string and nothing else.
        DateTime moment = new(2026, 8, 20, 9, 30, 0);
        Assert.Equal(
            "2026-08-20T09:30:00Z/10",
            CensusTableRow.Fingerprint(DateTime.SpecifyKind(moment, DateTimeKind.Unspecified), 10));
        Assert.Equal(
            "2026-08-20T09:30:00Z/10",
            CensusTableRow.Fingerprint(DateTime.SpecifyKind(moment, DateTimeKind.Utc), 10));
        Assert.Equal(
            CensusTableRow.Fingerprint(DateTime.SpecifyKind(moment, DateTimeKind.Local).ToUniversalTime(), 10),
            CensusTableRow.Fingerprint(DateTime.SpecifyKind(moment, DateTimeKind.Local), 10));
    }

    [Fact]
    public void EverySizeShapeATableCanHandBackReadsTheSame()
    {
        // A variant carrying a MAPI long is not contractually one CLR type, and a
        // fingerprint that changed shape with it would turn one item into two.
        DateTime moment = new(2026, 8, 20, 9, 30, 0, DateTimeKind.Utc);
        string expected = "2026-08-20T09:30:00Z/4096";

        Assert.Equal(expected, CensusTableRow.Fingerprint(moment, 4096));
        Assert.Equal(expected, CensusTableRow.Fingerprint(moment, 4096L));
        Assert.Equal(expected, CensusTableRow.Fingerprint(moment, (short)4096));
        Assert.Equal(expected, CensusTableRow.Fingerprint(moment, 4096u));
        Assert.Equal(expected, CensusTableRow.Fingerprint(moment, 4096d));
        Assert.Equal(expected, CensusTableRow.Fingerprint(moment, "4096"));
        Assert.Null(CensusTableRow.Fingerprint(moment, "not a number"));
        Assert.Null(CensusTableRow.Fingerprint("not a date", 4096));
    }

    [Fact]
    public void TheColumnSpellingsCoverBothFormsOutlookAccepts()
    {
        // One spelling has already cost this repository a shipped defect (Table.Sort). A
        // column that will not add leaves the folder counted rather than identified, so the
        // fallback spelling is what keeps the guard's identity half alive on a store that
        // only exposes the MAPI property.
        Assert.Contains("ReceivedTime", CensusTableRow.ReceivedColumnNames);
        Assert.Contains("urn:schemas:httpmail:datereceived", CensusTableRow.ReceivedColumnNames);
        Assert.Contains("Size", CensusTableRow.SizeColumnNames);
        Assert.Contains("http://schemas.microsoft.com/mapi/proptag/0x0E080003", CensusTableRow.SizeColumnNames);
        Assert.Contains("EntryID", CensusTableRow.IdColumnNames);
        Assert.Contains("Subject", CensusTableRow.SubjectColumnNames);
    }

    [Fact]
    public void TheBulkReadAsksForAPositiveNumberOfRows()
    {
        // A non-positive batch would ask the table for nothing on every pass, and every
        // folder would fall back to a count with no failure anywhere to say so.
        Assert.True(LiveOutlookTestMailer.CensusTableRowBatch > 0);
    }

    [Fact]
    public void TheColumnMapIsBuiltFromTheTablesOwnNames_FallbackSpellingIncluded()
    {
        CensusColumnMap plain = CensusTableRow.MapColumns(
            new[] { "EntryID", "Subject", "CreationTime", "ReceivedTime", "Size" });
        Assert.True(plain.IsUsable);
        Assert.Equal(0, plain.Id);
        Assert.Equal(1, plain.Subject);
        Assert.Equal(3, plain.Received);
        Assert.Equal(4, plain.Size);
        Assert.Equal(5, plain.ColumnCount);

        // A store that only accepted the namespace/proptag spellings is still fully usable.
        CensusColumnMap viaNamespace = CensusTableRow.MapColumns(
            new[]
            {
                "EntryID",
                "urn:schemas:httpmail:subject",
                "urn:schemas:httpmail:datereceived",
                "http://schemas.microsoft.com/mapi/proptag/0x0E080003",
            });
        Assert.True(viaNamespace.IsUsable);
        Assert.Equal(2, viaNamespace.Received);
        Assert.Equal(3, viaNamespace.Size);

        // A column that never landed leaves the map unusable, which counts the folder.
        Assert.False(CensusTableRow.MapColumns(new[] { "EntryID", "Subject", "ReceivedTime" }).IsUsable);
        Assert.False(CensusTableRow.MapColumns(Array.Empty<string>()).IsUsable);

        // Outlook decides how it spells a column name back to us, and the census must not
        // lose the identity reading over a capital letter.
        Assert.True(CensusTableRow.MapColumns(new[] { "entryid", "subject", "receivedtime", "size" }).IsUsable);
    }

    [Fact]
    public void ABulkReadIsAcceptedOnlyInTheExactShapeItWasAskedFor()
    {
        object?[,] block = new object?[2, 4];

        Assert.True(CensusTableRow.TryReadBlock(block, 2, 4, out Array? rows));
        Assert.NotNull(rows);

        // Its own transpose, a short block, a long block, a one-dimensional array and a
        // non-array all read as "this is not the answer I asked for".
        Assert.False(CensusTableRow.TryReadBlock(new object?[4, 2], 2, 4, out _));
        Assert.False(CensusTableRow.TryReadBlock(new object?[1, 4], 2, 4, out _));
        Assert.False(CensusTableRow.TryReadBlock(new object?[3, 4], 2, 4, out _));
        Assert.False(CensusTableRow.TryReadBlock(new object?[2, 3], 2, 4, out _));
        Assert.False(CensusTableRow.TryReadBlock(new object?[2, 5], 2, 4, out _));
        // Rank 1 with the RIGHT number of entries, so only the rank check can reject it:
        // a shape test that leans on the row count cannot tell a 1-D array from a 2-D one.
        Assert.False(CensusTableRow.TryReadBlock(new object?[2], 2, 4, out _));
        Assert.False(CensusTableRow.TryReadBlock(new object?[4], 2, 4, out _));
        Assert.False(CensusTableRow.TryReadBlock(null, 2, 4, out Array? none));
        Assert.Null(none);
    }

    [Fact]
    public void ABlockOfRowsProjectsInOrder_AndOneBadRowAbandonsTheFolder()
    {
        DateTime moment = new(2026, 8, 20, 9, 30, 0, DateTimeKind.Utc);
        object?[,] block = new object?[2, 4];
        block[0, IdIndex] = "ID-1";
        block[0, SubjectIndex] = "One";
        block[0, ReceivedIndex] = moment;
        block[0, SizeIndex] = 10;
        block[1, IdIndex] = "ID-2";
        block[1, SubjectIndex] = LiveOutlookTestMailer.SubjectTag + " two";
        block[1, ReceivedIndex] = moment;
        block[1, SizeIndex] = 20;

        List<CensusItem> walked = new();
        HashSet<string> seen = new(StringComparer.Ordinal);
        Assert.True(CensusTableRow.ProjectRows(block, Map(), walked, seen));
        Assert.Equal(new[] { "ID-1", "ID-2" }, walked.Select(i => i.Id));
        Assert.True(walked[1].Tagged);

        // The same id twice means the table shifted under the read, so the whole folder is
        // abandoned rather than deduplicated: a list built from two moments would report
        // whichever item was not in both as removed.
        Assert.False(CensusTableRow.ProjectRows(block, Map(), walked, seen));

        object?[,] nameless = new object?[1, 4];
        Assert.False(CensusTableRow.ProjectRows(nameless, Map(), new List<CensusItem>(), new HashSet<string>()));
    }

    [Fact]
    public void ABlockWithANonZeroOrigin_IsReadFromItsOwnBounds()
    {
        // A SAFEARRAY may declare any origin, and an assumed zero would shift every row.
        Array block = Array.CreateInstance(typeof(object), new[] { 1, 4 }, new[] { 1, 1 });
        block.SetValue("ID-1", 1, 1 + IdIndex);
        block.SetValue("One", 1, 1 + SubjectIndex);
        block.SetValue(new DateTime(2026, 8, 20, 9, 30, 0, DateTimeKind.Utc), 1, 1 + ReceivedIndex);
        block.SetValue(10, 1, 1 + SizeIndex);

        List<CensusItem> walked = new();
        Assert.True(CensusTableRow.ProjectRows(block, Map(), walked, new HashSet<string>(StringComparer.Ordinal)));
        Assert.Equal("ID-1", Assert.Single(walked).Id);
    }

    [Fact]
    public void AFiledItemProjectedFromTwoTableReadsIsStillRecognisedAsFiled()
    {
        // End to end over the projection: the exoneration the census exists to be able to
        // prove has to survive the row format, not just the comparison logic. Outlook
        // reissues an EntryID on a move, so the arrival is matched by the fingerprint.
        DateTime moment = new(2026, 8, 20, 9, 30, 0, DateTimeKind.Utc);
        CensusItem inbox = CensusTableRow.Project(Row("ID-BEFORE", "Invoice", moment, 8_192), Map())!.Value;
        CensusItem filed = CensusTableRow.Project(Row("ID-AFTER", "Invoice", moment, 8_192), Map())!.Value;

        Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> before = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Delegate"] = new Dictionary<string, FolderCensus>(StringComparer.OrdinalIgnoreCase)
            {
                ["Inbox"] = FolderCensus.WithItems(new[] { inbox }),
                ["Archive 2026"] = FolderCensus.WithItems(Array.Empty<CensusItem>()),
            },
        };
        Dictionary<string, IReadOnlyDictionary<string, FolderCensus>> after = new(StringComparer.OrdinalIgnoreCase)
        {
            ["Delegate"] = new Dictionary<string, FolderCensus>(StringComparer.OrdinalIgnoreCase)
            {
                ["Inbox"] = FolderCensus.WithItems(Array.Empty<CensusItem>()),
                ["Archive 2026"] = FolderCensus.WithItems(new[] { filed }),
            },
        };

        TripwireVerdict verdict = StoreCountTripwire.Evaluate(before, after, "Hub");

        Assert.False(verdict.Failed);
        Assert.Contains(verdict.Notes, n => n.Contains("filed (not loss)", StringComparison.Ordinal));
    }
}
