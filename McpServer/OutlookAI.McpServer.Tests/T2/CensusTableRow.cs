using System.Globalization;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Where the four census columns sit in a folder table's row, or -1 when the column is not
/// on the table at all.
/// <para>
/// Resolved ONCE per folder rather than per row: <c>Table.Columns</c> is a COM collection,
/// so naming a column costs a round trip per column per lookup, and the census's whole
/// reason for existing after 2026-08-20 is that round trips against an Exchange store are
/// the expensive thing.
/// </para>
/// </summary>
public readonly struct CensusColumnMap
{
    /// <summary>Builds a map. Every index is a zero-based position in a row's values, or -1.</summary>
    public CensusColumnMap(int id, int subject, int received, int size, int columnCount)
    {
        Id = id;
        Subject = subject;
        Received = received;
        Size = size;
        ColumnCount = columnCount;
    }

    /// <summary>Position of the EntryID column. A row without one identifies nothing.</summary>
    public int Id { get; }

    /// <summary>
    /// Position of the subject column. Read to set one boolean and never kept, but the
    /// column must be PRESENT: without it every item would silently read as untagged, and
    /// the attribution line would then say "undecidable" over a removal the suite itself
    /// caused. That is the one wrong answer this guard must not give quietly.
    /// </summary>
    public int Subject { get; }

    /// <summary>Position of the received-time column, half of the move-stable fingerprint.</summary>
    public int Received { get; }

    /// <summary>Position of the message-size column, the other half.</summary>
    public int Size { get; }

    /// <summary>How many columns the table carries. Used to check a bulk read's shape.</summary>
    public int ColumnCount { get; }

    /// <summary>
    /// True when this table can answer everything the census needs. A table missing any of
    /// the four degrades the FOLDER to a count rather than producing a weaker identity
    /// reading: identity without a fingerprint cannot prove a filing, so it would turn a
    /// person filing mail during a run into a suite failure, and identity without a subject
    /// cannot attribute one.
    /// </summary>
    public bool IsUsable => Id >= 0 && Subject >= 0 && Received >= 0 && Size >= 0;
}

/// <summary>
/// Turns one row of a folder's <c>Table</c> into a <see cref="CensusItem"/>.
/// <para>
/// Pure by design, and separate from the COM walk for that reason: the projection decides
/// what the tripwire compares (an opaque id, a move-stable fingerprint, a tag flag) and
/// what it discards (the subject), and neither of those may be settled by a rule that only
/// a live mailbox can execute.
/// </para>
/// <para>
/// S3/S4: a subject arrives here, sets a boolean and is dropped. Nothing built here carries
/// a subject, a sender or a body, so nothing another store's owner would recognise can
/// reach a census, a log or a failure message.
/// </para>
/// </summary>
public static class CensusTableRow
{
    /// <summary>
    /// Spellings tried for the received-time column, in order. Two of them because
    /// Microsoft documents the namespace form as valid for <c>Columns.Add</c> while the
    /// explicit name is what the object model calls the property, and this repository has
    /// already been bitten once by assuming a single spelling works everywhere
    /// (<c>OutlookComSession.DateSortProperties</c> carries the same pair for the same
    /// reason).
    /// </summary>
    public static IReadOnlyList<string> ReceivedColumnNames { get; } =
        new[] { "ReceivedTime", "urn:schemas:httpmail:datereceived" };

    /// <summary>
    /// Spellings tried for the message size. The proptag form is the fallback because
    /// <c>Size</c> is a computed property on the object model and a table may expose only
    /// the underlying MAPI property (PR_MESSAGE_SIZE, 0x0E08, type 3 = long).
    /// </summary>
    public static IReadOnlyList<string> SizeColumnNames { get; } =
        new[] { "Size", "http://schemas.microsoft.com/mapi/proptag/0x0E080003" };

    /// <summary>Spellings tried for the EntryID column. It is on every table by default.</summary>
    public static IReadOnlyList<string> IdColumnNames { get; } = new[] { "EntryID" };

    /// <summary>Spellings tried for the subject column. Also a default column.</summary>
    public static IReadOnlyList<string> SubjectColumnNames { get; } =
        new[] { "Subject", "urn:schemas:httpmail:subject" };

    /// <summary>
    /// Works out where the census columns sit from the table's own column NAMES, read back
    /// after the additions rather than assumed from them: <c>Columns.Add</c> can accept a
    /// spelling the rows then do not carry, and the names are the only account of the table
    /// that cannot disagree with the rows.
    /// </summary>
    public static CensusColumnMap MapColumns(IReadOnlyList<string> names)
    {
        ArgumentNullException.ThrowIfNull(names);
        return new CensusColumnMap(
            IndexOfColumn(names, IdColumnNames),
            IndexOfColumn(names, SubjectColumnNames),
            IndexOfColumn(names, ReceivedColumnNames),
            IndexOfColumn(names, SizeColumnNames),
            names.Count);
    }

    /// <summary>
    /// Accepts one bulk-read block only when its shape is EXACTLY the rows-by-columns block
    /// that was asked for.
    /// <para>
    /// A two-dimensional variant array carries no labels, so nothing in it distinguishes it
    /// from its own transpose, and reading it the wrong way round would invent items rather
    /// than fail. Requiring both dimensions to equal known numbers - the rows requested and
    /// the columns the table reported - leaves a transpose acceptable only when a folder
    /// happens to have exactly as many rows left as the table has columns, and the EntryID
    /// and duplicate checks in <see cref="ProjectRows"/> then reject that.
    /// </para>
    /// <para>
    /// A short block is a rejection too, not a partial answer: the folder promised a count
    /// and delivered fewer rows, so the reading spans two moments.
    /// </para>
    /// </summary>
    public static bool TryReadBlock(object? batch, int wantedRows, int columnCount, out Array? rows)
    {
        rows = batch as Array;
        if (rows == null
            || rows.Rank != 2
            || rows.GetLength(0) != wantedRows
            || rows.GetLength(1) != columnCount)
        {
            rows = null;
            return false;
        }

        return true;
    }

    /// <summary>
    /// Copies one bulk-read block into <paramref name="walked"/>, or returns false when any
    /// row fails to identify an item. A row the census cannot name is indistinguishable from
    /// an item that was deleted, so one bad row abandons the whole folder.
    /// <para>
    /// The lower bounds are read from the array rather than assumed to be zero, because a
    /// SAFEARRAY may declare any origin and an off-by-one origin would shift every row.
    /// </para>
    /// </summary>
    public static bool ProjectRows(Array rows, CensusColumnMap columns, List<CensusItem> walked, HashSet<string> seen)
    {
        ArgumentNullException.ThrowIfNull(rows);
        ArgumentNullException.ThrowIfNull(walked);
        ArgumentNullException.ThrowIfNull(seen);
        if (rows.Rank != 2)
        {
            return false;
        }

        int rowBase = rows.GetLowerBound(0);
        int columnBase = rows.GetLowerBound(1);
        int rowCount = rows.GetLength(0);
        int columnCount = rows.GetLength(1);
        object?[] values = new object?[columnCount];
        for (int r = 0; r < rowCount; r++)
        {
            for (int c = 0; c < columnCount; c++)
            {
                values[c] = rows.GetValue(rowBase + r, columnBase + c);
            }

            CensusItem? item = Project(values, columns);
            if (item == null || !seen.Add(item.Value.Id))
            {
                // A row with no id, or the same item twice: the reading shifted under us.
                return false;
            }

            walked.Add(item.Value);
        }

        return true;
    }

    /// <summary>
    /// Projects one row, or returns null when the row names no item.
    /// <para>
    /// Null is a hard signal, not a skip: a row inside a folder the census is walking that
    /// cannot say WHICH item it is makes the whole walk unusable, because an item the
    /// census failed to record reads exactly like an item that was deleted. The caller
    /// degrades the folder to a count.
    /// </para>
    /// </summary>
    /// <param name="values">One row's values, in table column order.</param>
    /// <param name="columns">Where the census columns sit. Must be <see cref="CensusColumnMap.IsUsable"/>.</param>
    public static CensusItem? Project(object?[] values, CensusColumnMap columns)
    {
        ArgumentNullException.ThrowIfNull(values);
        if (!columns.IsUsable || columns.Id >= values.Length || columns.Subject >= values.Length)
        {
            return null;
        }

        if (values[columns.Id] is not string id || id.Length == 0)
        {
            return null;
        }

        string? subject = values[columns.Subject] as string;
        bool tagged = subject != null
            && subject.IndexOf(LiveOutlookTestMailer.SubjectTag, StringComparison.OrdinalIgnoreCase) >= 0;

        return new CensusItem(id, Fingerprint(Value(values, columns.Received), Value(values, columns.Size)), tagged);
    }

    /// <summary>
    /// The move-stable key: a received instant and a byte size, both metadata. Null when
    /// either is missing, which is the ordinary reading for a draft or a report item (no
    /// received time) and means only that the item cannot be traced across a move.
    /// </summary>
    public static string? Fingerprint(object? received, object? size)
    {
        DateTime? utc = ReadUtc(received);
        long? bytes = ReadSize(size);
        if (utc == null || bytes == null)
        {
            return null;
        }

        return utc.Value.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture)
            + "/" + bytes.Value.ToString(CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// A table's date value as UTC.
    /// <para>
    /// An UNSPECIFIED kind is taken as already-UTC. That is what Microsoft documents for the
    /// <c>Table</c> object (it returns date-time values in UTC, unlike the object model,
    /// which returns local time) and it is the contract the rest of this solution already
    /// follows for an unspecified kind (<c>DaslDateLiteral.FormatUtc</c>). If that reading
    /// were ever wrong the tripwire's DECISIONS would not change - every value in every
    /// census comes through this one method, so two censuses still agree with each other -
    /// and only the instant PRINTED beside a departed item would be offset by the machine's
    /// UTC offset.
    /// </para>
    /// </summary>
    private static DateTime? ReadUtc(object? value)
    {
        if (value is not DateTime moment)
        {
            return null;
        }

        return moment.Kind switch
        {
            DateTimeKind.Utc => moment,
            DateTimeKind.Local => moment.ToUniversalTime(),
            _ => DateTime.SpecifyKind(moment, DateTimeKind.Utc),
        };
    }

    /// <summary>
    /// A table's size value as a byte count. Written out rather than handed to
    /// <c>Convert.ToInt64</c> because the variant a table hands back for a MAPI long is not
    /// contractually one CLR type, and a culture-sensitive conversion has no business
    /// deciding whether two censuses agree.
    /// </summary>
    private static long? ReadSize(object? value)
    {
        return value switch
        {
            int i => i,
            long l => l,
            short s => s,
            uint u => u,
            double d => (long)d,
            string text when long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long parsed)
                => parsed,
            _ => null,
        };
    }

    private static object? Value(object?[] values, int index)
    {
        return index >= 0 && index < values.Length ? values[index] : null;
    }

    /// <summary>
    /// Position of the FIRST accepted spelling among the table's column names, or -1. Order
    /// matters: the explicit name is asked for before the namespace form, so a table that
    /// carries both is read through the spelling the object model itself uses.
    /// </summary>
    private static int IndexOfColumn(IReadOnlyList<string> names, IReadOnlyList<string> spellings)
    {
        foreach (string spelling in spellings)
        {
            for (int i = 0; i < names.Count; i++)
            {
                if (string.Equals(names[i], spelling, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }
        }

        return -1;
    }
}
