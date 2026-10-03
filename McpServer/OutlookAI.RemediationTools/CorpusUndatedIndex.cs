using System.Globalization;
using System.Text;
using OutlookAI.Core.IndexSearch;

namespace OutlookAI.RemediationTools;

/// <summary>
/// What the Windows Search index makes of an UNDATED item - an appointment, a contact or a task -
/// while the undated probe holds it (<c>corpus-probe --undated-index-wait N</c>). Q98 (f), asked
/// 2026-10-03: the populations' undated rows exist for <c>LiveOrderKeyCollationTests</c>, which
/// assumes the index gives them no received date; the undated probe had already found that a
/// PST gives all three a DELIVERY TIME on their first save, so what is left to measure is the
/// index's own column for it, and every other column the product sorts on.
/// <para>
/// READ-ONLY and pure apart from the one client it is handed: it never writes, and the item it
/// looks for is created and deleted by the probe, under the probe's own two-key deletion rule.
/// Rows are found by the probe's subject word and matched in code against the probe's corpus tag
/// AND its kind - never by a name the index might spell differently from Outlook - so a row from
/// any other item, or from a previous kind's item the index has not yet dropped, cannot match.
/// </para>
/// </summary>
public static class CorpusUndatedIndex
{
    /// <summary>
    /// The columns printed for each row: the URL first (it is what proves the row is an Outlook
    /// item), the three the row is matched on, then the product's ORDER BY keys
    /// (<c>System.Message.DateReceived</c>, the default; <c>System.Size</c>) and every other date
    /// the index can carry for an Outlook item.
    /// </summary>
    public static readonly IReadOnlyList<string> Columns = new[]
    {
        "System.ItemUrl",
        "System.Subject",
        "System.ItemNameDisplay",
        "System.Contact.FullName",
        "System.Kind",
        "System.ItemType",
        "System.Message.DateReceived",
        "System.Size",
        "System.Message.DateSent",
        "System.DateModified",
        "System.DateCreated",
        "System.ItemDate",
        "System.StartDate",
        "System.EndDate",
        "System.DueDate",
        "System.ItemFolderPathDisplay",
        "System.Search.GatherTime",
    };

    /// <summary>The columns whose text the row is matched on.</summary>
    private static readonly string[] MatchColumns = { "System.Subject", "System.ItemNameDisplay", "System.Contact.FullName" };

    /// <summary>How often the reader asks the index again while it waits.</summary>
    public static readonly TimeSpan PollInterval = TimeSpan.FromSeconds(3);

    /// <summary>
    /// The statements the reader runs, each on its own so that one the provider refuses does not
    /// hide the other. A contact's subject may reach the index only as its name, hence two.
    /// </summary>
    public static IReadOnlyList<string> Statements()
    {
        string select = "SELECT TOP 100 " + string.Join(", ", Columns) + " FROM SystemIndex WHERE ";
        return new[]
        {
            select + "CONTAINS(System.Subject, '\"undated\"')",
            select + "CONTAINS(System.ItemNameDisplay, '\"undated\"')",
            select + "CONTAINS(System.Contact.FullName, '\"undated\"')",
        };
    }

    /// <summary>
    /// The probe tag inside <paramref name="subject"/> - <c>[OutlookAI-Corpus:&lt;id&gt;#&lt;ordinal&gt;]</c> -
    /// or null when it carries none.
    /// </summary>
    public static string? ProbeTag(string? subject)
    {
        if (string.IsNullOrEmpty(subject))
        {
            return null;
        }

        int open = subject.IndexOf(CorpusPlan.CorpusTagOpen, StringComparison.Ordinal);
        if (open < 0)
        {
            return null;
        }

        int close = subject.IndexOf(']', open);
        return close < 0 ? null : subject.Substring(open, close - open + 1);
    }

    /// <summary>
    /// True when <paramref name="row"/> is the index's row for the probe item of
    /// <paramref name="kind"/> that carries <paramref name="subject"/>: an Outlook URL, and a match
    /// column that holds the probe tag, the word <c>undated</c> and the kind's own word. The words
    /// are checked one by one rather than as a phrase, because a contact's name reaches the index
    /// re-ordered (file-as order) while the tag, which has no space in it, survives whole.
    /// </summary>
    public static bool IsRowFor(IReadOnlyDictionary<string, object?> row, string subject, CorpusItemKind kind)
    {
        ArgumentNullException.ThrowIfNull(row);
        string? tag = ProbeTag(subject);
        if (tag == null)
        {
            return false;
        }

        if (!row.TryGetValue("System.ItemUrl", out object? url) || url is not string text
            || !text.StartsWith("mapi", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        string kindWord = kind.ToString().ToLowerInvariant();
        foreach (string column in MatchColumns)
        {
            if (row.TryGetValue(column, out object? value) && value is string s
                && s.Contains(tag, StringComparison.Ordinal)
                && s.Contains("undated", StringComparison.OrdinalIgnoreCase)
                && s.Contains(kindWord, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// One row, every column on its own line, a missing value as <c>&lt;null&gt;</c>, a date in the
    /// invariant <c>yyyy-MM-dd HH:mm:ss</c> exactly as the provider returned it, and the encoded
    /// EntryID in the URL shortened to <c>&lt;id&gt;</c> (an id has no business in a log).
    /// </summary>
    public static string Describe(CorpusItemKind kind, IReadOnlyDictionary<string, object?> row, TimeSpan after)
    {
        ArgumentNullException.ThrowIfNull(row);
        var sb = new StringBuilder();
        sb.Append("  ").Append(kind.ToString().ToLowerInvariant()).Append(": in the index after ")
            .Append(((int)after.TotalSeconds).ToString(CultureInfo.InvariantCulture)).Append(" s");
        foreach (string column in Columns)
        {
            row.TryGetValue(column, out object? value);
            sb.AppendLine().Append("      ").Append(column).Append(" = ").Append(Format(column, value));
        }

        return sb.ToString();
    }

    private static string Format(string column, object? value)
    {
        switch (value)
        {
            case null:
            case DBNull:
                return "<null>";
            case DateTime dt:
                return dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            case string s when s.Length == 0:
                return "<empty>";
            case string s when column == "System.ItemUrl":
                return ShortenEncodedIds(s);
            case string s:
                return s;
            case System.Collections.IEnumerable many:
                var parts = new List<string>();
                foreach (object? part in many)
                {
                    parts.Add(Format(column, part));
                }

                return string.Join("|", parts);
            case IFormattable f:
                return f.ToString(null, CultureInfo.InvariantCulture);
            default:
                return value.ToString() ?? "<null>";
        }
    }

    private static string ShortenEncodedIds(string url)
    {
        var sb = new StringBuilder(url.Length);
        bool inId = false;
        foreach (char c in url)
        {
            bool encoded = c >= '가' && c <= '곿';
            if (encoded && !inId)
            {
                sb.Append("<id>");
            }
            else if (!encoded)
            {
                sb.Append(c);
            }

            inId = encoded;
        }

        return sb.ToString();
    }

    /// <summary>
    /// The observer <see cref="ComCorpusMailbox.ProbeUndated"/> calls while it holds each item:
    /// asks the index every <see cref="PollInterval"/> until a row matches (<see cref="IsRowFor"/>)
    /// or <paramref name="wait"/> runs out, and writes the row - or that there was none - to
    /// <paramref name="output"/>. It never throws: a refused statement or an unreachable index is
    /// written as such, because the probe must go on to delete the item it is holding.
    /// </summary>
    /// <param name="client">The index client; null opens the default one on first use.</param>
    /// <param name="elapsed">A monotonic clock reading for one wait (tests inject one).</param>
    /// <param name="sleep">How a poll interval passes (tests inject one).</param>
    /// <param name="written">
    /// The delivery time the probe WROTE on an item of this kind, or null - an all-kinds appointment or
    /// task (D62 (b), 2026-10-03). When there is one the reader does not stop at the first row: the index
    /// takes the item's FIRST save first, dated at its creation, so it keeps asking until the row's
    /// <c>System.Message.DateReceived</c> - read the way the product reads it - is the written instant, and
    /// says whether it got there and when. That is the measurement D62 (b) rests on: whether the index
    /// dates these kinds by their delivery time, which the plan can choose, or by their creation, which it
    /// cannot.
    /// </param>
    public static Action<CorpusItemKind, string> CreateReader(
        TextWriter output,
        TimeSpan wait,
        IIndexClient? client = null,
        Func<Func<TimeSpan>>? elapsed = null,
        Action<TimeSpan>? sleep = null,
        Func<CorpusItemKind, DateTime?>? written = null)
    {
        ArgumentNullException.ThrowIfNull(output);
        Func<Func<TimeSpan>> startClock = elapsed ?? (() =>
        {
            System.Diagnostics.Stopwatch sw = System.Diagnostics.Stopwatch.StartNew();
            return () => sw.Elapsed;
        });
        Action<TimeSpan> pause = sleep ?? Thread.Sleep;
        IIndexClient? index = client;
        var refused = new HashSet<string>(StringComparer.Ordinal);

        return (kind, subject) =>
        {
            try
            {
                if (index == null)
                {
                    index = IndexClientFactory.CreateAuto(out string report);
                    output.WriteLine("  index: " + report);
                }

                Func<TimeSpan> clock = startClock();
                DateTime? expected = written?.Invoke(kind);
                IReadOnlyDictionary<string, object?>? firstSeen = null;
                while (true)
                {
                    foreach (string sql in Statements())
                    {
                        IReadOnlyList<IReadOnlyDictionary<string, object?>> rows;
                        try
                        {
                            rows = index.ExecuteRows(sql, 100);
                        }
                        catch (Exception ex) when (ex is not OutOfMemoryException)
                        {
                            if (refused.Add(sql))
                            {
                                output.WriteLine("  index refused a statement (" + ex.GetType().Name + ": " + ex.Message + "): " + sql);
                            }

                            continue;
                        }

                        IReadOnlyDictionary<string, object?>? match = rows.FirstOrDefault(r => IsRowFor(r, subject, kind));
                        if (match != null)
                        {
                            if (expected == null)
                            {
                                output.WriteLine(Describe(kind, match, clock()));
                                return;
                            }

                            if (firstSeen == null)
                            {
                                firstSeen = match;
                                output.WriteLine(Describe(kind, match, clock()));
                            }

                            DateTime? indexed = IndexRowMapper.Map(match).DateReceivedUtc;
                            if (indexed != null && Math.Abs((indexed.Value - expected.Value).TotalSeconds) <= 2)
                            {
                                output.WriteLine("  " + kind.ToString().ToLowerInvariant() + ": DATED AS WRITTEN after "
                                    + ((int)clock().TotalSeconds).ToString(CultureInfo.InvariantCulture) + " s - System.Message.DateReceived "
                                    + indexed.Value.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture)
                                    + " is the delivery time the probe wrote, so the index dates this kind by its delivery time");
                                return;
                            }

                            break;
                        }
                    }

                    TimeSpan left = wait - clock();
                    if (left <= TimeSpan.Zero)
                    {
                        if (firstSeen != null && expected != null)
                        {
                            DateTime? last = IndexRowMapper.Map(firstSeen).DateReceivedUtc;
                            output.WriteLine("  " + kind.ToString().ToLowerInvariant() + ": NOT DATED AS WRITTEN after "
                                + ((int)wait.TotalSeconds).ToString(CultureInfo.InvariantCulture) + " s - the probe wrote "
                                + expected.Value.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture)
                                + " and the index still says "
                                + (last == null ? "<null>" : last.Value.ToString("yyyy-MM-ddTHH:mm:ssZ", CultureInfo.InvariantCulture))
                                + " (its first row); it may date this kind by something other than the delivery time");
                            return;
                        }

                        output.WriteLine("  " + kind.ToString().ToLowerInvariant() + ": NOT in the index after "
                            + ((int)wait.TotalSeconds).ToString(CultureInfo.InvariantCulture)
                            + " s - is a NOT elevated Outlook running on the profile that mounts this store?");
                        return;
                    }

                    pause(left < PollInterval ? left : PollInterval);
                }
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                output.WriteLine("  " + kind.ToString().ToLowerInvariant() + ": the index could not be read ("
                    + ex.GetType().Name + ": " + ex.Message + ")");
            }
        };
    }
}
