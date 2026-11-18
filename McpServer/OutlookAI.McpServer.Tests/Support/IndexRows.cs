namespace OutlookAI.McpServer.Tests;

/// <summary>
/// Rows the way an index stand-in returns them - case-insensitive column maps, as the OLE DB client
/// drains them - and the scope a statement's <c>SCOPE='...'</c> clause names. The one copy for the
/// suites' index stand-ins (Q154).
/// </summary>
internal static class IndexRows
{
    /// <summary>No rows.</summary>
    public static IReadOnlyList<IReadOnlyDictionary<string, object?>> None => Array.Empty<IReadOnlyDictionary<string, object?>>();

    /// <summary>One row of one column.</summary>
    public static IReadOnlyDictionary<string, object?> Row(string column, object? value)
    {
        return new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
        {
            [column] = value,
        };
    }

    /// <summary>Exactly one row of one column.</summary>
    public static IReadOnlyList<IReadOnlyDictionary<string, object?>> One(string column, object? value)
    {
        return new[] { Row(column, value) };
    }

    /// <summary>One <c>System.ItemUrl</c> row per URL, in order.</summary>
    public static IReadOnlyList<IReadOnlyDictionary<string, object?>> Urls(params string[] urls)
    {
        return urls.Select(u => Row("System.ItemUrl", u)).ToList();
    }

    /// <summary>A row's <paramref name="column"/> when it holds text; null otherwise.</summary>
    public static string? Text(IReadOnlyDictionary<string, object?> row, string column)
    {
        return row.TryGetValue(column, out object? value) && value is string s ? s : null;
    }

    /// <summary>The scope a statement's first <c>SCOPE='...'</c> clause names, or null when it names none.</summary>
    public static string? ScopeOf(string sql)
    {
        int start = sql.IndexOf("SCOPE='", StringComparison.Ordinal);
        if (start < 0)
        {
            return null;
        }

        start += "SCOPE='".Length;
        int end = sql.IndexOf('\'', start);
        return end < 0 ? sql.Substring(start) : sql.Substring(start, end - start);
    }
}
