using OutlookAI.Core.IndexSearch;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// The Windows Search index as the product questions it around a search, for a profile whose
/// indexed stores have the given URL prefixes: store discovery samples one mail item in each, the
/// freshness probe answers with <see cref="Frontier"/>, the reachability probe finds one item, and
/// any other statement - the search itself - returns <see cref="SearchRows"/>, up to its row limit.
/// The two probes answer only for a scope the index holds: a store's root, a folder beneath it, or
/// no scope at all while it holds a store. The one copy for the suites that drive the product's
/// search over a fabricated index (Q154).
/// <para>
/// <c>everyScope: true</c> answers both probes whatever they name, for a suite whose one store holds
/// everything it asks about. A suite whose index holds more than store prefixes - a delegate's
/// subtree - or whose probe tells a folder from a store derives, and overrides <see cref="Knows"/>,
/// <see cref="ProbeFinds"/> or <see cref="ProbedItemUrl"/>; one that answers a statement its own way
/// overrides <see cref="ExecuteRows"/> and hands the rest back to this one.
/// </para>
/// </summary>
internal class StubIndexClient : IIndexClient
{
    private const string DiscoveryTail = " System.ItemUrl FROM SystemIndex WHERE System.Kind='email'";

    private readonly IReadOnlyList<string> _knownPrefixes;

    private readonly bool _everyScope;

    internal StubIndexClient(IReadOnlyList<string> knownPrefixes, DateTime frontier, bool everyScope = false)
    {
        _knownPrefixes = knownPrefixes;
        Frontier = frontier;
        _everyScope = everyScope;
    }

    /// <summary>What the freshness probe reports as the newest item the index holds.</summary>
    internal DateTime Frontier { get; set; }

    /// <summary>What any other statement returns, cut to its row limit.</summary>
    internal IReadOnlyList<IReadOnlyDictionary<string, object?>> SearchRows { get; init; } = IndexRows.None;

    public IndexProviderKind Provider => IndexProviderKind.OleDb;

    public virtual IReadOnlyList<IReadOnlyDictionary<string, object?>> ExecuteRows(
        string sql, int maxRows, int? commandTimeoutSeconds = null)
    {
        if (IsDiscovery(sql))
        {
            return _knownPrefixes.Select(p => IndexRows.Row("System.ItemUrl", p + "/0/Inbox/sampled-item")).ToList();
        }

        if (IsFreshness(sql))
        {
            return Known(sql) ? IndexRows.One("System.Message.DateReceived", Frontier) : IndexRows.None;
        }

        if (IsProbe(sql))
        {
            return ProbeFinds(sql) ? IndexRows.One("System.ItemUrl", ProbedItemUrl(sql)) : IndexRows.None;
        }

        return SearchRows.Count <= maxRows ? SearchRows : SearchRows.Take(maxRows).ToList();
    }

    /// <summary>Store discovery: one sampled mail item per indexed store. Checked first.</summary>
    protected static bool IsDiscovery(string sql) => sql.EndsWith(DiscoveryTail, StringComparison.Ordinal);

    /// <summary>The freshness probe: the newest DateReceived in a scope. Checked second.</summary>
    protected static bool IsFreshness(string sql) => sql.Contains("System.Message.DateReceived FROM SystemIndex", StringComparison.Ordinal);

    /// <summary>The reachability probe: whether a scope holds any item. Checked third.</summary>
    protected static bool IsProbe(string sql) => sql.StartsWith("SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE", StringComparison.Ordinal);

    /// <summary>Whether a scope the statement names is one this index holds: a store's root, or a folder beneath it.</summary>
    protected virtual bool Knows(string scope)
    {
        return _knownPrefixes.Any(p =>
            string.Equals(scope, p, StringComparison.OrdinalIgnoreCase)
            || scope.StartsWith(p + "/0/", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Whether the reachability probe finds an item: where the index holds the scope it names.</summary>
    protected virtual bool ProbeFinds(string sql)
    {
        return Known(sql);
    }

    /// <summary>The item the reachability probe finds: one in the first store's Inbox.</summary>
    protected virtual string ProbedItemUrl(string sql)
    {
        return _knownPrefixes[0] + "/0/Inbox/probed-item";
    }

    private bool Known(string sql)
    {
        if (_everyScope)
        {
            return true;
        }

        string? scope = IndexRows.ScopeOf(sql);
        return scope == null ? _knownPrefixes.Count > 0 : Knows(scope);
    }
}
