namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Every xunit collection whose fixture arms the live-tier guards, named once so the
/// <c>[CollectionDefinition]</c> attribute and the code that reasons about collections
/// cannot drift apart.
/// <para>
/// WHY THIS EXISTS RATHER THAN A STRING LITERAL PER FIXTURE. The count tripwire has to
/// know when the run is OVER, and the only vantage point that knows which collections a
/// run contains is <see cref="SuiteCollectionOrderer"/>, which is handed the list before
/// anything is constructed. Comparing that list against a set of names means the two
/// sides must agree about the names, and a collection whose name is a literal in two
/// places is a collection whose verification can be lost to a typo. A constant cannot be
/// mistyped in only one of them.
/// </para>
/// <para>
/// <see cref="All"/> is pinned in CI against the collections the assembly actually
/// declares, so a live collection added later cannot quietly fall outside the guard: the
/// pin fails until the new collection is registered here.
/// </para>
/// </summary>
public static class LiveCollections
{
    /// <summary>Index-search discovery, decode-verify, recall and the completeness oracle.</summary>
    public const string Phase1 = "LivePhase1";

    /// <summary>Search/read round trips, health, freshness and the sweep cache.</summary>
    public const string Phase2 = "LivePhase2";

    /// <summary>Exhaustive and resumable scans, show-me, the UI search backend, the sort probe.</summary>
    public const string Phase3 = "LivePhase3";

    /// <summary>Draft creation, HTML drafts, signatures on drafts, update/discard.</summary>
    public const string Phase4 = "LivePhase4";

    /// <summary>The send path: tokens, expiry, fail-closed refusals.</summary>
    public const string Phase5 = "LivePhase5";

    /// <summary>Move/archive, folder scoping and sweep scoping (needs the hub's Archive folder).</summary>
    public const string MoveArchive = "LiveMoveArchive";

    /// <summary>manage_signature over COM and over stdio.</summary>
    public const string SignatureManage = "LiveSignatureManage";

    /// <summary>
    /// The stdio MCP shape tests that own no fixture state of their own. They spawn the
    /// built server exe per test and had, until this collection existed, no census, no
    /// preflight and no tripwire at all - three live classes running outside every guard.
    /// </summary>
    public const string McpToolShape = "LiveMcpToolShape";

    /// <summary>
    /// Outlook lifecycle: the disconnect/reattach proof and the headless guarantee. Forced
    /// LAST by <see cref="SuiteCollectionOrderer"/> because it bounces Outlook.
    /// </summary>
    public const string Lifecycle = "LiveLifecycle";

    /// <summary>
    /// Every guarded collection. Order is irrelevant here - the RUN's order is what decides
    /// which one verifies, and that comes from the collection orderer.
    /// </summary>
    public static IReadOnlyList<string> All { get; } = new[]
    {
        Phase1,
        Phase2,
        Phase3,
        Phase4,
        Phase5,
        MoveArchive,
        SignatureManage,
        McpToolShape,
        Lifecycle,
    };

    /// <summary>True when <paramref name="collectionName"/> names a guarded live collection.</summary>
    public static bool IsGuarded(string? collectionName)
    {
        return collectionName != null
            && All.Any(name => string.Equals(name, collectionName, StringComparison.Ordinal));
    }
}
