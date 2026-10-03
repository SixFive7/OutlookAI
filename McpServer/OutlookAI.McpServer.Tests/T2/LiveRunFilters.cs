namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// The <c>dotnet test --filter</c> expressions that choose which live tests run on which machine,
/// DERIVED from the trait vocabulary rather than typed - the one place they are spelled.
/// <para>
/// <b>Why derived.</b> A hand-kept filter string is exactly what this repository's history shows
/// drifting, and the filter is the thing that keeps a test off a machine it must not run on. Every
/// copy a person reads - the runbook, <c>Testbed/README.md</c>, the guest's hub-rebuild script, the
/// opt-in refusal - is pinned by <c>T1.LiveTierInventoryTests</c> to equal what is computed here, and
/// what is computed here is pinned to the vocabulary <c>T1.LiveTierInventoryTests</c> holds.
/// </para>
/// </summary>
public static class LiveRunFilters
{
    /// <summary>A delegate/shared mailbox: its index namespace drops every intermediate folder.</summary>
    public const string DelegateStore = "DelegateStore";

    /// <summary>
    /// A cached Exchange mailbox, whose object model hands out 70-byte Exchange entry ids (Q74 C1).
    /// A PST cannot be made to behave like one, so - like <see cref="DelegateStore"/> - no test guest
    /// can be given it.
    /// </summary>
    public const string CachedExchange = "CachedExchange";

    /// <summary>A populated Windows Search index - the capability the unindexed guest lacks.</summary>
    public const string SearchIndex = "SearchIndex";

    /// <summary>
    /// The trait a live test carries, PER METHOD, to say it changes nothing on the machine it runs on
    /// (Q74 layer 1): no mail item or folder, no signature, no registry value, nothing on the user's
    /// screen, through any path - its own body, its class's and its fixtures' constructors and teardown,
    /// or a tool it asks the MCP server for. <c>T1.ReadOnlyLiveTestTests</c> proves every carrier
    /// statically, as far as that can be proven, and fails the build for one that can reach a write.
    /// </summary>
    public const string WritesTrait = "Writes";

    /// <summary>The trait's one value. Absence means "may write" - the direction that fails safe.</summary>
    public const string WritesNone = "None";

    /// <summary>
    /// The capabilities only an Exchange profile has - the whole definition of the tests that cannot
    /// leave the maintainer's workstation. Since Q72 they are the only live tests that run there at all.
    /// </summary>
    public static IReadOnlyList<string> WorkstationOnlyCapabilities { get; } = new[] { DelegateStore, CachedExchange };

    /// <summary>
    /// The filter a test guest runs: every live test except the ones naming a capability no guest can
    /// be given. <c>Requires!=X</c> means no value of <c>Requires</c> equals X, which is what makes a
    /// multi-valued trait usable as an exclusion.
    /// </summary>
    public static string Guest { get; } =
        "Category=Live" + string.Concat(WorkstationOnlyCapabilities.Select(c => "&Requires!=" + c));

    /// <summary>The filter on the guest whose search index is switched off by design.</summary>
    public static string GuestUnindexed { get; } = Guest + "&Requires!=" + SearchIndex;

    /// <summary>
    /// The filter the maintainer's workstation runs, and the only one it may (Q72, Q74): the live tests
    /// that need an Exchange profile - nothing else may run there - AND that carry
    /// <see cref="WritesTrait"/>=<see cref="WritesNone"/>. Every such test must carry it
    /// (<c>T1.LiveTierInventoryTests</c>), so in practice the two halves select the same tests; both are
    /// kept so that neither alone decides.
    /// <para>
    /// <b>A VSTest property worth knowing.</b> A clause on a trait KEY that no test in the assembly
    /// carries is not evaluated at all - it matches everything (measured 2026-10-03:
    /// <c>Category=Live&amp;Bogus=None</c> lists every live test). So a misspelt key does not narrow a run,
    /// it silently stops narrowing it, which is one more reason this string is derived rather than typed,
    /// and why <c>T1.LiveTierInventoryTests</c> requires every key used here to be carried by a live test.
    /// </para>
    /// </summary>
    public static string Workstation { get; } =
        "Category=Live&" + WritesTrait + "=" + WritesNone + "&("
        + string.Join("|", WorkstationOnlyCapabilities.Select(c => "Requires=" + c)) + ")";
}
