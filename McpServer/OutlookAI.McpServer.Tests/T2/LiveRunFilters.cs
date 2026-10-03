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
    public const string WritesNothing = "Nothing";

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
    /// <see cref="WritesTrait"/>=<see cref="WritesNothing"/>. Every such test must carry it
    /// (<c>T1.LiveTierInventoryTests</c>), so in practice the two halves select the same tests; both are
    /// kept so that neither alone decides.
    /// <para>
    /// <b>Why the value is <see cref="WritesNothing"/> and never <c>None</c> - measured 2026-10-03.</b>
    /// To the VSTest filter a test that does not carry a trait has the value <c>None</c> for it:
    /// <c>Requires=None</c> lists all 3,144 tests that carry no <c>Requires</c>, and <c>Bogus=None</c> lists
    /// every test. A positive clause <c>Writes=Nothing</c> therefore selects every test that never declared
    /// anything - the exact opposite of its meaning - and the trait shipped that way for one commit before
    /// a <c>--list-tests</c> count showed it. A key or value spelt wrong otherwise selects NOTHING (measured:
    /// <c>Category=Live&amp;Bogus=Thing</c> lists none), which is the safe direction: an empty run.
    /// <c>T1.LiveTierInventoryTests</c> refuses <c>None</c> as a trait value and in any derived filter.
    /// </para>
    /// </summary>
    public static string Workstation { get; } =
        "Category=Live&" + WritesTrait + "=" + WritesNothing + "&("
        + string.Join("|", WorkstationOnlyCapabilities.Select(c => "Requires=" + c)) + ")";
}
