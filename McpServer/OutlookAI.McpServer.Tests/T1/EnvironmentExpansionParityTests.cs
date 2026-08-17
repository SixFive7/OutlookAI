using System.Collections.Generic;

using OutlookAI.Core.Services;
using OutlookAI.Services;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The <c>${VAR}</c> / <c>${VAR:-default}</c> expansion rule exists TWICE - the add-in's
/// <c>McpConfigEditor.ExpandEnvironmentReferences</c> writes a registration with it, and
/// the server's <c>HealthReporting.ExpandEnvironmentReferences</c> reads that registration
/// back and decides whether it still names the running executable.
/// <para>
/// The duplication is deliberate: Core takes no add-in dependency. What was NOT true is the
/// claim the source made about it - "both are pinned so they cannot drift apart unnoticed".
/// Each had its own suite with its own fixture literals (<c>${LOCALAPPDATA}</c>/<c>${NOPE}</c>
/// on one side, <c>${SET}</c>/<c>${UNSET}</c> on the other) and NOTHING ever ran one input
/// through both. A divergence would have shown up as a healthy registration reported as
/// drifted, or a drifted one reported as healthy - and both suites would have stayed green.
/// </para>
/// <para>
/// This is that missing test. One corpus, both implementations, character-for-character
/// equality. The tests project already compiles the add-in file (see the csproj), so it
/// costs one assertion and it pins the SHIPPED code on both sides rather than a
/// re-implementation of either.
/// </para>
/// </summary>
public sealed class EnvironmentExpansionParityTests
{
    /// <summary>
    /// The variable table both sides see. Deliberately includes an empty-valued variable:
    /// "set but empty" is the case where a naive implementation would use the value and a
    /// correct one falls back, and it is exactly the kind of edge two separate authors
    /// resolve differently.
    /// </summary>
    private static readonly Dictionary<string, string> Variables = new Dictionary<string, string>(System.StringComparer.Ordinal)
    {
        ["LOCALAPPDATA"] = @"C:\Users\tester\AppData\Local",
        ["SET"] = "value",
        ["EMPTY"] = "",
        ["WITH SPACES"] = "spaced",
        ["NESTED"] = "${SET}",
    };

    /// <summary>
    /// Every shape the rule has an opinion about, in one list, so a new edge case is added
    /// once and both implementations are held to it.
    /// </summary>
    public static TheoryData<string> Corpus()
    {
        return new TheoryData<string>
        {
            "",
            "no references at all",
            @"C:\Program Files\OutlookAI\OutlookAI.McpServer.exe",
            "${LOCALAPPDATA}/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe",
            "${SET}",
            "${UNSET}",
            "${EMPTY}",
            "${SET}${UNSET}${SET}",
            "prefix ${SET} suffix",
            "${SET:-fallback}",
            "${UNSET:-fallback}",
            "${EMPTY:-fallback}",
            "${UNSET:-}",
            "${UNSET:-with spaces and \\ backslash}",
            "${UNSET:-${SET}}",
            "${NESTED}",
            "${}",
            "${:-fallback}",
            "${unterminated",
            "trailing $",
            "$SET",
            "$$SET",
            "${WITH SPACES}",
            "${SET:-a:-b}",
            "}{",
            "${SET}}",
            "%LOCALAPPDATA%",
        };
    }

    /// <summary>
    /// The writing side and the reading side agree, exactly, on every input.
    /// <para>
    /// Break either implementation and this fails naming the input and both answers, which
    /// is the diagnostic the two independent suites could never produce.
    /// </para>
    /// </summary>
    [Theory]
    [MemberData(nameof(Corpus))]
    public void BothImplementations_ExpandIdentically(string input)
    {
        // The two lookups differ in NULLABILITY by design - the add-in's contract is "" for
        // unset, Core's is null - so the adapters here are the only difference the corpus is
        // allowed to see. Everything past this point must agree.
        string addIn = McpConfigEditor.ExpandEnvironmentReferences(
            input, name => Variables.TryGetValue(name, out string? v) ? v : "");
        string server = HealthReporting.ExpandEnvironmentReferences(
            input, name => Variables.TryGetValue(name, out string? v) ? v : null);

        Assert.True(
            string.Equals(addIn, server, System.StringComparison.Ordinal),
            $"the two ExpandEnvironmentReferences implementations disagree on \"{input}\": "
            + $"add-in produced \"{addIn}\", server produced \"{server}\". They are the write and read halves of one "
            + "registration check, so a divergence silently reports a healthy registration as drifted or the reverse.");
    }

    /// <summary>
    /// Guard against the corpus going vacuous: a MemberData source that silently returned
    /// nothing would leave the parity claim proven by zero cases.
    /// </summary>
    [Fact]
    public void Corpus_CoversEveryShapeTheRuleHasAnOpinionAbout()
    {
        Assert.True(Corpus().Count >= 20, $"the parity corpus must stay broad; it holds {Corpus().Count} inputs");
    }

    /// <summary>
    /// The portable command the add-in actually registers, expanded by the SERVER's reader,
    /// resolves to the same string the WRITER produced. This is the real-world instance of
    /// the parity above, using the shipped constant rather than a fixture spelling of it.
    /// </summary>
    [Fact]
    public void ThePortableInstalledCommand_RoundTripsThroughBothSides()
    {
        string addIn = McpConfigEditor.ExpandEnvironmentReferences(
            McpConfigEditor.PortableInstalledCommand,
            name => Variables.TryGetValue(name, out string? v) ? v : "");
        string server = HealthReporting.ExpandEnvironmentReferences(
            McpConfigEditor.PortableInstalledCommand,
            name => Variables.TryGetValue(name, out string? v) ? v : null);

        Assert.Equal(addIn, server);
        Assert.DoesNotContain("${", server, System.StringComparison.Ordinal);
    }
}
