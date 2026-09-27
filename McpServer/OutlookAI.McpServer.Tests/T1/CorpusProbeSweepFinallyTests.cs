using System.Reflection;
using System.Text.RegularExpressions;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins, from the source, that the cross-store residue sweep AFTER the probes runs even when a probe
/// throws - in <c>corpus-probe</c>, <c>corpus-build</c> and <c>corpus-reanchor</c> alike.
/// <para>
/// <b>Why.</b> On OAI-UNINDEXED, 2026-09-27, <c>corpus-probe</c>'s undated probe threw, and the sweep
/// written after it as a plain statement never ran: the one check that a probe left nothing in
/// another store was skipped on exactly the run that needed it. <c>corpus-build</c> and
/// <c>corpus-reanchor</c> already had theirs in a <c>finally</c>; the probe verb now does too, and this
/// test holds all three there. No pure test can reach it - the probes need a live Outlook - so it
/// reads the source.
/// </para>
/// </summary>
public sealed class CorpusProbeSweepFinallyTests
{
    private const string AfterCall = "SweepOtherStores(options, planOptions.CorpusId, output, \"after the probes\");";
    private const string BeforeCall = "SweepOtherStores(options, planOptions.CorpusId, output, \"before the probes\");";

    [Fact]
    public void EverySweepAfterTheProbes_IsTheFirstStatementOfAFinally()
    {
        string source = Source();
        List<int> sites = Sites(source, AfterCall);

        // probe, build, reanchor: a fourth site is a new verb, and it needs the same guarantee.
        Assert.Equal(3, sites.Count);
        foreach (int site in sites)
        {
            Assert.True(
                OpensAFinally(source, site),
                "a sweep 'after the probes' is not the first statement of a finally block, so a probe that "
                + "throws skips it (line " + LineOf(source, site) + " of CorpusCommands.cs)");
        }
    }

    [Fact]
    public void TheSweepsBeforeTheProbes_AreNotInAFinally_TheControl()
    {
        // The control: the scan tells the two apart. A sweep before the probes runs first, outside any
        // try, so it is never the first statement of a finally - if the scan said it was, the test
        // above would be proving nothing.
        string source = Source();
        List<int> sites = Sites(source, BeforeCall);
        Assert.Equal(3, sites.Count);
        foreach (int site in sites)
        {
            Assert.False(OpensAFinally(source, site), "line " + LineOf(source, site));
        }
    }

    /// <summary>
    /// True when the only text between the nearest preceding <c>finally</c> keyword and the call is
    /// the block's opening brace and whitespace.
    /// </summary>
    private static bool OpensAFinally(string source, int site)
    {
        int keyword = source.LastIndexOf("finally", site, StringComparison.Ordinal);
        if (keyword < 0)
        {
            return false;
        }

        string between = source.Substring(keyword + "finally".Length, site - keyword - "finally".Length);
        return Regex.IsMatch(between, @"^\s*\{\s*$");
    }

    private static List<int> Sites(string source, string call)
    {
        var sites = new List<int>();
        for (int at = source.IndexOf(call, StringComparison.Ordinal); at >= 0;
             at = source.IndexOf(call, at + call.Length, StringComparison.Ordinal))
        {
            sites.Add(at);
        }

        return sites;
    }

    private static int LineOf(string source, int index) => source.AsSpan(0, index).Count('\n') + 1;

    private static string Source()
        => File.ReadAllText(Path.Combine(RepoRoot(), "McpServer", "OutlookAI.RemediationTools", "CorpusCommands.cs"));

    private static string RepoRoot()
    {
        string testProjectDir =
            typeof(LiveTestSettings).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
        return Path.GetFullPath(Path.Combine(testProjectDir, "..", ".."));
    }
}
