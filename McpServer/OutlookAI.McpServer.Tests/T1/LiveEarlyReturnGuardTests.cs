using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using OutlookAI.McpServer.Tests.T2;
using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Pins the maintainer's decision Q76 (2026-09-27): three live test classes that could report
/// GREEN having proved nothing now say so, and cannot quietly go back.
/// <para>
/// <b>The defect.</b> Every one of these printed a line and returned - on every machine, so a run
/// that exercised nothing read exactly like one that exercised everything:
/// </para>
/// <list type="bullet">
/// <item><c>T2/LiveResumableScanTests</c> - FOUR early returns across three tests (first counted as
/// three: the superseded-token test has two), on a test hub whose exhaustive scan fits one page of
/// two, plus the acceptance itself, which on such a hub "paged" through one page and compared two
/// sets that agree by construction.</item>
/// <item><c>T2/LiveStaleIndexRowTests</c> - the return taken when the delegate folder tree does not
/// list the probe folder nested right now, so the locator assertion never ran.</item>
/// <item><c>T2/LiveDisconnectRecoveryTests</c> - the <c>SKIP:</c> taken when the Outbox is not
/// provably empty.</item>
/// </list>
/// <para>
/// <b>The fix is the repository's own idiom, not a new one.</b> The first two obtain their
/// population through <see cref="LivePopulationCoverage.Require"/> - a refusal on a Production
/// profile, a <c>PROVED NOTHING:</c> line on a Portable one, the answer decision 57 settled. The
/// Outbox branch FAILS on every profile instead, for the reason
/// <see cref="LiveDisconnectRecoveryTests.OutboxRefusal"/> documents.
/// </para>
/// <para>
/// <b>How it is pinned.</b> The callers are <c>Category=Live</c>, so CI can never run them. What a
/// pure function can carry is exercised here with the strings the live tests really use. What no
/// pure function can pin - that the live tests still ASK - is read out of their sources, the
/// substitute <see cref="LivePopulationCoverageTests"/> already uses. Those reads are exact about
/// how many guards each method holds, so removing any single guard fails this file; a control
/// against the pre-Q76 sources fails it too.
/// </para>
/// <para>
/// Synthetic names throughout. Nothing touches Outlook, a mailbox or a machine-local settings
/// file (S6).
/// </para>
/// </summary>
public sealed class LiveEarlyReturnGuardTests
{
    /// <summary>
    /// A guard call, as it appears at the start of its own line: the resumable scan's wrapper, the
    /// shared helper, or the bare Production check the older guards call.
    /// </summary>
    private static readonly Regex GuardCall = new(
        @"\bRequireResumeTokens\(|\bLivePopulationCoverage\.Require\(|\bRequireProductionPopulation\(",
        RegexOptions.CultureInvariant);

    /// <summary>A test method's signature line; the name is the group.</summary>
    private static readonly Regex TestSignature = new(
        @"^    public (?:void|async Task) (?<name>\w+)\(", RegexOptions.CultureInvariant);

    // =============================================== 1. the resumable scan, with its real strings

    [Fact]
    public void AScanThatHandsOutNoTokenSaysSoOnPortable_NamingThePageSizeAndTheRemedy()
    {
        List<string> lines = new();

        IReadOnlyList<string> tokens = LivePopulationCoverage.Require(
            Settings(LiveMachineProfile.Portable), Array.Empty<string>(),
            LiveResumableScanTests.ResumeTokenPopulation, "the changed-question refusal",
            LiveResumableScanTests.ResumeTokenRemedy, lines.Add);

        Assert.Empty(tokens);
        string provedNothing = Assert.Single(lines, l => l.StartsWith("PROVED NOTHING:", StringComparison.Ordinal));
        Assert.Contains("the changed-question refusal iterated nothing", provedNothing, StringComparison.Ordinal);
        Assert.Contains("the test hub paged at top 2", provedNothing, StringComparison.Ordinal);

        // The remedy is computed from the page size, and it has to be enough for the test that
        // needs the most: three pages, so a second token exists to supersede the first.
        Assert.Contains("at least 5 mail items", provedNothing, StringComparison.Ordinal);
        Assert.Contains("Reset-HubPopulation.ps1", provedNothing, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(ResumableScanPopulations))]
    public void AScanThatHandsOutNoTokenRefusesOnProduction_BeforeAnnouncingAnything(string population)
    {
        List<string> lines = new();

        InvalidOperationException refusal = Assert.Throws<InvalidOperationException>(
            () => LivePopulationCoverage.Require(
                Settings(LiveMachineProfile.Production), Array.Empty<string>(),
                population, "the superseded-token refusal", LiveResumableScanTests.ResumeTokenRemedy, lines.Add));

        Assert.Contains(population, refusal.Message, StringComparison.Ordinal);
        Assert.DoesNotContain(lines, l => l.StartsWith("PROVED NOTHING:", StringComparison.Ordinal));
    }

    [Fact]
    public void AScanThatHandsOutATokenIsCountedAndNotAnnounced()
    {
        List<string> lines = new();

        IReadOnlyList<string> tokens = LivePopulationCoverage.Require(
            Settings(LiveMachineProfile.Production), new[] { "scan-0123456789abcdef0123456789abcdef" },
            LiveResumableScanTests.ResumeTokenPopulation, "the resume-rung record",
            LiveResumableScanTests.ResumeTokenRemedy, lines.Add);

        Assert.Single(tokens);
        string coverage = Assert.Single(lines);
        Assert.StartsWith("coverage: 1 ", coverage, StringComparison.Ordinal);
    }

    public static IEnumerable<object[]> ResumableScanPopulations()
    {
        yield return new object[] { LiveResumableScanTests.ResumeTokenPopulation };
        yield return new object[] { LiveResumableScanTests.SecondResumeTokenPopulation };
    }

    // ================================ 2. the delegate nested-folder probe, with its real decision

    /// <summary>
    /// The guard fires on exactly the leaf-match shapes the old test skipped on, and on no other.
    /// Each case is <c>|</c>-separated matches of <c>/</c>-separated path segments; empty is none.
    /// </summary>
    [Theory]
    [InlineData("", 0)]
    [InlineData("Invoices", 0)]
    [InlineData("Invoices|Invoices", 0)]
    [InlineData("Archive/Invoices", 1)]
    [InlineData("Invoices|Archive/Invoices", 1)]
    [InlineData("Archive/Invoices|Team/2026/Invoices", 2)]
    public void TheNestedPathGuardFiresExactlyWhereTheOldSkipDid(string shape, int nested)
    {
        IReadOnlyList<IReadOnlyList<string>> matches = shape.Length == 0
            ? Array.Empty<IReadOnlyList<string>>()
            : shape.Split('|').Select(m => (IReadOnlyList<string>)m.Split('/')).ToList();

        // The condition the test used to log "skipped this run" on and return green.
        bool oldSkip = matches.Count == 0 || matches.All(m => m.Count <= 1);

        IReadOnlyList<IReadOnlyList<string>> found = LiveStaleIndexRowTests.NestedPaths(matches);
        Assert.Equal(nested, found.Count);
        Assert.Equal(oldSkip, found.Count == 0);
    }

    [Fact]
    public void ATreeListingNothingNestedSaysSoOnPortable_AndRefusesOnProduction()
    {
        IReadOnlyList<IReadOnlyList<string>> topLevelOnly = LiveStaleIndexRowTests.NestedPaths(
            new[] { (IReadOnlyList<string>)new[] { "Invoices" } });

        List<string> portable = new();
        _ = LivePopulationCoverage.Require(
            Settings(LiveMachineProfile.Portable), topLevelOnly,
            LiveStaleIndexRowTests.NestedPathPopulation, "the delegate leaf-name locator assertion",
            LiveStaleIndexRowTests.NestedPathRemedy, portable.Add);
        string line = Assert.Single(portable, l => l.StartsWith("PROVED NOTHING:", StringComparison.Ordinal));
        Assert.Contains("the delegate leaf-name locator assertion iterated nothing", line, StringComparison.Ordinal);
        Assert.Contains("lazily", line, StringComparison.Ordinal);
        Assert.Contains("delegateNestedFolderProbe", line, StringComparison.Ordinal);

        List<string> production = new();
        InvalidOperationException refusal = Assert.Throws<InvalidOperationException>(
            () => LivePopulationCoverage.Require(
                Settings(LiveMachineProfile.Production), topLevelOnly,
                LiveStaleIndexRowTests.NestedPathPopulation, "the delegate leaf-name locator assertion",
                LiveStaleIndexRowTests.NestedPathRemedy, production.Add));
        Assert.Contains(LiveStaleIndexRowTests.NestedPathPopulation, refusal.Message, StringComparison.Ordinal);
        Assert.DoesNotContain(production, l => l.StartsWith("PROVED NOTHING:", StringComparison.Ordinal));
    }

    [Fact]
    public void NestedPathsRefusesANullListRatherThanReadingItAsEmpty()
    {
        Assert.Throws<ArgumentNullException>(() => LiveStaleIndexRowTests.NestedPaths(null!));
    }

    // ============================================================== 3. the Outbox, on every profile

    [Fact]
    public void AnEmptyOutboxLetsTheScenarioRun()
    {
        Assert.Null(LiveDisconnectRecoveryTests.OutboxRefusal(0));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(5532)]
    public void AnOutboxHoldingMailFailsTheTest_NamingTheCountAndBothKindsOfMachine(int queued)
    {
        string? refusal = LiveDisconnectRecoveryTests.OutboxRefusal(queued);

        Assert.NotNull(refusal);
        Assert.Contains("holds " + queued.ToString(CultureInfo.InvariantCulture) + " item(s)", refusal, StringComparison.Ordinal);
        Assert.Contains("proved nothing about disconnect recovery", refusal, StringComparison.Ordinal);
        Assert.Contains("fails on every machine profile", refusal, StringComparison.Ordinal);
        Assert.Contains("On a test guest", refusal, StringComparison.Ordinal);
        Assert.Contains("On a working profile", refusal, StringComparison.Ordinal);

        // It is the failure message, so nothing in it may read as a pass.
        Assert.DoesNotContain("SKIP", refusal, StringComparison.Ordinal);
        Assert.DoesNotContain("PROVED NOTHING", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void AnUnreadableOutboxCountFailsTheSameWay()
    {
        // -1 is CountOutboxItems' "the walk itself failed". Unknown is unsafe under S7, the rule
        // LiveMailSink.EnsureOutboxDrained keeps as well.
        string? refusal = LiveDisconnectRecoveryTests.OutboxRefusal(-1);

        Assert.NotNull(refusal);
        Assert.Contains("could not be read", refusal, StringComparison.Ordinal);
        Assert.Contains("fails on every machine profile", refusal, StringComparison.Ordinal);
    }

    [Fact]
    public void TheOutboxBranchFailsTheTest_AndIsNotGatedOnTheProfile()
    {
        string source = LiveSource("LiveDisconnectRecoveryTests.cs");

        // The branch: from the S7 Outbox count to the line that announces the graceful close it
        // protects. Both anchors are code, not comments.
        int count = source.IndexOf("CountOutboxItems()", StringComparison.Ordinal);
        int close = source.IndexOf("parked Explorer window(s) gracefully", StringComparison.Ordinal);
        Assert.True(count >= 0 && close > count, "the Outbox branch of LiveDisconnectRecoveryTests was not found");
        string branch = source.Substring(count, close - count);

        Assert.Contains("OutboxRefusal(outboxItems)", branch, StringComparison.Ordinal);
        Assert.Contains("Assert.Fail(outboxRefusal)", branch, StringComparison.Ordinal);

        // Neither a skip nor the Production/Portable split: both would let a Portable run pass here.
        Assert.DoesNotContain("return;", branch, StringComparison.Ordinal);
        Assert.DoesNotContain("SKIP", branch, StringComparison.Ordinal);
        Assert.DoesNotContain("RequireProductionPopulation", branch, StringComparison.Ordinal);
        Assert.DoesNotContain("LivePopulationCoverage", branch, StringComparison.Ordinal);
    }

    // ======================================== 4. the population guards: read out of the sources

    /// <summary>
    /// How many guards each of these test methods holds - EXACT, so removing any single one fails.
    /// The superseded-token test has two because it needs a chain still open after its SECOND
    /// page; the delegate test has two because its probe can be missing as well as unlisted.
    /// </summary>
    [Theory]
    [InlineData("LiveResumableScanTests.cs", "APagedScan_ReturnsExactlyWhatOneUnpagedScanReturns_WithNoDuplicates", 1)]
    [InlineData("LiveResumableScanTests.cs", "APagedScan_ReportsWhichRungItResumedOn_SoTheSortQuestionIsAnsweredInPassing", 1)]
    [InlineData("LiveResumableScanTests.cs", "AResumeWithAChangedQuestion_IsRefused_AndTheRefusalNamesWhatChanged", 1)]
    [InlineData("LiveResumableScanTests.cs", "ASupersededToken_IsRefusedWithThePositionNeededToCarryOnWithoutIt", 2)]
    [InlineData("LiveStaleIndexRowTests.cs", "DelegateHitsInANestedFolder_AreReadable_ViaTheFlatLeafName", 2)]
    public void EachGuardedTestKeepsEveryOneOfItsGuards(string file, string method, int guards)
    {
        string[] body = TestMethodBody(file, method);

        int found = body.Count(l => GuardCall.IsMatch(l));
        Assert.True(
            found == guards,
            $"{file} {method} holds {found} guard call(s) where Q76 put {guards}. A guard removed is a test "
            + "that can pass having proved nothing again; one added belongs in this table.");
    }

    [Theory]
    [InlineData("LiveResumableScanTests.cs")]
    [InlineData("LiveStaleIndexRowTests.cs")]
    public void NoEarlyReturnInTheseFilesComesBeforeItsGuard(string file)
    {
        string[] lines = SourceLines(file);
        List<string> problems = new();
        foreach (string method in TestMethodNames(lines))
        {
            string[] body = TestMethodBody(file, method);
            int guardsSoFar = 0;
            int returnsSoFar = 0;
            for (int i = 0; i < body.Length; i++)
            {
                if (GuardCall.IsMatch(body[i]))
                {
                    guardsSoFar++;
                }

                if (!string.Equals(body[i].Trim(), "return;", StringComparison.Ordinal))
                {
                    continue;
                }

                // The k-th early return needs k guards before it. Anything less is a return that
                // was not preceded by its own refusal-or-announcement: the pre-Q76 shape.
                returnsSoFar++;
                if (guardsSoFar < returnsSoFar)
                {
                    problems.Add($"{method}: early return #{returnsSoFar} follows only {guardsSoFar} guard(s)");
                }
            }
        }

        Assert.True(problems.Count == 0, file + ": " + string.Join("; ", problems));
    }

    [Fact]
    public void ABareProductionCheckIsAlwaysFollowedByItsAnnouncement()
    {
        // RequireProductionPopulation throws on Production and does NOTHING on Portable, so on its
        // own it is exactly the silent green this file exists to stop. Wherever a test calls it
        // directly, a PROVED NOTHING line has to come before the return it guards.
        foreach (string file in new[] { "LiveResumableScanTests.cs", "LiveStaleIndexRowTests.cs" })
        {
            string[] lines = SourceLines(file);
            foreach (string method in TestMethodNames(lines))
            {
                string[] body = TestMethodBody(file, method);
                for (int i = 0; i < body.Length; i++)
                {
                    if (!body[i].Contains("RequireProductionPopulation(", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    int next = Array.FindIndex(body, i, l => string.Equals(l.Trim(), "return;", StringComparison.Ordinal));
                    Assert.True(next > i, $"{file} {method}: a bare RequireProductionPopulation with no return after it");
                    string between = string.Join("\n", body, i, next - i);
                    Assert.True(
                        between.Contains("\"PROVED NOTHING:", StringComparison.Ordinal),
                        $"{file} {method}: RequireProductionPopulation returns without a PROVED NOTHING line");
                }
            }
        }
    }

    [Fact]
    public void TheResumableScanGuardIsTheSharedHelper_AndTheDelegateGuardIsFedTheNestedPaths()
    {
        // One wrapper, and it adds nothing to the decision: the lines and the refusal are the
        // shared helper's, so this file's tests of that helper cover what the live run prints.
        string helper = MemberBody("LiveResumableScanTests.cs", "private IReadOnlyList<string> RequireResumeTokens(");
        Assert.Contains("LivePopulationCoverage.Require(", helper, StringComparison.Ordinal);
        Assert.Contains("_fixture.Settings", helper, StringComparison.Ordinal);
        Assert.Contains("_output.WriteLine", helper, StringComparison.Ordinal);

        // And the delegate test hands the helper the decision this file exercises above, not a
        // second opinion about which leaf matches count.
        string delegateTest = string.Join(
            "\n", TestMethodBody("LiveStaleIndexRowTests.cs", "DelegateHitsInANestedFolder_AreReadable_ViaTheFlatLeafName"));
        Assert.Contains("NestedPaths(matches)", delegateTest, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("LiveResumableScanTests.cs", "no rung was exercised")]
    [InlineData("LiveResumableScanTests.cs", "no token to refuse")]
    [InlineData("LiveResumableScanTests.cs", "no chain to supersede")]
    [InlineData("LiveResumableScanTests.cs", "- the chain finished")]
    [InlineData("LiveStaleIndexRowTests.cs", "the locator assertion is skipped this run")]
    [InlineData("LiveDisconnectRecoveryTests.cs", "Outbox item(s) (or count unavailable) - not closing anything")]
    public void TheLinesThatUsedToStandInForAResultAreGone(string file, string oldLine)
    {
        // Each of these was printed immediately before a green return, and was the whole of what
        // the run said about a test that had asserted nothing.
        Assert.DoesNotContain(oldLine, LiveSource(file), StringComparison.Ordinal);
    }

    // ================================================================================== helpers

    private static LiveTestSettings Settings(LiveMachineProfile profile)
    {
        return new LiveTestSettings
        {
            MachineProfile = profile,
            TestHubStoreDisplayName = "hub@example.test",
            ExpectedStoreDisplayNames = new List<string> { "hub@example.test" },
        };
    }

    private static IEnumerable<string> TestMethodNames(string[] lines)
    {
        for (int i = 0; i < lines.Length; i++)
        {
            string attribute = lines[i].Trim();
            if (!string.Equals(attribute, "[Fact]", StringComparison.Ordinal)
                && !attribute.StartsWith("[Theory", StringComparison.Ordinal))
            {
                continue;
            }

            for (int j = i + 1; j < lines.Length; j++)
            {
                Match signature = TestSignature.Match(lines[j]);
                if (signature.Success)
                {
                    yield return signature.Groups["name"].Value;
                    break;
                }
            }
        }
    }

    /// <summary>The lines strictly inside one test method's braces, at the class member indentation.</summary>
    private static string[] TestMethodBody(string file, string method)
    {
        string[] lines = SourceLines(file);
        int signature = Array.FindIndex(lines, l =>
        {
            Match m = TestSignature.Match(l);
            return m.Success && string.Equals(m.Groups["name"].Value, method, StringComparison.Ordinal);
        });
        Assert.True(signature >= 0, method + " was not found in " + file + " - this test has stopped proving anything");
        return BodyAfter(lines, signature, file, method);
    }

    private static string MemberBody(string file, string declarationStart)
    {
        string[] lines = SourceLines(file);
        int declaration = Array.FindIndex(lines, l => l.TrimStart().StartsWith(declarationStart, StringComparison.Ordinal));
        Assert.True(declaration >= 0, declarationStart + " was not found in " + file + " - this test has stopped proving anything");
        return string.Join("\n", BodyAfter(lines, declaration, file, declarationStart));
    }

    private static string[] BodyAfter(string[] lines, int declaration, string file, string what)
    {
        int open = Array.FindIndex(lines, declaration, l => string.Equals(l, "    {", StringComparison.Ordinal));
        int close = open < 0 ? -1 : Array.FindIndex(lines, open + 1, l => string.Equals(l, "    }", StringComparison.Ordinal));
        Assert.True(open > declaration && close > open, "could not find the body of " + what + " in " + file);
        return lines[(open + 1)..close];
    }

    private static string[] SourceLines(string fileName)
    {
        return LiveSource(fileName).Split('\n').Select(l => l.TrimEnd('\r')).ToArray();
    }

    private static string LiveSource(string fileName)
    {
        string dir = typeof(LiveTestSettings).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
        string path = Path.Combine(dir, "T2", fileName);
        Assert.True(File.Exists(path), "live test source is missing: " + path);
        return File.ReadAllText(path);
    }
}
