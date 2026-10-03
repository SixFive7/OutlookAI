using System.Globalization;
using System.Reflection;
using System.Text.RegularExpressions;
using OutlookAI.McpServer.Tests.T2;
using OutlookAI.McpServer.Tests.T3;
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
/// <b>The maintainer's decision Q101 (2026-10-03) extends it to the rest.</b> The Q57 pattern for
/// <c>T3/OutlookAvailabilityLiveTests</c> (a healthy Outlook; and a search error, now narrowed to the
/// one where the index is unreachable - any other error FAILS), <c>T3/ComHostSupervisionLiveTests</c>
/// (no COM host spawned), <c>T2/LiveUiSearchBackendTests</c> (a policy-hive value) and
/// <c>T2/LiveHeadlessGuaranteeTests</c> (no openable hub hit, no conversation). Its INVERSE for the
/// user-protection stops of <c>T2/LiveDisconnectRecoveryTests</c>:
/// <see cref="LivePopulationCoverage.StandAsideForAUser"/> skips with a line on Production, where a
/// person may be at the keyboard, and fails everywhere else. That test also takes its Inspector
/// and Outbox counts before it looks at windows at all (3(a)), and
/// <c>T2/LiveStaleIndexRowTests</c> waits a bounded time for the delegate tree before refusing (2(b)).
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
    /// shared helper, its user-protection inverse, or the bare Production check the older guards call.
    /// </summary>
    private static readonly Regex GuardCall = new(
        @"\bRequireResumeTokens\(|\bLivePopulationCoverage\.Require\(|\bLivePopulationCoverage\.StandAsideForAUser\(|\bRequireProductionPopulation\(",
        RegexOptions.CultureInvariant);

    /// <summary>
    /// Every live source this file reads for guards. A bare name is under <c>T2/</c>; the stdio
    /// tier's files carry their folder.
    /// </summary>
    public static IEnumerable<object[]> GuardedFiles()
    {
        yield return new object[] { "LiveResumableScanTests.cs" };
        yield return new object[] { "LiveStaleIndexRowTests.cs" };
        yield return new object[] { "LiveDisconnectRecoveryTests.cs" };
        yield return new object[] { "LiveUiSearchBackendTests.cs" };
        yield return new object[] { "LiveHeadlessGuaranteeTests.cs" };
        yield return new object[] { "T3/OutlookAvailabilityLiveTests.cs" };
        yield return new object[] { "T3/ComHostSupervisionLiveTests.cs" };
    }

    /// <summary>A test method's signature line; the name is the group.</summary>
    private static readonly Regex TestSignature = new(
        @"^    public (?:void|async Task) (?<name>\w+)\(", RegexOptions.CultureInvariant);

    /// <summary>
    /// A bare <c>return;</c> on its own line - with or without a comment after it. The comment form
    /// is how two of the returns Q101 converted were written (<c>return; // Outlook was healthy
    /// here...</c>), and the Q76 reader, which matched only the bare form, could not see them.
    /// </summary>
    private static readonly Regex EarlyReturnLine = new(@"^\s*return;\s*(?://.*)?$", RegexOptions.CultureInvariant);

    private static bool IsEarlyReturn(string line)
    {
        return EarlyReturnLine.IsMatch(line);
    }

    [Theory]
    [InlineData("            return;", true)]
    [InlineData("            return; // Outlook was healthy here; nothing transient to assert.", true)]
    [InlineData("            return;   //no space", true)]
    [InlineData("            return state;", false)]
    [InlineData("            // return;", false)]
    [InlineData("            returned = true;", false)]
    public void TheEarlyReturnReaderSeesTheCommentedFormToo(string line, bool isReturn)
    {
        // The reader every source-order check here rests on, held to the shapes it must and must
        // not match - a reader that stopped matching would turn those checks into vacuous passes.
        Assert.Equal(isReturn, IsEarlyReturn(line));
    }

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

        // The branch: from the S7 Outbox count to the first look at Outlook's windows, which since
        // Q101 3(a) comes straight after the two counts. Both anchors are code, not comments.
        int count = source.IndexOf("CountOutboxItems()", StringComparison.Ordinal);
        int close = source.IndexOf(WindowsBaseline, StringComparison.Ordinal);
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
    [InlineData("T3/OutlookAvailabilityLiveTests.cs", "ATransientOutlookState_AnswersFastAndCarriesRetryGuidance", 1)]
    [InlineData("T3/OutlookAvailabilityLiveTests.cs", "SearchAlwaysAnswers_AndSaysWhetherItIsComplete", 1)]
    [InlineData("T3/ComHostSupervisionLiveTests.cs", "NoComHostSurvivesTheServer", 1)]
    [InlineData("LiveUiSearchBackendTests.cs", "FlippingUserHiveValue_DrivesAdviceAndHealthField_BothStates", 1)]
    [InlineData("LiveHeadlessGuaranteeTests.cs", "NonShowMeOperations_NeverCreateAnOutlookWindow", 2)]
    [InlineData("LiveDisconnectRecoveryTests.cs", "OutlookExit_ReleasesHeldRefsInBackground_HealthProbes_GatewayReattaches", 5)]
    public void EachGuardedTestKeepsEveryOneOfItsGuards(string file, string method, int guards)
    {
        string[] body = TestMethodBody(file, method);

        int found = body.Count(l => GuardCall.IsMatch(l));
        Assert.True(
            found == guards,
            $"{file} {method} holds {found} guard call(s) where Q76/Q101 put {guards}. A guard removed is a test "
            + "that can pass having proved nothing again; one added belongs in this table.");
    }

    [Theory]
    [MemberData(nameof(GuardedFiles))]
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

                if (!IsEarlyReturn(body[i]))
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
        foreach (string file in GuardedFiles().Select(row => (string)row[0]))
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

                    int next = Array.FindIndex(body, i, IsEarlyReturn);
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
    [InlineData("T3/OutlookAvailabilityLiveTests.cs", "Outlook was healthy here; nothing transient to assert")]
    [InlineData("T3/OutlookAvailabilityLiveTests.cs", "No index on this machine; the freshness contract does not apply")]
    [InlineData("T3/ComHostSupervisionLiveTests.cs", "is all this run can honestly prove")]
    [InlineData("LiveUiSearchBackendTests.cs", "SKIP: policy-hive DisableServerAssistedSearch=")]
    [InlineData("LiveHeadlessGuaranteeTests.cs", "read/thread delta checks skipped this run")]
    [InlineData("LiveDisconnectRecoveryTests.cs", "SKIP: Outlook windows exist and the user was active")]
    [InlineData("LiveDisconnectRecoveryTests.cs", "open Inspector window(s) (possible unsent compose) - not closing anything")]
    [InlineData("LiveDisconnectRecoveryTests.cs", "SKIP: a window appeared during re-autostart")]
    [InlineData("LiveDisconnectRecoveryTests.cs", "SKIP: expected exactly our window before close")]
    [InlineData("LiveDisconnectRecoveryTests.cs", "SKIP(3b)")]
    public void TheLinesThatUsedToStandInForAResultAreGone(string file, string oldLine)
    {
        // Each of these was printed immediately before a green return, and was the whole of what
        // the run said about a test that had asserted nothing.
        Assert.DoesNotContain(oldLine, LiveSource(file), StringComparison.Ordinal);
    }

    // ============================ 5. Q101, the Q57 pattern: the real strings, through the real helper

    /// <summary>
    /// Every population Q101 put behind <see cref="LivePopulationCoverage.Require"/>, as the live test
    /// passes it: the population, what would not run, the remedy, and one phrase the remedy must keep.
    /// </summary>
    public static IEnumerable<object[]> Q101Populations()
    {
        yield return new object[]
        {
            OutlookAvailabilityLiveTests.TransientStatePopulation, "the retry-guidance check",
            OutlookAvailabilityLiveTests.TransientStateRemedy, "straight after Outlook was closed",
        };
        yield return new object[]
        {
            OutlookAvailabilityLiveTests.AnsweredSearchPopulation, "the freshness-contract assertions",
            OutlookAvailabilityLiveTests.AnsweredSearchRemedy, "WSearch",
        };
        yield return new object[]
        {
            ComHostSupervisionLiveTests.ComHostPopulation, "the check that the COM host dies with its server",
            ComHostSupervisionLiveTests.ComHostRemedy, "start Outlook before the run",
        };
        yield return new object[]
        {
            LiveUiSearchBackendTests.UserHiveInControlPopulation, "the two-state user-hive flip",
            LiveUiSearchBackendTests.UserHiveInControlRemedy, "POLICY hive",
        };
        yield return new object[]
        {
            LiveHeadlessGuaranteeTests.OpenableHubHitPopulation, "the read and thread window-delta checks",
            LiveHeadlessGuaranteeTests.OpenableHubHitRemedy, "Reset-HubPopulation.ps1",
        };
        yield return new object[]
        {
            LiveHeadlessGuaranteeTests.ConversationPopulation, "the thread window-delta check",
            LiveHeadlessGuaranteeTests.ConversationRemedy, "conversation",
        };
    }

    [Theory]
    [MemberData(nameof(Q101Populations))]
    public void AQ101PopulationThatIsMissingSaysSoOnPortable_NamingWhatDidNotRunAndTheRemedy(
        string population, string whatWouldNotRun, string remedy, string remedyMustSay)
    {
        List<string> lines = new();

        IReadOnlyList<string> found = LivePopulationCoverage.Require(
            Settings(LiveMachineProfile.Portable), Array.Empty<string>(), population, whatWouldNotRun, remedy, lines.Add);

        Assert.Empty(found);
        string provedNothing = Assert.Single(lines, l => l.StartsWith("PROVED NOTHING:", StringComparison.Ordinal));
        Assert.Contains(whatWouldNotRun + " iterated nothing", provedNothing, StringComparison.Ordinal);
        Assert.Contains(population, provedNothing, StringComparison.Ordinal);
        Assert.Contains(remedyMustSay, provedNothing, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(Q101Populations))]
    public void AQ101PopulationThatIsMissingRefusesOnProduction_BeforeAnnouncingAnything(
        string population, string whatWouldNotRun, string remedy, string remedyMustSay)
    {
        _ = remedyMustSay;
        List<string> lines = new();

        InvalidOperationException refusal = Assert.Throws<InvalidOperationException>(
            () => LivePopulationCoverage.Require(
                Settings(LiveMachineProfile.Production), Array.Empty<string>(), population, whatWouldNotRun, remedy, lines.Add));

        Assert.Contains(population, refusal.Message, StringComparison.Ordinal);
        Assert.DoesNotContain(lines, l => l.StartsWith("PROVED NOTHING:", StringComparison.Ordinal));
    }

    /// <summary>
    /// Each live guard is fed ITS population, remedy and description - not a neighbour's - so the
    /// two tests above exercise the strings the live run really prints.
    /// </summary>
    [Theory]
    [InlineData("T3/OutlookAvailabilityLiveTests.cs", "ATransientOutlookState_AnswersFastAndCarriesRetryGuidance",
        "TransientStatePopulation,", "TransientStateRemedy,", "\"the retry-guidance check\"")]
    [InlineData("T3/OutlookAvailabilityLiveTests.cs", "SearchAlwaysAnswers_AndSaysWhetherItIsComplete",
        "AnsweredSearchPopulation,", "AnsweredSearchRemedy,", "\"the freshness-contract assertions\"")]
    [InlineData("T3/ComHostSupervisionLiveTests.cs", "NoComHostSurvivesTheServer",
        "ComHostPopulation,", "ComHostRemedy,", "\"the check that the COM host dies with its server\"")]
    [InlineData("LiveUiSearchBackendTests.cs", "FlippingUserHiveValue_DrivesAdviceAndHealthField_BothStates",
        "UserHiveInControlPopulation,", "UserHiveInControlRemedy,", "\"the two-state user-hive flip\"")]
    [InlineData("LiveHeadlessGuaranteeTests.cs", "NonShowMeOperations_NeverCreateAnOutlookWindow",
        "OpenableHubHitPopulation,", "OpenableHubHitRemedy,", "\"the read and thread window-delta checks\"")]
    [InlineData("LiveHeadlessGuaranteeTests.cs", "NonShowMeOperations_NeverCreateAnOutlookWindow",
        "ConversationPopulation,", "ConversationRemedy,", "\"the thread window-delta check\"")]
    public void EachQ101GuardIsFedItsOwnStrings(
        string file, string method, string population, string remedy, string whatWouldNotRun)
    {
        string body = string.Join("\n", TestMethodBody(file, method));

        Assert.Contains(population, body, StringComparison.Ordinal);
        Assert.Contains(remedy, body, StringComparison.Ordinal);
        Assert.Contains(whatWouldNotRun, body, StringComparison.Ordinal);
    }

    [Fact]
    public void TheThreadCheckRunsOnlyBehindBothItsGuards()
    {
        // Line 83's silent skip, stated as the shape that replaced it: the thread call sits inside
        // the conversation guard, which sits inside the openable-hit guard - so neither check can be
        // skipped without its PROVED NOTHING line or its Production refusal coming first.
        string[] body = TestMethodBody("LiveHeadlessGuaranteeTests.cs", "NonShowMeOperations_NeverCreateAnOutlookWindow");
        int hitGuard = Array.FindIndex(body, l => l.Contains("OpenableHubHitPopulation,", StringComparison.Ordinal));
        int conversationGuard = Array.FindIndex(body, l => l.Contains("ConversationPopulation,", StringComparison.Ordinal));
        int thread = Array.FindIndex(body, l => l.Contains("service.Thread(", StringComparison.Ordinal));

        Assert.True(hitGuard >= 0 && conversationGuard > hitGuard && thread > conversationGuard,
            $"expected the hit guard, then the conversation guard, then the thread call; found lines {hitGuard}, {conversationGuard}, {thread}");
        Assert.Single(body, l => l.Contains("service.Thread(", StringComparison.Ordinal));
        Assert.DoesNotContain(body, l => l.Contains("read.ConversationId != null", StringComparison.Ordinal));
    }

    // ========================================= 6. Q101, the real bug: only the no-index error passes

    [Theory]
    [InlineData("unavailable: COMException", true)]
    [InlineData("unavailable: InvalidOperationException", true)]
    [InlineData("unavailable: OleDbException", true)]
    [InlineData("OleDb", false)]
    [InlineData("AdodbCom", false)]
    [InlineData("", false)]
    [InlineData(null, false)]
    public void OnlyTheProductsOwnUnreachableVerdictLetsASearchErrorThrough(string? provider, bool noIndex)
    {
        // The provider strings are outlook_health's index.provider: "OleDb" or "AdodbCom" when the
        // index answers, "unavailable: <exception type>" when MailService.Health could not reach it.
        // Null is health not saying - and that is NOT a reason to let an error through.
        Assert.Equal(noIndex, OutlookAvailabilityLiveTests.IndexIsUnreachable(provider));
    }

    [Fact]
    public void ASearchErrorFailsUnlessTheIndexIsUnreachable_AndThenItSaysSo()
    {
        string[] body = TestMethodBody("T3/OutlookAvailabilityLiveTests.cs", "SearchAlwaysAnswers_AndSaysWhetherItIsComplete");
        string joined = string.Join("\n", body);

        // The old blanket return is gone...
        Assert.DoesNotContain("TryGetProperty(\"error\", out _)", joined, StringComparison.Ordinal);

        // ...and in its place: read the error, assert it is the no-index one, THEN the guard, THEN
        // the only return - in that order.
        int error = Array.FindIndex(body, l => l.Contains("TryGetProperty(\"error\", out JsonElement error)", StringComparison.Ordinal));
        int verdict = Array.FindIndex(body, l => l.Contains("IndexIsUnreachable(provider)", StringComparison.Ordinal));
        int guard = Array.FindIndex(body, l => GuardCall.IsMatch(l));
        int firstReturn = Array.FindIndex(body, IsEarlyReturn);
        Assert.True(error >= 0 && verdict > error && guard > verdict && firstReturn > guard,
            $"expected error read < verdict < guard < return; found lines {error}, {verdict}, {guard}, {firstReturn}");
        Assert.Contains("Assert.True(", body[verdict - 1] + body[verdict], StringComparison.Ordinal);
        Assert.Single(body, IsEarlyReturn);
    }

    // ================================ 7. Q101, the inverse: a user-protection stop skips only for a user

    /// <summary>
    /// The user-protection stops of <c>LiveDisconnectRecoveryTests</c>: what the live test says it
    /// saw (shaped as it builds it), what did not run, and what it means where nobody is.
    /// </summary>
    public static IEnumerable<object[]> UserProtectionStops()
    {
        yield return new object[]
        {
            "2 Inspector window(s) are open (possibly an unsent compose)",
            LiveDisconnectRecoveryTests.Scenario, LiveDisconnectRecoveryTests.OpenInspectorOnAnUnattendedMachine,
        };
        yield return new object[]
        {
            "Outlook windows are open and the user was active 12 s ago",
            LiveDisconnectRecoveryTests.Scenario, LiveDisconnectRecoveryTests.RecentInputOnAnUnattendedMachine,
        };
        yield return new object[]
        {
            "1 Outlook window(s) appeared while Outlook re-autostarted headless",
            LiveDisconnectRecoveryTests.Scenario, LiveDisconnectRecoveryTests.WindowDuringRestartOnAnUnattendedMachine,
        };
        yield return new object[]
        {
            "2 Outlook windows were visible where only the one this test opened should be (ours has been closed)",
            LiveDisconnectRecoveryTests.Scenario, LiveDisconnectRecoveryTests.ExtraWindowOnAnUnattendedMachine,
        };
        yield return new object[]
        {
            "a real OutlookAISetup mutex is already held (an add-in install or update is running)",
            "the degraded-search check (3b), which has to hold that mutex itself",
            LiveDisconnectRecoveryTests.InstallerMutexOnAnUnattendedMachine,
        };
    }

    [Theory]
    [MemberData(nameof(UserProtectionStops))]
    public void AUserProtectionStopSkipsWithOneLine_WhereAUserMayBeAtTheKeyboard(
        string observed, string whatDidNotRun, string onAnUnattendedMachine)
    {
        List<string> lines = new();

        LivePopulationCoverage.StandAsideForAUser(
            Settings(LiveMachineProfile.Production), observed, whatDidNotRun, onAnUnattendedMachine, lines.Add);

        string line = Assert.Single(lines);
        Assert.StartsWith("SKIP (user protection): ", line, StringComparison.Ordinal);
        Assert.Contains(observed, line, StringComparison.Ordinal);
        Assert.Contains(whatDidNotRun + " did not run", line, StringComparison.Ordinal);
        Assert.Contains("machineProfile=Production", line, StringComparison.Ordinal);
        Assert.DoesNotContain("PROVED NOTHING", line, StringComparison.Ordinal);
    }

    [Theory]
    [MemberData(nameof(UserProtectionStops))]
    public void AUserProtectionStopFails_WhereNobodyShouldBe_BeforeSayingAnything(
        string observed, string whatDidNotRun, string onAnUnattendedMachine)
    {
        foreach (LiveMachineProfile unattended in new[] { LiveMachineProfile.Portable, (LiveMachineProfile)93 })
        {
            List<string> lines = new();

            InvalidOperationException refusal = Assert.Throws<InvalidOperationException>(
                () => LivePopulationCoverage.StandAsideForAUser(
                    Settings(unattended), observed, whatDidNotRun, onAnUnattendedMachine, lines.Add));

            // Nothing that could read as an acceptable skip is said before the failure.
            Assert.Empty(lines);
            Assert.Contains(observed, refusal.Message, StringComparison.Ordinal);
            Assert.Contains(whatDidNotRun + " did not run", refusal.Message, StringComparison.Ordinal);
            Assert.Contains(onAnUnattendedMachine, refusal.Message, StringComparison.Ordinal);
            Assert.Contains("FAILS instead of skipping", refusal.Message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void OnlyAProductionProfileMayHaveAUserAtTheKeyboard()
    {
        // The one place the inverse decision reads the profile. A value added to LiveMachineProfile
        // later, before anybody decides what it means, gets the louder answer - a failure.
        foreach (LiveMachineProfile profile in Enum.GetValues<LiveMachineProfile>())
        {
            Assert.Equal(profile == LiveMachineProfile.Production, Settings(profile).AUserMayBeAtTheKeyboard);
        }

        Assert.False(Settings((LiveMachineProfile)93).AUserMayBeAtTheKeyboard);
    }

    [Theory]
    [InlineData("", "the scenario", "why")]
    [InlineData("seen", " ", "why")]
    [InlineData("seen", "the scenario", "")]
    public void AUserProtectionStopMustSayWhatItSawWhatDidNotRunAndWhatItMeans(
        string observed, string whatDidNotRun, string onAnUnattendedMachine)
    {
        foreach (LiveMachineProfile profile in new[] { LiveMachineProfile.Production, LiveMachineProfile.Portable })
        {
            Assert.Throws<ArgumentException>(() => LivePopulationCoverage.StandAsideForAUser(
                Settings(profile), observed, whatDidNotRun, onAnUnattendedMachine, _ => { }));
        }
    }

    [Fact]
    public void EveryUserProtectionStopInTheDisconnectTestGoesThroughTheInverse()
    {
        // Five stops, each through the shared inverse and none through the Production/Portable split
        // the other way round, and no hand-written SKIP line left anywhere in the scenario.
        string[] body = TestMethodBody(
            "LiveDisconnectRecoveryTests.cs", "OutlookExit_ReleasesHeldRefsInBackground_HealthProbes_GatewayReattaches");

        Assert.Equal(5, body.Count(l => l.Contains("LivePopulationCoverage.StandAsideForAUser(", StringComparison.Ordinal)));
        Assert.DoesNotContain(body, l => l.Contains("LivePopulationCoverage.Require(", StringComparison.Ordinal));
        Assert.DoesNotContain(body, l => l.Contains("\"SKIP", StringComparison.Ordinal) || l.Contains("$\"SKIP", StringComparison.Ordinal));
        foreach (string meaning in new[]
                 {
                     "OpenInspectorOnAnUnattendedMachine,", "RecentInputOnAnUnattendedMachine,",
                     "WindowDuringRestartOnAnUnattendedMachine,", "ExtraWindowOnAnUnattendedMachine,",
                     "InstallerMutexOnAnUnattendedMachine,",
                 })
        {
            Assert.Single(body, l => l.Contains(meaning, StringComparison.Ordinal));
        }

        // Our own window is closed before the stop that follows a second one, so a failure on a guest
        // does not leave it behind.
        int close = Array.FindIndex(body, l => l.Contains("WindowProbe.PostClose(ourWindow);", StringComparison.Ordinal));
        int extraWindowStop = Array.FindIndex(body, l => l.Contains("ExtraWindowOnAnUnattendedMachine,", StringComparison.Ordinal));
        Assert.True(close >= 0 && close < extraWindowStop, "our window must be closed before the extra-window stop can fail");
    }

    // ========================================= 8. Q101 3(a): the S7 counts come before the window branch

    [Fact]
    public void TheInspectorAndOutboxCountsAreTakenBeforeTheTestLooksAtWindows()
    {
        string[] body = TestMethodBody(
            "LiveDisconnectRecoveryTests.cs", "OutlookExit_ReleasesHeldRefsInBackground_HealthProbes_GatewayReattaches");

        int inspectors = Assert.Single(Enumerable.Range(0, body.Length), i => body[i].Contains("GetOpenInspectors()", StringComparison.Ordinal));
        int outbox = Assert.Single(Enumerable.Range(0, body.Length), i => body[i].Contains("CountOutboxItems()", StringComparison.Ordinal));
        int baseline = Assert.Single(Enumerable.Range(0, body.Length), i => body[i].Contains(WindowsBaseline, StringComparison.Ordinal));
        int branch = Assert.Single(Enumerable.Range(0, body.Length), i => body[i].Contains("if (baselineWindows.Count > 0)", StringComparison.Ordinal));

        // Exactly once each, and both before the scenario looks at windows: a start with no window
        // open used to drive Outlook to exit with neither count taken.
        Assert.True(inspectors < baseline && outbox < baseline && baseline < branch,
            $"expected both counts before the windows baseline; found inspectors at {inspectors}, outbox at {outbox}, "
            + $"baseline at {baseline}, branch at {branch}");

        // And taken UNCONDITIONALLY: each is a statement at the method body's own indentation, exactly
        // as written - not nested in a block, and not behind a condition folded into the same line.
        Assert.Equal("        IReadOnlyList<ComInspectorInfo> inspectors = clock.Step(", body[inspectors - 2]);
        Assert.Equal("        int outboxItems = clock.Step(", body[outbox - 2]);

        // And the Inspector stop is the inverse, not a skip of its own.
        int inspectorStop = Array.FindIndex(body, l => l.Contains("OpenInspectorOnAnUnattendedMachine,", StringComparison.Ordinal));
        Assert.True(inspectorStop > inspectors && inspectorStop < outbox, "the Inspector stop must sit between the two counts");
    }

    // ======================================== 9. Q101 2(b): the delegate tree is waited for, boundedly

    [Fact]
    public void ATreeThatListsTheFolderNestedAtOnceCostsNoWait()
    {
        FakeClock clock = new();

        NestedListingWait listing = LiveStaleIndexRowTests.WaitForNestedListing(
            Walks(Paths("Archive/Invoices")), Budget, Poll, () => clock.Now, clock.Sleep);

        Assert.Equal(1, listing.Walks);
        Assert.Empty(clock.Slept);
        Assert.Single(LiveStaleIndexRowTests.NestedPaths(listing.Matches));
    }

    [Fact]
    public void ATreeThatCatchesUpIsWaitedFor_AndATopLevelMatchDoesNotEndTheWait()
    {
        FakeClock clock = new();

        NestedListingWait listing = LiveStaleIndexRowTests.WaitForNestedListing(
            Walks(Paths(""), Paths("Invoices"), Paths("Invoices|Archive/Invoices")), Budget, Poll, () => clock.Now, clock.Sleep);

        Assert.Equal(3, listing.Walks);
        Assert.Equal(new[] { Poll, Poll }, clock.Slept);
        Assert.Single(LiveStaleIndexRowTests.NestedPaths(listing.Matches));
    }

    [Fact]
    public void ATreeThatNeverListsItStopsAtTheBound_AndTheRefusalStillFires()
    {
        FakeClock clock = new();

        NestedListingWait listing = LiveStaleIndexRowTests.WaitForNestedListing(
            Walks(Paths("Invoices")), Budget, Poll, () => clock.Now, clock.Sleep);

        // A walk at 0, 15, ..., 300 s: 21 walks, exactly the budget slept, never more than one gap at a time.
        Assert.Equal(21, listing.Walks);
        Assert.Equal(Budget, clock.Slept.Aggregate(TimeSpan.Zero, (sum, s) => sum + s));
        Assert.All(clock.Slept, s => Assert.True(s <= Poll));
        Assert.Equal(Budget, listing.Waited);

        // And what it hands on is still empty, so the Production refusal is exactly as before.
        IReadOnlyList<IReadOnlyList<string>> nested = LiveStaleIndexRowTests.NestedPaths(listing.Matches);
        Assert.Throws<InvalidOperationException>(() => LivePopulationCoverage.Require(
            Settings(LiveMachineProfile.Production), nested, LiveStaleIndexRowTests.NestedPathPopulation,
            "the delegate leaf-name locator assertion", LiveStaleIndexRowTests.NestedPathRemedy, _ => { }));
    }

    [Fact]
    public void TheWaitNeverSleepsPastItsBudget_AndWalksOnceMoreAtTheEnd()
    {
        FakeClock clock = new();

        NestedListingWait listing = LiveStaleIndexRowTests.WaitForNestedListing(
            Walks(Paths("")), TimeSpan.FromSeconds(20), Poll, () => clock.Now, clock.Sleep);

        Assert.Equal(new[] { Poll, TimeSpan.FromSeconds(5) }, clock.Slept);
        Assert.Equal(3, listing.Walks);
    }

    [Fact]
    public void AWalkSlowerThanTheWholeBudgetIsNotRepeated()
    {
        FakeClock clock = new();
        int calls = 0;

        NestedListingWait listing = LiveStaleIndexRowTests.WaitForNestedListing(
            () =>
            {
                calls++;
                Assert.True(calls < 1000, "the delegate-tree wait walked a thousand times - its bound is gone");
                clock.Now += TimeSpan.FromSeconds(400);
                return Paths("");
            },
            Budget, Poll, () => clock.Now, clock.Sleep);

        Assert.Equal(1, listing.Walks);
        Assert.Empty(clock.Slept);
    }

    [Fact]
    public void TheBoundIsTheOneDocumented()
    {
        // The values the runbook and the constants' own justification name. A change is a decision
        // about how long a Production run waits on Exchange before refusing, so it shows up here.
        Assert.Equal(300, LiveStaleIndexRowTests.NestedListingWaitSeconds);
        Assert.Equal(15, LiveStaleIndexRowTests.NestedListingPollSeconds);
        Assert.True(LiveStaleIndexRowTests.NestedListingPollSeconds < LiveStaleIndexRowTests.NestedListingWaitSeconds);
        Assert.Throws<ArgumentOutOfRangeException>(() => LiveStaleIndexRowTests.WaitForNestedListing(
            Walks(Paths("")), Budget, TimeSpan.Zero, () => TimeSpan.Zero, _ => { }));
    }

    [Fact]
    public void TheDelegateTestWalksTheTreeOnlyThroughTheWait_BeforeItCanRefuse()
    {
        string[] body = TestMethodBody("LiveStaleIndexRowTests.cs", "DelegateHitsInANestedFolder_AreReadable_ViaTheFlatLeafName");
        string joined = string.Join("\n", body);

        // The live run's own bound, its own clock and a real sleep - each handed to the wait.
        Assert.Contains("LiveWaitBudget wait = LiveWaitBudget.OfSeconds(NestedListingWaitSeconds);", joined, StringComparison.Ordinal);
        Assert.Contains("            wait.Budget,", body);
        Assert.Contains("            TimeSpan.FromSeconds(NestedListingPollSeconds),", body);
        Assert.Contains("            () => wait.Elapsed,", body);
        Assert.Contains("            Thread.Sleep);", body);
        Assert.Single(body, l => l.Contains("FindFolderPathsByLeafName(", StringComparison.Ordinal));
        int wait = Array.FindIndex(body, l => l.Contains("WaitForNestedListing(", StringComparison.Ordinal));
        int walk = Array.FindIndex(body, l => l.Contains("FindFolderPathsByLeafName(", StringComparison.Ordinal));
        int refusal = Array.FindIndex(body, l => l.Contains("NestedPaths(matches)", StringComparison.Ordinal));
        Assert.True(wait >= 0 && walk == wait + 1 && refusal > walk,
            $"expected the only tree walk inside the wait, and the wait before the refusal; found lines {wait}, {walk}, {refusal}");
    }

    /// <summary>The live run's own bound, for the fake-clock tests.</summary>
    private static readonly TimeSpan Budget = TimeSpan.FromSeconds(LiveStaleIndexRowTests.NestedListingWaitSeconds);

    private static readonly TimeSpan Poll = TimeSpan.FromSeconds(LiveStaleIndexRowTests.NestedListingPollSeconds);

    /// <summary>A clock that moves only when the wait sleeps, and remembers every sleep.</summary>
    private sealed class FakeClock
    {
        public TimeSpan Now { get; set; }

        public List<TimeSpan> Slept { get; } = new();

        public void Sleep(TimeSpan duration)
        {
            Slept.Add(duration);
            Now += duration;
        }
    }

    /// <summary>
    /// A walk that answers each listing in turn and then repeats the last one - and refuses a
    /// thousandth call, so a wait that lost its bound fails here instead of hanging the suite.
    /// </summary>
    private static Func<IReadOnlyList<IReadOnlyList<string>>> Walks(params IReadOnlyList<IReadOnlyList<string>>[] listings)
    {
        int calls = 0;
        return () =>
        {
            calls++;
            Assert.True(calls < 1000, "the delegate-tree wait walked a thousand times - its bound is gone");
            return listings[Math.Min(calls - 1, listings.Length - 1)];
        };
    }

    /// <summary><c>|</c>-separated matches of <c>/</c>-separated segments, as in the shape theory above.</summary>
    private static IReadOnlyList<IReadOnlyList<string>> Paths(string shape)
    {
        return shape.Length == 0
            ? Array.Empty<IReadOnlyList<string>>()
            : shape.Split('|').Select(m => (IReadOnlyList<string>)m.Split('/')).ToList();
    }

    // ================================================================================== helpers

    /// <summary>
    /// The line where <c>LiveDisconnectRecoveryTests</c> first looks at Outlook's windows - code, not
    /// a comment, and since Q101 3(a) the line straight after the two S7 counts.
    /// </summary>
    private const string WindowsBaseline = "IReadOnlyList<IntPtr> baselineWindows = WindowProbe.VisibleOutlookWindows();";

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

    /// <summary>
    /// A live test's source as it is on disk now. A bare name is under <c>T2/</c>; a name with a
    /// folder (<c>T3/...</c>) is relative to the test project.
    /// </summary>
    private static string LiveSource(string fileName)
    {
        string dir = typeof(LiveTestSettings).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");
        string path = fileName.Contains('/', StringComparison.Ordinal)
            ? Path.Combine(dir, fileName.Replace('/', Path.DirectorySeparatorChar))
            : Path.Combine(dir, "T2", fileName);
        Assert.True(File.Exists(path), "live test source is missing: " + path);
        return File.ReadAllText(path);
    }
}
