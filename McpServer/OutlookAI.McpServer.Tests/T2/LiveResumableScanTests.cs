using System.Diagnostics;

using OutlookAI.Core.Services;

using Xunit;
using Xunit.Abstractions;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// GAP F2's acceptance: a paged exhaustive scan must return the SAME mail as an unpaged one -
/// no item skipped, no item twice.
/// <para>
/// This is the only test that can prove it. T1 drives the real <c>MailService</c> through a
/// stand-in session, so it pins the token's lifetime, the refusals, the payload and the walk
/// order - but a stand-in returns whatever the test tells it to, so it cannot say whether
/// Outlook's folder enumeration is stable, whether <c>Table.Sort</c> succeeds, or whether an
/// unsorted table hands back rows in the same order twice. Those three are the premises the
/// resumption ladder is built on, and only a real profile answers them.
/// </para>
/// <para>
/// READ-ONLY, and scoped to the tiny test-hub store (S2): it searches, it reads hits to
/// surface their EntryIDs from cache, and it touches nothing else. No item is created, moved,
/// edited or deleted, so no test artifact exists to sweep up afterwards.
/// </para>
/// <para>
/// <b>Every test here needs a CHAIN, and says so when there is none (Q76, 2026-09-27).</b> A hub
/// whose scan fits one page of <see cref="SmallPage"/> hands out no resume token, which leaves
/// three of these tests nothing to record, refuse or supersede - and the "paged" scan is one page,
/// so its two sets agree by construction. Those three used to return GREEN there - four early
/// returns between them, each with a log line and nothing else - and the acceptance passed without
/// ever resuming. Every test now asks <see cref="RequireResumeTokens"/>, which is
/// <see cref="LivePopulationCoverage"/>: a refusal on a Production profile, a
/// <c>PROVED NOTHING:</c> line on a Portable one. Pinned by <c>T1/LiveEarlyReturnGuardTests</c>.
/// </para>
/// </summary>
[Collection(LiveCollections.Phase3)]
[Trait("Category", "Live")]
public sealed class LiveResumableScanTests
{
    /// <summary>
    /// A page size small enough to force several pages over the hub corpus. Two rather than
    /// one because a one-item page cannot show a within-page ordering fault.
    /// </summary>
    private const int SmallPage = 2;

    /// <summary>
    /// A hard stop on paging, so a chain that fails to terminate fails the TEST rather than
    /// running until the suite is killed. Far above what the hub corpus needs.
    /// </summary>
    private const int MaxPages = 200;

    private readonly LivePhase3Fixture _fixture;
    private readonly ITestOutputHelper _output;

    public LiveResumableScanTests(LivePhase3Fixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
    }

    /// <summary>
    /// What every test here pages through, named as the Production refusal wraps it: a hub holding
    /// no more than one page hands out no token at all.
    /// </summary>
    internal static readonly string ResumeTokenPopulation =
        $"a resume token from an exhaustive scan of the test hub paged at top {SmallPage}";

    /// <summary>
    /// What <see cref="ASupersededToken_IsRefusedWithThePositionNeededToCarryOnWithoutIt"/> needs on
    /// top of that: the chain still open after its second page, so the first token has been
    /// superseded by something that is still live.
    /// </summary>
    internal static readonly string SecondResumeTokenPopulation =
        $"a second resume token from that scan, which needs more mail in the test hub than two pages at top {SmallPage}";

    /// <summary>What a reader of a PROVED NOTHING line here is to do about it.</summary>
    internal static readonly string ResumeTokenRemedy =
        $"To exercise it, give the test hub at least {(2 * SmallPage) + 1} mail items received since 2000 - more than "
        + $"two pages of {SmallPage}. On a test guest that is the hub population, once Testbed/guest/Reset-HubPopulation.ps1 "
        + "has rebuilt it at the start of the run (Docs/live-tier-on-the-vm.md section 3b); on a working profile the hub "
        + "is real mail, and a scan that ends this soon means it has been emptied.";

    private MailService Service => _fixture.Service;

    private string Hub => _fixture.Settings.TestHubStoreDisplayName;

    /// <summary>
    /// The resume tokens a test here is about to rely on - and the only way this file decides
    /// whether there are any. A null token is a page that ended the chain, so it counts as none.
    /// <para>
    /// <see cref="LivePopulationCoverage.Require"/> is the whole decision: the coverage line on
    /// every run, a refusal on a Production profile when there is nothing, and a
    /// <c>PROVED NOTHING:</c> line on a Portable one. A caller that gets an empty list back has
    /// already been announced, and may only return.
    /// </para>
    /// </summary>
    private IReadOnlyList<string> RequireResumeTokens(
        IEnumerable<string?> tokens, string population, string whatWouldNotRun)
    {
        return LivePopulationCoverage.Require(
            _fixture.Settings,
            tokens.OfType<string>().ToList(),
            population,
            whatWouldNotRun,
            ResumeTokenRemedy,
            _output.WriteLine);
    }

    [Fact]
    [Trait("Requires", "OutlookInstance")]
    [Trait("Writes", "Nothing")]
    public void APagedScan_ReturnsExactlyWhatOneUnpagedScanReturns_WithNoDuplicates()
    {
        // One page that covers the whole scope, as ground truth.
        SearchOutcome whole = Service.Search(NewRequest(top: 100));
        Assert.NotNull(whole.Exhaustive);
        Assert.Equal("complete", whole.Exhaustive!.StopReason);
        Assert.Null(whole.Exhaustive.NextToken);
        Assert.False(whole.Exhaustive.Truncated, "the control run must not hit the result cap");
        Assert.False(whole.Exhaustive.TimedOut, "the control run must not hit the time budget");

        HashSet<string> expected = EntryIdsOf(whole);
        Assert.True(expected.Count > 0, "the hub store returned nothing to page through");

        // The same scope, paged.
        HashSet<string> paged = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        List<string> duplicates = new List<string>();
        List<string> followed = new List<string>();
        string? token = null;
        int pages = 0;
        int itemsReturnedTotal = 0;
        Stopwatch clock = Stopwatch.StartNew();
        while (pages < MaxPages)
        {
            SearchRequest request = NewRequest(SmallPage);
            request.ResumeToken = token;
            SearchOutcome page = Service.Search(request);
            pages++;

            Assert.NotNull(page.Exhaustive);
            itemsReturnedTotal = page.Exhaustive!.ItemsReturnedTotal ?? itemsReturnedTotal;

            foreach (string entryId in EntryIdsOf(page))
            {
                if (!paged.Add(entryId))
                {
                    duplicates.Add(entryId);
                }
            }

            // Every page but the last is honest about being one, and carries the flag the
            // tool description tells an agent to relay.
            if (page.Exhaustive.NextToken == null)
            {
                Assert.Equal("complete", page.Exhaustive.StopReason);
                break;
            }

            Assert.NotEqual("complete", page.Exhaustive.StopReason);
            Assert.True(page.Degraded, "a page with more to come must be degraded - it is not the whole answer");
            Assert.Equal(FreshMerge.FreshnessPartial, page.Freshness);
            Assert.NotNull(page.Exhaustive.Position);
            token = page.Exhaustive.NextToken;
            followed.Add(token);
        }

        clock.Stop();
        _output.WriteLine(
            $"paged scan: pages={pages} top={SmallPage} unique={paged.Count} expected={expected.Count} "
            + $"duplicates={duplicates.Count} itemsReturnedTotal={itemsReturnedTotal} totalMs={clock.ElapsedMilliseconds}");

        Assert.True(pages < MaxPages, $"the chain did not terminate within {MaxPages} pages");

        // A scan that fitted ONE page never resumed, so the two sets below agree by construction and
        // say nothing about resumption. They are still asserted - a one-page answer that differs from
        // the unpaged one is a defect of its own - but the run is told first (Q76).
        _ = RequireResumeTokens(followed, ResumeTokenPopulation, "the paged-versus-unpaged acceptance");

        // THE ACCEPTANCE. Nothing skipped and nothing repeated - the two halves that make a
        // continuation token worth having rather than merely convenient.
        Assert.Empty(duplicates);
        AssertSetsEqual(expected, paged);
    }

    [Fact]
    [Trait("Requires", "OutlookInstance")]
    [Trait("Writes", "Nothing")]
    public void APagedScan_ReportsWhichRungItResumedOn_SoTheSortQuestionIsAnsweredInPassing()
    {
        // position.resumeTier is a cost signal AND evidence: "date" means Table.Sort works on
        // that folder, which is the open question behind the freshness sweep's own item cap.
        // Recorded here rather than asserted, because either answer is a legitimate outcome
        // and the ladder is built so that it changes the cost of a page, never its result.
        SearchRequest request = NewRequest(SmallPage);
        SearchOutcome first = Service.Search(request);

        Assert.NotNull(first.Exhaustive);
        if (RequireResumeTokens(new[] { first.Exhaustive!.NextToken }, ResumeTokenPopulation, "the resume-rung record").Count == 0)
        {
            return;
        }

        ScanPositionInfo position = first.Exhaustive.Position!;
        _output.WriteLine(
            $"resume: folder='{position.ResumeFolder}' within={position.ResumeWithinFolder} "
            + $"tier={position.ResumeTier} cursor={position.ResumeCursorUtc:o} "
            + $"folders={position.FoldersDone}/{position.FoldersTotal} stopReason={first.Exhaustive.StopReason}");

        Assert.Contains(position.ResumeTier, new[] { "date", "ordinal", "restart" });
        Assert.True(position.FoldersTotal > 0, "a scan that stopped must still have counted its scope");
        Assert.True(
            position.FoldersDone <= position.FoldersTotal,
            "folders finished cannot exceed folders in scope");
    }

    [Fact]
    [Trait("Requires", "OutlookInstance")]
    [Trait("Writes", "Nothing")]
    public void AResumeWithAChangedQuestion_IsRefused_AndTheRefusalNamesWhatChanged()
    {
        // The refusal path, against a real chain. Silently honouring it would answer a
        // different question under a claim of continuity; silently ignoring it would restart
        // the scan while the caller believed it was continuing.
        SearchOutcome first = Service.Search(NewRequest(SmallPage));
        Assert.NotNull(first.Exhaustive);
        if (RequireResumeTokens(new[] { first.Exhaustive!.NextToken }, ResumeTokenPopulation, "the changed-question refusal").Count == 0)
        {
            return;
        }

        SearchRequest changed = NewRequest(SmallPage);
        changed.ResumeToken = first.Exhaustive.NextToken;
        changed.AfterUtc = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);

        ArgumentException refused = Assert.Throws<ArgumentException>(() => Service.Search(changed));
        _output.WriteLine("refusal: " + refused.Message);
        Assert.Contains("after changed", refused.Message, StringComparison.Ordinal);

        // And the chain is untouched by the refusal: the original token still works, so a
        // caller who mistyped one argument has not thrown away the work already done.
        SearchRequest correct = NewRequest(SmallPage);
        correct.ResumeToken = first.Exhaustive.NextToken;
        SearchOutcome resumed = Service.Search(correct);
        Assert.True(resumed.Exhaustive!.Resumed);
    }

    [Fact]
    [Trait("Requires", "OutlookInstance")]
    [Trait("Writes", "Nothing")]
    public void ASupersededToken_IsRefusedWithThePositionNeededToCarryOnWithoutIt()
    {
        SearchOutcome first = Service.Search(NewRequest(SmallPage));
        Assert.NotNull(first.Exhaustive);
        if (RequireResumeTokens(new[] { first.Exhaustive!.NextToken }, ResumeTokenPopulation, "the superseded-token refusal").Count == 0)
        {
            return;
        }

        SearchRequest second = NewRequest(SmallPage);
        second.ResumeToken = first.Exhaustive.NextToken;
        SearchOutcome next = Service.Search(second);

        // A chain that FINISHED on its second page has nothing left that the first token could be
        // superseded by, so the replay below would be answered by a different refusal, or none.
        if (RequireResumeTokens(new[] { next.Exhaustive!.NextToken }, SecondResumeTokenPopulation, "the superseded-token refusal").Count == 0)
        {
            return;
        }

        SearchRequest replay = NewRequest(SmallPage);
        replay.ResumeToken = first.Exhaustive.NextToken;
        ArgumentException refused = Assert.Throws<ArgumentException>(() => Service.Search(replay));
        _output.WriteLine("refusal: " + refused.Message);

        Assert.Contains("superseded", refused.Message, StringComparison.Ordinal);
        Assert.Contains("folder(s)", refused.Message, StringComparison.Ordinal);
    }

    private SearchRequest NewRequest(int top)
    {
        return new SearchRequest
        {
            Exhaustive = true,
            Store = Hub,

            // The bound an exhaustive scan demands, wide enough to include the whole hub
            // corpus so the paged and unpaged runs answer the same question.
            AfterUtc = new DateTime(2000, 1, 1, 0, 0, 0, DateTimeKind.Utc),
            Top = top,
            SnippetChars = 0,
        };
    }

    /// <summary>
    /// The EntryIDs behind one answer's hits. Exhaustive hits carry real EntryIDs from birth,
    /// so <c>read</c> serves them from cache and this costs no locate.
    /// </summary>
    private HashSet<string> EntryIdsOf(SearchOutcome outcome)
    {
        HashSet<string> ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (HitSummary hit in outcome.Hits)
        {
            if (HubCorpus.IsTestArtifact(hit.Subject))
            {
                continue;
            }

            ReadOutcome read = Service.Read(hit.Id, maxBodyChars: 0);
            _ = ids.Add(read.EntryId);
        }

        return ids;
    }

    private void AssertSetsEqual(HashSet<string> expected, HashSet<string> actual)
    {
        List<string> missing = expected.Where(id => !actual.Contains(id)).ToList();
        List<string> extra = actual.Where(id => !expected.Contains(id)).ToList();
        if (missing.Count == 0 && extra.Count == 0)
        {
            return;
        }

        // Counts and id PREFIXES only - never a subject or a body (S4).
        Assert.Fail(
            $"the paged scan and the single-page scan disagree: {missing.Count} item(s) only the single page found "
            + $"({string.Join(",", missing.Take(5).Select(Prefix))}), {extra.Count} only the paged run found "
            + $"({string.Join(",", extra.Take(5).Select(Prefix))})");
    }

    private static string Prefix(string entryId)
    {
        return entryId.Length <= 8 ? entryId : entryId.Substring(0, 8);
    }
}
