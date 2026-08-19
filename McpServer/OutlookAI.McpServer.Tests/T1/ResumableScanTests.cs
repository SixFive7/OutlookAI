using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using OutlookAI.Core.Com;
using OutlookAI.Core.IndexSearch;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// GAP F2 - the exhaustive scan truncated in arbitrary tree order with no way to page past
/// it, and the mode chosen BECAUSE completeness matters was the one that could not deliver it.
/// <para>
/// Measured on the maintainer's real Exchange profile: a whole-store 60-day scan reached 3
/// folders of 32 before its budget expired, because one folder holds 108 144 items at roughly
/// 12 items/s. On a local PST the same walk runs at roughly 1 200 items/s and the clock
/// essentially never fires - there <c>top</c> stops it instead. One mechanism has to serve
/// both, and the two stop reasons want different advice, which is why the reason is RECORDED
/// rather than derived from two booleans that can both be true.
/// </para>
/// <para>
/// What this tier can reach: the token's lifetime and every way it is refused, the request
/// fingerprint that decides whether a resume answers the same question, the payload the token
/// rides in, and the walk order every page depends on. What it cannot reach is anything
/// requiring Outlook - whether the folder enumeration is really stable, whether
/// <c>Table.Sort</c> succeeds, whether an unsorted table returns rows in the same order twice.
/// Those are T2's (<c>LiveResumableScanTests</c>), and the ladder is built so that their
/// answers change the COST of a page and never the correctness of one.
/// </para>
/// </summary>
public sealed class ResumableScanTests
{
    private const string Store = "alice@example.com";
    private const string StorePrefix = "file:C:/store/alice";
    private static readonly DateTime Frontier = new DateTime(2026, 8, 1, 12, 0, 0, DateTimeKind.Utc);

    private static readonly ComStoreDetail[] ProfileStores = { new(Store, "store-alice", 0, true) };

    // ===================================================================== the token itself

    [Fact]
    public void AToken_RoundTripsAndCarriesTheWalkStateBackToTheChild()
    {
        // The whole mechanism in one test: page 1 stops with a position, the answer carries a
        // handle, and passing that handle back hands the CHILD exactly the cursor the child
        // produced - which is what makes the state survive a COM-host kill, since the child
        // that dies never held it.
        ScriptedSession scripted = new ScriptedSession();
        scripted.Enqueue(StoppedScan(ComScanStopReasons.TimeBudget, CursorAt("Archive", finished: "F1")));
        scripted.Enqueue(CompleteScan());
        MailService service = Service(scripted);

        SearchOutcome first = service.Search(Request());
        Assert.NotNull(first.Exhaustive!.NextToken);
        Assert.StartsWith("scan-", first.Exhaustive.NextToken!, StringComparison.Ordinal);

        SearchRequest second = Request();
        second.ResumeToken = first.Exhaustive.NextToken;
        SearchOutcome resumed = service.Search(second);

        Assert.Null(scripted.Cursors[0]);
        ComScanCursor sent = Assert.IsType<ComScanCursor>(scripted.Cursors[1]);
        Assert.Equal("archive-entry-id", sent.FolderEntryId);
        Assert.Equal(ScanResumeTierNames.Date, sent.Tier);
        Assert.Equal(new[] { "F1" }, sent.CompletedFolderEntryIds!.ToArray());

        // And the resumed page is honest about being one: complete or not, it is a
        // continuation and says so in both renderings.
        Assert.True(resumed.Exhaustive!.Resumed);
        Assert.Contains(FreshMerge.ScanGapResumed, resumed.Exhaustive.CoverageGaps!);
        Assert.True(resumed.Degraded);

        // The page that FINISHES a chain still reports the chain's total, not its own count.
        // It is the last chance to tell a caller what the whole scan cost them, and the page
        // itself carries one item of the two the chain returned.
        Assert.Equal(2, resumed.Exhaustive.ItemsReturnedTotal);
        Assert.Single(resumed.Hits);
    }

    [Fact]
    public void ANextToken_IsPresentExactlyWhenTheScanStoppedEarly()
    {
        // Pinned in BOTH directions on purpose. "Always present" teaches an agent to loop for
        // ever; "absent when it stopped" teaches it to call a partial answer complete. The
        // field's presence is the ONLY termination signal, because a short page is not one -
        // a walk can stop with zero admitted items and still have most of the store to go.
        MailService stopped = Service(new ScriptedSession(StoppedScan(ComScanStopReasons.ResultCap, CursorAt("Inbox"))));
        Assert.NotNull(stopped.Search(Request()).Exhaustive!.NextToken);

        MailService done = Service(new ScriptedSession(CompleteScan()));
        SearchOutcome outcome = done.Search(Request());
        Assert.Null(outcome.Exhaustive!.NextToken);
        Assert.Null(outcome.Exhaustive.Position);
    }

    [Fact]
    public void AScanThatStoppedWithNoPosition_IssuesNoToken_AndSaysSoRatherThanLookingComplete()
    {
        // The third state, and the one that would otherwise be indistinguishable from
        // success: the walk stopped and no resumable position could be formed. A missing
        // nextToken means "covered its scope" everywhere else, so this case has to say
        // out loud that it does not mean that here.
        MailService service = Service(new ScriptedSession(
            StoppedScan(ComScanStopReasons.TimeBudget, position: null)));

        SearchOutcome outcome = service.Search(Request());

        Assert.Null(outcome.Exhaustive!.NextToken);
        Assert.Equal(ComScanStopReasons.TimeBudget, outcome.Exhaustive.StopReason);
        Assert.Contains(
            outcome.Advice!,
            line => line.Contains("could NOT be made resumable", StringComparison.Ordinal));
    }

    [Fact]
    public void EveryWayATokenFails_HasItsOwnMessage_BecauseTheRemediesDiffer()
    {
        ExhaustiveScanCursors store = new ExhaustiveScanCursors();
        string fingerprint = ExhaustiveScanCursors.FingerprintOf(Request(), new[] { "test" });

        Assert.Equal(
            ScanTokenDecision.Malformed,
            store.Resolve("not-a-token", fingerprint, out ExhaustiveScanSession? none));
        Assert.Null(none);

        Assert.Equal(
            ScanTokenDecision.Unknown,
            store.Resolve("scan-" + new string('a', 32), fingerprint, out _));

        List<string> messages = new List<string>
        {
            MailService.DescribeResumeRefusal(ScanTokenDecision.Malformed, null, null),
            MailService.DescribeResumeRefusal(ScanTokenDecision.Unknown, null, null),
            MailService.DescribeResumeRefusal(ScanTokenDecision.Expired, null, null),
            MailService.DescribeResumeRefusal(ScanTokenDecision.Superseded, null, null),
            MailService.DescribeResumeRefusal(ScanTokenDecision.RequestChanged, null, new[] { "after" }),
        };

        Assert.Equal(messages.Count, messages.Distinct(StringComparer.Ordinal).Count());
        foreach (string message in messages)
        {
            Assert.Contains("resume_token", message, StringComparison.Ordinal);
        }

        // Each names the way back rather than only the problem.
        Assert.Contains("32 hex", messages[0], StringComparison.Ordinal);
        Assert.Contains("exhaustive.position", messages[1], StringComparison.Ordinal);
        Assert.Contains("expired", messages[2], StringComparison.Ordinal);
        Assert.Contains("superseded", messages[3], StringComparison.Ordinal);
        Assert.Contains("after changed", messages[4], StringComparison.Ordinal);
    }

    [Fact]
    public void AnExpiredToken_IsRefusedAsExpired_AndNotAsUnknown()
    {
        // Two different sentences for two different situations: "your chain aged out" points
        // at re-running from position, "this server has never seen it" points at a restart.
        // The clock is injected because the whole point of a time-to-live is that it elapses.
        DateTime now = new DateTime(2026, 8, 19, 3, 0, 0, DateTimeKind.Utc);
        ExhaustiveScanCursors store = new ExhaustiveScanCursors(TimeSpan.FromMinutes(30), () => now);
        string fingerprint = ExhaustiveScanCursors.FingerprintOf(Request(), new[] { "test" });
        string token = store.Issue(null, fingerprint, Position(), 12, out ExhaustiveScanSession session);

        Assert.Equal(ScanTokenDecision.Valid, store.Resolve(token, fingerprint, out _));

        now = now.AddMinutes(31);
        Assert.Equal(ScanTokenDecision.Expired, store.Resolve(token, fingerprint, out _));

        // And an expired chain is gone rather than lingering: a second resolve reports it
        // unknown, because the pruning that answered "expired" also removed it.
        Assert.Equal(ScanTokenDecision.Unknown, store.Resolve(token, fingerprint, out _));
        Assert.Equal(0, store.SessionCount);
        Assert.NotNull(session);
    }

    [Fact]
    public void ASupersededToken_IsRefused_AndTheRefusalCarriesThePositionSoRecoveryNeedsNoToken()
    {
        // A lost response is the reason a caller ever replays an older token, and honouring
        // one would be worse than refusing it: the chain's suppression set has already moved
        // past that position, so the replay would suppress exactly the rows it exists to
        // return. Refusing costs nothing only because the refusal carries the way back.
        ExhaustiveScanCursors store = new ExhaustiveScanCursors();
        string fingerprint = ExhaustiveScanCursors.FingerprintOf(Request(), new[] { "test" });
        string first = store.Issue(null, fingerprint, Position(), 40, out ExhaustiveScanSession session);
        string second = store.Issue(session, fingerprint, Position(), 35, out _);

        Assert.NotEqual(first, second);
        Assert.Equal(ScanTokenDecision.Valid, store.Resolve(second, fingerprint, out _));
        Assert.Equal(ScanTokenDecision.Superseded, store.Resolve(first, fingerprint, out ExhaustiveScanSession? live));

        Assert.NotNull(live);
        string message = MailService.DescribeResumeRefusal(ScanTokenDecision.Superseded, live, null);
        Assert.Contains("folder:'Archive'", message, StringComparison.Ordinal);
        Assert.Contains("before:'2026-06-11T08:14:22Z'", message, StringComparison.Ordinal);
        Assert.Contains("75 item(s)", message, StringComparison.Ordinal);
    }

    [Fact]
    public void AResumeWhoseQuestionChanged_IsRefused_AndTheRefusalNamesWhatChanged()
    {
        // Parameterised over every field the fingerprint covers, because a silent honour or a
        // silent ignore both answer a different question under a claim of continuity - and
        // one field quietly dropped from the fingerprint is exactly how that ships.
        SearchRequest baseline = Request();
        string original = ExhaustiveScanCursors.FingerprintOf(baseline, new[] { "test" });

        (string Label, SearchRequest Changed, IReadOnlyList<string> Terms)[] variants =
        {
            ("terms", Request(), new[] { "different" }),
            ("searchIn", With(r => r.SearchIn = SearchIn.SubjectOnly), new[] { "test" }),
            ("store", With(r => r.Store = "bob@example.com"), new[] { "test" }),
            ("folder", With(r => r.Folder = "Archive"), new[] { "test" }),
            ("includeSubfolders", With(r => r.IncludeSubfolders = false), new[] { "test" }),
            ("after", With(r => r.AfterUtc = Frontier.AddDays(-30)), new[] { "test" }),
            ("before", With(r => r.BeforeUtc = Frontier), new[] { "test" }),
            ("from", With(r => r.From = "carol@example.com"), new[] { "test" }),
            ("unreadOnly", With(r => r.UnreadOnly = true), new[] { "test" }),
            ("hasAttachments", With(r => r.HasAttachments = true), new[] { "test" }),
            ("orderBySize", With(r => r.OrderBySizeDescending = true), new[] { "test" }),
        };

        foreach ((string label, SearchRequest changed, IReadOnlyList<string> terms) in variants)
        {
            string other = ExhaustiveScanCursors.FingerprintOf(changed, terms);
            Assert.True(original != other, label + " must change the request fingerprint");

            IReadOnlyList<string> named = ExhaustiveScanCursors.DifferingArguments(original, other);
            Assert.Equal(new[] { label }, named.ToArray());

            string message = MailService.DescribeResumeRefusal(ScanTokenDecision.RequestChanged, null, named);
            Assert.Contains(label, message, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void TheStore_RefusesAResumeWhoseFingerprintDoesNotMatch_RatherThanHonouringIt()
    {
        // The comparison itself, not just the fingerprint that feeds it. Reverting this one
        // line - so a mismatch resolves as Valid - left the whole suite green until this test
        // existed, which is the worst shape a gap can have: every OTHER guard around it was
        // pinned, so the coverage looked complete while the decision that matters was
        // unprotected.
        ExhaustiveScanCursors store = new ExhaustiveScanCursors();
        string original = ExhaustiveScanCursors.FingerprintOf(Request(), new[] { "test" });
        string different = ExhaustiveScanCursors.FingerprintOf(With(r => r.Folder = "Archive"), new[] { "test" });
        string token = store.Issue(null, original, Position(), 7, out _);

        Assert.Equal(ScanTokenDecision.Valid, store.Resolve(token, original, out _));
        Assert.Equal(
            ScanTokenDecision.RequestChanged,
            store.Resolve(token, different, out ExhaustiveScanSession? session));

        // The session comes back with the refusal, so the message can name the way forward -
        // a refusal that also loses the position would cost the whole scan.
        Assert.NotNull(session);
        Assert.Equal(original, session!.Fingerprint);

        // And the refusal consumes nothing: the caller who re-sends the ORIGINAL question
        // must still be able to continue.
        Assert.Equal(ScanTokenDecision.Valid, store.Resolve(token, original, out _));
    }

    [Fact]
    public void AChangedResume_IsRefusedEndToEnd_AndLeavesTheChainUsable()
    {
        // The same decision through the real service, because that is where a caller meets
        // it: page one, then the same token with one argument moved. Silently honouring it
        // would answer a different question under a claim of continuity; silently ignoring it
        // would restart the scan while the caller believed they were continuing.
        ScriptedSession scripted = new ScriptedSession();
        scripted.Enqueue(StoppedScan(ComScanStopReasons.TimeBudget, CursorAt("Archive")));
        scripted.Enqueue(CompleteScan());
        MailService service = Service(scripted);

        SearchOutcome first = service.Search(Request());
        Assert.NotNull(first.Exhaustive!.NextToken);

        SearchRequest changed = Request();
        changed.ResumeToken = first.Exhaustive.NextToken;
        changed.AfterUtc = Frontier.AddDays(-30);

        ArgumentException refused = Assert.Throws<ArgumentException>(() => service.Search(changed));
        Assert.Contains("DIFFERENT query", refused.Message, StringComparison.Ordinal);
        Assert.Contains("after changed", refused.Message, StringComparison.Ordinal);

        // The child was never asked anything for the refused call - one scan call so far.
        Assert.Single(scripted.Cursors);

        // And the chain survives the refusal, so a mistyped argument does not throw away the
        // minutes the first page cost.
        SearchRequest correct = Request();
        correct.ResumeToken = first.Exhaustive.NextToken;
        Assert.True(service.Search(correct).Exhaustive!.Resumed);
    }

    [Fact]
    public void AResumeTokenWithoutExhaustive_IsRefused_RatherThanQuietlyIgnored()
    {
        // The silent half of the same failure. Without this the handle simply does nothing:
        // the caller gets a fresh indexed search that answers a different question, reports
        // itself as fresh and complete, and says nothing about the continuation it dropped.
        MailService service = Service(new ScriptedSession(CompleteScan()));
        SearchRequest request = Request();
        request.Exhaustive = false;
        request.ResumeToken = "scan-" + new string('a', 32);

        ArgumentException refused = Assert.Throws<ArgumentException>(() => service.Search(request));
        Assert.Contains("EXHAUSTIVE", refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void PresentationArguments_DoNotChangeTheQuestion_SoAPageMayBeSmallerThanTheLast()
    {
        // top and snippet_chars shape ONE page rather than the result set, so a caller who
        // decides page four should be cheaper must not be told they asked a different
        // question. Everything else being in the fingerprint is what makes this safe.
        string original = ExhaustiveScanCursors.FingerprintOf(Request(), new[] { "test" });
        Assert.Equal(original, ExhaustiveScanCursors.FingerprintOf(With(r => r.Top = 3), new[] { "test" }));
        Assert.Equal(original, ExhaustiveScanCursors.FingerprintOf(With(r => r.SnippetChars = 500), new[] { "test" }));
    }

    [Fact]
    public void AnOmittedArgument_AndOneSetToItsDefault_DoNotFingerprintAlike()
    {
        // Presence-first canonicalisation, the same rule DraftUpdateIntents uses and for the
        // same reason: no value a caller could type may be able to hash as "not supplied".
        string omitted = ExhaustiveScanCursors.FingerprintOf(Request(), null);
        string empty = ExhaustiveScanCursors.FingerprintOf(Request(), Array.Empty<string>());
        Assert.NotEqual(omitted, empty);

        Assert.NotEqual(
            ExhaustiveScanCursors.FingerprintOf(Request(), new[] { "test" }),
            ExhaustiveScanCursors.FingerprintOf(With(r => r.From = null), new[] { "test", "extra" }));
    }

    [Fact]
    public void TheStore_EvictsTheOldestChainAtCapacity_AndReportsTheEvictedOneUnknown()
    {
        // A bound that must not be silent. An evicted chain answers "unknown", which carries
        // the restart advice - as opposed to crashing, or worse, being resolved against
        // another chain's state.
        DateTime now = new DateTime(2026, 8, 19, 3, 0, 0, DateTimeKind.Utc);
        ExhaustiveScanCursors store = new ExhaustiveScanCursors(TimeSpan.FromHours(1), () => now);
        List<string> tokens = new List<string>();
        for (int i = 0; i <= ExhaustiveScanCursors.Capacity; i++)
        {
            now = now.AddSeconds(1);
            tokens.Add(store.Issue(null, "fingerprint-" + i, Position(), 1, out _));
        }

        Assert.Equal(ExhaustiveScanCursors.Capacity, store.SessionCount);
        Assert.Equal(ScanTokenDecision.Unknown, store.Resolve(tokens[0], "fingerprint-0", out _));
        Assert.Equal(
            ScanTokenDecision.Valid,
            store.Resolve(tokens[tokens.Count - 1], "fingerprint-" + ExhaustiveScanCursors.Capacity, out _));
    }

    [Fact]
    public void AChainThatFinished_StopsHoldingItsState()
    {
        // The finished-folder set and the suppression set exist to make the NEXT page
        // possible. There is no next page, so holding them is pure memory - and a token that
        // still resolved would let a caller "continue" a scan with nothing left to do.
        ScriptedSession scripted = new ScriptedSession();
        scripted.Enqueue(StoppedScan(ComScanStopReasons.ResultCap, CursorAt("Inbox")));
        scripted.Enqueue(CompleteScan());
        MailService service = Service(scripted);

        SearchOutcome first = service.Search(Request());
        SearchRequest second = Request();
        second.ResumeToken = first.Exhaustive!.NextToken;
        SearchOutcome finished = service.Search(second);

        Assert.Null(finished.Exhaustive!.NextToken);

        SearchRequest third = Request();
        third.ResumeToken = first.Exhaustive.NextToken;
        ArgumentException refused = Assert.Throws<ArgumentException>(() => service.Search(third));
        Assert.Contains("not known to this server", refused.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TokenShape_IsCheckedBeforeTheStoreIsConsulted()
    {
        Assert.True(ExhaustiveScanCursors.LooksLikeToken("scan-" + new string('0', 32)));
        Assert.False(ExhaustiveScanCursors.LooksLikeToken(null));
        Assert.False(ExhaustiveScanCursors.LooksLikeToken(string.Empty));
        Assert.False(ExhaustiveScanCursors.LooksLikeToken("confirm-" + new string('0', 32)));
        Assert.False(ExhaustiveScanCursors.LooksLikeToken("scan-" + new string('0', 31)));
        Assert.False(ExhaustiveScanCursors.LooksLikeToken("scan-" + new string('A', 32)));
        Assert.False(ExhaustiveScanCursors.LooksLikeToken("scan-" + new string('z', 32)));
    }

    // ===================================================================== the stop reason

    [Fact]
    public void TheStopReason_IsRecordedByTheWalk_AndNotDerivedFromTheBooleans()
    {
        // The reason this is a field at all. Both bounds can be latched on one result, and
        // the remedies point in different directions, so a derivation would pick one by
        // accident of which `if` came first. Here the walk says the CAP ended it while the
        // clock had also run out, and the payload repeats what the walk said.
        MailService service = Service(new ScriptedSession(new ComExhaustiveResult(
            new[] { Mail("EX1") },
            foldersScanned: 3,
            foldersSkipped: 0,
            engine: "ci_phrasematch",
            instantSearchEnabled: true,
            truncated: true,
            timedOut: true,
            stopReason: ComScanStopReasons.ResultCap,
            position: Position())));

        ExhaustiveInfo exhaustive = service.Search(Request()).Exhaustive!;

        Assert.Equal(ComScanStopReasons.ResultCap, exhaustive.StopReason);
        Assert.True(exhaustive.Truncated);
        Assert.True(exhaustive.TimedOut);

        // Both bounds still declare their own hole. The token is a REMEDY, not a repair.
        Assert.Contains(FreshMerge.ScanGapResultCap, exhaustive.CoverageGaps!);
        Assert.Contains(FreshMerge.ScanGapTimeBudget, exhaustive.CoverageGaps!);
    }

    [Fact]
    public void TheDepthGuard_IsNeverAStopReason_BecauseItNeverStopsTheWalk()
    {
        // depthLimitReached bounds one subtree and every sibling branch is still walked, so a
        // scan can report it and still have covered its scope. A payload implying the depth
        // guard ended the walk would send a caller after a broken folder tree that is not the
        // reason their answer is short.
        MailService service = Service(new ScriptedSession(new ComExhaustiveResult(
            new[] { Mail("EX1") },
            foldersScanned: 9,
            foldersSkipped: 0,
            engine: "ci_phrasematch",
            instantSearchEnabled: true,
            truncated: false,
            timedOut: false,
            depthLimitReached: true,
            stopReason: ComScanStopReasons.Complete)));

        ExhaustiveInfo exhaustive = service.Search(Request()).Exhaustive!;

        Assert.Equal(ComScanStopReasons.Complete, exhaustive.StopReason);
        Assert.True(exhaustive.DepthLimitReached);
        Assert.Null(exhaustive.NextToken);
        Assert.Contains(FreshMerge.ScanGapDepthLimit, exhaustive.CoverageGaps!);
    }

    [Fact]
    public void EveryStopReason_IsOneOfThreeValues_AndTheSetIsWhatTheWalkCanProduce()
    {
        IReadOnlyList<string> declared = typeof(ComScanStopReasons)
            .GetFields(BindingFlags.Public | BindingFlags.Static)
            .Where(f => f.IsLiteral && f.FieldType == typeof(string))
            .Select(f => (string)f.GetRawConstantValue()!)
            .OrderBy(v => v, StringComparer.Ordinal)
            .ToList();

        Assert.Equal(new[] { "complete", "result_cap", "time_budget" }, declared.ToArray());
        foreach (string value in declared)
        {
            Assert.DoesNotContain(" ", value, StringComparison.Ordinal);
            Assert.Equal(value.ToLowerInvariant(), value);
        }
    }

    // ===================================================================== the payload

    [Fact]
    public void AResumableStop_StillDegradesTheAnswer_BecauseARemedyIsNotARepair()
    {
        // The decision this pins: a scan that CAN be continued is still, as returned,
        // incomplete. Suppressing degraded because a remedy exists is precisely the "looks
        // complete and quietly is not" failure the whole coverage-code system was built
        // against, and it would be invisible - the payload would gain a field and lose a flag.
        MailService service = Service(new ScriptedSession(
            StoppedScan(ComScanStopReasons.TimeBudget, CursorAt("Archive"))));

        SearchOutcome outcome = service.Search(Request());

        Assert.NotNull(outcome.Exhaustive!.NextToken);
        Assert.True(outcome.Degraded);
        Assert.Equal(FreshMerge.FreshnessPartial, outcome.Freshness);
        Assert.Contains(FreshMerge.ScanGapTimeBudget, outcome.Exhaustive.CoverageGaps!);
    }

    [Fact]
    public void ThePosition_ReportsWhereToCarryOn_InParametersSearchAlreadyHas()
    {
        MailService service = Service(new ScriptedSession(
            StoppedScan(ComScanStopReasons.TimeBudget, CursorAt("Archive"))));

        ScanPositionInfo position = service.Search(Request()).Exhaustive!.Position!;

        Assert.Equal(4, position.FoldersDone);
        Assert.Equal(32, position.FoldersTotal);
        Assert.Equal("Archive", position.ResumeFolder);
        Assert.True(position.ResumeWithinFolder);
        Assert.Equal(new DateTime(2026, 6, 11, 8, 14, 22, DateTimeKind.Utc), position.ResumeCursorUtc);
        Assert.Equal(ScanResumeTierNames.Date, position.ResumeTier);
        Assert.Equal(1, position.Page);
    }

    [Fact]
    public void TheChainCounters_OnlyEverMoveForward()
    {
        // A caller decides when to stop paging from these numbers, so a total that dropped
        // (or a folder count that went backwards after a tree change) would make the one
        // visible cost of paging unreadable.
        ScriptedSession scripted = new ScriptedSession();
        scripted.Enqueue(StoppedScan(ComScanStopReasons.ResultCap, CursorAt("Inbox", foldersDone: 1)));
        scripted.Enqueue(StoppedScan(ComScanStopReasons.ResultCap, CursorAt("Archive", foldersDone: 4)));
        scripted.Enqueue(StoppedScan(ComScanStopReasons.ResultCap, CursorAt("Archive", foldersDone: 9)));
        MailService service = Service(scripted);

        List<int> totals = new List<int>();
        List<int> done = new List<int>();
        List<int> pages = new List<int>();
        string? token = null;
        for (int i = 0; i < 3; i++)
        {
            SearchRequest request = Request();
            request.ResumeToken = token;
            ExhaustiveInfo exhaustive = service.Search(request).Exhaustive!;
            totals.Add(exhaustive.ItemsReturnedTotal!.Value);
            done.Add(exhaustive.Position!.FoldersDone);
            pages.Add(exhaustive.Position.Page);
            token = exhaustive.NextToken;
        }

        Assert.Equal(new[] { 1, 2, 3 }, totals.ToArray());
        Assert.Equal(new[] { 1, 4, 9 }, done.ToArray());
        Assert.Equal(new[] { 1, 2, 3 }, pages.ToArray());
    }

    [Fact]
    public void EveryResumptionCode_NamesARemedyRatherThanOnlyAProblem()
    {
        (string Code, ExhaustiveInfo Scan, string Remedy)[] cases =
        {
            (
                FreshMerge.ScanGapResumed,
                new ExhaustiveInfo { Resumed = true, ItemsReturnedTotal = 300 },
                "per page"),
            (
                FreshMerge.ScanGapTreeChanged,
                new ExhaustiveInfo { TreeChangedFolders = 3 },
                "'after' bound"),
            (
                FreshMerge.ScanGapResumedUnsorted,
                new ExhaustiveInfo { ResumedUnsorted = true },
                "narrower 'after'/'before' window"),
            (
                FreshMerge.ScanGapResumePositionLost,
                new ExhaustiveInfo { ResumePositionLost = true },
                "restarted with duplicate suppression"),
            (
                FreshMerge.ScanGapDedupCapacity,
                new ExhaustiveInfo { DedupCapacityReached = true },
                "De-duplicate by id"),
        };

        foreach ((string code, ExhaustiveInfo scan, string remedy) in cases)
        {
            scan.CoverageGaps = new[] { code };
            string line = Assert.Single(MailService.DescribeExhaustiveCoverage(scan, top: 25));
            Assert.Contains(remedy, line, StringComparison.Ordinal);
        }
    }

    [Fact]
    public void BothStopReasonSentences_PointAtTheToken_AndSayWhichRemedyIsCheaper()
    {
        // F2's original complaint was that the answer was honest and unactionable: "raise
        // top" cannot help past the cap of 100 and a re-run re-walks the same tree. Both
        // sentences now name the continuation, and both distinguish it from narrowing - which
        // is the cheaper remedy for a cap stop and no remedy at all for a budget stop.
        ExhaustiveInfo capped = new ExhaustiveInfo
        {
            Truncated = true,
            NextToken = "scan-" + new string('a', 32),
            CoverageGaps = new[] { FreshMerge.ScanGapResultCap },
            Position = new ScanPositionInfo { ResumeFolder = "Archive" },
        };
        string capLine = Assert.Single(MailService.DescribeExhaustiveCoverage(capped, top: 100));
        Assert.Contains("resume_token", capLine, StringComparison.Ordinal);
        Assert.Contains("Archive", capLine, StringComparison.Ordinal);
        Assert.Contains("CHEAPER", capLine, StringComparison.Ordinal);

        ExhaustiveInfo expired = new ExhaustiveInfo
        {
            TimedOut = true,
            FoldersScanned = 3,
            NextToken = "scan-" + new string('b', 32),
            CoverageGaps = new[] { FreshMerge.ScanGapTimeBudget },
        };
        string budgetLine = Assert.Single(MailService.DescribeExhaustiveCoverage(expired, top: 100));
        Assert.Contains("resume_token", budgetLine, StringComparison.Ordinal);
        Assert.Contains("Nothing is cheaper", budgetLine, StringComparison.Ordinal);
    }

    // ===================================================================== the walk order

    [Fact]
    public void SiblingOrder_IsByNameThenByCollectionPosition_SoTwoWalksCannotDisagree()
    {
        // A continuation token is only as correct as the order it resumes into, and Microsoft
        // documents NO ordering for Folder.Folders - so the order exists only because this
        // comparator imposes it. One comparator, shared by the folder listing and the
        // resumable scan: two copies could drift into two different "stable" orders, and a
        // token would then resume into a tree the caller was never shown.
        List<(string Name, int Index)> siblings = new List<(string, int)>
        {
            ("Zebra", 1),
            ("archive", 2),
            ("Archive", 3),
            ("Inbox", 4),
        };

        siblings.Sort(OutlookComSession.CompareSiblings);

        Assert.Equal(
            new[] { "archive", "Archive", "Inbox", "Zebra" },
            siblings.Select(s => s.Name).ToArray());

        // The tiebreak, and why it is not decoration: two siblings may share a name, and
        // without a second key the sort is unstable exactly where resumption needs it - the
        // pair could swap between pages, so one folder is scanned twice and the other never.
        Assert.True(OutlookComSession.CompareSiblings(("Archive", 2), ("Archive", 3)) < 0);
        Assert.True(OutlookComSession.CompareSiblings(("Archive", 3), ("archive", 2)) > 0);
        Assert.Equal(0, OutlookComSession.CompareSiblings(("Archive", 3), ("archive", 3)));
    }

    [Fact]
    public void TheResumeDateBound_IsInclusive_AndDoesNotReplaceTheCallersOwnBeforeBound()
    {
        // The caller's 'before' is exclusive and stays exactly as they wrote it; the resume
        // cursor has to ADMIT its own instant, or items sharing it would be unreachable. The
        // ones already returned at that instant are excluded by id afterwards, because a date
        // alone cannot separate them.
        DateTime before = new DateTime(2026, 8, 1, 0, 0, 0, DateTimeKind.Utc);
        DateTime cursor = new DateTime(2026, 6, 11, 8, 14, 22, DateTimeKind.Utc);

        string plain = ExhaustiveDaslFilter.Build(
            new[] { "test" }, null, before, ExhaustiveEngine.Like, SearchIn.SubjectAndBody);
        string resumed = ExhaustiveDaslFilter.Build(
            new[] { "test" }, null, before, ExhaustiveEngine.Like, SearchIn.SubjectAndBody, cursor);

        Assert.Contains("\"urn:schemas:httpmail:datereceived\" < '2026-08-01", plain, StringComparison.Ordinal);
        Assert.DoesNotContain(" <= '", plain, StringComparison.Ordinal);

        Assert.Contains("\"urn:schemas:httpmail:datereceived\" < '2026-08-01", resumed, StringComparison.Ordinal);
        Assert.Contains(
            "\"urn:schemas:httpmail:datereceived\" <= '2026-06-11 08:14:22'", resumed, StringComparison.Ordinal);
    }

    // ===================================================================== fixtures

    /// <summary>The tier names as the payload spells them, kept beside the tests that read them.</summary>
    private static class ScanResumeTierNames
    {
        internal const string Date = "date";
    }

    private static SearchRequest Request()
    {
        return new SearchRequest
        {
            Query = "test",
            Store = Store,

            // An exhaustive scan refuses to run unbounded, so the request carries the bound
            // its own validation demands.
            Folder = "Inbox",
            Exhaustive = true,
            Top = 25,
            SnippetChars = 0,
        };
    }

    private static SearchRequest With(Action<SearchRequest> change)
    {
        SearchRequest request = Request();
        change(request);
        return request;
    }

    private static ComScanPosition Position()
    {
        return CursorAt("Archive");
    }

    private static ComScanPosition CursorAt(string folder, string? finished = null, int foldersDone = 4)
    {
        List<string> completed = new List<string>();
        if (finished != null)
        {
            completed.Add(finished);
        }

        ComScanCursor cursor = new ComScanCursor(
            completed,
            folder.ToLowerInvariant() + "-entry-id",
            "date",
            new DateTime(2026, 6, 11, 8, 14, 22, DateTimeKind.Utc),
            new[] { "TIE1" });

        return new ComScanPosition(
            cursor,
            foldersDone,
            32,
            folder,
            resumeWithinFolder: true,
            resumeCursorUtc: new DateTime(2026, 6, 11, 8, 14, 22, DateTimeKind.Utc),
            resumeTier: "date");
    }

    private static ComExhaustiveResult StoppedScan(string stopReason, ComScanPosition? position)
    {
        return new ComExhaustiveResult(
            new[] { Mail("EX1") },
            foldersScanned: 3,
            foldersSkipped: 0,
            engine: "ci_phrasematch",
            instantSearchEnabled: true,
            truncated: stopReason == ComScanStopReasons.ResultCap,
            timedOut: stopReason == ComScanStopReasons.TimeBudget,
            stopReason: stopReason,
            position: position);
    }

    private static ComExhaustiveResult CompleteScan()
    {
        return new ComExhaustiveResult(
            new[] { Mail("EX2") },
            foldersScanned: 9,
            foldersSkipped: 0,
            engine: "ci_phrasematch",
            instantSearchEnabled: true,
            truncated: false,
            timedOut: false,
            stopReason: ComScanStopReasons.Complete);
    }

    private static ComMailBrief Mail(string entryId)
    {
        return new ComMailBrief(
            entryId: entryId,
            storeDisplayName: Store,
            storeId: "store-alice",
            folderName: "Inbox",
            folderKind: "inbox",
            subject: "a test mail",
            senderName: "Bob",
            senderAddress: "bob@example.com",
            receivedTime: Frontier,
            isRead: true,
            hasAttachments: false,
            sizeBytes: 2048,
            body: "test body");
    }

    private static ComSweepResult Sweep(string? onlyStore)
    {
        return new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 4,
            foldersSkipped: 0,
            sweptFolders: new[] { Store + "/Inbox" },
            perStore: new[] { new ComStoreSweepCounters(Store, 4, 0, 0, 0) });
    }

    private static MailService Service(ScriptedSession scripted)
    {
        return new MailService(
            new DirectGateway(scripted.Build(ProfileStores, Sweep)), null, new StubIndexClient());
    }

    /// <summary>
    /// A scan that answers a scripted sequence of results and REMEMBERS the cursor it was
    /// handed on each call. The second half is what makes the round trip provable: the token
    /// is opaque to the parent, so the only way to show that page two continues page one is
    /// to look at what actually crossed into the child.
    /// </summary>
    private sealed class ScriptedSession
    {
        private readonly Queue<ComExhaustiveResult> _results = new Queue<ComExhaustiveResult>();

        internal ScriptedSession()
        {
        }

        internal ScriptedSession(ComExhaustiveResult single)
        {
            _results.Enqueue(single);
        }

        internal List<ComScanCursor?> Cursors { get; } = new List<ComScanCursor?>();

        internal void Enqueue(ComExhaustiveResult result)
        {
            _results.Enqueue(result);
        }

        internal IOutlookSession Build(
            IReadOnlyList<ComStoreDetail> stores, Func<string?, ComSweepResult> sweep)
        {
            return StandInSession.Create(stores, sweep, this);
        }

        internal ComExhaustiveResult Next(ComScanCursor? cursor)
        {
            Cursors.Add(cursor);
            return _results.Count > 0
                ? _results.Dequeue()
                : throw new InvalidOperationException("The script ran out of scan results.");
        }
    }

    /// <summary>An index that knows this one store and holds no search rows.</summary>
    private sealed class StubIndexClient : IIndexClient
    {
        private const string DiscoveryTail = " System.ItemUrl FROM SystemIndex WHERE System.Kind='email'";

        public IndexProviderKind Provider => IndexProviderKind.OleDb;

        public IReadOnlyList<IReadOnlyDictionary<string, object?>> ExecuteRows(
            string sql, int maxRows, int? commandTimeoutSeconds = null)
        {
            if (sql.EndsWith(DiscoveryTail, StringComparison.Ordinal))
            {
                return Rows(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                {
                    ["System.ItemUrl"] = StorePrefix + "/0/Inbox/sampled-item",
                });
            }

            if (sql.Contains("System.Message.DateReceived FROM SystemIndex", StringComparison.Ordinal))
            {
                return Rows(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                {
                    ["System.Message.DateReceived"] = Frontier,
                });
            }

            if (sql.StartsWith("SELECT TOP 1 System.ItemUrl FROM SystemIndex WHERE", StringComparison.Ordinal))
            {
                return Rows(new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase)
                {
                    ["System.ItemUrl"] = StorePrefix + "/0/Inbox/probed-item",
                });
            }

            return Array.Empty<IReadOnlyDictionary<string, object?>>();
        }

        private static IReadOnlyList<IReadOnlyDictionary<string, object?>> Rows(
            IReadOnlyDictionary<string, object?> row)
        {
            return new[] { row };
        }
    }

    private sealed class DirectGateway : IComGateway
    {
        private readonly IOutlookSession _session;

        internal DirectGateway(IOutlookSession session)
        {
            _session = session;
        }

        public event Action? OutlookGone
        {
            add { }
            remove { }
        }

        public bool IsConnected => true;

        public bool? QuitSinkActive => null;

        public bool ProbeConnected() => true;

        public T Run<T>(Func<IOutlookSession, T> operation) => operation(_session);

        public T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery) => operation(_session);

        public T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false)
            => operation(_session);

        public ComHostDiagnostics GetDiagnostics() => new ComHostDiagnostics("in-process", "ready");

        public void Dispose()
        {
        }
    }

    /// <summary>Answers the store list, the sweep and the scan; refuses everything else.</summary>
    private class StandInSession : DispatchProxy
    {
        private IReadOnlyList<ComStoreDetail> _stores = Array.Empty<ComStoreDetail>();
        private Func<string?, ComSweepResult> _sweep = _ => throw new NotSupportedException();
        private ScriptedSession? _scan;

        internal static IOutlookSession Create(
            IReadOnlyList<ComStoreDetail> stores,
            Func<string?, ComSweepResult> sweep,
            ScriptedSession scan)
        {
            object proxy = Create<IOutlookSession, StandInSession>()
                ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
            ((StandInSession)proxy)._stores = stores;
            ((StandInSession)proxy)._sweep = sweep;
            ((StandInSession)proxy)._scan = scan;
            return (IOutlookSession)proxy;
        }

        /// <inheritdoc />
        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            switch (targetMethod?.Name)
            {
                case nameof(IOutlookSession.GetProfileName):
                    return "T1 stand-in profile";

                case nameof(IOutlookSession.GetStoreDetails):
                    return _stores;

                // Argument 3 is onlyStoreDisplayName; see IOutlookSession.SweepFoldersNewerThan.
                case nameof(IOutlookSession.SweepFoldersNewerThan):
                    return _sweep(args?[3] as string);

                // Argument 9 is resumeFrom; see IOutlookSession.ExhaustiveScan.
                case nameof(IOutlookSession.ExhaustiveScan):
                    return _scan != null
                        ? _scan.Next(args != null && args.Length > 9 ? args[9] as ComScanCursor : null)
                        : throw new NotSupportedException("This fixture has no exhaustive scan.");

                default:
                    throw new NotSupportedException(
                        "The stand-in session does not implement " + (targetMethod?.Name ?? "?") + ".");
            }
        }
    }
}
