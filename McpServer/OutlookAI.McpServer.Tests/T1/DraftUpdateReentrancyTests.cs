using System.Reflection;

using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// <c>update_draft</c> is RE-ENTRANT: repeating an interrupted call finishes it instead of
/// performing it twice.
/// <para>
/// THE DEFECT THIS PINS. One contract method is not one COM call. <c>TryUpdateDraft</c> is
/// roughly twenty sequential cross-process calls with no transaction, and when an operation
/// misses its deadline the supervisor terminates the COM host - which lands BETWEEN two of
/// those calls far more often than inside one, and skips every <c>finally</c> block written
/// to leave the draft untouched. The user-visible harm, in order: attachments removed and
/// not re-added, which is loss of the user's own files; a subject assigned without the
/// conversation-index restore, which detaches the draft from its thread. Until this work the
/// caller was told the outcome was unknown and NOT to retry, because a retry removed what the
/// first attempt had added and doubled what it had attached.
/// </para>
/// <para>
/// THE KILL IS SIMULATED, and it has to be: reaching the real window needs a real Outlook
/// wedged for the length of an operation deadline. A stand-in session throws where the child
/// would have died, so the parent-side record, the pre-image and the resumed call are all
/// observed exactly as production would produce them; and the step that a repeat could get
/// wrong is pure logic, pinned here state by state rather than inferred.
/// </para>
/// <para>
/// No Outlook and no mailbox anywhere in this file.
/// </para>
/// </summary>
public sealed class DraftUpdateReentrancyTests : IDisposable
{
    /// <summary>A plausible bare EntryID: hex, even length, long enough to be accepted as one.</summary>
    private const string DraftId = "AB01CD02EF03AB04CD05EF06AB07CD08EF09AB10CD11EF12AB13CD14EF15AB16CD17EF18AB19CD20"
        + "EF21AB22CD23EF24AB25CD26EF27AB28CD29EF30AB31CD32EF33AB34CD35EF36AB37CD38EF39AB40";

    private const string OriginalIndex = "01CA0000000000000000000000000000000000000000000000ABCDEF";
    private const string OriginalTopic = "The original thread";

    private readonly string _directory;

    public DraftUpdateReentrancyTests()
    {
        _directory = Path.Combine(Path.GetTempPath(), "OutlookAI-McpTest-t1reentry-" + Guid.NewGuid().ToString("N")[..12]);
        _ = Directory.CreateDirectory(_directory);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(_directory, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    // ---------------------------------------------------------------- the replayability classification

    [Fact]
    public void AFirstAttempt_IsTheIdentityCase_SoNothingAboutOrdinaryUpdatesChanges()
    {
        // Before and now are the same list when nothing has been applied yet, and the plan
        // then reduces to remove-every-match plus add-everything - which is exactly what the
        // two loops it replaced did. The re-entrant path is the same rule evaluated against a
        // draft that moved, not a second mode with its own semantics.
        string[] onTheDraft = { "offer.pdf", "terms.txt" };

        DraftAttachmentWork plan = DraftAttachmentPlan.Build(
            onTheDraft, onTheDraft, new[] { @"C:\x\notes.txt" }, new[] { "terms.txt" });

        Assert.Equal(new[] { @"C:\x\notes.txt" }, plan.PathsToAdd);
        Assert.Equal("terms.txt", Assert.Single(plan.Removals).FileName);
        Assert.Equal(1, plan.Removals[0].Count);
        Assert.Empty(plan.AlreadyAttached);
        Assert.Empty(plan.AlreadyRemoved);
    }

    [Fact]
    public void AFileTheKilledAttemptAlreadyAttached_IsNotAttachedTwice()
    {
        // The accumulating step. Adding is the only part of the sequence where running it
        // again leaves something running it once would not.
        DraftAttachmentWork plan = DraftAttachmentPlan.Build(
            namesBefore: new[] { "offer.pdf" },
            namesNow: new[] { "offer.pdf", "notes.txt" },
            addPaths: new[] { @"C:\x\notes.txt" },
            removeNames: Array.Empty<string>());

        Assert.Empty(plan.PathsToAdd);
        Assert.Equal("notes.txt", Assert.Single(plan.AlreadyAttached));
        Assert.Empty(plan.Removals);
    }

    [Fact]
    public void OnlyTheFilesTheKilledAttemptDidNotReach_AreStillAttached()
    {
        // Additions run in request order, so the shortfall is the TAIL of the request.
        DraftAttachmentWork plan = DraftAttachmentPlan.Build(
            namesBefore: Array.Empty<string>(),
            namesNow: new[] { "a.txt" },
            addPaths: new[] { @"C:\x\a.txt", @"C:\x\b.txt", @"C:\x\c.txt" },
            removeNames: Array.Empty<string>());

        Assert.Equal(new[] { @"C:\x\b.txt", @"C:\x\c.txt" }, plan.PathsToAdd);
    }

    [Fact]
    public void ADraftThatAlreadyHadAFileOfThatName_StillGetsTheNewOne()
    {
        // The pre-image is what tells "already attached by the first attempt" from "the draft
        // always had one by that name". By name alone the two are indistinguishable, and
        // guessing either way is a silent wrong answer - a lost attachment or a duplicate.
        DraftAttachmentWork plan = DraftAttachmentPlan.Build(
            namesBefore: new[] { "notes.txt" },
            namesNow: new[] { "notes.txt" },
            addPaths: new[] { @"C:\x\notes.txt" },
            removeNames: Array.Empty<string>());

        Assert.Equal(new[] { @"C:\x\notes.txt" }, plan.PathsToAdd);
        Assert.Empty(plan.AlreadyAttached);
    }

    [Fact]
    public void AFileTheKilledAttemptAlreadyRemoved_StaysRemovedAndIsReportedRemoved()
    {
        // The subtractive step is idempotent in itself - removing an absent name does
        // nothing - but the REPORT is not: calling it "not found" would describe this attempt
        // rather than the outcome the caller asked for.
        DraftAttachmentWork plan = DraftAttachmentPlan.Build(
            namesBefore: new[] { "offer.pdf" },
            namesNow: Array.Empty<string>(),
            addPaths: Array.Empty<string>(),
            removeNames: new[] { "offer.pdf" });

        Assert.Empty(plan.Removals);
        Assert.Equal("offer.pdf", Assert.Single(plan.AlreadyRemoved));
    }

    [Fact]
    public void ANameThatMatchedNothingAtAll_IsNotReportedAsRemoved()
    {
        // The other half of the previous test, and the reason it needs the pre-image: a name
        // that was never on the draft must still come back as attachmentsNotFound.
        DraftAttachmentWork plan = DraftAttachmentPlan.Build(
            namesBefore: new[] { "offer.pdf" },
            namesNow: new[] { "offer.pdf" },
            addPaths: Array.Empty<string>(),
            removeNames: new[] { "ghost.pdf" });

        Assert.Empty(plan.Removals);
        Assert.Empty(plan.AlreadyRemoved);
    }

    [Theory]
    // Nothing applied yet: the old copy is still there alone. And "both halves ran, only the
    // new copy is left" is BYTE-IDENTICAL to it - which is the finding, not an oversight:
    // the two states cannot be told apart, so the plan must be right for both at once.
    [InlineData("nothing applied, or fully applied", new[] { "offer.pdf" }, 1)]
    // The kill landed between the addition and the removal - the state the reorder produces.
    [InlineData("added, not yet removed", new[] { "offer.pdf", "offer.pdf" }, 2)]
    public void ReplacingAFile_ConvergesFromEveryStateTheKillCanLeave(string state, string[] namesNow, int expectedRemovals)
    {
        Assert.NotEmpty(state);

        // A name that is both removed and added is the one case the pre-image CANNOT settle,
        // because the old copy and the new copy have the same name. The plan deliberately
        // redoes it - delete every current copy, attach every requested one - which converges
        // from all three states, at the cost of repeating work that may already have run. It
        // is safe to repeat precisely because the source is a file on disk.
        DraftAttachmentWork plan = DraftAttachmentPlan.Build(
            namesBefore: new[] { "offer.pdf" },
            namesNow: namesNow,
            addPaths: new[] { @"C:\x\offer.pdf" },
            removeNames: new[] { "offer.pdf" });

        Assert.Equal(new[] { @"C:\x\offer.pdf" }, plan.PathsToAdd);
        Assert.Equal(expectedRemovals, Assert.Single(plan.Removals).Count);
    }

    [Fact]
    public void RemovalsAreCountedAgainstTheDraftBeforeTheAdditions_WhichIsWhatMakesTheNewOrderSafe()
    {
        // Additions now run FIRST, so the window a kill can land in holds a DUPLICATE rather
        // than a hole. That only stays correct because Attachments.Add appends and the plan's
        // removals name the N lowest-indexed copies, counted before anything was attached.
        DraftAttachmentWork plan = DraftAttachmentPlan.Build(
            namesBefore: new[] { "offer.pdf", "offer.pdf" },
            namesNow: new[] { "offer.pdf", "offer.pdf" },
            addPaths: new[] { @"C:\x\offer.pdf" },
            removeNames: new[] { "offer.pdf" });

        Assert.Equal(2, Assert.Single(plan.Removals).Count);
    }

    [Fact]
    public void NamesAreMatchedCaseInsensitively_LikeOutlooksOwnAttachmentNames()
    {
        DraftAttachmentWork plan = DraftAttachmentPlan.Build(
            namesBefore: new[] { "Offer.PDF" },
            namesNow: new[] { "Offer.PDF" },
            addPaths: Array.Empty<string>(),
            removeNames: new[] { "offer.pdf" });

        Assert.Equal(1, Assert.Single(plan.Removals).Count);
    }

    [Fact]
    public void WithoutAResume_ThePreImageIsTheDraftsOwnCurrentNames()
    {
        // The selection the COM sequence makes, lifted out so it can be reverted: nothing
        // without a real Outlook can reach the call site, and getting it backwards is silent.
        DraftAttachmentWork plan = DraftAttachmentPlan.BuildForAttempt(
            null, new[] { "notes.txt" }, new[] { @"C:\x\notes.txt" }, Array.Empty<string>());

        Assert.Equal(new[] { @"C:\x\notes.txt" }, plan.PathsToAdd);
    }

    [Fact]
    public void WithAResume_TheRECORDEDNamesAreThePreImage()
    {
        DraftAttachmentWork plan = DraftAttachmentPlan.BuildForAttempt(
            new ComDraftUpdateResume(Array.Empty<string>()),
            new[] { "notes.txt" },
            new[] { @"C:\x\notes.txt" },
            Array.Empty<string>());

        Assert.Empty(plan.PathsToAdd);
        Assert.Equal("notes.txt", Assert.Single(plan.AlreadyAttached));
    }

    [Fact]
    public void TheThreadIndexRestored_ComesFromTheRecordWhenThereIsOne()
    {
        // The order matters and the wrong one is invisible: assigning Subject regenerates the
        // index, so preferring the live value on a repeat restores the value the interrupted
        // attempt already destroyed - and still reports the thread as preserved.
        Assert.Equal("live", ComDraftUpdateResume.ThreadIndexFor(null, "live"));
        Assert.Equal("live", ComDraftUpdateResume.ThreadIndexFor(new ComDraftUpdateResume(), "live"));
        Assert.Equal(
            "recorded",
            ComDraftUpdateResume.ThreadIndexFor(new ComDraftUpdateResume(null, "recorded", "topic"), "regenerated"));
        Assert.Equal(
            "topic",
            ComDraftUpdateResume.ThreadTopicFor(new ComDraftUpdateResume(null, "recorded", "topic"), "regenerated"));
    }

    // ---------------------------------------------------------------- what identifies a repeat

    [Fact]
    public void TheKey_IsDerivedFromTheRequest_SoTheCallerSuppliesNothing()
    {
        Assert.Equal(KeyFor("hello"), KeyFor("hello"));
    }

    [Theory]
    [InlineData("a different body")]
    public void AnyChangedArgument_IsANewRequestRatherThanARepeat(string body)
    {
        Assert.NotEqual(KeyFor("hello"), KeyFor(body));
    }

    [Fact]
    public void AnOmittedRecipientListAndAnEmptyOne_DoNotShareAKey()
    {
        // update_draft reads them as "leave alone" and "clear", which are opposite
        // instructions - hashing them the same would let one resume as the other.
        string omitted = DraftUpdateIntents.KeyFor(
            DraftId, null, null, null, null, null, null, null, null,
            Array.Empty<string>(), Array.Empty<string>(), true);
        string empty = DraftUpdateIntents.KeyFor(
            DraftId, null, null, null, Array.Empty<string>(), null, null, null, null,
            Array.Empty<string>(), Array.Empty<string>(), true);

        Assert.NotEqual(omitted, empty);
    }

    [Fact]
    public void ADifferentDraft_IsNeverAResumeEvenWithIdenticalArguments()
    {
        DraftUpdateIntents intents = new DraftUpdateIntents();
        intents.Begin("key", DraftId, new ComDraftUpdateResume());

        Assert.Null(intents.Resume("key", DraftId[..40] + DraftId[40..].Replace('A', 'B')));
        Assert.NotNull(intents.Resume("key", DraftId));
    }

    [Fact]
    public void ACallThatAnswered_LeavesNothingToResume()
    {
        // Two identical calls are NOT automatically a retry. Only a request whose outcome is
        // still unknown is resumable; once a call has answered, an identical one after it is
        // a fresh update and must behave like one.
        DraftUpdateIntents intents = new DraftUpdateIntents();
        intents.Begin("key", DraftId, new ComDraftUpdateResume());
        intents.Settle("key");

        Assert.Null(intents.Resume("key", DraftId));
    }

    [Fact]
    public void AnyOtherUpdateToTheSameDraft_DropsThePendingPreImage()
    {
        // The pre-image describes a draft that no longer exists once something else has
        // rewritten it, and resuming from it would be reasoning from a state that is gone.
        DraftUpdateIntents intents = new DraftUpdateIntents();
        intents.Begin("first", DraftId, new ComDraftUpdateResume());
        intents.Begin("second", DraftId, new ComDraftUpdateResume());

        Assert.Null(intents.Resume("first", DraftId));
        Assert.NotNull(intents.Resume("second", DraftId));
    }

    [Fact]
    public void ARepeatOfTheSameRequest_KeepsTheORIGINALPreImage()
    {
        // A resumed attempt that is itself interrupted must stay resumable from the state
        // before the FIRST attempt. Re-reading it now would capture the damage.
        DraftUpdateIntents intents = new DraftUpdateIntents();
        intents.Begin("key", DraftId, new ComDraftUpdateResume(new[] { "first.txt" }));
        intents.Begin("key", DraftId, new ComDraftUpdateResume(new[] { "second.txt" }));

        ComDraftUpdateResume? resume = intents.Resume("key", DraftId);
        Assert.NotNull(resume);
        Assert.Equal("first.txt", Assert.Single(resume!.AttachmentNamesBefore));
    }

    [Fact]
    public void AStaleIntent_IsNotResumable_BecauseTheUserMayHaveEditedTheDraftSince()
    {
        DateTime now = new DateTime(2026, 8, 19, 3, 0, 0, DateTimeKind.Utc);
        DraftUpdateIntents intents = new DraftUpdateIntents(TimeSpan.FromMinutes(10), () => now);
        intents.Begin("key", DraftId, new ComDraftUpdateResume());
        Assert.NotNull(intents.Resume("key", DraftId));

        now = now.AddMinutes(11);
        Assert.Null(intents.Resume("key", DraftId));
    }

    [Fact]
    public void DiscardingTheDraft_ForgetsItsPendingIntents()
    {
        DraftUpdateIntents intents = new DraftUpdateIntents();
        intents.Begin("key", DraftId, new ComDraftUpdateResume());
        intents.Forget(DraftId);

        Assert.Null(intents.Resume("key", DraftId));
        Assert.Equal(0, intents.PendingCount);
    }

    // ---------------------------------------------------------------- the kill, end to end

    [Fact]
    public void AKilledUpdate_TellsTheCallerToReIssueTheExactCall()
    {
        // The inversion is the whole of the re-entrancy work as the caller experiences it.
        string message = MailService.DescribeUpdateOutcomeUnknown("Outlook did not respond within 300000 ms.", resumable: true);

        Assert.Contains("UNKNOWN", message, StringComparison.Ordinal);
        Assert.Contains("RE-ISSUE THIS EXACT CALL", message, StringComparison.Ordinal);
        Assert.Contains("not attached twice", message, StringComparison.Ordinal);
        Assert.Contains("Outlook did not respond within 300000 ms.", message, StringComparison.Ordinal);
    }

    [Fact]
    public void AKilledUpdateWithNoRecord_StillSaysDoNotRetry()
    {
        // Honest degradation: with no pre-image there is nothing to resume from, and the
        // advice has to go back to what every killed mutation gets.
        string message = MailService.DescribeUpdateOutcomeUnknown("The Outlook COM host is not connected.", resumable: false);

        Assert.Contains("UNKNOWN", message, StringComparison.Ordinal);
        Assert.Contains("Do NOT simply retry it", message, StringComparison.Ordinal);
        Assert.DoesNotContain("RE-ISSUE", message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheFirstAttempt_RecordsThePreImageBeforeItTouchesAnything()
    {
        // "Record intent FIRST" is the whole mechanism, and the way it silently fails is that
        // the record is written after the call that dies. So the pre-image read must be
        // observed BEFORE the mutating call, not merely present.
        RecordingSession session = new RecordingSession { FailNextUpdate = true };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        _ = Assert.Throws<InvalidOperationException>(() => service.UpdateDraft(DraftId, subject: "New subject"));

        Assert.Equal(
            new[] { nameof(IOutlookSession.TryGetMailInfo), nameof(IOutlookSession.TryUpdateDraft) },
            session.Calls);
    }

    [Fact]
    public void ARepeatOfAKilledUpdate_CarriesThePreImageIntoTheComLayer()
    {
        // The end-to-end shape: attempt one dies where the COM host would have been killed,
        // attempt two arrives with the state the dead attempt saw and could not report.
        RecordingSession session = new RecordingSession { FailNextUpdate = true };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        _ = Assert.Throws<InvalidOperationException>(() => service.UpdateDraft(DraftId, subject: "New subject"));
        Assert.Null(session.LastResume);

        session.FailNextUpdate = false;
        UpdateDraftOutcome outcome = service.UpdateDraft(DraftId, subject: "New subject");

        Assert.NotNull(session.LastResume);
        Assert.Equal(OriginalIndex, session.LastResume!.ConversationIndex);
        Assert.Equal(OriginalTopic, session.LastResume!.ConversationTopic);
        Assert.True(outcome.Resumed);
        Assert.NotNull(outcome.ResumedAdvice);
    }

    [Fact]
    public void AResumeIsOfferedOnlyForTheSameRequest()
    {
        // A different request against the same draft is a different intention, and completing
        // the earlier one under its name would apply something nobody asked for now.
        RecordingSession session = new RecordingSession { FailNextUpdate = true };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        _ = Assert.Throws<InvalidOperationException>(() => service.UpdateDraft(DraftId, subject: "New subject"));

        session.FailNextUpdate = false;
        _ = service.UpdateDraft(DraftId, subject: "A different subject");

        Assert.Null(session.LastResume);
    }

    [Fact]
    public void AnUpdateThatAnswered_LeavesNoResumeForTheNextIdenticalCall()
    {
        RecordingSession session = new RecordingSession();
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        _ = service.UpdateDraft(DraftId, subject: "New subject");
        _ = service.UpdateDraft(DraftId, subject: "New subject");

        Assert.Null(session.LastResume);
    }

    [Fact]
    public void ANamedRefusal_SettlesTheIntent_BecauseItProvesNothingWasChanged()
    {
        // Every named refusal is decided before anything is written, so there is nothing left
        // to complete and no resume should be offered for it.
        RecordingSession session = new RecordingSession { UpdateRefusal = "AlreadySent" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        _ = Assert.Throws<DraftRefusedException>(() => service.UpdateDraft(DraftId, subject: "New subject"));

        session.UpdateRefusal = null;
        _ = service.UpdateDraft(DraftId, subject: "New subject");

        Assert.Null(session.LastResume);
    }

    [Fact]
    public void AnUnclassifiedComFailure_KeepsTheIntent_BecauseItCanLandPartWayThrough()
    {
        // The one refusal that does NOT prove the draft was left alone: it is raised by the
        // catch-all around the whole ~20-call sequence, so it can arrive after the body has
        // been committed through the inspector or after an attachment has gone. The wording
        // used to claim "Nothing was changed" here, which was simply untrue.
        RecordingSession session = new RecordingSession { UpdateRefusal = "COMException 0x80004005" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        DraftRefusedException refusal = Assert.Throws<DraftRefusedException>(
            () => service.UpdateDraft(DraftId, subject: "New subject"));
        Assert.Contains("UNKNOWN", refusal.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("Nothing was changed", refusal.Message, StringComparison.Ordinal);

        session.UpdateRefusal = null;
        _ = service.UpdateDraft(DraftId, subject: "New subject");

        Assert.NotNull(session.LastResume);
    }

    [Fact]
    public void ADraftNoStoreCouldOpen_KeepsItsOwnMessageAndLeavesNothingPending()
    {
        // "Not found" is the one non-refusal failure that still proves the negative: no store
        // opened the draft, so nothing was applied. Without its own branch it would be
        // swallowed by the unknown-outcome catch and the caller would be told to re-issue a
        // call against an id that does not resolve.
        RecordingSession session = new RecordingSession { UpdateRefusal = ComErrorTokens.ItemNotFound };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        InvalidOperationException failure = Assert.Throws<InvalidOperationException>(
            () => service.UpdateDraft(DraftId, subject: "New subject"));
        Assert.Contains("could not be opened", failure.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("UNKNOWN", failure.Message, StringComparison.Ordinal);

        session.UpdateRefusal = null;
        _ = service.UpdateDraft(DraftId, subject: "New subject");

        Assert.Null(session.LastResume);
    }

    [Fact]
    public void DiscardingTheDraft_DropsTheResumeOfferThatWouldOutliveIt()
    {
        // The pre-image is addressed by EntryID, and a discarded draft's id is dead. Leaving
        // the record would offer to complete an update against an item nothing resolves.
        RecordingSession session = new RecordingSession();
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        // A successful update first, because discard_draft only reaches drafts this server
        // touched - a killed update never gets that far.
        _ = service.UpdateDraft(DraftId, subject: "First");

        session.FailNextUpdate = true;
        _ = Assert.Throws<InvalidOperationException>(() => service.UpdateDraft(DraftId, subject: "Second"));

        session.FailNextUpdate = false;
        _ = service.DiscardDraft(DraftId);
        _ = service.UpdateDraft(DraftId, subject: "Second");

        Assert.Null(session.LastResume);
    }

    [Fact]
    public void AnUpdateThatNeedsNoPreImage_TakesNoExtraReads()
    {
        // Body and recipients are assignment-shaped: writing them twice writes the same
        // value. Nothing about such a request can be got wrong by a repeat, so it must not
        // pay for a pre-image read.
        RecordingSession session = new RecordingSession();
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        _ = service.UpdateDraft(DraftId, body: "rewritten");

        Assert.DoesNotContain(nameof(IOutlookSession.TryGetMailInfo), session.Calls);
    }

    [Fact]
    public void AttachingAFile_TakesTheAttachmentPreImage()
    {
        // Adding is the accumulating step, and "is this file already on?" cannot be answered
        // from the request alone - so this is the request shape that pays for the second read.
        string path = Path.Combine(_directory, "notes.txt");
        File.WriteAllText(path, "payload");

        RecordingSession session = new RecordingSession { AttachmentNames = new[] { "offer.pdf" } };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        _ = service.UpdateDraft(DraftId, attachments: new[] { path });

        Assert.Contains(nameof(IOutlookSession.SnapshotAttachmentsById), session.Calls);
    }

    private static string KeyFor(string body)
    {
        return DraftUpdateIntents.KeyFor(
            DraftId, ComDraftBody.FromText(body), "subject", null, null, null, null, null, null,
            Array.Empty<string>(), Array.Empty<string>(), true);
    }

    /// <summary>Runs operations straight against the stand-in session, with no budget layer.</summary>
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

    /// <summary>
    /// A session that answers the reads, records the order it was called in, and dies on the
    /// update when told to - which is what a killed COM host looks like from the parent: a
    /// <see cref="TimeoutException"/> and no report of what was applied.
    /// </summary>
    private sealed class RecordingSession
    {
        private readonly List<string> _calls = new List<string>();

        internal RecordingSession()
        {
            AsSession = Proxy.Create(this);
        }

        internal IOutlookSession AsSession { get; }

        /// <summary>True to make the next update die the way a killed child does.</summary>
        internal bool FailNextUpdate { get; set; }

        /// <summary>COM-side refusal token to answer the update with, instead of a result.</summary>
        internal string? UpdateRefusal { get; set; }

        /// <summary>What the draft is carrying, for the attachment pre-image.</summary>
        internal IReadOnlyList<string> AttachmentNames { get; set; } = Array.Empty<string>();

        /// <summary>The resume argument the last update was handed.</summary>
        internal ComDraftUpdateResume? LastResume { get; private set; }

        internal IReadOnlyList<string> Calls
        {
            get
            {
                lock (_calls)
                {
                    return _calls.ToList();
                }
            }
        }

        private static ComDraftInfo Snapshot()
        {
            return new ComDraftInfo(
                DraftId, "Work", "store-work", "Drafts", "folder-drafts", "A subject",
                "someone@example.com", OriginalIndex, "conv-1", Array.Empty<ComRecipientInfo>(), OriginalTopic);
        }

        private object? Handle(MethodInfo method, object?[]? args)
        {
            lock (_calls)
            {
                _calls.Add(method.Name);
            }

            switch (method.Name)
            {
                case nameof(IOutlookSession.GetStoreDetails):
                    return Array.Empty<ComStoreDetail>();
                case nameof(IOutlookSession.TryGetMailInfo):
                    return Snapshot();
                case nameof(IOutlookSession.SnapshotAttachmentsById):
                    return AttachmentNames.Select((n, i) => new ComAttachmentInfo(i + 1, n, 10L)).ToList();
                case nameof(IOutlookSession.TryDiscardDraft):
                    return new ComDraftDiscardResult(DraftId, DraftId, "Work", "Drafts", "Deleted Items", "A subject");
                case nameof(IOutlookSession.TryUpdateDraft):
                    LastResume = args?[12] as ComDraftUpdateResume;
                    if (FailNextUpdate)
                    {
                        throw new TimeoutException("Outlook did not respond to 'TryUpdateDraft' within 300000 ms.");
                    }

                    if (UpdateRefusal != null)
                    {
                        SetError(method, args, UpdateRefusal);
                        return null;
                    }

                    return new ComDraftUpdateResult(
                        Snapshot(), new[] { "subject" }, Array.Empty<string>(),
                        AttachmentNames.Select((n, i) => new ComAttachmentInfo(i + 1, n, 10L)).ToList(),
                        Array.Empty<string>(), Array.Empty<string>(), Array.Empty<string>(),
                        false, false, false, null, false, null, true);
                default:
                    return null;
            }
        }

        private static void SetError(MethodInfo method, object?[]? args, string reason)
        {
            ParameterInfo[] parameters = method.GetParameters();
            for (int i = 0; args != null && i < parameters.Length && i < args.Length; i++)
            {
                if (parameters[i].IsOut && string.Equals(parameters[i].Name, "error", StringComparison.Ordinal))
                {
                    args[i] = reason;
                }
            }
        }

        internal class Proxy : DispatchProxy
        {
            private RecordingSession _owner = null!;

            internal static IOutlookSession Create(RecordingSession owner)
            {
                object proxy = Create<IOutlookSession, Proxy>()
                    ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
                ((Proxy)proxy)._owner = owner;
                return (IOutlookSession)proxy;
            }

            /// <inheritdoc />
            protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
            {
                ArgumentNullException.ThrowIfNull(targetMethod);

                object? handled = _owner.Handle(targetMethod, args);

                // Every unassigned out slot is filled: an unassigned one is unboxed on the way
                // back and would fail the call for a reason unrelated to what is under test.
                ParameterInfo[] parameters = targetMethod.GetParameters();
                for (int i = 0; args != null && i < parameters.Length && i < args.Length; i++)
                {
                    if (parameters[i].IsOut && args[i] == null)
                    {
                        Type slot = parameters[i].ParameterType.GetElementType()!;
                        args[i] = slot.IsValueType ? Activator.CreateInstance(slot) : null;
                    }
                }

                if (handled != null)
                {
                    return handled;
                }

                Type returnType = targetMethod.ReturnType;
                return returnType != typeof(void) && returnType.IsValueType
                    ? Activator.CreateInstance(returnType)
                    : null;
            }
        }
    }
}
