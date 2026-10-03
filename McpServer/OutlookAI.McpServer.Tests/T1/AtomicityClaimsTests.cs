using System.Reflection;
using System.Text.Json;

using OutlookAI.ComHost.Supervision;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;
using OutlookAI.McpServer.Tools;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Every sentence in this product that says something did NOT happen, pinned.
/// <para>
/// THE DEFECT THIS FILE EXISTS FOR. A read-only audit on 2026-08-19 enumerated the 31 claims
/// of atomicity or non-effect the product makes and found sixteen of them wrong - all wrong
/// the same way. A sentence written for the refusals it sat next to had been attached to a
/// catch-all that reached much further, so a call that failed AFTER changing mail reported
/// that nothing had changed. Nothing in the suite could notice, because the claim lived only
/// in prose and prose is the one thing no test checks by accident.
/// </para>
/// <para>
/// The sharpest example is the reason this file leads with it. <c>db34923</c> corrected
/// <c>update_draft</c>'s message so an unclassified COM failure stops claiming "Nothing was
/// changed", and pinned it - by asserting on <c>refusal.Message</c>. One layer up, the tool
/// surface attached a single advice string to every reason code, so the payload that reached
/// the wire said UNKNOWN in its message and "Nothing was changed or deleted." in its advice.
/// The fix had been undone before it shipped, by a field nothing asserted on.
/// </para>
/// <para>
/// No Outlook and no mailbox anywhere in this file.
/// </para>
/// </summary>
public sealed class AtomicityClaimsTests
{
    /// <summary>A plausible bare EntryID: hex, even length, long enough to be accepted as one.</summary>
    private const string DraftId = "AB01CD02EF03AB04CD05EF06AB07CD08EF09AB10CD11EF12AB13CD14EF15AB16CD17EF18AB19CD20"
        + "EF21AB22CD23EF24AB25CD26EF27AB28CD29EF30AB31CD32EF33AB34CD35EF36AB37CD38EF39AB40";

    private const string ItemId = "01AA02BB03CC04DD05EE06FF07AA08BB09CC10DD11EE12FF13AA14BB15CC16DD17EE18FF19AA20BB"
        + "21CC22DD23EE24FF25AA26BB27CC28DD29EE30FF31AA32BB33CC34DD35EE36FF37AA38BB39CC40DD";

    // ---------------------------------------------------------------- row 1: the advice that undid a shipped fix

    [Fact]
    public void AnInterruptedUpdate_DoesNotClaimNothingChanged_InItsADVICEEither()
    {
        // THE ASSERTION THAT WAS MISSING. The message was pinned and the advice was not, so
        // one payload carried both answers to the same question.
        string advice = OutlookTools.DraftRefusalAdvice(MailService.ComFailureRefusal);

        Assert.DoesNotContain("Nothing was changed", advice, StringComparison.Ordinal);
        Assert.Contains("UNKNOWN", advice, StringComparison.Ordinal);
        Assert.Equal(MutationOutcome.Unknown, OutlookTools.DraftRefusalOutcome(MailService.ComFailureRefusal));
    }

    [Theory]
    [InlineData("not_a_mail_item")]
    [InlineData("not_an_unsent_draft")]
    [InlineData("not_in_drafts_folder")]
    [InlineData("drafts_folder_unreadable")]
    [InlineData("compose_surface_unavailable")]
    [InlineData("signature_file_missing")]
    [InlineData("not_created_by_this_server")]
    public void ANamedRefusal_KeepsTheClaimItEarns(string reason)
    {
        // The other half, and it matters just as much: every named refusal IS decided before
        // the first write, so removing the claim from all of them would be a different bug -
        // an agent told to go looking for changes that provably were not made.
        Assert.Contains("Nothing was changed or deleted", OutlookTools.DraftRefusalAdvice(reason), StringComparison.Ordinal);
        Assert.Equal(MutationOutcome.Unchanged, OutlookTools.DraftRefusalOutcome(reason));
    }

    // ---------------------------------------------------------------- the shared opening sentence

    [Fact]
    public void AnInterruptedREAD_SaysNothingChanged()
    {
        Assert.Equal(MutationOutcome.Unchanged, MutationOutcome.ForInterrupted(nameof(IOutlookSession.TryReadItem)));
        Assert.Contains("READS", MutationOutcome.DescribeInterrupted(nameof(IOutlookSession.TryReadItem)), StringComparison.Ordinal);
        Assert.DoesNotContain("UNKNOWN", MutationOutcome.DescribeInterrupted(nameof(IOutlookSession.TryReadItem)), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(nameof(IOutlookSession.TrySendDraft))]
    [InlineData(nameof(IOutlookSession.TryUpdateDraft))]
    [InlineData(nameof(IOutlookSession.TryDiscardDraft))]
    [InlineData(nameof(IOutlookSession.TryMoveItemToPath))]
    [InlineData(nameof(IOutlookSession.TryCreateNewDraft))]
    public void AnInterruptedMUTATION_SaysTheOutcomeIsUnknown(string operation)
    {
        Assert.Equal(MutationOutcome.Unknown, MutationOutcome.ForInterrupted(operation));
        Assert.Contains("UNKNOWN", MutationOutcome.DescribeInterrupted(operation), StringComparison.Ordinal);
    }

    [Fact]
    public void AnUnclassifiedOperation_ReadsAsAMutation_BecauseForgettingMustFailClosed()
    {
        // Inherited from ComSessionOperations and worth pinning here too: the failure mode of
        // adding a contract method without classifying it must be an over-cautious answer,
        // never a false "nothing happened".
        Assert.Equal(MutationOutcome.Unknown, MutationOutcome.ForInterrupted("TryInventANewOperation"));
        Assert.Equal(MutationOutcome.Applied, MutationOutcome.ForCompleted("TryInventANewOperation"));
    }

    [Fact]
    public void AnAnswerTooLargeToReturn_NeverSaysNothingWasChangedOverAMutation()
    {
        // Row 3: the only claim in the product that asserted a mutation both happened and did
        // not. The work runs to completion and only the FRAMING of the reply fails, so
        // "the work itself succeeded and nothing was changed" was self-contradictory.
        string mutating = MutationOutcome.DescribeAnswerLost(nameof(IOutlookSession.TryCreateNewDraft));

        Assert.DoesNotContain("nothing was changed", mutating, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("SUCCEEDED", mutating, StringComparison.Ordinal);
        Assert.Contains("do NOT repeat it", mutating, StringComparison.Ordinal);
        Assert.Equal(MutationOutcome.Applied, MutationOutcome.ForCompleted(nameof(IOutlookSession.TryCreateNewDraft)));

        // A read keeps the old, true, sentence.
        string reading = MutationOutcome.DescribeAnswerLost(nameof(IOutlookSession.ExhaustiveScan));
        Assert.Contains("nothing was changed", reading, StringComparison.Ordinal);
        Assert.Equal(MutationOutcome.Unchanged, MutationOutcome.ForCompleted(nameof(IOutlookSession.ExhaustiveScan)));
    }

    [Fact]
    public void ABareComFailure_StopsSayingRetry_WhenTheCallChangesMail()
    {
        // Row 15. "Outlook rejected the operation; check outlook_health and retry" was the
        // advice for everything alike, including a send.
        Assert.Contains("retry", OutlookTools.ComFailureAdvice(nameof(IOutlookSession.TryReadItem)), StringComparison.OrdinalIgnoreCase);

        string mutating = OutlookTools.ComFailureAdvice(nameof(IOutlookSession.TrySendDraft));
        Assert.Contains("UNKNOWN", mutating, StringComparison.Ordinal);
        Assert.Contains("do not", mutating, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void AFailureThatCannotBeAttributed_StatesNoOutcomeAtAll()
    {
        // Absent is deliberately NOT the same as "unchanged": a request that dispatched no
        // COM call may still have failed for a reason this server cannot classify, and
        // guessing "unchanged" is exactly the habit the whole audit was about.
        string advice = OutlookTools.ComFailureAdvice(null);

        Assert.Contains("NOT stated", advice, StringComparison.Ordinal);
        Assert.DoesNotContain("CHANGES mail and it did not answer", advice, StringComparison.Ordinal);
    }

    [Fact]
    public void AKilledMutation_IsNotToldThatTheNextCallStartsClean_AndNothingElse()
    {
        // Row 7. The deadline path is the COMMON way out of a killed child and it was the one
        // without the treatment: nine tools got "the next call starts clean", a statement
        // about the HOST that reads as a statement about the mail.
        string mutating = OutlookTools.TimeoutAdvice(nameof(IOutlookSession.TryMoveItemToPath));

        Assert.Contains("UNKNOWN", mutating, StringComparison.Ordinal);
        Assert.Contains("look at the current state", mutating, StringComparison.Ordinal);

        string reading = OutlookTools.TimeoutAdvice(nameof(IOutlookSession.SweepFoldersNewerThan));
        Assert.DoesNotContain("UNKNOWN", reading, StringComparison.Ordinal);
        Assert.DoesNotContain("look at the current state", reading, StringComparison.Ordinal);
    }

    [Fact]
    public void AnAnswerTooLarge_StopsInvitingASecondMutation()
    {
        // Row 3's tool-layer half. "Retrying the SAME request will refuse again" is right for
        // a search and, over an operation that already ran and changed mail, is an
        // instruction to do it twice for nothing.
        string mutating = OutlookTools.ResponseTooLargeAdvice(nameof(IOutlookSession.TryUpdateDraft));

        Assert.DoesNotContain("Retrying the SAME request", mutating, StringComparison.Ordinal);
        Assert.Contains("SUCCEEDED", mutating, StringComparison.Ordinal);

        string reading = OutlookTools.ResponseTooLargeAdvice(nameof(IOutlookSession.ExhaustiveScan));
        Assert.Contains("Retrying the SAME request", reading, StringComparison.Ordinal);
        Assert.Contains("ask for less", reading, StringComparison.Ordinal);
    }

    [Fact]
    public void AnInterruptedRequestsOutcome_TravelsOnTheExceptionThatCarriesTheMessage()
    {
        // Row 24 is the site the audit called the model and its wording is untouched. What it
        // gained is the machine-readable half, because only the raising site knows the
        // operation - and "the host could not be started" is NOT the same claim as "your mail
        // is untouched", which is why the other constructor leaves it null.
        Assert.Equal(
            MutationOutcome.Unknown,
            new ComHostUnavailableException("interrupted", MutationOutcome.Unknown).Outcome);
        Assert.Null(new ComHostUnavailableException("could not start").Outcome);
    }

    // ---------------------------------------------------------------- row 2: the discard claim

    [Fact]
    public void AnInterruptedDiscard_StopsClaimingNothingWasChanged()
    {
        // The route that is open is Delete() itself, sitting bare in TryDiscardDraft's outer
        // try - NOT the post-delete re-locate, which has its own catch for every COM-failure
        // type. Either way the sentence "Nothing was changed" was unassertable.
        RecordingSession session = new RecordingSession { DiscardRefusal = "COMException 0x800706BE" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        service.DraftRegistry.Register(DraftId);
        DraftRefusedException refusal = Assert.Throws<DraftRefusedException>(() => service.DiscardDraft(DraftId));

        Assert.Equal(MailService.ComFailureRefusal, refusal.Reason);
        Assert.DoesNotContain("Nothing was changed", refusal.Message, StringComparison.Ordinal);
        Assert.Contains("UNKNOWN", refusal.Message, StringComparison.Ordinal);
        Assert.Contains("DELETED ITEMS", refusal.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(MutationOutcome.Unknown, OutlookTools.DraftRefusalOutcome(refusal.Reason));
    }

    // ---------------------------------------------------------------- row 5: the orphan draft

    [Fact]
    public void ACreateThatFailedAfterTheSave_LeavesADraftDiscardDraftCanReach()
    {
        // THE BEHAVIOURAL HALF. Save() commits the draft and four COM steps follow it; a
        // failure in any of them used to end with the caller never learning the id, so the
        // orphan was out of reach of the only cleanup tool the product has - discard_draft
        // refuses anything not in this registry.
        RecordingSession session = new RecordingSession
        {
            CreateRefusal = "COMException 0x80004005",
            SavedDraftEntryId = DraftId,
        };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(
            () => service.NewDraft("me@example.com", "them@example.com", null, "A subject", "body"));

        Assert.True(service.DraftRegistry.Contains(DraftId));
        Assert.Equal(MutationOutcome.Unknown, failure.Outcome);
        Assert.Contains(DraftId, failure.Message, StringComparison.Ordinal);
        Assert.Contains("ALREADY SAVED", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ACreateThatFailedWithNoIdToReport_StillTellsTheCallerToLookInDrafts()
    {
        // The COM layer could not read the id back, which is itself a COM failure - so the
        // weaker sentence is the honest one, and it must still send the caller looking.
        RecordingSession session = new RecordingSession { CreateRefusal = "COMException 0x80004005" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(
            () => service.NewDraft("me@example.com", "them@example.com", null, "A subject", "body"));

        Assert.Equal(MutationOutcome.Unknown, failure.Outcome);
        Assert.Contains("MAY HAVE BEEN SAVED", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnAccountThatDoesNotExist_StillProvesNothingWasCreated()
    {
        // Decided before anything is created, so this one keeps the strong claim. Removing it
        // everywhere would be the opposite failure.
        RecordingSession session = new RecordingSession { CreateRefusal = "AccountNotFound" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(
            () => service.NewDraft("nobody@example.com", "them@example.com", null, "A subject", "body"));

        Assert.Equal(MutationOutcome.Unchanged, failure.Outcome);
        Assert.Contains("Nothing was created", failure.Message, StringComparison.Ordinal);
    }

    // ---------------------------------------------------------------- row 6: folders made before a refusal

    [Fact]
    public void AMoveRefusedAfterCreatingFolders_NamesTheFoldersItMade()
    {
        // move_mail to 'Deleted Items/foo' with create_folder:true really does create foo
        // inside Deleted Items and then refuse, because the folder is resolved before the
        // guard that inspects it. The folder used to be created and never mentioned.
        RecordingSession session = new RecordingSession
        {
            MoveRefusal = "TargetIsDeletedItems",
            CreatedFolderPaths = new[] { "Deleted Items/foo" },
        };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        MoveMailOutcome outcome = service.MoveMail(new[] { ItemId }, "Deleted Items/foo", createFolder: true);

        MoveItemView item = Assert.Single(outcome.Items);
        Assert.False(item.Ok);
        Assert.Contains("Deleted Items/foo", item.Error!, StringComparison.Ordinal);
        Assert.Contains("CREATED before this failed", item.Error!, StringComparison.Ordinal);
        Assert.Equal(new[] { "Deleted Items/foo" }, outcome.CreatedFolders);
    }

    [Fact]
    public void AMoveThatFailedWithoutCreatingAnything_SaysNothingAboutFolders()
    {
        RecordingSession session = new RecordingSession { MoveRefusal = "TargetIsOutbox" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        MoveMailOutcome outcome = service.MoveMail(new[] { ItemId }, "Outbox", createFolder: true);

        MoveItemView item = Assert.Single(outcome.Items);
        Assert.DoesNotContain("CREATED before", item.Error!, StringComparison.Ordinal);
        Assert.Null(outcome.CreatedFolders);
    }

    // ---------------------------------------------------------------- rows 4 and 16: Ok is a boolean over three states

    [Fact]
    public void AMoveRefusal_IsReportedAsUnchanged()
    {
        RecordingSession session = new RecordingSession { MoveRefusal = "AlreadyInTargetFolder" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        MoveItemView item = Assert.Single(service.MoveMail(new[] { ItemId }, "Archive").Items);

        Assert.False(item.Ok);
        Assert.Equal(MutationOutcome.Unchanged, item.Outcome);
    }

    [Fact]
    public void AMoveWhoseFolderCheckCouldNotBeMade_IsRefused_AndProvablyMovedNothing()
    {
        // Decision A's fail-closed rule: the guard could not establish where Deleted Items or
        // the Outbox is, so it refused - before Move() - and the reply must say so rather than
        // send the caller looking for a move that cannot have happened.
        RecordingSession session = new RecordingSession { MoveRefusal = SpecialFolderGuards.TargetGuardUnreadable };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        MoveItemView item = Assert.Single(service.MoveMail(new[] { ItemId }, "Projects").Items);

        Assert.False(item.Ok);
        Assert.Equal(MutationOutcome.Unchanged, item.Outcome);
        Assert.Contains("could not be checked", item.Error!, StringComparison.Ordinal);
        Assert.DoesNotContain("UNKNOWN", item.Error!, StringComparison.Ordinal);
    }

    [Fact]
    public void ADraftWhoseDraftsFolderCouldNotBeFound_IsANamedRefusal_ThatClaimsNothingItDidNotCheck()
    {
        // The Drafts gate's own unreadable answer. Unnamed, it would fall into the catch-all and
        // tell the caller the draft MAY have been deleted; named as not_in_drafts_folder, it
        // would claim something the gate never established.
        RecordingSession session = new RecordingSession { DiscardRefusal = SpecialFolderGuards.DraftsFolderUnreadable };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        service.DraftRegistry.Register(DraftId);
        DraftRefusedException refusal = Assert.Throws<DraftRefusedException>(() => service.DiscardDraft(DraftId));

        Assert.Equal("drafts_folder_unreadable", refusal.Reason);
        Assert.DoesNotContain("UNKNOWN", refusal.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("does not live in a Drafts folder", refusal.Message, StringComparison.Ordinal);
        Assert.Equal(MutationOutcome.Unchanged, OutlookTools.DraftRefusalOutcome(refusal.Reason));
    }

    // ---------------------------------------------------------------- Q85: the destinations may create, and must report

    private const string MadeDrafts = "tier@vm.invalid/Drafts";

    private const string MadeDeletedItems = "tier@vm.invalid/Deleted Items";

    [Fact]
    public void ANewDraftThatCreatedTheDraftsFolder_SaysSo_AndOneThatDidNot_SaysNothing()
    {
        RecordingSession session = new RecordingSession { CreatedFolder = MadeDrafts };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        DraftOutcome made = service.NewDraft("me@example.com", "them@example.com", null, "A subject", "body", display: false);
        Assert.Equal(new[] { MadeDrafts }, made.CreatedFolders);

        // Backward compatible: nothing created, no field at all - not an empty list.
        session.CreatedFolder = null;
        DraftOutcome plain = service.NewDraft("me@example.com", "them@example.com", null, "A subject", "body", display: false);
        Assert.Null(plain.CreatedFolders);
    }

    [Fact]
    public void ANewDraftThatFailedAfterCreatingTheDraftsFolder_StillReportsIt()
    {
        // The folder is made before the draft is saved, so a failure after it leaves the folder
        // behind: the failure says so, in its field and in its sentence.
        RecordingSession session = new RecordingSession { CreatedFolder = MadeDrafts, CreateRefusal = "COMException 0x80004005" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(
            () => service.NewDraft("me@example.com", "them@example.com", null, "A subject", "body", display: false));

        Assert.Equal(new[] { MadeDrafts }, failure.CreatedFolders);
        Assert.Contains("CREATED before this failed", failure.Message, StringComparison.Ordinal);
        Assert.Contains(MadeDrafts, failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ANewDraftRefusedBeforeAnythingWasLookedUp_ReportsNoFolder()
    {
        RecordingSession session = new RecordingSession { CreateRefusal = "AccountNotFound" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(
            () => service.NewDraft("nobody@example.com", "them@example.com", null, "A subject", "body", display: false));

        Assert.Null(failure.CreatedFolders);
        Assert.DoesNotContain("CREATED", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AReplyOrForwardThatCreatedTheSourceStoresDraftsFolder_SaysSo()
    {
        RecordingSession session = new RecordingSession { CreatedFolder = MadeDrafts };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        DraftOutcome reply = service.ReplyDraft(ItemId, "body", display: false);
        DraftOutcome forward = service.ForwardDraft(ItemId, "body", "them@example.com", display: false);

        Assert.Equal(new[] { MadeDrafts }, reply.CreatedFolders);
        Assert.Equal(new[] { MadeDrafts }, forward.CreatedFolders);
    }

    [Fact]
    public void AReplyThatFailedAfterCreatingTheDraftsFolder_StillReportsIt()
    {
        RecordingSession session = new RecordingSession { CreatedFolder = MadeDrafts, CreateRefusal = "COMException 0x80004005" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(
            () => service.ReplyDraft(ItemId, "body", display: false));

        Assert.Equal(new[] { MadeDrafts }, failure.CreatedFolders);
        Assert.Contains("CREATED before this failed", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ADiscardThatCreatedDeletedItems_SaysSo_AndOneThatDidNot_SaysNothing()
    {
        RecordingSession session = new RecordingSession { CreatedFolder = MadeDeletedItems };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        service.DraftRegistry.Register(DraftId);
        DiscardDraftOutcome made = service.DiscardDraft(DraftId);
        Assert.Equal(new[] { MadeDeletedItems }, made.CreatedFolders);

        session.CreatedFolder = null;
        service.DraftRegistry.Register(DraftId);
        DiscardDraftOutcome plain = service.DiscardDraft(DraftId);
        Assert.Null(plain.CreatedFolders);
    }

    [Fact]
    public void ADiscardThatFailedAfterCreatingDeletedItems_StillReportsIt()
    {
        RecordingSession session = new RecordingSession { CreatedFolder = MadeDeletedItems, DiscardRefusal = "COMException 0x800706BE" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        service.DraftRegistry.Register(DraftId);
        DraftRefusedException refusal = Assert.Throws<DraftRefusedException>(() => service.DiscardDraft(DraftId));

        Assert.Equal(MailService.ComFailureRefusal, refusal.Reason);
        Assert.Equal(new[] { MadeDeletedItems }, refusal.CreatedFolders);
        Assert.Contains("CREATED before this failed", refusal.Message, StringComparison.Ordinal);
        Assert.Contains("WHETHER THE DRAFT WAS DELETED IS UNKNOWN", refusal.Message, StringComparison.Ordinal);
    }

    // ---------------------------------------------------------------- Q96: what a failed draft call may claim

    [Fact]
    public void ANewDraftWhoseDraftsLookupFailed_SaysNoDraftWasCreated_NotThatOneMayHaveBeenSaved()
    {
        // Q96 (iii). The Drafts lookup comes before Outlook is given anything to save, so "a
        // draft may have been saved - check Drafts before retrying" was a claim that cannot be
        // true there. What is known is said instead: no draft, and - the lookup's own re-check
        // having found the top level as it was - no folder either.
        RecordingSession session = new RecordingSession
        {
            CreateRefusal = ComErrorTokens.With(ComErrorTokens.DraftsFolderUnavailable, "COMException 0x80004005"),
        };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(
            () => service.NewDraft("me@example.com", "them@example.com", null, "A subject", "body", display: false));

        Assert.Equal(MutationOutcome.Unchanged, failure.Outcome);
        Assert.Contains("NO DRAFT WAS CREATED", failure.Message, StringComparison.Ordinal);
        Assert.Contains("COMException 0x80004005", failure.Message, StringComparison.Ordinal);
        Assert.Contains("me@example.com", failure.Message, StringComparison.Ordinal);
        Assert.Contains("No folder was created either", failure.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("MAY HAVE BEEN SAVED", failure.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("check Drafts", failure.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Null(failure.CreatedFolders);
    }

    [Fact]
    public void ANewDraftWhoseDraftsLookupMadeTheFolderThenFailed_ReportsIt_InTheFieldAndTheSentence()
    {
        // Q96 (ii): the re-check found the folder the failed call made, and it travels exactly as
        // a folder made on any other failure does - and "no folder was created" is not said.
        RecordingSession session = new RecordingSession
        {
            CreateRefusal = ComErrorTokens.With(ComErrorTokens.DraftsFolderUnavailable, "COMException 0x80004005"),
            CreatedFolder = MadeDrafts,
        };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(
            () => service.NewDraft("me@example.com", "them@example.com", null, "A subject", "body", display: false));

        Assert.Equal(MutationOutcome.Unchanged, failure.Outcome);
        Assert.Equal(new[] { MadeDrafts }, failure.CreatedFolders);
        Assert.Contains("NO DRAFT WAS CREATED", failure.Message, StringComparison.Ordinal);
        Assert.Contains("CREATED before this failed", failure.Message, StringComparison.Ordinal);
        Assert.Contains(MadeDrafts, failure.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("No folder was created", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ANewDraftWhoseDraftsLookupCouldNotBeReChecked_SaysItCouldNotTell_AndClaimsNoFolder()
    {
        // Never claim what cannot be proven, in either direction: not "created", and not
        // "nothing was created" either.
        RecordingSession session = new RecordingSession
        {
            CreateRefusal = ComErrorTokens.With(ComErrorTokens.DraftsFolderCreationUnverified, "COMException 0x800706BA"),
        };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(
            () => service.NewDraft("me@example.com", "them@example.com", null, "A subject", "body", display: false));

        Assert.Equal(MutationOutcome.Unchanged, failure.Outcome);
        Assert.Contains("NO DRAFT WAS CREATED", failure.Message, StringComparison.Ordinal);
        Assert.Contains("could NOT be checked", failure.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("No folder was created", failure.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("CREATED before this failed", failure.Message, StringComparison.Ordinal);
        Assert.Null(failure.CreatedFolders);
    }

    [Theory]
    [InlineData("new")]
    [InlineData("reply")]
    [InlineData("forward")]
    public void ADraftThatFailedBeforeTheCompose_SaysNoDraftWasSaved(string tool)
    {
        // Q96 (iii), "check the other draft tools for the same wording": Reply()/Forward()
        // failing, or the account pin, came before anything that can save a draft, and the
        // reply and forward tools said "a draft may have been saved" over it too.
        RecordingSession session = new RecordingSession
        {
            CreateRefusal = ComErrorTokens.With(ComErrorTokens.DraftNotStarted, "COMException 0x80004005"),
        };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(() =>
        {
            _ = tool switch
            {
                "new" => service.NewDraft("me@example.com", "them@example.com", null, "A subject", "body", display: false),
                "reply" => service.ReplyDraft(ItemId, "body", display: false),
                _ => service.ForwardDraft(ItemId, "body", "them@example.com", display: false),
            };
        });

        Assert.Equal(MutationOutcome.Unchanged, failure.Outcome);
        Assert.Contains("NO DRAFT WAS SAVED", failure.Message, StringComparison.Ordinal);
        Assert.Contains("COMException 0x80004005", failure.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("MAY HAVE BEEN SAVED", failure.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("ALREADY SAVED", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ANewDraftThatFailedBeforeTheCompose_AfterItsLookupMadeTheFolder_StillNamesTheFolder()
    {
        // No draft, and still a folder: the lookup succeeded, made Drafts, and the next step
        // failed. "No draft was saved" must not swallow the residue sentence.
        RecordingSession session = new RecordingSession
        {
            CreateRefusal = ComErrorTokens.With(ComErrorTokens.DraftNotStarted, "COMException 0x80004005"),
            CreatedFolder = MadeDrafts,
        };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(
            () => service.NewDraft("me@example.com", "them@example.com", null, "A subject", "body", display: false));

        Assert.Equal(new[] { MadeDrafts }, failure.CreatedFolders);
        Assert.Contains("NO DRAFT WAS SAVED", failure.Message, StringComparison.Ordinal);
        Assert.Contains("CREATED before this failed", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AFailedLookupThatMadeTwoFolders_ReportsBoth()
    {
        // With no folder handed back, the re-check reports every folder that appeared; the
        // contract carries a list for exactly this, and the failure keeps both.
        RecordingSession session = new RecordingSession
        {
            CreateRefusal = ComErrorTokens.With(ComErrorTokens.DraftsFolderUnavailable, "COMException 0x80004005"),
            CreatedFolder = MadeDrafts,
            AlsoCreatedFolder = "tier@vm.invalid/Junk Email",
        };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(
            () => service.NewDraft("me@example.com", "them@example.com", null, "A subject", "body", display: false));

        Assert.Equal(new[] { MadeDrafts, "tier@vm.invalid/Junk Email" }, failure.CreatedFolders);
        Assert.Contains("2 folders were CREATED before this failed", failure.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(ComErrorTokens.DraftNotStarted)]
    [InlineData(ComErrorTokens.DraftsFolderUnavailable)]
    [InlineData(ComErrorTokens.DraftsFolderCreationUnverified)]
    public void TheDraftTokens_RoundTrip_AndNeverMatchEachOtherOrTheRetryToken(string token)
    {
        string written = ComErrorTokens.With(token, "COMException 0x80004005");

        Assert.True(ComErrorTokens.TryRead(written, token, out string detail));
        Assert.Equal("COMException 0x80004005", detail);
        Assert.True(ComErrorTokens.TryRead(token, token, out string bare));
        Assert.Equal("unknown", bare);
        foreach (string other in new[] { ComErrorTokens.DraftNotStarted, ComErrorTokens.DraftsFolderUnavailable, ComErrorTokens.DraftsFolderCreationUnverified, ComErrorTokens.ItemNotFound })
        {
            if (other != token)
            {
                Assert.False(ComErrorTokens.TryRead(written, other, out _), token + " was read as " + other);
            }
        }

        // The cross-store retry keys on ItemNotFound by exact match; a draft token must never
        // look like "the item was not opened", or a failure after a draft was made would fan out.
        Assert.NotEqual(ComErrorTokens.ItemNotFound, written);
        Assert.False(ComErrorTokens.TryRead(token + "X:detail", token, out _));
        Assert.False(ComErrorTokens.TryRead(null, token, out _));
        Assert.False(ComErrorTokens.TryRead("COMException 0x80004005", token, out _));
    }

    [Fact]
    public void TheErrorPayload_CarriesCreatedFolders_OnlyWhenAFolderWasMade()
    {
        // The failure half of the field, on the wire. Absent means absent - the error shape
        // every other failure has always had is untouched.
        string with = ErrorJson(OutlookTools.Error("OperationFailed", "m", null, outcome: MutationOutcome.Unknown, createdFolders: new[] { MadeDrafts }));
        string without = ErrorJson(OutlookTools.Error("OperationFailed", "m", null, outcome: MutationOutcome.Unknown));

        using (JsonDocument parsed = JsonDocument.Parse(with))
        {
            JsonElement created = parsed.RootElement.GetProperty("error").GetProperty("createdFolders");
            Assert.Equal(MadeDrafts, Assert.Single(created.EnumerateArray()).GetString());
        }

        using (JsonDocument parsed = JsonDocument.Parse(without))
        {
            Assert.False(parsed.RootElement.GetProperty("error").TryGetProperty("createdFolders", out _));
        }
    }

    private static string ErrorJson(ModelContextProtocol.Protocol.CallToolResult result)
    {
        Assert.True(result.IsError);
        return Assert.IsType<ModelContextProtocol.Protocol.TextContentBlock>(Assert.Single(result.Content)).Text;
    }

    [Fact]
    public void AnUnclassifiedMoveFailure_IsReportedAsUnknown_AndStopsSayingRetry()
    {
        // Row 16. Every named code is decided before Move() runs; the catch-all is not, so
        // "check outlook_health and retry" was advice to repeat a move that may already have
        // happened - against an EntryID a successful move has already invalidated.
        RecordingSession session = new RecordingSession { MoveRefusal = "COMException 0x800706BE" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        MoveItemView item = Assert.Single(service.MoveMail(new[] { ItemId }, "Archive").Items);

        Assert.False(item.Ok);
        Assert.Equal(MutationOutcome.Unknown, item.Outcome);
        Assert.Contains("UNKNOWN", item.Error!, StringComparison.Ordinal);
    }

    [Fact]
    public void AMoveThatHappenedAndCouldNotBeAudited_IsReportedAsApplied()
    {
        // The case no wording could repair: Ok=false over an item that MOVED. The message has
        // always said so and it travelled through a field documented as "nothing was moved
        // for this item".
        Assert.Equal(MutationOutcome.Unchanged, MailService.MoveFailureOutcome("TargetFolderCreateFailed"));
        Assert.Equal(MutationOutcome.Unchanged, MailService.MoveFailureOutcome("CrossStoreTarget:Work"));
        Assert.Equal(MutationOutcome.Unknown, MailService.MoveFailureOutcome(null));
        Assert.Equal(MutationOutcome.Unknown, MailService.MoveFailureOutcome("COMException 0x80004005"));
    }

    // ---------------------------------------------------------------- rows 11 and 14

    [Fact]
    public void AFailedAttachmentSave_NamesThePathAPartialFileWouldBeAt()
    {
        RecordingSession session = new RecordingSession
        {
            AttachmentRefusal = "IOException: IOException",
            AttemptedPath = @"C:\scratch\offer.pdf",
        };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(
            () => service.SaveAttachment(ItemId, 1, @"C:\scratch"));

        Assert.Equal(MutationOutcome.Unknown, failure.Outcome);
        Assert.Contains(@"C:\scratch\offer.pdf", failure.Message, StringComparison.Ordinal);
        Assert.Contains("PARTIAL OR COMPLETE FILE MAY EXIST", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnAttachmentIndexOutOfRange_StillProvesNothingWasWritten()
    {
        RecordingSession session = new RecordingSession { AttachmentRefusal = "AttachmentIndexOutOfRange (count=0)" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(
            () => service.SaveAttachment(ItemId, 1, @"C:\scratch"));

        Assert.Equal(MutationOutcome.Unchanged, failure.Outcome);
        Assert.Contains("Nothing was written to disk", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AFailedDisplay_AdmitsTheWindowMayBeOpenAndTheMailMarkedRead()
    {
        // Display() is the LAST call in the sequence, so a failure reported over it is
        // precisely the case where it may already have happened.
        RecordingSession session = new RecordingSession { DisplayRefusal = "COMException 0x800706BE" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(() => service.OpenInOutlook(ItemId));

        Assert.Equal(MutationOutcome.Unknown, failure.Outcome);
        Assert.Contains("MARKED READ", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnItemThatCouldNotBeOpenedAtAll_StillProvesNothingWasDisplayed()
    {
        RecordingSession session = new RecordingSession { DisplayRefusal = ComErrorTokens.ItemNotFound };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(() => service.OpenInOutlook(ItemId));

        Assert.Equal(MutationOutcome.Unchanged, failure.Outcome);
        Assert.Contains("nothing was displayed", failure.Message, StringComparison.Ordinal);
    }

    // ---------------------------------------------------------------- the function the audit left unread

    [Fact]
    public void AFailedNavigation_AdmitsTheWindowMayHaveMovedAnyway()
    {
        // BuildNavigationError was the one function section 5 of the audit deprioritised and
        // never read. Read: its catch-all sits after an Explorer may have been CREATED and
        // shown, after CurrentFolder was set, and for show_search_results after Search() was
        // issued - so "Outlook could not show the requested view" was the same defect.
        RecordingSession session = new RecordingSession { NavigationRefusal = "COMException 0x80004005" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(() => service.GotoFolder("Work", "Inbox"));

        Assert.Equal(MutationOutcome.Unknown, failure.Outcome);
        Assert.Contains("MAY HAVE MOVED ANYWAY", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AStoreThatDoesNotExist_StillProvesTheWindowDidNotMove()
    {
        RecordingSession session = new RecordingSession { NavigationRefusal = "StoreNotFound" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(() => service.GotoFolder("Ghost"));

        Assert.Equal(MutationOutcome.Unchanged, failure.Outcome);
        Assert.Contains("nothing was opened or moved on screen", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void APerItemMoveTimeout_ReportsUnknownRatherThanNotMoved()
    {
        // Row 26's message has said the outcome is UNKNOWN since the batch-budget work; it
        // travelled through a field documented as "nothing was moved for this item". The
        // message and the field now agree.
        RecordingSession session = new RecordingSession { MoveTimeout = true };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        MoveItemView item = Assert.Single(service.MoveMail(new[] { ItemId }, "Archive").Items);

        Assert.False(item.Ok);
        Assert.Equal(MutationOutcome.Unknown, item.Outcome);
        Assert.Contains("UNKNOWN", item.Error!, StringComparison.Ordinal);
    }

    [Fact]
    public void APerItemArchiveTimeout_ReportsUnknownToo()
    {
        // archive_mail keeps its own copy of the per-item timeout arm, on its own COM call,
        // and a test that drove only move_mail left it unguarded. Proved by mutation rather
        // than assumed: reverting archive's outcome to "unchanged" passed the whole suite
        // while reverting move's did not.
        RecordingSession session = new RecordingSession { ArchiveTimeout = true };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        MoveItemView item = Assert.Single(service.ArchiveMail(new[] { ItemId }).Items);

        Assert.False(item.Ok);
        Assert.Equal(MutationOutcome.Unknown, item.Outcome);
        Assert.Contains("UNKNOWN", item.Error!, StringComparison.Ordinal);
    }

    [Fact]
    public void ADerivedDraftWhoseSourceWasNeverOpened_ProvesNothingWasCreated()
    {
        // ItemNotFound is set at GetItemFromID and nowhere else, so it proves the negative - and
        // since Q96 (iii) it is not the only one: DraftNotStarted marks every failure before the
        // compose. Everything after that may have saved a draft by the time it fails.
        RecordingSession session = new RecordingSession { CreateRefusal = ComErrorTokens.ItemNotFound };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(
            () => service.ReplyDraft(ItemId, "body"));

        Assert.Equal(MutationOutcome.Unchanged, failure.Outcome);
        Assert.Contains("no draft was created", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ADerivedDraftThatFailedAfterTheSave_SaysADraftMayExist()
    {
        RecordingSession session = new RecordingSession
        {
            CreateRefusal = "COMException 0x80004005",
            SavedDraftEntryId = DraftId,
        };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        OperationOutcomeException failure = Assert.Throws<OperationOutcomeException>(
            () => service.ReplyDraft(ItemId, "body"));

        Assert.Equal(MutationOutcome.Unknown, failure.Outcome);
        Assert.True(service.DraftRegistry.Contains(DraftId));
        Assert.Contains("ALREADY SAVED", failure.Message, StringComparison.Ordinal);
    }

    // ---------------------------------------------------------------- rows 9 and 10: the send path's two wrinkles

    [Fact]
    public void AnAbandonedSend_StopsNamingAStepItCannotKnowFailed()
    {
        // Row 9. "The draft could not be re-opened for sending" named ONE step out of the
        // whole sequence and named the wrong one: this is the catch-all, so the failure can
        // be anywhere between the open and the moment before Send(). "Nothing was sent"
        // survives the audit and is why the outcome is unchanged - Send() has its own catch,
        // which is what makes the negative provable rather than hoped for.
        OperationOutcomeException failure = SendAndExpectFailure("COMException 0x80004005");

        Assert.Equal(MutationOutcome.Unchanged, failure.Outcome);
        Assert.DoesNotContain("could not be re-opened", failure.Message, StringComparison.Ordinal);
        Assert.Contains("Nothing was sent", failure.Message, StringComparison.Ordinal);
        Assert.Contains("send-account pin", failure.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AnAbortedSend_SaysTheDraftsSendAccountWasRewritten()
    {
        // Row 10. "Nothing was sent" was true and was never the misleading half. What was
        // missing is that the abort is not a no-op on the DRAFT: SendUsingAccount is written
        // to the item BEFORE the readback that failed, and no path restores it.
        SendRefusedException refusal = Assert.Throws<SendRefusedException>(
            () => SendWithToken("SendIdentityVerificationFailed"));

        Assert.Contains("Nothing was sent", refusal.Message, StringComparison.Ordinal);
        Assert.Contains("WROTE the draft's send-account pin", refusal.Message, StringComparison.Ordinal);
        Assert.Contains("did not restore it", refusal.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void AFailedSendCall_StillPointsAtTheOutbox()
    {
        // Unchanged by this pass and pinned so it stays that way: Outlook answering with an
        // error from Send() is the one send failure where the mail may already be queued.
        OperationOutcomeException failure = SendAndExpectFailure("SendCallFailed:COMException 0x80004005");

        Assert.Equal(MutationOutcome.Unknown, failure.Outcome);
        Assert.Contains("Outbox", failure.Message, StringComparison.Ordinal);
    }

    private static OperationOutcomeException SendAndExpectFailure(string sendError)
    {
        return Assert.Throws<OperationOutcomeException>(() => SendWithToken(sendError));
    }

    /// <summary>
    /// Drives the real two-step send flow against a stand-in: the first call issues the
    /// token, the second consumes it and reaches the failure the test is about.
    /// </summary>
    private static void SendWithToken(string sendError)
    {
        RecordingSession session = new RecordingSession { SendRefusal = sendError };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        SendOutcome issued = service.Send(ItemId);
        Assert.NotNull(issued.ConfirmToken);
        _ = service.Send(ItemId, issued.ConfirmToken);
    }

    // ---------------------------------------------------------------- row 25: the send path, confirmed

    [Fact]
    public void AKilledSend_StillSaysTheOutcomeIsUnknownAndNamesTheOutbox()
    {
        // The send path comes out of the audit CONFIRMED rather than fixed - Send() is the
        // last call in the sequence and has its own catch, which is what makes "Nothing was
        // sent" provable elsewhere on that path rather than hoped for. Pinned here so the
        // sweep that changed its neighbours leaves a record that this one was checked.
        string message = MailService.DescribeSendOutcomeUnknown("Outlook did not respond");

        Assert.Contains("UNKNOWN", message, StringComparison.Ordinal);
        Assert.Contains("OUTBOX", message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("Nothing was sent", message, StringComparison.Ordinal);
    }

    // ---------------------------------------------------------------- the descriptions that teach it

    /// <summary>
    /// The thirteen tools that can change something. Named rather than derived, deliberately:
    /// deriving the list from the same classification the code uses would make this test
    /// agree with any mistake that classification makes.
    /// </summary>
    private static readonly string[] MutatingTools =
    {
        "save_attachment", "move_mail", "archive_mail", "open_in_outlook", "goto_folder",
        "show_search_results", "new_draft", "reply_draft", "replyall_draft", "forward_draft",
        "update_draft", "discard_draft", "send",
    };

    [Fact]
    public void EveryToolThatCanChangeSomething_TeachesTheOutcomeField()
    {
        // A field nothing reads is cost with no benefit, and the only thing that makes an
        // agent read it is the description. Without this, removing the clause is invisible:
        // DescriptionBudgetCiTests measures sizes and asserts nothing about content.
        List<string> missing = new List<string>();
        foreach (string tool in MutatingTools)
        {
            if (!DescriptionOf(tool).Contains("outcome:", StringComparison.Ordinal))
            {
                missing.Add(tool);
            }
        }

        Assert.True(missing.Count == 0, "these tools no longer teach the outcome field: " + string.Join(", ", missing));
    }

    [Fact]
    public void TheReadOnlyTools_DoNotCarryTheClause()
    {
        // The other half: they can only ever answer "unchanged", and search is the one
        // description already close to the client's measured 2048-code-unit cut.
        Assert.DoesNotContain("outcome:", DescriptionOf("search"), StringComparison.Ordinal);
        Assert.DoesNotContain("outcome:", DescriptionOf("read"), StringComparison.Ordinal);
    }

    [Fact]
    public void DiscardDraft_StopsPromisingThatAFailureIsAlwaysANoOp()
    {
        // "it never silently does nothing" is true of REFUSALS and was silently untrue of
        // failures: a COM failure during the delete leaves the outcome unknown.
        string description = DescriptionOf("discard_draft");

        Assert.Contains("never silently does nothing", description, StringComparison.Ordinal);
        Assert.Contains("UNKNOWN", description, StringComparison.Ordinal);
        Assert.Contains("Deleted Items rather than assuming nothing happened", description, StringComparison.Ordinal);
    }

    private static string DescriptionOf(string toolName)
    {
        foreach (MethodInfo method in typeof(OutlookTools).GetMethods(BindingFlags.Public | BindingFlags.Static))
        {
            foreach (object attribute in method.GetCustomAttributes(inherit: false))
            {
                if (string.Equals(attribute.GetType().Name, "McpServerToolAttribute", StringComparison.Ordinal)
                    && string.Equals(
                        (string?)attribute.GetType().GetProperty("Name")?.GetValue(attribute), toolName, StringComparison.Ordinal))
                {
                    return method
                        .GetCustomAttributes(typeof(System.ComponentModel.DescriptionAttribute), inherit: false)
                        .Cast<System.ComponentModel.DescriptionAttribute>()
                        .Single()
                        .Description;
                }
            }
        }

        throw new InvalidOperationException("No tool named '" + toolName + "' - this test can no longer prove anything.");
    }

    // ---------------------------------------------------------------- support

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
    /// A session that refuses whichever operation the test is about, in exactly the shape the
    /// COM layer refuses it: a null return plus a content-free token in the <c>error</c> out
    /// parameter, with the other out parameters filled in the way the real one fills them.
    /// </summary>
    private sealed class RecordingSession
    {
        internal RecordingSession()
        {
            AsSession = Proxy.Create(this);
        }

        internal IOutlookSession AsSession { get; }

        internal string? CreateRefusal { get; set; }

        internal string? SavedDraftEntryId { get; set; }

        internal string? DiscardRefusal { get; set; }

        internal string? MoveRefusal { get; set; }

        internal IReadOnlyList<string>? CreatedFolderPaths { get; set; }

        internal string? AttachmentRefusal { get; set; }

        internal string? AttemptedPath { get; set; }

        internal string? DisplayRefusal { get; set; }

        internal string? NavigationRefusal { get; set; }

        internal bool MoveTimeout { get; set; }

        internal bool ArchiveTimeout { get; set; }

        internal string? SendRefusal { get; set; }

        /// <summary>The folder the draft/discard call reports CREATING (Q85), in the COM layer's <c>store/path</c> form.</summary>
        internal string? CreatedFolder { get; set; }

        /// <summary>A second folder the same call reports - what a failed lookup's re-check can find (Q96 (ii)).</summary>
        internal string? AlsoCreatedFolder { get; set; }

        /// <summary>The <c>createdFolders</c> out parameter as the COM layer fills it: null when nothing was made.</summary>
        private IReadOnlyList<string>? CreatedFolders()
        {
            List<string> made = new[] { CreatedFolder, AlsoCreatedFolder }.Where(f => f != null).Select(f => f!).ToList();
            return made.Count > 0 ? made : null;
        }

        private static ComDraftInfo Snapshot()
        {
            return new ComDraftInfo(
                DraftId, "Work", "store-work", "Drafts", "folder-drafts", "A subject",
                "someone@example.com", null, "conv-1", Array.Empty<ComRecipientInfo>(), "A subject");
        }

        private object? Handle(MethodInfo method, object?[]? args)
        {
            switch (method.Name)
            {
                case nameof(IOutlookSession.GetStoreDetails):
                    return Array.Empty<ComStoreDetail>();

                case nameof(IOutlookSession.TryCreateNewDraft):
                case nameof(IOutlookSession.TryCreateDerivedDraft):
                    SetOut(method, args, "savedDraftEntryId", SavedDraftEntryId);
                    SetOut(method, args, "createdFolders", CreatedFolders());
                    if (CreateRefusal != null)
                    {
                        SetOut(method, args, "error", CreateRefusal);
                        return null;
                    }

                    return new ComDraftCreateResult(Snapshot(), true, false, 0, 0, false, null, false);

                case nameof(IOutlookSession.TryDiscardDraft):
                    SetOut(method, args, "createdFolders", CreatedFolders());
                    if (DiscardRefusal != null)
                    {
                        SetOut(method, args, "error", DiscardRefusal);
                        return null;
                    }

                    return new ComDraftDiscardResult(DraftId, DraftId, "Work", "Drafts", "Deleted Items", "A subject");

                case nameof(IOutlookSession.TryMoveItemToPath):
                    SetOut(method, args, "createdFolderPaths", CreatedFolderPaths);
                    if (MoveTimeout)
                    {
                        // What a killed COM host looks like from the parent: no answer, and
                        // no report of what was applied.
                        throw new TimeoutException("Outlook did not respond to 'TryMoveItemToPath' within 240000 ms.");
                    }

                    if (MoveRefusal != null)
                    {
                        SetOut(method, args, "error", MoveRefusal);
                        return null;
                    }

                    return new ComMoveItemResult(ItemId, ItemId, "Work", "Inbox", "Archive", Array.Empty<string>());

                case nameof(IOutlookSession.TrySaveAttachment):
                    // A VALUE-TYPE out parameter has to be filled explicitly: DispatchProxy
                    // hands the args array to Invoke with null in every out slot, so leaving
                    // sizeBytes alone unboxes null into a long and the proxy throws.
                    SetOut(method, args, "sizeBytes", 0L);
                    SetOut(method, args, "attemptedPath", AttemptedPath);
                    SetOut(method, args, "error", AttachmentRefusal);
                    return null;

                case nameof(IOutlookSession.TryGetMailInfo):
                    return Snapshot();

                case nameof(IOutlookSession.TryResolveArchiveFolder):
                case nameof(IOutlookSession.TryResolveOrCreateArchiveFolder):
                    return new ComArchiveFolderInfo("Work", "store-work", "folder-archive", "Archive", "Archive", "designated");

                case nameof(IOutlookSession.TryMoveItemToFolderId):
                    if (ArchiveTimeout)
                    {
                        throw new TimeoutException("Outlook did not respond to 'TryMoveItemToFolderId' within 240000 ms.");
                    }

                    return new ComMoveItemResult(ItemId, ItemId, "Work", "Inbox", "Archive", Array.Empty<string>());

                case nameof(IOutlookSession.TryGetSendableDraftState):
                    return new ComSendableDraftState(
                        ItemId, "store-work", "Work", "Drafts", "A subject", false, "body",
                        "me@example.com", Array.Empty<ComRecipientInfo>());

                case nameof(IOutlookSession.TrySendDraft):
                    SetOut(method, args, "error", SendRefusal);
                    return null;

                case nameof(IOutlookSession.TryDisplayItem):
                    SetOut(method, args, "error", DisplayRefusal);
                    return null;

                case nameof(IOutlookSession.TryGotoFolder):
                case nameof(IOutlookSession.TryShowSearchResults):
                    SetOut(method, args, "error", NavigationRefusal);
                    return null;

                default:
                    return null;
            }
        }

        private static void SetOut(MethodInfo method, object?[]? args, string name, object? value)
        {
            ParameterInfo[] parameters = method.GetParameters();
            for (int i = 0; args != null && i < parameters.Length && i < args.Length; i++)
            {
                if (parameters[i].IsOut && string.Equals(parameters[i].Name, name, StringComparison.Ordinal))
                {
                    args[i] = value;
                }
            }
        }

        internal class Proxy : DispatchProxy
        {
            private RecordingSession _owner = null!;

            internal static IOutlookSession Create(RecordingSession owner)
            {
                object proxy = Create<IOutlookSession, Proxy>()!;
                ((Proxy)proxy)._owner = owner;
                return (IOutlookSession)proxy;
            }

            protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
            {
                return _owner.Handle(targetMethod!, args);
            }
        }
    }
}
