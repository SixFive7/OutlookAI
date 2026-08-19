using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The drafting surface's one silent cap: the addresses Outlook could NOT resolve were cut
/// at <see cref="MailService.UnresolvedRecipientsCap"/> with no truncation flag and no
/// total, while the RESOLVED recipient list two properties above carried exactly that pair.
/// <para>
/// WHY THIS MATTERED MORE THAN ITS "tangential" LABEL SUGGESTED. An agent told "these 20
/// addresses did not resolve" out of 27 has a false-complete list, on the highest-stakes
/// surface in the product: the remedy for an unresolved recipient is to ask the USER about
/// it, so the seven past the cap were never mentioned to anybody, and the draft went out - or
/// failed to - carrying a fault its own report had declared absent. The resolved list can
/// survive being short (the draft holds the recipients either way, and every operation reads
/// the full COM-side list); this one cannot.
/// </para>
/// <para>
/// Two tiers here, deliberately. The pure classifiers pin the DECISION; the two private
/// mapping methods are invoked by reflection to pin that the decision actually reaches the
/// payload - that half is what a classifier test alone cannot prove, and it was exactly the
/// half that was missing. No Outlook, no mailbox, no audit line: the mappers are pure
/// functions of a snapshot plus a gateway they only ever use for a best-effort attachment
/// re-read that is allowed to fail.
/// </para>
/// </summary>
public sealed class UnresolvedRecipientReportingTests
{
    private const string Store = "alice@example.com";

    // ============================================================ the pure classifier

    [Fact]
    public void UnderTheCap_EverythingIsListed_AndNothingClaimsTruncation()
    {
        IReadOnlyList<string>? capped = MailService.CapUnresolvedRecipients(
            Addresses(3), out int total, out bool truncated);

        Assert.Equal(3, capped!.Count);
        Assert.Equal(3, total);
        Assert.False(truncated);
        Assert.Null(MailService.DescribeUnresolvedRecipientCap(total, truncated));
    }

    [Fact]
    public void ExactlyTheCap_IsNotTruncated()
    {
        // The boundary the has-more pair exists to get right: 20 of 20 is a complete list.
        IReadOnlyList<string>? capped = MailService.CapUnresolvedRecipients(
            Addresses(MailService.UnresolvedRecipientsCap), out int total, out bool truncated);

        Assert.Equal(MailService.UnresolvedRecipientsCap, capped!.Count);
        Assert.Equal(MailService.UnresolvedRecipientsCap, total);
        Assert.False(truncated);
    }

    [Fact]
    public void OverTheCap_CutsTheListAndReportsTheRealTotal()
    {
        IReadOnlyList<string>? capped = MailService.CapUnresolvedRecipients(
            Addresses(27), out int total, out bool truncated);

        Assert.Equal(MailService.UnresolvedRecipientsCap, capped!.Count);
        Assert.Equal(27, total);
        Assert.True(truncated);

        // Order preserved: the FIRST cap-many are the ones named, so the list is a prefix of
        // what the caller passed rather than an arbitrary slice.
        Assert.Equal("bad1@example.com", capped[0]);
        Assert.Equal($"bad{MailService.UnresolvedRecipientsCap}@example.com", capped[^1]);
    }

    [Theory]
    [InlineData(null)]
    [InlineData(0)]
    public void NothingFailedToResolve_IsAnAbsentField_NotAnEmptyArray(int? count)
    {
        IReadOnlyList<string>? source = count == null ? null : Array.Empty<string>();

        Assert.Null(MailService.CapUnresolvedRecipients(source, out int total, out bool truncated));
        Assert.Equal(0, total);
        Assert.False(truncated);
    }

    // ================================================================= the prose half

    [Fact]
    public void TheSentence_NamesTheTotal_TheCap_AndWhatIsNotListed()
    {
        string advice = MailService.DescribeUnresolvedRecipientCap(27, truncated: true)!;

        Assert.Contains("27", advice, StringComparison.Ordinal);
        Assert.Contains(
            MailService.UnresolvedRecipientsCap.ToString(System.Globalization.CultureInfo.InvariantCulture),
            advice,
            StringComparison.Ordinal);

        // The remainder, spelled out: an agent relaying "20 failed" out of 27 is the defect.
        Assert.Contains("7", advice, StringComparison.Ordinal);
        Assert.Contains("NOT listed", advice, StringComparison.Ordinal);
    }

    [Fact]
    public void AnUncutList_GetsNoSentence()
    {
        // The cry-wolf guard: a draft whose three addresses all failed is fully reported by
        // the list itself, and a sentence there would train the reader to ignore this one.
        Assert.Null(MailService.DescribeUnresolvedRecipientCap(3, truncated: false));
    }

    // ============================================================== reaching the payload

    [Fact]
    public void ANewDraft_CarriesTheFlagTheTotalAndTheSentence()
    {
        DraftOutcome outcome = MapNewDraft(27);

        Assert.Equal(MailService.UnresolvedRecipientsCap, outcome.UnresolvedRecipients!.Count);
        Assert.True(outcome.UnresolvedRecipientsTruncated);
        Assert.Equal(27, outcome.UnresolvedRecipientsTotal);
        Assert.NotNull(outcome.UnresolvedRecipientsAdvice);
        Assert.Contains("27", outcome.UnresolvedRecipientsAdvice!, StringComparison.Ordinal);
    }

    [Fact]
    public void ANewDraftUnderTheCap_LooksExactlyAsItAlwaysDid()
    {
        // The regression half. Nothing new may appear on the ordinary draft, or every draft
        // payload grows two fields and a sentence that say nothing.
        DraftOutcome outcome = MapNewDraft(2);

        Assert.Equal(2, outcome.UnresolvedRecipients!.Count);
        Assert.Null(outcome.UnresolvedRecipientsTruncated);
        Assert.Null(outcome.UnresolvedRecipientsTotal);
        Assert.Null(outcome.UnresolvedRecipientsAdvice);
    }

    [Fact]
    public void ADraftWhereEveryAddressResolved_ReportsNothingAtAll()
    {
        DraftOutcome outcome = MapNewDraft(0);

        Assert.Null(outcome.UnresolvedRecipients);
        Assert.Null(outcome.UnresolvedRecipientsTruncated);
        Assert.Null(outcome.UnresolvedRecipientsTotal);
        Assert.Null(outcome.UnresolvedRecipientsAdvice);
    }

    [Fact]
    public void AnUpdatedDraft_ReportsTheCutTheSameWay()
    {
        // update_draft REPLACES the recipient list, so it is the call most likely to hand
        // Outlook a long address list in one go - it must not report the cut differently
        // from the creators.
        UpdateDraftOutcome outcome = MapUpdate(27);

        Assert.Equal(MailService.UnresolvedRecipientsCap, outcome.UnresolvedRecipients!.Count);
        Assert.True(outcome.UnresolvedRecipientsTruncated);
        Assert.Equal(27, outcome.UnresolvedRecipientsTotal);
        Assert.Equal(
            MailService.DescribeUnresolvedRecipientCap(27, truncated: true),
            outcome.UnresolvedRecipientsAdvice);
    }

    [Fact]
    public void AnUpdatedDraftUnderTheCap_ReportsNoCut()
    {
        UpdateDraftOutcome outcome = MapUpdate(5);

        Assert.Equal(5, outcome.UnresolvedRecipients!.Count);
        Assert.Null(outcome.UnresolvedRecipientsTruncated);
        Assert.Null(outcome.UnresolvedRecipientsTotal);
        Assert.Null(outcome.UnresolvedRecipientsAdvice);
    }

    // =================================================================== fixtures

    private static IReadOnlyList<string> Addresses(int count)
    {
        return Enumerable.Range(1, count).Select(i => $"bad{i}@example.com").ToList();
    }

    private static ComDraftInfo Draft()
    {
        return new ComDraftInfo(
            entryId: "0000DRAFT",
            storeDisplayName: Store,
            storeId: "store-alice",
            parentFolderName: "Drafts",
            parentFolderEntryId: "0000DRAFTS",
            subject: "a draft",
            sendUsingAccountSmtp: Store,
            conversationIndex: null,
            conversationId: null,
            recipients: new[] { new ComRecipientInfo("to", "Bob", "bob@example.com") });
    }

    /// <summary>
    /// Drives the real <c>MailService.ToDraftOutcome</c>. Private, so reflection - the
    /// alternative is <c>NewDraft</c>, which writes a real audit line under
    /// %LOCALAPPDATA% and needs a whole compose path stubbed to prove one mapping.
    /// </summary>
    private static DraftOutcome MapNewDraft(int unresolvedCount)
    {
        ComDraftCreateResult created = new ComDraftCreateResult(
            Draft(),
            accountResolved: true,
            signatureInjected: false,
            bodyTextCharsBeforeSignature: 0,
            bodyTextCharsAfterSignature: 0,
            movedToDrafts: false,
            initialSaveFolderName: "Drafts",
            displayed: false,
            unresolvedRecipients: Addresses(unresolvedCount));

        return (DraftOutcome)Invoke(
            "ToDraftOutcome",
            "new",
            created,
            null,
            null,
            ComDraftBody.FromText("body"),
            Array.Empty<string>(),
            Array.Empty<DraftAttachmentFile>());
    }

    private static UpdateDraftOutcome MapUpdate(int unresolvedCount)
    {
        ComDraftUpdateResult updated = new ComDraftUpdateResult(
            Draft(),
            changedFields: new[] { "to" },
            unresolvedRecipients: Addresses(unresolvedCount),
            attachments: Array.Empty<ComAttachmentInfo>(),
            attachmentsAdded: Array.Empty<string>(),
            attachmentsRemoved: Array.Empty<string>(),
            attachmentsFailed: Array.Empty<string>(),
            bodyReplaced: false,
            bodyPlacedViaWordEditor: false,
            displayed: false,
            signatureOverrideName: null,
            signatureOverrideApplied: false,
            signatureOverrideError: null,
            conversationTopicPreserved: null);

        return (UpdateDraftOutcome)Invoke(
            "ToUpdateOutcome", updated, null, null, Array.Empty<string>(), Array.Empty<string>(), false);
    }

    private static object Invoke(string method, params object?[] args)
    {
        using MailService service = new MailService(new DeadGateway(), null, null);
        MethodInfo target = typeof(MailService)
            .GetMethod(method, BindingFlags.Instance | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException(
                "MailService." + method + " is gone - this test pins the mapping it performs.");

        try
        {
            return target.Invoke(service, args)
                ?? throw new InvalidOperationException(method + " returned null.");
        }
        catch (TargetInvocationException ex) when (ex.InnerException != null)
        {
            throw ex.InnerException;
        }
    }

    /// <summary>
    /// An Outlook that is not there. The mappers touch COM for one thing only - the
    /// best-effort attachment re-read, which is wrapped in a catch-all precisely so
    /// reporting can never cost a caller their draft - so this exercises that path too.
    /// </summary>
    private sealed class DeadGateway : IComGateway
    {
        public event Action? OutlookGone
        {
            add { }
            remove { }
        }

        public bool IsConnected => false;

        public bool? QuitSinkActive => null;

        public bool ProbeConnected() => false;

        public T Run<T>(Func<IOutlookSession, T> operation) => throw Gone();

        public T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery) => throw Gone();

        public T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false)
            => throw Gone();

        public ComHostDiagnostics GetDiagnostics() => new ComHostDiagnostics("in-process", "down");

        public void Dispose()
        {
        }

        private static OutlookUnavailableException Gone()
        {
            return new OutlookUnavailableException("Outlook is not running.");
        }
    }
}
