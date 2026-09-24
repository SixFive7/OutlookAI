using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

using Xunit;
using Xunit.Sdk;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// Q84 (maintainer decision (c), 2026-09-24): read-only lookups never create folders; only the
/// move-to-archive path may, and it says so.
/// <para>
/// THE DEFECT. Measured on a test guest's POP3 PST (Outlook LTSC 2024 16.0.17932), resolving
/// the designated Archive folder the way the product did - <c>Store.GetDefaultFolder(39)</c>,
/// then a verification that asked <c>GetDefaultFolder</c> for each core default folder - took
/// the store from 14 folders to 16: it CREATED <c>Archive</c>, and the verification's
/// <c>GetDefaultFolder(23)</c> CREATED <c>Junk Email</c>. The freshness sweep asked the same
/// <c>GetDefaultFolder(23)</c> on every search.
/// </para>
/// <para>
/// HOW IT IS PINNED WITHOUT A MAILBOX. <see cref="FakeStore"/> is an Outlook store modelled on
/// that measurement and on MS-OXOSFLD: <c>GetDefaultFolder</c> on a folder the store lacks
/// CREATES it and designates it (section 3.1.4.1), for every folder id - pessimistically, since
/// only 23 and 39 were measured - and everything is recorded. The resolver
/// (<see cref="SpecialFolders"/>) and the archive resolution
/// (<see cref="ArchiveFolderResolution"/>) run against it UNCHANGED, so what is asserted here is
/// the shipped decision logic. The COM half (<see cref="ComSpecialFolderStore"/>) is only
/// provable on a guest; see the report that accompanied this change.
/// </para>
/// <para>
/// THE CONTROLS. <see cref="PreFixReadOnlyResolution"/> replays the code this replaced, and on
/// the measured shape the fake reproduces the measured damage exactly - Archive and Junk Email,
/// and nothing else - which is what proves the fake can see the defect at all. The same
/// assertion every read-only test makes is then shown to FAIL on that replay. And
/// <see cref="EveryCreatingLookupInTheProduct_IsOnTheReviewedList"/> fails the build of trust
/// the moment any product member calls <c>GetDefaultFolder</c> - or the product's own creating
/// entry point - without being on a list that says why it may.
/// </para>
/// </summary>
public sealed class ReadOnlyFolderLookupTests
{
    private const string TierStore = "tier@vm.invalid";

    private const string ItemId = "01AA02BB03CC04DD05EE06FF07AA08BB09CC10DD11EE12FF13AA14BB15CC16DD17EE18FF19AA20BB"
        + "21CC22DD23EE24FF25AA26BB27CC28DD29EE30FF31AA32BB33CC34DD35EE36FF37AA38BB39CC40DD";

    private const string SecondItemId = "02AA02BB03CC04DD05EE06FF07AA08BB09CC10DD11EE12FF13AA14BB15CC16DD17EE18FF19AA20BB"
        + "21CC22DD23EE24FF25AA26BB27CC28DD29EE30FF31AA32BB33CC34DD35EE36FF37AA38BB39CC40DD";

    // ------------------------------------------------------------------ the control: the measured damage, replayed

    [Fact]
    public void Control_ThePreFixLookup_OnTheMeasuredPst_CreatesExactlyWhatTheGuestSaw()
    {
        // 14 folders before, no Archive, no Junk Email; after: Archive and Junk Email created.
        // If the fake did not reproduce this, every "created nothing" below would prove nothing.
        FakeStore pst = FakeStore.MeasuredTierPst();
        Assert.Equal(14, pst.FolderCount);

        (string? error, string? entryId, string? via) = PreFixReadOnlyResolution(pst);

        Assert.Null(error);
        Assert.Equal(pst.EntryIdOfSpecial(ArchiveFolderResolution.OlFolderArchive), entryId);
        Assert.Equal(ArchiveFolderResolution.ViaOutlookDefaultFolder, via);
        Assert.Equal(new[] { ArchiveFolderResolution.OlFolderArchive, SpecialFolders.OlFolderJunk }, pst.Created);
        Assert.Equal(16, pst.FolderCount);
    }

    [Fact]
    public void Control_TheNoCreationAssertion_FailsWhenAReadOnlyPathIsPointedBackAtGetDefaultFolder()
    {
        // The assertion the read-only tests rely on, run against a read-only path that asks
        // GetDefaultFolder the way the pre-fix code did. It must fail - otherwise the pins are
        // decoration.
        FakeStore pst = FakeStore.MeasuredTierPst();
        _ = PreFixReadOnlyResolution(pst);

        Assert.ThrowsAny<XunitException>(() => AssertCreatedNothing(pst));

        // And the narrowest possible regression - ONE direct call where the sweep asks for Junk.
        FakeStore sweepTarget = FakeStore.MeasuredTierPst();
        _ = sweepTarget.GetDefaultFolder(SpecialFolders.OlFolderJunk);
        Assert.ThrowsAny<XunitException>(() => AssertCreatedNothing(sweepTarget));
    }

    // ------------------------------------------------------------------ read-only archive resolution

    [Fact]
    public void ReadOnlyArchive_OnThePstWithoutOne_AnswersNoDesignatedArchiveFolder_AndCreatesNothing()
    {
        FakeStore pst = FakeStore.MeasuredTierPst();

        ArchiveFolderAnswer answer = ArchiveFolderResolution.ResolveReadOnly(pst, FakeStore.StoreId);

        Assert.Equal(ArchiveFolderResolution.NoDesignatedArchiveFolder, answer.Error);
        Assert.Null(answer.EntryId);
        Assert.False(answer.Created);
        AssertCreatedNothing(pst);
        Assert.DoesNotContain(ArchiveFolderResolution.OlFolderArchive, pst.GetDefaultFolderCalls);
        Assert.Equal(14, pst.FolderCount);
    }

    [Fact]
    public void ReadOnlyArchive_OnAPstThatHasOne_FindsItFromTheInboxDesignation_WithoutAskingForIt()
    {
        FakeStore pst = FakeStore.MeasuredTierPst();
        pst.AddSpecial(ArchiveFolderResolution.OlFolderArchive, "Archive");

        ArchiveFolderAnswer answer = ArchiveFolderResolution.ResolveReadOnly(pst, FakeStore.StoreId);

        Assert.Null(answer.Error);
        Assert.Equal(pst.EntryIdOfSpecial(ArchiveFolderResolution.OlFolderArchive), answer.EntryId);
        Assert.Equal("Archive", answer.Name);
        Assert.Equal(ArchiveFolderResolution.ViaInboxArchiveProperty, answer.Via);
        Assert.False(answer.Created);
        AssertCreatedNothing(pst);
        Assert.DoesNotContain(ArchiveFolderResolution.OlFolderArchive, pst.GetDefaultFolderCalls);

        // The verification ran, and asked for Junk Email - which this store does not have -
        // without making it.
        Assert.DoesNotContain(SpecialFolders.OlFolderJunk, pst.GetDefaultFolderCalls);
    }

    [Fact]
    public void ReadOnlyArchive_OnTheExchangeShape_IsExactlyWhatItWasBefore()
    {
        // "Exchange must keep working exactly as before": same folder, same via, and the SAME
        // sequence of GetDefaultFolder calls the replaced code made - and no property reads.
        FakeStore mailbox = FakeStore.ExchangeMailbox();
        FakeStore twin = FakeStore.ExchangeMailbox();

        ArchiveFolderAnswer answer = ArchiveFolderResolution.ResolveReadOnly(mailbox, FakeStore.StoreId);
        (string? preError, string? preEntryId, string? preVia) = PreFixReadOnlyResolution(twin);

        Assert.Null(answer.Error);
        Assert.Null(preError);
        Assert.Equal(preEntryId, answer.EntryId);
        Assert.Equal(preVia, answer.Via);
        Assert.Equal(ArchiveFolderResolution.ViaOutlookDefaultFolder, answer.Via);
        Assert.Equal(twin.GetDefaultFolderCalls, mailbox.GetDefaultFolderCalls);
        Assert.Equal(new[] { 39, 3, 4, 5, 6, 16, 23 }, mailbox.GetDefaultFolderCalls);
        Assert.Equal(0, mailbox.PropertyReads);
        Assert.Equal(0, mailbox.RootListings);
        AssertCreatedNothing(mailbox);
    }

    [Fact]
    public void ReadOnlyArchive_OnExchange_KeepsTheD39StorePropertyFallback()
    {
        // GetDefaultFolder(39) answering null on an Exchange store fell back to the store
        // object's PR_IPM_ARCHIVE_ENTRYID before, and still does.
        FakeStore mailbox = FakeStore.ExchangeMailbox();
        mailbox.ArchiveCallAnswersNull = true;
        mailbox.DesignatedOnStore.Add(ArchiveFolderResolution.OlFolderArchive);

        ArchiveFolderAnswer answer = ArchiveFolderResolution.ResolveReadOnly(mailbox, FakeStore.StoreId);

        Assert.Null(answer.Error);
        Assert.Equal(ArchiveFolderResolution.ViaStoreArchiveProperty, answer.Via);
        AssertCreatedNothing(mailbox);
    }

    [Fact]
    public void ReadOnlyArchive_WhoseDesignationCannotBeRead_SaysSo_AndCreatesNothing()
    {
        FakeStore pst = FakeStore.MeasuredTierPst();
        pst.AddSpecial(ArchiveFolderResolution.OlFolderArchive, "Archive");
        pst.InboxDesignationRead = PropertyReadStatus.Failed;

        ArchiveFolderAnswer answer = ArchiveFolderResolution.ResolveReadOnly(pst, FakeStore.StoreId);

        Assert.Equal(ArchiveFolderResolution.ArchiveDesignationUnreadable, answer.Error);
        AssertCreatedNothing(pst);
        Assert.DoesNotContain(ArchiveFolderResolution.OlFolderArchive, pst.GetDefaultFolderCalls);
    }

    [Fact]
    public void ReadOnlyArchive_ThatDesignatesACoreDefault_IsRefused_AndCreatesNothing()
    {
        // The paranoia check survives, without the creating calls it used to make.
        FakeStore pst = FakeStore.MeasuredTierPst();
        pst.DesignateOnInbox(ArchiveFolderResolution.OlFolderArchive, pst.EntryIdOfSpecial(SpecialFolders.OlFolderSentMail)!);

        ArchiveFolderAnswer answer = ArchiveFolderResolution.ResolveReadOnly(pst, FakeStore.StoreId);

        Assert.Equal("ArchiveFolderVerificationFailed:coreDefault", answer.Error);
        AssertCreatedNothing(pst);
    }

    // ------------------------------------------------------------------ the move path: may create, must say so

    [Fact]
    public void MoveArchive_OnThePstWithoutOne_CreatesTheArchiveOnly_AndReportsIt()
    {
        FakeStore pst = FakeStore.MeasuredTierPst();

        ArchiveFolderAnswer answer = ArchiveFolderResolution.ResolveForMove(pst, FakeStore.StoreId);

        Assert.Null(answer.Error);
        Assert.True(answer.Created);
        Assert.Equal("\\\\" + TierStore + "\\Archive", answer.FolderPath);
        Assert.Equal(ArchiveFolderResolution.ViaOutlookDefaultFolder, answer.Via);

        // Exactly one creating call, and exactly the folder the caller asked to move mail into:
        // the verification no longer makes Junk Email on the way.
        Assert.Equal(1, pst.GetDefaultFolderCalls.Count(id => id == ArchiveFolderResolution.OlFolderArchive));
        Assert.Equal(new[] { ArchiveFolderResolution.OlFolderArchive }, pst.Created);
        Assert.Equal(15, pst.FolderCount);
    }

    [Fact]
    public void MoveArchive_OnAPstThatAlreadyHasOne_ResolvesItAsBefore_AndReportsNothingCreated()
    {
        FakeStore pst = FakeStore.MeasuredTierPst();
        pst.AddSpecial(ArchiveFolderResolution.OlFolderArchive, "Archive");

        ArchiveFolderAnswer answer = ArchiveFolderResolution.ResolveForMove(pst, FakeStore.StoreId);

        Assert.Null(answer.Error);
        Assert.False(answer.Created);
        Assert.Equal(ArchiveFolderResolution.ViaOutlookDefaultFolder, answer.Via);
        AssertCreatedNothing(pst);
    }

    [Fact]
    public void MoveArchive_WhoseCreatedFolderIsRefusedByVerification_StillReportsTheCreation()
    {
        // A folder made and then refused still exists, and this server cannot delete folders.
        FakeStore pst = FakeStore.MeasuredTierPst();
        pst.CreatedArchiveItemType = 1; // a non-mail folder: the verification refuses it

        ArchiveFolderAnswer answer = ArchiveFolderResolution.ResolveForMove(pst, FakeStore.StoreId);

        Assert.Equal("ArchiveFolderVerificationFailed:itemType", answer.Error);
        Assert.True(answer.Created);
        Assert.Equal("\\\\" + TierStore + "\\Archive", answer.FolderPath);
    }

    [Fact]
    public void MoveArchive_OnExchange_IsTheReadOnlyResolution_AndNeverListsTheRoot()
    {
        FakeStore mailbox = FakeStore.ExchangeMailbox();

        ArchiveFolderAnswer answer = ArchiveFolderResolution.ResolveForMove(mailbox, FakeStore.StoreId);

        Assert.Null(answer.Error);
        Assert.False(answer.Created);
        Assert.Equal(ArchiveFolderResolution.ViaOutlookDefaultFolder, answer.Via);
        Assert.Equal(new[] { 39, 3, 4, 5, 6, 16, 23 }, mailbox.GetDefaultFolderCalls);
        Assert.Equal(0, mailbox.RootListings);
    }

    [Fact]
    public void MoveArchive_ThatCannotListTheRootFirst_DoesNotAskAtAll()
    {
        // Without the "before" list a creation could not be recognised, so it is not risked.
        FakeStore pst = FakeStore.MeasuredTierPst();
        pst.RootListingFails = true;

        ArchiveFolderAnswer answer = ArchiveFolderResolution.ResolveForMove(pst, FakeStore.StoreId);

        Assert.Equal(ArchiveFolderResolution.ArchiveFolderStateUnreadable, answer.Error);
        Assert.Empty(pst.GetDefaultFolderCalls);
        AssertCreatedNothing(pst);
    }

    [Fact]
    public void MoveArchive_WhenOutlookReturnsNothing_ClaimsNothingCreatedOnlyIfTheRootProvesIt()
    {
        FakeStore unchanged = FakeStore.MeasuredTierPst();
        unchanged.ArchiveCallAnswersNull = true;
        Assert.Equal(
            ArchiveFolderResolution.NoDesignatedArchiveFolder,
            ArchiveFolderResolution.ResolveForMove(unchanged, FakeStore.StoreId).Error);

        FakeStore changed = FakeStore.MeasuredTierPst();
        changed.ArchiveCallAnswersNull = true;
        changed.ArchiveCallAddsUnreturnedRootFolder = true;
        Assert.Equal(
            ArchiveFolderResolution.ArchiveFolderCreationUnverified,
            ArchiveFolderResolution.ResolveForMove(changed, FakeStore.StoreId).Error);
    }

    // ------------------------------------------------------------------ the resolver the sweep and the others use

    [Theory]
    [InlineData(SpecialFolders.OlFolderInbox)]
    [InlineData(SpecialFolders.OlFolderSentMail)]
    [InlineData(SpecialFolders.OlFolderDeletedItems)]
    [InlineData(SpecialFolders.OlFolderOutbox)]
    [InlineData(SpecialFolders.OlFolderDrafts)]
    public void Resolve_OnTheMeasuredPst_FindsWhatItHas_WithoutCreatingAnything(int folderId)
    {
        FakeStore pst = FakeStore.MeasuredTierPst();

        OutlookComSession.DefaultFolderResolution resolution = SpecialFolders.Resolve(pst, folderId, out object? folder, out _);

        Assert.Equal(OutlookComSession.DefaultFolderResolution.Resolved, resolution);
        Assert.Equal(pst.EntryIdOfSpecial(folderId), pst.EntryIdOf(folder!));
        AssertCreatedNothing(pst);
    }

    [Theory]
    [InlineData(SpecialFolders.OlFolderJunk)]
    [InlineData(ArchiveFolderResolution.OlFolderArchive)]
    [InlineData(SpecialFolders.OlFolderConflicts)]
    [InlineData(SpecialFolders.OlFolderSyncIssues)]
    [InlineData(SpecialFolders.OlFolderLocalFailures)]
    [InlineData(SpecialFolders.OlFolderServerFailures)]
    public void Resolve_OnTheMeasuredPst_AnswersAbsentForWhatItLacks_WithoutAskingForIt(int folderId)
    {
        // The sweep's Junk Email (23) is the one that used to be made on every search.
        FakeStore pst = FakeStore.MeasuredTierPst();

        OutlookComSession.DefaultFolderResolution resolution = SpecialFolders.Resolve(pst, folderId, out object? folder, out _);

        Assert.Equal(OutlookComSession.DefaultFolderResolution.Absent, resolution);
        Assert.Null(folder);
        Assert.DoesNotContain(folderId, pst.GetDefaultFolderCalls);
        AssertCreatedNothing(pst);
    }

    [Fact]
    public void Resolve_TheSweepsFourFolders_OnTheMeasuredPst_CreatesNothing()
    {
        // Inbox, Sent Items, Deleted Items, Junk Email - the default sweep scope, in its order.
        FakeStore pst = FakeStore.MeasuredTierPst();
        List<OutlookComSession.DefaultFolderResolution> verdicts = new List<OutlookComSession.DefaultFolderResolution>();
        foreach (int folderId in new[] { 6, 5, 3, 23 })
        {
            verdicts.Add(SpecialFolders.Resolve(pst, folderId, out object? folder, out _));
            pst.Release(folder);
        }

        Assert.Equal(
            new[]
            {
                OutlookComSession.DefaultFolderResolution.Resolved,
                OutlookComSession.DefaultFolderResolution.Resolved,
                OutlookComSession.DefaultFolderResolution.Resolved,
                OutlookComSession.DefaultFolderResolution.Absent,
            },
            verdicts);
        AssertCreatedNothing(pst);
        Assert.Equal(14, pst.FolderCount);
    }

    [Fact]
    public void Resolve_OnAPstThatIsNotADeliveryStore_NeverAsksForAnInbox()
    {
        // A data file with Deleted Items only: its mask says so, and nothing else is asked for.
        FakeStore dataFile = FakeStore.NonDeliveryPst();

        foreach (int folderId in new[] { 6, 5, 4, 16, 23, 39 })
        {
            Assert.Equal(
                OutlookComSession.DefaultFolderResolution.Absent,
                SpecialFolders.Resolve(dataFile, folderId, out _, out _));
        }

        Assert.Equal(
            OutlookComSession.DefaultFolderResolution.Resolved,
            SpecialFolders.Resolve(dataFile, SpecialFolders.OlFolderDeletedItems, out object? deleted, out _));
        dataFile.Release(deleted);

        Assert.Equal(new[] { SpecialFolders.OlFolderDeletedItems }, dataFile.GetDefaultFolderCalls);
        AssertCreatedNothing(dataFile);
    }

    [Theory]
    [InlineData(PropertyReadStatus.Failed)]
    [InlineData(PropertyReadStatus.NotFound)]
    public void Resolve_WhenTheStoreWillNotSayWhatItHas_IsUnreadable_AndAsksForNothing(PropertyReadStatus mask)
    {
        // Unreadable, never absent - and never a guess that ends in a creation.
        FakeStore pst = FakeStore.MeasuredTierPst();
        pst.MaskRead = mask;

        foreach (int folderId in new[] { 6, 5, 3, 4, 16, 23, 39 })
        {
            Assert.Equal(
                OutlookComSession.DefaultFolderResolution.Unreadable,
                SpecialFolders.Resolve(pst, folderId, out _, out _));
        }

        Assert.Empty(pst.GetDefaultFolderCalls);
        AssertCreatedNothing(pst);
    }

    [Fact]
    public void Resolve_ADesignationPointingAtAFolderThatIsGone_IsAbsent_AndCreatesNothing()
    {
        FakeStore pst = FakeStore.MeasuredTierPst();
        pst.DesignateOnInbox(SpecialFolders.OlFolderJunk, "00000000DEADBEEFDEADBEEFDEADBEEFDEADBEEF01020304");

        Assert.Equal(
            OutlookComSession.DefaultFolderResolution.Absent,
            SpecialFolders.Resolve(pst, SpecialFolders.OlFolderJunk, out _, out _));
        AssertCreatedNothing(pst);
    }

    [Fact]
    public void Resolve_ADesignatedJunkFolder_IsFoundThroughTheInbox_NotThroughGetDefaultFolder()
    {
        FakeStore pst = FakeStore.MeasuredTierPst();
        pst.AddSpecial(SpecialFolders.OlFolderJunk, "Junk Email");

        Assert.Equal(
            OutlookComSession.DefaultFolderResolution.Resolved,
            SpecialFolders.Resolve(pst, SpecialFolders.OlFolderJunk, out object? junk, out SpecialFolderSource source));

        Assert.Equal(SpecialFolderSource.InboxDesignation, source);
        Assert.Equal(pst.EntryIdOfSpecial(SpecialFolders.OlFolderJunk), pst.EntryIdOf(junk!));
        Assert.DoesNotContain(SpecialFolders.OlFolderJunk, pst.GetDefaultFolderCalls);
    }

    [Fact]
    public void Resolve_OnExchange_IsTheCallItAlwaysWas()
    {
        FakeStore mailbox = FakeStore.ExchangeMailbox();

        foreach (int folderId in new[] { 6, 5, 3, 23 })
        {
            Assert.Equal(
                OutlookComSession.DefaultFolderResolution.Resolved,
                SpecialFolders.Resolve(mailbox, folderId, out _, out SpecialFolderSource source));
            Assert.Equal(SpecialFolderSource.DefaultFolderCall, source);
        }

        Assert.Equal(new[] { 6, 5, 3, 23 }, mailbox.GetDefaultFolderCalls);
        Assert.Equal(0, mailbox.PropertyReads);
    }

    [Fact]
    public void Resolve_AnUnreadableStoreType_TakesTheBranchThatNeverCreates()
    {
        FakeStore pst = FakeStore.MeasuredTierPst();
        pst.ExchangeStoreTypeValue = null;

        Assert.False(SpecialFolders.IsExchangeStore(null));
        Assert.Equal(
            OutlookComSession.DefaultFolderResolution.Absent,
            SpecialFolders.Resolve(pst, SpecialFolders.OlFolderJunk, out _, out _));
        AssertCreatedNothing(pst);
    }

    [Fact]
    public void Resolve_AFolderNothingDocumentsTheLocationOf_IsUnreadable_NotAsked()
    {
        // Calendar (9) has no designation this resolver reads: its existence cannot be
        // established without the creating call, so the call is not made.
        FakeStore pst = FakeStore.MeasuredTierPst();

        Assert.Equal(
            OutlookComSession.DefaultFolderResolution.Unreadable,
            SpecialFolders.Resolve(pst, 9, out _, out _));
        Assert.Empty(pst.GetDefaultFolderCalls);
    }

    // ------------------------------------------------------------------ the documented layouts

    [Fact]
    public void TheValidFolderMaskBits_AreMapiDefsValues()
    {
        // Microsoft's MAPIDefS.h: FOLDER_IPM_INBOX_VALID 0x02, OUTBOX 0x04, WASTEBASKET 0x08,
        // SENTMAIL 0x10. A wrong bit here reads "absent" for a folder that exists.
        Assert.Equal(0x02, SpecialFolders.ValidFolderBit(SpecialFolders.OlFolderInbox));
        Assert.Equal(0x04, SpecialFolders.ValidFolderBit(SpecialFolders.OlFolderOutbox));
        Assert.Equal(0x08, SpecialFolders.ValidFolderBit(SpecialFolders.OlFolderDeletedItems));
        Assert.Equal(0x10, SpecialFolders.ValidFolderBit(SpecialFolders.OlFolderSentMail));
        Assert.Null(SpecialFolders.ValidFolderBit(SpecialFolders.OlFolderJunk));
    }

    [Fact]
    public void TheAdditionalRenIndexes_AreTheMsOxosfldTable()
    {
        // MS-OXOSFLD 2.2.4: 0 Conflicts, 1 Sync Issues, 2 Local Failures, 3 Server Failures,
        // 4 Junk E-mail; OlDefaultFolders 19, 20, 21, 22, 23.
        Assert.Equal(0, SpecialFolders.AdditionalRenIndex(19));
        Assert.Equal(1, SpecialFolders.AdditionalRenIndex(20));
        Assert.Equal(2, SpecialFolders.AdditionalRenIndex(21));
        Assert.Equal(3, SpecialFolders.AdditionalRenIndex(22));
        Assert.Equal(4, SpecialFolders.AdditionalRenIndex(23));
        Assert.Null(SpecialFolders.AdditionalRenIndex(ArchiveFolderResolution.OlFolderArchive));
    }

    [Fact]
    public void ThePropertyTags_AreTheDocumentedOnes()
    {
        // MS-OXPROPS: PidTagIpmArchiveEntryId 0x35FF PtypBinary (2.752), PidTagIpmDraftsEntryId
        // 0x36D7 PtypBinary (2.754), PidTagAdditionalRenEntryIds 0x36D8 PtypMultipleBinary
        // (2.509); MAPITags.h: PR_VALID_FOLDER_MASK 0x35DF PT_LONG.
        Assert.Equal("http://schemas.microsoft.com/mapi/proptag/0x35FF0102", ArchiveFolderResolution.ArchiveEntryIdPropertySchema);
        Assert.Equal("http://schemas.microsoft.com/mapi/proptag/0x36D70102", SpecialFolders.DraftsEntryIdSchema);
        Assert.Equal("http://schemas.microsoft.com/mapi/proptag/0x36D81102", SpecialFolders.AdditionalRenEntryIdsSchema);
        Assert.Equal("http://schemas.microsoft.com/mapi/proptag/0x35DF0003", SpecialFolders.ValidFolderMaskSchema);
        Assert.Equal(SpecialFolders.DraftsEntryIdSchema, SpecialFolders.DesignationSchema(SpecialFolders.OlFolderDrafts));
        Assert.Equal(ArchiveFolderResolution.ArchiveEntryIdPropertySchema, SpecialFolders.DesignationSchema(39));
        Assert.Equal(SpecialFolders.AdditionalRenEntryIdsSchema, SpecialFolders.DesignationSchema(23));
    }

    [Fact]
    public void JunkIsReadAtIndexFour_OfTheShapeTheLiveProbeRecorded()
    {
        // D39's live probe of an Inbox: the classic five slots plus a 4-byte non-EntryID trailer.
        object[] measured =
        {
            Bytes(0x10), Bytes(0x11), Bytes(0x12), Bytes(0x13), Bytes(0x14), new byte[] { 1, 2, 3, 4 },
        };

        DesignatedEntryId junk = SpecialFolders.ReadEntryIdAt(measured, 4);

        Assert.Equal(DesignatedEntryIdKind.EntryId, junk.Kind);
        Assert.Equal(Convert.ToHexString(Bytes(0x14)), junk.Hex);

        // Reading the first usable slot - what the archive helper does - would pick Conflicts.
        Assert.NotEqual(junk.Hex, ArchiveFolderResolution.TryReadEntryIdHex(measured));
    }

    [Fact]
    public void AnEmptySlot_OrAnArrayTooShort_DesignatesNothing()
    {
        Assert.Equal(DesignatedEntryIdKind.None, SpecialFolders.ReadEntryIdAt(new object[] { Bytes(1), Bytes(2) }, 4).Kind);
        Assert.Equal(DesignatedEntryIdKind.None, SpecialFolders.ReadEntryIdAt(new object[] { Bytes(1), Bytes(2), Bytes(3), Bytes(4), Array.Empty<byte>() }, 4).Kind);
        Assert.Equal(DesignatedEntryIdKind.None, SpecialFolders.ReadEntryIdAt(new object?[] { null, null, null, null, null }, 4).Kind);
        Assert.Equal(DesignatedEntryIdKind.None, SpecialFolders.ReadEntryIdAt(null, 4).Kind);
    }

    [Fact]
    public void AValueOfTheWrongShape_IsUnrecognised_NotAbsent()
    {
        // A single byte array read by position would pick a BYTE; a string is not this property.
        Assert.Equal(DesignatedEntryIdKind.Unrecognised, SpecialFolders.ReadEntryIdAt(Bytes(7), 4).Kind);
        Assert.Equal(DesignatedEntryIdKind.Unrecognised, SpecialFolders.ReadEntryIdAt("00112233", 0).Kind);
        Assert.Equal(DesignatedEntryIdKind.Unrecognised, SpecialFolders.ReadEntryIdAt(new object[] { "x", "y", "z", "w", 42 }, 4).Kind);
        Assert.Equal(DesignatedEntryIdKind.Unrecognised, SpecialFolders.ReadEntryId(42).Kind);
        Assert.Equal(DesignatedEntryIdKind.Unrecognised, SpecialFolders.ReadEntryId("not hex").Kind);
    }

    [Fact]
    public void ASingleBinaryDesignation_ReadsAsItAlwaysDid()
    {
        Assert.Equal(Convert.ToHexString(Bytes(9)), SpecialFolders.ReadEntryId(Bytes(9)).Hex);
        Assert.Equal(DesignatedEntryIdKind.None, SpecialFolders.ReadEntryId(new byte[] { 1, 2 }).Kind);
        Assert.Equal(DesignatedEntryIdKind.None, SpecialFolders.ReadEntryId(null).Kind);
    }

    [Fact]
    public void TheMaskReadsAsAnInteger_AndOnlyAsOne()
    {
        Assert.True(SpecialFolders.TryReadMask(0xFF, out int bits));
        Assert.Equal(0xFF, bits);
        Assert.True(SpecialFolders.TryReadMask((short)0x1E, out bits));
        Assert.Equal(0x1E, bits);
        Assert.False(SpecialFolders.TryReadMask("255", out _));
        Assert.False(SpecialFolders.TryReadMask(null, out _));
    }

    [Fact]
    public void OnlyMapiNotFound_ReadsAsNotThere()
    {
        Assert.True(SpecialFolders.IsMapiNotFound(new COMException("The property is unknown or cannot be found.", SpecialFolders.MapiENotFound)));
        Assert.False(SpecialFolders.IsMapiNotFound(new COMException("RPC server unavailable", unchecked((int)0x800706BA))));
        Assert.False(SpecialFolders.IsMapiNotFound(new InvalidCastException()));
        Assert.False(SpecialFolders.IsMapiNotFound(null));
    }

    [Fact]
    public void OnlyANonExchangeStore_TakesTheProvingBranch()
    {
        Assert.True(SpecialFolders.IsExchangeStore(0));  // olPrimaryExchangeMailbox
        Assert.True(SpecialFolders.IsExchangeStore(1));  // olExchangeMailbox (delegate)
        Assert.True(SpecialFolders.IsExchangeStore(2));  // olExchangePublicFolder
        Assert.True(SpecialFolders.IsExchangeStore(4));  // olAdditionalExchangeMailbox
        Assert.False(SpecialFolders.IsExchangeStore(3)); // olNotExchange
    }

    // ------------------------------------------------------------------ the contract and the tool

    [Fact]
    public void TheReadOnlyResolution_StaysRetryable_AndTheCreatingOneDoesNot()
    {
        Assert.True(ComSessionOperations.IsRetryable(nameof(IOutlookSession.TryResolveArchiveFolder)));
        Assert.False(ComSessionOperations.IsRetryable(nameof(IOutlookSession.TryResolveOrCreateArchiveFolder)));
        Assert.Contains(nameof(IOutlookSession.TryResolveOrCreateArchiveFolder), ComSessionOperations.MutatingOperations);
    }

    [Fact]
    public void ArchiveMail_UsesTheCreatingResolution_AndReportsAFolderItMadeEvenWhenNothingMoved()
    {
        ArchiveSession session = new ArchiveSession { CreatedFolderPath = "Archive", MoveRefusal = "TargetFolderNotFound" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        ArchiveMailOutcome outcome = service.ArchiveMail(new[] { ItemId });

        Assert.Equal(0, outcome.Archived);
        Assert.Equal(new[] { TierStore + "/Archive" }, outcome.CreatedFolders);
        Assert.Contains(nameof(IOutlookSession.TryResolveOrCreateArchiveFolder), session.Calls);
        Assert.DoesNotContain(nameof(IOutlookSession.TryResolveArchiveFolder), session.Calls);
    }

    [Fact]
    public void ArchiveMail_RefusedAfterCreating_NamesTheFolderInTheItemError_AndTheOutcome()
    {
        ArchiveSession session = new ArchiveSession
        {
            CreatedFolderPath = "Archive",
            ResolveError = "ArchiveFolderVerificationFailed:itemType",
        };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        ArchiveMailOutcome outcome = service.ArchiveMail(new[] { ItemId });

        MoveItemView item = Assert.Single(outcome.Items);
        Assert.False(item.Ok);
        Assert.Equal(MutationOutcome.Unchanged, item.Outcome);
        Assert.Contains("CREATED before this failed", item.Error!, StringComparison.Ordinal);
        Assert.Contains("Archive", item.Error!, StringComparison.Ordinal);
        Assert.Equal(new[] { TierStore + "/Archive" }, outcome.CreatedFolders);
        Assert.Null(outcome.ArchiveFolders);
    }

    [Fact]
    public void ArchiveMail_ThatCreatedNothing_ReportsNoCreatedFolders()
    {
        ArchiveSession session = new ArchiveSession { MoveRefusal = "TargetFolderNotFound" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        ArchiveMailOutcome outcome = service.ArchiveMail(new[] { ItemId });

        Assert.Null(outcome.CreatedFolders);
    }

    [Fact]
    public void ArchiveMail_ResolvesEachStoreOnce_SoACreationIsReportedOnce()
    {
        ArchiveSession session = new ArchiveSession { CreatedFolderPath = "Archive", MoveRefusal = "TargetFolderNotFound" };
        using MailService service = new MailService(new DirectGateway(session.AsSession));

        ArchiveMailOutcome outcome = service.ArchiveMail(new[] { ItemId, SecondItemId });

        Assert.Equal(1, session.Calls.Count(c => c == nameof(IOutlookSession.TryResolveOrCreateArchiveFolder)));
        Assert.Equal(new[] { TierStore + "/Archive" }, outcome.CreatedFolders);
    }

    [Theory]
    [InlineData(ArchiveFolderResolution.NoDesignatedArchiveFolder, "Nothing was created")]
    [InlineData(ArchiveFolderResolution.ArchiveFolderStateUnreadable, "Nothing was moved or created")]
    [InlineData(ArchiveFolderResolution.ArchiveDesignationUnreadable, "Nothing was created")]
    [InlineData(ArchiveFolderResolution.ArchiveFolderCreationUnverified, "may have CREATED")]
    public void EveryArchiveRefusal_SaysWhatItCanProveAboutCreation(string token, string claim)
    {
        string text = MailService.DescribeArchiveResolutionFailure("Store X", token);

        Assert.Contains("Store X", text, StringComparison.Ordinal);
        Assert.Contains(claim, text, StringComparison.Ordinal);
    }

    [Fact]
    public void TheUnverifiedCreation_NeverClaimsNothingWasCreated()
    {
        string text = MailService.DescribeArchiveResolutionFailure("Store X", ArchiveFolderResolution.ArchiveFolderCreationUnverified);

        Assert.DoesNotContain("Nothing was created", text, StringComparison.Ordinal);
        Assert.Contains("Nothing was moved", text, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ the write tools' guards (decision A, direction 1)
    //
    // move_mail's "not the Outbox, not Deleted Items or under it" and update_draft's/discard_draft's
    // "in the Drafts folder" only COMPARE folder identities, so they never need the folder to
    // exist. They run here unchanged against the same fake stores, and the control at the end of
    // the section runs the guards as they were against the same shapes.

    [Fact]
    public void Control_TheOldGuards_OnStoresLackingTheirFolders_CreateThem_AndTheAssertionSeesIt()
    {
        // The fake makes every missing folder it is asked for: measured for Drafts on a guest's
        // data file, and pessimistic for the Outbox, which the same file was measured NOT to get.
        // Either way, if the old guards did not fail AssertCreatedNothing here, the no-creation
        // pins on the new ones below would be decoration.
        FakeStore bare = FakeStore.WithoutSpecialFolders();
        FakeFolder target = bare.AddPlain("Projects");
        _ = PreFixMoveGuard(bare, target.EntryId, Chain(target.EntryId));
        Assert.Equal(new[] { SpecialFolders.OlFolderDeletedItems, SpecialFolders.OlFolderOutbox }, bare.Created);
        Assert.ThrowsAny<XunitException>(() => AssertCreatedNothing(bare));

        FakeStore dataFile = FakeStore.NonDeliveryPst();
        FakeFolder folder = dataFile.AddPlain("Projects");
        _ = PreFixDraftsGate(dataFile, Chain(folder.EntryId));
        Assert.Equal(new[] { SpecialFolders.OlFolderDrafts }, dataFile.Created);
        Assert.ThrowsAny<XunitException>(() => AssertCreatedNothing(dataFile));
    }

    [Fact]
    public void MoveGuard_OnAStoreWithNeitherFolder_AsksForNeither_AndRefusesNothingOnThatGround()
    {
        // Nothing can be moved into a Deleted Items or an Outbox the store does not have.
        FakeStore bare = FakeStore.WithoutSpecialFolders();
        FakeFolder target = bare.AddPlain("Projects");

        Assert.Null(SpecialFolderGuards.MoveTargetRefusal(bare, target.EntryId, Chain(target.EntryId)));

        Assert.Empty(bare.GetDefaultFolderCalls);
        AssertCreatedNothing(bare);
    }

    [Fact]
    public void MoveGuard_OnADataFileThatIsNoDeliveryStore_AsksOnlyForTheFolderItHas()
    {
        FakeStore dataFile = FakeStore.NonDeliveryPst();
        FakeFolder target = dataFile.AddPlain("Projects");

        Assert.Null(SpecialFolderGuards.MoveTargetRefusal(dataFile, target.EntryId, Chain(target.EntryId)));

        // Deleted Items is proven by the store's mask and asked for; the Outbox is not there, and is not.
        Assert.Equal(new[] { SpecialFolders.OlFolderDeletedItems }, dataFile.GetDefaultFolderCalls);
        AssertCreatedNothing(dataFile);
    }

    [Fact]
    public void MoveGuard_StillRefusesTheOutbox_AndDeletedItems_AndEverythingUnderIt()
    {
        FakeStore pst = FakeStore.MeasuredTierPst();
        string outbox = pst.EntryIdOfSpecial(SpecialFolders.OlFolderOutbox)!;
        string deleted = pst.EntryIdOfSpecial(SpecialFolders.OlFolderDeletedItems)!;
        FakeFolder underDeleted = pst.AddPlain("Old");
        FakeFolder plain = pst.AddPlain("Projects");

        Assert.Equal(SpecialFolderGuards.TargetIsOutbox, SpecialFolderGuards.MoveTargetRefusal(pst, outbox, Chain(outbox)));
        Assert.Equal(SpecialFolderGuards.TargetIsDeletedItems, SpecialFolderGuards.MoveTargetRefusal(pst, deleted, Chain(deleted)));
        Assert.Equal(
            SpecialFolderGuards.TargetIsDeletedItems,
            SpecialFolderGuards.MoveTargetRefusal(pst, underDeleted.EntryId, Chain(underDeleted.EntryId, deleted)));
        Assert.Null(SpecialFolderGuards.MoveTargetRefusal(pst, plain.EntryId, Chain(plain.EntryId)));

        AssertCreatedNothing(pst);
    }

    [Theory]
    [InlineData(PropertyReadStatus.Failed)]
    [InlineData(PropertyReadStatus.NotFound)]
    public void MoveGuard_FailsClosed_WhenTheStoreWillNotSayWhichFoldersItHas(PropertyReadStatus mask)
    {
        FakeStore pst = FakeStore.MeasuredTierPst();
        pst.MaskRead = mask;
        FakeFolder plain = pst.AddPlain("Projects");

        Assert.Equal(
            SpecialFolderGuards.TargetGuardUnreadable,
            SpecialFolderGuards.MoveTargetRefusal(pst, plain.EntryId, Chain(plain.EntryId)));

        Assert.Empty(pst.GetDefaultFolderCalls);
        AssertCreatedNothing(pst);
    }

    [Fact]
    public void MoveGuard_FailsClosed_WhenAFolderResolvesButWillNotSayWhoItIs()
    {
        FakeStore pst = FakeStore.MeasuredTierPst();
        pst.EntryIdUnreadable.Add(pst.EntryIdOfSpecial(SpecialFolders.OlFolderOutbox)!);
        FakeFolder plain = pst.AddPlain("Projects");

        Assert.Equal(
            SpecialFolderGuards.TargetGuardUnreadable,
            SpecialFolderGuards.MoveTargetRefusal(pst, plain.EntryId, Chain(plain.EntryId)));
    }

    [Fact]
    public void MoveGuard_NamesAProvenRefusal_EvenWhenTheOtherLookupFailed()
    {
        // Refused either way; the reason given is the one that is actually known.
        FakeStore pst = FakeStore.MeasuredTierPst();
        pst.EntryIdUnreadable.Add(pst.EntryIdOfSpecial(SpecialFolders.OlFolderOutbox)!);
        string deleted = pst.EntryIdOfSpecial(SpecialFolders.OlFolderDeletedItems)!;
        FakeFolder underDeleted = pst.AddPlain("Old");

        Assert.Equal(
            SpecialFolderGuards.TargetIsDeletedItems,
            SpecialFolderGuards.MoveTargetRefusal(pst, underDeleted.EntryId, Chain(underDeleted.EntryId, deleted)));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    public void MoveGuard_FailsClosed_OnATargetThatWillNotSayWhoItIs(string? targetEntryId)
    {
        FakeStore pst = FakeStore.MeasuredTierPst();

        Assert.Equal(SpecialFolderGuards.TargetGuardUnreadable, SpecialFolderGuards.MoveTargetRefusal(pst, targetEntryId, Chain()));
    }

    [Fact]
    public void MoveGuard_OnTheExchangeShape_AsksExactlyWhatTheOldGuardAsked_AndAnswersTheSame()
    {
        // "Exchange stays unchanged": the same GetDefaultFolder calls in the same order, no
        // property reads, and the same verdict for every target the old guard could decide.
        FakeStore mailbox = FakeStore.ExchangeMailbox();
        FakeStore twin = FakeStore.ExchangeMailbox();
        string outbox = mailbox.EntryIdOfSpecial(SpecialFolders.OlFolderOutbox)!;
        string deleted = mailbox.EntryIdOfSpecial(SpecialFolders.OlFolderDeletedItems)!;
        string underDeleted = mailbox.AddPlain("Old").EntryId;
        string plain = mailbox.AddPlain("Projects").EntryId;
        Assert.Equal(underDeleted, twin.AddPlain("Old").EntryId);
        Assert.Equal(plain, twin.AddPlain("Projects").EntryId);

        (string Target, string[] Folders, string? Expected)[] cases =
        {
            (outbox, new[] { outbox }, SpecialFolderGuards.TargetIsOutbox),
            (deleted, new[] { deleted }, SpecialFolderGuards.TargetIsDeletedItems),
            (underDeleted, new[] { underDeleted, deleted }, SpecialFolderGuards.TargetIsDeletedItems),
            (plain, new[] { plain }, null),
        };
        foreach ((string target, string[] chain, string? expected) in cases)
        {
            Assert.Equal(expected, SpecialFolderGuards.MoveTargetRefusal(mailbox, target, Chain(chain)));
            Assert.Equal(expected, PreFixMoveGuard(twin, target, Chain(chain)));
        }

        Assert.Equal(twin.GetDefaultFolderCalls, mailbox.GetDefaultFolderCalls);
        Assert.Equal(new[] { 3, 4, 3, 4, 3, 4, 3, 4 }, mailbox.GetDefaultFolderCalls);
        Assert.Equal(0, mailbox.PropertyReads);
        AssertCreatedNothing(mailbox);
    }

    [Fact]
    public void MoveGuard_OnExchange_ALookupThatFails_NowRefuses_WhereTheOldGuardLetTheMoveThrough()
    {
        // The one Exchange difference, and it IS the fail-closed rule: GetDefaultFolder(3)
        // failing used to read as "no Deleted Items", so a move INTO Deleted Items went ahead.
        FakeStore mailbox = FakeStore.ExchangeMailbox();
        FakeStore twin = FakeStore.ExchangeMailbox();
        mailbox.DefaultFolderCallFails.Add(SpecialFolders.OlFolderDeletedItems);
        twin.DefaultFolderCallFails.Add(SpecialFolders.OlFolderDeletedItems);
        string deleted = mailbox.EntryIdOfSpecial(SpecialFolders.OlFolderDeletedItems)!;

        Assert.Null(PreFixMoveGuard(twin, deleted, Chain(deleted)));
        Assert.Equal(SpecialFolderGuards.TargetGuardUnreadable, SpecialFolderGuards.MoveTargetRefusal(mailbox, deleted, Chain(deleted)));
        Assert.Equal(twin.GetDefaultFolderCalls, mailbox.GetDefaultFolderCalls);
    }

    [Fact]
    public void DraftsGate_OnAStoreWithoutDrafts_IsNotADraft_AndCreatesNothing()
    {
        // A store with no Drafts folder holds no draft.
        FakeStore dataFile = FakeStore.NonDeliveryPst();
        FakeFolder folder = dataFile.AddPlain("Projects");

        Assert.Equal(SpecialFolderGuards.NotInDraftsFolder, SpecialFolderGuards.DraftsFolderRefusal(dataFile, Chain(folder.EntryId)));

        Assert.DoesNotContain(SpecialFolders.OlFolderDrafts, dataFile.GetDefaultFolderCalls);
        AssertCreatedNothing(dataFile);
    }

    [Fact]
    public void DraftsGate_FindsDraftsThroughTheInboxDesignation_WithoutAskingForIt()
    {
        FakeStore pst = FakeStore.MeasuredTierPst();
        string drafts = pst.EntryIdOfSpecial(SpecialFolders.OlFolderDrafts)!;
        string inbox = pst.EntryIdOfSpecial(SpecialFolders.OlFolderInbox)!;
        FakeFolder underDrafts = pst.AddPlain("Held");

        Assert.Null(SpecialFolderGuards.DraftsFolderRefusal(pst, Chain(drafts)));
        Assert.Null(SpecialFolderGuards.DraftsFolderRefusal(pst, Chain(underDrafts.EntryId, drafts)));
        Assert.Equal(SpecialFolderGuards.NotInDraftsFolder, SpecialFolderGuards.DraftsFolderRefusal(pst, Chain(inbox)));

        Assert.DoesNotContain(SpecialFolders.OlFolderDrafts, pst.GetDefaultFolderCalls);
        AssertCreatedNothing(pst);
    }

    [Theory]
    [InlineData(PropertyReadStatus.Failed, PropertyReadStatus.Found)]
    [InlineData(PropertyReadStatus.NotFound, PropertyReadStatus.Found)]
    [InlineData(PropertyReadStatus.Found, PropertyReadStatus.Failed)]
    public void DraftsGate_FailsClosed_WhenWhereDraftsIsCannotBeRead(PropertyReadStatus mask, PropertyReadStatus inboxDesignation)
    {
        FakeStore pst = FakeStore.MeasuredTierPst();
        pst.MaskRead = mask;
        pst.InboxDesignationRead = inboxDesignation;
        string drafts = pst.EntryIdOfSpecial(SpecialFolders.OlFolderDrafts)!;

        // Even for an item that IS in Drafts: unproven is refused, never assumed.
        Assert.Equal(SpecialFolderGuards.DraftsFolderUnreadable, SpecialFolderGuards.DraftsFolderRefusal(pst, Chain(drafts)));

        Assert.DoesNotContain(SpecialFolders.OlFolderDrafts, pst.GetDefaultFolderCalls);
        AssertCreatedNothing(pst);
    }

    [Fact]
    public void DraftsGate_AnInboxThatDesignatesNoDrafts_ReadsAsNoDraftsFolder()
    {
        // THE ASSUMPTION A GUEST HAS TO CONFIRM. On a store that is not Exchange the gate finds
        // Drafts where MS-OXOSFLD 2.2.3 puts it for the mailbox owner - designated on the Inbox,
        // or on the store object. A store whose designation lives anywhere else reads as having
        // no Drafts folder, and every draft in it is refused as not_in_drafts_folder.
        FakeStore pst = FakeStore.MeasuredTierPst();
        pst.InboxDesignationRead = PropertyReadStatus.NotFound;
        string drafts = pst.EntryIdOfSpecial(SpecialFolders.OlFolderDrafts)!;

        Assert.Equal(SpecialFolderGuards.NotInDraftsFolder, SpecialFolderGuards.DraftsFolderRefusal(pst, Chain(drafts)));
        AssertCreatedNothing(pst);
    }

    [Fact]
    public void DraftsGate_OnTheExchangeShape_AsksExactlyWhatTheOldGateAsked_AndAnswersTheSame()
    {
        FakeStore mailbox = FakeStore.ExchangeMailbox();
        FakeStore twin = FakeStore.ExchangeMailbox();
        string drafts = mailbox.EntryIdOfSpecial(SpecialFolders.OlFolderDrafts)!;
        string inbox = mailbox.EntryIdOfSpecial(SpecialFolders.OlFolderInbox)!;

        Assert.Null(SpecialFolderGuards.DraftsFolderRefusal(mailbox, Chain(drafts)));
        Assert.Null(PreFixDraftsGate(twin, Chain(drafts)));
        Assert.Equal(SpecialFolderGuards.NotInDraftsFolder, SpecialFolderGuards.DraftsFolderRefusal(mailbox, Chain(inbox)));
        Assert.Equal(SpecialFolderGuards.NotInDraftsFolder, PreFixDraftsGate(twin, Chain(inbox)));

        Assert.Equal(twin.GetDefaultFolderCalls, mailbox.GetDefaultFolderCalls);
        Assert.Equal(new[] { 16, 16 }, mailbox.GetDefaultFolderCalls);
        Assert.Equal(0, mailbox.PropertyReads);
        AssertCreatedNothing(mailbox);
    }

    [Fact]
    public void DraftsGate_OnExchange_ALookupThatFails_IsStillRefused_NowWithTheTrueReason()
    {
        // The old gate refused too, but by claiming the item was not in Drafts - which it had
        // not checked. Same refusal; the reason now says the check could not be made.
        FakeStore mailbox = FakeStore.ExchangeMailbox();
        FakeStore twin = FakeStore.ExchangeMailbox();
        mailbox.DefaultFolderCallFails.Add(SpecialFolders.OlFolderDrafts);
        twin.DefaultFolderCallFails.Add(SpecialFolders.OlFolderDrafts);
        string drafts = mailbox.EntryIdOfSpecial(SpecialFolders.OlFolderDrafts)!;

        Assert.Equal(SpecialFolderGuards.NotInDraftsFolder, PreFixDraftsGate(twin, Chain(drafts)));
        Assert.Equal(SpecialFolderGuards.DraftsFolderUnreadable, SpecialFolderGuards.DraftsFolderRefusal(mailbox, Chain(drafts)));
    }

    [Theory]
    [InlineData(SpecialFolderGuards.TargetIsOutbox)]
    [InlineData(SpecialFolderGuards.TargetIsDeletedItems)]
    [InlineData(SpecialFolderGuards.TargetGuardUnreadable)]
    public void EveryMoveGuardRefusal_IsNamed_AndProvablyMovedNothing(string code)
    {
        // An unnamed code falls into the catch-all, which tells the caller the move MAY have
        // happened. Every guard refusal is decided before Move() is called.
        Assert.Equal(MutationOutcome.Unchanged, MailService.MoveFailureOutcome(code));
        Assert.DoesNotContain(
            "UNKNOWN",
            MailService.DescribeMoveFailure(code, "Projects", null, createFolder: false),
            StringComparison.Ordinal);
    }

    /// <summary>A folder chain, the item's or target's own folder first: true for any EntryID on it.</summary>
    private static Func<string, bool> Chain(params string[] entryIds)
    {
        return wanted => entryIds.Any(entryId => string.Equals(entryId, wanted, StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// move_mail's target guard as it was, transliterated onto <see cref="ISpecialFolderStore"/>
    /// call for call: <c>GetDefaultFolder(3)</c>, then <c>(4)</c>, a failed lookup read as "no
    /// such folder" and its check skipped.
    /// </summary>
    private static string? PreFixMoveGuard(FakeStore store, string targetEntryId, Func<string, bool> targetIsOrIsUnder)
    {
        string? deletedItemsEntryId = PreFixEntryId(store, SpecialFolders.OlFolderDeletedItems);
        string? outboxEntryId = PreFixEntryId(store, SpecialFolders.OlFolderOutbox);
        if (outboxEntryId != null && string.Equals(targetEntryId, outboxEntryId, StringComparison.OrdinalIgnoreCase))
        {
            return SpecialFolderGuards.TargetIsOutbox;
        }

        return deletedItemsEntryId != null && targetIsOrIsUnder(deletedItemsEntryId)
            ? SpecialFolderGuards.TargetIsDeletedItems
            : null;
    }

    /// <summary>The update_draft/discard_draft gate as it was: <c>GetDefaultFolder(16)</c>, a failed lookup read as "not in Drafts".</summary>
    private static string? PreFixDraftsGate(FakeStore store, Func<string, bool> itemFolderIsOrIsUnder)
    {
        string? draftsEntryId = PreFixEntryId(store, SpecialFolders.OlFolderDrafts);
        return draftsEntryId != null && itemFolderIsOrIsUnder(draftsEntryId) ? null : SpecialFolderGuards.NotInDraftsFolder;
    }

    private static string? PreFixEntryId(FakeStore store, int folderId)
    {
        try
        {
            object? folder = store.GetDefaultFolder(folderId);
            return folder == null ? null : store.EntryIdOf(folder);
        }
        catch (COMException)
        {
            return null;
        }
    }

    // ------------------------------------------------------------------ the source-level control

    /// <summary>
    /// Every member of the shipped product that may call a CREATING lookup, and why. A member
    /// not on this list that calls <c>.GetDefaultFolder(</c>, <c>.GetSharedDefaultFolder(</c> or
    /// <c>GetDefaultFolderMayCreate(</c> fails the test - which is what happens if a read-only
    /// path (the sweep, navigation, the read-only archive lookup and its verification, the probes,
    /// the Outbox count, the default-folder info) or one of the write tools' guards (move_mail's
    /// Deleted Items/Outbox check, the update_draft/discard_draft Drafts gate - off this list since
    /// decision A) is pointed back at <c>GetDefaultFolder</c>.
    /// </summary>
    private static readonly Dictionary<string, string> ReviewedCreatingLookups = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["McpServer/OutlookAI.Core/Com/SpecialFolders.cs::CallDefaultFolder"] =
            "the resolver's ONE call: an Exchange store as before, any other store only once the folder is proven to exist",
        ["McpServer/OutlookAI.Core/Com/SpecialFolders.cs::GetDefaultFolder"] =
            "ComSpecialFolderStore's COM implementation of ISpecialFolderStore.GetDefaultFolder",
        ["McpServer/OutlookAI.Core/Com/SpecialFolders.cs::GetDefaultFolderMayCreate"] =
            "the declaration of the product's explicit creating entry point",
        ["McpServer/OutlookAI.Core/Com/ArchiveFolderResolution.cs::ResolveForMove"] =
            "archive_mail: the one path allowed to create (Q84 (c)); it reports what it created",
        ["McpServer/OutlookAI.Core/Com/ComposeSurface.cs::TryPinProcess"] =
            "NOT CHANGED (D49 process pin): NameSpace.GetDefaultFolder(6), the default store's Inbox; see the Q84 report",
        ["McpServer/OutlookAI.Core/Com/OutlookComSession.cs::EnsureVisibleExplorer"] =
            "NOT CHANGED (display tools): NameSpace.GetDefaultFolder(6), the default store's Inbox; see the Q84 report",
        ["McpServer/OutlookAI.Core/Com/OutlookComSession.cs::TryCreateNewDraft"] =
            "WRITE PATH, not changed: the new draft's destination, Drafts of the account's delivery store",
        ["McpServer/OutlookAI.Core/Com/OutlookComSession.cs::TryCreateDerivedDraft"] =
            "WRITE PATH, not changed: relocating a reply/forward into its source store's Drafts; see the Q84 report",
        ["McpServer/OutlookAI.Core/Com/OutlookComSession.cs::TryDiscardDraft"] =
            "WRITE PATH, not changed: naming the Deleted Items a discarded draft went to",
    };

    [Fact]
    public void EveryCreatingLookupInTheProduct_IsOnTheReviewedList()
    {
        List<string> found = FindCreatingLookups();

        Assert.True(found.Count > 0, "the scan found no GetDefaultFolder call at all - it has stopped proving anything");

        List<string> unreviewed = found.Distinct(StringComparer.Ordinal)
            .Where(site => !ReviewedCreatingLookups.ContainsKey(site))
            .OrderBy(site => site, StringComparer.Ordinal)
            .ToList();
        Assert.True(
            unreviewed.Count == 0,
            "These product members call a lookup that CREATES a folder on a store that lacks it (measured on a PST "
            + "for GetDefaultFolder 23 and 39), and are not on the reviewed list: " + string.Join(", ", unreviewed)
            + ". A read-only path must use SpecialFolders.Resolve instead (Q84). Only the move-to-archive path may create.");

        List<string> stale = ReviewedCreatingLookups.Keys
            .Where(site => !found.Contains(site, StringComparer.Ordinal))
            .OrderBy(site => site, StringComparer.Ordinal)
            .ToList();
        Assert.True(
            stale.Count == 0,
            "These reviewed sites no longer call a creating lookup - take them off the list so it keeps meaning something: "
            + string.Join(", ", stale));
    }

    [Theory]
    [InlineData("ResolveDefaultFolder")]
    [InlineData("CountOutboxItems")]
    [InlineData("ResolveNavigationFolder")]
    [InlineData("ResolveProbeFolder")]
    [InlineData("TryGetDefaultFolderInfo")]
    [InlineData("ResolveArchiveFolderCore")]
    public void EveryReadOnlyLookup_GoesThroughTheResolver(string member)
    {
        // The positive half of the scan above: the members that used to call GetDefaultFolder
        // directly now reach SpecialFolders.Resolve, which is pinned behaviourally above.
        string body = MemberBody("McpServer/OutlookAI.Core/Com/OutlookComSession.cs", member);

        Assert.True(
            body.Contains("SpecialFolders.Resolve(", StringComparison.Ordinal)
                || body.Contains("ResolveDefaultFolder(special", StringComparison.Ordinal)
                || body.Contains("ArchiveFolderResolution.ResolveReadOnly(", StringComparison.Ordinal),
            member + " no longer reaches the non-creating resolver");
        Assert.DoesNotContain(".GetDefaultFolder(", body, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("VerifyMoveTarget", "SpecialFolderGuards.MoveTargetRefusal(")]
    [InlineData("CheckInDraftsFolder", "SpecialFolderGuards.DraftsFolderRefusal(")]
    public void EveryWriteToolGuard_GoesThroughTheNonCreatingGuard(string member, string guard)
    {
        // Decision A, direction 1: the two checks that only compare folder identities left the
        // reviewed list above. This is the positive half - each reaches its guard, which is
        // pinned behaviourally in the section on the write tools' guards.
        string body = MemberBody("McpServer/OutlookAI.Core/Com/OutlookComSession.cs", member);

        Assert.Contains(guard, body, StringComparison.Ordinal);
        Assert.DoesNotContain(".GetDefaultFolder(", body, StringComparison.Ordinal);
    }

    // ------------------------------------------------------------------ support

    /// <summary>The one assertion every read-only test makes.</summary>
    private static void AssertCreatedNothing(FakeStore store)
    {
        Assert.True(
            store.Created.Count == 0,
            "a read-only lookup CREATED folder(s) " + string.Join(", ", store.Created) + " in the store");
    }

    /// <summary>
    /// The read-only archive resolution this change replaced, transliterated onto
    /// <see cref="ISpecialFolderStore"/> call for call: <c>GetDefaultFolder(39)</c>, the store
    /// property fallback, then the verification that asked <c>GetDefaultFolder</c> for Deleted
    /// Items, Outbox, Sent Items, Inbox, Drafts and Junk in that order.
    /// </summary>
    private static (string? Error, string? EntryId, string? Via) PreFixReadOnlyResolution(FakeStore store)
    {
        object? folder;
        string via = ArchiveFolderResolution.ViaOutlookDefaultFolder;
        try
        {
            folder = store.GetDefaultFolder(ArchiveFolderResolution.OlFolderArchive);
        }
        catch (COMException)
        {
            folder = null;
        }

        if (folder == null)
        {
            via = ArchiveFolderResolution.ViaStoreArchiveProperty;
            PropertyRead read = store.ReadStoreProperty(ArchiveFolderResolution.ArchiveEntryIdPropertySchema);
            string? hex = read.Status == PropertyReadStatus.Found ? ArchiveFolderResolution.TryReadEntryIdHex(read.Value) : null;
            if (hex != null && store.OpenFolder(hex, out object? opened) == PropertyReadStatus.Found)
            {
                folder = opened;
            }
        }

        if (folder == null)
        {
            return (ArchiveFolderResolution.NoDesignatedArchiveFolder, null, null);
        }

        SpecialFolderFacts facts = store.Describe(folder);
        if (!string.Equals(facts.StoreId, FakeStore.StoreId, StringComparison.OrdinalIgnoreCase))
        {
            return ("ArchiveFolderVerificationFailed:store", null, null);
        }

        if (facts.DefaultItemType != 0)
        {
            return ("ArchiveFolderVerificationFailed:itemType", null, null);
        }

        foreach (int coreDefault in new[] { 3, 4, 5, 6, 16, 23 })
        {
            try
            {
                object? defaultFolder = store.GetDefaultFolder(coreDefault);
                if (string.Equals(store.EntryIdOf(defaultFolder!), facts.EntryId, StringComparison.OrdinalIgnoreCase))
                {
                    return ("ArchiveFolderVerificationFailed:coreDefault", null, null);
                }
            }
            catch (COMException)
            {
            }
        }

        return (null, facts.EntryId, via);
    }

    private static byte[] Bytes(byte seed)
    {
        byte[] bytes = new byte[24];
        for (int i = 0; i < bytes.Length; i++)
        {
            bytes[i] = (byte)(seed + i);
        }

        return bytes;
    }

    private static string RepoRoot()
    {
        string testProjectDir = typeof(ReadOnlyFolderLookupTests).Assembly
                .GetCustomAttributes<AssemblyMetadataAttribute>()
                .FirstOrDefault(a => a.Key == "TestProjectDir")?.Value
            ?? throw new InvalidOperationException("AssemblyMetadata 'TestProjectDir' is missing.");

        // <repo>/McpServer/OutlookAI.McpServer.Tests/ -> <repo>
        return Path.GetFullPath(Path.Combine(testProjectDir, "..", ".."));
    }

    /// <summary>A creating call, or the product's own creating entry point, on a line of code.</summary>
    private static readonly Regex CreatingLookup = new Regex(
        @"\.GetDefaultFolder\s*\(|\.GetSharedDefaultFolder\s*\(|\bGetDefaultFolderMayCreate\s*\(",
        RegexOptions.CultureInvariant);

    /// <summary>
    /// A member declaration at class-member indentation: EXACTLY eight spaces in these sources,
    /// so a statement indented deeper (an <c>if (</c> inside a method) is never taken for one.
    /// </summary>
    private static readonly Regex MemberDeclaration = new Regex(
        @"^ {8}(?![ /\[{}#])[^=;]*?\b(?<name>\w+)\s*(?:<[^>]*>)?\s*\(",
        RegexOptions.CultureInvariant);

    /// <summary>
    /// Every product line that calls a creating lookup, as <c>file::member</c>. Scans the three
    /// MCP server projects and the add-in; RemediationTools (an operator tool) and the tests are
    /// not the product.
    /// </summary>
    private static List<string> FindCreatingLookups()
    {
        string root = RepoRoot();
        List<string> files = new List<string>();
        foreach (string project in new[] { "OutlookAI.Core", "OutlookAI.McpServer", "OutlookAI.ComHost" })
        {
            string dir = Path.Combine(root, "McpServer", project);
            Assert.True(Directory.Exists(dir), "product project is missing: " + dir);
            files.AddRange(Directory.EnumerateFiles(dir, "*.cs", SearchOption.AllDirectories)
                .Where(f => !IsBuildOutput(f)));
        }

        files.AddRange(Directory.EnumerateFiles(root, "*.cs", SearchOption.TopDirectoryOnly));
        foreach (string addInDir in new[] { "Services", "TaskPane" })
        {
            files.AddRange(Directory.EnumerateFiles(Path.Combine(root, addInDir), "*.cs", SearchOption.AllDirectories));
        }

        List<string> sites = new List<string>();
        foreach (string file in files)
        {
            string relative = Path.GetRelativePath(root, file).Replace('\\', '/');
            string[] lines = File.ReadAllLines(file);
            for (int i = 0; i < lines.Length; i++)
            {
                string trimmed = lines[i].TrimStart();
                if (trimmed.StartsWith("//", StringComparison.Ordinal) || trimmed.StartsWith("*", StringComparison.Ordinal))
                {
                    continue;
                }

                if (CreatingLookup.IsMatch(lines[i]))
                {
                    sites.Add(relative + "::" + EnclosingMember(lines, i));
                }
            }
        }

        return sites;
    }

    private static bool IsBuildOutput(string path)
    {
        string normalized = path.Replace('\\', '/');
        return normalized.Contains("/bin/", StringComparison.OrdinalIgnoreCase)
            || normalized.Contains("/obj/", StringComparison.OrdinalIgnoreCase);
    }

    private static string EnclosingMember(string[] lines, int index)
    {
        for (int i = index; i >= 0; i--)
        {
            Match declaration = MemberDeclaration.Match(lines[i]);
            if (declaration.Success)
            {
                return declaration.Groups["name"].Value;
            }
        }

        return "<no member>";
    }

    /// <summary>The source of one member: its declaration line through the next member declaration.</summary>
    private static string MemberBody(string relativePath, string member)
    {
        string[] lines = File.ReadAllLines(Path.Combine(RepoRoot(), relativePath));
        int start = Array.FindIndex(lines, line =>
        {
            Match declaration = MemberDeclaration.Match(line);
            return declaration.Success && declaration.Groups["name"].Value == member;
        });
        Assert.True(start >= 0, member + " was not found in " + relativePath + " - this test has stopped proving anything");

        int end = start + 1;
        while (end < lines.Length && !MemberDeclaration.IsMatch(lines[end]))
        {
            end++;
        }

        return string.Join("\n", lines, start, end - start);
    }

    /// <summary>One folder of a <see cref="FakeStore"/>.</summary>
    private sealed class FakeFolder
    {
        internal FakeFolder(string entryId, string name, string folderPath, int defaultItemType)
        {
            EntryId = entryId;
            Name = name;
            FolderPath = folderPath;
            DefaultItemType = defaultItemType;
        }

        internal string EntryId { get; }

        internal string Name { get; }

        internal string FolderPath { get; }

        internal int DefaultItemType { get; }
    }

    /// <summary>
    /// An Outlook store as far as special folders go: which exist, where each is designated
    /// (the store's PR_VALID_FOLDER_MASK for Inbox/Outbox/Deleted/Sent; the Inbox's binary
    /// identification properties and PR_ADDITIONAL_REN_ENTRYIDS for the rest), and what
    /// <c>GetDefaultFolder</c> does about one that is missing - it MAKES it at the root and
    /// designates it, as measured for 23 and 39 on a POP3 PST and as MS-OXOSFLD 3.1.4.1 tells a
    /// client to. Every call is recorded.
    /// </summary>
    private sealed class FakeStore : ISpecialFolderStore
    {
        internal const string StoreId = "0000000038A1BB1005E5101AA1BB08002B2A56C20000";

        private readonly string _displayName;
        private readonly Dictionary<int, FakeFolder> _special = new Dictionary<int, FakeFolder>();
        private readonly Dictionary<string, FakeFolder> _byEntryId = new Dictionary<string, FakeFolder>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<int, string> _inboxDesignations = new Dictionary<int, string>();
        private readonly List<string> _rootChildren = new List<string>();
        private int _nextId = 0x100;

        private FakeStore(string displayName, int? exchangeStoreType)
        {
            _displayName = displayName;
            ExchangeStoreTypeValue = exchangeStoreType;
        }

        internal int? ExchangeStoreTypeValue { get; set; }

        internal PropertyReadStatus MaskRead { get; set; } = PropertyReadStatus.Found;

        internal PropertyReadStatus InboxDesignationRead { get; set; } = PropertyReadStatus.Found;

        internal bool RootListingFails { get; set; }

        internal bool ArchiveCallAnswersNull { get; set; }

        internal bool ArchiveCallAddsUnreturnedRootFolder { get; set; }

        internal int CreatedArchiveItemType { get; set; }

        internal HashSet<int> DesignatedOnStore { get; } = new HashSet<int>();

        /// <summary>Folder ids whose <c>GetDefaultFolder</c> call fails, the way a COM call does.</summary>
        internal HashSet<int> DefaultFolderCallFails { get; } = new HashSet<int>();

        /// <summary>Entry ids of folders that open but will not say their EntryID.</summary>
        internal HashSet<string> EntryIdUnreadable { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        internal List<int> GetDefaultFolderCalls { get; } = new List<int>();

        internal List<int> Created { get; } = new List<int>();

        internal int PropertyReads { get; private set; }

        internal int RootListings { get; private set; }

        internal int FolderCount => _byEntryId.Count;

        public int? ExchangeStoreType => ExchangeStoreTypeValue;

        /// <summary>
        /// The measured guest store: a POP3 PST with 14 folders and no Archive or Junk Email.
        /// Which nine besides the five core defaults it held was not recorded, so they are
        /// modelled as plain folders.
        /// </summary>
        internal static FakeStore MeasuredTierPst()
        {
            FakeStore store = new FakeStore(TierStore, SpecialFolders.OlNotExchange);
            store.AddSpecial(SpecialFolders.OlFolderDeletedItems, "Deleted Items");
            store.AddSpecial(SpecialFolders.OlFolderInbox, "Inbox");
            store.AddSpecial(SpecialFolders.OlFolderOutbox, "Outbox");
            store.AddSpecial(SpecialFolders.OlFolderSentMail, "Sent Items");
            store.AddSpecial(SpecialFolders.OlFolderDrafts, "Drafts");
            foreach (string plain in new[] { "Calendar", "Contacts", "Journal", "Notes", "Tasks", "RSS Feeds", "Conversation History", "Search Folders", "Tier Plain" })
            {
                store.AddPlain(plain);
            }

            return store;
        }

        /// <summary>A data file that is no account's delivery store: Deleted Items and nothing else special.</summary>
        internal static FakeStore NonDeliveryPst()
        {
            FakeStore store = new FakeStore("Outlook Data File", SpecialFolders.OlNotExchange);
            store.AddSpecial(SpecialFolders.OlFolderDeletedItems, "Deleted Items");
            store.AddPlain("Search Folders");
            return store;
        }

        /// <summary>A data file with no special folder at all - not even Deleted Items.</summary>
        internal static FakeStore WithoutSpecialFolders()
        {
            FakeStore store = new FakeStore("Bare Data File", SpecialFolders.OlNotExchange);
            store.AddPlain("Search Folders");
            return store;
        }

        /// <summary>A primary Exchange mailbox: every default folder the server keeps is there.</summary>
        internal static FakeStore ExchangeMailbox()
        {
            FakeStore store = new FakeStore("someone@example.com", 0);
            foreach ((int id, string name) in new[]
            {
                (3, "Deleted Items"), (4, "Outbox"), (5, "Sent Items"), (6, "Inbox"), (16, "Drafts"), (23, "Junk Email"), (39, "Archive"),
            })
            {
                store.AddSpecial(id, name);
            }

            return store;
        }

        internal string? EntryIdOfSpecial(int folderId)
        {
            return _special.TryGetValue(folderId, out FakeFolder? folder) ? folder.EntryId : null;
        }

        /// <summary>Adds a special folder at the root, designated where the store type designates it.</summary>
        internal FakeFolder AddSpecial(int folderId, string name, int defaultItemType = 0)
        {
            FakeFolder folder = AddPlain(name, defaultItemType);
            _special[folderId] = folder;
            if (SpecialFolders.ValidFolderBit(folderId) == null)
            {
                _inboxDesignations[folderId] = folder.EntryId;
            }

            return folder;
        }

        internal FakeFolder AddPlain(string name, int defaultItemType = 0)
        {
            string entryId = "0000000038A1BB1005E5101AA1BB08002B2A56C2" + (_nextId++).ToString("X8", System.Globalization.CultureInfo.InvariantCulture);
            FakeFolder folder = new FakeFolder(entryId, name, "\\\\" + _displayName + "\\" + name, defaultItemType);
            _byEntryId[entryId] = folder;
            _rootChildren.Add(entryId);
            return folder;
        }

        /// <summary>Designates an entry id on the Inbox - one that may name no folder at all.</summary>
        internal void DesignateOnInbox(int folderId, string entryId)
        {
            _inboxDesignations[folderId] = entryId;
        }

        public PropertyRead ReadStoreProperty(string schemaName)
        {
            PropertyReads++;
            if (schemaName == SpecialFolders.ValidFolderMaskSchema)
            {
                if (MaskRead == PropertyReadStatus.NotFound)
                {
                    return PropertyRead.Missing();
                }

                if (MaskRead == PropertyReadStatus.Failed)
                {
                    return PropertyRead.Failure();
                }

                int mask = 0x01; // FOLDER_IPM_SUBTREE_VALID
                foreach (int id in new[] { 3, 4, 5, 6 })
                {
                    if (_special.ContainsKey(id))
                    {
                        mask |= SpecialFolders.ValidFolderBit(id)!.Value;
                    }
                }

                return PropertyRead.Found(mask);
            }

            if (schemaName == ArchiveFolderResolution.ArchiveEntryIdPropertySchema
                && DesignatedOnStore.Contains(ArchiveFolderResolution.OlFolderArchive)
                && _special.TryGetValue(ArchiveFolderResolution.OlFolderArchive, out FakeFolder? archive))
            {
                return PropertyRead.Found(Convert.FromHexString(archive.EntryId));
            }

            return PropertyRead.Missing();
        }

        public PropertyRead ReadFolderProperty(object folder, string schemaName)
        {
            PropertyReads++;
            if (!_special.TryGetValue(SpecialFolders.OlFolderInbox, out FakeFolder? inbox) || !ReferenceEquals(folder, inbox))
            {
                return PropertyRead.Missing();
            }

            if (InboxDesignationRead != PropertyReadStatus.Found)
            {
                return InboxDesignationRead == PropertyReadStatus.NotFound ? PropertyRead.Missing() : PropertyRead.Failure();
            }

            if (schemaName == SpecialFolders.AdditionalRenEntryIdsSchema)
            {
                bool any = false;
                object[] slots = new object[6];
                for (int index = 0; index < 5; index++)
                {
                    if (_inboxDesignations.TryGetValue(19 + index, out string? entryId))
                    {
                        slots[index] = Convert.FromHexString(entryId);
                        any = true;
                    }
                    else
                    {
                        slots[index] = Array.Empty<byte>();
                    }
                }

                slots[5] = new byte[] { 0, 0, 0, 0 };
                return any ? PropertyRead.Found(slots) : PropertyRead.Missing();
            }

            int? folderId = schemaName == SpecialFolders.DraftsEntryIdSchema
                ? SpecialFolders.OlFolderDrafts
                : schemaName == ArchiveFolderResolution.ArchiveEntryIdPropertySchema
                    ? ArchiveFolderResolution.OlFolderArchive
                    : null;
            return folderId.HasValue && _inboxDesignations.TryGetValue(folderId.Value, out string? designated)
                ? PropertyRead.Found(Convert.FromHexString(designated))
                : PropertyRead.Missing();
        }

        public object? GetDefaultFolder(int olDefaultFolderId)
        {
            GetDefaultFolderCalls.Add(olDefaultFolderId);
            if (DefaultFolderCallFails.Contains(olDefaultFolderId))
            {
                throw new COMException("The operation failed.", unchecked((int)0x80004005));
            }

            if (olDefaultFolderId == ArchiveFolderResolution.OlFolderArchive && ArchiveCallAnswersNull)
            {
                if (ArchiveCallAddsUnreturnedRootFolder)
                {
                    AddPlain("Archive");
                }

                return null;
            }

            if (_special.TryGetValue(olDefaultFolderId, out FakeFolder? existing))
            {
                return existing;
            }

            // What Outlook did on the guest: make the folder the store lacks, and designate it.
            string name = olDefaultFolderId switch
            {
                ArchiveFolderResolution.OlFolderArchive => "Archive",
                SpecialFolders.OlFolderJunk => "Junk Email",
                _ => "Special " + olDefaultFolderId.ToString(System.Globalization.CultureInfo.InvariantCulture),
            };
            int itemType = olDefaultFolderId == ArchiveFolderResolution.OlFolderArchive ? CreatedArchiveItemType : 0;
            FakeFolder made = AddSpecial(olDefaultFolderId, name, itemType);
            Created.Add(olDefaultFolderId);
            return made;
        }

        public PropertyReadStatus OpenFolder(string entryIdHex, out object? folder)
        {
            if (_byEntryId.TryGetValue(entryIdHex, out FakeFolder? found))
            {
                folder = found;
                return PropertyReadStatus.Found;
            }

            folder = null;
            return PropertyReadStatus.NotFound;
        }

        public IReadOnlyList<string>? ListRootChildEntryIds()
        {
            RootListings++;
            return RootListingFails ? null : _rootChildren.ToList();
        }

        public string? EntryIdOf(object folder)
        {
            string entryId = ((FakeFolder)folder).EntryId;
            return EntryIdUnreadable.Contains(entryId) ? null : entryId;
        }

        public SpecialFolderFacts Describe(object folder)
        {
            FakeFolder f = (FakeFolder)folder;
            return new SpecialFolderFacts(f.EntryId, f.Name, f.FolderPath, StoreId, f.DefaultItemType);
        }

        public void Release(object? comObject)
        {
        }
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
    /// The archive_mail side of the contract: one item in the tier store, a resolution that
    /// reports what it created, and a move that is refused - so no audit line is ever written.
    /// </summary>
    private sealed class ArchiveSession
    {
        internal ArchiveSession()
        {
            AsSession = Proxy.Create(this);
        }

        internal IOutlookSession AsSession { get; }

        internal List<string> Calls { get; } = new List<string>();

        internal string? CreatedFolderPath { get; set; }

        internal string? ResolveError { get; set; }

        internal string? MoveRefusal { get; set; }

        private object? Handle(MethodInfo method, object?[]? args)
        {
            Calls.Add(method.Name);
            switch (method.Name)
            {
                case nameof(IOutlookSession.GetStoreDetails):
                    return Array.Empty<ComStoreDetail>();

                case nameof(IOutlookSession.TryGetMailInfo):
                    return new ComDraftInfo(
                        ItemId, TierStore, "store-tier", "Inbox", "folder-inbox", "A subject",
                        "someone@example.com", null, "conv-1", Array.Empty<ComRecipientInfo>(), "A subject");

                case nameof(IOutlookSession.TryResolveOrCreateArchiveFolder):
                    SetOut(method, args, "createdFolderPath", CreatedFolderPath);
                    if (ResolveError != null)
                    {
                        SetOut(method, args, "error", ResolveError);
                        return null;
                    }

                    return new ComArchiveFolderInfo(TierStore, "store-tier", "folder-archive", "Archive", "Archive", "outlookDefaultFolder");

                case nameof(IOutlookSession.TryMoveItemToFolderId):
                    SetOut(method, args, "error", MoveRefusal);
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
            private ArchiveSession _owner = null!;

            internal static IOutlookSession Create(ArchiveSession owner)
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
