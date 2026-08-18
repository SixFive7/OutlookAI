using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// How the freshness sweep tells "this store has no Junk Email folder" from "this folder
/// is there and will not open" (<see cref="OutlookComSession.ClassifyDefaultFolder"/>),
/// and what each verdict does to the coverage counters, the gap codes and the advice.
/// <para>
/// WHAT IS PROVEN HERE AND WHAT IS NOT. The classification is a pure function and every
/// case of it is pinned below, including the accounting the sweep loop builds on top of
/// it. The COM call that PRODUCES the signal needs a real mailbox, so the answers are
/// modelled here: null for a store without the folder, a caught COM-call failure for one
/// that will not open. That half rests on the documented contract of
/// Store.GetDefaultFolder - "If the default folder of the requested type does not exist,
/// GetDefaultFolder returns Null (Nothing in Visual Basic)" - and is exercised for real by
/// the live tier (T2 LiveSweepScopeTests, whose swept/skipped/absent sum must equal the
/// default folder set).
/// </para>
/// <para>
/// THE DEFECT. Resolution used to be one try/catch: a null answer was dereferenced one
/// line later, the C# dynamic binder turned that into a RuntimeBinderException, and the
/// catch counted it as a SKIPPED folder - identical to a folder that genuinely would not
/// open. Once folders_skipped became one of the coverage codes that set
/// <c>degraded: true</c> and <c>freshness: "partial"</c>, every non-folder-scoped search on
/// a profile holding one store without a Junk Email or Deleted Items folder reported itself
/// degraded, and <c>degraded</c> is the field the search tool tells an agent to relay to the
/// user. A flag that cries wolf is worse than no flag.
/// </para>
/// </summary>
public sealed class DefaultFolderResolutionTests
{
    /// <summary>MAPI_E_NOT_FOUND - the HRESULT most likely to tempt someone into reading a failure as absence.</summary>
    private const int MapiNotFound = unchecked((int)0x8004010F);

    /// <summary>RPC_S_SERVER_UNAVAILABLE - Outlook went away mid-call (live-observed shape).</summary>
    private const int RpcServerUnavailable = unchecked((int)0x800706BA);

    [Fact]
    public void ResolvedFolder_IsSwept()
    {
        // The everyday case: the store handed the folder over.
        Assert.Equal(
            OutlookComSession.DefaultFolderResolution.Resolved,
            OutlookComSession.ClassifyDefaultFolder(new object(), failure: null));
    }

    [Fact]
    public void NullAnswer_IsAbsence_BecauseThatIsWhatTheApiReturnsForAFolderTheStoreDoesNotHave()
    {
        // The signal the whole fix keys on, and it is the RETURN VALUE, not an error:
        // Store.GetDefaultFolder returns Null when the default folder of the requested type
        // does not exist in the store (Office VBA reference, "Return value"). Nothing is
        // thrown, so nothing about the failure path can be used to recognise this case.
        Assert.Equal(
            OutlookComSession.DefaultFolderResolution.Absent,
            OutlookComSession.ClassifyDefaultFolder(folder: null, failure: null));
    }

    [Fact]
    public void EveryComCallFailure_IsUnreadable_NotAbsence()
    {
        // The classifier is handed exactly the exception shapes OutlookComSession.
        // IsComCallFailure admits, because those are the only ones the resolver catches;
        // anything else propagates and degrades the whole sweep, as it always did.
        List<Exception> failures = new List<Exception>
        {
            new COMException("folder will not open", RpcServerUnavailable),
            new COMException("operation failed", unchecked((int)0x80004005)),
            new ArgumentException("late-bound E_INVALIDARG"),
            new InvalidCastException("late-bound type mismatch"),
            new MissingMemberException("Could not get dispatch ID for GetDefaultFolder"),
            new InvalidComObjectException("detached RCW"),
            new Microsoft.CSharp.RuntimeBinder.RuntimeBinderException(
                "Cannot perform runtime binding on a null reference"),
        };

        foreach (Exception failure in failures)
        {
            Assert.Equal(
                OutlookComSession.DefaultFolderResolution.Unreadable,
                OutlookComSession.ClassifyDefaultFolder(folder: null, failure));
        }

        // The last entry above is the exact shape the OLD code produced for an ABSENT
        // folder: null came back, the next line dereferenced it dynamically, and the binder
        // threw. That is why absence and unreadability were indistinguishable - the null was
        // converted into an exception before anyone could look at it.
    }

    [Fact]
    public void AFailureThatLooksLikeNotFound_IsStillUnreadable()
    {
        // Deliberate, and the fail-safe direction: reading a thrown "not found" as absence
        // would silently drop a folder whose mail really is missing from the answer, while
        // reading absence as a failure only over-reports. Only the documented null return
        // means absence.
        Assert.Equal(
            OutlookComSession.DefaultFolderResolution.Unreadable,
            OutlookComSession.ClassifyDefaultFolder(folder: null, new COMException("not found", MapiNotFound)));
    }

    [Fact]
    public void AFailureWins_EvenIfSomethingCameBackWithIt()
    {
        // Defensive: a call that both returned and failed is not a store telling us the
        // folder does not exist, so it can never be counted as absence.
        Assert.Equal(
            OutlookComSession.DefaultFolderResolution.Unreadable,
            OutlookComSession.ClassifyDefaultFolder(new object(), new COMException("half-open", RpcServerUnavailable)));
    }

    [Fact]
    public void AStoreMissingItsJunkFolder_ContributesThreeSweptAndNothingSkipped()
    {
        // THE case that made every search report itself degraded: Inbox, Sent Items and
        // Deleted Items open; the store has no Junk Email folder at all.
        SweepAccounting tally = SweepDefaultFolders(
            Answer.Resolved(), Answer.Resolved(), Answer.Resolved(), Answer.Absent());

        Assert.Equal(3, tally.Swept);
        Assert.Equal(0, tally.Skipped);
        Assert.Equal(1, tally.Absent);
    }

    [Fact]
    public void AnUnreadableFolder_IsStillSkipped()
    {
        // The other half of the split, unchanged: a folder that is there and will not open
        // has no freshness coverage, so it must keep being counted and reported.
        SweepAccounting tally = SweepDefaultFolders(
            Answer.Resolved(), Answer.Resolved(), Answer.Failed(), Answer.Resolved());

        Assert.Equal(3, tally.Swept);
        Assert.Equal(1, tally.Skipped);
        Assert.Equal(0, tally.Absent);
    }

    [Fact]
    public void TheThreeOutcomesTogether_AccountForTheWholeDefaultFolderSet()
    {
        // The accounting rule: whatever happens per folder, swept + skipped + absent equals
        // the folder set the sweep set out to cover. Nothing may vanish from the arithmetic,
        // and an absent folder may not be booked as a failure.
        SweepAccounting tally = SweepDefaultFolders(
            Answer.Resolved(), Answer.Failed(), Answer.Absent(), Answer.Absent());

        Assert.Equal(1, tally.Swept);
        Assert.Equal(1, tally.Skipped);
        Assert.Equal(2, tally.Absent);
        Assert.Equal(
            OutlookComSession.DefaultSweepFolderKinds.Count,
            tally.Swept + tally.Skipped + tally.Absent);
    }

    [Fact]
    public void AnAbsentFolder_IsCarriedOutOfTheComLayerApartFromSkippedFolders()
    {
        ComSweepResult result = new ComSweepResult(
            Array.Empty<ComMailBrief>(),
            foldersSwept: 3,
            foldersSkipped: 0,
            sweptFolders: new[] { "alice@example.com/Inbox", "alice@example.com/Sent Items", "alice@example.com/Deleted Items" },
            foldersAbsent: 1);

        Assert.Equal(0, result.FoldersSkipped);
        Assert.Equal(0, result.FoldersFailed);
        Assert.Equal(1, result.FoldersAbsent);

        // A folder-scoped sweep is asked for a NAMED folder, so it never reports absence -
        // a named folder that does not resolve is a skip, and the default is 0.
        ComSweepResult scoped = new ComSweepResult(Array.Empty<ComMailBrief>(), foldersSwept: 1, foldersSkipped: 0);
        Assert.Equal(0, scoped.FoldersAbsent);
    }

    [Fact]
    public void AnAbsentFolder_RaisesNoGapCode_AndKeepsTheSearchUndegraded()
    {
        // The end of the chain, and the whole point of the fix.
        SweepInfo sweep = new SweepInfo
        {
            Performed = true,
            Scope = MailService.DefaultSweepScopeDescription,
            FoldersSwept = 3,
            FoldersSkipped = 0,
            FoldersAbsent = 1,
        };

        sweep.CoverageGaps = FreshMerge.DescribeCoverageGaps(sweep);

        Assert.Null(sweep.CoverageGaps);
        Assert.Equal(FreshMerge.FreshnessLive, FreshMerge.ClassifyFreshness(sweep));
        Assert.Empty(MailService.DescribeSweepCoverage(sweep, "12 minutes", folderScoped: false));
    }

    [Fact]
    public void AnUnreadableFolder_StillRaisesFoldersSkipped_AndStillDegrades()
    {
        // The regression guard on the other side: the fix must not buy quiet by dropping a
        // real coverage hole.
        SweepInfo sweep = new SweepInfo
        {
            Performed = true,
            Scope = MailService.DefaultSweepScopeDescription,
            FoldersSwept = 3,
            FoldersSkipped = 1,
        };

        sweep.CoverageGaps = FreshMerge.DescribeCoverageGaps(sweep);

        Assert.NotNull(sweep.CoverageGaps);
        Assert.Contains(FreshMerge.GapFoldersSkipped, sweep.CoverageGaps!);
        Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyFreshness(sweep));
        Assert.Contains(
            MailService.DescribeSweepCoverage(sweep, "12 minutes", folderScoped: false),
            line => line.Contains("skipped 1 folder(s)", StringComparison.Ordinal));
    }

    [Fact]
    public void AStoreWithoutASingleDefaultFolder_IsStillReportedAsNoCoverage()
    {
        // Absence is not a gap, but a sweep that ended up covering NOTHING is - the answer
        // really was not checked against live Outlook, whatever the reason.
        SweepInfo sweep = new SweepInfo
        {
            Performed = true,
            Scope = MailService.DefaultSweepScopeDescription,
            FoldersSwept = 0,
            FoldersSkipped = 0,
            FoldersAbsent = OutlookComSession.DefaultSweepFolderKinds.Count,
        };

        sweep.CoverageGaps = FreshMerge.DescribeCoverageGaps(sweep);

        Assert.Contains(FreshMerge.GapNothingSwept, sweep.CoverageGaps!);
        Assert.Equal(FreshMerge.FreshnessPartial, FreshMerge.ClassifyFreshness(sweep));
    }

    // --- a model of the shipped default-folder loop, driven by the pure classification ---

    private static SweepAccounting SweepDefaultFolders(params Answer[] answers)
    {
        SweepAccounting tally = new SweepAccounting();
        foreach (Answer answer in answers)
        {
            switch (OutlookComSession.ClassifyDefaultFolder(answer.Folder, answer.Failure))
            {
                case OutlookComSession.DefaultFolderResolution.Absent:
                    tally.Absent++;
                    break;
                case OutlookComSession.DefaultFolderResolution.Unreadable:
                    tally.Skipped++;
                    break;
                default:
                    tally.Swept++;
                    break;
            }
        }

        return tally;
    }

    /// <summary>One store's answer for one default folder, as the COM call would deliver it.</summary>
    private sealed class Answer
    {
        private Answer(object? folder, Exception? failure)
        {
            Folder = folder;
            Failure = failure;
        }

        internal object? Folder { get; }

        internal Exception? Failure { get; }

        internal static Answer Resolved() => new Answer(new object(), null);

        /// <summary>The store has no such default folder: a null return, nothing thrown.</summary>
        internal static Answer Absent() => new Answer(null, null);

        /// <summary>The folder is there and the call failed.</summary>
        internal static Answer Failed() => new Answer(null, new COMException("folder will not open", RpcServerUnavailable));
    }

    private sealed class SweepAccounting
    {
        internal int Swept { get; set; }

        internal int Skipped { get; set; }

        internal int Absent { get; set; }
    }
}
