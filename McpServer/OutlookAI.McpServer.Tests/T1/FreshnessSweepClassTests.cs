using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using OutlookAI.ComHost.Supervision;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The freshness sweep may take ten minutes. Every other tool may not.
/// <para>
/// THE CHANGE THIS GUARDS (2026-08-24). The maintainer set the sweep budget to 600 s on the
/// standing rule that completeness outranks performance. That number does not fit under the
/// ordinary hang detector: <c>ComHostPolicy.TimeoutIndicatesUnresponsiveness</c> reads a
/// caller budget at or above its class deadline as evidence that OUTLOOK is deaf rather than
/// that the work was big, so a 600 s sweep judged against the 300 s ordinary deadline turns
/// every slow search on a large mailbox into a strike against the circuit breaker - two of
/// them and every COM request fails fast for the cooldown. The alternative on the table was
/// to raise the ordinary deadline instead, which buys the sweep its time by making
/// <c>read</c>, <c>move_mail</c>, <c>new_draft</c> and <c>list_folders</c> each wait eleven
/// minutes to discover a wedged Outlook, and twice that before failing fast.
/// </para>
/// <para>
/// So the sweep got a class of its own, exactly as the exhaustive scan already had one. THE
/// WHOLE VALUE OF THAT IS THE PART THAT DID NOT MOVE, and "did not move" is the hardest kind
/// of behaviour to notice losing: stretching a quick tool's detector changes nothing any
/// other test observes, produces no error, and is only visible on a machine whose Outlook
/// has actually wedged. That is what this file is for. It walks the WHOLE contract and
/// asserts the effective hang detector of every method on it, so a tool acquiring the long
/// one - by being added to the freshness set, by the sweep set swallowing the fallback, or
/// by the two class deadlines being collapsed into one number - fails here and names itself.
/// </para>
/// <para>
/// It is also why the classification lives in <see cref="ComOperationClasses"/> in Core
/// rather than in the private <c>ClassifyOperation</c> it replaced. That method sat on a
/// <c>DispatchProxy</c> which cannot be built without a supervisor and a child process, so
/// the single most load-bearing decision in the budget ladder was on a line no test in CI
/// could execute.
/// </para>
/// </summary>
public sealed class FreshnessSweepClassTests
{
    /// <summary>
    /// The effective hang detector for one contract call with no per-call override - i.e.
    /// exactly what <c>RemoteSessionProxy.Invoke</c> computes before it dispatches.
    /// </summary>
    private static long DetectorFor(string operation)
    {
        return ComHostPolicy.DeadlineFor(ComOperationClasses.ClassOf(operation), null);
    }

    private static IReadOnlyList<string> ContractOperations()
    {
        return typeof(IOutlookSession).GetMethods(BindingFlags.Public | BindingFlags.Instance)
            .Select(m => m.Name)
            .OrderBy(n => n, StringComparer.Ordinal)
            .ToArray();
    }

    /// <summary>
    /// THE DELIVERABLE. Every method on the contract, with the detector it actually gets -
    /// and the assertion that everything outside the three named long classes gets the
    /// ordinary one.
    /// <para>
    /// Read from the interface rather than from a list, so a contract method added later is
    /// covered the day it appears; and asserted by VALUE rather than by class name, because
    /// "still classified Operation" would stay green if the ordinary deadline itself were
    /// dragged up to hold the sweep - which is the change this whole design exists to avoid.
    /// </para>
    /// </summary>
    [Fact]
    public void EveryToolOtherThanTheLongOnes_KeepsTheOrdinaryHangDetector()
    {
        HashSet<string> longClasses = new HashSet<string>(
            ComOperationClasses.HealthProbeOperations
                .Concat(ComOperationClasses.ExhaustiveScanOperations)
                .Concat(ComOperationClasses.FreshnessSweepOperations),
            StringComparer.Ordinal);

        List<string> stretched = new List<string>();
        int ordinary = 0;

        foreach (string operation in ContractOperations())
        {
            if (longClasses.Contains(operation))
            {
                continue;
            }

            ordinary++;
            if (DetectorFor(operation) != ComOperationBudgets.OperationDeadlineMs)
            {
                stretched.Add($"{operation} -> {DetectorFor(operation)} ms ({ComOperationClasses.ClassOf(operation)})");
            }
        }

        Assert.True(
            stretched.Count == 0,
            $"these operations no longer reclaim a wedged Outlook in {ComOperationBudgets.OperationDeadlineMs} ms: "
            + string.Join(", ", stretched)
            + ". Every tool a user waits on keeps the ordinary hang detector - that is the entire reason the sweep "
            + "and the exhaustive scan have classes of their own instead of a bigger shared number.");

        // The loop really ran over the contract. A classification that swallowed everything
        // into the long sets would otherwise satisfy the assertion above by iterating zero
        // times, which is the way an exhaustive check quietly stops being one.
        Assert.True(
            ordinary >= ContractOperations().Count - longClasses.Count,
            "the ordinary-class count must account for every contract method outside the long classes");
        Assert.True(ordinary > 15, $"only {ordinary} contract operations were checked, which cannot be the whole contract");
    }

    /// <summary>
    /// The tools named in the argument for keeping the ordinary detector, pinned by name so
    /// the intent survives any future reshuffling of the sets. These are what a user is
    /// sitting in front of when Outlook wedges.
    /// </summary>
    [Theory]
    [InlineData(nameof(IOutlookSession.TryReadItem))]
    [InlineData(nameof(IOutlookSession.TryMoveItemToPath))]
    [InlineData(nameof(IOutlookSession.TryMoveItemToFolderId))]
    [InlineData(nameof(IOutlookSession.TryCreateNewDraft))]
    [InlineData(nameof(IOutlookSession.TryCreateDerivedDraft))]
    [InlineData(nameof(IOutlookSession.TryUpdateDraft))]
    [InlineData(nameof(IOutlookSession.TryDiscardDraft))]
    [InlineData(nameof(IOutlookSession.TrySendDraft))]
    [InlineData(nameof(IOutlookSession.ListFolders))]
    [InlineData(nameof(IOutlookSession.GetAccounts))]
    [InlineData(nameof(IOutlookSession.TrySaveAttachment))]
    public void AQuickTool_IsNeverDispatchedOnTheSweepsDetector(string operation)
    {
        Assert.Equal(ComHostOperationClass.Operation, ComOperationClasses.ClassOf(operation));
        Assert.Equal(ComOperationBudgets.OperationDeadlineMs, (int)DetectorFor(operation));

        // Stated as the inequality rather than only as the value, so the failure message
        // says what was lost: this tool would wait the sweep's detector for a wedge.
        Assert.True(
            DetectorFor(operation) < ComOperationBudgets.FreshnessSweepDeadlineMs,
            $"{operation} must not wait the freshness class's {ComOperationBudgets.FreshnessSweepDeadlineMs} ms before a "
            + "wedged Outlook is reclaimed");
        Assert.True(
            DetectorFor(operation) < ComOperationBudgets.ExhaustiveScanDeadlineMs,
            $"{operation} must not wait the exhaustive class's {ComOperationBudgets.ExhaustiveScanDeadlineMs} ms either");

        // And its own budget-free expiry still counts as a hang, which is what the short
        // detector is FOR. A tool that stopped counting would fail fast never.
        Assert.True(ComHostPolicy.TimeoutIndicatesUnresponsiveness(ComOperationClasses.ClassOf(operation), null));
    }

    /// <summary>
    /// The two operations in the freshness class, and why there are two: the sweep and
    /// <c>thread</c>'s conversation walk share <c>MailService.SweepBudgetMs</c>, so they have
    /// to share the threshold that budget is judged against.
    /// </summary>
    [Fact]
    public void TheSweepAndTheThreadWalk_AreTheFreshnessClass()
    {
        Assert.Equal(
            ComHostOperationClass.FreshnessSweep,
            ComOperationClasses.ClassOf(nameof(IOutlookSession.SweepFoldersNewerThan)));
        Assert.Equal(
            ComHostOperationClass.FreshnessSweep,
            ComOperationClasses.ClassOf(nameof(IOutlookSession.TryGetConversationItems)));

        // They share the budget, which is why they must share the class.
        Assert.Equal(MailService.SweepBudgetMs, MailService.ThreadWalkBudgetMs);

        // Neither of their ordinary expiries is evidence of a wedge...
        Assert.False(ComHostPolicy.TimeoutIndicatesUnresponsiveness(
            ComOperationClasses.ClassOf(nameof(IOutlookSession.SweepFoldersNewerThan)), MailService.SweepBudgetMs));
        Assert.False(ComHostPolicy.TimeoutIndicatesUnresponsiveness(
            ComOperationClasses.ClassOf(nameof(IOutlookSession.TryGetConversationItems)), MailService.ThreadWalkBudgetMs));

        // ...while a freshness call that declares NO budget still trips the detector, so the
        // class buys the sweep room without disarming the class.
        Assert.True(ComHostPolicy.TimeoutIndicatesUnresponsiveness(ComHostOperationClass.FreshnessSweep, null));
        Assert.True(ComHostPolicy.TimeoutIndicatesUnresponsiveness(
            ComHostOperationClass.FreshnessSweep, ComOperationBudgets.FreshnessSweepDeadlineMs));
    }

    /// <summary>
    /// The rule is applied PER CLASS, and the same budget gets opposite answers under the
    /// two classes. This is the assertion that fails if the classes stop being distinguished
    /// - whether by collapsing the switch, by giving the freshness class the ordinary
    /// deadline, or by comparing every class against one number.
    /// </summary>
    [Fact]
    public void TheUnresponsivenessThreshold_DistinguishesTheClasses()
    {
        const int budget = ComOperationBudgets.FreshnessSweepBudgetMs;

        Assert.False(ComHostPolicy.TimeoutIndicatesUnresponsiveness(ComHostOperationClass.FreshnessSweep, budget));
        Assert.True(ComHostPolicy.TimeoutIndicatesUnresponsiveness(ComHostOperationClass.Operation, budget));

        // The classes resolve to genuinely different deadlines. Equal deadlines make the
        // classification decorative: every assertion about "which class" would still pass
        // while the behaviour it stands for had gone.
        Assert.NotEqual(
            ComHostPolicy.DeadlineFor(ComHostOperationClass.Operation, null),
            ComHostPolicy.DeadlineFor(ComHostOperationClass.FreshnessSweep, null));
        Assert.NotEqual(
            ComHostPolicy.DeadlineFor(ComHostOperationClass.ExhaustiveScan, null),
            ComHostPolicy.DeadlineFor(ComHostOperationClass.FreshnessSweep, null));

        // Five classes, five distinct deadlines - the mechanism is a table, not a synonym.
        long[] deadlines = Enum.GetValues<ComHostOperationClass>()
            .Select(c => ComHostPolicy.DeadlineFor(c, null))
            .ToArray();
        Assert.Equal(deadlines.Length, deadlines.Distinct().Count());
    }

    /// <summary>
    /// The freshness ladder, every rung, as inequalities rather than as five numbers that
    /// happen to agree today.
    /// <para>
    /// The equality cases are the defects, not the edges. An inner budget EQUAL to its outer
    /// one can never degrade gracefully - the walk stops only once elapsed has passed it and
    /// then still has to serialize its answer, while the watchdog fires at <c>&gt;=</c> -
    /// and a caller budget EQUAL to its class deadline is read as a hang detector, which is
    /// the breaker outage. Both were real: the first cost the exhaustive scan its documented
    /// partial-results answer, the second is what blocked this change for a day.
    /// </para>
    /// </summary>
    [Fact]
    public void TheFreshnessLadder_IsOrderedAndDerived()
    {
        Assert.True(
            MailService.SweepWorkBudgetMs < MailService.SweepBudgetMs,
            $"the sweep's inner budget ({MailService.SweepWorkBudgetMs} ms) must be strictly inside the gateway budget it "
            + $"runs under ({MailService.SweepBudgetMs} ms), or a long sweep is a timeout and a host kill rather than a "
            + "partial answer");
        Assert.True(
            MailService.SweepBudgetMs < ComOperationBudgets.FreshnessSweepDeadlineMs,
            $"the sweep's declared budget ({MailService.SweepBudgetMs} ms) must be strictly inside its class deadline "
            + $"({ComOperationBudgets.FreshnessSweepDeadlineMs} ms); at or above it, every ordinary slow sweep counts "
            + "toward the circuit breaker");
        Assert.True(
            ComOperationBudgets.OperationDeadlineMs < ComOperationBudgets.FreshnessSweepDeadlineMs,
            $"the freshness class ({ComOperationBudgets.FreshnessSweepDeadlineMs} ms) exists precisely because it is "
            + $"longer than the ordinary operation deadline ({ComOperationBudgets.OperationDeadlineMs} ms); equal or "
            + "shorter means the class buys nothing and should not exist");

        // Derived, not written twice. The return trip is reserved out of the outer budget,
        // exactly as it is for the exhaustive scan.
        Assert.Equal(
            MailService.SweepBudgetMs - ComOperationBudgets.ResultReturnHeadroomMs,
            MailService.SweepWorkBudgetMs);
        Assert.Equal(ComOperationBudgets.FreshnessSweepBudgetMs, MailService.SweepBudgetMs);
        Assert.Equal(ComOperationBudgets.FreshnessSweepWorkBudgetMs, MailService.SweepWorkBudgetMs);
        Assert.True(ComOperationBudgets.ResultReturnHeadroomMs > 0, "the return-trip headroom is the whole mechanism");

        // THE DERIVATION OF THE CLASS DEADLINE, pinned from outside because it cannot be
        // written as a const: ComOperationBudgets lives in Core/Com and may not reference the
        // service layer, so the sum of "one whole composed search plus the return trip" has
        // no compilation that can state it. Same idiom as SweepBodyCapTests. Without this the
        // 675 000 is a literal nobody can check, and narrowing the sweep budget later would
        // silently leave the class deadline oversized.
        Assert.Equal(
            MailService.SearchBudgetMs + ComOperationBudgets.ResultReturnHeadroomMs,
            ComOperationBudgets.FreshnessSweepDeadlineMs);
    }

    /// <summary>
    /// The exhaustive scan is untouched by all of this - its class, its deadline and its
    /// relationship to the ordinary one are exactly what they were.
    /// <para>
    /// Deliberate, and stated here as well as in <c>BudgetCompositionTests</c>: 615 s is the
    /// subject of a measurement that has never been run on either machine, and moving it to
    /// make room for the sweep would have destroyed what that measurement is meant to
    /// inform. The freshness class is what made moving it unnecessary.
    /// </para>
    /// </summary>
    [Fact]
    public void TheExhaustiveScanClass_IsUnchangedByTheFreshnessSplit()
    {
        Assert.Equal(615_000, ComOperationBudgets.ExhaustiveScanDeadlineMs);
        Assert.Equal(
            ComHostOperationClass.ExhaustiveScan,
            ComOperationClasses.ClassOf(nameof(IOutlookSession.ExhaustiveScan)));
        Assert.Equal(
            ComOperationBudgets.ExhaustiveScanDeadlineMs,
            (int)DetectorFor(nameof(IOutlookSession.ExhaustiveScan)));
        Assert.True(ComOperationBudgets.ExhaustiveScanDeadlineMs > ComOperationBudgets.OperationDeadlineMs);
    }

    /// <summary>
    /// The health probe keeps the shortest detector of all, which is the one place a budget
    /// expiring IS the answer.
    /// </summary>
    [Fact]
    public void TheHealthProbe_KeepsTheShortestDetector()
    {
        Assert.Equal(
            ComHostOperationClass.HealthProbe,
            ComOperationClasses.ClassOf(nameof(IOutlookSession.GetProfileName)));
        Assert.Equal(
            ComOperationBudgets.HealthProbeDeadlineMs,
            (int)DetectorFor(nameof(IOutlookSession.GetProfileName)));
        Assert.True(ComOperationBudgets.HealthProbeDeadlineMs < ComOperationBudgets.OperationDeadlineMs);
    }

    /// <summary>
    /// Fail-closed, in the direction that costs a budget rather than a hang detector: an
    /// unclassified name gets the SHORT deadline. The opposite default would mean a contract
    /// method added without a decision silently inherits ten minutes.
    /// </summary>
    [Theory]
    [InlineData("SomeMethodAddedLater")]
    [InlineData("")]
    [InlineData(null)]
    public void AnOperationNobodyClassified_GetsTheOrdinaryDetector(string? operation)
    {
        Assert.Equal(ComHostOperationClass.Operation, ComOperationClasses.ClassOf(operation));
        Assert.Equal(
            ComOperationBudgets.OperationDeadlineMs,
            (int)ComHostPolicy.DeadlineFor(ComOperationClasses.ClassOf(operation), null));
    }

    /// <summary>
    /// The three long-class sets are disjoint and every name in them is really on the
    /// contract. Written with <c>nameof</c>, so a rename breaks the build there - but a
    /// DELETION would leave a set naming a method that no longer exists, and the fallback is
    /// silent, so the operation would quietly change class.
    /// </summary>
    [Fact]
    public void TheClassTable_NamesOnlyRealContractOperations_AndNeverTwice()
    {
        HashSet<string> contract = new HashSet<string>(ContractOperations(), StringComparer.Ordinal);

        string[][] sets =
        {
            ComOperationClasses.HealthProbeOperations.ToArray(),
            ComOperationClasses.ExhaustiveScanOperations.ToArray(),
            ComOperationClasses.FreshnessSweepOperations.ToArray(),
        };

        foreach (string[] set in sets)
        {
            Assert.NotEmpty(set);
            Assert.Empty(set.Except(contract, StringComparer.Ordinal));
        }

        int named = sets.Sum(s => s.Length);
        int distinct = sets.SelectMany(s => s).Distinct(StringComparer.Ordinal).Count();
        Assert.Equal(named, distinct);
    }
}
