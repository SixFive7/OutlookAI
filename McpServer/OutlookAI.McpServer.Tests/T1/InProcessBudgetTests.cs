using System;
using System.Runtime.InteropServices;

using OutlookAI.ComHost.Host;
using OutlookAI.Core.Com;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The in-process gateway's budget is real.
/// <para>
/// THE DEFECT THIS PINS. <c>ComGateway.Run(operation, budgetMilliseconds, allowConnectFloor)</c>
/// was <c>{ return Run(operation); }</c> - it accepted a budget and discarded it. That is
/// defensible inside the COM host child, where the parent's watchdog is the real bound. It
/// was not defensible for <c>MailService.CreateDefault()</c>, which uses the same gateway,
/// and on which EVERY live (T2) fixture is built: the whole live tier therefore ran with no
/// budget, no aggregate, no breaker and no hang detector, so everything sized against those
/// budgets was unverified there by construction and a hung live run had nothing to stop it.
/// </para>
/// <para>
/// What the in-process side can enforce is the AGGREGATE, not the single call: a blocked
/// outbound COM call is not cancellable and the caller cannot kill itself. So the proxy
/// refuses to START a call once the budget is spent, which is exactly the half that bounds
/// a multi-call service operation.
/// </para>
/// </summary>
public sealed class InProcessBudgetTests
{
    [Theory]
    // Untouched budget, and the ordinary shrinking case.
    [InlineData(180_000L, 0L, 180_000L)]
    [InlineData(180_000L, 20_000L, 160_000L)]
    // Spent exactly, and past spent: clamped at zero rather than going negative.
    [InlineData(180_000L, 180_000L, 0L)]
    [InlineData(180_000L, 200_000L, 0L)]
    public void RemainingBudget_ShrinksWithElapsedAndNeverGoesNegative(long budget, long elapsed, long expected)
    {
        Assert.Equal(expected, BudgetedSessionProxy.RemainingMilliseconds(budget, elapsed));
    }

    [Theory]
    // Comfortably inside.
    [InlineData(180_000L, 0L, true)]
    // Exactly at the dispatch floor: still dispatched, so the boundary is inclusive.
    [InlineData(180_000L, 180_000L - BudgetedSessionProxy.MinimumRemainingMilliseconds, true)]
    // One millisecond below it: reported as spent instead of dispatched, because a
    // sub-second remainder cannot finish anything and a bare timeout is a worse answer.
    [InlineData(180_000L, 180_000L - BudgetedSessionProxy.MinimumRemainingMilliseconds + 1, false)]
    [InlineData(180_000L, 180_000L, false)]
    public void DispatchFloor_RefusesARemainderTooSmallToFinishAnything(long budget, long elapsed, bool expected)
    {
        Assert.Equal(expected, BudgetedSessionProxy.CanDispatch(budget, elapsed));
    }

    [Fact]
    public void NoBudget_IsNotWrapped_BecauseNoBudgetMustNeverMeanNoCalls()
    {
        IOutlookSession inner = ComHostFaultInjection.FaultingSession.Create("com");
        Assert.Same(inner, BudgetedSessionProxy.Wrap(inner, 0));
        Assert.Same(inner, BudgetedSessionProxy.Wrap(inner, -1));
    }

    [Fact]
    public void AnUnspentBudget_LetsTheCallThrough_AndKeepsTheFailuresTypeAndMessage()
    {
        // The reflective hop is the hazard: this repository has already had a reflective
        // layer flatten every deliberate error into "Exception has been thrown by the target
        // of an invocation", which broke both the tool layer's advice (it branches on
        // exception TYPE) and ComGateway's disconnect rebuild (it branches on COMException
        // HRESULTs). Adding another proxy must not reintroduce that.
        IOutlookSession budgeted = BudgetedSessionProxy.Wrap(
            ComHostFaultInjection.FaultingSession.Create("com"), 180_000);

        COMException failure = Assert.Throws<COMException>(() => budgeted.GetProfileName());
        Assert.Equal(ComHostFaultInjection.SessionComMessage, failure.Message);
        Assert.Equal(ComHostFaultInjection.SessionComHResult, failure.HResult);
        Assert.DoesNotContain("target of an invocation", failure.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ASpentBudget_RefusesTheCallInsteadOfStartingIt()
    {
        // A budget below the dispatch floor is spent the moment it is created, which makes
        // this deterministic rather than a race against a stopwatch. The inner session
        // throws on every call, so reaching it at all would surface as a COMException -
        // getting a TimeoutException instead proves the call was never started.
        IOutlookSession budgeted = BudgetedSessionProxy.Wrap(
            ComHostFaultInjection.FaultingSession.Create("com"), 1);

        TimeoutException expired = Assert.Throws<TimeoutException>(() => budgeted.GetProfileName());

        // The message names the operation and the budget, because the caller's remedy is to
        // narrow the request rather than to retry it unchanged.
        Assert.Contains(nameof(IOutlookSession.GetProfileName), expired.Message, StringComparison.Ordinal);
        Assert.Contains("1 ms", expired.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TheRefusalMessage_IsBuiltFromItsInputs()
    {
        Assert.Equal(
            BudgetedSessionProxy.BudgetExhaustedMessage("SweepFoldersNewerThan", 165_000),
            BudgetedSessionProxy.BudgetExhaustedMessage("SweepFoldersNewerThan", 165_000));
        Assert.Contains(
            "SweepFoldersNewerThan",
            BudgetedSessionProxy.BudgetExhaustedMessage("SweepFoldersNewerThan", 165_000),
            StringComparison.Ordinal);
    }
}
