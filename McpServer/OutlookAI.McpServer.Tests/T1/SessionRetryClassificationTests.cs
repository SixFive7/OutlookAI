using System.Reflection;
using System.Runtime.InteropServices;

using OutlookAI.ComHost.Host;
using OutlookAI.ComHost.Supervision;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The COM host may re-run a failed call against a rebuilt session. This pins WHICH calls,
/// and that the answer is decided by classification rather than by luck.
/// <para>
/// The hazard is specific. <c>ComGateway.Run</c>'s one-shot rebuild catches the
/// RPC_E_DISCONNECTED family, and that family contains <c>RPC_S_CALL_FAILED</c>
/// (0x800706BE) - the HRESULT whose documented meaning is that the call MAY OR MAY NOT have
/// executed. A re-run is therefore a possible second execution. On 2026-08-18 the wrapper
/// fix (<c>GatewayRoutingProxy</c>) made that rebuild reachable inside the COM host for the
/// first time, and <c>TrySendDraft</c> is on the same contract, reached the same way. Its
/// confirm token is consumed on the PARENT before the call is sent, so on a re-run none of
/// the send guards - two-step token, content hash, identity check - is still standing. A
/// duplicate send cannot be recalled.
/// </para>
/// <para>
/// Two levels are pinned, because either alone would rot. The TABLE must cover the contract
/// exactly (so a method added later cannot slip through unclassified), and the ROUTING
/// PROXY must actually ask for what the table says (so a correct table cannot sit unread).
/// </para>
/// </summary>
public sealed class SessionRetryClassificationTests
{
    [Fact]
    public void EveryOperationOnTheContract_IsClassifiedExactlyOnce()
    {
        // THE GUARD. Adding a method to IOutlookSession fails here until someone decides
        // whether re-running it is safe. Read from the interface, never from a copy of it.
        HashSet<string> contract = ContractOperations();
        HashSet<string> classified = new(
            ComSessionOperations.ReadOnlyOperations.Concat(ComSessionOperations.MutatingOperations),
            StringComparer.Ordinal);

        Assert.Empty(contract.Except(classified, StringComparer.Ordinal));
        Assert.Empty(classified.Except(contract, StringComparer.Ordinal));
        Assert.Empty(ComSessionOperations.ReadOnlyOperations.Intersect(
            ComSessionOperations.MutatingOperations, StringComparer.Ordinal));

        // Counted, so a rename that swaps one name for another cannot pass unnoticed.
        Assert.Equal(contract.Count, ComSessionOperations.ReadOnlyOperations.Count
            + ComSessionOperations.MutatingOperations.Count);
    }

    [Fact]
    public void TheSendPath_IsNotRetryable_AndTheReadThatPrecedesItIs()
    {
        // The operation this whole classification exists for, named on its own so the intent
        // survives any future reshuffling of the sets.
        Assert.False(ComSessionOperations.IsRetryable(nameof(IOutlookSession.TrySendDraft)));

        // Its sibling reads the draft and hashes it. Nothing is consumed, so a rebuilt
        // session may be asked the same question again.
        Assert.True(ComSessionOperations.IsRetryable(nameof(IOutlookSession.TryGetSendableDraftState)));
    }

    [Theory]
    [InlineData(nameof(IOutlookSession.TryCreateNewDraft))]
    [InlineData(nameof(IOutlookSession.TryCreateDerivedDraft))]
    [InlineData(nameof(IOutlookSession.TryUpdateDraft))]
    [InlineData(nameof(IOutlookSession.TryDiscardDraft))]
    [InlineData(nameof(IOutlookSession.TryMoveItemToPath))]
    [InlineData(nameof(IOutlookSession.TryMoveItemToFolderId))]
    [InlineData(nameof(IOutlookSession.TrySaveAttachment))]
    public void EveryOperationThatLeavesSomethingBehind_IsNotRetryable(string operation)
    {
        // Not just send: a re-run of any of these leaves a second draft, a second move
        // attempt against an item that is no longer where it was, or a second file on disk.
        Assert.False(ComSessionOperations.IsRetryable(operation));
    }

    /// <summary>
    /// The classification reaches the CALLER too, not only the retry decision.
    /// <para>
    /// When one request wedges and its COM host is reclaimed, every innocent sibling dies
    /// with the connection and is told why. That message used to end "This request itself
    /// was not at fault; retry it." for all of them alike - correct for a read, and
    /// actively dangerous for a mutation: the child was terminated at an unknown point, so
    /// a killed <c>TrySendDraft</c> may already have submitted the mail and a killed
    /// <c>TryUpdateDraft</c> may have applied part of its sequence. "Retry it" is then
    /// advice to send twice.
    /// </para>
    /// </summary>
    [Theory]
    [InlineData(nameof(IOutlookSession.SweepFoldersNewerThan), true)]
    [InlineData(nameof(IOutlookSession.ExhaustiveScan), true)]
    [InlineData(nameof(IOutlookSession.TryReadItem), true)]
    [InlineData(nameof(IOutlookSession.TrySendDraft), false)]
    [InlineData(nameof(IOutlookSession.TryUpdateDraft), false)]
    [InlineData(nameof(IOutlookSession.TryMoveItemToPath), false)]
    [InlineData("SomeMethodAddedLater", false)]
    public void AnInterruptedRequest_IsToldWhetherItIsSafeToRetry(string operation, bool retryable)
    {
        foreach (string? cause in new[] { null, "'TryUpdateDraft' exceeded its 300000 ms budget." })
        {
            string message = ComHostSupervisor.DescribeInterruption(operation, cause);

            Assert.Contains(operation, message, System.StringComparison.Ordinal);
            if (retryable)
            {
                Assert.Contains("retrying it is safe", message, System.StringComparison.Ordinal);
                Assert.DoesNotContain("UNKNOWN", message, System.StringComparison.Ordinal);
            }
            else
            {
                Assert.Contains("UNKNOWN", message, System.StringComparison.Ordinal);
                Assert.DoesNotContain("retrying it is safe", message, System.StringComparison.Ordinal);
            }
        }
    }

    [Fact]
    public void AnOperationNobodyClassified_IsTreatedAsMutating()
    {
        // Fail-closed. If the guard above is ever bypassed, the cost is a read that stopped
        // recovering from a disconnect - never a write that started duplicating itself.
        Assert.False(ComSessionOperations.IsRetryable("SomeMethodAddedLater"));
        Assert.False(ComSessionOperations.IsRetryable(null));
        Assert.False(ComSessionOperations.IsClassified("SomeMethodAddedLater"));
    }

    [Fact]
    public void TheRoutingProxy_AsksForTheRebuildOnEveryRead_AndOnNoWrite()
    {
        // The table read at the one place that can act on it. Every contract method is
        // driven through the real proxy against a session that fails the way a lost Outlook
        // does; what is asserted is the recovery the proxy requested BEFORE the call ran.
        IOutlookSession failing = ComHostFaultInjection.FaultingSession.Create("com");
        RecordingGateway gateway = new RecordingGateway(failing);
        IOutlookSession proxy = GatewayRoutingProxy.Create(gateway);

        int reads = 0;
        int writes = 0;
        foreach (MethodInfo method in typeof(IOutlookSession).GetMethods())
        {
            gateway.Forget();

            TargetInvocationException thrown = Assert.Throws<TargetInvocationException>(
                () => method.Invoke(proxy, DefaultArgumentsFor(method)));
            _ = Assert.IsType<COMException>(thrown.InnerException);

            bool retryable = ComSessionOperations.IsRetryable(method.Name);
            Assert.Equal(
                retryable ? ComSessionRecovery.RebuildOnce : ComSessionRecovery.None,
                gateway.RequestedRecovery);

            if (retryable)
            {
                reads++;
            }
            else
            {
                writes++;
            }
        }

        // Both halves were actually exercised - a classification that silently collapsed to
        // one answer would otherwise satisfy every assertion above.
        Assert.Equal(ComSessionOperations.ReadOnlyOperations.Count, reads);
        Assert.Equal(ComSessionOperations.MutatingOperations.Count, writes);
    }

    // ------------------------------------------- the OTHER retry: store by store, in one session

    // A different question from the one above, and it had drifted in both directions at
    // once. This is not the disconnect rebuild; it is the loop that answers "the caller gave
    // a bare EntryID and we do not know which store it lives in". reply/replyall/forward
    // retried on ANY failure, so a compose or Save error fanned out across every store -
    // and creating a derived draft is not idempotent, so a run that still ended in failure
    // could leave one orphaned draft per store, none of whose ids the caller ever learns.
    // update_draft and discard_draft asked for the "ItemNotFound" token, which their COM
    // layer never set - so their loop was dead code and a draft in a non-default store came
    // back as an opaque "COMException 0x..." instead of being found.

    [Fact]
    public void ABareEntryIdThatWasNotFound_IsLookedForInTheOtherStores()
    {
        Assert.True(MailService.ShouldSearchOtherStores(
            storeId: null, succeeded: false, error: MailService.ItemNotFoundToken));
    }

    [Fact]
    public void AFailureWithAnyOtherCause_StopsAtTheFirstAttempt()
    {
        // The tightening. Every one of these happens AFTER the item was opened, so a second
        // attempt would repeat work rather than find something - and for a derived draft
        // that work leaves a mail behind.
        foreach (string? other in new[] { "AlreadySent", "NotInDraftsFolder", "COMException 0x80040111", "NoInspector", null })
        {
            Assert.False(MailService.ShouldSearchOtherStores(storeId: null, succeeded: false, error: other));
            Assert.False(MailService.KeepSearchingStores(succeeded: false, error: other));
        }
    }

    [Fact]
    public void AKnownStore_IsNeverSecondGuessed()
    {
        // The caller (or the hit cache) already said where the item lives. Not finding it
        // there is an answer, not a reason to go looking elsewhere.
        Assert.False(MailService.ShouldSearchOtherStores(
            storeId: "store-id", succeeded: false, error: MailService.ItemNotFoundToken));
    }

    [Fact]
    public void ASuccessEndsTheSearch_Immediately()
    {
        Assert.False(MailService.ShouldSearchOtherStores(
            storeId: null, succeeded: true, error: null));
        Assert.False(MailService.KeepSearchingStores(succeeded: true, error: null));
        Assert.True(MailService.KeepSearchingStores(succeeded: false, error: MailService.ItemNotFoundToken));
    }

    [Fact]
    public void TheTokenIsSpeltExactlyAsTheComLayerSetsIt()
    {
        // The two halves of this rule live in different assemblies and are joined by a
        // string. That is how update_draft's retry came to be unreachable for as long as it
        // was: nothing failed, nothing warned, the loop simply never ran.
        Assert.Equal("ItemNotFound", MailService.ItemNotFoundToken);
    }

    private static HashSet<string> ContractOperations()
    {
        return new HashSet<string>(
            typeof(IOutlookSession).GetMethods().Select(m => m.Name),
            StringComparer.Ordinal);
    }

    /// <summary>
    /// Placeholder arguments for a reflective call. Nothing reads them - the stand-in
    /// session fails before it looks at anything - so the values only have to be legal.
    /// </summary>
    private static object?[] DefaultArgumentsFor(MethodInfo method)
    {
        ParameterInfo[] parameters = method.GetParameters();
        object?[] args = new object?[parameters.Length];
        for (int i = 0; i < parameters.Length; i++)
        {
            Type type = parameters[i].ParameterType;
            if (type.IsByRef)
            {
                type = type.GetElementType()!;
            }

            args[i] = type.IsValueType ? Activator.CreateInstance(type) : null;
        }

        return args;
    }

    /// <summary>
    /// Stands where <c>ComGateway</c> stands and remembers what was asked of it, so the
    /// decision can be observed without an Outlook to disconnect from.
    /// </summary>
    private sealed class RecordingGateway : IComGateway
    {
        private readonly IOutlookSession _session;

        internal RecordingGateway(IOutlookSession session)
        {
            _session = session;
        }

        public event Action? OutlookGone
        {
            add { }
            remove { }
        }

        internal ComSessionRecovery? RequestedRecovery { get; private set; }

        public bool IsConnected => true;

        public bool? QuitSinkActive => null;

        public bool ProbeConnected() => true;

        internal void Forget() => RequestedRecovery = null;

        public T Run<T>(Func<IOutlookSession, T> operation) => Run(operation, ComSessionRecovery.None);

        public T Run<T>(Func<IOutlookSession, T> operation, ComSessionRecovery recovery)
        {
            RequestedRecovery = recovery;
            return operation(_session);
        }

        public T Run<T>(Func<IOutlookSession, T> operation, int budgetMilliseconds, bool allowConnectFloor = false)
        {
            return Run(operation);
        }

        public ComHostDiagnostics GetDiagnostics() => new ComHostDiagnostics("in-process", "ready");

        public void Dispose()
        {
        }
    }
}
