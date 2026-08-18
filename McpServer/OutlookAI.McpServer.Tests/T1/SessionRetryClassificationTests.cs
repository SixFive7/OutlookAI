using System.Reflection;
using System.Runtime.InteropServices;

using OutlookAI.ComHost.Host;
using OutlookAI.Core.Com;

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
