using System.IO.Pipes;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

using OutlookAI.ComHost.Host;
using OutlookAI.ComHost.Protocol;
using OutlookAI.ComHost.Supervision;
using OutlookAI.Core.Com;
using OutlookAI.Core.Services;

using Xunit;

namespace OutlookAI.McpServer.Tests.T1;

/// <summary>
/// The per-operation audit of the COM-host boundary: what a deliberate, actionable failure
/// looks like by the time it reaches the caller, for EVERY method on
/// <see cref="IOutlookSession"/> rather than for the one method a probe happened to try.
/// <para>
/// The defect this exists for was found on 2026-08-18 and had been live since the process
/// split. Two reflective hops sit in series - the routing proxy's <c>targetMethod.Invoke</c>
/// and the host server's <c>method.Invoke</c> - so a failure from the session arrived
/// wrapped twice, the server peeled exactly one layer, and what went on the wire was the
/// INNER wrapper: type "TargetInvocationException", message "Exception has been thrown by
/// the target of an invocation.", no HRESULT, no reason. Every deliberate error from all 26
/// contract methods read the same, and an agent that cannot tell "no such folder" from
/// "something went wrong" retries blindly instead of calling <c>list_folders</c>.
/// </para>
/// <para>
/// It went unnoticed because nothing had looked at the boundary as a whole. That is what
/// this file is: the enumeration is taken from the interface by reflection, so an operation
/// added later is audited the day it is added, and the count is asserted so coverage cannot
/// quietly shrink.
/// </para>
/// <para>
/// The whole child side runs for real - routing proxy, host server, framing, JSON - over an
/// in-process named pipe, with the parent's own <see cref="ComHostErrorMapper"/> rebuilding
/// what comes back. Faking any of it would remove the point: every one of those layers is
/// somewhere an error could be flattened, and the one that did flatten it looked correct in
/// isolation. No Outlook, no process spawn, no mailbox.
/// </para>
/// </summary>
public sealed class ComHostErrorFidelityTests
{
    /// <summary>
    /// The contract size at the time of the audit. Asserted as a floor, not an equality:
    /// growing the contract is fine, auditing fewer operations than were audited here is
    /// not.
    /// </summary>
    private const int AuditedOperationCount = 26;

    /// <summary>Bounds every wait, so a boundary bug fails this suite instead of wedging CI.</summary>
    private static readonly TimeSpan Patience = TimeSpan.FromSeconds(30);

    [Fact]
    public async Task EveryOperationOnTheContract_DeliversItsOwnFailureIntact()
    {
        // The audit. One specific, actionable message per operation - modelled on the real
        // one that went missing - each naming its own operation, so a boundary that mixed
        // two up, or answered with a constant, fails here rather than reading plausibly.
        MethodInfo[] contract = typeof(IOutlookSession).GetMethods();
        await using Boundary boundary = await Boundary.StartAsync(
            ScriptedSession.Create(operation => new InvalidOperationException(FolderMessageFor(operation))));

        int audited = 0;
        foreach (MethodInfo method in contract)
        {
            Exception surfaced = await boundary.FailureOfAsync(method.Name);

            InvalidOperationException typed = Assert.IsType<InvalidOperationException>(surfaced);
            Assert.Equal(FolderMessageFor(method.Name), typed.Message);
            audited++;
        }

        Assert.Equal(contract.Length, audited);
        Assert.True(
            audited >= AuditedOperationCount,
            "The audit covered " + audited + " operations; " + AuditedOperationCount + " were covered when it was written.");
    }

    [Fact]
    public async Task EveryOperationThatReportsFailureByToken_DeliversTheTokenItself()
    {
        // The OTHER channel, and the one the cross-store retries live on. Most of this
        // contract reports failure through `out string? error` and returns null rather than
        // throwing, so the exception path above says nothing about it: the token travels as
        // a by-ref output inside a SUCCESSFUL frame. It is what decides whether a bare
        // EntryID is looked for in the other stores, so losing it is a behaviour change with
        // no error anywhere to notice.
        MethodInfo[] byToken = typeof(IOutlookSession)
            .GetMethods()
            .Where(m => m.GetParameters().Any(IsErrorOutput))
            .ToArray();

        Assert.NotEmpty(byToken);

        await using Boundary boundary = await Boundary.StartAsync(ScriptedSession.Create(_ => null));

        foreach (MethodInfo method in byToken)
        {
            ComHostResponse response = await boundary.CallAsync(method.Name);

            Assert.True(response.Ok, method.Name + " reported a failure it was not asked to report.");
            Assert.NotNull(response.Outputs);
            Assert.True(
                response.Outputs!.TryGetValue("error", out JsonElement carried),
                method.Name + " lost its 'error' output crossing the boundary.");
            Assert.Equal(ComErrorTokens.ItemNotFound, carried.GetString());
        }
    }

    [Theory]
    [MemberData(nameof(ModelledFailures))]
    public async Task AModelledFailure_ArrivesAsItself(string label)
    {
        // Type identity is not cosmetic here. OutlookTools.GuardAsync branches on the
        // exception TYPE to choose the payload an agent sees, and ComGateway keys its
        // disconnect rebuild on COMException HRESULTs - both of which were unreachable for
        // as long as everything arrived as a reflection wrapper.
        //
        // The label rather than the exception itself is the theory datum: xunit cannot
        // serialize an Exception, and passing one collapses fourteen named cases into a
        // single opaque test that names nothing when it fails.
        Exception raised = Raise(label);
        await using Boundary boundary = await Boundary.StartAsync(ScriptedSession.Create(_ => raised));

        Exception surfaced = await boundary.FailureOfAsync(nameof(IOutlookSession.GetProfileName));

        Assert.Equal(raised.GetType(), surfaced.GetType());
        Assert.Equal(raised.Message, surfaced.Message);
    }

    [Fact]
    public async Task AComFailure_KeepsItsHResult()
    {
        // The HRESULT is the only thing that distinguishes "Outlook disconnected, rebuild
        // and re-run the read" from "Outlook refused this". It is written on the wire only
        // when the error IS a COMException by then, which - before the unwrap fix - it never
        // was.
        COMException raised = new COMException("Outlook let go of the session.", unchecked((int)0x80010108));
        await using Boundary boundary = await Boundary.StartAsync(ScriptedSession.Create(_ => raised));

        COMException surfaced = Assert.IsType<COMException>(
            await boundary.FailureOfAsync(nameof(IOutlookSession.TryReadItem)));

        Assert.Equal(raised.Message, surfaced.Message);
        Assert.Equal(unchecked((int)0x80010108), surfaced.HResult);
    }

    [Fact]
    public async Task ARefusal_KeepsItsMachineReadableReason()
    {
        // A refusal's Reason is what the tool layer publishes as `error.reason`, and it is
        // the field an agent is told to branch on. It is read off the exception reflectively
        // on the child side, so a wrapper in the way loses it silently - it is simply absent
        // from the frame, and an absent reason looks exactly like an operation that has none.
        DraftRefusedException raised = new DraftRefusedException(
            "not_created_by_this_server", "This draft was not created by this server, so it will not be discarded.");
        await using Boundary boundary = await Boundary.StartAsync(ScriptedSession.Create(_ => raised));

        DraftRefusedException surfaced = Assert.IsType<DraftRefusedException>(
            await boundary.FailureOfAsync(nameof(IOutlookSession.TryDiscardDraft)));

        Assert.Equal("not_created_by_this_server", surfaced.Reason);
        Assert.Equal(raised.Message, surfaced.Message);
    }

    [Fact]
    public async Task AFailureTheParentDoesNotModel_StillArrivesUnderItsOwnName()
    {
        // The default branch of the mapper. The type cannot be rebuilt, so the caller gets a
        // ComHostRemoteException - but it carries the CHILD's type name, and the tool layer
        // reports that rather than the transport's. The message is the child's own either
        // way; what would otherwise be lost is the word saying what kind of failure it was.
        // Invariant 10 in check-pinned-constants.ps1 stops the COM layer adding deliberate
        // failures that land here.
        await using Boundary boundary = await Boundary.StartAsync(
            ScriptedSession.Create(_ => new FileNotFoundException("The signature file has gone.")));

        ComHostRemoteException surfaced = Assert.IsType<ComHostRemoteException>(
            await boundary.FailureOfAsync(nameof(IOutlookSession.GetAccounts)));

        Assert.Equal(nameof(FileNotFoundException), surfaced.RemoteType);
        Assert.Equal("The signature file has gone.", surfaced.Message);
    }

    [Fact]
    public void AFailureWithNoTypeAtAll_StillNamesSomething()
    {
        // ComHostError.Type defaults to empty, and the supervisor synthesises an error of
        // its own when a frame arrives carrying none. The tool layer now reports RemoteType
        // verbatim as the error type, so a blank one would publish `"type": ""` - a worse
        // answer than naming the transport the failure crossed.
        ComHostRemoteException rebuilt = Assert.IsType<ComHostRemoteException>(
            ComHostErrorMapper.ToException(new ComHostError { Type = string.Empty, Message = "no detail" }));

        Assert.Equal(nameof(ComHostRemoteException), rebuilt.RemoteType);
        Assert.Equal("no detail", rebuilt.Message);
    }

    [Fact]
    public async Task NoFailureArrivesAsAReflectionWrapper()
    {
        // The regression test proper, stated as the negative it is: whatever else changes
        // about the boundary, the sentence below must never be what a caller reads.
        await using Boundary boundary = await Boundary.StartAsync(
            ScriptedSession.Create(operation => new InvalidOperationException(FolderMessageFor(operation))));

        foreach (MethodInfo method in typeof(IOutlookSession).GetMethods())
        {
            Exception surfaced = await boundary.FailureOfAsync(method.Name);

            Assert.NotEqual(nameof(TargetInvocationException), surfaced.GetType().Name);
            Assert.DoesNotContain(
                "Exception has been thrown by the target of an invocation",
                surfaced.Message,
                StringComparison.Ordinal);
        }
    }

    /// <summary>
    /// Every exception type <see cref="ComHostErrorMapper"/> models. Each is raised on the
    /// child side and expected back unchanged; a type that stops being modelled falls to the
    /// mapper's default branch and fails its case here.
    /// </summary>
    public static TheoryData<string> ModelledFailures()
    {
        return new TheoryData<string>
        {
            "send refusal",
            "draft refusal",
            "outlook unavailable",
            "com failure",
            "detached rcw",
            "disposed",
            "null argument",
            "range",
            "argument",
            "invalid operation",
            "not supported",
            "denied",
            "io",
            "timeout",
        };
    }

    private static Exception Raise(string label)
    {
        switch (label)
        {
            case "send refusal":
                return new SendRefusedException("stale_token", "The confirm token no longer matches this draft.");
            case "draft refusal":
                return new DraftRefusedException("already_sent", "This draft has already been sent.");
            case "outlook unavailable":
                return new OutlookUnavailableException("Outlook is being updated; try again shortly.");
            case "com failure":
                return new COMException("Outlook rejected the call.", unchecked((int)0x80010108));
            case "detached rcw":
                return new InvalidComObjectException("The COM object has been separated from its underlying RCW.");
            case "disposed":
                return new ObjectDisposedException(string.Empty, "The session has been disposed.");
            case "null argument":
                return new ArgumentNullException(paramName: null, message: "Value cannot be null.");
            case "range":
                return new ArgumentOutOfRangeException(paramName: null, message: "absoluteWalkCap must be positive.");
            case "argument":
                return new ArgumentException("EntryID must not be blank.");
            case "invalid operation":
                return new InvalidOperationException("Folder 'Nope' was not found in store 'Work'. Use list_folders for paths.");
            case "not supported":
                return new NotSupportedException("This store does not support the operation.");
            case "denied":
                return new UnauthorizedAccessException("Access to the target directory is denied.");
            case "io":
                return new IOException("The attachment could not be written.");
            case "timeout":
                return new TimeoutException("The scan ran out of its time budget.");
            default:
                throw new ArgumentException("Unknown failure label '" + label + "'.", nameof(label));
        }
    }

    /// <summary>The shape of message that actually went missing: specific, and carrying the next action.</summary>
    private static string FolderMessageFor(string operation)
    {
        return "Folder 'Nope/Missing' was not found in store 'Work' (no child folder named 'Nope') while running '"
            + operation + "'. Use list_folders for paths.";
    }

    private static bool IsErrorOutput(ParameterInfo parameter)
    {
        return parameter.IsOut
            && string.Equals(parameter.Name, "error", StringComparison.Ordinal)
            && parameter.ParameterType.GetElementType() == typeof(string);
    }

    /// <summary>
    /// The real child side of the pipe, in this process: <see cref="GatewayRoutingProxy"/>
    /// over a stand-in session, served by the real <see cref="ComHostServer"/>, with the
    /// parent end speaking the real framing.
    /// </summary>
    private sealed class Boundary : IAsyncDisposable
    {
        private readonly NamedPipeServerStream _childEnd;
        private readonly NamedPipeClientStream _parentEnd;
        private readonly CancellationTokenSource _shutdown;
        private readonly Task _serving;
        private long _nextId;

        private Boundary(NamedPipeServerStream childEnd, NamedPipeClientStream parentEnd, Task serving, CancellationTokenSource shutdown)
        {
            _childEnd = childEnd;
            _parentEnd = parentEnd;
            _serving = serving;
            _shutdown = shutdown;
        }

        internal static async Task<Boundary> StartAsync(IOutlookSession session)
        {
            string name = ComHostProtocol.NewPipeName();
            NamedPipeServerStream childEnd = new NamedPipeServerStream(
                name, PipeDirection.InOut, 1, PipeTransmissionMode.Byte, PipeOptions.Asynchronous);
            NamedPipeClientStream parentEnd = new NamedPipeClientStream(
                ".", name, PipeDirection.InOut, PipeOptions.Asynchronous);

            Task listening = childEnd.WaitForConnectionAsync();
            await parentEnd.ConnectAsync((int)Patience.TotalMilliseconds).ConfigureAwait(false);
            await listening.WaitAsync(Patience).ConfigureAwait(false);

            ComHostServer server = new ComHostServer(
                childEnd, GatewayRoutingProxy.Create(new PassThroughGateway(session)), typeof(IOutlookSession));

            CancellationTokenSource shutdown = new CancellationTokenSource();
            Task serving = server.ServeAsync(shutdown.Token);
            Boundary boundary = new Boundary(childEnd, parentEnd, serving, shutdown);

            // The host announces itself before it serves anything; consuming that frame here
            // keeps the framing in step for every later call.
            ComHostResponse ready = await boundary.ReadAsync().ConfigureAwait(false);
            Assert.Equal(ComHostEvents.Ready, ready.Event);

            return boundary;
        }

        /// <summary>Calls one operation and returns the raw response frame.</summary>
        internal async Task<ComHostResponse> CallAsync(string operation)
        {
            long id = Interlocked.Increment(ref _nextId);
            byte[] frame = ComHostProtocol.EncodeFrame(new ComHostRequest { Id = id, Operation = operation });
            await _parentEnd.WriteAsync(frame).ConfigureAwait(false);
            await _parentEnd.FlushAsync().ConfigureAwait(false);

            ComHostResponse response = await ReadAsync().ConfigureAwait(false);
            Assert.Equal(id, response.Id);
            return response;
        }

        /// <summary>Calls one operation and rebuilds the failure exactly as the parent does.</summary>
        internal async Task<Exception> FailureOfAsync(string operation)
        {
            ComHostResponse response = await CallAsync(operation).ConfigureAwait(false);

            Assert.False(response.Ok, operation + " answered successfully; it was asked to fail.");
            Assert.NotNull(response.Error);
            return ComHostErrorMapper.ToException(response.Error!);
        }

        public async ValueTask DisposeAsync()
        {
            await _shutdown.CancelAsync().ConfigureAwait(false);
            _parentEnd.Dispose();

            try
            {
                await _serving.WaitAsync(Patience).ConfigureAwait(false);
            }
            catch (Exception ex) when (ex is OperationCanceledException || ex is IOException || ex is ObjectDisposedException)
            {
                // Tearing the pipe down is how this host is meant to stop.
            }

            _childEnd.Dispose();
            _shutdown.Dispose();
        }

        private async Task<ComHostResponse> ReadAsync()
        {
            ComHostResponse? response = await ComHostProtocol
                .ReadFrameAsync<ComHostResponse>(_parentEnd, CancellationToken.None)
                .WaitAsync(Patience)
                .ConfigureAwait(false);

            Assert.NotNull(response);
            return response!;
        }
    }

    /// <summary>
    /// Stands where <c>ComGateway</c> stands on the child side and does nothing but run the
    /// operation, so what this file observes is the boundary and never the gateway.
    /// </summary>
    private sealed class PassThroughGateway : IComGateway
    {
        private readonly IOutlookSession _session;

        internal PassThroughGateway(IOutlookSession session)
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
    /// A stand-in Outlook session that answers however the test asks, per operation.
    /// <para>
    /// A <see cref="DispatchProxy"/> for the reason the production fault session is one: it
    /// is a REAL object reached through <c>targetMethod.Invoke</c>, so whatever it raises
    /// travels the exact reflective hop a real session's failure travels - which is where
    /// the message used to be lost. Twenty-six hand-written stubs would not.
    /// </para>
    /// </summary>
    internal class ScriptedSession : DispatchProxy
    {
        private Func<string, Exception?> _script = _ => null;

        internal static IOutlookSession Create(Func<string, Exception?> script)
        {
            object proxy = Create<IOutlookSession, ScriptedSession>()
                ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
            ((ScriptedSession)proxy)._script = script;
            return (IOutlookSession)proxy;
        }

        /// <inheritdoc />
        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            ArgumentNullException.ThrowIfNull(targetMethod);

            Exception? failure = _script(targetMethod.Name);
            if (failure != null)
            {
                throw failure;
            }

            // Nothing scripted: report failure the way most of this contract does - null,
            // with the not-found token in the by-ref error parameter. EVERY out parameter is
            // assigned, not just that one: an out slot left as the proxy handed it over is
            // unboxed on the way back and takes the call down with a failure that has
            // nothing to do with the boundary (`out long sizeBytes` on save_attachment
            // found this).
            ParameterInfo[] parameters = targetMethod.GetParameters();
            for (int i = 0; args != null && i < parameters.Length && i < args.Length; i++)
            {
                ParameterInfo parameter = parameters[i];
                if (!parameter.IsOut)
                {
                    continue;
                }

                if (IsErrorOutput(parameter))
                {
                    args[i] = ComErrorTokens.ItemNotFound;
                    continue;
                }

                Type slot = parameter.ParameterType.GetElementType()!;
                args[i] = slot.IsValueType ? Activator.CreateInstance(slot) : null;
            }

            Type returnType = targetMethod.ReturnType;
            return returnType != typeof(void) && returnType.IsValueType
                ? Activator.CreateInstance(returnType)
                : null;
        }
    }
}
