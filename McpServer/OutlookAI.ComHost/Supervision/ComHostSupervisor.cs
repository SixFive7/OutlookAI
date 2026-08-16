using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO.Pipes;
using System.Runtime.Versioning;
using System.Text.Json;
using OutlookAI.ComHost.Protocol;

namespace OutlookAI.ComHost.Supervision
{
    /// <summary>
    /// Parent-side owner of the COM child: spawns it, talks to it, bounds every request,
    /// and kills it when a request exceeds its budget.
    /// <para>
    /// The deadline watchdog is armed per request and runs INDEPENDENTLY of whoever is
    /// awaiting the result. That matters: if the MCP client cancels, the caller stops
    /// waiting, but the operation may still be wedged inside Outlook. Because the child
    /// serves requests serially, an abandoned-but-wedged operation would block every
    /// later one. Keeping the watchdog armed means the wedge is still reclaimed, even
    /// though nobody is listening any more.
    /// </para>
    /// </summary>
    [SupportedOSPlatform("windows")]
    internal sealed class ComHostSupervisor : IDisposable
    {
        private const string ChildExeName = "OutlookAI.ComHost.exe";
        private const int ReadyTimeoutMilliseconds = 30_000;

        private readonly bool _allowStartingOutlook;
        private readonly SemaphoreSlim _startLock = new SemaphoreSlim(1, 1);
        private readonly ConcurrentDictionary<long, PendingRequest> _pending = new ConcurrentDictionary<long, PendingRequest>();
        private readonly object _stateLock = new object();

        private ChildJobObject? _job;
        private Process? _child;
        private NamedPipeServerStream? _pipe;
        private CancellationTokenSource? _childCts;
        private TaskCompletionSource<bool>? _ready;
        private long _nextId;
        private volatile ComHostState _state = ComHostState.None;
        private int _consecutiveStartFailures;
        private long _lastStartFailureTimestamp;
        private int _restartCount;

        /// <summary>
        /// Incremented every time a child is installed. Teardown is scoped to a
        /// generation so a late kill cannot destroy its own replacement - see
        /// <see cref="TearDownChild"/>.
        /// </summary>
        private int _generation;
        private string? _lastFailureMessage;
        private bool _childHasServed;
        private bool _disposed;

        internal ComHostSupervisor(bool allowStartingOutlook)
        {
            _allowStartingOutlook = allowStartingOutlook;
        }

        /// <summary>Raised when the child reports its Outlook went away.</summary>
        internal event Action? OutlookGone;

        /// <summary>Current lifecycle state, for health reporting.</summary>
        internal ComHostState State => _state;

        /// <summary>How many times the child has been replaced this process lifetime.</summary>
        internal int RestartCount => Volatile.Read(ref _restartCount);

        /// <summary>The child's PID, or null when no child is running. Used by tests to prove a respawn happened.</summary>
        internal int? ChildProcessId
        {
            get
            {
                lock (_stateLock)
                {
                    try
                    {
                        return _child is { HasExited: false } ? _child.Id : null;
                    }
                    catch (InvalidOperationException)
                    {
                        return null;
                    }
                }
            }
        }

        /// <summary>Last failure the supervisor recorded, for health reporting. Null when healthy.</summary>
        internal string? LastFailureMessage
        {
            get
            {
                lock (_stateLock)
                {
                    return _lastFailureMessage;
                }
            }
        }

        /// <summary>
        /// Invokes a contract operation on the child, bounded by its deadline.
        /// </summary>
        /// <exception cref="ComHostTimeoutException">The deadline expired; the child was killed.</exception>
        /// <exception cref="ComHostUnavailableException">No child could be started.</exception>
        internal async Task<ComHostInvocationResult> InvokeAsync(
            string operation,
            object? arguments,
            ComHostOperationClass operationClass,
            long? deadlineOverrideMilliseconds,
            CancellationToken cancellationToken)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);

            long deadline = ComHostPolicy.DeadlineFor(operationClass, deadlineOverrideMilliseconds);

            // The first operation on a fresh child also pays for establishing the COM
            // session, which may cold-start OUTLOOK.EXE, so it gets a wider floor than the
            // ordinary budget - otherwise a legitimate cold start looks like a wedge.
            //
            // But ONLY when the caller expressed no opinion. An explicit budget is a
            // deliberate statement of intent and outranks the floor. This was wrong when
            // first written and outlook_health paid for it: its explicit 5 s probe was
            // silently widened to the 90 s floor, and because health makes two gateway
            // calls it could block for ~180 s - against a wedged Outlook, measured at
            // 200 s+ on 2026-08-16. The one tool that must always answer was the one made
            // to wait longest.
            if (deadlineOverrideMilliseconds is not > 0
                && !Volatile.Read(ref _childHasServed)
                && deadline < ComHostPolicy.ConnectFloorMilliseconds)
            {
                deadline = ComHostPolicy.ConnectFloorMilliseconds;
            }
            await EnsureStartedAsync(cancellationToken).ConfigureAwait(false);

            long id = Interlocked.Increment(ref _nextId);
            PendingRequest pending = new PendingRequest(operation, deadline);
            if (!_pending.TryAdd(id, pending))
            {
                throw new InvalidOperationException("Duplicate request id.");
            }

            ArmDeadline(id, pending);

            try
            {
                JsonElement? argumentElement = arguments == null
                    ? null
                    : JsonSerializer.SerializeToElement(arguments, ComHostProtocol.Json);

                await SendAsync(new ComHostRequest { Id = id, Operation = operation, Arguments = argumentElement })
                    .ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                _ = _pending.TryRemove(id, out _);
                pending.Dispose();
                throw new ComHostUnavailableException("Could not send the request to the COM host.", ex);
            }

            ComHostResponse response;
            try
            {
                response = await pending.Completion.Task.WaitAsync(cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                _ = _pending.TryRemove(id, out _);
                pending.Dispose();
            }

            if (!response.Ok)
            {
                throw ComHostErrorMapper.ToException(response.Error
                    ?? new ComHostError { Type = "Exception", Message = "The COM host reported a failure with no detail." });
            }

            Volatile.Write(ref _childHasServed, true);
            return new ComHostInvocationResult(response.Result, response.Outputs);
        }

        /// <summary>
        /// Arms the per-request deadline. Deliberately not tied to the caller's token -
        /// see the class remarks.
        /// </summary>
        private void ArmDeadline(long id, PendingRequest pending)
        {
            int generation;
            lock (_stateLock)
            {
                generation = _generation;
            }

            _ = Task.Delay(TimeSpan.FromMilliseconds(pending.DeadlineMilliseconds), pending.DeadlineCts.Token)
                .ContinueWith(
                    _ =>
                    {
                        if (!_pending.ContainsKey(id))
                        {
                            return;
                        }

                        // The operation is still outstanding past its budget. Reclaim it
                        // the only way a blocked COM call can be reclaimed.
                        //
                        // The ordering here is subtle and was wrong twice, in opposite
                        // directions, so it is spelled out:
                        //
                        // 1. Record the decision to replace the host - failure text, state
                        //    and restart count - BEFORE releasing the caller. The caller
                        //    issues its next request the instant it is released, and that
                        //    request must not observe a host that still looks Ready, nor a
                        //    restart count that has not caught up yet.
                        // 2. Complete the request as a TIMEOUT, before the kill. Killing
                        //    first tears down the connection, and teardown fails everything
                        //    outstanding as "the host stopped" - which would win the race
                        //    and report the vaguer cause, hiding both that we ended it and
                        //    why.
                        // 3. Kill last. It is the slowest step and nothing waits on it.
                        BeginReplacement($"'{pending.Operation}' exceeded its {pending.DeadlineMilliseconds} ms budget; the COM host was restarted.");
                        pending.Completion.TrySetException(
                            new ComHostTimeoutException(pending.Operation, pending.DeadlineMilliseconds));
                        KillChild($"deadline exceeded on '{pending.Operation}'", generation);
                    },
                    CancellationToken.None,
                    TaskContinuationOptions.OnlyOnRanToCompletion | TaskContinuationOptions.ExecuteSynchronously,
                    TaskScheduler.Default);
        }

        private async Task EnsureStartedAsync(CancellationToken cancellationToken)
        {
            DispatchVerdict verdict = ComHostPolicy.DecideDispatch(new DispatchInput(
                _state,
                Volatile.Read(ref _consecutiveStartFailures),
                MillisecondsSince(Volatile.Read(ref _lastStartFailureTimestamp)),
                _allowStartingOutlook));

            switch (verdict)
            {
                case DispatchVerdict.Dispatch:
                    return;

                case DispatchVerdict.RefuseBackoff:
                    throw new ComHostUnavailableException(
                        $"The Outlook COM host failed to start {Volatile.Read(ref _consecutiveStartFailures)} times in a row; "
                        + "further attempts are paused briefly. Check outlook_health.");

                case DispatchVerdict.RefuseUnavailable:
                    throw new ComHostUnavailableException(
                        "Outlook may not be started right now (the OutlookAI installer is running, or autostart is disabled). Retry shortly.");

                case DispatchVerdict.StartThenDispatch:
                default:
                    break;
            }

            await _startLock.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                // Re-check: another caller may have started it while we queued.
                if (_state == ComHostState.Ready)
                {
                    return;
                }

                await StartChildAsync(cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                _ = _startLock.Release();
            }
        }

        private async Task StartChildAsync(CancellationToken cancellationToken)
        {
            TearDownChild();
            Volatile.Write(ref _childHasServed, false);

            string exePath = ResolveChildExecutable();
            string pipeName = ComHostProtocol.NewPipeName();

            NamedPipeServerStream pipe = new NamedPipeServerStream(
                pipeName,
                PipeDirection.InOut,
                maxNumberOfServerInstances: 1,
                PipeTransmissionMode.Byte,
                PipeOptions.Asynchronous);

            ProcessStartInfo startInfo = new ProcessStartInfo(exePath)
            {
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardError = true,
                // The child inherits the parent's working directory by default, which
                // Claude Code sets to the user's project folder. Anchor it to the exe
                // directory so nothing resolves relative to an arbitrary location.
                WorkingDirectory = Path.GetDirectoryName(exePath) ?? Environment.CurrentDirectory,
            };
            startInfo.Environment[ComHostProtocol.PipeNameVariable] = pipeName;
            startInfo.Environment[ComHostProtocol.ParentPidVariable] =
                Environment.ProcessId.ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (!_allowStartingOutlook)
            {
                startInfo.ArgumentList.Add("--no-autostart");
            }

            Process child;
            try
            {
                child = Process.Start(startInfo)
                    ?? throw new ComHostUnavailableException($"Process.Start returned no process for {exePath}.");
            }
            catch (Exception ex) when (ex is not ComHostUnavailableException)
            {
                pipe.Dispose();
                NoteStartFailure(ex.Message);
                throw new ComHostUnavailableException($"Could not start the Outlook COM host at {exePath}.", ex);
            }

            ChildJobObject job = ChildJobObject.CreateOrInert();
            _ = job.TryAssign(child);

            CancellationTokenSource childCts = new CancellationTokenSource();
            TaskCompletionSource<bool> ready = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);

            int generation;
            lock (_stateLock)
            {
                _pipe = pipe;
                _child = child;
                _job = job;
                _childCts = childCts;
                _ready = ready;
                _state = ComHostState.Starting;
                generation = ++_generation;
            }

            DrainStandardError(child);

            try
            {
                using CancellationTokenSource connectCts =
                    CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, childCts.Token);
                connectCts.CancelAfter(ReadyTimeoutMilliseconds);
                await pipe.WaitForConnectionAsync(connectCts.Token).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                // Name what actually happened. Reporting every failure here as a timeout
                // once masked a disposed-pipe race for several test runs, because the
                // message asserted a cause rather than reporting one.
                string reason = ex is OperationCanceledException
                    ? $"the COM host did not connect within {ReadyTimeoutMilliseconds} ms"
                    : $"the COM host connection failed to establish ({ex.GetType().Name}: {ex.Message})";
                NoteStartFailure(reason);
                TearDownChild();
                throw new ComHostUnavailableException("The Outlook COM host did not connect.", ex);
            }

            _ = Task.Run(() => ReadLoopAsync(pipe, generation, childCts.Token), CancellationToken.None);

            try
            {
                await ready.Task.WaitAsync(TimeSpan.FromMilliseconds(ReadyTimeoutMilliseconds), cancellationToken)
                    .ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                NoteStartFailure("the COM host connected but never reported ready");
                TearDownChild();
                throw new ComHostUnavailableException("The Outlook COM host never reported ready.", ex);
            }

            lock (_stateLock)
            {
                _state = ComHostState.Ready;

                // _lastFailureMessage is deliberately NOT cleared here. A recovered wedge
                // must still leave a trace: the restart count says one happened, and this
                // says what wedged - which is the actionable half. Clearing it on a
                // successful restart would leave health reporting "restarted once" with no
                // explanation, and invisibility is the exact failure this work exists to
                // end. It reads as "last failure this session", not "current failure".
            }

            Volatile.Write(ref _consecutiveStartFailures, 0);
        }

        private async Task ReadLoopAsync(Stream pipe, int generation, CancellationToken cancellationToken)
        {
            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    ComHostResponse? response = await ComHostProtocol
                        .ReadFrameAsync<ComHostResponse>(pipe, cancellationToken)
                        .ConfigureAwait(false);
                    if (response == null)
                    {
                        break;
                    }

                    if (!string.IsNullOrEmpty(response.Event))
                    {
                        HandleEvent(response.Event);
                        continue;
                    }

                    if (_pending.TryGetValue(response.Id, out PendingRequest? pending))
                    {
                        _ = pending.Completion.TrySetResult(response);
                    }

                    // An unmatched id is a response to a request whose caller gave up.
                    // Dropping it is correct - the deadline watchdog already decided.
                }
            }
            catch (OperationCanceledException)
            {
                // Normal teardown.
            }
            catch (Exception ex)
            {
                RecordFailure($"The COM host connection failed: {ex.Message}");
            }
            finally
            {
                OnChildConnectionLost(generation);
            }
        }

        private void HandleEvent(string name)
        {
            if (string.Equals(name, ComHostEvents.Ready, StringComparison.Ordinal))
            {
                TaskCompletionSource<bool>? ready;
                lock (_stateLock)
                {
                    ready = _ready;
                }

                _ = ready?.TrySetResult(true);
                return;
            }

            if (string.Equals(name, ComHostEvents.OutlookGone, StringComparison.Ordinal))
            {
                OutlookGone?.Invoke();
            }
        }

        private void OnChildConnectionLost(int generation)
        {
            lock (_stateLock)
            {
                if (_generation == generation && _state != ComHostState.None)
                {
                    _state = ComHostState.Faulted;
                }
            }

            // Everything still outstanding dies with the connection. Failing them
            // explicitly is the whole point: silence is the bug being fixed.
            foreach (KeyValuePair<long, PendingRequest> entry in _pending.ToArray())
            {
                if (_pending.TryRemove(entry.Key, out PendingRequest? pending))
                {
                    _ = pending.Completion.TrySetException(new ComHostUnavailableException(
                        $"The Outlook COM host stopped before '{pending.Operation}' completed."));
                }
            }
        }

        private async Task SendAsync(ComHostRequest request)
        {
            NamedPipeServerStream? pipe;
            lock (_stateLock)
            {
                pipe = _pipe;
            }

            if (pipe == null || !pipe.IsConnected)
            {
                throw new ComHostUnavailableException("The Outlook COM host is not connected.");
            }

            byte[] frame = ComHostProtocol.EncodeFrame(request);
            await pipe.WriteAsync(frame).ConfigureAwait(false);
            await pipe.FlushAsync().ConfigureAwait(false);
        }

        private void KillChild(string why, int generation)
        {
            Process? child;
            lock (_stateLock)
            {
                if (_generation != generation)
                {
                    return;
                }

                child = _child;
                _state = ComHostState.Faulted;
            }

            if (child == null)
            {
                return;
            }

            try
            {
                if (!child.HasExited)
                {
                    // Not counted here: BeginReplacement already recorded the decision.
                    // Counting in both places would double-count one reclaimed wedge.
                    child.Kill(entireProcessTree: true);
                }
            }
            catch (Exception)
            {
                // Already gone, or access denied on an exiting process. Either way the
                // connection teardown below covers it.
            }

            _ = why;
            TearDownChild(generation);
        }

        /// <summary>
        /// Tears down the current child, or - when <paramref name="onlyGeneration"/> is
        /// given - only if that is still the current one.
        /// <para>
        /// The scoping is load-bearing. The deadline watchdog releases the waiting caller
        /// BEFORE it kills, so the caller can have started a replacement child by the time
        /// the kill runs. An unscoped teardown would then capture the NEW pipe and dispose
        /// it, and the replacement would fail to connect - reported, misleadingly, as a
        /// connect timeout. That is a real bug this fixes, not a theoretical one.
        /// </para>
        /// </summary>
        private void TearDownChild(int? onlyGeneration = null)
        {
            Process? child;
            NamedPipeServerStream? pipe;
            CancellationTokenSource? cts;
            ChildJobObject? job;

            lock (_stateLock)
            {
                if (onlyGeneration is int generation && _generation != generation)
                {
                    // Superseded: this kill refers to a child that has already been
                    // replaced. Doing nothing is exactly right.
                    return;
                }

                child = _child;
                pipe = _pipe;
                cts = _childCts;
                job = _job;
                _child = null;
                _pipe = null;
                _childCts = null;
                _job = null;
                _ready = null;
                if (_state != ComHostState.None)
                {
                    _state = ComHostState.Faulted;
                }
            }

            try
            {
                cts?.Cancel();
            }
            catch (ObjectDisposedException)
            {
            }

            cts?.Dispose();
            pipe?.Dispose();

            try
            {
                if (child is { HasExited: false })
                {
                    child.Kill(entireProcessTree: true);
                }
            }
            catch (Exception)
            {
                // Best-effort.
            }

            child?.Dispose();

            // Closing the job handle terminates anything still in it.
            job?.Dispose();
        }

        private static void DrainStandardError(Process child)
        {
            // The child writes diagnostics to stderr. Read them so a full pipe buffer can
            // never block it, and surface them on our own stderr where the MCP client
            // shows server logs. Never stdout - that carries JSON-RPC.
            _ = Task.Run(async () =>
            {
                try
                {
                    string? line;
                    while ((line = await child.StandardError.ReadLineAsync().ConfigureAwait(false)) != null)
                    {
                        await Console.Error.WriteLineAsync("[com-host] " + line).ConfigureAwait(false);
                    }
                }
                catch (Exception)
                {
                    // The child exited; nothing further to drain.
                }
            });
        }

        /// <summary>
        /// Locates the child beside this assembly. Correct in both layouts: installed at
        /// %LOCALAPPDATA%\OutlookAI\Setup\McpServer\, and dev builds where a ProjectReference
        /// copies the child's apphost next to the parent. Never the current working
        /// directory - Claude Code sets that to the user's project folder.
        /// </summary>
        private static string ResolveChildExecutable()
        {
            string path = Path.Combine(AppContext.BaseDirectory, ChildExeName);
            if (!File.Exists(path))
            {
                throw new ComHostUnavailableException(
                    $"The Outlook COM host executable is missing. Expected it at: {path}. "
                    + "This indicates an incomplete installation.");
            }

            return path;
        }

        private void NoteStartFailure(string message)
        {
            _ = Interlocked.Increment(ref _consecutiveStartFailures);
            Volatile.Write(ref _lastStartFailureTimestamp, Stopwatch.GetTimestamp());
            RecordFailure(message);
        }

        /// <summary>
        /// Marks the host as no longer usable and counts the replacement, atomically and
        /// BEFORE any waiting caller is released.
        /// <para>
        /// This is what makes the restart observable at the moment it is decided rather
        /// than whenever the kill happens to finish. A caller released first would race
        /// ahead and see a host that still looked Ready - and try to send on a pipe being
        /// torn down underneath it.
        /// </para>
        /// </summary>
        private void BeginReplacement(string message)
        {
            lock (_stateLock)
            {
                _lastFailureMessage = message;
                _state = ComHostState.Faulted;
            }

            _ = Interlocked.Increment(ref _restartCount);
        }

        private void RecordFailure(string message)
        {
            lock (_stateLock)
            {
                _lastFailureMessage = message;
            }
        }

        private static long MillisecondsSince(long timestamp)
        {
            if (timestamp == 0)
            {
                return long.MaxValue;
            }

            return (long)Stopwatch.GetElapsedTime(timestamp).TotalMilliseconds;
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            TearDownChild();
            _startLock.Dispose();
        }

        private sealed class PendingRequest : IDisposable
        {
            internal PendingRequest(string operation, long deadlineMilliseconds)
            {
                Operation = operation;
                DeadlineMilliseconds = deadlineMilliseconds;
            }

            internal string Operation { get; }

            internal long DeadlineMilliseconds { get; }

            internal CancellationTokenSource DeadlineCts { get; } = new CancellationTokenSource();

            internal TaskCompletionSource<ComHostResponse> Completion { get; } =
                new TaskCompletionSource<ComHostResponse>(TaskCreationOptions.RunContinuationsAsynchronously);

            public void Dispose()
            {
                try
                {
                    DeadlineCts.Cancel();
                }
                catch (ObjectDisposedException)
                {
                }

                DeadlineCts.Dispose();
            }
        }
    }

    /// <summary>
    /// One COM-host invocation's outcome: the return value plus any by-ref parameter
    /// values. Returned together rather than stashed on the supervisor, so concurrent
    /// callers cannot read each other's output parameters.
    /// </summary>
    internal readonly record struct ComHostInvocationResult(
        JsonElement? Result,
        IReadOnlyDictionary<string, JsonElement>? Outputs);
}
