using System.Collections.Concurrent;
using System.Diagnostics;
using System.IO.Pipes;
using System.Runtime.Versioning;
using System.Text.Json;
using OutlookAI.ComHost.Protocol;
using OutlookAI.Core.Com;

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
        private int _consecutiveTimeouts;
        private long _lastTimeoutTimestamp;
        private long _lastStartAttemptTimestamp;
        private int _warmUpInFlight;
        private string _lastLivenessDetail = string.Empty;
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

        /// <summary>Consecutive operation timeouts; 0 once Outlook answers again.</summary>
        internal int ConsecutiveTimeouts => Volatile.Read(ref _consecutiveTimeouts);

        /// <summary>Outlook's externally observed state, judged without COM.</summary>
        internal OutlookLivenessState Liveness => OutlookLiveness.Probe();

        /// <summary>Detail behind the last liveness observation, for health output.</summary>
        internal string LivenessDetail
        {
            get
            {
                lock (_stateLock)
                {
                    return _lastLivenessDetail;
                }
            }
        }

        /// <summary>Whether COM requests are currently being failed fast because Outlook is not answering.</summary>
        internal bool IsUnresponsive => ComHostPolicy.DecideBreaker(new BreakerInput(
            Volatile.Read(ref _consecutiveTimeouts),
            MillisecondsSince(Volatile.Read(ref _lastTimeoutTimestamp)))) != BreakerVerdict.Closed;

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
            bool allowConnectFloor,
            CancellationToken cancellationToken)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);

            GateOutlookLiveness();

            BreakerVerdict verdict = ComHostPolicy.DecideBreaker(new BreakerInput(
                Volatile.Read(ref _consecutiveTimeouts),
                MillisecondsSince(Volatile.Read(ref _lastTimeoutTimestamp))));

            if (verdict == BreakerVerdict.Open)
            {
                // Fail in microseconds rather than spending another full budget to
                // rediscover what the last two requests established.
                throw new ComHostUnresponsiveException(
                    Volatile.Read(ref _consecutiveTimeouts), ComHostPolicy.UnresponsiveRetryAfterSeconds);
            }

            if (verdict == BreakerVerdict.HalfOpen && !await ProbeAliveAsync(cancellationToken).ConfigureAwait(false))
            {
                throw new ComHostUnresponsiveException(
                    Volatile.Read(ref _consecutiveTimeouts), ComHostPolicy.UnresponsiveRetryAfterSeconds);
            }

            return await InvokeCoreAsync(
                    operation, arguments, operationClass, deadlineOverrideMilliseconds, allowConnectFloor, cancellationToken)
                .ConfigureAwait(false);
        }

        /// <summary>
        /// Consults Windows about Outlook before spending anything on COM.
        /// <para>
        /// Three things fall out of one free check. A hung Outlook is refused instantly
        /// instead of after a 30-120 s budget. A starting Outlook produces retry guidance
        /// instead of a long block. And a not-running Outlook is started at most once per
        /// cooldown, which is the anti-churn guard: the 2026-08-16 root-cause analysis
        /// found the wedged instance was one we had started, 39 seconds after starting a
        /// previous one, most likely activating Outlook while the prior instance was still
        /// exiting.
        /// </para>
        /// </summary>
        /// <summary>
        /// Test-only override of the liveness probe, e.g. <c>responsive</c>.
        /// <para>
        /// Needed because the gate is now so effective that it hides the paths beneath it:
        /// on a machine where Outlook is genuinely hung, the gate refuses instantly and the
        /// injected-fault tests for the timeout / kill / respawn machinery never execute.
        /// Forcing the observed state keeps those tests deterministic wherever they run.
        /// </para>
        /// </summary>
        internal const string LivenessOverrideVariable = "OUTLOOKAI_COMHOST_LIVENESS";

        private static readonly OutlookLivenessState? LivenessOverride = ReadLivenessOverride();

        private static OutlookLivenessState? ReadLivenessOverride()
        {
            string? raw = Environment.GetEnvironmentVariable(LivenessOverrideVariable);
            return Enum.TryParse(raw, ignoreCase: true, out OutlookLivenessState parsed) ? parsed : null;
        }

        private void GateOutlookLiveness()
        {
            OutlookLivenessState liveness = OutlookLiveness.Probe(out string detail);
            if (LivenessOverride is OutlookLivenessState forced)
            {
                liveness = forced;
                detail = "forced by " + LivenessOverrideVariable;
            }
            lock (_stateLock)
            {
                _lastLivenessDetail = detail;
            }

            long sinceStart = MillisecondsSince(Volatile.Read(ref _lastStartAttemptTimestamp));
            LivenessVerdict verdict = ComHostPolicy.DecideLiveness(liveness, sinceStart, _allowStartingOutlook);
            int retryAfter = ComHostPolicy.RetryAfterSecondsFor(verdict, sinceStart);

            switch (verdict)
            {
                case LivenessVerdict.Proceed:
                    return;

                case LivenessVerdict.Hung:
                    RecordFailure($"Outlook is not responding ({detail}); requests needing it are refused immediately.");
                    throw new ComHostUnresponsiveException(detail, retryAfter);

                case LivenessVerdict.Starting:
                    BeginWarmUp();
                    throw new ComHostStartingException(retryAfter, detail);

                case LivenessVerdict.StartSuppressed:
                    if (!_allowStartingOutlook)
                    {
                        throw new ComHostUnavailableException(
                            "Outlook is not running and may not be started right now (the OutlookAI installer is running, "
                            + "or autostart is disabled). Retry shortly.");
                    }

                    throw new ComHostStartingException(retryAfter, "a start was attempted moments ago");

                case LivenessVerdict.MayStart:
                default:
                    Volatile.Write(ref _lastStartAttemptTimestamp, Stopwatch.GetTimestamp());
                    BeginWarmUp();
                    throw new ComHostStartingException(
                        ComHostPolicy.StartingRetryAfterSeconds, "Outlook was not running and is being started");
            }
        }

        /// <summary>
        /// Starts the COM host and warms up its Outlook connection in the background, so
        /// no caller ever waits out a cold start. At most one warm-up runs at a time.
        /// </summary>
        private void BeginWarmUp()
        {
            if (Interlocked.CompareExchange(ref _warmUpInFlight, 1, 0) != 0)
            {
                return;
            }

            _ = Task.Run(async () =>
            {
                try
                {
                    _ = await InvokeCoreAsync(
                            nameof(OutlookAI.Core.Com.IOutlookSession.GetProfileName),
                            null,
                            ComHostOperationClass.Connect,
                            null,
                            allowConnectFloor: false,
                            CancellationToken.None)
                        .ConfigureAwait(false);
                }
                catch (Exception)
                {
                    // A failed warm-up is not an error in itself: the caller was already
                    // told to retry, and the next attempt re-evaluates from scratch.
                }
                finally
                {
                    Volatile.Write(ref _warmUpInFlight, 0);
                }
            });
        }

        /// <summary>
        /// One cheap liveness probe, used to decide whether a wedged Outlook has come
        /// back. Deliberately GetProfileName on the health budget rather than the caller's
        /// real request: re-probing with a full 120 s operation would make every cooldown
        /// expiry cost two minutes again, which is the cost this whole mechanism exists to
        /// avoid.
        /// </summary>
        private async Task<bool> ProbeAliveAsync(CancellationToken cancellationToken)
        {
            try
            {
                _ = await InvokeCoreAsync(
                        nameof(OutlookAI.Core.Com.IOutlookSession.GetProfileName),
                        null,
                        ComHostOperationClass.HealthProbe,
                        ComHostPolicy.HealthProbeDeadlineMilliseconds,
                        allowConnectFloor: false,
                        cancellationToken)
                    .ConfigureAwait(false);
                return true;
            }
            catch (OperationCanceledException)
            {
                throw;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private async Task<ComHostInvocationResult> InvokeCoreAsync(
            string operation,
            object? arguments,
            ComHostOperationClass operationClass,
            long? deadlineOverrideMilliseconds,
            bool allowConnectFloor,
            CancellationToken cancellationToken)
        {

            long deadline = ComHostPolicy.DeadlineFor(operationClass, deadlineOverrideMilliseconds);

            // The first operation on a fresh child also pays for establishing the COM
            // session, which may cold-start OUTLOOK.EXE, so it gets a wider floor than the
            // ordinary budget - otherwise a legitimate cold start looks like a wedge.
            //
            // But ONLY when the caller expressed no opinion, or explicitly opted in. An
            // explicit budget is a deliberate statement of intent and outranks the floor.
            // This was wrong when first written and outlook_health paid for it: its
            // explicit 5 s probe was silently widened to the 90 s floor, and because health
            // makes two gateway calls it could block for ~180 s - against a wedged Outlook,
            // measured at 200 s+ on 2026-08-16. The one tool that must always answer was
            // the one made to wait longest.
            //
            // The opt-in exists because the OTHER explicit-budget caller wanted the
            // opposite: the freshness sweep's 30 s is a budget for the SWEEP, and
            // suppressing the floor meant that on a fresh host the first search had to fit
            // the COM attach and the whole sweep into it. Where attaching to a large OST
            // takes longer than 30 s that sweep could never succeed - every attempt timed
            // out, killed the host, bumped the restart count and blamed the sweep.
            if ((deadlineOverrideMilliseconds is not > 0 || allowConnectFloor)
                && !Volatile.Read(ref _childHasServed)
                && deadline < ComHostPolicy.ConnectFloorMilliseconds)
            {
                deadline = ComHostPolicy.ConnectFloorMilliseconds;
            }

            // The child-start handshake used to sit entirely outside the deadline system.
            // It is now bounded by this operation's own deadline (floored, so a test that
            // shortens the budget does not start failing on child startup instead) - and
            // the floor gives way to a budget the CALLER declared, on the same terms as the
            // connect floor above. Without that, outlook_health's explicit 5 s could spend
            // 10 s in handshake before its own clock even started.
            bool callerDeclaredBudget = deadlineOverrideMilliseconds is > 0 && !allowConnectFloor;
            await EnsureStartedAsync(
                    ComHostPolicy.HandshakeBudgetFor(deadline, callerDeclaredBudget), cancellationToken)
                .ConfigureAwait(false);

            long id = Interlocked.Increment(ref _nextId);
            PendingRequest pending = new PendingRequest(
                operation,
                deadline,
                countsTowardUnresponsive: ComHostPolicy.TimeoutIndicatesUnresponsiveness(
                    operationClass, deadlineOverrideMilliseconds));
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

            // Any answer at all means Outlook is talking again.
            Volatile.Write(ref _consecutiveTimeouts, 0);
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
                        //
                        // 4. Count it toward the breaker only when it is EVIDENCE. An
                        //    expiring caller-declared work budget says the work was big;
                        //    an expiring hang detector says Outlook is deaf. Counting both
                        //    meant two ordinary slow searches on a large mailbox opened the
                        //    breaker and failed every request for 30 s - an outage caused
                        //    by nothing but the size of the mailbox. The kill and the
                        //    restart happen either way: a blocked COM call cannot be
                        //    reclaimed any other way, and the child serves serially.
                        // Not `_ =` here: the enclosing lambda's parameter is named _.
                        if (pending.CountsTowardUnresponsive)
                        {
                            Interlocked.Increment(ref _consecutiveTimeouts);
                            Volatile.Write(ref _lastTimeoutTimestamp, Stopwatch.GetTimestamp());
                        }

                        BeginReplacement(pending.CountsTowardUnresponsive
                            ? $"'{pending.Operation}' exceeded its {pending.DeadlineMilliseconds} ms budget; the COM host was restarted."
                            : $"'{pending.Operation}' ran past the {pending.DeadlineMilliseconds} ms budget its caller set for it, so the "
                              + "COM host was restarted to reclaim the call. Outlook was not judged unresponsive by this.");
                        pending.Completion.TrySetException(
                            new ComHostTimeoutException(pending.Operation, pending.DeadlineMilliseconds));
                        KillChild($"deadline exceeded on '{pending.Operation}'", generation);
                    },
                    CancellationToken.None,
                    TaskContinuationOptions.OnlyOnRanToCompletion | TaskContinuationOptions.ExecuteSynchronously,
                    TaskScheduler.Default);
        }

        private async Task EnsureStartedAsync(long handshakeBudgetMilliseconds, CancellationToken cancellationToken)
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

                await StartChildAsync(handshakeBudgetMilliseconds, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                _ = _startLock.Release();
            }
        }

        private async Task StartChildAsync(long handshakeBudgetMilliseconds, CancellationToken cancellationToken)
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

            // ONE budget across BOTH halves of the handshake. Each half used to get a full
            // 30 s of its own, so a slow child start could cost 60 s before any deadline
            // applied at all - outside the deadline system, and unshortenable by
            // OUTLOOKAI_COMHOST_DEADLINE_MS. The handshake is one thing; it gets one clock.
            Stopwatch handshake = Stopwatch.StartNew();
            try
            {
                using CancellationTokenSource connectCts =
                    CancellationTokenSource.CreateLinkedTokenSource(cancellationToken, childCts.Token);
                connectCts.CancelAfter(TimeSpan.FromMilliseconds(handshakeBudgetMilliseconds));
                await pipe.WaitForConnectionAsync(connectCts.Token).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                // Name what actually happened. Reporting every failure here as a timeout
                // once masked a disposed-pipe race for several test runs, because the
                // message asserted a cause rather than reporting one.
                string reason = ex is OperationCanceledException
                    ? $"the COM host did not connect within {handshakeBudgetMilliseconds} ms"
                    : $"the COM host connection failed to establish ({ex.GetType().Name}: {ex.Message})";
                NoteStartFailure(reason);
                TearDownChild();
                throw new ComHostUnavailableException("The Outlook COM host did not connect.", ex);
            }

            _ = Task.Run(() => ReadLoopAsync(pipe, generation, childCts.Token), CancellationToken.None);

            long readyBudget = handshakeBudgetMilliseconds - handshake.ElapsedMilliseconds;
            if (readyBudget < 1)
            {
                readyBudget = 1;
            }

            try
            {
                await ready.Task.WaitAsync(TimeSpan.FromMilliseconds(readyBudget), cancellationToken)
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

                    // Counted here rather than where the failure reaches the caller: the
                    // child refuses an oversized answer in ITS process, so its own counter
                    // is invisible to health, and a refusal whose caller had already given
                    // up would go unrecorded if this waited for a pending match.
                    if (response.Error != null
                        && string.Equals(response.Error.Type, nameof(ComHostResponseTooLargeException), StringComparison.Ordinal))
                    {
                        ComHostFrameMeter.Shared.RecordRefusal();
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
            //
            // Say WHY, not just what. The COM host serves requests serially, so when one
            // request wedges and is reclaimed, its innocent siblings die with it - and
            // "the host stopped" leaves the caller with no idea that something else was
            // at fault or that retrying is reasonable.
            string? cause;
            lock (_stateLock)
            {
                cause = _lastFailureMessage;
            }

            foreach (KeyValuePair<long, PendingRequest> entry in _pending.ToArray())
            {
                if (_pending.TryRemove(entry.Key, out PendingRequest? pending))
                {
                    _ = pending.Completion.TrySetException(
                        new ComHostUnavailableException(DescribeInterruption(pending.Operation, cause)));
                }
            }
        }

        /// <summary>
        /// What an interrupted request is told, which depends on WHAT IT WAS DOING.
        /// <para>
        /// This used to end "This request itself was not at fault; retry it." for every
        /// victim alike. That is right for a read and wrong for a mutation: the child was
        /// terminated at an unknown point, so a killed <c>TrySendDraft</c> may already have
        /// submitted the mail and a killed <c>TryUpdateDraft</c> may have applied part of
        /// its ~20-call sequence, and "retry it" is then advice to send twice or to append
        /// the attachments a second time. The classification is the same
        /// <see cref="ComSessionOperations"/> the in-process gateway keys its
        /// disconnect-rebuild on, so there is one answer to "may this be run again" in the
        /// product rather than two, and it fails closed: an unclassified name is treated as
        /// mutating.
        /// </para>
        /// <para>
        /// Pure and internal so T1 can pin both halves without a child process.
        /// </para>
        /// </summary>
        internal static string DescribeInterruption(string operation, string? cause)
        {
            string opening = string.IsNullOrEmpty(cause)
                ? $"The Outlook COM host stopped before '{operation}' completed."
                : $"'{operation}' was interrupted because the Outlook COM host was restarted: {cause}";

            return ComSessionOperations.IsRetryable(operation)
                ? opening + " This request itself was not at fault, and it only READ - retrying it is safe."
                : opening + " This request itself was not at fault, but it CHANGES mail and the COM host ended before it "
                    + "answered, so whether it took effect is UNKNOWN. Do not simply retry it: check the current state "
                    + "first (read the item, or search for it) and decide from what you find.";
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

        /// <summary>
        /// How long a child gets to notice its pipe closed and exit on its own before it is
        /// terminated (2026-08-19).
        /// <para>
        /// The child already has a clean-exit path - <c>ComHostServer.ServeAsync</c> returns
        /// on EOF, with the comment "Exiting here is what makes a parent shutdown reliably
        /// take the child with it" - and <see cref="TearDownChild"/> already disposes the
        /// pipe before killing. But the kill followed the dispose in the same statement
        /// block, so the child was terminated microseconds after EOF and that path was dead
        /// code in practice. This gap makes it reachable: on an ORDERLY teardown the child
        /// runs its own finally blocks and releases its COM references, which is what the
        /// kill skips.
        /// </para>
        /// <para>
        /// 250 ms is chosen to be invisible: it is paid only when a child is being replaced
        /// or the server is shutting down, never on a served request. On the DEADLINE path
        /// it costs nothing at all, because <see cref="KillChild"/> has already terminated
        /// the process and <c>HasExited</c> is true by the time the wait is reached.
        /// </para>
        /// </summary>
        private const int CleanExitGraceMilliseconds = 250;

        /// <summary>
        /// Terminates the child that missed its deadline.
        /// <para>
        /// WHY A HARD KILL, AND WHY THAT DOES NOT CONTRADICT <c>PumpedStaRunner</c>. Inside
        /// this same child, <c>PumpedStaRunner.Dispose</c> refuses to abort its STA thread
        /// on the grounds that "a COM call could be mid-flight". Read side by side the two
        /// disciplines look opposed; they are not, and the difference is which process
        /// survives. <c>Thread.Abort</c> injects an asynchronous exception into a thread in
        /// a process that KEEPS RUNNING, leaving that process holding half-released COM
        /// proxies, a corrupted apartment and an unusable Application reference - a
        /// permanently broken child that still answers the pipe. <c>TerminateProcess</c>
        /// destroys the whole address space at once, and Windows tears down the LRPC
        /// endpoints and Outlook's references to the dead client with it. The child's
        /// refusal protects the child's own remaining life; this kill ends that life, so
        /// there is nothing left to protect.
        /// </para>
        /// <para>
        /// WHAT THE KILL DOES NOT REACH: OUTLOOK.EXE. It is not in the tree, established
        /// four ways (2026-08-18 audit) - there is exactly one <c>Process.Start</c> in this
        /// whole path and it starts the COM host; Outlook is COM-ACTIVATED, so the SCM's
        /// service host is its parent, which this repo already measured on a live wedge
        /// (<c>OUTLOOK.EXE -Embedding</c>, parent <c>svchost.exe</c>, see
        /// <see cref="AutostartCooldownMilliseconds"/>); job membership is inherited only
        /// by processes the job's members create, and the child creates none; and
        /// <c>entireProcessTree</c> walks recorded parentage. So this is a CLIENT death,
        /// not a server kill: no store is left half-written by it.
        /// </para>
        /// <para>
        /// WHAT IT DOES COST is caller certainty, and that is handled elsewhere: the
        /// pending request is completed as a timeout before this runs, its siblings are
        /// told whether their own operation is safe to retry
        /// (<see cref="OnChildConnectionLost"/>), and a killed send is reported with the
        /// Outbox warning by <c>MailService.Send</c>. What no grace period can fix is that
        /// a wedged child cannot answer a polite request at all: <c>ComHostServer.ServeAsync</c>
        /// calls <c>Invoke</c> synchronously inside its read loop, so while wedged the
        /// child is not reading the pipe and a stop frame is structurally undeliverable.
        /// </para>
        /// </summary>
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

            // Closing the pipe is EOF to the child, and the child exits on EOF by design.
            // Give it that chance before terminating it: an orderly exit runs its finally
            // blocks and releases its COM references, which a kill skips. Free on the
            // deadline path (KillChild has already terminated it), and bounded at
            // CleanExitGraceMilliseconds everywhere else.
            try
            {
                if (child is { HasExited: false })
                {
                    _ = child.WaitForExit(CleanExitGraceMilliseconds);
                }
            }
            catch (Exception)
            {
                // Already gone, or the handle is unusable. The kill below covers it.
            }

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
            internal PendingRequest(string operation, long deadlineMilliseconds, bool countsTowardUnresponsive)
            {
                Operation = operation;
                DeadlineMilliseconds = deadlineMilliseconds;
                CountsTowardUnresponsive = countsTowardUnresponsive;
            }

            internal string Operation { get; }

            internal long DeadlineMilliseconds { get; }

            /// <summary>
            /// Whether this request's deadline expiring is evidence that Outlook is
            /// unresponsive, rather than that the work was bigger than the budget its
            /// caller chose for it. Decided once, at dispatch, by
            /// <see cref="ComHostPolicy.TimeoutIndicatesUnresponsiveness"/> - the watchdog
            /// fires long after the inputs are out of scope.
            /// </summary>
            internal bool CountsTowardUnresponsive { get; }

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
