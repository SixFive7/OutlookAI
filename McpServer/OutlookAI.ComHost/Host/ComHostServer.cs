using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using OutlookAI.ComHost.Protocol;

namespace OutlookAI.ComHost.Host
{
    /// <summary>
    /// The child side of the pipe: reads operation requests, invokes them against the
    /// real COM session, and writes results back.
    /// <para>
    /// Dispatch is by reflection over the shared contract interface rather than a
    /// hand-written switch. Both ends bind to the same <see cref="Type"/>, so a method
    /// added to the contract is automatically callable, and there is no second list of
    /// operation names to drift out of step with the first.
    /// </para>
    /// <para>
    /// Requests are handled STRICTLY SERIALLY. That is not a limitation to be optimised
    /// away later - all Outlook work funnels onto one pumped STA thread regardless, so
    /// concurrency here would only queue deeper. Serial handling also gives the parent a
    /// simple, true model: if a request has not answered, that request is the one that is
    /// stuck.
    /// </para>
    /// </summary>
    internal sealed class ComHostServer
    {
        private readonly Stream _pipe;
        private readonly object _target;
        private readonly Type _contract;
        private readonly SemaphoreSlim _writeLock = new SemaphoreSlim(1, 1);
        private readonly Dictionary<string, MethodInfo> _operations;
        private readonly int _maxFrameBytes;

        /// <param name="pipe">The connected pipe to the parent.</param>
        /// <param name="target">The object the contract calls are dispatched to.</param>
        /// <param name="contract">The interface whose methods are callable.</param>
        /// <param name="maxFrameBytes">
        /// Test seam. Production takes the default, so the shipped ceiling is
        /// <see cref="ComHostProtocol.MaxFrameBytes"/>; a test lowers it to reach the
        /// refusal path below without building a 64 MB answer to provoke it.
        /// </param>
        internal ComHostServer(Stream pipe, object target, Type contract, int maxFrameBytes = ComHostProtocol.MaxFrameBytes)
        {
            _pipe = pipe;
            _target = target;
            _contract = contract;
            _maxFrameBytes = maxFrameBytes;
            _operations = contract
                .GetMethods()
                .GroupBy(m => m.Name, StringComparer.Ordinal)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.Ordinal);
        }

        /// <summary>Serves until the pipe closes or <paramref name="cancellationToken"/> fires.</summary>
        internal async Task ServeAsync(CancellationToken cancellationToken)
        {
            await SendEventAsync(ComHostEvents.Ready).ConfigureAwait(false);

            while (!cancellationToken.IsCancellationRequested)
            {
                ComHostRequest? request;
                try
                {
                    request = await ComHostProtocol
                        .ReadFrameAsync<ComHostRequest>(_pipe, cancellationToken)
                        .ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    return;
                }
                catch (IOException)
                {
                    // Parent went away mid-frame.
                    return;
                }

                if (request == null)
                {
                    // Clean EOF: the parent closed the pipe. Exiting here is what makes a
                    // parent shutdown reliably take the child with it, independently of
                    // the job object.
                    return;
                }

                ComHostResponse response = Invoke(request);
                try
                {
                    await WriteAsync(response, cancellationToken).ConfigureAwait(false);
                }
                catch (ComHostProtocolException ex)
                {
                    // The answer is too large to frame. Answering with the refusal is the
                    // whole fix: this used to leave the serve loop, print to stderr and end
                    // the process, so a caller that asked for too much was told the COM host
                    // had gone away - and every LATER request died with the host too,
                    // including the ones that would have fitted.
                    //
                    // The refusal frame is a few hundred bytes, so it cannot hit the same
                    // limit. Only an IOException can stop it, and that means the parent is
                    // already gone.
                    try
                    {
                        await WriteAsync(TooLarge(request, ex), cancellationToken).ConfigureAwait(false);
                    }
                    catch (IOException)
                    {
                        return;
                    }
                }
                catch (IOException)
                {
                    return;
                }
            }
        }

        /// <summary>
        /// Turns a framing refusal into an answer the caller can act on.
        /// <para>
        /// Names the operation as well as the size and the limit. The size alone says the
        /// answer was too big; the operation is what tells the caller WHICH request to ask
        /// for less of, and it is known here and nowhere further up.
        /// </para>
        /// </summary>
        private static ComHostResponse TooLarge(ComHostRequest request, ComHostProtocolException refusal)
        {
            return Failure(
                request.Id,
                nameof(ComHostResponseTooLargeException),
                $"The answer to '{request.Operation}' was too large to return in one piece: {refusal.Message} "
                + "The work itself succeeded and nothing was changed; only the reply was refused.");
        }

        /// <summary>Pushes an unsolicited event to the parent (e.g. Outlook exited).</summary>
        internal async Task SendEventAsync(string eventName)
        {
            try
            {
                await WriteAsync(
                    new ComHostResponse { Id = 0, Event = eventName, Ok = true },
                    CancellationToken.None).ConfigureAwait(false);
            }
            catch (Exception)
            {
                // An event is advisory. Losing one must never take down the host: the
                // parent's own liveness checks cover the same ground more slowly.
            }
        }

        private async Task WriteAsync(ComHostResponse response, CancellationToken cancellationToken)
        {
            // Encoded OUTSIDE the write lock on purpose: a refusal must not have taken the
            // lock, or the substitute frame the serve loop writes next would deadlock
            // against it.
            byte[] frame = ComHostProtocol.EncodeFrame(response, _maxFrameBytes);
            await _writeLock.WaitAsync(cancellationToken).ConfigureAwait(false);
            try
            {
                await _pipe.WriteAsync(frame, cancellationToken).ConfigureAwait(false);
                await _pipe.FlushAsync(cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                _ = _writeLock.Release();
            }
        }

        private ComHostResponse Invoke(ComHostRequest request)
        {
            if (!_operations.TryGetValue(request.Operation, out MethodInfo? method))
            {
                return Failure(
                    request.Id,
                    nameof(MissingMethodException),
                    $"Operation '{request.Operation}' is not part of {_contract.Name}.");
            }

            object?[] arguments;
            try
            {
                arguments = BindArguments(method, request.Arguments);
            }
            catch (Exception ex)
            {
                return Failure(request.Id, nameof(ArgumentException), $"Could not bind arguments for '{request.Operation}': {ex.Message}");
            }

            try
            {
                // Test-only, and a no-op unless OUTLOOKAI_COMHOST_FAULT is set. Applied
                // before the call reaches Outlook so the timeout/kill/respawn path is
                // exercisable on a machine with no Outlook at all.
                ComHostFaultInjection.Apply(request.Operation);

                object? result = method.Invoke(_target, arguments);
                JsonElement? encoded = null;
                if (method.ReturnType != typeof(void) && result != null)
                {
                    encoded = JsonSerializer.SerializeToElement(result, method.ReturnType, ComHostProtocol.Json);
                }

                return new ComHostResponse
                {
                    Id = request.Id,
                    Ok = true,
                    Result = encoded,
                    Outputs = CollectOutputs(method, arguments),
                };
            }
            catch (Exception ex)
            {
                // No TargetInvocationException case of its own: FromException peels the
                // whole chain. One unwrap rule, in one place - two of them is how the
                // second stops matching the first.
                return FromException(request.Id, ex);
            }
        }

        private static object?[] BindArguments(MethodInfo method, JsonElement? argumentObject)
        {
            ParameterInfo[] parameters = method.GetParameters();
            if (parameters.Length == 0)
            {
                return Array.Empty<object?>();
            }

            object?[] bound = new object?[parameters.Length];
            for (int i = 0; i < parameters.Length; i++)
            {
                ParameterInfo parameter = parameters[i];
                string name = parameter.Name ?? throw new ArgumentException("Contract parameters must be named.");

                if (argumentObject is { ValueKind: JsonValueKind.Object } obj
                    && obj.TryGetProperty(name, out JsonElement value)
                    && value.ValueKind != JsonValueKind.Undefined)
                {
                    bound[i] = value.Deserialize(parameter.ParameterType, ComHostProtocol.Json);
                    continue;
                }

                // Absent argument: use the declared default when there is one, otherwise
                // the type's default. Optional parameters are common on this contract.
                bound[i] = parameter.HasDefaultValue
                    ? parameter.DefaultValue
                    : DefaultOf(parameter.ParameterType);
            }

            return bound;
        }

        private static object? DefaultOf(Type type)
        {
            if (type.IsByRef)
            {
                type = type.GetElementType()!;
            }

            return type.IsValueType ? Activator.CreateInstance(type) : null;
        }

        /// <summary>
        /// Captures the post-invocation values of by-ref parameters so the caller's own
        /// <c>out</c> variables can be filled in on the far side.
        /// </summary>
        private static Dictionary<string, JsonElement>? CollectOutputs(MethodInfo method, object?[] arguments)
        {
            Dictionary<string, JsonElement>? outputs = null;
            ParameterInfo[] parameters = method.GetParameters();
            for (int i = 0; i < parameters.Length; i++)
            {
                ParameterInfo parameter = parameters[i];
                if (!parameter.ParameterType.IsByRef)
                {
                    continue;
                }

                Type valueType = parameter.ParameterType.GetElementType()!;
                outputs ??= new Dictionary<string, JsonElement>(StringComparer.Ordinal);
                outputs[parameter.Name!] = JsonSerializer.SerializeToElement(arguments[i], valueType, ComHostProtocol.Json);
            }

            return outputs;
        }

        private static ComHostResponse FromException(long id, Exception ex)
        {
            Exception actual = Unwrap(ex);
            string? reason = TryReadReason(actual);
            return new ComHostResponse
            {
                Id = id,
                Ok = false,
                Error = new ComHostError
                {
                    Type = actual.GetType().Name,
                    Message = actual.Message,
                    HResult = actual is COMException com ? com.HResult : null,
                    Reason = reason,
                },
            };
        }

        /// <summary>
        /// Peels reflection wrappers off a failure before it becomes the wire error.
        /// <para>
        /// This is the LAST point at which a wrapper can be caught: past here, Type,
        /// Message, HResult and Reason are all that survive, and every one of them is read
        /// off the exception handed in. A <see cref="TargetInvocationException"/> answers
        /// them with "TargetInvocationException", "Exception has been thrown by the target
        /// of an invocation.", null and null - which is precisely the four-way loss that
        /// hid a good, specific, actionable message from every caller until 2026-08-18.
        /// </para>
        /// <para>
        /// The routing proxy now unwraps its own hop, so in a correct build this loop runs
        /// once (for this method's own reflective dispatch) and no further. It is written as
        /// a loop anyway because the failure it guards against is silent: a reflective layer
        /// added anywhere below would re-flatten everything and break no test that does not
        /// specifically look for it. The depth cap only stops a self-referential chain from
        /// spinning; nothing legitimate approaches it.
        /// </para>
        /// </summary>
        private static Exception Unwrap(Exception ex)
        {
            const int MaxDepth = 8;
            Exception current = ex;
            int depth = 0;
            while (depth++ < MaxDepth && current is TargetInvocationException { InnerException: { } inner })
            {
                current = inner;
            }

            return current;
        }

        /// <summary>
        /// Draft and send refusals carry a machine-readable Reason the tool layer surfaces
        /// verbatim. Read reflectively so this file needs no reference to those types.
        /// </summary>
        private static string? TryReadReason(Exception ex)
        {
            PropertyInfo? property = ex.GetType().GetProperty("Reason", BindingFlags.Public | BindingFlags.Instance);
            if (property == null || property.PropertyType != typeof(string))
            {
                return null;
            }

            return property.GetValue(ex) as string;
        }

        private static ComHostResponse Failure(long id, string type, string message)
        {
            return new ComHostResponse
            {
                Id = id,
                Ok = false,
                Error = new ComHostError { Type = type, Message = message },
            };
        }
    }
}
