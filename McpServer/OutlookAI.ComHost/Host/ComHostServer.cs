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

        internal ComHostServer(Stream pipe, object target, Type contract)
        {
            _pipe = pipe;
            _target = target;
            _contract = contract;
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
                catch (IOException)
                {
                    return;
                }
            }
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
            byte[] frame = ComHostProtocol.EncodeFrame(response);
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
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                return FromException(request.Id, ex.InnerException);
            }
            catch (Exception ex)
            {
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
            string? reason = TryReadReason(ex);
            return new ComHostResponse
            {
                Id = id,
                Ok = false,
                Error = new ComHostError
                {
                    Type = ex.GetType().Name,
                    Message = ex.Message,
                    HResult = ex is COMException com ? com.HResult : null,
                    Reason = reason,
                },
            };
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
