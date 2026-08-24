using System.Reflection;
using System.Runtime.Versioning;
using System.Text.Json;
using OutlookAI.ComHost.Protocol;
using OutlookAI.ComHost.Supervision;
using OutlookAI.Core.Com;

namespace OutlookAI.ComHost.Client
{
    /// <summary>
    /// Parent-side <see cref="IOutlookSession"/>: every call becomes one bounded round
    /// trip to the COM host.
    /// <para>
    /// A <see cref="DispatchProxy"/> so the two ends of the contract cannot drift. The
    /// child dispatches by reflecting over the same interface, so adding a method to
    /// <see cref="IOutlookSession"/> makes it work on both sides at once, with no
    /// hand-written pair to keep in step.
    /// </para>
    /// <para>
    /// The call blocks the calling thread. That is safe here in a way it was not before,
    /// and the distinction is the whole point of this architecture: the wait is bounded
    /// by the supervisor's deadline, and the supervisor can always make it end by killing
    /// the child. The old in-process design blocked on a COM call that nothing could
    /// interrupt.
    /// </para>
    /// </summary>
    [SupportedOSPlatform("windows")]
    public class RemoteSessionProxy : DispatchProxy
    {
        private ComHostSupervisor _supervisor = null!;

        /// <summary>Creates a session proxy bound to <paramref name="supervisor"/>.</summary>
        internal static IOutlookSession Create(ComHostSupervisor supervisor)
        {
            ArgumentNullException.ThrowIfNull(supervisor);

            object proxy = Create<IOutlookSession, RemoteSessionProxy>()
                ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
            ((RemoteSessionProxy)proxy)._supervisor = supervisor;
            return (IOutlookSession)proxy;
        }

        /// <inheritdoc />
        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            ArgumentNullException.ThrowIfNull(targetMethod);

            ParameterInfo[] parameters = targetMethod.GetParameters();
            Dictionary<string, object?> arguments = BuildArguments(parameters, args);

            // Which hang detector stands over this call. The table is in Core
            // (ComOperationClasses) rather than here, because "the sweep may take ten
            // minutes, read may not" is the load-bearing decision of the whole budget
            // ladder and a private method on a DispatchProxy is a line no test can execute.
            ComHostOperationClass operationClass = ComOperationClasses.ClassOf(targetMethod.Name);

            // Shrink this call to what is left of the enclosing operation's AGGREGATE
            // budget. Without it, a lambda that makes many contract calls got a full
            // deadline for each and the operation as a whole was unbounded.
            long callDeadline = ComHostPolicy.DeadlineFor(
                operationClass, ComHostRequestContext.DeadlineOverrideMilliseconds);
            long effectiveDeadline = ComHostPolicy.EffectiveDeadlineMilliseconds(
                callDeadline, ComHostRequestContext.RemainingAggregateMilliseconds);
            if (effectiveDeadline <= 0)
            {
                // The AGGREGATE is what ran out, not this call's own budget, so the message
                // says so: the earlier round trips of this same operation spent it.
                throw new TimeoutException(
                    $"The Outlook operation ran out of its overall time budget before '{targetMethod.Name}' could be sent "
                    + "to the COM host - earlier steps of the same operation used it up. Results are incomplete; narrow "
                    + "the request (fewer ids, a smaller folder scope) and try again.");
            }

            // Recorded BEFORE the round trip, because the failure this serves is one that
            // never comes back: a call that raises a bare COMException carries no operation
            // name of its own, and the tool layer needs one to decide whether "retry" is
            // safe advice.
            ComHostRequestContext.NoteOperation(targetMethod.Name);

            ComHostInvocationResult invocation = _supervisor.InvokeAsync(
                    targetMethod.Name,
                    arguments.Count == 0 ? null : arguments,
                    operationClass,
                    effectiveDeadline,
                    ComHostRequestContext.AllowConnectFloor,
                    ComHostRequestContext.Token)
                .GetAwaiter()
                .GetResult();

            WriteBackOutputs(parameters, args, invocation.Outputs);

            if (targetMethod.ReturnType == typeof(void) || invocation.Result is not JsonElement element)
            {
                return null;
            }

            return element.Deserialize(targetMethod.ReturnType, ComHostProtocol.Json);
        }

        private static Dictionary<string, object?> BuildArguments(ParameterInfo[] parameters, object?[]? args)
        {
            Dictionary<string, object?> arguments = new Dictionary<string, object?>(StringComparer.Ordinal);
            if (args == null)
            {
                return arguments;
            }

            for (int i = 0; i < parameters.Length && i < args.Length; i++)
            {
                ParameterInfo parameter = parameters[i];

                // A pure `out` parameter carries no input value; sending its uninitialised
                // slot would be noise at best and a deserialization failure at worst.
                if (parameter.IsOut)
                {
                    continue;
                }

                arguments[parameter.Name!] = args[i];
            }

            return arguments;
        }

        private static void WriteBackOutputs(
            ParameterInfo[] parameters,
            object?[]? args,
            IReadOnlyDictionary<string, JsonElement>? outputs)
        {
            if (args == null || outputs == null || outputs.Count == 0)
            {
                return;
            }

            for (int i = 0; i < parameters.Length && i < args.Length; i++)
            {
                ParameterInfo parameter = parameters[i];
                if (!parameter.ParameterType.IsByRef)
                {
                    continue;
                }

                if (!outputs.TryGetValue(parameter.Name!, out JsonElement value))
                {
                    continue;
                }

                Type valueType = parameter.ParameterType.GetElementType()!;
                args[i] = value.ValueKind == JsonValueKind.Null
                    ? null
                    : value.Deserialize(valueType, ComHostProtocol.Json);
            }
        }
    }
}
