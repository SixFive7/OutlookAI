using System.Reflection;
using OutlookAI.Core.Com;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// A COM session that hands every call to one handler: a test's own fake, which answers the members it models,
/// for the suites whose fakes keep state of their own rather than being scripts on <see cref="StandInSession"/>.
/// Seven suites declared a proxy of their own that this replaces (the code-sharing refactor's S5a, the survey's
/// T03), three of them with the defaults <see cref="Create"/> describes. Not sealed: DispatchProxy derives from
/// its TProxy at runtime and refuses a sealed one.
/// </summary>
internal class ForwardingSession : DispatchProxy
{
    private Func<MethodInfo, object?[]?, object?> _handle = (member, _) => throw new NotSupportedException(member.Name);
    private bool _defaults;

    /// <summary>
    /// A session whose every call goes to <paramref name="handle"/>, given the member and its arguments. With
    /// <paramref name="defaultsForTheRest"/>, an out argument the handler left unassigned, and a value-typed
    /// result it did not give, are their type's default: an unassigned slot would be unboxed on the way back
    /// and fail the call for a reason that has nothing to do with the test.
    /// </summary>
    internal static IOutlookSession Create(Func<MethodInfo, object?[]?, object?> handle, bool defaultsForTheRest = false)
    {
        object proxy = Create<IOutlookSession, ForwardingSession>()
            ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
        ((ForwardingSession)proxy)._handle = handle;
        ((ForwardingSession)proxy)._defaults = defaultsForTheRest;
        return (IOutlookSession)proxy;
    }

    /// <inheritdoc />
    protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
    {
        ArgumentNullException.ThrowIfNull(targetMethod);
        object? handled = _handle(targetMethod, args);
        if (!_defaults)
        {
            return handled;
        }

        ParameterInfo[] parameters = targetMethod.GetParameters();
        for (int i = 0; args != null && i < parameters.Length && i < args.Length; i++)
        {
            if (parameters[i].IsOut && args[i] == null)
            {
                Type slot = parameters[i].ParameterType.GetElementType()!;
                args[i] = slot.IsValueType ? Activator.CreateInstance(slot) : null;
            }
        }

        if (handled != null)
        {
            return handled;
        }

        Type returnType = targetMethod.ReturnType;
        return returnType != typeof(void) && returnType.IsValueType ? Activator.CreateInstance(returnType) : null;
    }
}
