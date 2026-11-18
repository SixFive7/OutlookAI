using System.Reflection;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// The arguments a <see cref="DispatchProxy"/> stand-in receives, as the stand-ins for
/// <c>IOutlookSession</c> fill them in. The one copy (Q154).
/// </summary>
internal static class ProxyArgs
{
    /// <summary>Sets the out parameter named <paramref name="name"/>, where <paramref name="method"/> has one.</summary>
    public static void SetOut(MethodInfo method, object?[]? args, string name, object? value)
    {
        ParameterInfo[] parameters = method.GetParameters();
        for (int i = 0; args != null && i < parameters.Length && i < args.Length; i++)
        {
            if (parameters[i].IsOut && string.Equals(parameters[i].Name, name, StringComparison.Ordinal))
            {
                args[i] = value;
            }
        }
    }
}
