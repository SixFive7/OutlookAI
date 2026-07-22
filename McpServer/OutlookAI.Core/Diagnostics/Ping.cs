using System;
using System.Reflection;
using System.Runtime.Versioning;

namespace OutlookAI.Core.Diagnostics
{
    /// <summary>
    /// Phase-0 scaffold diagnostics: pure, host-neutral logic behind the MCP <c>echo</c> tool.
    /// Proves the Server -> Core reference chain and gives the T1 tier a real cross-assembly
    /// target. Everything in Core must stay host-neutral (v3.MD section 0.5.2): no MCP types,
    /// no console assumptions, no per-session state.
    /// </summary>
    public static class Ping
    {
        /// <summary>Prefix every echo reply carries so tests can assert provenance.</summary>
        public const string EchoPrefix = "OutlookAI.Core echo: ";

        /// <summary>Returns <paramref name="message"/> wrapped in the canonical echo envelope.</summary>
        /// <exception cref="ArgumentNullException"><paramref name="message"/> is null.</exception>
        public static string Echo(string message)
        {
            if (message == null)
            {
                throw new ArgumentNullException(nameof(message));
            }

            return EchoPrefix + message;
        }

        /// <summary>
        /// The target framework the loaded Core assembly was compiled for, e.g.
        /// ".NETFramework,Version=v4.8" or ".NETCoreApp,Version=v10.0". Lets hosts and
        /// tests verify which of the two Core targets they are running against.
        /// </summary>
        public static string TargetFramework =>
            typeof(Ping).Assembly.GetCustomAttribute<TargetFrameworkAttribute>()?.FrameworkName ?? "unknown";
    }
}
