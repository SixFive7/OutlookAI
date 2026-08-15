using System.Globalization;

namespace OutlookAI.ComHost.Host
{
    /// <summary>
    /// Test-only fault injection for the COM host, driven by an environment variable.
    /// <para>
    /// This exists because the behaviour that matters most in this architecture -
    /// "a wedged Outlook call becomes a bounded, structured failure and the next call
    /// succeeds" - is otherwise only reachable by genuinely wedging a real Outlook. That
    /// is not reproducible, not safe to do on a machine holding real mail, and impossible
    /// in CI.
    /// </para>
    /// <para>
    /// Faults are applied BEFORE the call reaches Outlook, on purpose: it means the whole
    /// timeout / kill / respawn path can be exercised on a machine with no Outlook at all,
    /// which is exactly where CI runs.
    /// </para>
    /// <para>
    /// Syntax, in <c>OUTLOOKAI_COMHOST_FAULT</c>:
    /// <list type="bullet">
    /// <item><c>hang:*</c> or <c>hang:OperationName</c> - never answer (simulates a wedged COM call)</item>
    /// <item><c>crash:*</c> or <c>crash:OperationName</c> - exit the process abruptly</item>
    /// <item><c>delay:2500:*</c> - answer, but only after the given milliseconds</item>
    /// <item><c>throw:*</c> - fail with a COMException</item>
    /// </list>
    /// Unset - the overwhelmingly normal case - costs one null check per call.
    /// </para>
    /// </summary>
    internal static class ComHostFaultInjection
    {
        /// <summary>Environment variable naming the fault to inject. Unset in production.</summary>
        internal const string Variable = "OUTLOOKAI_COMHOST_FAULT";

        private static readonly string? Spec = Environment.GetEnvironmentVariable(Variable);

        /// <summary>True when any fault is configured. Reported by health so an injected fault is never mistaken for a real one.</summary>
        internal static bool IsActive => !string.IsNullOrWhiteSpace(Spec);

        /// <summary>The configured fault specification, for diagnostics.</summary>
        internal static string? Description => Spec;

        /// <summary>Applies any configured fault for <paramref name="operation"/>.</summary>
        internal static void Apply(string operation)
        {
            if (string.IsNullOrWhiteSpace(Spec))
            {
                return;
            }

            string[] parts = Spec!.Split(':', StringSplitOptions.TrimEntries);
            if (parts.Length < 2)
            {
                return;
            }

            string kind = parts[0];
            string target = parts[^1];
            if (!Matches(target, operation))
            {
                return;
            }

            switch (kind.ToLowerInvariant())
            {
                case "hang":
                    // Block this thread forever. The host serves serially, so this models a
                    // wedged COM call faithfully: nothing else is answered either, and only
                    // the parent killing the process can end it.
                    Thread.Sleep(Timeout.Infinite);
                    return;

                case "crash":
                    Environment.FailFast($"Injected COM host crash on '{operation}'.");
                    return;

                case "delay":
                    if (parts.Length >= 3 && int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out int ms))
                    {
                        Thread.Sleep(ms);
                    }

                    return;

                case "throw":
                    throw new System.Runtime.InteropServices.COMException(
                        $"Injected COM failure on '{operation}'.", unchecked((int)0x80010108));

                default:
                    return;
            }
        }

        private static bool Matches(string target, string operation)
        {
            return target == "*" || string.Equals(target, operation, StringComparison.OrdinalIgnoreCase);
        }
    }
}
