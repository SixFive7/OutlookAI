using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using OutlookAI.Core.Com;

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
    /// <item><c>sessionthrow:folder:*</c> / <c>sessionthrow:com:*</c> - fail INSIDE the
    /// Outlook session, behind the routing proxy's reflective invoke</item>
    /// </list>
    /// Unset - the overwhelmingly normal case - costs one null check per call.
    /// </para>
    /// <para>
    /// A second variable, <c>OUTLOOKAI_COMHOST_EXIT_DELAY_MS</c>, makes this process slow to
    /// EXIT rather than slow to answer. See <see cref="ApplyExitDelay"/>.
    /// </para>
    /// <para>
    /// <c>sessionthrow</c> is deliberately NOT another <see cref="Apply"/> kind, and the
    /// difference is the whole reason it exists. <c>Apply</c> runs in
    /// <see cref="ComHostServer"/>, ABOVE the routing proxy, so a fault raised there never
    /// crosses the reflective invoke where child-side exceptions used to be flattened into
    /// "Exception has been thrown by the target of an invocation." A regression test built
    /// on <c>throw</c> would therefore have passed both before and after that fix.
    /// <c>sessionthrow</c> substitutes the SESSION, so the failure travels the exact path a
    /// real one does - and still needs no Outlook, because the substitute never connects.
    /// </para>
    /// </summary>
    internal static class ComHostFaultInjection
    {
        /// <summary>The <c>sessionthrow</c> kind, which fails behind the routing proxy rather than in front of it.</summary>
        internal const string SessionThrowKind = "sessionthrow";

        /// <summary>
        /// The message the <c>folder</c> shape raises. Deliberately modelled on the real
        /// exhaustive folder-resolution error - the one whose loss started this - so the
        /// test asserts on the shape of message that actually goes missing: specific, and
        /// carrying the next action.
        /// </summary>
        internal const string SessionFolderMessage =
            "Folder 'Nope/Missing' was not found in store 'InjectedStore' (no child folder named 'Nope'). "
            + "Use list_folders for paths.";

        /// <summary>The message the <c>com</c> shape raises.</summary>
        internal const string SessionComMessage = "Injected COM failure raised inside the Outlook session.";

        /// <summary>RPC_E_DISCONNECTED - the HRESULT that must survive the crossing for the disconnect logic to work.</summary>
        internal const int SessionComHResult = unchecked((int)0x80010108);

        /// <summary>Environment variable naming the fault to inject. Unset in production.</summary>
        internal const string Variable = "OUTLOOKAI_COMHOST_FAULT";

        /// <summary>
        /// Environment variable making this process SLOW TO EXIT, in milliseconds. Unset in
        /// production, and deliberately a second variable rather than another
        /// <see cref="Variable"/> kind: every kind there names an operation, and this one
        /// does not - it fires once, on the way out, whatever the child was serving. A test
        /// also needs it ALONGSIDE an operation fault, which one spec string cannot express.
        /// </summary>
        internal const string ExitDelayVariable = "OUTLOOKAI_COMHOST_EXIT_DELAY_MS";

        /// <summary>
        /// The ceiling on <see cref="ExitDelayVariable"/>. A typo cannot park the child for
        /// longer than the parent would ever wait for it anyway.
        /// </summary>
        internal const int MaxExitDelayMilliseconds = 30_000;

        private static readonly string? Spec = Environment.GetEnvironmentVariable(Variable);

        private static readonly string? ExitDelaySpec = Environment.GetEnvironmentVariable(ExitDelayVariable);

        /// <summary>True when any fault is configured. Reported by health so an injected fault is never mistaken for a real one.</summary>
        internal static bool IsActive => !string.IsNullOrWhiteSpace(Spec) || !string.IsNullOrWhiteSpace(ExitDelaySpec);

        /// <summary>
        /// The configured fault specification, for diagnostics. The exit delay is appended
        /// rather than folded in, so a run that sets only <see cref="Variable"/> - which is
        /// every existing test - still reports exactly the string it set.
        /// </summary>
        internal static string? Description => string.IsNullOrWhiteSpace(ExitDelaySpec)
            ? Spec
            : $"{Spec} +exitdelay:{ExitDelaySpec}ms".TrimStart();

        /// <summary>
        /// Makes this process take a while to exit, when <see cref="ExitDelayVariable"/> asks
        /// for it. A no-op - one null check - in every production run.
        /// <para>
        /// It stands in for the one thing a CI box cannot produce: the time a real
        /// <c>OutlookComSession.Dispose</c> takes to unadvise its sink, close the pin
        /// Explorer and release its references. Without it, the parent's shutdown grace is
        /// unfalsifiable in CI - a child with no Outlook session exits instantly, so a grace
        /// of 250 ms, 2000 ms or none at all all look identical from outside. With it, the
        /// grace has an observable consequence: the parent either waits for the child or
        /// terminates it mid-exit.
        /// </para>
        /// <para>
        /// This is a delay, NOT a COM release. It proves the child was GIVEN time to run its
        /// exit path; it cannot prove anything about what that path does to Outlook, which
        /// needs a real Outlook and lives in the live tier.
        /// </para>
        /// </summary>
        internal static void ApplyExitDelay()
        {
            if (string.IsNullOrWhiteSpace(ExitDelaySpec)
                || !int.TryParse(ExitDelaySpec, NumberStyles.Integer, CultureInfo.InvariantCulture, out int delay)
                || delay <= 0)
            {
                return;
            }

            Thread.Sleep(Math.Min(delay, MaxExitDelayMilliseconds));
        }

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

        /// <summary>
        /// Hands back a stand-in session that fails <paramref name="operation"/>, when one
        /// is configured. False - and no allocation - in every production run.
        /// </summary>
        internal static bool TrySessionFault(string operation, out IOutlookSession? session)
        {
            session = null;
            if (string.IsNullOrWhiteSpace(Spec))
            {
                return false;
            }

            string[] parts = Spec!.Split(':', StringSplitOptions.TrimEntries);

            // kind:shape:target - the shape is required, so a two-part spec is not this.
            if (parts.Length < 3 || !string.Equals(parts[0], SessionThrowKind, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (!Matches(parts[^1], operation))
            {
                return false;
            }

            session = FaultingSession.Create(parts[1]);
            return true;
        }

        private static bool Matches(string target, string operation)
        {
            return target == "*" || string.Equals(target, operation, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Stands in for the real Outlook session and fails every call it is given.
        /// <para>
        /// A <see cref="DispatchProxy"/> for two reasons. It satisfies all 26 contract
        /// members without 26 stubs that would need updating whenever the contract grows.
        /// And - the load-bearing one - it is a real object reached through
        /// <c>targetMethod.Invoke(session, args)</c>, so the exception it raises travels the
        /// same reflective hop a real session's does, which is where the message used to be
        /// lost.
        /// </para>
        /// </summary>
        internal class FaultingSession : DispatchProxy
        {
            private string _shape = string.Empty;

            internal static IOutlookSession Create(string shape)
            {
                object proxy = Create<IOutlookSession, FaultingSession>()
                    ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
                ((FaultingSession)proxy)._shape = shape;
                return (IOutlookSession)proxy;
            }

            /// <inheritdoc />
            protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
            {
                throw Build(_shape);
            }

            private static Exception Build(string shape)
            {
                return string.Equals(shape, "com", StringComparison.OrdinalIgnoreCase)
                    ? new COMException(SessionComMessage, SessionComHResult)
                    : new InvalidOperationException(SessionFolderMessage);
            }
        }
    }
}
