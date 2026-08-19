using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.ExceptionServices;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// An <see cref="IOutlookSession"/> that refuses to START a contract call once the
    /// enclosing operation's budget is spent. The in-process half of the aggregate budget
    /// the COM host enforces by killing a child.
    /// <para>
    /// WHY IT EXISTS. <see cref="ComGateway"/>'s budget overload used to be
    /// <c>{ return Run(operation); }</c> - it accepted a budget and discarded it. That is
    /// correct INSIDE the COM host child, where the parent's watchdog is the real bound.
    /// But <c>MailService.CreateDefault()</c> builds on the same gateway, and EVERY live
    /// (T2) fixture is built on <c>CreateDefault()</c> - so the whole live tier ran with no
    /// budget, no aggregate, no breaker and no hang detector, and anything sized against
    /// those budgets was unverified by construction. A live-tier run that hung was
    /// diagnosed as an Outlook problem when the tier simply had nothing to stop it.
    /// </para>
    /// <para>
    /// WHAT IT CAN AND CANNOT DO, stated plainly because the difference is the entire
    /// reason the COM host exists. It CANNOT bound one call: a blocked outbound COM call is
    /// not cancellable, and killing the caller is not an option when the caller is us. It
    /// CAN bound a SEQUENCE, by checking the clock between calls and refusing the next one
    /// - which is exactly the aggregate half of what <c>RemoteSessionProxy</c> does, and is
    /// what turns a multi-call service operation from unbounded into bounded-with-one-call
    /// of overshoot.
    /// </para>
    /// <para>
    /// The clock starts AFTER the session is connected, so a cold start is never charged to
    /// the caller's work budget. That matches the remote gateway's <c>allowConnectFloor</c>
    /// opt-in for the callers that use it, and for the callers that do not it is simply
    /// honest: in-process there is nothing that could have bounded the connect either.
    /// </para>
    /// <para>
    /// Failures are re-thrown with <see cref="ExceptionDispatchInfo"/> rather than
    /// unwrapped by hand. Reflection wraps everything in
    /// <see cref="TargetInvocationException"/>, and this repository has already paid for
    /// that once: a reflective layer on the COM-host path flattened every deliberate error
    /// into "Exception has been thrown by the target of an invocation", which broke both
    /// the tool layer's advice (it branches on exception TYPE) and
    /// <see cref="ComGateway"/>'s disconnect-and-rebuild (it branches on
    /// <c>COMException</c> HRESULTs). Capturing and re-throwing keeps type, message, data
    /// and the original stack.
    /// </para>
    /// </summary>
    public class BudgetedSessionProxy : DispatchProxy
    {
        /// <summary>
        /// Least budget a call may be dispatched with. Mirrors
        /// <c>ComHostPolicy.MinimumDispatchDeadlineMilliseconds</c>: below this the budget
        /// is spent for practical purposes, and reporting that is more useful than starting
        /// work that has no time to finish.
        /// </summary>
        public const int MinimumRemainingMilliseconds = 1_000;

        private IOutlookSession _inner = null!;
        private long _startTimestamp;
        private long _budgetMilliseconds;

        /// <summary>
        /// Wraps <paramref name="inner"/> so contract calls stop once
        /// <paramref name="budgetMilliseconds"/> of wall clock has been spent. A
        /// non-positive budget returns the session unwrapped - "no budget" must not become
        /// "no calls".
        /// </summary>
        public static IOutlookSession Wrap(IOutlookSession inner, int budgetMilliseconds)
        {
            if (inner == null)
            {
                throw new ArgumentNullException(nameof(inner));
            }

            if (budgetMilliseconds <= 0)
            {
                return inner;
            }

            object proxy = Create<IOutlookSession, BudgetedSessionProxy>()
                ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
            BudgetedSessionProxy typed = (BudgetedSessionProxy)proxy;
            typed._inner = inner;
            typed._budgetMilliseconds = budgetMilliseconds;
            typed._startTimestamp = Stopwatch.GetTimestamp();
            return (IOutlookSession)proxy;
        }

        /// <summary>
        /// What is left of <paramref name="budgetMilliseconds"/> after
        /// <paramref name="elapsedMilliseconds"/>, clamped at zero. Pure so T1 can pin the
        /// boundary without a COM session.
        /// </summary>
        public static long RemainingMilliseconds(long budgetMilliseconds, long elapsedMilliseconds)
        {
            long remaining = budgetMilliseconds - elapsedMilliseconds;
            return remaining > 0 ? remaining : 0;
        }

        /// <summary>
        /// Whether a call may still be started. Pure, and deliberately the same shape as the
        /// remote side's dispatch floor: a sub-second remainder is reported as spent rather
        /// than dispatched.
        /// </summary>
        public static bool CanDispatch(long budgetMilliseconds, long elapsedMilliseconds)
        {
            return RemainingMilliseconds(budgetMilliseconds, elapsedMilliseconds) >= MinimumRemainingMilliseconds;
        }

        /// <summary>The message a caller gets when the aggregate is spent. Public so T1 pins it.</summary>
        public static string BudgetExhaustedMessage(string operation, long budgetMilliseconds)
        {
            return $"The Outlook operation ran out of its {budgetMilliseconds} ms time budget before '{operation}' could "
                + "be started - earlier steps of the same operation used it up. Results are incomplete; narrow the "
                + "request (fewer ids, a smaller folder scope) and try again.";
        }

        /// <inheritdoc />
        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            if (targetMethod == null)
            {
                throw new ArgumentNullException(nameof(targetMethod));
            }

            // Stopwatch.GetElapsedTime is net7+; Core still gates on net48 for the v3.1
            // event host, so the same arithmetic is written out.
            long elapsed = (Stopwatch.GetTimestamp() - _startTimestamp) * 1000L / Stopwatch.Frequency;
            if (!CanDispatch(_budgetMilliseconds, elapsed))
            {
                throw new TimeoutException(BudgetExhaustedMessage(targetMethod.Name, _budgetMilliseconds));
            }

            try
            {
                return targetMethod.Invoke(_inner, args);
            }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
                throw; // Unreachable; the line above always throws.
            }
        }
    }
}
