using System;
using System.Collections.Generic;

namespace OutlookAI.Core.Com
{
    /// <summary>Deadline class of an operation.</summary>
    /// <remarks>
    /// It lives in Core, beside <see cref="ComOperationBudgets"/>, for the reason that type
    /// records: the dependency only runs one way (<c>OutlookAI.ComHost</c> references
    /// <c>OutlookAI.Core</c>, never the reverse), so a class that is only nameable in the
    /// supervisor's assembly cannot be decided anywhere the service layer - or a test - can
    /// see. It used to be declared beside <c>ComHostPolicy</c> and assigned by a private
    /// method on a <c>DispatchProxy</c>, which meant the single decision "which tool gets
    /// which hang detector" sat on a line no test could execute.
    /// </remarks>
    public enum ComHostOperationClass
    {
        /// <summary>An ordinary mailbox operation.</summary>
        Operation = 0,

        /// <summary>Establishing the session, possibly cold-starting Outlook.</summary>
        Connect = 1,

        /// <summary>A health probe, which must degrade rather than block.</summary>
        HealthProbe = 2,

        /// <summary>
        /// An exhaustive folder scan: the one operation a caller picks BECAUSE
        /// completeness matters more than speed, and the one whose budget expiry is a
        /// documented answer rather than an incident. It has its own deadline so that
        /// giving it ten minutes does not give every other tool ten minutes.
        /// </summary>
        ExhaustiveScan = 3,

        /// <summary>
        /// A freshness check layered over an answer the index has already produced - the
        /// search path's folder sweep and <c>thread</c>'s conversation walk.
        /// <para>
        /// Same argument as <see cref="ExhaustiveScan"/>, different work. The sweep runs on
        /// every search over a profile that may hold tens of gigabytes, its expiry is a
        /// documented partial answer rather than an incident, and the maintainer's standing
        /// rule is that completeness outranks speed. Giving it ten minutes must not give
        /// <c>read</c> and <c>move_mail</c> ten minutes before a wedged Outlook is
        /// reclaimed - which is exactly what happened when the sweep shared
        /// <see cref="Operation"/> and its budget was pushed past that class's deadline.
        /// </para>
        /// </summary>
        FreshnessSweep = 4,
    }

    /// <summary>
    /// Which deadline class each <see cref="IOutlookSession"/> operation is dispatched
    /// under - i.e. which hang detector stands over it.
    /// <para>
    /// <b>Why this is a table in Core rather than a branch in the transport.</b> The
    /// classification used to be a private <c>ClassifyOperation</c> on
    /// <c>RemoteSessionProxy</c>, a <c>DispatchProxy</c> that cannot be constructed without
    /// a live supervisor and a child process. So the one decision that says "the sweep may
    /// take ten minutes, <c>read</c> may not" was unreachable from any test: stretching a
    /// quick tool's detector, or dropping the sweep back onto the ordinary class, changed
    /// behaviour that nothing in CI could observe. Moving the decision here does not make it
    /// cleverer, it makes it drivable - <c>T1 FreshnessSweepClassTests</c> now walks the
    /// whole contract and asserts the class of every method on it.
    /// </para>
    /// <para>
    /// <b>The rule.</b> A long class is for work a caller chose knowing it would be slow and
    /// whose expiry is a documented partial answer. Everything else - every tool a user
    /// waits on, every mutation - keeps the ordinary hang detector, because that is the only
    /// thing that reclaims a wedged Outlook for them. Fail-closed: a name this table does
    /// not know answers <see cref="ComHostOperationClass.Operation"/>, so a contract method
    /// added without a decision gets the SHORT detector rather than silently inheriting a
    /// long one.
    /// </para>
    /// </summary>
    public static class ComOperationClasses
    {
        /// <summary>
        /// The cheapest call on the contract, which is what makes it the liveness probe: a
        /// slow one is itself the signal, so it must not wait an ordinary budget to say so.
        /// </summary>
        private static readonly HashSet<string> HealthProbeNames = new HashSet<string>(StringComparer.Ordinal)
        {
            nameof(IOutlookSession.GetProfileName),
        };

        /// <summary>
        /// <see cref="IOutlookSession.ExhaustiveScan"/> and nothing else. Its own inner
        /// budget stops it gracefully well before the class deadline, so reaching the outer
        /// one really is a wedge.
        /// </summary>
        private static readonly HashSet<string> ExhaustiveScanNames = new HashSet<string>(StringComparer.Ordinal)
        {
            nameof(IOutlookSession.ExhaustiveScan),
        };

        /// <summary>
        /// The two live freshness checks, and the reason they are one class:
        /// <list type="bullet">
        /// <item><description><see cref="IOutlookSession.SweepFoldersNewerThan"/> - the
        /// search path's sweep of the window the index has not caught up with.</description></item>
        /// <item><description><see cref="IOutlookSession.TryGetConversationItems"/> -
        /// <c>thread</c>'s live conversation walk, which <c>MailService.ThreadWalkBudgetMs</c>
        /// already declared to be "the same budget the freshness sweep runs under ... because
        /// it is the same kind of work: a bounded live COM check layered over an answer the
        /// index already produced". It shares the budget, so it has to share the threshold
        /// that budget is judged against - otherwise the walk's ordinary expiry starts
        /// counting toward the circuit breaker while the sweep's does not.</description></item>
        /// </list>
        /// </summary>
        private static readonly HashSet<string> FreshnessSweepNames = new HashSet<string>(StringComparer.Ordinal)
        {
            nameof(IOutlookSession.SweepFoldersNewerThan),
            nameof(IOutlookSession.TryGetConversationItems),
        };

        /// <summary>The health-probe class's operations, for the T1 guard.</summary>
        public static IReadOnlyCollection<string> HealthProbeOperations => HealthProbeNames;

        /// <summary>The exhaustive-scan class's operations, for the T1 guard.</summary>
        public static IReadOnlyCollection<string> ExhaustiveScanOperations => ExhaustiveScanNames;

        /// <summary>The freshness class's operations, for the T1 guard.</summary>
        public static IReadOnlyCollection<string> FreshnessSweepOperations => FreshnessSweepNames;

        /// <summary>
        /// The deadline class <paramref name="operationName"/> is dispatched under.
        /// Fail-closed to <see cref="ComHostOperationClass.Operation"/>, which is the SHORT
        /// hang detector: forgetting to classify a new contract method costs it a longer
        /// budget it might have wanted, never costs every other tool their hang detection.
        /// </summary>
        public static ComHostOperationClass ClassOf(string? operationName)
        {
            if (operationName == null)
            {
                return ComHostOperationClass.Operation;
            }

            if (HealthProbeNames.Contains(operationName))
            {
                return ComHostOperationClass.HealthProbe;
            }

            if (ExhaustiveScanNames.Contains(operationName))
            {
                return ComHostOperationClass.ExhaustiveScan;
            }

            return FreshnessSweepNames.Contains(operationName)
                ? ComHostOperationClass.FreshnessSweep
                : ComHostOperationClass.Operation;
        }
    }
}
