using System.Diagnostics;

namespace OutlookAI.ComHost.Client
{
    /// <summary>
    /// Ambient per-request state for calls that reach the COM host.
    /// <para>
    /// The service layer is synchronous across roughly 4,000 lines and 26 gateway call
    /// sites, and threading a <see cref="CancellationToken"/> through all of it would be
    /// a very large change for a small gain: the MCP SDK suppresses the response of a
    /// cancelled request regardless, so the token's only real job here is to stop the
    /// caller waiting. An <see cref="AsyncLocal{T}"/> carries it instead, set once by the
    /// tool layer and read by the session proxy.
    /// </para>
    /// <para>
    /// The same channel carries the AGGREGATE budget of the enclosing gateway operation.
    /// A per-call deadline bounds one round trip; a gateway operation is a lambda that may
    /// make many of them, and without an aggregate each one independently got a full
    /// budget, so the operation as a whole had no bound. <see cref="RemoteSessionProxy"/>
    /// shrinks every call's deadline to what is left.
    /// </para>
    /// <para>
    /// Note what this deliberately does NOT do: cancelling does not abandon the deadline.
    /// The supervisor keeps its watchdog armed for the abandoned operation, because the
    /// COM host serves requests serially and a wedged-but-unwatched call would otherwise
    /// block every later one.
    /// </para>
    /// </summary>
    public static class ComHostRequestContext
    {
        private static readonly AsyncLocal<CancellationToken> CurrentToken = new AsyncLocal<CancellationToken>();
        private static readonly AsyncLocal<long?> CurrentDeadlineOverride = new AsyncLocal<long?>();
        private static readonly AsyncLocal<AggregateBudget?> CurrentAggregate = new AsyncLocal<AggregateBudget?>();
        private static readonly AsyncLocal<bool> CurrentConnectFloorOptIn = new AsyncLocal<bool>();

        /// <summary>The cancellation token of the MCP request being served, or None.</summary>
        public static CancellationToken Token => CurrentToken.Value;

        /// <summary>An explicit per-call deadline in milliseconds, or null for the policy default.</summary>
        public static long? DeadlineOverrideMilliseconds => CurrentDeadlineOverride.Value;

        /// <summary>
        /// What is left of the enclosing gateway operation's aggregate budget, or null when
        /// no enclosing operation declared one. Can go negative-in-spirit: the value is
        /// clamped at zero and the proxy refuses to dispatch below the dispatch floor.
        /// </summary>
        public static long? RemainingAggregateMilliseconds => CurrentAggregate.Value?.RemainingMilliseconds;

        /// <summary>
        /// True when this call's explicit budget is for its own WORK and the supervisor may
        /// still apply the cold-start connect floor on top of it.
        /// <para>
        /// An explicit budget normally outranks the floor, and it has to: outlook_health's
        /// 5 s probe was once silently widened to 90 s and health took 200 s+ to answer.
        /// But the freshness sweep's explicit 30 s was suppressing the floor too, so on a
        /// FRESH host the very first search had to fit the COM attach and the whole sweep
        /// into 30 s - on a machine where attaching to a large OST takes longer than that,
        /// the sweep could never succeed, and each attempt killed the host and restarted
        /// Outlook. So the caller says which it means.
        /// </para>
        /// </summary>
        public static bool AllowConnectFloor => CurrentConnectFloorOptIn.Value;

        /// <summary>Establishes the ambient context for the current logical call.</summary>
        public static IDisposable Enter(
            CancellationToken cancellationToken,
            long? deadlineOverrideMilliseconds = null,
            long? aggregateBudgetMilliseconds = null,
            bool allowConnectFloor = false)
        {
            Scope scope = new Scope(
                CurrentToken.Value, CurrentDeadlineOverride.Value, CurrentAggregate.Value, CurrentConnectFloorOptIn.Value);
            CurrentToken.Value = cancellationToken;
            CurrentDeadlineOverride.Value = deadlineOverrideMilliseconds;
            CurrentConnectFloorOptIn.Value = allowConnectFloor;

            // Every Enter fully defines the ambient state, including clearing an aggregate
            // it does not declare - a scope that inherited a stale budget would bound work
            // it knows nothing about.
            CurrentAggregate.Value = aggregateBudgetMilliseconds is > 0
                ? new AggregateBudget(Stopwatch.GetTimestamp(), aggregateBudgetMilliseconds.Value)
                : null;

            return scope;
        }

        /// <summary>An aggregate budget as an absolute start plus a total, so it survives async hops.</summary>
        private readonly struct AggregateBudget
        {
            private readonly long _startTimestamp;
            private readonly long _totalMilliseconds;

            internal AggregateBudget(long startTimestamp, long totalMilliseconds)
            {
                _startTimestamp = startTimestamp;
                _totalMilliseconds = totalMilliseconds;
            }

            internal long RemainingMilliseconds
            {
                get
                {
                    long elapsed = (long)Stopwatch.GetElapsedTime(_startTimestamp).TotalMilliseconds;
                    long remaining = _totalMilliseconds - elapsed;
                    return remaining > 0 ? remaining : 0;
                }
            }
        }

        private sealed class Scope : IDisposable
        {
            private readonly CancellationToken _previousToken;
            private readonly long? _previousDeadline;
            private readonly AggregateBudget? _previousAggregate;
            private readonly bool _previousConnectFloorOptIn;
            private bool _disposed;

            internal Scope(
                CancellationToken previousToken,
                long? previousDeadline,
                AggregateBudget? previousAggregate,
                bool previousConnectFloorOptIn)
            {
                _previousToken = previousToken;
                _previousDeadline = previousDeadline;
                _previousAggregate = previousAggregate;
                _previousConnectFloorOptIn = previousConnectFloorOptIn;
            }

            public void Dispose()
            {
                if (_disposed)
                {
                    return;
                }

                _disposed = true;
                CurrentToken.Value = _previousToken;
                CurrentDeadlineOverride.Value = _previousDeadline;
                CurrentAggregate.Value = _previousAggregate;
                CurrentConnectFloorOptIn.Value = _previousConnectFloorOptIn;
            }
        }
    }
}
