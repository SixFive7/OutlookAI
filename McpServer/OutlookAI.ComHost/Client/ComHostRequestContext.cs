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

        /// <summary>The cancellation token of the MCP request being served, or None.</summary>
        public static CancellationToken Token => CurrentToken.Value;

        /// <summary>An explicit per-call deadline in milliseconds, or null for the policy default.</summary>
        public static long? DeadlineOverrideMilliseconds => CurrentDeadlineOverride.Value;

        /// <summary>Establishes the ambient context for the current logical call.</summary>
        public static IDisposable Enter(CancellationToken cancellationToken, long? deadlineOverrideMilliseconds = null)
        {
            CancellationToken previousToken = CurrentToken.Value;
            long? previousDeadline = CurrentDeadlineOverride.Value;
            CurrentToken.Value = cancellationToken;
            CurrentDeadlineOverride.Value = deadlineOverrideMilliseconds;
            return new Scope(previousToken, previousDeadline);
        }

        private sealed class Scope : IDisposable
        {
            private readonly CancellationToken _previousToken;
            private readonly long? _previousDeadline;
            private bool _disposed;

            internal Scope(CancellationToken previousToken, long? previousDeadline)
            {
                _previousToken = previousToken;
                _previousDeadline = previousDeadline;
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
            }
        }
    }
}
