namespace OutlookAI.Core.Com
{
    /// <summary>
    /// What the service layer can say about how Outlook is being reached, without knowing
    /// whether that is in-process or through the killable COM host.
    /// <para>
    /// Lives in Core, and carries only primitives, so <c>outlook_health</c> can report the
    /// supervision state without Core taking a dependency on the host assembly.
    /// </para>
    /// </summary>
    public sealed class ComHostDiagnostics
    {
        /// <summary>Creates the report.</summary>
        public ComHostDiagnostics(
            string mode,
            string state,
            int? processId = null,
            int restartCount = 0,
            string? lastFailure = null,
            string? injectedFault = null,
            bool unresponsive = false,
            int consecutiveTimeouts = 0,
            long? largestFrameBytes = null,
            long? frameLimitBytes = null,
            int? framesRefusedTooLarge = null)
        {
            Unresponsive = unresponsive;
            ConsecutiveTimeouts = consecutiveTimeouts;
            Mode = mode;
            State = state;
            ProcessId = processId;
            RestartCount = restartCount;
            LastFailure = lastFailure;
            InjectedFault = injectedFault;
            LargestFrameBytes = largestFrameBytes;
            FrameLimitBytes = frameLimitBytes;
            FramesRefusedTooLarge = framesRefusedTooLarge;
        }

        /// <summary>"child-process" when COM runs in the supervised host, "in-process" otherwise.</summary>
        public string Mode { get; }

        /// <summary>Lifecycle state: none, starting, ready or faulted.</summary>
        public string State { get; }

        /// <summary>PID of the COM host, when one is running.</summary>
        public int? ProcessId { get; }

        /// <summary>
        /// How many times the COM host has been replaced this session. A climbing count is
        /// the signal that Outlook keeps wedging - each restart is one reclaimed hang.
        /// </summary>
        public int RestartCount { get; }

        /// <summary>
        /// The last supervision failure this session, or null when nothing has gone wrong.
        /// Deliberately NOT cleared by a successful restart: a recovered wedge must still
        /// be explainable afterwards, or the recovery hides the fault.
        /// </summary>
        public string? LastFailure { get; }

        /// <summary>
        /// True while requests needing Outlook are being refused immediately because it
        /// has repeatedly failed to answer. Self-clearing: Outlook is re-probed after a
        /// cooldown and any success closes it.
        /// </summary>
        public bool Unresponsive { get; }

        /// <summary>Consecutive operation timeouts; 0 once Outlook answers again.</summary>
        public int ConsecutiveTimeouts { get; }

        /// <summary>
        /// Set only when a test fault is configured. Reported so an injected failure can
        /// never be mistaken for a real one while reading a health report.
        /// </summary>
        public string? InjectedFault { get; }

        /// <summary>
        /// The largest single message that has crossed the pipe to the COM host, in bytes,
        /// SINCE THIS SERVER PROCESS STARTED. Null when Outlook is reached in-process,
        /// where there is no pipe and so no such thing.
        /// <para>
        /// This is a measurement, not a problem. It exists because the frame limit beside
        /// it was chosen as "far above any real payload" without anyone measuring what a
        /// real payload weighs. Read the two together: the ratio is the headroom, and it is
        /// the only evidence that says whether the limit is set sensibly.
        /// </para>
        /// <para>
        /// Deliberately NOT reset when the COM host restarts. The child is restartable and
        /// its own counters die with it; the question this answers is about the product,
        /// and the profiles that produce the biggest answers are exactly the ones whose
        /// hosts restart most.
        /// </para>
        /// </summary>
        public long? LargestFrameBytes { get; }

        /// <summary>
        /// The ceiling a single message must fit under. Reported alongside the high-water
        /// mark so the headroom can be read without knowing the constant.
        /// </summary>
        public long? FrameLimitBytes { get; }

        /// <summary>
        /// How many answers were refused this session for exceeding that ceiling. Zero
        /// normally; anything else means a caller was told "too large" instead of getting
        /// its answer, and the request that caused it needs narrowing rather than retrying.
        /// </summary>
        public int? FramesRefusedTooLarge { get; }
    }
}
