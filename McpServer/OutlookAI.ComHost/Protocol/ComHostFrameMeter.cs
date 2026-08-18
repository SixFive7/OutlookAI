namespace OutlookAI.ComHost.Protocol
{
    /// <summary>
    /// Records how big the frames crossing this pipe actually get.
    /// <para>
    /// The point is to answer a question nobody could answer before: is
    /// <see cref="ComHostProtocol.MaxFrameBytes"/> the right number? The limit was chosen
    /// as "far above any real payload", but the largest frame the product actually
    /// produces had never been measured, so the headroom was an assumption rather than an
    /// observation. A high-water mark reported through <c>outlook_health</c> turns it into
    /// evidence, at the cost of one volatile read per frame - and an interlocked compare
    /// only on the rare frame that sets a new peak.
    /// </para>
    /// <para>
    /// Lifetime is ONE PROCESS, and for the MCP server that means one server session. It
    /// deliberately survives COM host restarts: the child is restartable and its counters
    /// die with it, but "how big do real answers get" is a question about the product, not
    /// about one child, and resetting it on every restart would erase exactly the evidence
    /// gathered by the degraded profiles most likely to produce a big frame.
    /// </para>
    /// <para>
    /// Each process counts what it saw. In the MCP server that is every request it encoded,
    /// every answer it read back, and every refusal either end reported. In the COM host it
    /// is the mirror image, read by nothing - the child has no health surface of its own.
    /// </para>
    /// </summary>
    internal sealed class ComHostFrameMeter
    {
        /// <summary>The per-process instrument. Static because framing is, and because the counter must outlive any one connection.</summary>
        internal static ComHostFrameMeter Shared { get; } = new ComHostFrameMeter();

        private long _largestFrameBytes;
        private int _framesRefusedTooLarge;

        /// <summary>Payload bytes of the largest single frame seen, excluding the 4-byte length prefix. Zero until one crosses.</summary>
        internal long LargestFrameBytes => Volatile.Read(ref _largestFrameBytes);

        /// <summary>How many frames were refused for exceeding the limit. Non-zero means a caller was told "too large" instead of getting an answer.</summary>
        internal int FramesRefusedTooLarge => Volatile.Read(ref _framesRefusedTooLarge);

        /// <summary>
        /// Notes a frame that was built or received intact. Written as a compare-and-swap
        /// loop rather than a lock because this runs on every single call in both
        /// directions, and once the high-water mark is above the current frame - which is
        /// the overwhelmingly common case - it costs one read and no write at all.
        /// </summary>
        internal void RecordFrame(long payloadBytes)
        {
            long observed = Volatile.Read(ref _largestFrameBytes);
            while (payloadBytes > observed)
            {
                long previous = Interlocked.CompareExchange(ref _largestFrameBytes, payloadBytes, observed);
                if (previous == observed)
                {
                    return;
                }

                observed = previous;
            }
        }

        /// <summary>
        /// Notes a frame that was refused for being over the limit.
        /// <para>
        /// Refused frames are counted but do NOT move the high-water mark. That mark says
        /// how close ordinary traffic came to the limit, and folding an oversized answer
        /// into it would report a size that never crossed the pipe and make every future
        /// reading look closer to the ceiling than it was. The refused size is not lost:
        /// it is named in the error the caller receives.
        /// </para>
        /// </summary>
        internal void RecordRefusal()
        {
            _ = Interlocked.Increment(ref _framesRefusedTooLarge);
        }
    }
}
