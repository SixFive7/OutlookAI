using System;
using System.Diagnostics;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// UTC "now" for measuring how long ago something in THIS PROCESS happened - a clock that
    /// only moves forward, expressed as a <see cref="DateTime"/> so it drops into the existing
    /// TTL comparisons and the existing injectable-clock seams unchanged.
    /// <para>
    /// <b>Which clock to use where.</b> Two different jobs wear the same type in this codebase
    /// and only one of them is safe on the wall clock:
    /// </para>
    /// <list type="bullet">
    /// <item><description>
    /// <b>Measuring an elapsed duration</b> - a cache TTL, a token lifetime, "how old is this
    /// entry". Both ends are readings this process took, and only their DIFFERENCE is meant.
    /// Use this class. <see cref="DateTime.UtcNow"/> is wrong here because it can jump: a
    /// backwards jump (an NTP correction, a user changing the clock, a VM resuming from a
    /// snapshot) extends every live TTL by the size of the jump, and a forwards jump expires
    /// them all at once.
    /// </description></item>
    /// <item><description>
    /// <b>Recording or comparing an absolute instant</b> - a log timestamp, a backup filename,
    /// a DASL date filter, the age of a file on disk, anything compared against a
    /// <c>DateReceived</c> from Outlook or a frontier from the index. The other side of those
    /// comparisons is real calendar time this process did not observe, so it must be real
    /// calendar time on this side too. Use <see cref="DateTime.UtcNow"/>, not this.
    /// </description></item>
    /// </list>
    /// <para>
    /// The value is a wall-clock anchor taken once, plus monotonic elapsed time since. So it
    /// is a real UTC instant on the way in and stays legible in a debugger, but it will drift
    /// from the system clock over a long-lived process by exactly the corrections the system
    /// clock accepted - which is the point, and is why it must never be used for the second
    /// job above.
    /// </para>
    /// <para>
    /// <see cref="Stopwatch"/> rather than <c>Environment.TickCount64</c> because Core also
    /// targets net48, which does not have it (the same constraint <c>PumpedStaRunner</c>
    /// records). The two fields are read in textual order - static initialisers run top to
    /// bottom - so the stopwatch starts a fraction after the anchor is read; that skews the
    /// reading earlier by well under a microsecond, and in the direction that lets an entry
    /// live a hair longer rather than expiring it early.
    /// </para>
    /// </summary>
    public static class MonotonicClock
    {
        private static readonly DateTime AnchorUtc = DateTime.UtcNow;
        private static readonly Stopwatch SinceAnchor = Stopwatch.StartNew();

        /// <summary>
        /// The anchor instant plus monotonic elapsed time since process start. Only
        /// DIFFERENCES between two readings of this are meaningful - see the class remarks
        /// before using it anywhere an absolute instant is wanted.
        /// </summary>
        public static DateTime UtcNow => AnchorUtc + SinceAnchor.Elapsed;
    }
}
