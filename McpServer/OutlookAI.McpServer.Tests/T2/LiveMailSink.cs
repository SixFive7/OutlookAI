using System.Globalization;
using System.Net.Sockets;
using OutlookAI.Core.Com;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Coordinates of the local mail sink a test machine uses to make sends REAL, when it has
/// one. Absent means the profile has genuine transport - which is the maintainer's own
/// machine - and every check here is then a no-op.
/// </summary>
public sealed class MailSinkSettings
{
    /// <summary>Host the sink accepts submissions on. Loopback, always.</summary>
    public string SubmitHost { get; set; } = "127.0.0.1";

    /// <summary>Port the sink accepts submissions on.</summary>
    public int SubmitPort { get; set; } = 25;

    /// <summary>Host the sink serves delivered mail back from.</summary>
    public string RetrieveHost { get; set; } = "127.0.0.1";

    /// <summary>Port the sink serves delivered mail back from (POP3).</summary>
    public int RetrievePort { get; set; } = 110;

    /// <summary>
    /// How long a reachability probe waits for each listener. Short: the sink is on
    /// loopback, so a connect that is not instant is a connect that is not going to happen,
    /// and a long timeout here would just delay a failure the operator has to fix anyway.
    /// </summary>
    public int ConnectTimeoutMs { get; set; } = 2000;

    /// <summary>Whether both ports are named.</summary>
    public bool IsComplete
        => !string.IsNullOrWhiteSpace(SubmitHost) && SubmitPort > 0
            && !string.IsNullOrWhiteSpace(RetrieveHost) && RetrievePort > 0;
}

/// <summary>
/// Proves the local mail sink is there before the live tier sends anything into it.
/// <para>
/// <b>Why a sink at all, and why it must deliver BACK.</b> The test VM's dummy account
/// pointed at an unroutable server. A send therefore QUEUES and never leaves - and the
/// Outbox is in the mandatory zero-artifact sweep, so every run that sent anything would
/// fail its own teardown, forever, on an artifact nothing could remove. Six live methods
/// additionally need the mail to actually arrive. Weakening the sweep was rejected outright:
/// that guard exists because real mail was once destroyed, and the Outbox is the folder that
/// most reliably catches a genuine send-path leak.
/// </para>
/// <para>
/// <b>The sink is NOT in this repository, and that is a decision rather than an omission.</b>
/// A loopback SMTP-plus-POP3 server is a few hundred lines of RFC 1939, and its failure
/// modes - dot-stuffing a body line that starts with a period, UIDL identities that move when
/// the store is recreated, STAT octet counts - all produce INTERMITTENT wrong answers against
/// Outlook, which is the fussiest POP3 client there is. This suite's whole design is the
/// elimination of intermittent artifacts; writing a new source of them to serve it would be
/// a bad trade. A maintained, permissively licensed component that already does exactly this
/// is used instead, it is not a dependency of the product or of the build, and the runbook
/// says which one and how to install it. See <c>Docs/live-tier-on-the-vm.md</c>.
/// </para>
/// <para>
/// <b>What is checked here.</b> Two things, both cheap and both decisive. The listeners
/// answer a TCP connect - a sink that is not running is the single likeliest cause of a
/// send-path failure, and it is indistinguishable from a code fault once the mail is in the
/// Outbox. And the Outbox is EMPTY before anything runs - because if delivery is not really
/// happening, that is where the evidence accumulates, and starting a run on top of it means
/// the teardown sweep will blame this run for the last one's residue.
/// </para>
/// </summary>
public static class LiveMailSink
{
    private static readonly object Gate = new();
    private static bool _checked;
    private static volatile bool _nudgeWhileWaiting;

    /// <summary>
    /// Whether an arrival wait should keep asking Outlook to deliver. True only on a machine
    /// whose settings declare a sink - a profile with real transport needs no prompting and
    /// should not be prodded on a timer by a test suite.
    /// <para>
    /// Process-wide rather than per test because it is a fact about the MACHINE, learned once
    /// from the settings file, and the wait helper is reached from six call sites that have
    /// no settings of their own.
    /// </para>
    /// </summary>
    internal static bool NudgeWhileWaiting => _nudgeWhileWaiting;

    /// <summary>
    /// Proves the sink is reachable, or throws naming the repair. A no-op on a machine whose
    /// settings declare no sink, and computed once per process.
    /// </summary>
    public static void EnsureReachable(LiveTestSettings settings)
    {
        ArgumentNullException.ThrowIfNull(settings);
        lock (Gate)
        {
            if (_checked)
            {
                return;
            }

            MailSinkSettings? sink = settings.MailSink;
            if (sink == null)
            {
                _checked = true;
                return;
            }

            _nudgeWhileWaiting = true;

            (bool reachable, string message) = Probe(sink);
            Console.WriteLine("[sink] " + message);
            if (!reachable)
            {
                throw new InvalidOperationException(message);
            }

            _checked = true;
        }
    }

    /// <summary>
    /// Connects to both listeners and says what it found. Separate from
    /// <see cref="EnsureReachable"/>, which latches once per process, so the decision itself
    /// can be exercised repeatedly against listeners a test starts and stops.
    /// <para>
    /// BOTH halves are probed even when the first fails. A sink that accepts submissions and
    /// cannot hand them back is the failure this whole design exists to avoid, and reporting
    /// only the first fault would send an operator to fix the half that was already working.
    /// </para>
    /// </summary>
    internal static (bool Reachable, string Message) Probe(MailSinkSettings sink)
    {
        ArgumentNullException.ThrowIfNull(sink);
        string? submit = ProbeListener(sink.SubmitHost, sink.SubmitPort, sink.ConnectTimeoutMs);
        string? retrieve = ProbeListener(sink.RetrieveHost, sink.RetrievePort, sink.ConnectTimeoutMs);
        if (submit == null && retrieve == null)
        {
            return (true, string.Format(
                CultureInfo.InvariantCulture,
                "submission {0}:{1} and retrieval {2}:{3} both answering.",
                sink.SubmitHost,
                sink.SubmitPort,
                sink.RetrieveHost,
                sink.RetrievePort));
        }

        return (false,
            "The local mail sink is not answering, so nothing this run sends can be delivered and every send would "
            + "leave a permanent artifact in the Outbox that the zero-artifact sweep then fails on."
            + (submit ?? string.Empty) + (retrieve ?? string.Empty)
            + " Start the sink service and re-run; see the mail-sink section of "
            + "Docs/live-tier-on-the-vm.md. If this machine has real transport and needs no sink, remove the "
            + "'mailSink' block from the live-test settings.");
    }

    /// <summary>
    /// Fails the run when the Outbox already holds mail. Separate from
    /// <see cref="EnsureReachable"/> because it needs a COM session and that one deliberately
    /// does not - reachability has to be answerable before Outlook is started.
    /// <para>
    /// It reports a COUNT and never a subject: the Outbox of a production profile is the
    /// user's own unsent mail.
    /// </para>
    /// </summary>
    public static void EnsureOutboxDrained(OutlookComSession session, LiveTestSettings settings)
    {
        ArgumentNullException.ThrowIfNull(session);
        ArgumentNullException.ThrowIfNull(settings);
        if (settings.MailSink == null)
        {
            return;
        }

        // Profile-wide, and -1 means the walk itself failed. Unknown is treated as unsafe,
        // which is the rule the quit-when-safe check beside it already keeps: the whole
        // point of the number is to decide whether it is safe to proceed, and a number
        // nobody could read decides nothing.
        int queued = session.CountOutboxItems();
        if (queued < 0)
        {
            throw new InvalidOperationException(
                "The Outbox item count could not be read, so this run cannot tell an empty Outbox from one holding "
                + "a previous run's undelivered mail. Unknown is unsafe here: if delivery is not happening, the "
                + "zero-artifact sweep at the end will blame this run for the last one's residue.");
        }

        if (queued > 0)
        {
            throw new InvalidOperationException(string.Format(
                CultureInfo.InvariantCulture,
                "The profile's Outbox already holds {0} item(s) before this run started. On a machine with a local "
                + "sink that is the signature of delivery not happening: sends queue, nothing drains them, and the "
                + "residue outlives every teardown. Fix the sink and clear the Outbox before running the tier - "
                + "otherwise this run's zero-artifact sweep will fail on the last run's mail.",
                queued));
        }
    }

    /// <summary>
    /// Asks Outlook to flush the Outbox and fetch what the sink has for it. Best-effort by
    /// design: a profile may refuse, and the arrival wait remains the authority either way.
    /// A no-op unless this machine declares a sink.
    /// <para>
    /// It has to be re-issued WHILE waiting rather than fired once after <c>Send()</c>.
    /// Microsoft documents <c>SendAndReceive</c> as asynchronous with no completion signal,
    /// so a single call can perfectly well finish its fetch BEFORE the submission it
    /// triggered reaches the sink - after which nothing asks again until Outlook's own
    /// schedule comes round, which defaults to thirty minutes and is far outside any
    /// deadline this suite keeps.
    /// </para>
    /// </summary>
    internal static void NudgeDelivery()
    {
        if (!_nudgeWhileWaiting)
        {
            return;
        }

        LiveOutlookTestMailer.RequestDelivery();
    }

    /// <summary>
    /// Connects and disconnects. Returns null when the listener answered, or a sentence
    /// naming what did not - never an exception, because the caller wants to report BOTH
    /// halves rather than the first one to fail.
    /// </summary>
    private static string? ProbeListener(string host, int port, int timeoutMs)
    {
        try
        {
            using var client = new TcpClient();
            if (!client.ConnectAsync(host, port).Wait(timeoutMs))
            {
                return string.Format(
                    CultureInfo.InvariantCulture,
                    " Nothing answered {0}:{1} within {2} ms.",
                    host,
                    port,
                    timeoutMs);
            }

            return null;
        }
        catch (Exception ex) when (ex is SocketException or AggregateException or ObjectDisposedException)
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                " Connecting to {0}:{1} failed: {2}.",
                host,
                port,
                ex.GetBaseException().Message);
        }
    }
}
