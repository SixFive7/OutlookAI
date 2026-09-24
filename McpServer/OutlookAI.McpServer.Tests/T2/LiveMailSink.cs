using System.Globalization;
using System.Net.Sockets;
using System.Text;
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
/// fail its own teardown, forever, on an artifact nothing could remove.
/// </para>
/// <para>
/// <b>"Six live methods additionally need the mail to actually arrive" was wrong - it is 13,
/// corrected 2026-09-15.</b> Six FILES hold an arrival wait (see <c>LiveInboxArrival</c>'s own
/// header, which counts the copies it consolidated); that file count was read as a method count,
/// and three more waits were never in the consolidation at all. Of 127 live methods, however,
/// exactly ONE needs the wire in a way nothing can substitute:
/// <c>T3/Phase5LiveMcpToolShapeTests.SendTool_TwoStepFlow_RoundTrip_OverRealStdio_WithAuditLines</c>,
/// the only test that calls the product's own <c>send</c> tool and lets Outlook submit. The other
/// twelve need an item to appear in the Inbox with a known subject, which direct PST creation
/// produces - as <see cref="LiveOutlookTestMailer"/>'s own attachment helper already does,
/// deliberately, because drafts are indexed exactly like received mail.
/// Weakening the sweep was rejected outright:
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
/// a bad trade. On the test guests the sink is <b>Inbucket 3.1.1</b> (MIT), chosen 2026-09-24
/// because its POP3 accepts a login with no password and shows each login only its own
/// mailbox. It is media, not a dependency of the product or of the build:
/// <c>Testbed/MEDIA.md</c> records it and <c>Testbed/guest/Install-MailSink.ps1</c> installs it
/// and proves a full SMTP-to-POP3 round trip. See <c>Docs/live-tier-on-the-vm.md</c> section 2.7.
/// </para>
/// <para>
/// <b>What is checked here.</b> Three things, all cheap and all decisive. The listeners
/// answer a TCP connect - a sink that is not running is the single likeliest cause of a
/// send-path failure, and it is indistinguishable from a code fault once the mail is in the
/// Outbox. Each listener then GREETS in the protocol its port is declared for - an SMTP
/// <c>220</c> on submission, a POP3 <c>+OK</c> on retrieval - because anything that binds a
/// port passes a connect, including some other program, or the two halves configured the wrong
/// way round, and both of those fail later as a two-minute arrival timeout that names nothing.
/// And the Outbox is EMPTY before anything runs - because if delivery is not really happening,
/// that is where the evidence accumulates, and starting a run on top of it means the teardown
/// sweep will blame this run for the last one's residue.
/// </para>
/// <para>
/// <b>What is deliberately NOT checked here: a round trip.</b> Proving that the sink hands back
/// what it was given means submitting a message, and this suite cannot see which sink it is
/// talking to or how its mailboxes are arranged - on a sink that shows every login every
/// message, a probe would be downloaded into the hub store by the next send/receive. So the
/// suite asks only questions with no side effects, and the round trip is proved where the sink's
/// arrangement is known: by the installer's <c>-Verify</c>, into mailboxes no account reads.
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

            (bool speaking, string greetings) = ProbeGreetings(sink);
            Console.WriteLine("[sink] " + greetings);
            if (!speaking)
            {
                throw new InvalidOperationException(greetings);
            }

            _checked = true;
        }
    }

    /// <summary>
    /// Reads each listener's greeting and checks it is the protocol that port is declared for:
    /// an SMTP <c>220</c> on submission, a POP3 <c>+OK</c> on retrieval. Nothing is submitted
    /// and nothing is logged in to - each side is sent <c>QUIT</c> straight after its greeting -
    /// so this has no side effect on any sink, whatever its mailboxes look like.
    /// <para>
    /// Separate from <see cref="Probe"/>, which only connects: a connect is what anything that
    /// binds the port passes, including another program or the two halves swapped. BOTH halves
    /// are read even when the first is wrong, for the reason <see cref="Probe"/> gives.
    /// </para>
    /// </summary>
    internal static (bool Ready, string Message) ProbeGreetings(MailSinkSettings sink)
    {
        ArgumentNullException.ThrowIfNull(sink);
        string? submit = ProbeGreeting(
            sink.SubmitHost, sink.SubmitPort, sink.ConnectTimeoutMs, "submission", "an SMTP greeting (220)", IsSmtpGreeting);
        string? retrieve = ProbeGreeting(
            sink.RetrieveHost, sink.RetrievePort, sink.ConnectTimeoutMs, "retrieval", "a POP3 greeting (+OK)", IsPop3Greeting);
        if (submit == null && retrieve == null)
        {
            return (true, string.Format(
                CultureInfo.InvariantCulture,
                "submission {0}:{1} greets as SMTP and retrieval {2}:{3} as POP3.",
                sink.SubmitHost,
                sink.SubmitPort,
                sink.RetrieveHost,
                sink.RetrievePort));
        }

        return (false,
            "The local mail sink's ports answer, but not with the protocols the 'mailSink' block declares, so "
            + "a send would reach the wrong program or none and sit in the Outbox."
            + (submit ?? string.Empty) + (retrieve ?? string.Empty)
            + " Check what holds those ports and that submitPort and retrievePort are not swapped; on a test "
            + "guest, Testbed/guest/Install-MailSink.ps1 -Verify says which. See the mail-sink section of "
            + "Docs/live-tier-on-the-vm.md.");
    }

    /// <summary>An SMTP greeting is a 220 reply; a multi-line one ends on its <c>220 </c> line.</summary>
    internal static bool IsSmtpGreeting(IReadOnlyList<string> lines)
    {
        ArgumentNullException.ThrowIfNull(lines);
        return lines.Count > 0
            && lines[^1].StartsWith("220", StringComparison.Ordinal)
            && (lines[^1].Length == 3 || lines[^1][3] == ' ');
    }

    /// <summary>A POP3 greeting is a single <c>+OK</c> line (RFC 1939 section 4).</summary>
    internal static bool IsPop3Greeting(IReadOnlyList<string> lines)
    {
        ArgumentNullException.ThrowIfNull(lines);
        return lines.Count == 1 && lines[0].StartsWith("+OK", StringComparison.Ordinal);
    }

    /// <summary>
    /// Connects, reads the greeting - following an SMTP continuation (<c>220-</c>) to its last
    /// line - says <c>QUIT</c>, and returns null when the greeting is the expected one, or a
    /// sentence naming what arrived instead. Never throws: the caller reports both halves.
    /// </summary>
    private static string? ProbeGreeting(
        string host, int port, int timeoutMs, string half, string expected, Func<IReadOnlyList<string>, bool> accept)
    {
        try
        {
            using var client = new TcpClient();
            if (!client.ConnectAsync(host, port).Wait(timeoutMs))
            {
                return string.Format(
                    CultureInfo.InvariantCulture, " The {0} port {1}:{2} did not accept a connection within {3} ms.", half, host, port, timeoutMs);
            }

            client.ReceiveTimeout = timeoutMs;
            client.SendTimeout = timeoutMs;
            using NetworkStream stream = client.GetStream();
            using var reader = new StreamReader(stream, Encoding.Latin1);
            using var writer = new StreamWriter(stream, Encoding.Latin1) { NewLine = "\r\n", AutoFlush = true };

            var lines = new List<string>();
            for (int i = 0; i < 10; i++)
            {
                string? line = reader.ReadLine();
                if (line == null)
                {
                    break;
                }

                lines.Add(line);
                bool smtpContinuation = line.Length > 3 && line[3] == '-' && char.IsAsciiDigit(line[0]);
                if (!smtpContinuation)
                {
                    break;
                }
            }

            if (accept(lines))
            {
                // Polite, so the sink's own log records a clean close rather than a dropped
                // client. Best-effort: the verdict is already made.
                try
                {
                    writer.WriteLine("QUIT");
                    _ = reader.ReadLine();
                }
                catch (IOException)
                {
                }

                return null;
            }

            string first = lines.Count == 0 ? "nothing at all" : "'" + Printable(lines[0]) + "'";
            return string.Format(
                CultureInfo.InvariantCulture,
                " The {0} port {1}:{2} answered {3} where {4} was expected.",
                half,
                host,
                port,
                first,
                expected);
        }
        catch (Exception ex) when (ex is SocketException or AggregateException or ObjectDisposedException or IOException)
        {
            return string.Format(
                CultureInfo.InvariantCulture,
                " Reading the {0} greeting from {1}:{2} failed: {3}.",
                half,
                host,
                port,
                ex.GetBaseException().Message);
        }
    }

    /// <summary>At most 80 characters, control characters shown as '?', for a message a human reads.</summary>
    private static string Printable(string line)
    {
        var chars = new StringBuilder();
        foreach (char c in line.Length > 80 ? line[..80] : line)
        {
            chars.Append(char.IsControl(c) ? '?' : c);
        }

        return chars.ToString();
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
