using System.Globalization;
using System.Net;
using System.Net.Sockets;
using System.Text;
using Xunit;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Pins <see cref="LiveMailSink.ProbeGreetings"/>: a sink port must greet in the protocol it is
/// declared for, not merely accept a connection. Pure loopback, no Outlook and no sink - each
/// test stands up tiny listeners that say one scripted greeting.
/// </summary>
public class LiveMailSinkGreetingTests
{
    [Fact]
    public void Greetings_PassWhenEachPortSpeaksItsOwnProtocol()
    {
        using var smtp = new GreetingListener("220 inbucket Inbucket SMTP ready\r\n");
        using var pop3 = new GreetingListener("+OK Inbucket POP3 server ready <1.2@inbucket>\r\n");

        (bool ready, string message) = LiveMailSink.ProbeGreetings(new MailSinkSettings
        {
            SubmitPort = smtp.Port,
            RetrievePort = pop3.Port,
        });

        Assert.True(ready, message);
        Assert.Contains("greets as SMTP", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Greetings_FollowAMultiLineSmtpGreetingToItsLastLine()
    {
        using var smtp = new GreetingListener("220-sink.vm.invalid first line\r\n220 second line\r\n");
        using var pop3 = new GreetingListener("+OK ready\r\n");

        (bool ready, string message) = LiveMailSink.ProbeGreetings(new MailSinkSettings
        {
            SubmitPort = smtp.Port,
            RetrievePort = pop3.Port,
        });

        Assert.True(ready, message);
    }

    [Fact]
    public void Greetings_NameBothHalvesWhenThePortsAreSwapped()
    {
        // Both ports answer a connect, so Probe passes this configuration. Only the greeting
        // shows that submission and retrieval point at each other's listener.
        using var smtp = new GreetingListener("220 smtp here\r\n");
        using var pop3 = new GreetingListener("+OK pop3 here\r\n");

        (bool ready, string message) = LiveMailSink.ProbeGreetings(new MailSinkSettings
        {
            SubmitPort = pop3.Port,
            RetrievePort = smtp.Port,
            ConnectTimeoutMs = 2000,
        });

        Assert.False(ready);
        Assert.Contains("submission port 127.0.0.1:" + pop3.Port.ToString(CultureInfo.InvariantCulture), message, StringComparison.Ordinal);
        Assert.Contains("'+OK pop3 here'", message, StringComparison.Ordinal);
        Assert.Contains("retrieval port 127.0.0.1:" + smtp.Port.ToString(CultureInfo.InvariantCulture), message, StringComparison.Ordinal);
        Assert.Contains("'220 smtp here'", message, StringComparison.Ordinal);
        Assert.Contains("Outbox", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Greetings_RefuseAListenerThatSaysNothing()
    {
        using var smtp = new GreetingListener("220 smtp here\r\n");
        var silent = new TcpListener(IPAddress.Loopback, 0);
        silent.Start();
        try
        {
            int silentPort = ((IPEndPoint)silent.LocalEndpoint).Port;
            (bool ready, string message) = LiveMailSink.ProbeGreetings(new MailSinkSettings
            {
                SubmitPort = smtp.Port,
                RetrievePort = silentPort,
                ConnectTimeoutMs = 750,
            });

            Assert.False(ready);
            Assert.Contains(silentPort.ToString(CultureInfo.InvariantCulture), message, StringComparison.Ordinal);
            Assert.DoesNotContain("submission port", message, StringComparison.Ordinal);
        }
        finally
        {
            silent.Stop();
        }
    }

    [Fact]
    public void Greetings_RefuseAnSmtpServerThatTurnsClientsAway()
    {
        using var smtp = new GreetingListener("554 no service here\r\n");
        using var pop3 = new GreetingListener("+OK\r\n");

        (bool ready, string message) = LiveMailSink.ProbeGreetings(new MailSinkSettings
        {
            SubmitPort = smtp.Port,
            RetrievePort = pop3.Port,
        });

        Assert.False(ready);
        Assert.Contains("'554 no service here'", message, StringComparison.Ordinal);
        Assert.Contains("SMTP greeting (220)", message, StringComparison.Ordinal);
    }

    [Fact]
    public void GreetingShapes_AreJudgedByTheirProtocolsRules()
    {
        Assert.True(LiveMailSink.IsSmtpGreeting(["220 ready"]));
        Assert.True(LiveMailSink.IsSmtpGreeting(["220"]));
        Assert.True(LiveMailSink.IsSmtpGreeting(["220-a", "220 b"]));
        Assert.False(LiveMailSink.IsSmtpGreeting(["220-a"]));
        Assert.False(LiveMailSink.IsSmtpGreeting(["554 go away"]));
        Assert.False(LiveMailSink.IsSmtpGreeting(["+OK"]));
        Assert.False(LiveMailSink.IsSmtpGreeting([]));

        Assert.True(LiveMailSink.IsPop3Greeting(["+OK ready"]));
        Assert.False(LiveMailSink.IsPop3Greeting(["-ERR busy"]));
        Assert.False(LiveMailSink.IsPop3Greeting(["220 smtp"]));
        Assert.False(LiveMailSink.IsPop3Greeting([]));
    }

    /// <summary>
    /// A loopback listener that writes one scripted greeting to every connection, answers a
    /// line from the client (the probe's QUIT) and hangs up.
    /// </summary>
    private sealed class GreetingListener : IDisposable
    {
        private readonly TcpListener _listener;
        private readonly byte[] _greeting;
        private readonly CancellationTokenSource _stop = new();
        private readonly Task _serve;

        internal GreetingListener(string greeting)
        {
            _greeting = Encoding.ASCII.GetBytes(greeting);
            _listener = new TcpListener(IPAddress.Loopback, 0);
            _listener.Start();
            Port = ((IPEndPoint)_listener.LocalEndpoint).Port;
            _serve = Task.Run(ServeAsync);
        }

        internal int Port { get; }

        public void Dispose()
        {
            _stop.Cancel();
            _listener.Stop();
            try
            {
                _serve.Wait(2000);
            }
            catch (AggregateException)
            {
            }

            _stop.Dispose();
        }

        private async Task ServeAsync()
        {
            while (!_stop.IsCancellationRequested)
            {
                TcpClient client;
                try
                {
                    client = await _listener.AcceptTcpClientAsync(_stop.Token).ConfigureAwait(false);
                }
                catch (Exception ex) when (ex is OperationCanceledException or SocketException or ObjectDisposedException)
                {
                    return;
                }

                using (client)
                {
                    try
                    {
                        NetworkStream stream = client.GetStream();
                        await stream.WriteAsync(_greeting, _stop.Token).ConfigureAwait(false);
                        using var read = CancellationTokenSource.CreateLinkedTokenSource(_stop.Token);
                        read.CancelAfter(2000);
                        byte[] buffer = new byte[256];
                        int n = await stream.ReadAsync(buffer, read.Token).ConfigureAwait(false);
                        if (n > 0)
                        {
                            await stream.WriteAsync(Encoding.ASCII.GetBytes("+OK bye\r\n"), _stop.Token).ConfigureAwait(false);
                        }
                    }
                    catch (Exception ex) when (ex is OperationCanceledException or IOException or SocketException or ObjectDisposedException)
                    {
                    }
                }
            }
        }
    }
}
