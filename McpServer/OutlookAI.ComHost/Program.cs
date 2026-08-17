using System.Diagnostics;
using System.IO.Pipes;
using OutlookAI.ComHost.Host;
using OutlookAI.ComHost.Protocol;
using OutlookAI.Core.Com;

namespace OutlookAI.ComHost
{
    /// <summary>
    /// The COM child process.
    /// <para>
    /// This process exists to be killable. Everything that can block on Outlook lives
    /// here - the pumped STA thread, the Application object, the pin Explorer, every
    /// late-bound call - so that when Outlook stops answering, the parent can reclaim
    /// the wedged thread and its COM references the only way Windows allows: by ending
    /// the process that holds them. An in-process timeout cannot do this. A blocked
    /// outbound COM call is not cancellable, and even releasing its RCWs marshals back
    /// into the same wedged apartment.
    /// </para>
    /// <para>
    /// Consequently this process deliberately has NO internal timeouts on COM work. Its
    /// only contract is: serve the pipe, and die quietly when told. See "Why two processes"
    /// in McpServer/README.md.
    /// </para>
    /// </summary>
    internal static class Program
    {
        private static async Task<int> Main(string[] args)
        {
            string? pipeName = Environment.GetEnvironmentVariable(ComHostProtocol.PipeNameVariable);
            if (string.IsNullOrWhiteSpace(pipeName))
            {
                await Console.Error
                    .WriteLineAsync($"{ComHostProtocol.PipeNameVariable} is not set. This process is started by OutlookAI.McpServer, not directly.")
                    .ConfigureAwait(false);
                return 2;
            }

            bool allowStartingOutlook = !args.Contains("--no-autostart", StringComparer.OrdinalIgnoreCase);

            using CancellationTokenSource shutdown = new CancellationTokenSource();
            WatchParent(shutdown);

            using NamedPipeClientStream pipe = new NamedPipeClientStream(
                ".", pipeName, PipeDirection.InOut, PipeOptions.Asynchronous);

            try
            {
                await pipe.ConnectAsync(ConnectTimeoutMs, shutdown.Token).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                await Console.Error.WriteLineAsync($"Could not connect to the supervising pipe: {ex.Message}").ConfigureAwait(false);
                return 3;
            }

            using ComGateway gateway = new ComGateway(allowStartingOutlook);
            IOutlookSession session = GatewayRoutingProxy.Create(gateway);
            ComHostServer server = new ComHostServer(pipe, session, typeof(IOutlookSession));

            // Tell the parent when Outlook goes away, so it can drop cached state without
            // waiting to discover it on the next call. Advisory: the parent re-probes
            // anyway, this only makes it prompt.
            gateway.OutlookGone += () => _ = server.SendEventAsync(ComHostEvents.OutlookGone);

            try
            {
                await server.ServeAsync(shutdown.Token).ConfigureAwait(false);
                return 0;
            }
            catch (OperationCanceledException)
            {
                return 0;
            }
            catch (Exception ex)
            {
                await Console.Error.WriteLineAsync($"COM host terminated: {ex}").ConfigureAwait(false);
                return 1;
            }
        }

        /// <summary>
        /// The child's half of the pipe handshake. Shared with the parent's
        /// (<c>ComHostPolicy.HandshakeBudgetMilliseconds</c>) rather than declared as a
        /// second 30 s: it is ONE handshake, and two ends of one protocol each owning their
        /// own literal is how the two halves drift apart. The parent may use LESS than this
        /// when the triggering operation's deadline is shorter - it is the side that tears
        /// down, so the child only needs the ceiling.
        /// </summary>
        private const int ConnectTimeoutMs = ComOperationBudgets.HandshakeBudgetMs;

        /// <summary>
        /// Exits if the parent disappears. The job object already covers the normal case;
        /// this covers the rest - a parent killed in a way that never closes its job
        /// handle, or a job that could not be created at all. An orphaned COM host is
        /// exactly the leak this architecture exists to stop, so it is worth two guards.
        /// </summary>
        private static void WatchParent(CancellationTokenSource shutdown)
        {
            string? raw = Environment.GetEnvironmentVariable(ComHostProtocol.ParentPidVariable);
            if (!int.TryParse(raw, out int parentPid))
            {
                return;
            }

            try
            {
                Process parent = Process.GetProcessById(parentPid);
                parent.EnableRaisingEvents = true;
                parent.Exited += (_, _) =>
                {
                    try
                    {
                        shutdown.Cancel();
                    }
                    catch (ObjectDisposedException)
                    {
                        // Already shutting down.
                    }
                };

                if (parent.HasExited)
                {
                    shutdown.Cancel();
                }
            }
            catch (ArgumentException)
            {
                // Parent already gone.
                shutdown.Cancel();
            }
            catch (InvalidOperationException)
            {
                shutdown.Cancel();
            }
        }
    }
}
