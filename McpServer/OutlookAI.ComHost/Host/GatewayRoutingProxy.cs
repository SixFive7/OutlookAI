using System.Reflection;
using OutlookAI.Core.Com;

namespace OutlookAI.ComHost.Host
{
    /// <summary>
    /// Child-side proxy: turns every <see cref="IOutlookSession"/> call into
    /// <c>ComGateway.Run(session =&gt; session.TheMethod(args))</c>.
    /// <para>
    /// A <see cref="DispatchProxy"/> rather than 23 hand-written forwarding methods.
    /// Those methods would be pure ceremony, and every one of them would be a place for
    /// the two ends of the contract to drift apart silently.
    /// </para>
    /// <para>
    /// Routing through the gateway rather than straight at the session is deliberate: it
    /// keeps the existing COM-level recovery - the liveness ping, the one-shot rebuild
    /// when Outlook exited under a live proxy - on the child side where the session
    /// actually lives. The parent supervises the process; the gateway supervises the COM
    /// connection. Neither has to know about the other's failure mode.
    /// </para>
    /// </summary>
    public class GatewayRoutingProxy : DispatchProxy
    {
        private ComGateway _gateway = null!;

        /// <summary>Creates a proxy routing <see cref="IOutlookSession"/> calls through <paramref name="gateway"/>.</summary>
        public static IOutlookSession Create(ComGateway gateway)
        {
            ArgumentNullException.ThrowIfNull(gateway);

            object proxy = Create<IOutlookSession, GatewayRoutingProxy>()
                ?? throw new InvalidOperationException("DispatchProxy.Create returned null.");
            ((GatewayRoutingProxy)proxy)._gateway = gateway;
            return (IOutlookSession)proxy;
        }

        /// <inheritdoc />
        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            if (targetMethod == null)
            {
                throw new ArgumentNullException(nameof(targetMethod));
            }

            return _gateway.Run<object?>(session => targetMethod.Invoke(session, args));
        }
    }
}
