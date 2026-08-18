using System.Reflection;
using System.Runtime.ExceptionServices;
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
        private IComGateway _gateway = null!;

        /// <summary>Creates a proxy routing <see cref="IOutlookSession"/> calls through <paramref name="gateway"/>.</summary>
        public static IOutlookSession Create(IComGateway gateway)
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

            // Test-only, and a no-op unless OUTLOOKAI_COMHOST_FAULT names a session fault.
            // It stands in for the session ITSELF rather than for Outlook, which is what
            // makes the child-side error path reachable on a machine with no Outlook.
            if (ComHostFaultInjection.TrySessionFault(targetMethod.Name, out IOutlookSession? faulted))
            {
                return InvokeSession(targetMethod, faulted!, args);
            }

            return _gateway.Run<object?>(session => InvokeSession(targetMethod, session, args));
        }

        /// <summary>
        /// Calls one contract method on the session and rethrows what the SESSION threw,
        /// never the reflection wrapper around it.
        /// <para>
        /// <see cref="MethodBase.Invoke(object, object[])"/> wraps whatever the target
        /// throws in a <see cref="TargetInvocationException"/> whose own message is the
        /// content-free "Exception has been thrown by the target of an invocation." Letting
        /// that wrapper escape cost two distinct things, and both were live defects:
        /// </para>
        /// <para>
        /// 1. <c>ComGateway.Run</c> keys its one-shot rebuild on the exception TYPE
        /// (RPC_E_DISCONNECTED and friends). Every call arrived wrapped, so the filter never
        /// matched and the disconnect recovery was dead code in this process.
        /// </para>
        /// <para>
        /// 2. <c>ComHostServer</c> reflects too, so the wrapper was wrapped a second time.
        /// It peels one layer, put the INNER wrapper on the wire, and every deliberate error
        /// the session raises - "Folder 'X' was not found in store 'Y' ... use list_folders
        /// for paths" among them - reached the agent as that one useless sentence. Nothing
        /// noticed until a probe happened to hit it on 2026-08-18.
        /// </para>
        /// <para>
        /// <see cref="ExceptionDispatchInfo"/> rather than <c>throw ex.InnerException</c>:
        /// the latter resets the stack to this line and discards every frame inside the
        /// session, which is the only record of where the failure actually happened.
        /// </para>
        /// </summary>
        private static object? InvokeSession(MethodInfo targetMethod, IOutlookSession session, object?[]? args)
        {
            try
            {
                return targetMethod.Invoke(session, args);
            }
            catch (TargetInvocationException ex) when (ex.InnerException != null)
            {
                ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
                throw; // Unreachable - Throw() does not return, but the compiler cannot know.
            }
        }
    }
}
