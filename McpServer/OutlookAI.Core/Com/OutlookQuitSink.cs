using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Minimal IDispatch event sink for the Outlook Application source dispinterface,
    /// advised on the pumped STA (v3.MD section 0.5.2 obligation - advise sinks need a
    /// real message pump). Only DISPID 61447 (Quit) is mapped; other event DISPIDs
    /// resolve to DISP_E_MEMBERNOTFOUND, which COM event sources ignore by design.
    ///
    /// Role (SF-2, calibrated by probe 2026-07-23): DEFENSE-IN-DEPTH ONLY. On this
    /// build the Quit event does NOT reach out-of-process sinks when another client
    /// drives Application.Quit (measured: advise succeeds, event never fires, Outlook
    /// parks awaiting the held refs) - the process-exit watcher in OutlookComSession is
    /// the load-bearing release signal. The sink stays advised because it is nearly
    /// free and covers any shutdown path that DOES raise the event. If it ever fires,
    /// the handler must return fast (fast-shutdown discipline), so it only queues the
    /// notification.
    /// </summary>
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public sealed class OutlookQuitSink : IOutlookApplicationEventsSink
    {
        private readonly Action _onQuit;

        /// <summary>Creates the sink; <paramref name="onQuit"/> is queued to the thread pool on Quit.</summary>
        public OutlookQuitSink(Action onQuit)
        {
            _onQuit = onQuit ?? throw new ArgumentNullException(nameof(onQuit));
        }

        /// <summary>Outlook's Quit event (DISPID 61447): arrives on the pumped STA - return fast.</summary>
        public void Quit()
        {
            System.Threading.ThreadPool.QueueUserWorkItem(_ => _onQuit());
        }

        /// <summary>
        /// Advises <paramref name="sink"/> on <paramref name="application"/>'s event
        /// connection point. Tries the version-specific source dispinterfaces newest
        /// first (GUIDs verified against the Outlook PIA on this machine, 2026-07-23;
        /// the Quit DISPID is 61447 in all of them). Returns null when no connection
        /// point is available (caller degrades to process-exit watching only).
        /// </summary>
        public static OutlookQuitSinkRegistration? TryAdvise(object application, OutlookQuitSink sink)
        {
            if (application == null || sink == null || !(application is IConnectionPointContainer container))
            {
                return null;
            }

            foreach (Guid sourceIid in new[]
            {
                new Guid("0006302C-0000-0000-C000-000000000046"), // ApplicationEvents_11 (Outlook 2003+)
                new Guid("0006300E-0000-0000-C000-000000000046"), // ApplicationEvents_10
                new Guid("0006304E-0000-0000-C000-000000000046"), // ApplicationEvents
            })
            {
                IConnectionPoint? point = null;
                try
                {
                    Guid iid = sourceIid;
                    container.FindConnectionPoint(ref iid, out point);
                    if (point == null)
                    {
                        continue;
                    }

                    point.Advise(sink, out int cookie);
                    return new OutlookQuitSinkRegistration(point, cookie);
                }
                catch (Exception ex) when (!(ex is OutOfMemoryException))
                {
                    if (point != null && Marshal.IsComObject(point))
                    {
                        Marshal.ReleaseComObject(point);
                    }
                }
            }

            return null;
        }
    }

    /// <summary>
    /// Outlook Application events source dispinterface, minimally declared: the CLR
    /// dispatches incoming Invoke calls by the [DispId] mapping, so declaring only Quit
    /// is sufficient for a sink that ignores every other event.
    /// </summary>
    [ComImport]
    [Guid("0006302C-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IOutlookApplicationEventsSink
    {
        /// <summary>Fires when Outlook begins shutting down.</summary>
        [DispId(61447)]
        void Quit();
    }

    /// <summary>Advise cookie holder; Unadvise is best-effort (Outlook may already be gone).</summary>
    public sealed class OutlookQuitSinkRegistration
    {
        private IConnectionPoint? _point;
        private readonly int _cookie;

        internal OutlookQuitSinkRegistration(IConnectionPoint point, int cookie)
        {
            _point = point;
            _cookie = cookie;
        }

        /// <summary>Unadvises and releases the connection point (swallow-all: teardown of a possibly-dead server).</summary>
        public void Unadvise()
        {
            IConnectionPoint? point = _point;
            _point = null;
            if (point == null)
            {
                return;
            }

            try
            {
                point.Unadvise(_cookie);
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                // Outlook already exited - nothing to unadvise on.
            }
            finally
            {
                if (Marshal.IsComObject(point))
                {
                    Marshal.ReleaseComObject(point);
                }
            }
        }
    }
}
