using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// One dedicated STA thread that runs a REAL Win32 message pump plus a serialized
    /// work queue - the v3.MD section-0.5.2 obligation for the ComGateway thread: COM
    /// cross-apartment calls into this STA (and, in v3.1, out-of-process event advise
    /// sinks) arrive as window messages, so the thread must dispatch messages while it
    /// waits for work, not merely block on a queue.
    ///
    /// Implementation: <c>MsgWaitForMultipleObjectsEx(QS_ALLINPUT | MWMO_INPUTAVAILABLE)</c>
    /// over a work-available event; queued work runs between message bursts, and every
    /// pending window message is translated + dispatched. All Outlook COM objects are
    /// created and used exclusively on this thread (v3.MD section 12: never marshal
    /// Outlook objects across threads).
    /// </summary>
    internal sealed class PumpedStaRunner : IDisposable
    {
        private readonly ConcurrentQueue<KeyValuePair<Func<object?>, TaskCompletionSource<object?>>> _queue =
            new ConcurrentQueue<KeyValuePair<Func<object?>, TaskCompletionSource<object?>>>();

        private readonly AutoResetEvent _workAvailable = new AutoResetEvent(false);
        private readonly ManualResetEventSlim _pumpReady = new ManualResetEventSlim(false);
        private readonly Thread _thread;
        private volatile bool _shutdown;
        private volatile bool _disposed;

        internal PumpedStaRunner(string threadName)
        {
            _thread = new Thread(PumpLoop)
            {
                IsBackground = true,
                Name = threadName,
            };
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
            _pumpReady.Wait();
        }

        /// <summary>Runs <paramref name="work"/> on the pumped STA thread and returns its result.</summary>
        internal T Run<T>(Func<T> work)
        {
            if (work == null)
            {
                throw new ArgumentNullException(nameof(work));
            }

            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(PumpedStaRunner));
            }

            if (Thread.CurrentThread.ManagedThreadId == _thread.ManagedThreadId)
            {
                // Re-entrant call from the STA thread itself (e.g. from a future event
                // sink): execute inline instead of deadlocking on the queue.
                return work();
            }

            TaskCompletionSource<object?> completion =
                new TaskCompletionSource<object?>(TaskCreationOptions.RunContinuationsAsynchronously);
            _queue.Enqueue(new KeyValuePair<Func<object?>, TaskCompletionSource<object?>>(() => work(), completion));
            _workAvailable.Set();
            return (T)completion.Task.GetAwaiter().GetResult()!;
        }

        /// <summary>Runs <paramref name="work"/> on the pumped STA thread.</summary>
        internal void Run(Action work)
        {
            if (work == null)
            {
                throw new ArgumentNullException(nameof(work));
            }

            Run<object?>(() =>
            {
                work();
                return null;
            });
        }

        private void PumpLoop()
        {
            // Touching the message queue forces its creation before callers may rely on it.
            NativeMethods.PeekMessage(out NativeMethods.MSG _, IntPtr.Zero, 0, 0, NativeMethods.PM_NOREMOVE);
            _pumpReady.Set();

            IntPtr[] handles = { _workAvailable.SafeWaitHandle.DangerousGetHandle() };
            while (!_shutdown)
            {
                uint wait = NativeMethods.MsgWaitForMultipleObjectsEx(
                    1,
                    handles,
                    NativeMethods.Infinite,
                    NativeMethods.QS_ALLINPUT,
                    NativeMethods.MWMO_INPUTAVAILABLE);

                if (wait == NativeMethods.WaitObject0)
                {
                    DrainWorkQueue();
                }
                else if (wait == NativeMethods.WaitObject0 + 1)
                {
                    PumpPendingMessages();
                }
                else
                {
                    // WAIT_FAILED or an unexpected code - bail out rather than spin hot.
                    break;
                }
            }

            DrainWorkQueue();
            FailPendingWork();
        }

        private void PumpPendingMessages()
        {
            while (NativeMethods.PeekMessage(out NativeMethods.MSG msg, IntPtr.Zero, 0, 0, NativeMethods.PM_REMOVE))
            {
                if (msg.message == NativeMethods.WM_QUIT)
                {
                    _shutdown = true;
                    return;
                }

                NativeMethods.TranslateMessage(ref msg);
                NativeMethods.DispatchMessage(ref msg);
            }
        }

        private void DrainWorkQueue()
        {
            while (_queue.TryDequeue(out KeyValuePair<Func<object?>, TaskCompletionSource<object?>> item))
            {
                try
                {
                    item.Value.SetResult(item.Key());
                }
                catch (Exception ex)
                {
                    item.Value.SetException(ex);
                }
            }
        }

        private void FailPendingWork()
        {
            while (_queue.TryDequeue(out KeyValuePair<Func<object?>, TaskCompletionSource<object?>> item))
            {
                item.Value.TrySetException(new ObjectDisposedException(nameof(PumpedStaRunner)));
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _shutdown = true;
            _workAvailable.Set();
            if (!_thread.Join(TimeSpan.FromSeconds(15)))
            {
                // Background thread; the process may exit with it still parked. Never
                // abort - a COM call could be mid-flight.
            }

            _workAvailable.Dispose();
            _pumpReady.Dispose();
        }

        private static class NativeMethods
        {
            internal const uint PM_NOREMOVE = 0x0000;
            internal const uint PM_REMOVE = 0x0001;
            internal const uint WM_QUIT = 0x0012;
            internal const uint QS_ALLINPUT = 0x04FF;
            internal const uint MWMO_INPUTAVAILABLE = 0x0004;
            internal const uint Infinite = 0xFFFFFFFF;
            internal const uint WaitObject0 = 0;

            [StructLayout(LayoutKind.Sequential)]
            internal struct POINT
            {
                public int X;
                public int Y;
            }

            [StructLayout(LayoutKind.Sequential)]
            internal struct MSG
            {
                public IntPtr hwnd;
                public uint message;
                public IntPtr wParam;
                public IntPtr lParam;
                public uint time;
                public POINT pt;
            }

            [DllImport("user32.dll", SetLastError = true)]
            internal static extern uint MsgWaitForMultipleObjectsEx(
                uint nCount, IntPtr[] pHandles, uint dwMilliseconds, uint dwWakeMask, uint dwFlags);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool PeekMessage(out MSG lpMsg, IntPtr hWnd, uint wMsgFilterMin, uint wMsgFilterMax, uint wRemoveMsg);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool TranslateMessage([In] ref MSG lpMsg);

            [DllImport("user32.dll")]
            internal static extern IntPtr DispatchMessage([In] ref MSG lpMsg);
        }
    }
}
