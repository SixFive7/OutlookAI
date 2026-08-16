using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace OutlookAI.Core.Com
{
    /// <summary>How Outlook looks from the outside, judged without touching COM.</summary>
    public enum OutlookLivenessState
    {
        /// <summary>No OUTLOOK.EXE for this user.</summary>
        NotRunning = 0,

        /// <summary>Running, but it has not created its windows yet - almost certainly still starting.</summary>
        Starting = 1,

        /// <summary>Running and pumping messages.</summary>
        Responsive = 2,

        /// <summary>Running, but Windows itself reports its windows as hung. COM calls into it will not return.</summary>
        Hung = 3,
    }

    /// <summary>
    /// Answers "is Outlook actually alive?" using only Win32 - no COM, no automation, no
    /// chance of blocking.
    /// <para>
    /// This matters because the expensive way to discover a wedged Outlook is to make a
    /// COM call and wait out its budget. Measured on 2026-08-16, that cost 30-120 s per
    /// request. Windows already knows the answer: it tracks whether a window's owning
    /// thread is servicing its message queue, and will tell us in microseconds through
    /// <c>IsHungAppWindow</c>. Asking first turns a 30 s discovery into a free one.
    /// </para>
    /// <para>
    /// Deliberately conservative. Only <see cref="OutlookLivenessState.Hung"/> is treated
    /// as a hard "do not call" signal, and it is reported only when Windows flags EVERY
    /// candidate window as hung - a single busy window during a long operation is normal
    /// and must not be mistaken for a wedge.
    /// </para>
    /// </summary>
    public static class OutlookLiveness
    {
        private const string ProcessName = "OUTLOOK";

        /// <summary>
        /// Window classes that only exist once Outlook's UI thread is genuinely up.
        /// Checking these rather than every window avoids judging Outlook by helper
        /// windows that other components create on their own threads.
        /// </summary>
        private static readonly string[] UiWindowClasses =
        {
            "rctrl_renwnd32",  // Explorer / Inspector - the real Outlook UI windows
            "mspim_wnd32",     // "Microsoft Outlook" hidden top-level window
        };

        /// <summary>Probes Outlook. Never throws, never blocks; worst case a few hundred microseconds.</summary>
        public static OutlookLivenessState Probe()
        {
            return Probe(out _);
        }

        /// <summary>Probes Outlook, reporting how many candidate windows were seen and how many were hung.</summary>
        public static OutlookLivenessState Probe(out string detail)
        {
            detail = string.Empty;
            List<int> pids = new List<int>();
            try
            {
                foreach (Process process in Process.GetProcessesByName(ProcessName))
                {
                    try
                    {
                        pids.Add(process.Id);
                    }
                    finally
                    {
                        process.Dispose();
                    }
                }
            }
            catch (Exception)
            {
                // Enumeration itself failed; report the safe answer rather than guessing.
                detail = "process enumeration failed";
                return OutlookLivenessState.NotRunning;
            }

            if (pids.Count == 0)
            {
                detail = "no OUTLOOK.EXE";
                return OutlookLivenessState.NotRunning;
            }

            int candidates = 0;
            int hung = 0;
            try
            {
                NativeMethods.EnumWindows(
                    (handle, unusedLParam) =>
                    {
                        // Not `_` for the lParam: the discards below would then assign to it.
                        uint owner;
                        _ = NativeMethods.GetWindowThreadProcessId(handle, out owner);
                        if (!pids.Contains((int)owner))
                        {
                            return true;
                        }

                        StringBuilder className = new StringBuilder(64);
                        _ = NativeMethods.GetClassNameW(handle, className, className.Capacity);
                        string name = className.ToString();
                        for (int i = 0; i < UiWindowClasses.Length; i++)
                        {
                            if (string.Equals(name, UiWindowClasses[i], StringComparison.Ordinal))
                            {
                                candidates++;
                                if (NativeMethods.IsHungAppWindow(handle))
                                {
                                    hung++;
                                }

                                break;
                            }
                        }

                        return true;
                    },
                    IntPtr.Zero);
            }
            catch (Exception)
            {
                detail = "window enumeration failed";
                return OutlookLivenessState.Starting;
            }

            if (candidates == 0)
            {
                // Running with no UI windows yet. On a normal start this lasts a second or
                // two; treating it as "starting" rather than "hung" is what keeps a cold
                // start from being misreported as a failure.
                detail = "running, no UI windows yet";
                return OutlookLivenessState.Starting;
            }

            detail = hung + " of " + candidates + " UI windows hung";
            return hung == candidates ? OutlookLivenessState.Hung : OutlookLivenessState.Responsive;
        }

        /// <summary>Human-readable one-liner for health output.</summary>
        public static string Describe(OutlookLivenessState state)
        {
            switch (state)
            {
                case OutlookLivenessState.NotRunning:
                    return "not running";
                case OutlookLivenessState.Starting:
                    return "starting";
                case OutlookLivenessState.Responsive:
                    return "responsive";
                case OutlookLivenessState.Hung:
                    return "not responding";
                default:
                    return "unknown";
            }
        }

        private static class NativeMethods
        {
            internal delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

            [DllImport("user32.dll")]
            internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            internal static extern int GetClassNameW(IntPtr hWnd, StringBuilder className, int maxCount);

            /// <summary>
            /// Windows' own judgement: true when the window's thread has stopped servicing
            /// its message queue. This is the same signal Explorer uses to grey out a
            /// window and offer to close it.
            /// </summary>
            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool IsHungAppWindow(IntPtr hWnd);
        }
    }
}
