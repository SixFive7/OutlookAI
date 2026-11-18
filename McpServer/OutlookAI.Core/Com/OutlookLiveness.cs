using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using OutlookAI.Core.AddIn;

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
    /// One look at Outlook from the outside: Windows' verdict, plus what the same window walk saw on
    /// the way - the process, the thread that owns Outlook's windows, and whether any of them is on
    /// screen. Everything the start checks (<see cref="OutlookStartObserver"/>) need from the windows
    /// comes out of this one walk, so they cost no more than the liveness probe always did.
    /// </summary>
    public sealed class OutlookLivenessSnapshot
    {
        internal OutlookLivenessSnapshot(
            OutlookLivenessState state,
            string detail,
            int? processId,
            int? mainThreadId,
            IntPtr mainWindow,
            bool hasVisibleWindow)
        {
            State = state;
            Detail = detail;
            ProcessId = processId;
            MainThreadId = mainThreadId;
            MainWindow = mainWindow;
            HasVisibleWindow = hasVisibleWindow;
        }

        /// <summary>Windows' verdict on Outlook.</summary>
        public OutlookLivenessState State { get; }

        /// <summary>What the probe saw, in a few words, for health output and the start log.</summary>
        public string Detail { get; }

        /// <summary>
        /// The OUTLOOK.EXE this look found, or null when none runs. The first one Windows lists - the
        /// same rule <c>OutlookComSession</c> watches its process by: Outlook runs once per user.
        /// </summary>
        public int? ProcessId { get; }

        /// <summary>
        /// The thread that owns Outlook's UI windows - its main thread - or null before Outlook has
        /// created any. Preferably the owner of the hidden "Microsoft Outlook" window, which exists in
        /// an Outlook with no window on screen too.
        /// </summary>
        public int? MainThreadId { get; }

        /// <summary>One of the main thread's windows, for asking whether that thread answers. Zero when none.</summary>
        internal IntPtr MainWindow { get; }

        /// <summary>
        /// Whether an Explorer or Inspector window of Outlook's is visible - a window the user can see,
        /// minimised ones included. False for an Outlook started by a program, with no window.
        /// </summary>
        public bool HasVisibleWindow { get; }
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
        /// <summary>The Explorer and Inspector window class - the windows a user sees.</summary>
        private const string ExplorerWindowClass = "rctrl_renwnd32";

        /// <summary>The hidden "Microsoft Outlook" top-level window, which exists in a window-less Outlook too.</summary>
        private const string HiddenMainWindowClass = "mspim_wnd32";

        /// <summary>
        /// Window classes that only exist once Outlook's UI thread is genuinely up.
        /// Checking these rather than every window avoids judging Outlook by helper
        /// windows that other components create on their own threads.
        /// </summary>
        private static readonly string[] UiWindowClasses =
        {
            ExplorerWindowClass,     // Explorer / Inspector - the real Outlook UI windows
            HiddenMainWindowClass,   // "Microsoft Outlook" hidden top-level window
        };

        /// <summary>Probes Outlook. Never throws, never blocks; worst case a few hundred microseconds.</summary>
        public static OutlookLivenessState Probe()
        {
            return Probe(out _);
        }

        /// <summary>Probes Outlook, reporting how many candidate windows were seen and how many were hung.</summary>
        public static OutlookLivenessState Probe(out string detail)
        {
            OutlookLivenessSnapshot snapshot = Snapshot();
            detail = snapshot.Detail;
            return snapshot.State;
        }

        /// <summary>
        /// Probes Outlook and keeps what the window walk saw on the way. Never throws, never blocks; the
        /// same one pass over the windows <see cref="Probe(out string)"/> has always made.
        /// </summary>
        public static OutlookLivenessSnapshot Snapshot()
        {
            List<int> pids = new List<int>();
            try
            {
                foreach (Process process in OutlookProcesses.InThisSession())
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
                return new OutlookLivenessSnapshot(
                    OutlookLivenessState.NotRunning, "process enumeration failed", null, null, IntPtr.Zero, false);
            }

            if (pids.Count == 0)
            {
                return new OutlookLivenessSnapshot(OutlookLivenessState.NotRunning, "no OUTLOOK.EXE", null, null, IntPtr.Zero, false);
            }

            int candidates = 0;
            int hung = 0;
            bool visible = false;
            IntPtr mainWindow = IntPtr.Zero;
            int? mainThread = null;
            bool mainWindowIsHidden = false;
            try
            {
                NativeWindows.EnumWindows(
                    (handle, unusedLParam) =>
                    {
                        // Not `_` for the lParam: the discards below would then assign to it.
                        uint owner;
                        uint thread = NativeWindows.GetWindowThreadProcessId(handle, out owner);
                        if (!pids.Contains((int)owner))
                        {
                            return true;
                        }

                        StringBuilder className = new StringBuilder(64);
                        _ = NativeWindows.GetClassNameW(handle, className, className.Capacity);
                        string name = className.ToString();
                        for (int i = 0; i < UiWindowClasses.Length; i++)
                        {
                            if (string.Equals(name, UiWindowClasses[i], StringComparison.Ordinal))
                            {
                                candidates++;
                                if (NativeWindows.IsHungAppWindow(handle))
                                {
                                    hung++;
                                }

                                bool isHiddenMain = string.Equals(name, HiddenMainWindowClass, StringComparison.Ordinal);
                                if (mainWindow == IntPtr.Zero || (isHiddenMain && !mainWindowIsHidden))
                                {
                                    mainWindow = handle;
                                    mainThread = unchecked((int)thread);
                                    mainWindowIsHidden = isHiddenMain;
                                }

                                if (!isHiddenMain && NativeWindows.IsWindowVisible(handle))
                                {
                                    visible = true;
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
                return new OutlookLivenessSnapshot(OutlookLivenessState.Starting, "window enumeration failed", pids[0], null, IntPtr.Zero, false);
            }

            if (candidates == 0)
            {
                // Running with no UI windows yet. On a normal start this lasts a second or
                // two; treating it as "starting" rather than "hung" is what keeps a cold
                // start from being misreported as a failure. How long is too long is the
                // start policy's question (OutlookStartPolicy.StartingLimitMilliseconds).
                return new OutlookLivenessSnapshot(
                    OutlookLivenessState.Starting, "running, no UI windows yet", pids[0], null, IntPtr.Zero, false);
            }

            string detail = hung + " of " + candidates + " UI windows hung";
            return new OutlookLivenessSnapshot(
                hung == candidates ? OutlookLivenessState.Hung : OutlookLivenessState.Responsive,
                detail,
                pids[0],
                mainThread,
                mainWindow,
                visible);
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
    }
}
