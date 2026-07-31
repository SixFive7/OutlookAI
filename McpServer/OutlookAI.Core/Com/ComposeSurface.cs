using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// THE INVISIBLE COMPOSE SURFACE (v3.MD D49, soak fix 22).
    /// <para>
    /// Phase-1 measured fact: <c>Inspector.WordEditor</c> is unobtainable while Outlook
    /// is window-less - it does not return null, it THROWS
    /// <c>COMException "The operation failed."</c> - and EVERY route that makes it
    /// obtainable requires a real top-level window. There is no window-free Word path.
    /// What CAN be controlled is whether a HUMAN ever sees that window.
    /// </para>
    /// <para>
    /// So this class does exactly two things, and they fix two DIFFERENT problems that
    /// D48 conflated:
    /// </para>
    /// <list type="number">
    /// <item><description>
    /// <b>The process pin</b> (<see cref="TryPinProcess"/>) - a NON-DISPLAYED
    /// <c>Explorers.Add(folder, 0)</c>. It does nothing whatsoever for the editor, but it
    /// is a complete cure for the LIFETIME problem: measured, a window-less Outlook EXITS
    /// when the only Inspector is closed, which is what killed the compose path's own
    /// instance, made <c>update_draft</c> fail with <c>com_failure</c>, and produced the
    /// three <c>RPC_S_SERVER_UNAVAILABLE</c> suite failures. With a non-displayed Explorer
    /// held the instance survives, and D33 promotability is preserved (probe-measured:
    /// launching outlook.exe with the pin held still gives the user a normal window).
    /// </description></item>
    /// <item><description>
    /// <b>The editor promotion</b> (<see cref="PromoteForWordEditor"/>) - park the
    /// inspector's (already existing, invisible) window OFF-SCREEN, <c>Activate()</c> it,
    /// then <c>ShowWindow(SW_HIDE)</c> anything that became visible. Measured 53-79 ms to
    /// a live WordEditor with the window ending <c>IsWindowVisible == false</c>: absent
    /// from the screen, the taskbar and Alt-Tab. Ordering is load-bearing - off-screen
    /// FIRST, because <c>Activate()</c> alone paints the window where a user would see it
    /// (that is what the shipped <c>update_draft</c> path has been doing).
    /// </description></item>
    /// </list>
    /// <para>
    /// SAFETY RULE, and the reason the window selection is a pure function
    /// (<see cref="SelectWindowsToPark"/> / <see cref="SelectWindowsToHide"/>): a window
    /// the user could already see is NEVER touched. Promotion parks only windows that are
    /// <c>IsWindowVisible == false</c> at entry and hides only windows that became visible
    /// DURING the promotion. A minimized or DWM-cloaked Outlook still reports
    /// <c>IsWindowVisible == true</c>, so the user's own minimized Outlook can never be
    /// selected.
    /// </para>
    /// </summary>
    public static class ComposeSurface
    {
        /// <summary>Where a parked window is moved. Far outside any real virtual screen.</summary>
        public const int ParkX = -32000;

        /// <summary>Where a parked window is moved.</summary>
        public const int ParkY = -32000;

        /// <summary>
        /// One top-level window, reduced to what the selection rules need. A record so the
        /// rules can be exercised in T1 without a window manager.
        /// </summary>
        public sealed class WindowState
        {
            public WindowState(IntPtr handle, bool visible)
            {
                Handle = handle;
                Visible = visible;
            }

            public IntPtr Handle { get; }

            /// <summary>
            /// Win32 <c>IsWindowVisible</c>. NOTE: this is TRUE for a minimized window and
            /// TRUE for a DWM-cloaked one - which is exactly why it is the right predicate
            /// for "the user might be able to see this", and why nothing here uses a
            /// rectangle test to decide ownership.
            /// </summary>
            public bool Visible { get; }
        }

        /// <summary>
        /// Windows to park off-screen BEFORE promoting: every currently INVISIBLE window.
        /// A visible one is the user's and is never returned.
        /// </summary>
        public static IReadOnlyList<IntPtr> SelectWindowsToPark(IEnumerable<WindowState> windows)
        {
            if (windows == null)
            {
                return Array.Empty<IntPtr>();
            }

            List<IntPtr> parked = new List<IntPtr>();
            foreach (WindowState w in windows)
            {
                if (!w.Visible)
                {
                    parked.Add(w.Handle);
                }
            }

            return parked;
        }

        /// <summary>
        /// Windows to hide AFTER promoting: those visible now that were not visible before.
        /// Expressed as a set difference so a window CREATED by the promotion is covered
        /// too, and so a window the user already had open never can be.
        /// </summary>
        public static IReadOnlyList<IntPtr> SelectWindowsToHide(
            IEnumerable<WindowState> before,
            IEnumerable<WindowState> after)
        {
            HashSet<IntPtr> visibleBefore = new HashSet<IntPtr>();
            if (before != null)
            {
                foreach (WindowState w in before)
                {
                    if (w.Visible)
                    {
                        _ = visibleBefore.Add(w.Handle);
                    }
                }
            }

            List<IntPtr> toHide = new List<IntPtr>();
            if (after == null)
            {
                return toHide;
            }

            foreach (WindowState w in after)
            {
                if (w.Visible && !visibleBefore.Contains(w.Handle))
                {
                    toHide.Add(w.Handle);
                }
            }

            return toHide;
        }

        /// <summary>
        /// Promotes <paramref name="inspectorObject"/> so its <c>WordEditor</c> becomes
        /// obtainable, without ever showing a window to the user. Returns the Word
        /// document, or null when the surface could not be promoted (the CALLER must then
        /// refuse or report - never silently compose a lesser draft, D49).
        /// </summary>
        internal static object? PromoteForWordEditor(object inspectorObject, out string? error)
        {
            error = null;
            if (inspectorObject == null)
            {
                error = "NoInspector";
                return null;
            }

            IReadOnlyList<WindowState> before;
            try
            {
                before = SnapshotOutlookWindows();
                foreach (IntPtr h in SelectWindowsToPark(before))
                {
                    Park(h);
                }
            }
            catch (Exception)
            {
                // Window discipline is best-effort; a failure here must not stop the
                // promotion, it only risks a brief flash.
                before = Array.Empty<WindowState>();
            }

            try
            {
                ((dynamic)inspectorObject).Activate();
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                error = "ActivateFailed";
            }

            HideAnythingThatBecameVisible(before);

            // Outlook can reposition/re-show the promoted window once its own layout pass
            // runs; one pumped settle plus a second sweep covers that without a race.
            PumpedStaRunner.PumpedWait(60);
            HideAnythingThatBecameVisible(before);

            object? document = null;
            try
            {
                document = ((dynamic)inspectorObject).WordEditor;
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                // Headless throws here rather than returning null (Phase-1 finding 1).
            }

            if (document == null && error == null)
            {
                error = "NoWordEditor";
            }

            return document;
        }

        /// <summary>
        /// Re-hides anything the promotion put on screen. Safe to call repeatedly.
        /// </summary>
        internal static void HideAnythingThatBecameVisible(IReadOnlyList<WindowState> before)
        {
            try
            {
                IReadOnlyList<WindowState> after = SnapshotOutlookWindows();
                foreach (IntPtr h in SelectWindowsToHide(before, after))
                {
                    Park(h);
                    _ = NativeMethods.ShowWindow(h, NativeMethods.SW_HIDE);
                }
            }
            catch (Exception)
            {
                // Best-effort by construction.
            }
        }

        /// <summary>
        /// Creates the non-displayed Explorer that keeps a window-less Outlook alive across
        /// an <c>Inspector.Close</c>. Returns the Explorer RCW (the SESSION owns it and must
        /// release - never Close - it) or null when a window/Explorer already exists.
        /// </summary>
        internal static object? TryPinProcess(object applicationObject, object namespaceObject, out string? error)
        {
            error = null;
            object? explorers = null;
            object? folder = null;
            try
            {
                explorers = ((dynamic)applicationObject).Explorers;
                int count = (int)((dynamic)explorers!).Count;
                if (count > 0)
                {
                    // Something already holds Outlook open - adding another Explorer would
                    // be pure cost, and D33 says we create no surface we do not need.
                    return null;
                }

                folder = ((dynamic)namespaceObject).GetDefaultFolder(6); // olFolderInbox

                // 0 = olFolderDisplayNormal. Deliberately NOT followed by Display():
                // an Explorer that is never displayed owns an INVISIBLE window, which
                // pins the process without putting anything on screen (Phase-1 row 2).
                return ((dynamic)explorers!).Add(folder, 0);
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                error = OutlookComSession.DescribeComFailure(ex);
                return null;
            }
            finally
            {
                OutlookComSession.Release(folder);
                OutlookComSession.Release(explorers);
            }
        }

        /// <summary>Every top-level window owned by the OUTLOOK.EXE processes.</summary>
        public static IReadOnlyList<WindowState> SnapshotOutlookWindows()
        {
            HashSet<int> pids = new HashSet<int>();
            foreach (Process p in Process.GetProcessesByName("OUTLOOK"))
            {
                try
                {
                    _ = pids.Add(p.Id);
                }
                finally
                {
                    p.Dispose();
                }
            }

            List<WindowState> windows = new List<WindowState>();
            if (pids.Count == 0)
            {
                return windows;
            }

            _ = NativeMethods.EnumWindows(
                (h, l) =>
                {
                    _ = NativeMethods.GetWindowThreadProcessId(h, out uint pid);
                    if (pids.Contains((int)pid))
                    {
                        windows.Add(new WindowState(h, NativeMethods.IsWindowVisible(h)));
                    }

                    return true;
                },
                IntPtr.Zero);

            return windows;
        }

        /// <summary>Count of windows a HUMAN can currently see (visible, not minimized, on screen).</summary>
        public static int CountUserVisibleWindows()
        {
            int visible = 0;
            foreach (WindowState w in SnapshotOutlookWindows())
            {
                if (!w.Visible)
                {
                    continue;
                }

                if (!NativeMethods.GetWindowRect(w.Handle, out NativeMethods.RECT r))
                {
                    continue;
                }

                if (r.Right <= r.Left || r.Bottom <= r.Top)
                {
                    continue;
                }

                if (r.Left <= ParkX / 2 && r.Top <= ParkY / 2)
                {
                    continue; // parked
                }

                visible++;
            }

            return visible;
        }

        private static void Park(IntPtr handle)
        {
            _ = NativeMethods.SetWindowPos(
                handle,
                IntPtr.Zero,
                ParkX,
                ParkY,
                0,
                0,
                NativeMethods.SWP_NOSIZE | NativeMethods.SWP_NOZORDER | NativeMethods.SWP_NOACTIVATE);
        }

        private static class NativeMethods
        {
            internal const uint SWP_NOSIZE = 0x0001;
            internal const uint SWP_NOZORDER = 0x0004;
            internal const uint SWP_NOACTIVATE = 0x0010;
            internal const int SW_HIDE = 0;

            internal delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

            [StructLayout(LayoutKind.Sequential)]
            internal struct RECT
            {
                public int Left;
                public int Top;
                public int Right;
                public int Bottom;
            }

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

            [DllImport("user32.dll")]
            internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool IsWindowVisible(IntPtr hWnd);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool SetWindowPos(IntPtr hWnd, IntPtr insertAfter, int x, int y, int cx, int cy, uint flags);

            [DllImport("user32.dll")]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool ShowWindow(IntPtr hWnd, int cmdShow);
        }
    }
}
