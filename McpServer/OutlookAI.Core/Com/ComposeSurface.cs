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
    /// What it rests on, measured on Office LTSC 2021 (16.0.14334) by the D49 probe of 2026-07-31 - 22 runs on
    /// one machine, kept with their sample sizes in <c>Docs/research/q149-draft-text/</c>:
    /// <c>Inspector.WordEditor</c> is unobtainable while Outlook has no window that is displayed or was
    /// activated - it does not return null, it THROWS <c>COMException "The operation failed."</c> (2 of 2
    /// cold starts, and every such read D46 and D48 made). The window it needs is the PROCESS's, not the
    /// draft's own: once Outlook shows a window, a never-shown inspector hands out its editor directly
    /// (D45: 31 of 31), and after one activation in a process later never-activated inspectors got it too
    /// (the maintainer's audit log, 20 of 21). What CAN be controlled is whether a HUMAN ever sees the
    /// window. The same questions on Office LTSC 2024 (16.0.17932) were measured by the Q149 (d) suite
    /// (<c>T2/LiveQ149WindowlessComposeTests</c>) on 2026-11-17, with its results in that folder: the editor
    /// refused the same way (L2, 20 of 20 on the two test guests) - after 30 seconds of reads, and with a
    /// window shown by Windows, as well (W2-W5) - and after one promotion in a process the later
    /// never-activated inspectors handed it out (W6, 60 of 60), their edits committing through
    /// <c>ModifiedFormPages.Add("Message")</c> and <c>Close(olSave)</c> (W7, 60 of 60) and not through
    /// <c>CurrentItem.Save()</c> (W6, 0 of 60).
    /// </para>
    /// <para>
    /// So this class does exactly two things, and they fix two DIFFERENT problems that
    /// D48 conflated:
    /// </para>
    /// <list type="number">
    /// <item><description>
    /// <b>The process pin</b> (<see cref="TryPinProcess"/>) - a NON-DISPLAYED
    /// <c>Explorers.Add(folder, 0)</c>. It does nothing whatsoever for the editor, but it
    /// is a complete cure for the LIFETIME problem: measured on Office LTSC 2021 (2 of 2 cold
    /// starts) and on Office LTSC 2024 (Q149 L1: 40 of 40 shipped composes without the pin, 0 of 10
    /// with it), a window-less Outlook EXITS when the only Inspector is closed, which is what killed the compose path's own
    /// instance, made <c>update_draft</c> fail with <c>com_failure</c>, and produced the
    /// three <c>RPC_S_SERVER_UNAVAILABLE</c> suite failures. With a non-displayed Explorer
    /// held the instance survives, and D33 promotability is preserved (probe-measured:
    /// launching outlook.exe with the pin held still gives the user a normal window).
    /// </description></item>
    /// <item><description>
    /// <b>The editor promotion</b> (<see cref="PromoteForWordEditor"/>) - park the
    /// inspector's (already existing, invisible) window OFF-SCREEN, <c>Activate()</c> it,
    /// then <c>ShowWindow(SW_HIDE)</c> anything that became visible. Measured on Office LTSC 2021
    /// (2 runs): 53-79 ms from the <c>Activate()</c> to a live WordEditor - the activation itself took
    /// 314-611 ms - with the window ending <c>IsWindowVisible == false</c>: absent
    /// from the screen, the taskbar and Alt-Tab. Ordering is load-bearing - off-screen
    /// FIRST, because <c>Activate()</c> alone paints the window where a user would see it
    /// (that is what the shipped <c>update_draft</c> path has been doing).
    /// <para>
    /// On Office LTSC 2024 the park does not hold in a window-less Outlook (Q149, 2026-11-17): in 50 of 50
    /// composes on the two test guests <c>Activate()</c> showed the inspector at Outlook's own position,
    /// (25,25) 768x519, on screen for 130-646 ms until the sweep below hid it, and gave Outlook the
    /// foreground. In a windowed Outlook it stayed off-screen (B0, 40 of 40).
    /// </para>
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
        /// <summary>
        /// IUnknown pointers of every lifetime pin this PROCESS holds. COM identity is the
        /// only sound way to recognise a pin: it has no distinguishing folder, caption or
        /// window title, and several server sessions can be live at once, so per-session
        /// ownership is not enough to answer "is this Explorer a pin?".
        /// </summary>
        private static readonly HashSet<IntPtr> PinPointers = new HashSet<IntPtr>();

        private static readonly object PinLock = new object();

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


        /// <summary>True when <paramref name="explorer"/> IS one of this process's pins.</summary>
        public static bool IsPin(object? explorer)
        {
            if (explorer == null)
            {
                return false;
            }

            IntPtr p = IntPtr.Zero;
            try
            {
                p = Marshal.GetIUnknownForObject(explorer);
                lock (PinLock)
                {
                    return PinPointers.Contains(p);
                }
            }
            catch (Exception)
            {
                return false;
            }
            finally
            {
                if (p != IntPtr.Zero)
                {
                    _ = Marshal.Release(p);
                }
            }
        }

        private static void RememberPin(object explorer, bool remember)
        {
            IntPtr p = IntPtr.Zero;
            try
            {
                p = Marshal.GetIUnknownForObject(explorer);
                lock (PinLock)
                {
                    if (remember)
                    {
                        _ = PinPointers.Add(p);
                    }
                    else
                    {
                        _ = PinPointers.Remove(p);
                    }
                }
            }
            catch (Exception)
            {
            }
            finally
            {
                if (p != IntPtr.Zero)
                {
                    _ = Marshal.Release(p);
                }
            }
        }

        /// <summary>Drops <paramref name="explorer"/> from the pin registry (teardown).</summary>
        public static void ForgetPin(object? explorer)
        {
            if (explorer != null)
            {
                RememberPin(explorer, remember: false);
            }
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
            return PromoteForWordEditor(inspectorObject, out error, out _);
        }

        /// <summary>
        /// <see cref="PromoteForWordEditor(object, out string)"/>, also saying whether Outlook had no window a
        /// person could see when the promotion began - the case the promotion exists for, and the one the
        /// audit line's <c>composeSurface</c> reports as <c>wordEditorPromoted</c>. Since 2026-10-04 every
        /// compose promotes, windowed or not, because an edit on a re-opened draft commits only through an
        /// activated inspector; this is what still tells the two apart.
        /// <para>
        /// A refusal keeps its cause: <paramref name="error"/> is <see cref="ComErrorTokens.NoWordEditor"/>
        /// or <see cref="ComErrorTokens.ActivateFailed"/> with the failure's type and HRESULT after the colon
        /// (<see cref="ComErrorTokens.With"/>), where it used to be the bare word.
        /// </para>
        /// </summary>
        internal static object? PromoteForWordEditor(object inspectorObject, out string? error, out bool outlookWasWindowless)
        {
            error = null;
            outlookWasWindowless = false;
            if (inspectorObject == null)
            {
                error = "NoInspector";
                return null;
            }

            IReadOnlyList<WindowState> before;
            try
            {
                before = SnapshotOutlookWindows();
                outlookWasWindowless = !before.Any(w => w.Visible);
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
                error = ComErrorTokens.With(ComErrorTokens.ActivateFailed, OutlookComSession.DescribeComFailure(ex));
            }

            HideAnythingThatBecameVisible(before);

            // Outlook can reposition/re-show the promoted window once its own layout pass
            // runs; one pumped settle plus a second sweep covers that without a race.
            PumpedStaRunner.PumpedWait(60);
            HideAnythingThatBecameVisible(before);

            object? document = null;
            string? editorFailure = null;
            try
            {
                document = ((dynamic)inspectorObject).WordEditor;
            }
            catch (Exception ex) when (OutlookComSession.IsComCallFailure(ex))
            {
                // A window-less Outlook throws here rather than returning null - on Office LTSC
                // 2021 COMException "The operation failed." (D49). Kept with its HRESULT: the
                // bare word made every refusal in the audit log undiagnosable.
                editorFailure = OutlookComSession.DescribeComFailure(ex);
            }

            if (document == null && error == null)
            {
                error = ComErrorTokens.With(ComErrorTokens.NoWordEditor, editorFailure ?? "returned null");
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
                    // Anything already holding Outlook open - the user's own window, or
                    // another server session's pin - makes a second one pure cost, and
                    // D33 says we create no surface we do not need. A pin is only ever
                    // RELEASED, never closed, on session teardown, so relying on another
                    // session's is safe: it outlives that session.
                    return null;
                }

                folder = ((dynamic)namespaceObject).GetDefaultFolder(6); // olFolderInbox

                // 0 = olFolderDisplayNormal. Deliberately NOT followed by Display():
                // an Explorer that is never displayed owns an INVISIBLE window, which
                // pins the process without putting anything on screen (D49 probe row 2, Office LTSC 2021).
                object created = ((dynamic)explorers!).Add(folder, 0);
                RememberPin(created, remember: true);
                return created;
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

        /// <summary>
        /// Adds the Explorer a show-me call puts on screen - and never hands back an Explorer that
        /// already existed, a lifetime pin above all.
        /// <para>
        /// MEASURED on Office LTSC 2024 (16.0.17932, the unindexed test guest, 2026-10-03, D49 probe
        /// v3): <c>Explorers.Add</c> on the folder a NON-DISPLAYED Explorer already shows returns THAT
        /// Explorer - the same COM object, <c>Explorers.Count</c> unchanged - while an Add on any other
        /// folder makes a new one, and the non-displayed Explorer then keeps Outlook running after the
        /// new window is closed. The pin (<see cref="TryPinProcess"/>) is a non-displayed Explorer on
        /// the default Inbox, so a show-me call for that Inbox was handed the pin, displayed it, and the
        /// user's close took the pin - and Outlook - with it: the D49 regression the guest's live runs
        /// reported while the session still said it was pinned.
        /// </para>
        /// <para>
        /// So an Explorer that comes back as a pin, or without the count going up, is released -
        /// never displayed, never closed - and the Explorer is added on <paramref name="alternateFolder"/>
        /// instead; <paramref name="rerouted"/> tells the caller to navigate it to
        /// <paramref name="folder"/> once it is on screen. The count is checked as well as the pin
        /// registry because a pin another session made is reached through another apartment's proxy,
        /// whose pointer the registry does not hold. When the alternate is unavailable or comes back
        /// the same way, the answer is an error: an existing Explorer is never returned from here.
        /// </para>
        /// </summary>
        /// <param name="add"><c>Explorers.Add(folder, olFolderDisplayNormal)</c>.</param>
        /// <param name="countExplorers">A fresh <c>Explorers.Count</c>; negative when it cannot be read.</param>
        /// <param name="folder">The folder the caller wants on screen.</param>
        /// <param name="alternateFolder">Another folder to open the window on - the store's top folder. Called at most once; the result is released here.</param>
        /// <param name="rerouted">True when the returned Explorer was added on the alternate folder.</param>
        /// <param name="error">Content-free reason when nothing was returned.</param>
        internal static object? AddShowMeExplorer(
            Func<object, object?> add,
            Func<int> countExplorers,
            object folder,
            Func<object?> alternateFolder,
            out bool rerouted,
            out string? error)
        {
            rerouted = false;
            error = null;
            int before = countExplorers();
            object? explorer = add(folder);
            if (!IsExistingExplorer(explorer, before, countExplorers))
            {
                return explorer;
            }

            OutlookComSession.Release(explorer);
            object? alternate = null;
            try
            {
                alternate = alternateFolder();
                if (alternate == null)
                {
                    error = "Outlook handed back an Explorer that already existed - its hidden lifetime window - for this folder, "
                        + "and no other folder was available to open a new window on.";
                    return null;
                }

                before = countExplorers();
                explorer = add(alternate);
                if (IsExistingExplorer(explorer, before, countExplorers))
                {
                    OutlookComSession.Release(explorer);
                    error = "Outlook handed back an Explorer that already existed for every folder tried, so no window was opened.";
                    return null;
                }

                rerouted = true;
                return explorer;
            }
            finally
            {
                OutlookComSession.Release(alternate);
            }
        }

        /// <summary>An Explorer <c>Explorers.Add</c> did not create: a known pin, or the count did not go up.</summary>
        private static bool IsExistingExplorer(object? explorer, int countBefore, Func<int> countExplorers)
        {
            if (explorer == null || IsPin(explorer))
            {
                return true;
            }

            if (countBefore < 0)
            {
                return false;
            }

            int after = countExplorers();
            return after >= 0 && after <= countBefore;
        }

        /// <summary>Every top-level window owned by the OUTLOOK.EXE processes.</summary>
        public static IReadOnlyList<WindowState> SnapshotOutlookWindows()
        {
            HashSet<int> pids = new HashSet<int>();
            foreach (Process p in OutlookProcesses.InThisSession())
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
