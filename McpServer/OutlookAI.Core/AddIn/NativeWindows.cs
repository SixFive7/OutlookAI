using System;
using System.Runtime.InteropServices;
using System.Text;

namespace OutlookAI.Core.AddIn
{
    /// <summary>
    /// The Win32 window calls OutlookAI makes: finding a process's top-level windows, reading their
    /// state and placing them. One set of declarations for Core and the add-in (the code-sharing
    /// refactor's stage S4, survey group A09); until then each file that needed them declared its own,
    /// each a different subset. Nothing here finds Outlook: a caller passes the window or the process.
    /// </summary>
    public static class NativeWindows
    {
        /// <summary><c>SetWindowPos</c>: keep the window's size.</summary>
        public const uint SWP_NOSIZE = 0x0001;

        /// <summary><c>SetWindowPos</c>: keep the window's place in the Z order.</summary>
        public const uint SWP_NOZORDER = 0x0004;

        /// <summary><c>SetWindowPos</c>: do not activate the window.</summary>
        public const uint SWP_NOACTIVATE = 0x0010;

        /// <summary><c>ShowWindow</c>: hide the window.</summary>
        public const int SW_HIDE = 0;

        /// <summary>The callback <see cref="EnumWindows"/> calls once per top-level window; false stops it.</summary>
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        /// <summary>A window's rectangle in screen coordinates.</summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            /// <summary>The left edge.</summary>
            public int Left;

            /// <summary>The top edge.</summary>
            public int Top;

            /// <summary>The right edge.</summary>
            public int Right;

            /// <summary>The bottom edge.</summary>
            public int Bottom;
        }

        /// <summary>Calls <paramref name="callback"/> for every top-level window.</summary>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

        /// <summary>The thread that made the window, and the process it belongs to.</summary>
        [DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        /// <summary>
        /// Whether the window has the visible style - TRUE for a minimized window and for a
        /// DWM-cloaked one, which is why it is the right test for "the user might be able to see this".
        /// </summary>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        /// <summary>Whether the handle still names a window.</summary>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);

        /// <summary>Whether the window is minimized.</summary>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsIconic(IntPtr hWnd);

        /// <summary>The window's rectangle in screen coordinates.</summary>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        /// <summary>Moves, sizes or reorders the window, as <paramref name="flags"/> say.</summary>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr insertAfter, int x, int y, int cx, int cy, uint flags);

        /// <summary>Shows or hides the window.</summary>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int cmdShow);

        /// <summary>The window's class name, into <paramref name="className"/>.</summary>
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassNameW(IntPtr hWnd, StringBuilder className, int maxCount);

        /// <summary>
        /// Windows' own judgement: true when the window's thread has stopped servicing
        /// its message queue. This is the same signal Explorer uses to grey out a
        /// window and offer to close it.
        /// </summary>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsHungAppWindow(IntPtr hWnd);

        /// <summary>This process's id, as Windows numbers it.</summary>
        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentProcessId();

        /// <summary>
        /// The window of an Outlook Explorer or Inspector, through <see cref="IOleWindow"/>; zero when Outlook will
        /// not say. It reads the window only - nothing is shown, activated or created.
        /// </summary>
        public static IntPtr WindowOf(object? outlookWindowObject)
        {
            if (outlookWindowObject == null)
            {
                return IntPtr.Zero;
            }

            try
            {
                if (outlookWindowObject is IOleWindow window && window.GetWindow(out IntPtr handle) == 0)
                {
                    return handle;
                }
            }
            catch (Exception ex) when (ex is InvalidCastException || ex is COMException || ex is InvalidComObjectException)
            {
            }

            return IntPtr.Zero;
        }

        /// <summary>
        /// Whether an Outlook Explorer or Inspector is on screen for a person: its window exists and has the visible
        /// style. A compose window an agent's draft is written in without being shown is not.
        /// </summary>
        public static bool IsShown(object? outlookWindowObject)
        {
            IntPtr window = WindowOf(outlookWindowObject);
            return window != IntPtr.Zero && IsWindowVisible(window);
        }

        /// <summary>The OLE window interface every Outlook Explorer and Inspector implements: the one way from one to its window.</summary>
        [ComImport]
        [Guid("00000114-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IOleWindow
        {
            /// <summary>The window.</summary>
            [PreserveSig]
            int GetWindow(out IntPtr phwnd);

            /// <summary>Context-sensitive help, never used here.</summary>
            [PreserveSig]
            int ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool enterMode);
        }
    }
}
