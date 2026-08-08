using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// Thrown when a screenshot could NOT be taken safely - specifically when Outlook did
/// not actually reach the foreground, so the screen rectangle would have contained
/// whatever unrelated windows sat on top of it. Callers treat this as "S5 evidence
/// skipped", never as a test failure: the evidence is a nice-to-have, capturing an
/// unrelated application's window is not.
/// </summary>
internal sealed class ScreenCaptureSkippedException : Exception
{
    /// <summary>Creates the skip signal.</summary>
    public ScreenCaptureSkippedException(string message)
        : base(message)
    {
    }
}

/// <summary>
/// Screenshot evidence helper for the S5 cases (v3.MD section 0.5.1): captures ONLY the
/// Outlook main window's screen rectangle - never the full desktop, so no other
/// application's content can leak into the image. Callers are responsible for the S5
/// content rules (view scoped to the test-hub store, navigation pane hidden) and for
/// saving into the gitignored screenshots directory (S6).
/// <para>
/// ⚠ THE RECT PROBLEM AND ITS FIX (soak fix 18 found it, soak fix 19 fixed it): this
/// grabs a SCREEN RECTANGLE, so it captures whatever pixels sit at those coordinates -
/// not Outlook's own bitmap. <c>SetForegroundWindow</c> is a REQUEST Windows routinely
/// denies (the caller must already own the foreground or have input focus), and it lost
/// that race repeatedly during the batch-B run: two captures contained an unrelated file
/// manager and editor. Every capture now VERIFIES through <c>GetForegroundWindow</c>
/// that the target really is frontmost - BEFORE and AFTER the grab, since a notification
/// can steal focus mid-capture - retries a few times, and throws
/// <see cref="ScreenCaptureSkippedException"/> (deleting any partial file) instead of
/// writing a polluted PNG. No image at all is strictly better than an image of something
/// that was never ours to capture.
/// </para>
/// </summary>
internal static class ScreenCapture
{
    private const string OutlookExplorerWindowClass = "rctrl_renwnd32";

    /// <summary>How many times the foreground is requested before giving up.</summary>
    private const int ForegroundAttempts = 4;

    /// <summary>Settle time after a z-order change, per attempt (ms).</summary>
    private const int RepaintDelayMs = 500;

    /// <summary>
    /// Finds the Outlook Explorer window (by caption when possible, else the largest
    /// visible window of Outlook's Explorer class), brings it to the foreground and
    /// captures its rectangle to a PNG. Returns the absolute file path. Throws
    /// <see cref="ScreenCaptureSkippedException"/> when Outlook could not be made
    /// frontmost.
    /// </summary>
    public static string CaptureOutlookWindow(string? caption, string targetDirectory, string fileName)
    {
        IntPtr window = FindOutlookWindow(caption);
        if (window == IntPtr.Zero)
        {
            throw new InvalidOperationException("No visible Outlook Explorer window found to capture.");
        }

        return CaptureVerifiedForeground(window, targetDirectory, fileName);
    }

    /// <summary>
    /// Captures the visible Outlook window whose caption CONTAINS
    /// <paramref name="captionFragment"/> (Inspector windows share Outlook's window
    /// class with Explorers, so the Phase-4 draft screenshot targets the compose window
    /// by its unique tagged subject). Throws when no such window is on screen, and
    /// <see cref="ScreenCaptureSkippedException"/> when it will not come forward.
    /// </summary>
    public static string CaptureOutlookWindowByCaptionFragment(string captionFragment, string targetDirectory, string fileName)
    {
        if (string.IsNullOrWhiteSpace(captionFragment))
        {
            throw new ArgumentException("Caption fragment required.", nameof(captionFragment));
        }

        IntPtr window = FindOutlookWindowByCaptionFragment(captionFragment);
        if (window == IntPtr.Zero)
        {
            throw new InvalidOperationException("No visible Outlook window with the requested caption fragment found.");
        }

        return CaptureVerifiedForeground(window, targetDirectory, fileName);
    }

    /// <summary>
    /// True when any VISIBLE Outlook-class window's caption contains the fragment -
    /// the never-displayed assert for the identity drafts (Q-it2-3a).
    /// </summary>
    public static bool AnyVisibleOutlookWindowWithCaptionFragment(string captionFragment)
    {
        return FindOutlookWindowByCaptionFragment(captionFragment) != IntPtr.Zero;
    }

    /// <summary>
    /// The single capture path, shared by both entry points: request the foreground,
    /// PROVE the window got it, grab its rectangle, then prove it still had the
    /// foreground afterwards. Any failure of that proof means the pixels may belong to
    /// another application, so no file is left behind.
    /// </summary>
    private static string CaptureVerifiedForeground(IntPtr window, string targetDirectory, string fileName)
    {
        for (int attempt = 1; attempt <= ForegroundAttempts; attempt++)
        {
            NativeMethods.SetForegroundWindow(window);
            Thread.Sleep(RepaintDelayMs); // repaint after the z-order change

            if (NativeMethods.GetForegroundWindow() != window)
            {
                continue;
            }

            // Prefer the DWM extended frame bounds: GetWindowRect includes the invisible
            // resize border, whose pixels belong to whatever window lies BEHIND Outlook -
            // the frame bounds keep foreign content out of the capture entirely.
            NativeMethods.RECT rect;
            int hr = NativeMethods.DwmGetWindowAttribute(
                window, NativeMethods.DWMWA_EXTENDED_FRAME_BOUNDS, out rect, Marshal.SizeOf<NativeMethods.RECT>());
            if (hr != 0 && !NativeMethods.GetWindowRect(window, out rect))
            {
                throw new InvalidOperationException("GetWindowRect failed for the Outlook window.");
            }

            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;
            if (width <= 0 || height <= 0)
            {
                throw new InvalidOperationException($"Outlook window has an empty rectangle ({width}x{height}).");
            }

            Directory.CreateDirectory(targetDirectory);
            string path = Path.Combine(targetDirectory, fileName);
            using (Bitmap bitmap = new(width, height))
            {
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.CopyFromScreen(rect.Left, rect.Top, 0, 0, new Size(width, height));
                }

                // Re-verify BEFORE writing: a toast or an installer dialog can appear
                // between the foreground check and CopyFromScreen, and that is exactly
                // what would land in the image.
                if (NativeMethods.GetForegroundWindow() != window)
                {
                    continue;
                }

                bitmap.Save(path, ImageFormat.Png);
            }

            return path;
        }

        throw new ScreenCaptureSkippedException(
            "Outlook did not hold the foreground for the capture after " + ForegroundAttempts
            + " attempts, so the screen rectangle could have contained unrelated windows. No screenshot was written.");
    }

    private static IntPtr FindOutlookWindowByCaptionFragment(string captionFragment)
    {
        IntPtr found = IntPtr.Zero;
        NativeMethods.EnumWindows((hwnd, _) =>
        {
            if (!NativeMethods.IsWindowVisible(hwnd))
            {
                return true;
            }

            var className = new System.Text.StringBuilder(64);
            _ = NativeMethods.GetClassNameW(hwnd, className, className.Capacity);
            if (!string.Equals(className.ToString(), OutlookExplorerWindowClass, StringComparison.Ordinal))
            {
                return true;
            }

            var caption = new System.Text.StringBuilder(512);
            _ = NativeMethods.GetWindowTextW(hwnd, caption, caption.Capacity);
            if (caption.ToString().IndexOf(captionFragment, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                found = hwnd;
                return false;
            }

            return true;
        }, IntPtr.Zero);

        return found;
    }

    private static IntPtr FindOutlookWindow(string? caption)
    {
        if (!string.IsNullOrEmpty(caption))
        {
            IntPtr exact = NativeMethods.FindWindowW(OutlookExplorerWindowClass, caption);
            if (exact != IntPtr.Zero && NativeMethods.IsWindowVisible(exact))
            {
                return exact;
            }
        }

        // Fallback: the largest visible window of Outlook's Explorer window class
        // (captions can change asynchronously while a search is populating).
        IntPtr best = IntPtr.Zero;
        long bestArea = 0;
        NativeMethods.EnumWindows((hwnd, _) =>
        {
            if (!NativeMethods.IsWindowVisible(hwnd))
            {
                return true;
            }

            var className = new System.Text.StringBuilder(64);
            _ = NativeMethods.GetClassNameW(hwnd, className, className.Capacity);
            if (!string.Equals(className.ToString(), OutlookExplorerWindowClass, StringComparison.Ordinal))
            {
                return true;
            }

            if (NativeMethods.GetWindowRect(hwnd, out NativeMethods.RECT r))
            {
                long area = (long)(r.Right - r.Left) * (r.Bottom - r.Top);
                if (area > bestArea)
                {
                    bestArea = area;
                    best = hwnd;
                }
            }

            return true;
        }, IntPtr.Zero);

        return best;
    }

    private static class NativeMethods
    {
        [StructLayout(LayoutKind.Sequential)]
        internal struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        internal delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        internal static extern IntPtr FindWindowW(string? lpClassName, string? lpWindowName);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        internal static extern int GetClassNameW(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        internal static extern int GetWindowTextW(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        internal const int DWMWA_EXTENDED_FRAME_BOUNDS = 9;

        [DllImport("dwmapi.dll")]
        internal static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out RECT pvAttribute, int cbAttribute);
    }
}
