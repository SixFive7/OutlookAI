using System.Runtime.InteropServices;
using System.Text;
using OutlookAI.Core.AddIn;

namespace OutlookAI.McpServer.Tests;

/// <summary>
/// The Win32 calls the live tests make that the product does not, declared once: the screen capture's and
/// the visible-window census's, and those the pane tests, the compose probe and the restart helper each
/// declared for themselves (the code-sharing refactor's S5a, the survey's A09). The window calls the product
/// makes too are Core's <see cref="NativeWindows"/>, which the tests call; a call one test file alone makes
/// stays in that file.
/// </summary>
internal static class TestWin32
{
    /// <summary>DwmGetWindowAttribute's window frame as drawn, a <see cref="NativeWindows.RECT"/>.</summary>
    internal const int DWMWA_EXTENDED_FRAME_BOUNDS = 9;

    /// <summary>Non-zero when DWM cloaks the window: drawn nowhere, though IsWindowVisible says TRUE.</summary>
    internal const int DWMWA_CLOAKED = 14;

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern IntPtr FindWindowW(string? lpClassName, string? lpWindowName);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    internal static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool EnumChildWindows(IntPtr parent, NativeWindows.EnumWindowsProc callback, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern int GetWindowTextW(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    internal static extern int GetWindowTextLengthW(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    internal static extern bool PostMessageW(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

    [DllImport("dwmapi.dll")]
    internal static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out NativeWindows.RECT pvAttribute, int cbAttribute);

    [DllImport("dwmapi.dll")]
    internal static extern int DwmGetWindowAttribute(IntPtr hwnd, int dwAttribute, out int pvAttribute, int cbAttribute);
}
