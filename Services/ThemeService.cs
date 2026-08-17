using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32;

namespace OutlookAI.Services
{
    public static class ThemeService
    {
        public static bool IsDarkMode { get; private set; }

        public static Color Background { get; private set; }
        public static Color ControlBackground { get; private set; }
        public static Color TextBoxBackground { get; private set; }
        public static Color Text { get; private set; }
        public static Color SecondaryText { get; private set; }
        public static Color Accent { get; private set; }
        public static Color Border { get; private set; }
        public static Color StatusError { get; private set; }
        public static Color StatusSuccess { get; private set; }
        public static Color LinkError { get; private set; }
        public static Color ButtonFace { get; private set; }
        public static Color ButtonText { get; private set; }
        public static Color ResultBackground { get; private set; }

        static ThemeService()
        {
            SetLightDefaults();
            try { Detect(); } catch { }

            try
            {
                // Re-detect on ANY preference change. Windows dark/light app-mode toggles
                // arrive under inconsistent categories (General/Color/VisualStyle), so we must
                // not filter by category. Detect() is cheap and ThemeChanged fires only on a
                // real light<->dark flip.
                SystemEvents.UserPreferenceChanged += (s, e) =>
                {
                    try { Detect(); } catch { }
                };
            }
            catch { }
        }

        private static void SetLightDefaults()
        {
            IsDarkMode = false;
            Background = Color.FromArgb(250, 249, 248);
            ControlBackground = SystemColors.Control;
            TextBoxBackground = SystemColors.Window;
            Text = SystemColors.ControlText;
            SecondaryText = Color.Gray;
            Accent = Color.FromArgb(0, 120, 212);
            Border = SystemColors.ControlDark;
            StatusError = Color.DarkRed;
            StatusSuccess = Color.DarkGreen;
            LinkError = Color.IndianRed;
            ButtonFace = SystemColors.ButtonFace;
            ButtonText = SystemColors.ControlText;
            ResultBackground = SystemColors.Window;
        }

        /// <summary>Raised when the detected theme actually flips (light&lt;-&gt;dark) at runtime.
        /// May fire on a non-UI thread (driven by SystemEvents.UserPreferenceChanged).</summary>
        public static event EventHandler ThemeChanged;

        public static void Detect()
        {
            ApplyMode(DetectDarkMode());
        }

        private static void ApplyMode(bool dark)
        {
            bool changed = dark != IsDarkMode;
            IsDarkMode = dark;

            if (dark)
            {
                Background = Color.FromArgb(32, 32, 32);
                ControlBackground = Color.FromArgb(45, 45, 48);
                TextBoxBackground = Color.FromArgb(51, 51, 55);
                Text = Color.FromArgb(230, 230, 230);
                SecondaryText = Color.FromArgb(160, 160, 160);
                Accent = Color.FromArgb(75, 156, 245);
                Border = Color.FromArgb(70, 70, 74);
                StatusError = Color.FromArgb(255, 120, 120);
                StatusSuccess = Color.FromArgb(120, 220, 120);
                LinkError = Color.FromArgb(255, 140, 140);
                ButtonFace = Color.FromArgb(55, 55, 60);
                ButtonText = Color.FromArgb(230, 230, 230);
                ResultBackground = Color.FromArgb(40, 40, 44);
            }
            else
            {
                SetLightDefaults();
            }

            // Notify open panes only when the theme actually changed (not on every
            // unrelated UserPreferenceChanged). Isolate handlers so one can't break the rest.
            if (changed)
            {
                var handler = ThemeChanged;
                if (handler != null)
                {
                    foreach (EventHandler h in handler.GetInvocationList())
                    {
                        try { h(null, EventArgs.Empty); } catch { }
                    }
                }
            }
        }

        private static bool DetectDarkMode()
        {
            try
            {
                foreach (var ver in OfficeVersions.Supported)
                {
                    using (var key = Registry.CurrentUser.OpenSubKey($@"SOFTWARE\Microsoft\Office\{ver}\Common"))
                    {
                        var theme = key?.GetValue("UI Theme");
                        if (theme is int themeValue)
                        {
                            // Office "Office Theme" values: 0=Colorful, 3=Dark Gray, 4=Black,
                            // 5=White, 6=use system. Only "Black" is a true dark surface;
                            // Colorful/Dark Gray/White keep a light content area. 6 (and any
                            // other value) falls through to the Windows app-mode check below.
                            if (themeValue == 4)
                                return true;
                            if (themeValue == 0 || themeValue == 3 || themeValue == 5)
                                return false;
                        }
                    }
                }
            }
            catch { }

            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    var value = key?.GetValue("AppsUseLightTheme");
                    if (value is int lightTheme)
                        return lightTheme == 0;
                }
            }
            catch { }

            return false;
        }

        // ===== Live watch for the Office "Office Theme" dropdown =====
        // Office theme changes write HKCU\...\Office\<ver>\Common\UI Theme but do NOT raise
        // SystemEvents.UserPreferenceChanged, so we watch the key directly and re-Detect on change.
        private static readonly object _watchGate = new object();
        private static Thread _watchThread;
        private static ManualResetEvent _stopWatch;
        private static volatile bool _watching;

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern int RegOpenKeyEx(IntPtr hKey, string subKey, int options, int samDesired, out IntPtr phkResult);
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern int RegNotifyChangeKeyValue(IntPtr hKey, bool watchSubtree, int notifyFilter, IntPtr hEvent, bool asynchronous);
        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern int RegCloseKey(IntPtr hKey);
        private static readonly IntPtr HKEY_CURRENT_USER = new IntPtr(unchecked((int)0x80000001));
        private const int KEY_NOTIFY = 0x0010;
        private const int REG_NOTIFY_CHANGE_LAST_SET = 0x0004;

        public static void StartWatching()
        {
            lock (_watchGate)
            {
                if (_watching) return;
                _watching = true;
                var stop = new ManualResetEvent(false);
                _stopWatch = stop;
                _watchThread = new Thread(() => WatchLoop(stop)) { IsBackground = true, Name = "OfficeThemeWatcher" };
                _watchThread.Start();
            }
        }

        public static void StopWatching()
        {
            Thread t;
            ManualResetEvent stop;
            lock (_watchGate)
            {
                if (!_watching) return;
                _watching = false;
                t = _watchThread; stop = _stopWatch;
                _watchThread = null; _stopWatch = null;
            }
            try { stop?.Set(); } catch { }
            bool joined = true;
            try { joined = t == null || t.Join(2000); } catch { }
            if (joined) { try { stop?.Dispose(); } catch { } }
        }

        // Receives the stop-event as a captured local so it never reads a field that
        // StopWatching may have nulled (avoids a WaitAny(null) crash on this background thread).
        private static void WatchLoop(ManualResetEvent stop)
        {
            IntPtr hKey = IntPtr.Zero;
            foreach (var ver in OfficeVersions.Supported)
            {
                if (RegOpenKeyEx(HKEY_CURRENT_USER, $@"SOFTWARE\Microsoft\Office\{ver}\Common", 0, KEY_NOTIFY, out hKey) == 0 && hKey != IntPtr.Zero)
                    break;
                hKey = IntPtr.Zero;
            }
            if (hKey == IntPtr.Zero) return;

            try
            {
                using (var changed = new AutoResetEvent(false))
                {
                    var handles = new WaitHandle[] { changed, stop };
                    while (_watching)
                    {
                        if (RegNotifyChangeKeyValue(hKey, false, REG_NOTIFY_CHANGE_LAST_SET, changed.SafeWaitHandle.DangerousGetHandle(), true) != 0)
                            break;
                        if (WaitHandle.WaitAny(handles) == 1 || !_watching)
                            break;
                        try { Detect(); } catch { }
                    }
                }
            }
            finally { RegCloseKey(hKey); }
        }
    }
}
