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

        /// <summary>Office's per-user "Office Theme" value, under the Common key of the Office
        /// major that is actually installed.</summary>
        private const string UiThemeValueName = "UI Theme";

        private static bool DetectDarkMode()
        {
            return DecideDarkMode(OfficeVersions.HasOutlookKey, ReadUiTheme, WindowsAppsAreDark);
        }

        /// <summary>
        /// The whole light/dark decision, with the registry replaced by three functions so the
        /// branches this machine cannot produce - an Office major other than the installed one, a
        /// detected Office with no Common key, and nothing detected at all - are reachable
        /// without a second Office install.
        /// <paramref name="outlookKeyExists"/> is the
        /// <see cref="OfficeVersions.TryDetectOutlookVersion(Func{string, bool}, out string)"/>
        /// seam. <paramref name="readUiTheme"/> is handed the FULL HKCU path of the Common key
        /// and returns its <c>UI Theme</c> value, or null when the key or the value is absent.
        /// None of the three may throw.
        /// </summary>
        internal static bool DecideDarkMode(
            Func<string, bool> outlookKeyExists,
            Func<string, object> readUiTheme,
            Func<bool> windowsAppsAreDark)
        {
            string themeKeyPath;
            if (TryGetOfficeThemeKeyPath(outlookKeyExists, out themeKeyPath))
            {
                bool? officeSaysDark = OfficeThemeIsDark(readUiTheme(themeKeyPath));
                if (officeSaysDark.HasValue)
                    return officeSaysDark.Value;
            }

            // DELIBERATE FALLBACK, and it answers three different questions with one value: no
            // Office major was detected, the detected one has no Common key, or Office is set to
            // "use system" (6). All three mean Office has no opinion of its own, and the Windows
            // app mode is then the RIGHT answer rather than a shrug - it is precisely what Office
            // itself follows for 6. What we no longer do is read SOME OTHER major's Common key
            // because it happened to open first: that answered confidently with a setting
            // belonging to an Office that is not the one running, which is the silent-wrong-read
            // shape rather than this deliberate one.
            return windowsAppsAreDark();
        }

        /// <summary>
        /// The HKCU Common key whose UI Theme applies, chosen from the Office version actually
        /// installed rather than from whichever supported major's key opens first. False means
        /// detection found no Office at all; <paramref name="keyPath"/> is empty then, because
        /// there is nothing to read and nothing to watch.
        /// </summary>
        internal static bool TryGetOfficeThemeKeyPath(Func<string, bool> outlookKeyExists, out string keyPath)
        {
            string version;
            if (!OfficeVersions.TryDetectOutlookVersion(outlookKeyExists, out version))
            {
                // Detection hands back OfficeVersions.Fallback here, and this caller deliberately
                // discards it. OutlookTuningService has to write SOMEWHERE and so must take the
                // guess; a read that has a better source available (Windows) should not read
                // 16.0's theme on a machine we could not prove runs 16.0.
                keyPath = string.Empty;
                return false;
            }

            keyPath = OfficeVersions.CommonKeyPath(version);
            return true;
        }

        /// <summary>
        /// Office "Office Theme" values: 0=Colorful, 3=Dark Gray, 4=Black, 5=White, 6=use system.
        /// Only Black is a true dark surface; Colorful/Dark Gray/White keep a light content area.
        /// Null means "Office has no opinion" - 6, a non-int value, or anything unrecognised -
        /// and sends the caller to the Windows app mode.
        /// </summary>
        internal static bool? OfficeThemeIsDark(object uiTheme)
        {
            if (uiTheme is int themeValue)
            {
                if (themeValue == 4)
                    return true;
                if (themeValue == 0 || themeValue == 3 || themeValue == 5)
                    return false;
            }

            return null;
        }

        /// <summary>Live UI Theme read. Null - "no opinion" - for a missing key, a missing value
        /// or any registry failure, because <see cref="DecideDarkMode"/> takes seams that do not
        /// throw.</summary>
        private static object ReadUiTheme(string commonKeyPath)
        {
            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(commonKeyPath))
                {
                    return key == null ? null : key.GetValue(UiThemeValueName);
                }
            }
            catch { return null; }
        }

        /// <summary>Live Windows app-mode read: AppsUseLightTheme 0 = dark. Absent or unreadable
        /// counts as light, which is what this add-in has always defaulted to.</summary>
        private static bool WindowsAppsAreDark()
        {
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

        /// <summary>
        /// Gives up the watch state from the watcher thread itself, for the two cases where there
        /// is nothing to subscribe to. Only if this thread is still the CURRENT watcher: a
        /// StopWatching that already ran, or a later StartWatching, owns the state instead and
        /// must not have it torn out from under it. Whoever clears the state disposes the event,
        /// and both do it holding nothing else, so the two cannot race over it.
        /// Leaving _watching set here would be the worse bug of the two available: StartWatching
        /// would refuse for the rest of the session to start a watcher that never existed.
        /// </summary>
        private static void ReleaseWatchState(ManualResetEvent stop)
        {
            ManualResetEvent toDispose = null;
            lock (_watchGate)
            {
                if (_watching && ReferenceEquals(_stopWatch, stop))
                {
                    _watching = false;
                    _watchThread = null;
                    _stopWatch = null;
                    toDispose = stop;
                }
            }
            if (toDispose != null) { try { toDispose.Dispose(); } catch { } }
        }

        // Receives the stop-event as a captured local so it never reads a field that
        // StopWatching may have nulled (avoids a WaitAny(null) crash on this background thread).
        private static void WatchLoop(ManualResetEvent stop)
        {
            // Watch the Common key of the Office that is actually installed - the same key
            // DetectDarkMode reads, from the same detection, so the watcher cannot end up
            // subscribed to a hive the reader ignores. It used to take whichever supported
            // major's Common key opened first, which on a machine carrying a left-over Common
            // from an earlier Office would have watched a key nothing writes any more: the theme
            // dropdown would then appear to do nothing until the next restart.
            string themeKeyPath = string.Empty;
            bool detected;
            try { detected = TryGetOfficeThemeKeyPath(OfficeVersions.HasOutlookKey, out themeKeyPath); }
            catch { detected = false; }

            if (!detected)
            {
                // No Office detected, so there is no Office theme to follow: Detect() answers
                // from the Windows app mode in this state, and SystemEvents.UserPreferenceChanged
                // already carries changes to THAT. Nothing was opened, so nothing leaks - but the
                // watch state has to go back, or the add-in spends the session believing a
                // watcher is running.
                ReleaseWatchState(stop);
                return;
            }

            IntPtr hKey;
            if (RegOpenKeyEx(HKEY_CURRENT_USER, themeKeyPath, 0, KEY_NOTIFY, out hKey) != 0 || hKey == IntPtr.Zero)
            {
                // Office is installed but has never written its Common key. A failed
                // RegOpenKeyEx assigns no handle, so there is nothing to close here either.
                ReleaseWatchState(stop);
                return;
            }

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
            finally
            {
                RegCloseKey(hKey);
                // Covers the other way out of that loop: RegNotifyChangeKeyValue failing, which
                // ends the watch just as finally as a failed open did. A no-op on the ordinary
                // shutdown path, where StopWatching already owns the state and disposes the
                // event itself once it has joined this thread.
                ReleaseWatchState(stop);
            }
        }
    }
}
