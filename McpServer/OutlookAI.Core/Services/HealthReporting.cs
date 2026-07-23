using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;

using Microsoft.Win32;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// Probe + mapping helpers behind the <c>health</c> tool (Phase 7). The pure
    /// mapping functions are T1-tested; the machine probes are small, content-free and
    /// never throw (health must always produce a report). Registry-only where possible:
    /// the tuning summary reads the add-in's HKCU state directly so the server stays
    /// decoupled from add-in code (R12/section 0.5.3).
    /// </summary>
    public static class HealthReporting
    {
        /// <summary>Registry path of the add-in's tuning state (mirrors OutlookTuningService.TuningKeyPath).</summary>
        public const string TuningKeyPath = @"Software\OutlookAI\Tuning";

        private const string WSearchServiceKeyPath = @"SYSTEM\CurrentControlSet\Services\WSearch";
        private const string OutlookAppPathKeyPath = @"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE";

        /// <summary>
        /// Maps a Windows service Start registry value to a compact mode name (pure,
        /// T1-pinned): 2 automatic, 3 manual, 4 disabled.
        /// </summary>
        public static string DescribeServiceStartMode(int? startValue)
        {
            switch (startValue)
            {
                case 2:
                    return "automatic";
                case 3:
                    return "manual";
                case 4:
                    return "disabled";
                case null:
                    return "unknown";
                default:
                    return "other(" + startValue.Value.ToString(CultureInfo.InvariantCulture) + ")";
            }
        }

        /// <summary>
        /// Builds the tuning summary from a value reader (pure mapping, T1-tested with a
        /// fabricated reader; production passes a registry-backed one). A missing
        /// Initialized value means the add-in never ran its tuning service here.
        /// </summary>
        public static TuningHealthView ReadTuningState(Func<string, object?> readValue)
        {
            if (readValue == null)
            {
                throw new ArgumentNullException(nameof(readValue));
            }

            if (AsBool(readValue("Initialized")) != true)
            {
                return new TuningHealthView { Managed = false };
            }

            string? conflicts = readValue("PolicyConflicts") as string;
            return new TuningHealthView
            {
                Managed = true,
                Enabled = AsBool(readValue("Enabled")),
                SearchEnabled = AsBool(readValue("SearchEnabled")),
                CachingEnabled = AsBool(readValue("CachingEnabled")),
                OstEnabled = AsBool(readValue("OstEnabled")),
                RestartNeeded = AsBool(readValue("RestartNeeded")),
                PolicyConflicts = string.IsNullOrWhiteSpace(conflicts) ? null : conflicts,
                LastReconcileUtc = readValue("LastReconcileUtc") as string,
            };
        }

        /// <summary>Reads the tuning summary from the live HKCU registry state.</summary>
        public static TuningHealthView ReadTuningStateFromRegistry()
        {
            try
            {
                using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(TuningKeyPath, writable: false))
                {
                    if (key == null)
                    {
                        return new TuningHealthView { Managed = false };
                    }

                    return ReadTuningState(name => key.GetValue(name));
                }
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                return new TuningHealthView { Managed = false };
            }
        }

        /// <summary>WSearch service Start value from HKLM (null when unreadable/absent).</summary>
        public static int? TryReadWSearchStartValue()
        {
            try
            {
                using (RegistryKey? key = Registry.LocalMachine.OpenSubKey(WSearchServiceKeyPath, writable: false))
                {
                    return key?.GetValue("Start") as int?;
                }
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                return null;
            }
        }

        /// <summary>Whether a process with the given image name (no extension) is running (null = probe failed).</summary>
        public static bool? TryIsProcessRunning(string processName)
        {
            try
            {
                Process[] processes = Process.GetProcessesByName(processName);
                try
                {
                    return processes.Length > 0;
                }
                finally
                {
                    foreach (Process process in processes)
                    {
                        process.Dispose();
                    }
                }
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                return null;
            }
        }

        /// <summary>
        /// Whether the running OUTLOOK.EXE is HEADLESS - no main window, tray icon only
        /// (SF-3: the D17 autostart state; empirically mapped 2026-07-23). Null when
        /// Outlook is not running or the probe failed; false when a window exists (a
        /// user session, or a headless one promoted by a normal launch).
        /// </summary>
        public static bool? TryGetOutlookHeadless()
        {
            try
            {
                Process[] processes = Process.GetProcessesByName("OUTLOOK");
                try
                {
                    if (processes.Length == 0)
                    {
                        return null;
                    }

                    return processes[0].MainWindowHandle == IntPtr.Zero;
                }
                finally
                {
                    foreach (Process process in processes)
                    {
                        process.Dispose();
                    }
                }
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                return null;
            }
        }

        /// <summary>Registry DWORD (boxed int) to bool: nonzero true, zero false, anything else null.</summary>
        private static bool? AsBool(object? value)
        {
            if (value is int number)
            {
                return number != 0;
            }

            return null;
        }

        /// <summary>
        /// Installed classic-Outlook build via the App Paths registration + file version
        /// (no COM - works with Outlook closed; null when not installed/readable).
        /// </summary>
        public static string? TryGetOutlookVersion()
        {
            try
            {
                string? exePath;
                using (RegistryKey? key = Registry.LocalMachine.OpenSubKey(OutlookAppPathKeyPath, writable: false))
                {
                    exePath = key?.GetValue(null) as string;
                }

                if (string.IsNullOrWhiteSpace(exePath))
                {
                    return null;
                }

                exePath = exePath!.Trim('"');
                if (!File.Exists(exePath))
                {
                    return null;
                }

                return FileVersionInfo.GetVersionInfo(exePath).FileVersion;
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                return null;
            }
        }
    }
}
