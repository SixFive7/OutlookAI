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

        /// <summary>User-hive Outlook search key carrying DisableServerAssistedSearch (D22; the tuning Search group writes here).</summary>
        public const string OutlookSearchUserKeyPath = @"Software\Microsoft\Office\16.0\Outlook\Search";

        /// <summary>Policy-hive Outlook search key - authoritative over the user hive when its value exists (ADMX-managed).</summary>
        public const string OutlookSearchPolicyKeyPath = @"Software\Policies\Microsoft\Office\16.0\Outlook\Search";

        /// <summary>uiSearchBackend value: Outlook UI search queries the local SystemIndex - the same corpus agent search uses.</summary>
        public const string UiSearchBackendLocal = "local";

        /// <summary>uiSearchBackend value: Outlook UI search goes through the Exchange service search (server-capped, differently ranked - may diverge from agent search).</summary>
        public const string UiSearchBackendServerAssisted = "server-assisted";

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

        /// <summary>
        /// Reads the tuning summary from the live HKCU registry state. Also stamps the
        /// EFFECTIVE UI search backend - a live-registry fact independent of desired
        /// state, reported even when the add-in never initialized tuning here.
        /// </summary>
        public static TuningHealthView ReadTuningStateFromRegistry()
        {
            TuningHealthView view;
            try
            {
                using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(TuningKeyPath, writable: false))
                {
                    view = key == null
                        ? new TuningHealthView { Managed = false }
                        : ReadTuningState(name => key.GetValue(name));
                }
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                view = new TuningHealthView { Managed = false };
            }

            view.UiSearchBackend = ReadUiSearchBackendFromRegistry();
            return view;
        }

        /// <summary>
        /// Maps the DisableServerAssistedSearch DWORDs to the effective UI search backend
        /// (pure, T1-pinned). The policy-hive value is authoritative when present; with
        /// neither hive set (or an explicit 0) Outlook uses its server-assisted default.
        /// </summary>
        public static string DescribeUiSearchBackend(int? policyValue, int? userValue)
        {
            int effective = (policyValue ?? userValue) ?? 0;
            return effective != 0 ? UiSearchBackendLocal : UiSearchBackendServerAssisted;
        }

        /// <summary>
        /// EFFECTIVE Outlook UI search backend from the live registry (D22 coupling made
        /// visible, D35): "local" when DisableServerAssistedSearch is in force - Outlook's
        /// search box then queries the same SystemIndex corpus agent search uses -
        /// "server-assisted" when the value is absent/0 (UI results server-capped and
        /// differently ranked, so they can diverge from agent search). Never throws.
        /// </summary>
        public static string ReadUiSearchBackendFromRegistry()
        {
            return DescribeUiSearchBackend(
                TryReadCurrentUserDword(OutlookSearchPolicyKeyPath, "DisableServerAssistedSearch"),
                TryReadCurrentUserDword(OutlookSearchUserKeyPath, "DisableServerAssistedSearch"));
        }

        /// <summary>HKCU DWORD read that treats absent keys/values/non-DWORD data and failures as null.</summary>
        private static int? TryReadCurrentUserDword(string keyPath, string valueName)
        {
            try
            {
                using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(keyPath, writable: false))
                {
                    return key?.GetValue(valueName) as int?;
                }
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                return null;
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

        // ===== MCP registration (Phase 8) =====

        /// <summary>Registry path of the add-in's registration state (mirrors McpRegistrationService.McpKeyPath).</summary>
        public const string McpRegistrationKeyPath = @"Software\OutlookAI\Mcp";

        /// <summary>registration.status: the registered command IS the running executable.</summary>
        public const string RegistrationOk = "ok";

        /// <summary>registration.status: an entry exists but names a different executable.</summary>
        public const string RegistrationDrifted = "drifted";

        /// <summary>registration.status: no outlookai entry at all.</summary>
        public const string RegistrationAbsent = "absent";

        /// <summary>registration.status: the config exists but could not be parsed (so it is never rewritten).</summary>
        public const string RegistrationUnreadable = "unreadable";

        /// <summary>registration.status: the config could not be examined at all.</summary>
        public const string RegistrationUnknown = "unknown";

        /// <summary>
        /// Compares the registered command against the running executable (pure, T1-pinned).
        /// Path comparison, not string comparison: the registration and the process path can
        /// differ only in casing or separators and still be the same file.
        /// </summary>
        public static string DescribeMcpRegistration(string? registeredCommand, string? runningFrom)
        {
            if (string.IsNullOrWhiteSpace(registeredCommand))
            {
                return RegistrationAbsent;
            }

            if (string.IsNullOrWhiteSpace(runningFrom))
            {
                return RegistrationUnknown;
            }

            return SamePath(registeredCommand!, runningFrom!) ? RegistrationOk : RegistrationDrifted;
        }

        /// <summary>Whether two paths name the same file, tolerating case and separator differences.</summary>
        public static bool SamePath(string a, string b)
        {
            try
            {
                return string.Equals(
                    Path.GetFullPath(a).TrimEnd('\\'),
                    Path.GetFullPath(b).TrimEnd('\\'),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
            }
        }

        /// <summary>
        /// The observed registration plus what the add-in last recorded. Never throws, and
        /// never writes: health only ever reports on this file, the add-in owns repairing it.
        /// </summary>
        public static McpRegistrationHealthView ReadMcpRegistration(string? runningFrom)
        {
            var view = new McpRegistrationHealthView { RunningFrom = runningFrom };

            bool readable;
            view.RegisteredCommand = TryReadRegisteredCommand(out readable);
            view.Status = readable
                ? DescribeMcpRegistration(view.RegisteredCommand, runningFrom)
                : RegistrationUnreadable;

            try
            {
                using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(McpRegistrationKeyPath, writable: false))
                {
                    if (key != null)
                    {
                        view.AddInStatus = key.GetValue("Status") as string;
                        view.AddInLastReconcileUtc = key.GetValue("LastReconcileUtc") as string;
                        view.AddInResolvedServerPath = key.GetValue("ResolvedServerPath") as string;
                        object? healed = key.GetValue("Healed");
                        if (healed is int healedInt)
                        {
                            view.AddInHealed = healedInt != 0;
                        }
                    }
                }
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                // Leave the add-in fields null; the observed status above still stands.
            }

            if (string.IsNullOrEmpty(view.AddInStatus)) view.AddInStatus = null;
            if (string.IsNullOrEmpty(view.AddInLastReconcileUtc)) view.AddInLastReconcileUtc = null;
            if (string.IsNullOrEmpty(view.AddInResolvedServerPath)) view.AddInResolvedServerPath = null;

            return view;
        }

        /// <summary>
        /// mcpServers.outlookai.command from ~/.claude.json. <paramref name="readable"/> is
        /// false only when the file exists but could not be read or understood - an absent
        /// file is perfectly readable and simply has no entry.
        /// </summary>
        private static string? TryReadRegisteredCommand(out bool readable)
        {
            readable = true;
            try
            {
                string path = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                    ".claude.json");

                if (!File.Exists(path))
                {
                    return null;
                }

                // FileShare.ReadWrite: the Claude Code CLI owns this file and may hold it open.
                string text;
                using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var reader = new StreamReader(stream))
                {
                    text = reader.ReadToEnd();
                }

                return ExtractRegisteredCommand(text, out readable);
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                readable = false;
                return null;
            }
        }

        /// <summary>
        /// Pulls mcpServers -> outlookai -> command out of raw JSON text (pure, T1-pinned).
        /// Hand-scanned rather than deserialized so this compiles on BOTH Core targets:
        /// System.Text.Json is not available on net48, and Core takes no JSON dependency
        /// (v3.MD D18 v2 / section 0.5.2 - Core must stay host-neutral and reference-light).
        /// Read-only by construction: nothing here can modify the file.
        /// <paramref name="readable"/> is false when the text is not a JSON object at all.
        /// </summary>
        public static string? ExtractRegisteredCommand(string? json, out bool readable)
        {
            readable = true;
            if (string.IsNullOrWhiteSpace(json))
            {
                return null;
            }

            int i = SkipWhitespace(json!, 0);
            if (i >= json!.Length || json[i] != '{')
            {
                readable = false;
                return null;
            }

            int servers = FindMemberValue(json, i, "mcpServers");
            if (servers < 0 || json[servers] != '{')
            {
                return null;
            }

            int entry = FindMemberValue(json, servers, "outlookai");
            if (entry < 0 || json[entry] != '{')
            {
                return null;
            }

            int command = FindMemberValue(json, entry, "command");
            if (command < 0 || json[command] != '"')
            {
                return null;
            }

            // (FindMemberValue never returns an index past the end, so the reads above are
            // in bounds; truncated input yields -1 instead.)

            int end = SkipString(json, command);
            return end < 0 ? null : Unescape(json.Substring(command + 1, end - command - 2));
        }

        /// <summary>
        /// Start index of the value of <paramref name="name"/> among the DIRECT members of the
        /// object beginning at <paramref name="objectStart"/>, or -1. Members are walked one at
        /// a time and each value is skipped whole, so a same-named key nested deeper can never
        /// be mistaken for a direct member.
        /// </summary>
        private static int FindMemberValue(string json, int objectStart, string name)
        {
            int i = SkipWhitespace(json, objectStart + 1);
            if (i < json.Length && json[i] == '}')
            {
                return -1;
            }

            while (i < json.Length)
            {
                if (json[i] != '"')
                {
                    return -1;
                }

                int keyEnd = SkipString(json, i);
                if (keyEnd < 0)
                {
                    return -1;
                }

                string key = Unescape(json.Substring(i + 1, keyEnd - i - 2));

                i = SkipWhitespace(json, keyEnd);
                if (i >= json.Length || json[i] != ':')
                {
                    return -1;
                }

                int valueStart = SkipWhitespace(json, i + 1);
                if (valueStart >= json.Length)
                {
                    // Truncated right after the colon: no value to point at.
                    return -1;
                }

                if (key == name)
                {
                    return valueStart;
                }

                int valueEnd = SkipValue(json, valueStart);
                if (valueEnd < 0)
                {
                    return -1;
                }

                i = SkipWhitespace(json, valueEnd);
                if (i >= json.Length || json[i] != ',')
                {
                    return -1;
                }

                i = SkipWhitespace(json, i + 1);
            }

            return -1;
        }

        private static int SkipWhitespace(string json, int i)
        {
            while (i < json.Length && char.IsWhiteSpace(json[i]))
            {
                i++;
            }

            return i;
        }

        /// <summary>Index just past the closing quote of the string starting at i, or -1.</summary>
        private static int SkipString(string json, int i)
        {
            i++;
            while (i < json.Length)
            {
                char c = json[i];
                if (c == '\\')
                {
                    i += 2;
                    continue;
                }

                if (c == '"')
                {
                    return i + 1;
                }

                i++;
            }

            return -1;
        }

        /// <summary>Index just past the end of the value starting at i, or -1.</summary>
        private static int SkipValue(string json, int i)
        {
            if (i >= json.Length)
            {
                return -1;
            }

            char c = json[i];
            if (c == '"')
            {
                return SkipString(json, i);
            }

            if (c == '{' || c == '[')
            {
                int depth = 0;
                while (i < json.Length)
                {
                    char d = json[i];
                    if (d == '"')
                    {
                        int end = SkipString(json, i);
                        if (end < 0)
                        {
                            return -1;
                        }

                        i = end;
                        continue;
                    }

                    if (d == '{' || d == '[')
                    {
                        depth++;
                    }
                    else if (d == '}' || d == ']')
                    {
                        depth--;
                        if (depth == 0)
                        {
                            return i + 1;
                        }
                    }

                    i++;
                }

                return -1;
            }

            while (i < json.Length)
            {
                char d = json[i];
                if (d == ',' || d == '}' || d == ']' || char.IsWhiteSpace(d))
                {
                    return i;
                }

                i++;
            }

            return i;
        }

        private static string Unescape(string s)
        {
            if (s.IndexOf('\\') < 0)
            {
                return s;
            }

            var sb = new System.Text.StringBuilder(s.Length);
            for (int i = 0; i < s.Length; i++)
            {
                if (s[i] != '\\' || i + 1 >= s.Length)
                {
                    sb.Append(s[i]);
                    continue;
                }

                i++;
                char c = s[i];
                switch (c)
                {
                    case 'n': sb.Append('\n'); break;
                    case 'r': sb.Append('\r'); break;
                    case 't': sb.Append('\t'); break;
                    case 'b': sb.Append('\b'); break;
                    case 'f': sb.Append('\f'); break;
                    case 'u':
                        if (i + 4 < s.Length
                            && int.TryParse(s.Substring(i + 1, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int code))
                        {
                            sb.Append((char)code);
                            i += 4;
                        }
                        else
                        {
                            sb.Append(c);
                        }

                        break;
                    default: sb.Append(c); break;
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// Path of the executable hosting this process - what a registration must name to
        /// spawn this server. Not Environment.ProcessPath: that does not exist on net48, and
        /// Core compiles for both targets (R10).
        /// </summary>
        public static string? CurrentProcessPath()
        {
            try
            {
                using (Process process = Process.GetCurrentProcess())
                {
                    return process.MainModule?.FileName;
                }
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                return null;
            }
        }
    }
}
