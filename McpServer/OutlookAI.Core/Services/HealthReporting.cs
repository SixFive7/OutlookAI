using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;

using Microsoft.Win32;

// The one definition of everything the add-in and this server have to agree about: the HKCU keys
// the add-in writes and this file reads, and the Claude Code configuration both of them look at.
// Core's AddIn\AddInServerContract.cs, which the add-in takes them from too (Q163 (a), Q167 (a)),
// so these are not copies of the add-in's constants - they are the add-in's constants.
using Contract = global::OutlookAI.Core.AddIn.AddInServerContract;
// And the Office majors, plus the hive paths built from whichever one this machine has.
// Core's AddIn\OfficeVersions.cs is the add-in's too, so the Outlook Search key this file reads is
// built by the same expression the add-in writes it with - see OutlookSearchUserKeyPath below.
using OfficeVersions = global::OutlookAI.Core.AddIn.OfficeVersions;
using AddInServerContract = global::OutlookAI.Core.AddIn.AddInServerContract;
using McpConfigEditor = global::OutlookAI.Core.AddIn.McpConfigEditor;

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
        /// <summary>
        /// Registry path of the add-in's tuning state. NOT a copy of the add-in's constant - the
        /// same one, <see cref="AddInServerContract.TuningKeyPath"/>, which the add-in writes with,
        /// so there is nothing here to keep in step.
        /// </summary>
        public const string TuningKeyPath = Contract.TuningKeyPath;

        /// <summary>
        /// The value name this file reads out of both Search keys below - NOT a copy of the
        /// add-in's: it is <see cref="AddInServerContract.DisableServerAssistedSearchValueName"/>,
        /// so the name the tuning service writes and the name read back here are one constant.
        /// </summary>
        public const string DisableServerAssistedSearchValueName =
            Contract.DisableServerAssistedSearchValueName;

        /// <summary>
        /// User-hive Outlook search key carrying DisableServerAssistedSearch (D22; the tuning
        /// Search group writes here). Aimed at the Office major this machine actually has - it
        /// was a hardcoded 16.0, which read a non-existent hive on Outlook 2013 or a future 17.0
        /// and reported the resulting nothing as "server-assisted", the default.
        /// </summary>
        public static readonly string OutlookSearchUserKeyPath =
            BuildOutlookSearchUserKeyPath(OutlookProfileRegistry.OfficeVersion);

        /// <summary>Policy-hive Outlook search key - authoritative over the user hive when its value exists (ADMX-managed).</summary>
        public static readonly string OutlookSearchPolicyKeyPath =
            BuildOutlookSearchPolicyKeyPath(OutlookProfileRegistry.OfficeVersion);

        /// <summary>
        /// The user-hive search key for an arbitrary Office major (pure, so the 15.0 and 17.0
        /// shapes are assertable on a machine that has neither).
        /// <para>
        /// Built by the SHARED <c>OfficeVersions.OutlookSearchKeyPath</c>, which the add-in's
        /// <c>OutlookTuningService</c> also builds its Search key from. Sharing that file used to
        /// settle only the VERSION in this path: the add-in concatenated the rest of it by hand,
        /// so one address had two spellings across a boundary no compiler crosses, and a typo in
        /// either aimed at a key Outlook never touches with both halves still compiling.
        /// </para>
        /// </summary>
        public static string BuildOutlookSearchUserKeyPath(string officeVersion)
        {
            if (officeVersion == null)
            {
                throw new ArgumentNullException(nameof(officeVersion));
            }

            return OfficeVersions.OutlookSearchKeyPath(officeVersion);
        }

        /// <summary>
        /// The policy-hive mirror of the key above, from the same shared file: a policy value
        /// outranks the user hive, so a report that ignored it would call a policy-disabled
        /// machine tuned.
        /// <para>
        /// READ HERE AND NOWHERE ELSE, and that asymmetry is deliberate rather than an omission:
        /// the add-in sets a user preference, not search policy, so it has no reason to write
        /// this key. It is NOT true that the add-in leaves the Policies hive alone in general -
        /// its tuning service writes five values under
        /// <c>...\Policies\...\Outlook\Cached Mode</c> (D25's sync-slider settings), which is why
        /// the policy ROOT is built by <c>OfficeVersions.PolicyOutlookKeyPath</c> on both sides.
        /// </para>
        /// </summary>
        public static string BuildOutlookSearchPolicyKeyPath(string officeVersion)
        {
            if (officeVersion == null)
            {
                throw new ArgumentNullException(nameof(officeVersion));
            }

            return OfficeVersions.PolicyOutlookSearchKeyPath(officeVersion);
        }

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

            // Every value name below comes from the shared contract file the add-in's
            // OutlookTuningService writes with, so a rename cannot land on one side only.
            if (AsBool(readValue(Contract.TuningInitializedValueName)) != true)
            {
                return new TuningHealthView { Managed = false };
            }

            string? conflicts = readValue(Contract.TuningPolicyConflictsValueName) as string;
            string? needsAdministrator = readValue(Contract.TuningNeedsAdministratorValueName) as string;
            return new TuningHealthView
            {
                Managed = true,
                Enabled = AsBool(readValue(Contract.TuningEnabledValueName)),
                SearchEnabled = AsBool(readValue(Contract.TuningSearchEnabledValueName)),
                CachingEnabled = AsBool(readValue(Contract.TuningCachingEnabledValueName)),
                OstEnabled = AsBool(readValue(Contract.TuningOstEnabledValueName)),
                RestartNeeded = AsBool(readValue(Contract.TuningRestartNeededValueName)),
                PolicyConflicts = string.IsNullOrWhiteSpace(conflicts) ? null : conflicts,
                NeedsAdministrator = string.IsNullOrWhiteSpace(needsAdministrator) ? null : needsAdministrator,
                LastReconcileUtc = readValue(Contract.TuningLastReconcileUtcValueName) as string,
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
                TryReadCurrentUserDword(OutlookSearchPolicyKeyPath, Contract.DisableServerAssistedSearchValueName),
                TryReadCurrentUserDword(OutlookSearchUserKeyPath, Contract.DisableServerAssistedSearchValueName));
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

        // ===== Office major version =====
        // Not the same fact as TryGetOutlookVersion above, and the two are worth telling apart.
        // That one is the installed BUILD (OUTLOOK.EXE's file version, e.g. 16.0.14332.20255);
        // this one is the registry HIVE MAJOR - which Software\Microsoft\Office\<major>\ subtree
        // every registry-backed answer in this server is read out of. The server used to assume
        // 16.0 for all of them, so a 15.0 or 17.0 machine got empty accounts, empty signature
        // defaults and a default-looking search setting, with nothing anywhere saying why.

        /// <summary>
        /// The Office major this machine has, or null when NONE of the supported majors is
        /// registered. Detected once per process (see
        /// <see cref="OutlookProfileRegistry.DetectedOfficeVersion"/>); never throws.
        /// </summary>
        public static string? DetectedOfficeVersion()
        {
            try
            {
                return OutlookProfileRegistry.DetectedOfficeVersion;
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                return null;
            }
        }

        /// <summary>
        /// The same detection driven by a caller-supplied "does this HKCU key exist?" predicate
        /// (pure, T1-tested exactly the way <see cref="ReadTuningState"/> is). This machine has a
        /// single Office version installed, so the 15.0, 17.0 and nothing-found branches have no
        /// other way of being exercised without a second Outlook.
        /// </summary>
        public static string? DetectOfficeVersion(Func<string, bool> outlookKeyExists)
        {
            return OutlookProfileRegistry.DetectOfficeVersion(outlookKeyExists);
        }

        /// <summary>
        /// The <c>problems</c> line for a machine where none of the supported Office majors is
        /// registered. Assembled from the supported list and the hive actually being read, so it
        /// cannot go stale when a version is added, and it says the one thing this tool exists to
        /// say: the empty answers are not a broken profile, they are the wrong hive.
        /// </summary>
        public static readonly string NoOfficeVersionProblem =
            "No supported Office version is registered on this machine: none of "
            + string.Join(", ", OutlookProfileRegistry.SupportedOfficeVersions)
            + @" has an HKCU\Software\Microsoft\Office\<version>\Outlook key. Everything this server reads from "
            + "the registry (accounts, per-account signature defaults, the Outlook search settings reported here) "
            + "is therefore being read from " + OutlookProfileRegistry.OutlookRootKeyPath + " as a fallback and can "
            + "come back EMPTY even though Outlook itself is working. If this Outlook is newer than the versions "
            + "listed above, this product needs that version added; otherwise Outlook has not written its profile "
            + "key yet, which a normal Outlook start fixes.";

        // ===== MCP registration (Phase 8) =====

        /// <summary>
        /// Registry path of the add-in's registration state - the same constant the add-in's
        /// <c>McpRegistrationService</c> writes with, <see cref="AddInServerContract.McpKeyPath"/>,
        /// not a mirror of it.
        /// </summary>
        public const string McpRegistrationKeyPath = Contract.McpKeyPath;

        // The registration.status vocabulary. These five strings are also PUBLISHED - README.md
        // lists them for agents and McpServer/README.md explains each one - and Markdown cannot
        // read a C# constant, so Tools/Checks/check-pinned-constants.ps1 (#6) compares the two
        // rather than leaving it to memory. Note this set is the SERVER's own verdict about
        // ~/.claude.json; the add-in's own status codes (McpRegistrationService.Status*, surfaced
        // here verbatim as AddInStatus) are a different vocabulary for a different field.

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
        ///
        /// The registered command is EXPANDED first. Claude Code resolves <c>${VAR}</c> in
        /// <c>command</c> when it reads its configuration, and the add-in registers the
        /// portable <c>${LOCALAPPDATA}/…</c> spelling on purpose, so comparing the raw text
        /// would report a perfectly correct registration as drift.
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

            string resolved = McpConfigEditor.ExpandEnvironmentReferences(registeredCommand);
            if (string.IsNullOrWhiteSpace(resolved))
            {
                // Every reference in it expanded to nothing: it names no file at all, which
                // is a broken registration, not this one.
                return RegistrationDrifted;
            }

            return McpConfigEditor.SamePath(resolved, runningFrom!) ? RegistrationOk : RegistrationDrifted;
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
                        // Value names from the shared contract - the add-in's Persist() writes
                        // these very constants.
                        view.AddInStatus = key.GetValue(Contract.McpStatusValueName) as string;
                        view.AddInLastReconcileUtc = key.GetValue(Contract.McpLastReconcileUtcValueName) as string;
                        view.AddInResolvedServerPath = key.GetValue(Contract.McpResolvedServerPathValueName) as string;
                        object? healed = key.GetValue(Contract.McpHealedValueName);
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
                // The one place the path is built, for the add-in and for this report alike.
                string path = McpConfigEditor.UserConfigPath();
                bool exists = File.Exists(path);
                string text = "";
                if (exists)
                {
                    // FileShare.ReadWrite: the Claude Code CLI owns this file and may hold it open.
                    using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    using (var reader = new StreamReader(stream))
                    {
                        text = reader.ReadToEnd();
                    }
                }

                return ReadRegisteredCommand(exists, text, out readable);
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                readable = false;
                return null;
            }
        }

        /// <summary>
        /// The registration a configuration file holds, judged by the rule the add-in repairs it by,
        /// <see cref="McpConfigEditor.IsReadableConfig"/> (pure, T1-pinned). An ABSENT file is readable
        /// and holds nothing. A file that is there but reads back empty is UNREADABLE, not
        /// unregistered: the Claude Code CLI truncates the file before it flushes, and a read taken in
        /// between is a failed read. Until stage S3 of the code-sharing refactor this report called
        /// such a file "absent" while the add-in called it unreadable.
        /// </summary>
        public static string? ReadRegisteredCommand(bool fileExists, string? text, out bool readable)
        {
            readable = McpConfigEditor.IsReadableConfig(fileExists, text);
            if (!readable || !fileExists || text == null)
            {
                return null;
            }

            return McpConfigEditor.TryReadServerCommand(text, out string command) ? command : null;
        }

        /// <summary>
        /// Pulls mcpServers -> outlookai -> command out of raw JSON text (pure, T1-pinned), through the
        /// add-in's own reader, <see cref="McpConfigEditor"/>: one scanner for the file the add-in
        /// writes and this report reads, where until stage S3 of the code-sharing refactor there were
        /// two (Q154). Read-only by construction: nothing here can modify the file.
        /// <paramref name="readable"/> is false when the text is not one well-formed JSON object - the
        /// add-in will not touch such a file either - and empty text is readable with no entry.
        /// </summary>
        public static string? ExtractRegisteredCommand(string? json, out bool readable)
        {
            readable = true;
            if (json == null || json.Trim().Length == 0)
            {
                return null;
            }

            if (!McpConfigEditor.TryValidateJsonObject(json, out _))
            {
                readable = false;
                return null;
            }

            return McpConfigEditor.TryReadServerCommand(json, out string command) ? command : null;
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
