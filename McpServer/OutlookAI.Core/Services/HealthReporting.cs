using System;
using System.Diagnostics;
using System.Globalization;
using System.IO;

using Microsoft.Win32;

// The one definition of everything the add-in and this server have to agree about: the HKCU keys
// the add-in writes and this file reads, and the Claude Code configuration both of them look at.
// Services\AddInServerContract.cs is LINKED into this project (see OutlookAI.Core.csproj), so
// these are not copies of the add-in's constants - they are the add-in's constants.
using Contract = global::OutlookAI.Services.AddInServerContract;
// And the Office majors, plus the hive paths built from whichever one this machine has.
// Services\OfficeVersions.cs is LINKED here too, so the Outlook Search key this file reads is
// built by the same expression the add-in writes it with - see OutlookSearchUserKeyPath below.
using OfficeVersions = global::OutlookAI.Services.OfficeVersions;

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
        /// same one: <c>Services\AddInServerContract.cs</c> is compiled into this assembly as well
        /// as into the add-in, so there is nothing here to keep in step. This member stays public
        /// only because the contract type is internal (CS0436 - see that file's header) and
        /// callers outside Core name this one.
        /// </summary>
        public const string TuningKeyPath = Contract.TuningKeyPath;

        /// <summary>
        /// The value name this file reads out of both Search keys below - NOT a copy of the
        /// add-in's: <c>Services\AddInServerContract.cs</c> is compiled into this assembly as
        /// well as into the add-in, so the name the tuning service writes and the name read back
        /// here are one constant. Public only because the contract type is internal (CS0436 - see
        /// that file's header) and the T2 live test that flips the value names this one.
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
            return new TuningHealthView
            {
                Managed = true,
                Enabled = AsBool(readValue(Contract.TuningEnabledValueName)),
                SearchEnabled = AsBool(readValue(Contract.TuningSearchEnabledValueName)),
                CachingEnabled = AsBool(readValue(Contract.TuningCachingEnabledValueName)),
                OstEnabled = AsBool(readValue(Contract.TuningOstEnabledValueName)),
                RestartNeeded = AsBool(readValue(Contract.TuningRestartNeededValueName)),
                PolicyConflicts = string.IsNullOrWhiteSpace(conflicts) ? null : conflicts,
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
        /// <c>McpRegistrationService</c> writes with, from the linked contract file, not a mirror
        /// of it. Public for the same reason as <see cref="TuningKeyPath"/>.
        /// </summary>
        public const string McpRegistrationKeyPath = Contract.McpKeyPath;

        // The registration.status vocabulary. These five strings are also PUBLISHED - README.md
        // lists them for agents and McpServer/README.md explains each one - and Markdown cannot
        // read a C# constant, so .github/scripts/check-pinned-constants.ps1 (#6) compares the two
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

            string resolved = ExpandEnvironmentReferences(registeredCommand);
            if (string.IsNullOrWhiteSpace(resolved))
            {
                // Every reference in it expanded to nothing: it names no file at all, which
                // is a broken registration, not this one.
                return RegistrationDrifted;
            }

            return SamePath(resolved, runningFrom!) ? RegistrationOk : RegistrationDrifted;
        }

        /// <summary>
        /// Expands the <c>${VAR}</c> and <c>${VAR:-default}</c> forms Claude Code documents
        /// for <c>command</c>, <c>args</c> and <c>env</c> (pure when given a
        /// <paramref name="lookup"/>; T1-pinned). Anything else - a bare <c>$</c>, an
        /// unterminated or empty <c>${}</c> - is literal text and is copied through, which is
        /// the only safe reading for a value about to be compared against a real file.
        ///
        /// The add-in's <c>McpConfigEditor.ExpandEnvironmentReferences</c> is the same rule on
        /// the writing side. This one stays a second IMPLEMENTATION rather than a shared file -
        /// unlike the constants in <c>AddInServerContract</c>, which Core now links - because the
        /// two signatures differ where it matters: this one is nullable-annotated and defaults its
        /// lookup to the process environment, the add-in's takes a required non-nullable lookup,
        /// and reconciling them would change a public surface on both sides for no behavioural
        /// gain. They cannot drift apart unnoticed because
        /// <c>EnvironmentExpansionParityTests</c> runs ONE shared corpus through BOTH
        /// implementations and asserts they agree character for character. That claim used to
        /// be made here and was false: each side had its own suite with its own fixture
        /// literals and nothing ever fed one input to both, even though the test project
        /// already compiles the add-in file.
        /// </summary>
        public static string ExpandEnvironmentReferences(string? value, Func<string, string?>? lookup = null)
        {
            // Bound to a non-nullable local once: net48's reference assemblies are not
            // annotated, so the null-state of string.IsNullOrEmpty's argument does not flow
            // there the way it does on net10, and Core compiles warning-free for both.
            string text = value ?? "";
            if (text.Length == 0 || text.IndexOf("${", StringComparison.Ordinal) < 0)
            {
                return text;
            }

            lookup ??= static name =>
            {
                try
                {
                    return Environment.GetEnvironmentVariable(name);
                }
                catch (Exception ex) when (!(ex is OutOfMemoryException))
                {
                    return null;
                }
            };

            var sb = new System.Text.StringBuilder(text.Length);
            int i = 0;
            while (i < text.Length)
            {
                if (text[i] != '$' || i + 1 >= text.Length || text[i + 1] != '{')
                {
                    sb.Append(text[i]);
                    i++;
                    continue;
                }

                int close = text.IndexOf('}', i + 2);
                if (close < 0)
                {
                    sb.Append(text, i, text.Length - i);
                    break;
                }

                string inner = text.Substring(i + 2, close - i - 2);
                string name = inner;
                string fallback = "";
                int marker = inner.IndexOf(":-", StringComparison.Ordinal);
                if (marker >= 0)
                {
                    name = inner.Substring(0, marker);
                    fallback = inner.Substring(marker + 2);
                }

                if (name.Length == 0)
                {
                    sb.Append(text, i, close - i + 1);
                    i = close + 1;
                    continue;
                }

                string resolved = "";
                try
                {
                    resolved = lookup(name) ?? "";
                }
                catch (Exception ex) when (!(ex is OutOfMemoryException))
                {
                    resolved = "";
                }

                sb.Append(resolved.Length > 0 ? resolved : fallback);
                i = close + 1;
            }

            return sb.ToString();
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
                // Same file name, resolved the same way, as the add-in's
                // McpRegistrationService.ClaudeConfigPath - from the shared contract.
                string path = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                    Contract.ClaudeConfigFileName);

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

            // mcpServers / outlookai / command all come from the shared contract: the add-in
            // WRITES these three names and this scanner READS them, so renaming our entry (which
            // would deregister the server on every machine) can only ever be a one-place edit.
            int servers = FindMemberValue(json, i, Contract.ServersProperty);
            if (servers < 0 || json[servers] != '{')
            {
                return null;
            }

            int entry = FindMemberValue(json, servers, Contract.ServerName);
            if (entry < 0 || json[entry] != '{')
            {
                return null;
            }

            int command = FindMemberValue(json, entry, Contract.CommandProperty);
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
