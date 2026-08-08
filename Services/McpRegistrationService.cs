using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;
using Microsoft.Win32;

namespace OutlookAI.Services
{
    /// <summary>
    /// Keeps Claude Code's user-global MCP registration pointing at the OutlookAI mail
    /// server that is actually installed (v3 plan D6 v2 / R11 / Phase 8).
    ///
    /// This is the same on-every-Outlook-start reconcile shape as
    /// <see cref="OutlookTuningService"/>, and for the same reasons: it heals drift instead
    /// of relying on a one-shot install-time write. The drift it exists to fix is real —
    /// during development the entry points at a build-output path, and an installed copy
    /// must take that over without the user editing anything.
    ///
    /// The file it edits (<c>~/.claude.json</c>) belongs to the Claude Code CLI, is large,
    /// and is rewritten by that CLI, so the discipline here is deliberately defensive:
    ///  - a file that does not PARSE is never written to (a truncating rewrite would cost the
    ///    user every project entry in it);
    ///  - only the <c>mcpServers</c> value is re-rendered; every other byte of the file is
    ///    spliced through unchanged, so unrelated settings cannot be reformatted or reordered;
    ///  - the result is re-parsed and checked before it replaces anything;
    ///  - the replacement is atomic and keeps a backup;
    ///  - a correct entry is a no-op — nothing is written, and no backup is churned.
    ///
    /// Everything is per-user (HKCU + the user profile); no elevation, no COM, no Outlook
    /// object model — so this can run off the UI thread without touching the add-in's COM
    /// ownership rules.
    /// </summary>
    internal static class McpRegistrationService
    {
        internal const string McpKeyPath = @"Software\OutlookAI\Mcp";
        internal const string InstallDirValueName = "InstallDir";
        private const string AppKeyPath = @"Software\OutlookAI";

        /// <summary>Server name under <c>mcpServers</c>; matches what Phase 2 registered.</summary>
        internal const string ServerName = "outlookai";

        internal const string RelativeServerPath = @"McpServer\OutlookAI.McpServer.exe";

        // Status codes. Also written to HKCU so the MCP server can report them in
        // outlook_health without having to guess what the add-in did.
        internal const string StatusOk = "ok";
        internal const string StatusHealed = "healed";
        internal const string StatusNoClaude = "claude_code_not_installed";
        internal const string StatusNoServer = "server_not_installed";
        internal const string StatusNoRuntime = "dotnet_runtime_missing";
        internal const string StatusParseFailed = "config_unreadable";
        internal const string StatusError = "error";

        internal const string DotnetRuntimeDownloadUrl = "https://dotnet.microsoft.com/download/dotnet/10.0";

        private static readonly object _gate = new object();

        internal sealed class RegistrationSnapshot
        {
            /// <summary>One of the Status* constants.</summary>
            public string Status { get; internal set; }

            /// <summary>Human-readable detail; never null.</summary>
            public string Detail { get; internal set; }

            /// <summary>Path this reconcile wants registered, or null when none was resolved.</summary>
            public string ResolvedServerPath { get; internal set; }

            /// <summary>Path currently registered in ~/.claude.json, or null.</summary>
            public string RegisteredCommand { get; internal set; }

            /// <summary>True when this run rewrote the registration.</summary>
            public bool Healed { get; internal set; }

            public string LastReconcileUtc { get; internal set; }
        }

        // ===== Public surface =====

        /// <summary>
        /// Reconciles the registration. Never throws: every failure becomes a status the
        /// settings dialog and outlook_health can show.
        /// </summary>
        internal static RegistrationSnapshot Reconcile()
        {
            lock (_gate)
            {
                RegistrationSnapshot snap;
                try
                {
                    snap = ReconcileCore();
                }
                catch (Exception ex)
                {
                    snap = new RegistrationSnapshot
                    {
                        Status = StatusError,
                        Detail = ex.Message,
                        Healed = false,
                    };
                }

                snap.LastReconcileUtc = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
                try { Persist(snap); }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine("MCP registration persist: " + ex.Message); }
                return snap;
            }
        }

        /// <summary>Last recorded state, for the settings dialog. Never throws.</summary>
        internal static RegistrationSnapshot GetSnapshot()
        {
            try
            {
                return new RegistrationSnapshot
                {
                    Status = ReadString("Status") ?? StatusError,
                    Detail = ReadString("Detail") ?? "",
                    ResolvedServerPath = ReadString("ResolvedServerPath"),
                    RegisteredCommand = ReadString("Command"),
                    Healed = ReadDword("Healed") == 1,
                    LastReconcileUtc = ReadString("LastReconcileUtc"),
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP registration snapshot: " + ex.Message);
                return new RegistrationSnapshot { Status = StatusError, Detail = ex.Message };
            }
        }

        // ===== Reconcile =====

        private static RegistrationSnapshot ReconcileCore()
        {
            string configPath = ClaudeConfigPath();
            string serverPath = ResolveInstalledServerPath();

            var snap = new RegistrationSnapshot
            {
                ResolvedServerPath = serverPath,
                Detail = "",
                Healed = false,
            };

            // (1) No Claude Code on this machine: nothing to register into. Not an error —
            // the add-in is perfectly usable on its own.
            if (!IsClaudeCodeInstalled(configPath))
            {
                snap.Status = StatusNoClaude;
                snap.Detail = "Claude Code was not found on this machine, so there is nothing to register with.";
                return snap;
            }

            // (2) Add-in installed without the mail server beside it (an older install layout,
            // or a developer add-in build). Leave whatever is registered alone — in the
            // developer case that entry is deliberately pointing at a build output.
            if (serverPath == null)
            {
                snap.Status = StatusNoServer;
                snap.Detail = "The OutlookAI mail server was not found next to the installed add-in, so the existing registration was left unchanged.";
                snap.RegisteredCommand = TryReadRegisteredCommand(configPath);
                return snap;
            }

            snap.RegisteredCommand = TryReadRegisteredCommand(configPath);

            // (3) The server needs the .NET 10 runtime. Registering a server that cannot start
            // would surface as a failed MCP server in every Claude session, so report instead
            // and heal on a later start once the runtime is there.
            if (!IsDotnetRuntime10Installed())
            {
                snap.Status = StatusNoRuntime;
                snap.Detail = "The .NET 10 runtime the mail server needs is not installed, so registration was skipped. Install it from " + DotnetRuntimeDownloadUrl + " and restart Outlook.";
                return snap;
            }

            string raw;
            try
            {
                raw = File.Exists(configPath) ? File.ReadAllText(configPath) : "";
            }
            catch (Exception ex)
            {
                snap.Status = StatusError;
                snap.Detail = "Could not read " + configPath + ": " + ex.Message;
                return snap;
            }

            Dictionary<string, object> root = null;
            if (raw.Trim().Length > 0)
            {
                root = TryParseObject(raw);
                if (root == null)
                {
                    // Never rewrite a file we could not understand — that is how config gets lost.
                    snap.Status = StatusParseFailed;
                    snap.Detail = "Claude Code's configuration file could not be read as JSON, so it was left untouched.";
                    return snap;
                }
            }

            // (4) Already correct — do nothing at all. No write, no backup churn.
            if (IsSamePath(snap.RegisteredCommand, serverPath) && IsStdioEntry(root))
            {
                snap.Status = StatusOk;
                snap.Detail = "Registered and pointing at the installed mail server.";
                return snap;
            }

            string updated;
            string spliceError;
            if (!TryBuildUpdatedConfig(raw, root, serverPath, out updated, out spliceError))
            {
                snap.Status = StatusError;
                snap.Detail = "Could not update the registration safely: " + spliceError + ". The file was left untouched.";
                return snap;
            }

            string writeError;
            if (!TryWriteAtomically(configPath, updated, out writeError))
            {
                snap.Status = StatusError;
                snap.Detail = "Could not write the registration: " + writeError;
                return snap;
            }

            snap.Status = StatusHealed;
            snap.Healed = true;
            snap.RegisteredCommand = serverPath;
            snap.Detail = "Registration was pointing elsewhere and has been repaired.";
            return snap;
        }

        // ===== Paths and probes =====

        internal static string ClaudeConfigPath()
        {
            return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".claude.json");
        }

        /// <summary>
        /// Claude Code is present if its CLI is where the add-in already expects it
        /// (see <see cref="ClaudeService"/>) or if it has written its config file.
        /// </summary>
        private static bool IsClaudeCodeInstalled(string configPath)
        {
            try
            {
                if (File.Exists(configPath))
                    return true;
                string cli = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                    @".local\bin\claude.exe");
                return File.Exists(cli);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// The installed server executable, or null when there isn't one. Prefers the install
        /// directory the installer recorded; falls back to walking up from this assembly (an
        /// installed add-in lives under {app}\Application Files\OutlookAI_x_y_z_w\, so {app}
        /// is two levels up, but that layout is not worth hard-coding).
        /// </summary>
        internal static string ResolveInstalledServerPath()
        {
            try
            {
                string installDir = ReadAppString(InstallDirValueName);
                if (!string.IsNullOrEmpty(installDir))
                {
                    string candidate = Path.Combine(installDir, RelativeServerPath);
                    if (File.Exists(candidate))
                        return Path.GetFullPath(candidate);
                }

                string dir = Path.GetDirectoryName(typeof(McpRegistrationService).Assembly.Location);
                for (int i = 0; i < 4 && !string.IsNullOrEmpty(dir); i++)
                {
                    string candidate = Path.Combine(dir, RelativeServerPath);
                    if (File.Exists(candidate))
                        return Path.GetFullPath(candidate);
                    dir = Path.GetDirectoryName(dir);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP server path resolve: " + ex.Message);
            }

            return null;
        }

        /// <summary>
        /// True when some Microsoft.NETCore.App 10.x shared framework is present. Exactly
        /// 10.x: the server's default roll-forward (Minor) accepts a newer 10.x but not 11.x.
        /// Probes the filesystem because the sharedfx registry key is absent on machines that
        /// do have the runtime.
        /// </summary>
        internal static bool IsDotnetRuntime10Installed()
        {
            try
            {
                string root = null;
                using (var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedhost"))
                {
                    if (key != null)
                    {
                        var path = key.GetValue("Path") as string;
                        if (!string.IsNullOrEmpty(path))
                            root = Path.Combine(path, @"shared\Microsoft.NETCore.App");
                    }
                }

                if (string.IsNullOrEmpty(root))
                {
                    root = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles),
                        @"dotnet\shared\Microsoft.NETCore.App");
                }

                if (!Directory.Exists(root))
                    return false;

                foreach (var d in Directory.GetDirectories(root))
                {
                    string name = Path.GetFileName(d);
                    if (name != null && name.StartsWith("10.", StringComparison.Ordinal))
                        return true;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("dotnet runtime probe: " + ex.Message);
            }

            return false;
        }

        // ===== JSON: read, splice, verify =====

        private static Dictionary<string, object> TryParseObject(string json)
        {
            try
            {
                var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
                return serializer.Deserialize<Dictionary<string, object>>(json);
            }
            catch
            {
                return null;
            }
        }

        private static string TryReadRegisteredCommand(string configPath)
        {
            try
            {
                if (!File.Exists(configPath))
                    return null;
                var root = TryParseObject(File.ReadAllText(configPath));
                return ReadRegisteredCommand(root);
            }
            catch
            {
                return null;
            }
        }

        private static string ReadRegisteredCommand(Dictionary<string, object> root)
        {
            var entry = ReadServerEntry(root);
            if (entry == null)
                return null;
            object command;
            return entry.TryGetValue("command", out command) ? command as string : null;
        }

        private static Dictionary<string, object> ReadServerEntry(Dictionary<string, object> root)
        {
            if (root == null)
                return null;
            object servers;
            if (!root.TryGetValue("mcpServers", out servers))
                return null;
            var map = servers as Dictionary<string, object>;
            if (map == null)
                return null;
            object entry;
            if (!map.TryGetValue(ServerName, out entry))
                return null;
            return entry as Dictionary<string, object>;
        }

        private static bool IsStdioEntry(Dictionary<string, object> root)
        {
            var entry = ReadServerEntry(root);
            if (entry == null)
                return false;
            object type;
            if (!entry.TryGetValue("type", out type))
                return false;
            return string.Equals(type as string, "stdio", StringComparison.OrdinalIgnoreCase);
        }

        internal static bool IsSamePath(string a, string b)
        {
            if (string.IsNullOrEmpty(a) || string.IsNullOrEmpty(b))
                return false;
            try
            {
                return string.Equals(
                    Path.GetFullPath(a).TrimEnd('\\'),
                    Path.GetFullPath(b).TrimEnd('\\'),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
            }
        }

        /// <summary>
        /// Produces the new file content by replacing ONLY the <c>mcpServers</c> value (or
        /// inserting the whole property when absent). Everything outside that span is copied
        /// through byte for byte, so no unrelated setting is reformatted. The result is
        /// re-parsed and checked before it is returned.
        /// </summary>
        internal static bool TryBuildUpdatedConfig(
            string raw,
            Dictionary<string, object> root,
            string serverPath,
            out string updated,
            out string error)
        {
            updated = null;
            error = null;

            var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

            // Preserve any other servers, and any extra keys on our own entry, verbatim.
            var servers = new Dictionary<string, object>(StringComparer.Ordinal);
            object existingServers;
            if (root != null && root.TryGetValue("mcpServers", out existingServers))
            {
                var map = existingServers as Dictionary<string, object>;
                if (map != null)
                {
                    foreach (var kv in map)
                        servers[kv.Key] = kv.Value;
                }
                else if (existingServers != null)
                {
                    error = "the mcpServers setting is not an object";
                    return false;
                }
            }

            Dictionary<string, object> entry = null;
            object ours;
            if (servers.TryGetValue(ServerName, out ours))
                entry = ours as Dictionary<string, object>;
            if (entry == null)
                entry = new Dictionary<string, object>(StringComparer.Ordinal);
            else
                entry = new Dictionary<string, object>(entry, StringComparer.Ordinal);

            entry["type"] = "stdio";
            entry["command"] = serverPath;
            if (!entry.ContainsKey("args"))
                entry["args"] = new object[0];
            if (!entry.ContainsKey("env"))
                entry["env"] = new Dictionary<string, object>(StringComparer.Ordinal);

            servers[ServerName] = entry;

            string serversJson;
            try
            {
                serversJson = serializer.Serialize(servers);
            }
            catch (Exception ex)
            {
                error = "could not render the mcpServers section (" + ex.Message + ")";
                return false;
            }

            int expectedTopLevelKeys;
            if (raw.Trim().Length == 0)
            {
                updated = "{\"mcpServers\":" + serversJson + "}";
                expectedTopLevelKeys = 1;
            }
            else
            {
                int valueStart, valueEnd;
                if (TryFindTopLevelValueSpan(raw, "mcpServers", out valueStart, out valueEnd))
                {
                    updated = raw.Substring(0, valueStart) + serversJson + raw.Substring(valueEnd);
                    expectedTopLevelKeys = root.Count;
                }
                else
                {
                    int brace = raw.IndexOf('{');
                    if (brace < 0)
                    {
                        error = "the configuration file is not a JSON object";
                        return false;
                    }
                    string separator = HasAnyContentAfterBrace(raw, brace) ? "," : "";
                    updated = raw.Substring(0, brace + 1)
                        + "\"mcpServers\":" + serversJson + separator
                        + raw.Substring(brace + 1);
                    expectedTopLevelKeys = root.Count + 1;
                }
            }

            // Verify before anyone gets to write it.
            var reparsed = TryParseObject(updated);
            if (reparsed == null)
            {
                error = "the updated configuration did not parse back";
                updated = null;
                return false;
            }
            if (reparsed.Count != expectedTopLevelKeys)
            {
                error = "the updated configuration changed the number of top-level settings";
                updated = null;
                return false;
            }
            if (!IsSamePath(ReadRegisteredCommand(reparsed), serverPath))
            {
                error = "the updated configuration does not name the installed server";
                updated = null;
                return false;
            }

            return true;
        }

        private static bool HasAnyContentAfterBrace(string raw, int brace)
        {
            for (int i = brace + 1; i < raw.Length; i++)
            {
                char c = raw[i];
                if (char.IsWhiteSpace(c))
                    continue;
                return c != '}';
            }
            return false;
        }

        /// <summary>
        /// Finds the character span of a top-level property's VALUE in raw JSON text. Depth
        /// aware and string/escape aware, so a key of the same name nested inside a project
        /// entry is not mistaken for the top-level one.
        /// </summary>
        internal static bool TryFindTopLevelValueSpan(string json, string key, out int valueStart, out int valueEnd)
        {
            valueStart = -1;
            valueEnd = -1;

            int depth = 0;
            int i = 0;
            while (i < json.Length)
            {
                char c = json[i];

                if (c == '"')
                {
                    int stringStart = i;
                    int stringEnd = SkipString(json, i);
                    if (stringEnd < 0)
                        return false;

                    if (depth == 1)
                    {
                        int afterKey = SkipWhitespace(json, stringEnd);
                        if (afterKey < json.Length && json[afterKey] == ':')
                        {
                            string name = json.Substring(stringStart + 1, stringEnd - stringStart - 2);
                            if (Unescape(name) == key)
                            {
                                int start = SkipWhitespace(json, afterKey + 1);
                                int end = SkipValue(json, start);
                                if (end < 0)
                                    return false;
                                valueStart = start;
                                valueEnd = end;
                                return true;
                            }
                            // Not our key: jump past its value so a nested object cannot be
                            // scanned at the wrong depth.
                            int skipTo = SkipValue(json, SkipWhitespace(json, afterKey + 1));
                            if (skipTo < 0)
                                return false;
                            i = skipTo;
                            continue;
                        }
                    }

                    i = stringEnd;
                    continue;
                }

                if (c == '{' || c == '[')
                    depth++;
                else if (c == '}' || c == ']')
                    depth--;

                i++;
            }

            return false;
        }

        private static int SkipWhitespace(string json, int i)
        {
            while (i < json.Length && char.IsWhiteSpace(json[i]))
                i++;
            return i;
        }

        /// <summary>Index just past the closing quote, or -1 when unterminated.</summary>
        private static int SkipString(string json, int i)
        {
            i++; // opening quote
            while (i < json.Length)
            {
                char c = json[i];
                if (c == '\\')
                {
                    i += 2;
                    continue;
                }
                if (c == '"')
                    return i + 1;
                i++;
            }
            return -1;
        }

        /// <summary>Index just past the end of the value starting at i, or -1.</summary>
        private static int SkipValue(string json, int i)
        {
            if (i >= json.Length)
                return -1;

            char c = json[i];
            if (c == '"')
                return SkipString(json, i);

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
                            return -1;
                        i = end;
                        continue;
                    }
                    if (d == '{' || d == '[')
                        depth++;
                    else if (d == '}' || d == ']')
                    {
                        depth--;
                        if (depth == 0)
                            return i + 1;
                    }
                    i++;
                }
                return -1;
            }

            // Number, true, false, null: runs until a structural character.
            while (i < json.Length)
            {
                char d = json[i];
                if (d == ',' || d == '}' || d == ']' || char.IsWhiteSpace(d))
                    return i;
                i++;
            }
            return i;
        }

        private static string Unescape(string s)
        {
            if (s.IndexOf('\\') < 0)
                return s;
            var sb = new StringBuilder(s.Length);
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
                        if (i + 4 < s.Length)
                        {
                            int code;
                            if (int.TryParse(s.Substring(i + 1, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out code))
                            {
                                sb.Append((char)code);
                                i += 4;
                                break;
                            }
                        }
                        sb.Append(c);
                        break;
                    default: sb.Append(c); break;
                }
            }
            return sb.ToString();
        }

        /// <summary>
        /// Writes via a sibling temp file and <see cref="File.Replace(string,string,string)"/>,
        /// which swaps the file in one operation and leaves the previous content as a backup.
        /// A half-written config is therefore not reachable even if the machine dies mid-write.
        /// </summary>
        private static bool TryWriteAtomically(string configPath, string content, out string error)
        {
            error = null;
            string dir = Path.GetDirectoryName(configPath);
            string temp = Path.Combine(dir, ".claude.json.outlookai-new");
            string backup = Path.Combine(dir, ".claude.json.outlookai-backup");

            try
            {
                // No BOM: the CLI reads this file as plain UTF-8.
                File.WriteAllText(temp, content, new UTF8Encoding(false));

                if (File.Exists(configPath))
                    File.Replace(temp, configPath, backup, true);
                else
                    File.Move(temp, configPath);

                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                try { if (File.Exists(temp)) File.Delete(temp); }
                catch { }
                return false;
            }
        }

        // ===== Registry bookkeeping =====

        private static void Persist(RegistrationSnapshot snap)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(McpKeyPath))
            {
                if (key == null)
                    return;
                key.SetValue("Status", snap.Status ?? "", RegistryValueKind.String);
                key.SetValue("Detail", snap.Detail ?? "", RegistryValueKind.String);
                key.SetValue("Command", snap.RegisteredCommand ?? "", RegistryValueKind.String);
                key.SetValue("ResolvedServerPath", snap.ResolvedServerPath ?? "", RegistryValueKind.String);
                key.SetValue("Healed", snap.Healed ? 1 : 0, RegistryValueKind.DWord);
                key.SetValue("LastReconcileUtc", snap.LastReconcileUtc ?? "", RegistryValueKind.String);
            }
        }

        private static string ReadString(string valueName)
        {
            using (var key = Registry.CurrentUser.OpenSubKey(McpKeyPath))
            {
                if (key == null)
                    return null;
                var s = key.GetValue(valueName) as string;
                return string.IsNullOrEmpty(s) ? null : s;
            }
        }

        private static string ReadAppString(string valueName)
        {
            using (var key = Registry.CurrentUser.OpenSubKey(AppKeyPath))
            {
                if (key == null)
                    return null;
                return key.GetValue(valueName) as string;
            }
        }

        private static int? ReadDword(string valueName)
        {
            using (var key = Registry.CurrentUser.OpenSubKey(McpKeyPath))
            {
                if (key == null)
                    return null;
                var v = key.GetValue(valueName);
                return v is int ? (int?)(int)v : null;
            }
        }
    }
}
