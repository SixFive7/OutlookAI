using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace OutlookAI.Services
{
    /// <summary>
    /// Pure text surgery on Claude Code's MCP configuration files, split out of
    /// <see cref="McpRegistrationService"/> so the paranoid parts can be pinned by the T1
    /// unit suite (McpServer/OutlookAI.McpServer.Tests/T1/McpConfigEditorTests.cs LINKS
    /// this very file, so the tests exercise the code that ships). Nothing here touches
    /// the filesystem, the registry, Outlook or COM: every entry point takes text and
    /// returns text.
    ///
    /// Two files, one shape:
    ///  - <c>~/.claude.json</c> — user scope, applies to every project of this user;
    ///  - a project's <c>.mcp.json</c> — committed to source control, and approved once
    ///    per project by Claude Code before it is used.
    /// Both carry a top-level <c>mcpServers</c> object, so one splice discipline serves
    /// both, and it is the same discipline the user-scope reconcile has always used:
    ///  - a file that does not PARSE is never edited — it is refused and reported, because
    ///    a truncating rewrite would cost the user everything else in that file;
    ///  - only our own member is re-rendered; every other byte is copied through, so no
    ///    unrelated setting is reformatted, reordered or lost;
    ///  - the result is re-parsed AND cross-checked (our command reads back, and no other
    ///    server or top-level setting was gained or lost) BEFORE the caller may write it.
    ///
    /// Framework-neutral by construction: no <c>JavaScriptSerializer</c> (net48-only) and
    /// no <c>System.Text.Json</c> (absent on net48), so the same source compiles into the
    /// net48 add-in and into the .NET 10 test host.
    /// </summary>
    internal static class McpConfigEditor
    {
        /// <summary>
        /// Server name under <c>mcpServers</c>, in both scopes. From
        /// <see cref="AddInServerContract"/>: the MCP server reads this same member out of
        /// <c>~/.claude.json</c> to answer "am I the server Claude Code is configured to spawn?",
        /// so the write side and the read side share the one definition instead of each spelling
        /// it out.
        /// </summary>
        internal const string ServerName = AddInServerContract.ServerName;

        /// <summary>The top-level container both file shapes share; read by the server too.</summary>
        internal const string ServersProperty = AddInServerContract.ServersProperty;

        /// <summary>Project-scope file name, at the root of the project folder.</summary>
        internal const string ProjectConfigFileName = ".mcp.json";

        /// <summary>
        /// The portable spelling of the installed server. Claude Code expands
        /// <c>${VAR}</c> / <c>${VAR:-default}</c> in <c>command</c>, <c>args</c> and
        /// <c>env</c> in BOTH file shapes, so this survives a roaming profile, a renamed
        /// user, or a teammate cloning the repo on their own machine — none of which a
        /// resolved absolute path survives. Forward slashes on purpose: they need no JSON
        /// escaping, and Windows accepts them.
        /// </summary>
        internal const string PortableInstalledCommand =
            "${LOCALAPPDATA}/OutlookAI/Setup/McpServer/OutlookAI.McpServer.exe";

        /// <summary>Nesting limit for the validator, so a hostile/broken file cannot blow the stack.</summary>
        private const int MaxDepth = 200;

        // ===== Environment-variable references =====

        /// <summary>
        /// Expands the <c>${VAR}</c> and <c>${VAR:-default}</c> forms Claude Code documents
        /// for <c>command</c>/<c>args</c>/<c>env</c>. Anything else — a bare <c>$</c>, an
        /// unterminated <c>${</c> — is literal text and is copied through untouched, which
        /// is the safe reading for a path we are about to compare against a real file.
        /// <paramref name="lookup"/> returns "" for a variable that is not set.
        /// </summary>
        internal static string ExpandEnvironmentReferences(string value, Func<string, string> lookup)
        {
            if (string.IsNullOrEmpty(value) || value.IndexOf("${", StringComparison.Ordinal) < 0)
                return value ?? "";
            if (lookup == null)
                return value;

            var sb = new StringBuilder(value.Length);
            int i = 0;
            while (i < value.Length)
            {
                if (value[i] != '$' || i + 1 >= value.Length || value[i + 1] != '{')
                {
                    sb.Append(value[i]);
                    i++;
                    continue;
                }

                int close = value.IndexOf('}', i + 2);
                if (close < 0)
                {
                    // Unterminated: literal, not an error. Copy the rest verbatim.
                    sb.Append(value, i, value.Length - i);
                    break;
                }

                string inner = value.Substring(i + 2, close - i - 2);
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
                    // "${}" names nothing. Degenerate, so literal — same reasoning as an
                    // unterminated reference: never invent a path.
                    sb.Append(value, i, close - i + 1);
                    i = close + 1;
                    continue;
                }

                string resolved = "";
                try { resolved = lookup(name) ?? ""; }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine("env lookup: " + ex.Message); }

                sb.Append(resolved.Length > 0 ? resolved : fallback);
                i = close + 1;
            }

            return sb.ToString();
        }

        /// <summary>Whether a value carries a <c>${...}</c> reference for Claude Code to expand.</summary>
        internal static bool ContainsEnvironmentReference(string value)
        {
            return !string.IsNullOrEmpty(value) && value.IndexOf("${", StringComparison.Ordinal) >= 0;
        }

        /// <summary>The process environment as an expansion lookup ("" for anything unset).</summary>
        internal static Func<string, string> ProcessEnvironmentLookup()
        {
            return name =>
            {
                try { return Environment.GetEnvironmentVariable(name) ?? ""; }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("env read: " + ex.Message);
                    return "";
                }
            };
        }

        /// <summary>
        /// What to register for a resolved server executable: the portable
        /// <see cref="PortableInstalledCommand"/> when it expands to exactly that file on
        /// this machine, otherwise the resolved path itself. The fallback is the point —
        /// a developer build, or an install into a non-default directory, must still be
        /// registered correctly rather than aspirationally.
        /// </summary>
        internal static string PreferredCommand(string resolvedServerPath, Func<string, string> lookup)
        {
            if (string.IsNullOrEmpty(resolvedServerPath))
                return "";
            return SamePath(ExpandEnvironmentReferences(PortableInstalledCommand, lookup), resolvedServerPath)
                ? PortableInstalledCommand
                : resolvedServerPath;
        }

        /// <summary>Whether two paths name the same file, tolerating case and separator differences.</summary>
        internal static bool SamePath(string a, string b)
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
            catch (Exception ex)
            {
                // A path with invalid characters is not a path we can normalize; comparing
                // the raw text is still better than throwing out of a reconcile.
                System.Diagnostics.Debug.WriteLine("path compare: " + ex.Message);
                return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
            }
        }

        // The opt-in toggle, and every "may Outlook change this by itself?" question, live in
        // McpRegistrationDecision — also pure, also linked into the T1 suite. This file stays
        // what its name says: text in, text out.

        // ===== Reading: the one state "empty" must not be read as "nothing configured" =====

        /// <summary>
        /// Whether a config file that is ON DISK read back with no content at all.
        ///
        /// This is the difference between "there is no configuration yet" and "we failed to
        /// read the configuration", and everything downstream turns on it. Both files here
        /// are rewritten by the Claude Code CLI, which truncates and then flushes; reads are
        /// deliberately opened with <c>FileShare.ReadWrite</c> so the CLI holding the file
        /// open does not fail us, and that is exactly what makes the window between those two
        /// steps observable. A zero-length read taken for "nothing configured yet" would have
        /// the caller write a fresh one-property document over a config it merely failed to
        /// read — costing the user every other setting in it.
        ///
        /// So: only a file that is genuinely ABSENT may take a create-new path. A file that
        /// exists and reads empty is unreadable, and unreadable means untouched.
        /// </summary>
        internal static bool ExistsButReadsEmpty(bool fileExists, string raw)
        {
            return fileExists && (raw == null || raw.Trim().Length == 0);
        }

        // ===== Rendering =====

        /// <summary>The stdio entry as one line — used when splicing into a file we did not format.</summary>
        internal static string RenderEntryInline(string command)
        {
            return "{ \"type\": \"stdio\", \"command\": \"" + EscapeJsonString(command ?? "")
                   + "\", \"args\": [], \"env\": {} }";
        }

        /// <summary>
        /// A whole new project <c>.mcp.json</c>. Pretty-printed with LF endings because this
        /// file is normally committed and read in diffs.
        /// </summary>
        internal static string RenderNewProjectDocument(string command)
        {
            return "{\n"
                 + "  \"" + ServersProperty + "\": {\n"
                 + "    \"" + ServerName + "\": {\n"
                 + "      \"type\": \"stdio\",\n"
                 + "      \"command\": \"" + EscapeJsonString(command ?? "") + "\",\n"
                 + "      \"args\": [],\n"
                 + "      \"env\": {}\n"
                 + "    }\n"
                 + "  }\n"
                 + "}\n";
        }

        /// <summary>JSON string-body escaping (the surrounding quotes are the caller's).</summary>
        internal static string EscapeJsonString(string s)
        {
            if (s == null)
                return "";
            var sb = new StringBuilder(s.Length + 8);
            foreach (char c in s)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < ' ')
                            sb.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture));
                        else
                            sb.Append(c);
                        break;
                }
            }
            return sb.ToString();
        }

        // ===== Builders =====

        /// <summary>
        /// Produces the content of a project <c>.mcp.json</c> that registers our server,
        /// MERGING into <paramref name="raw"/> rather than replacing it: other servers and
        /// every other byte survive. A file that is genuinely ABSENT
        /// (<paramref name="fileExists"/> false) yields a fresh pretty-printed document; a
        /// file that does not parse is REFUSED (false + <paramref name="error"/>) because
        /// rewriting a file we cannot read would destroy whatever is in it — and so is a file
        /// that exists yet reads empty, see <see cref="ExistsButReadsEmpty"/>.
        ///
        /// An <c>outlookai</c> entry that already names this command is left byte-identical;
        /// one naming something else is replaced wholesale, because it is our entry and the
        /// button's promise is that afterwards it says exactly this.
        /// </summary>
        internal static bool TryBuildProjectConfig(
            string raw, bool fileExists, string command, out string updated, out string error)
        {
            updated = "";
            error = "";

            if (string.IsNullOrEmpty(command))
            {
                error = "no server executable was resolved";
                return false;
            }

            raw = raw ?? "";
            string candidate;

            if (ExistsButReadsEmpty(fileExists, raw))
            {
                // On disk, yet nothing came back: a failed read, not an empty project. Writing
                // a whole new document here would replace whatever the file really holds.
                error = "the file exists but read back empty, which usually means another program is rewriting it right now";
                return false;
            }

            if (raw.Trim().Length == 0)
            {
                candidate = RenderNewProjectDocument(command);
            }
            else
            {
                if (!TryValidateJsonObject(raw, out error))
                    return false;

                string entry = RenderEntryInline(command);
                int serversStart, serversEnd;

                if (!TryFindTopLevelValueSpan(raw, ServersProperty, out serversStart, out serversEnd))
                {
                    // No mcpServers at all: add the whole property, keeping the file's own
                    // indentation for the line we introduce.
                    int brace = raw.IndexOf('{');
                    if (brace < 0)
                    {
                        error = "the file is not a JSON object";
                        return false;
                    }
                    string lead = LeadingLayoutAfter(raw, brace);
                    string separator = HasAnyContentAfterBrace(raw, brace) ? "," : "";
                    candidate = raw.Substring(0, brace + 1)
                              + lead + "\"" + ServersProperty + "\": { \"" + ServerName + "\": " + entry + " }" + separator
                              + raw.Substring(brace + 1);
                }
                else if (raw[serversStart] != '{')
                {
                    error = "the mcpServers setting is not an object";
                    return false;
                }
                else
                {
                    int memberStart, valueStart, valueEnd;
                    if (TryFindMemberSpan(raw, serversStart, ServerName, out memberStart, out valueStart, out valueEnd))
                    {
                        // An entry that already names this exact command is left completely
                        // alone — same rule as the user-scope reconcile. Re-running the
                        // button must not reformat a file under source control, and must
                        // not produce a diff that says nothing.
                        string existing;
                        candidate = TryReadServerCommand(raw, out existing)
                                    && string.Equals(existing, command, StringComparison.Ordinal)
                            ? raw
                            : raw.Substring(0, valueStart) + entry + raw.Substring(valueEnd);
                    }
                    else
                    {
                        string lead = LeadingLayoutAfter(raw, serversStart);
                        string separator = IsEmptyObject(raw, serversStart) ? "" : ",";
                        candidate = raw.Substring(0, serversStart + 1)
                                  + lead + "\"" + ServerName + "\": " + entry + separator
                                  + raw.Substring(serversStart + 1);
                    }
                }
            }

            if (!VerifyAddition(raw, candidate, command, out error))
                return false;

            updated = candidate;
            return true;
        }

        /// <summary>
        /// Produces the content of a config with our entry REMOVED — the deregistration path
        /// behind turning the "all my projects" toggle off. Absent entry, absent
        /// <c>mcpServers</c> and an empty file are all success with
        /// <paramref name="changed"/> false (nothing to do, nothing written); a file that
        /// does not parse is refused, exactly as on the way in.
        ///
        /// "Ours" means what we could have written: an <c>outlookai</c> member whose value is
        /// a JSON OBJECT. A member of that name holding anything else — a string, say — is
        /// left alone, because the presence detector that decides whether the user ever opted
        /// in does not count it either. Matching by name alone here would have a fresh install
        /// (opted out by default, so this path runs) silently delete a malformed-but-present
        /// entry the user had put there by hand.
        ///
        /// Only the member is cut, plus the comma that separated it and the indentation that
        /// belonged to it — whitespace is the sole unrelated byte this touches, and only so
        /// the file is not left with a dangling blank line.
        /// </summary>
        internal static bool TryBuildConfigWithoutServer(string raw, out string updated, out bool changed, out string error)
        {
            updated = raw ?? "";
            changed = false;
            error = "";

            if (updated.Trim().Length == 0)
                return true;

            if (!TryValidateJsonObject(updated, out error))
                return false;

            int serversStart, serversEnd;
            if (!TryFindTopLevelValueSpan(updated, ServersProperty, out serversStart, out serversEnd))
                return true;
            if (updated[serversStart] != '{')
                return true; // not an object: not something we ever wrote, so not ours to remove

            int memberStart, valueStart, valueEnd;
            if (!TryFindMemberSpan(updated, serversStart, ServerName, out memberStart, out valueStart, out valueEnd))
                return true;
            if (updated[valueStart] != '{')
                return true; // named ours, but not an object: not something we ever wrote, so not ours to remove

            int floor = serversStart + 1;
            int cutStart;
            int cutEnd;
            int afterValue = SkipWhitespace(updated, valueEnd);
            if (afterValue < updated.Length && updated[afterValue] == ',')
            {
                // A member follows: take our member, its indentation, and the separating comma.
                cutStart = TrimBackWhitespace(updated, memberStart, floor);
                cutEnd = afterValue + 1;
            }
            else
            {
                // Last member: the comma BEFORE us is the one that has to go.
                cutStart = TrimBackWhitespace(updated, memberStart, floor);
                cutEnd = valueEnd;
                if (cutStart > floor && updated[cutStart - 1] == ',')
                    cutStart--;
            }

            string candidate = updated.Substring(0, cutStart) + updated.Substring(cutEnd);
            if (!VerifyRemoval(updated, candidate, out error))
                return false;

            updated = candidate;
            changed = true;
            return true;
        }

        // ===== Verification (nothing is written until these pass) =====

        private static bool VerifyAddition(string raw, string candidate, string command, out string error)
        {
            error = "";

            if (!TryValidateJsonObject(candidate, out error))
            {
                error = "the updated file did not parse back (" + error + ")";
                return false;
            }

            string readBack;
            if (!TryReadServerCommand(candidate, out readBack) || !string.Equals(readBack, command, StringComparison.Ordinal))
            {
                error = "the updated file does not name the server";
                return false;
            }

            var expectedServers = ListServerNames(raw);
            if (!expectedServers.Contains(ServerName))
                expectedServers.Add(ServerName);
            if (!SameMultiset(expectedServers, ListServerNames(candidate)))
            {
                error = "the update would have changed which MCP servers are configured";
                return false;
            }

            var expectedKeys = ListTopLevelKeys(raw);
            if (!expectedKeys.Contains(ServersProperty))
                expectedKeys.Add(ServersProperty);
            if (!SameMultiset(expectedKeys, ListTopLevelKeys(candidate)))
            {
                error = "the update would have changed the top-level settings";
                return false;
            }

            return true;
        }

        private static bool VerifyRemoval(string raw, string candidate, out string error)
        {
            error = "";

            if (!TryValidateJsonObject(candidate, out error))
            {
                error = "the updated file did not parse back (" + error + ")";
                return false;
            }

            string readBack;
            if (TryReadServerCommand(candidate, out readBack))
            {
                error = "the entry was still present after removal";
                return false;
            }

            var expectedServers = ListServerNames(raw);
            expectedServers.Remove(ServerName);
            if (!SameMultiset(expectedServers, ListServerNames(candidate)))
            {
                error = "the removal would have changed which other MCP servers are configured";
                return false;
            }

            if (!SameMultiset(ListTopLevelKeys(raw), ListTopLevelKeys(candidate)))
            {
                error = "the removal would have changed the top-level settings";
                return false;
            }

            return true;
        }

        /// <summary>
        /// Whether two name lists hold the same names with the same multiplicities.
        ///
        /// A MULTISET comparison, deliberately: count plus a one-directional
        /// <c>Contains</c> would call <c>[a, a, b]</c> and <c>[a, b, b]</c> equal, and this is
        /// the safety assertion standing in front of every write — a check that can be
        /// satisfied by the wrong file is worse than no check at all. Order is deliberately
        /// NOT compared: a splice inserts our member at the front of a block, so the order
        /// legitimately differs from the file we started with.
        ///
        /// Internal rather than private so the T1 suite can pin it directly.
        /// </summary>
        internal static bool SameMultiset(List<string> expected, List<string> actual)
        {
            if (expected == null || actual == null)
                return ReferenceEquals(expected, actual);
            if (expected.Count != actual.Count)
                return false;

            var remaining = new List<string>(actual);
            foreach (string name in expected)
            {
                int at = remaining.IndexOf(name);
                if (at < 0)
                    return false;
                remaining.RemoveAt(at);
            }
            return true;
        }

        // ===== Readers =====

        /// <summary><c>mcpServers.outlookai.command</c>, or false when there is no such string.</summary>
        internal static bool TryReadServerCommand(string json, out string command)
        {
            command = "";
            if (string.IsNullOrEmpty(json))
                return false;

            int serversStart, serversEnd;
            if (!TryFindTopLevelValueSpan(json, ServersProperty, out serversStart, out serversEnd))
                return false;
            if (json[serversStart] != '{')
                return false;

            int memberStart, entryStart, entryEnd;
            if (!TryFindMemberSpan(json, serversStart, ServerName, out memberStart, out entryStart, out entryEnd))
                return false;
            if (json[entryStart] != '{')
                return false;

            int commandMember, commandStart, commandEnd;
            if (!TryFindMemberSpan(json, entryStart, AddInServerContract.CommandProperty, out commandMember, out commandStart, out commandEnd))
                return false;
            if (json[commandStart] != '"')
                return false;

            command = Unescape(json.Substring(commandStart + 1, commandEnd - commandStart - 2));
            return true;
        }

        /// <summary>Names under <c>mcpServers</c> (empty when absent or not an object).</summary>
        internal static List<string> ListServerNames(string json)
        {
            var names = new List<string>();
            if (string.IsNullOrEmpty(json))
                return names;

            int serversStart, serversEnd;
            if (!TryFindTopLevelValueSpan(json, ServersProperty, out serversStart, out serversEnd))
                return names;
            if (json[serversStart] != '{')
                return names;

            return ListMemberNames(json, serversStart);
        }

        /// <summary>Top-level property names (empty when the text is not a JSON object).</summary>
        internal static List<string> ListTopLevelKeys(string json)
        {
            var names = new List<string>();
            if (string.IsNullOrEmpty(json))
                return names;
            int brace = SkipWhitespace(json, 0);
            if (brace >= json.Length || json[brace] != '{')
                return names;
            return ListMemberNames(json, brace);
        }

        private static List<string> ListMemberNames(string json, int objectStart)
        {
            var names = new List<string>();
            int i = SkipWhitespace(json, objectStart + 1);
            string name;
            int memberStart, valueStart, valueEnd;
            while (TryNextMember(json, ref i, out name, out memberStart, out valueStart, out valueEnd))
                names.Add(name);
            return names;
        }

        // ===== Member walking =====

        /// <summary>
        /// Span of a DIRECT member of the object beginning at <paramref name="objectStart"/>:
        /// <paramref name="memberStart"/> is its opening key quote, and
        /// [<paramref name="valueStart"/>, <paramref name="valueEnd"/>) is its value.
        /// Members are walked one at a time and each value is skipped whole, so a same-named
        /// key nested deeper can never be mistaken for a direct member.
        /// </summary>
        internal static bool TryFindMemberSpan(
            string json, int objectStart, string name,
            out int memberStart, out int valueStart, out int valueEnd)
        {
            memberStart = -1;
            valueStart = -1;
            valueEnd = -1;
            if (string.IsNullOrEmpty(json) || objectStart < 0 || objectStart >= json.Length || json[objectStart] != '{')
                return false;

            int i = SkipWhitespace(json, objectStart + 1);
            string key;
            int ms, vs, ve;
            while (TryNextMember(json, ref i, out key, out ms, out vs, out ve))
            {
                if (key == name)
                {
                    memberStart = ms;
                    valueStart = vs;
                    valueEnd = ve;
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// Span of a TOP-LEVEL property's value in raw JSON text. Kept as its own entry point
        /// because that is what the user-scope reconcile splices, and because the nested
        /// <c>mcpServers</c> block inside a <c>projects</c> entry of <c>~/.claude.json</c>
        /// must never be mistaken for it.
        /// </summary>
        internal static bool TryFindTopLevelValueSpan(string json, string key, out int valueStart, out int valueEnd)
        {
            valueStart = -1;
            valueEnd = -1;
            if (string.IsNullOrEmpty(json))
                return false;

            int brace = SkipWhitespace(json, 0);
            if (brace >= json.Length || json[brace] != '{')
                return false;

            int memberStart;
            return TryFindMemberSpan(json, brace, key, out memberStart, out valueStart, out valueEnd);
        }

        /// <summary>
        /// Reads the member at <paramref name="i"/> and advances past its separator. False at
        /// the object's closing brace AND on anything malformed — callers here run on text the
        /// validator has already accepted, where the two cases coincide.
        /// </summary>
        private static bool TryNextMember(
            string json, ref int i,
            out string name, out int memberStart, out int valueStart, out int valueEnd)
        {
            name = "";
            memberStart = -1;
            valueStart = -1;
            valueEnd = -1;

            if (i >= json.Length || json[i] != '"')
                return false;

            int keyEnd = SkipString(json, i);
            if (keyEnd < 0)
                return false;

            memberStart = i;
            name = Unescape(json.Substring(i + 1, keyEnd - i - 2));

            int colon = SkipWhitespace(json, keyEnd);
            if (colon >= json.Length || json[colon] != ':')
                return false;

            valueStart = SkipWhitespace(json, colon + 1);
            if (valueStart >= json.Length)
                return false;

            valueEnd = SkipValue(json, valueStart);
            if (valueEnd < 0)
                return false;

            int next = SkipWhitespace(json, valueEnd);
            i = (next < json.Length && json[next] == ',') ? SkipWhitespace(json, next + 1) : next;
            return true;
        }

        /// <summary>True when the object beginning at <paramref name="objectStart"/> has no members.</summary>
        internal static bool IsEmptyObject(string json, int objectStart)
        {
            int i = SkipWhitespace(json, objectStart + 1);
            return i < json.Length && json[i] == '}';
        }

        /// <summary>
        /// The whitespace run right after an opening brace, when it spans a line — the file's
        /// own indentation for the first member. Reused as the prefix of a member we insert
        /// there, so an inserted line lands where the reader expects it instead of being
        /// jammed against the brace. "" when the object is written compactly.
        /// </summary>
        private static string LeadingLayoutAfter(string json, int brace)
        {
            int i = brace + 1;
            int start = i;
            while (i < json.Length && char.IsWhiteSpace(json[i]))
                i++;
            string run = json.Substring(start, i - start);
            return run.IndexOf('\n') >= 0 ? run : "";
        }

        /// <summary>True when anything other than the closing brace follows <paramref name="brace"/>.</summary>
        internal static bool HasAnyContentAfterBrace(string raw, int brace)
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

        private static int TrimBackWhitespace(string json, int from, int floor)
        {
            while (from > floor && char.IsWhiteSpace(json[from - 1]))
                from--;
            return from;
        }

        // ===== Lenient scanning (post-validation splicing) =====

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

        // ===== Strict validation (the gate that decides we may touch the file at all) =====

        /// <summary>
        /// Whether <paramref name="text"/> is one complete JSON OBJECT and nothing else.
        /// Strict on purpose — this is the guard in front of every write, and "it looked
        /// close enough" is how a user loses the rest of their configuration. The lenient
        /// scanners above then run only on text that got through here.
        /// </summary>
        internal static bool TryValidateJsonObject(string text, out string error)
        {
            error = "";
            if (string.IsNullOrEmpty(text))
            {
                error = "the file is empty";
                return false;
            }

            int i = SkipWhitespace(text, 0);
            if (i >= text.Length || text[i] != '{')
            {
                error = "the file is not a JSON object";
                return false;
            }

            int end = ScanValue(text, i, 0);
            if (end < 0)
            {
                error = "the file is not well-formed JSON";
                return false;
            }

            if (SkipWhitespace(text, end) != text.Length)
            {
                error = "the file has trailing content after the JSON object";
                return false;
            }

            return true;
        }

        /// <summary>Index just past a strictly well-formed JSON value at i, or -1.</summary>
        private static int ScanValue(string json, int i, int depth)
        {
            if (depth > MaxDepth || i >= json.Length)
                return -1;

            char c = json[i];
            switch (c)
            {
                case '{': return ScanObject(json, i, depth);
                case '[': return ScanArray(json, i, depth);
                case '"': return ScanString(json, i);
                case 't': return MatchLiteral(json, i, "true");
                case 'f': return MatchLiteral(json, i, "false");
                case 'n': return MatchLiteral(json, i, "null");
                default: return ScanNumber(json, i);
            }
        }

        private static int ScanObject(string json, int i, int depth)
        {
            i = SkipWhitespace(json, i + 1);
            if (i < json.Length && json[i] == '}')
                return i + 1;

            while (true)
            {
                if (i >= json.Length || json[i] != '"')
                    return -1;
                i = ScanString(json, i);
                if (i < 0)
                    return -1;

                i = SkipWhitespace(json, i);
                if (i >= json.Length || json[i] != ':')
                    return -1;

                i = ScanValue(json, SkipWhitespace(json, i + 1), depth + 1);
                if (i < 0)
                    return -1;

                i = SkipWhitespace(json, i);
                if (i >= json.Length)
                    return -1;
                if (json[i] == '}')
                    return i + 1;
                if (json[i] != ',')
                    return -1;
                i = SkipWhitespace(json, i + 1);
            }
        }

        private static int ScanArray(string json, int i, int depth)
        {
            i = SkipWhitespace(json, i + 1);
            if (i < json.Length && json[i] == ']')
                return i + 1;

            while (true)
            {
                i = ScanValue(json, i, depth + 1);
                if (i < 0)
                    return -1;

                i = SkipWhitespace(json, i);
                if (i >= json.Length)
                    return -1;
                if (json[i] == ']')
                    return i + 1;
                if (json[i] != ',')
                    return -1;
                i = SkipWhitespace(json, i + 1);
            }
        }

        /// <summary>Strict string scan: valid escapes only, no raw control characters.</summary>
        private static int ScanString(string json, int i)
        {
            i++; // opening quote
            while (i < json.Length)
            {
                char c = json[i];
                if (c == '"')
                    return i + 1;
                if (c < ' ')
                    return -1;
                if (c != '\\')
                {
                    i++;
                    continue;
                }

                i++;
                if (i >= json.Length)
                    return -1;
                char e = json[i];
                if (e == 'u')
                {
                    if (i + 4 >= json.Length)
                        return -1;
                    for (int k = 1; k <= 4; k++)
                    {
                        if (!IsHex(json[i + k]))
                            return -1;
                    }
                    i += 5;
                    continue;
                }
                if (e != '"' && e != '\\' && e != '/' && e != 'b' && e != 'f' && e != 'n' && e != 'r' && e != 't')
                    return -1;
                i++;
            }
            return -1;
        }

        private static bool IsHex(char c)
        {
            return (c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F');
        }

        private static int MatchLiteral(string json, int i, string literal)
        {
            if (i + literal.Length > json.Length)
                return -1;
            return string.CompareOrdinal(json, i, literal, 0, literal.Length) == 0 ? i + literal.Length : -1;
        }

        private static int ScanNumber(string json, int i)
        {
            int start = i;
            if (i < json.Length && json[i] == '-')
                i++;

            if (i >= json.Length || !IsDigit(json[i]))
                return -1;
            if (json[i] == '0')
                i++;
            else
                while (i < json.Length && IsDigit(json[i])) i++;

            if (i < json.Length && json[i] == '.')
            {
                i++;
                if (i >= json.Length || !IsDigit(json[i]))
                    return -1;
                while (i < json.Length && IsDigit(json[i])) i++;
            }

            if (i < json.Length && (json[i] == 'e' || json[i] == 'E'))
            {
                i++;
                if (i < json.Length && (json[i] == '+' || json[i] == '-'))
                    i++;
                if (i >= json.Length || !IsDigit(json[i]))
                    return -1;
                while (i < json.Length && IsDigit(json[i])) i++;
            }

            return i > start ? i : -1;
        }

        private static bool IsDigit(char c)
        {
            return c >= '0' && c <= '9';
        }
    }
}
