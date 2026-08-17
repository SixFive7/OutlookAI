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
    /// Keeps Claude Code's MCP registration pointing at the OutlookAI mail server that is
    /// actually installed (v3 plan D6 v2 / R11 / Phase 8).
    ///
    /// Two scopes, deliberately different in character:
    ///
    ///  - USER scope (<c>~/.claude.json</c>, every project of this user) is OPT-IN. When it
    ///    is on, this reconciles on every Outlook start — the same shape as
    ///    <see cref="OutlookTuningService"/>, and for the same reason: it heals drift instead
    ///    of relying on a one-shot install-time write. When it is off, an entry we (or an
    ///    earlier version) put there is removed and then left alone; turning the toggle off
    ///    IS the deregistration path.
    ///
    ///    Whether the opt-in is on is never GUESSED. Where the user's intent is not already
    ///    known — a first run, or a world that has changed under a stored preference — this
    ///    asks, in Outlook, and does nothing at all until it is answered. The rules are in
    ///    <see cref="McpRegistrationDecision"/> (pure, T1-pinned); this file supplies the
    ///    evidence, applies the verdict, and hands the question to
    ///    <c>TaskPane.McpRegistrationPrompt</c>. Two consequences worth stating outright:
    ///    an entry counts as ours ONLY when its command already resolves to the installed
    ///    server, so a command the user chose is never adopted, overwritten or deleted; and
    ///    while a question is outstanding nothing is written anywhere — including when the
    ///    question cannot be asked because Outlook is running in the background with no
    ///    window, which is the state agent sessions rely on.
    ///
    ///  - PROJECT scope (<c>&lt;project&gt;/.mcp.json</c>) is a one-shot, explicit action from
    ///    the settings dialog. That file is normally committed to source control, and Claude
    ///    Code asks the user to approve the server the first time it opens the project — that
    ///    prompt is the user's, not ours to answer.
    ///
    /// The files edited belong to the Claude Code CLI, are rewritten by it, and in the
    /// user-scope case are large, so the discipline is deliberately defensive:
    ///  - a file that does not PARSE is never written to (a truncating rewrite would cost the
    ///    user every project entry in it);
    ///  - nor is one that EXISTS but reads back empty — that is a failed read, not an empty
    ///    machine, and only a genuinely absent file may be created from scratch;
    ///  - only our own value is re-rendered; every other byte of the file is spliced through
    ///    unchanged, so unrelated settings cannot be reformatted or reordered;
    ///  - the result is re-parsed and cross-checked before it replaces anything;
    ///  - the replacement is atomic;
    ///  - a correct entry is a no-op — nothing is written, and no backup is churned.
    /// The text surgery itself lives in <see cref="McpConfigEditor"/>, which is unit-tested
    /// (T1) because it is the part that can silently destroy a user's configuration.
    ///
    /// Everything is per-user (HKCU + the user profile); no elevation, no COM, no Outlook
    /// object model — so this can run off the UI thread without touching the add-in's COM
    /// ownership rules.
    /// </summary>
    internal static class McpRegistrationService
    {
        /// <summary>
        /// The key this service owns. Its path, and every value name written under it, live in
        /// <see cref="AddInServerContract"/>: the MCP server's <c>outlook_health</c> reads this
        /// key to report what the last reconcile did, and that key IS the contract between the
        /// two components. One definition compiled into both, rather than a comment on each side
        /// claiming to mirror the other.
        /// </summary>
        internal const string McpKeyPath = AddInServerContract.McpKeyPath;
        internal const string InstallDirValueName = "InstallDir";
        private const string AppKeyPath = @"Software\OutlookAI";

        /// <summary>Server name under <c>mcpServers</c>; matches what Phase 2 registered.</summary>
        internal const string ServerName = McpConfigEditor.ServerName;

        internal const string RelativeServerPath = @"McpServer\OutlookAI.McpServer.exe";

        /// <summary>
        /// Tri-state opt-in for user-scope registration: 1 on, 0 off, ABSENT never decided.
        /// The absent state is load-bearing — it is what makes the difference between "the
        /// user does not want this" and "the user has not been asked yet", and only the first
        /// of those may be acted on. See <see cref="McpRegistrationDecision.Decide"/>.
        /// </summary>
        internal const string GlobalRegistrationValueName = AddInServerContract.McpGlobalRegistrationValueName;

        // Status codes. Also written to HKCU (AddInServerContract.McpStatusValueName) so the MCP
        // server can report them in outlook_health without having to guess what the add-in did.
        //
        // Note what is and is not enforced. The server surfaces this string VERBATIM as
        // registration.addInStatus and never compares it against anything, so NO COMPILATION can
        // notice a rename - the add-in goes on writing, the server goes on reporting, and only
        // the published meaning is wrong. The only other place these values appear is the list in
        // McpServer/README.md, and since that is Markdown it cannot read a C# constant either -
        // so check-pinned-constants.ps1 #7 compares the two as a SET, in both directions, exactly
        // as #6 does for the server's own registration.status vocabulary. Adding or renaming a
        // code here therefore fails the build until that list is updated to match. (Before #7
        // existed, awaiting_choice had been added here and never published there.)
        internal const string StatusOk = "ok";
        internal const string StatusHealed = "healed";
        internal const string StatusDisabled = "not_registered_by_choice";
        internal const string StatusRemoved = "removed";
        internal const string StatusNoClaude = "claude_code_not_installed";
        internal const string StatusNoServer = "server_not_installed";
        internal const string StatusNoRuntime = "dotnet_runtime_missing";
        internal const string StatusParseFailed = "config_unreadable";
        internal const string StatusAwaitingChoice = "awaiting_choice";
        internal const string StatusError = "error";

        internal const string DotnetRuntimeDownloadUrl = "https://dotnet.microsoft.com/download/dotnet/10.0";

        private static readonly object _gate = new object();

        /// <summary>
        /// Where a question goes when this decides one has to be asked. Installed by the
        /// add-in at startup (<c>TaskPane.McpRegistrationPrompt</c>); null in any host without
        /// a UI, which simply means the question is deferred to a later session.
        ///
        /// The handler is called on whatever thread the reconcile is running on — normally a
        /// background one — and MUST return immediately: it marshals to Outlook's UI thread
        /// and shows the dialog later, once startup has settled and only if a human can
        /// actually see it. Nothing here waits for the answer; the answer comes back as a
        /// separate <see cref="ApplyUserChoice"/> call.
        /// </summary>
        internal static Action<McpRegistrationDecision.PromptKind> PromptHost;

        /// <summary>
        /// Whether a question has already been handed to the host in this Outlook session.
        /// Process lifetime IS the session, which is exactly the granularity wanted: asked
        /// once, and if it went unanswered — dismissed, or never shown because nobody was
        /// looking — not asked again until Outlook next starts.
        /// </summary>
        private static bool _promptHandedToHost;

        internal sealed class RegistrationSnapshot
        {
            /// <summary>One of the Status* constants.</summary>
            public string Status { get; internal set; }

            /// <summary>Human-readable detail; never null.</summary>
            public string Detail { get; internal set; }

            /// <summary>Path this reconcile wants registered, or null when none was resolved.</summary>
            public string ResolvedServerPath { get; internal set; }

            /// <summary>Command currently registered in ~/.claude.json, or null.</summary>
            public string RegisteredCommand { get; internal set; }

            /// <summary>True when this run rewrote the registration (added, repaired or removed it).</summary>
            public bool Healed { get; internal set; }

            /// <summary>Whether user-scope ("all my projects") registration is opted in.</summary>
            public bool GlobalRegistrationEnabled { get; internal set; }

            public string LastReconcileUtc { get; internal set; }
        }

        // ===== Public surface =====

        /// <summary>
        /// Reconciles the user-scope registration. Never throws: every failure becomes a
        /// status the settings dialog and outlook_health can show.
        /// </summary>
        internal static RegistrationSnapshot Reconcile()
        {
            return Reconcile(false);
        }

        /// <summary>
        /// <paramref name="intentJustDeclared"/> is true only when the user has just said what
        /// they want — ticked the box in the settings dialog, or answered one of our prompts.
        /// It is what turns "the world disagrees with your stored preference, what now?" into
        /// simply doing what was asked. See <see cref="McpRegistrationDecision.Decide"/>.
        /// </summary>
        private static RegistrationSnapshot Reconcile(bool intentJustDeclared)
        {
            lock (_gate)
            {
                RegistrationSnapshot snap;
                try
                {
                    snap = ReconcileCore(intentJustDeclared);
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
                    Status = ReadString(AddInServerContract.McpStatusValueName) ?? StatusError,
                    Detail = ReadString(AddInServerContract.McpDetailValueName) ?? "",
                    ResolvedServerPath = ReadString(AddInServerContract.McpResolvedServerPathValueName),
                    RegisteredCommand = ReadString(AddInServerContract.McpCommandValueName),
                    Healed = ReadDword(AddInServerContract.McpHealedValueName) == 1,
                    GlobalRegistrationEnabled = ReadDword(GlobalRegistrationValueName) == 1,
                    LastReconcileUtc = ReadString(AddInServerContract.McpLastReconcileUtcValueName),
                };
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP registration snapshot: " + ex.Message);
                return new RegistrationSnapshot { Status = StatusError, Detail = ex.Message };
            }
        }

        /// <summary>
        /// Records the user's choice for user-scope registration. The following
        /// <see cref="Reconcile()"/> acts on it — on for the usual write, off for the removal.
        /// Never throws.
        /// </summary>
        internal static void SetGlobalRegistrationEnabled(bool enabled)
        {
            try { WriteDword(GlobalRegistrationValueName, enabled ? 1 : 0); }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine("MCP toggle persist: " + ex.Message); }
        }

        /// <summary>
        /// The user has just decided: remember it and act on it now, in one step. This is the
        /// single way an answer enters the system — the settings dialog's tick box and every
        /// button on the startup prompt all end up here — so the dialog and the prompt cannot
        /// end up disagreeing about what was chosen or what is registered.
        ///
        /// The reconcile that follows treats the preference as freshly declared, so it acts
        /// instead of asking again: on registers (taking the <c>outlookai</c> name over from
        /// whatever held it), off removes an entry of OURS and leaves anything else alone.
        /// Never throws.
        /// </summary>
        internal static RegistrationSnapshot ApplyUserChoice(bool enabled)
        {
            SetGlobalRegistrationEnabled(enabled);
            return Reconcile(true);
        }

        /// <summary>
        /// Whether a question is STILL worth asking, re-evaluated from the live configuration.
        /// <see cref="McpRegistrationDecision.PromptKind.None"/> means drop it.
        ///
        /// The prompt is shown seconds after the reconcile that raised it, and the world can
        /// move in between — the user may have ticked the box in the settings dialog, or the
        /// CLI may have written the entry. Re-checking here is what stops a stale question
        /// from being put to the user. Reads only; changes nothing. Never throws.
        /// </summary>
        internal static McpRegistrationDecision.PromptKind PendingPrompt()
        {
            try
            {
                lock (_gate)
                {
                    string configPath = ClaudeConfigPath();
                    bool configExists = File.Exists(configPath);
                    string raw = configExists ? ReadSharedSettled(configPath) : "";

                    Dictionary<string, object> root;
                    bool parsed = TryParseConfig(configExists, raw, out root);
                    string serverPath = ResolveInstalledServerPath();

                    var decision = McpRegistrationDecision.Decide(
                        ReadDword(GlobalRegistrationValueName),
                        ClassifyEntry(parsed, root, serverPath),
                        canAskUser: true,
                        intentJustDeclared: false);

                    string status, detail;
                    if (decision.Prompt == McpRegistrationDecision.PromptKind.None
                        || !CanAnswerPrompt(decision.Prompt, configExists, serverPath, out status, out detail))
                    {
                        return McpRegistrationDecision.PromptKind.None;
                    }

                    return decision.Prompt;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP pending prompt: " + ex.Message);
                return McpRegistrationDecision.PromptKind.None;
            }
        }

        /// <summary>
        /// Whether a question may be handed to the host at all: there has to BE a host, and
        /// this session must not have asked already. Deliberately does not ask whether a
        /// window is visible — at the moment the startup reconcile runs, a perfectly normal
        /// interactive Outlook has not painted its window yet. That check belongs to the host,
        /// which makes it when it is about to show the dialog rather than minutes earlier.
        /// </summary>
        private static bool CanAskUser()
        {
            return PromptHost != null && !_promptHandedToHost;
        }

        /// <summary>
        /// The command a manual <c>claude mcp add</c> would need, for the settings dialog's
        /// copy button. The portable <c>${LOCALAPPDATA}</c> spelling when this machine's
        /// install is in the default place, the resolved path otherwise; "" when no server
        /// was found.
        /// </summary>
        internal static string ResolvePreferredCommand()
        {
            try
            {
                string serverPath = ResolveInstalledServerPath();
                if (serverPath == null)
                    return "";
                return McpConfigEditor.PreferredCommand(serverPath, McpConfigEditor.ProcessEnvironmentLookup());
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP preferred command: " + ex.Message);
                return "";
            }
        }

        /// <summary>
        /// Writes (or merges into) <c>.mcp.json</c> in <paramref name="folderPath"/> so this
        /// one project gets the mail server. Never throws.
        ///
        /// Merge, never overwrite: other servers and every other byte survive, and a file that
        /// does not parse is refused outright rather than replaced. Nothing here approves the
        /// server for the project — Claude Code asks the user on first use of that folder, and
        /// answering that prompt on their behalf is not ours to do.
        /// </summary>
        internal static bool TryRegisterInProject(string folderPath, out string configPath, out string error)
        {
            configPath = "";
            error = "";

            try
            {
                if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
                {
                    error = "That folder no longer exists.";
                    return false;
                }

                string serverPath = ResolveInstalledServerPath();
                if (serverPath == null)
                {
                    error = "The OutlookAI mail server was not found next to the installed add-in, so there is nothing to register.";
                    return false;
                }

                string command = McpConfigEditor.PreferredCommand(serverPath, McpConfigEditor.ProcessEnvironmentLookup());
                configPath = Path.Combine(folderPath, McpConfigEditor.ProjectConfigFileName);

                // Whether the file is THERE is as load-bearing as what it says: only a
                // genuinely absent one may be created from scratch. See
                // McpConfigEditor.ExistsButReadsEmpty.
                bool fileExists = File.Exists(configPath);
                string raw = "";
                if (fileExists)
                {
                    try
                    {
                        raw = ReadSharedSettled(configPath);
                    }
                    catch (Exception ex)
                    {
                        error = "Could not read the existing " + McpConfigEditor.ProjectConfigFileName + ": " + ex.Message;
                        return false;
                    }
                }

                string updated, buildError;
                if (!McpConfigEditor.TryBuildProjectConfig(raw, fileExists, command, out updated, out buildError))
                {
                    error = "The existing " + McpConfigEditor.ProjectConfigFileName
                            + " could not be updated safely — " + buildError + ". It was left untouched.";
                    return false;
                }

                // Already exactly right: no write, no backup churn, no pointless source-control diff.
                if (string.Equals(updated, raw, StringComparison.Ordinal))
                    return true;

                string writeError;
                if (!TryWriteAtomically(configPath, updated, false, out writeError))
                {
                    error = "Could not write " + configPath + ": " + writeError;
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        // ===== Reconcile =====

        private static RegistrationSnapshot ReconcileCore(bool intentJustDeclared)
        {
            string configPath = ClaudeConfigPath();
            string serverPath = ResolveInstalledServerPath();

            var snap = new RegistrationSnapshot
            {
                ResolvedServerPath = serverPath,
                Detail = "",
                Healed = false,
            };

            // Read the config exactly once: everything below is a decision about this text.
            bool configExists;
            string raw;
            try
            {
                configExists = File.Exists(configPath);
                raw = configExists ? ReadSharedSettled(configPath) : "";
            }
            catch (Exception ex)
            {
                snap.Status = StatusError;
                snap.Detail = "Could not read " + configPath + ": " + ex.Message;
                snap.GlobalRegistrationEnabled = ReadDword(GlobalRegistrationValueName) == 1;
                return snap;
            }

            Dictionary<string, object> root;
            bool parsed = TryParseConfig(configExists, raw, out root);
            snap.RegisteredCommand = parsed ? ReadRegisteredCommand(root) : null;

            // (0) The two facts every decision below is made from: what the user has already
            // said (tri-state, and "nothing" is a state), and what is actually in the file.
            var entry = ClassifyEntry(parsed, root, serverPath);
            int? stored = ReadDword(GlobalRegistrationValueName);

            // Resolved before anything else can stand in the way, so the settings dialog
            // always has a state to show. An entry that is ALREADY ours is the only thing
            // that counts as having opted in without saying so.
            snap.GlobalRegistrationEnabled = McpRegistrationDecision.ResolveOptIn(
                stored, entry == McpRegistrationDecision.EntryState.Ours);

            var decision = McpRegistrationDecision.Decide(stored, entry, CanAskUser(), intentJustDeclared);

            // (1) A question is outstanding. NOTHING is written — not the config, not the
            // opt-in — until it is answered, whether or not it can be put to anyone now.
            if (decision.Prompt != McpRegistrationDecision.PromptKind.None)
                return AskTheUser(snap, decision.Prompt, configExists, serverPath);

            if (decision.Deferred)
            {
                snap.Status = StatusAwaitingChoice;
                snap.Detail = "OutlookAI needs you to decide what happens to the mail server's entry in Claude Code, "
                              + "and could not ask this time. Nothing was changed; it will ask again the next time "
                              + "Outlook starts.";
                return snap;
            }

            // (2) Nothing to do at all: the entry is off and absent, an entry that is not ours
            // is none of our business, or the file could not be read.
            if (decision.Action == McpRegistrationDecision.RegistrationAction.None)
                return NothingToDo(snap, entry);

            if (decision.Action == McpRegistrationDecision.RegistrationAction.Remove)
                return ReconcileOptedOut(snap, configPath, raw, parsed);

            // (3) No Claude Code on this machine: nothing to register into. Not an error —
            // the add-in is perfectly usable on its own.
            if (!IsClaudeCodeInstalled(configExists))
            {
                snap.Status = StatusNoClaude;
                snap.Detail = "Claude Code was not found on this machine, so there is nothing to register with.";
                return snap;
            }

            // (4) Adopting an entry that is already ours: record the opt-in BEFORE the
            // environment gates below can return early. It is a fact about this machine —
            // the entry is there and it names our server — and leaving it undecided merely
            // because today's runtime is missing would have this ask a question tomorrow
            // that has effectively already been answered.
            if (decision.Action == McpRegistrationDecision.RegistrationAction.AdoptAndRegister)
            {
                try { WriteDword(GlobalRegistrationValueName, 1); }
                catch (Exception ex) { System.Diagnostics.Debug.WriteLine("MCP toggle adopt: " + ex.Message); }
                snap.GlobalRegistrationEnabled = true;
            }

            // (5) Add-in installed without the mail server beside it (an older install layout,
            // or a developer add-in build). Leave whatever is registered alone — in the
            // developer case that entry is deliberately pointing at a build output.
            if (serverPath == null)
            {
                snap.Status = StatusNoServer;
                snap.Detail = "The OutlookAI mail server was not found next to the installed add-in, so the existing registration was left unchanged.";
                return snap;
            }

            // (6) The server needs the .NET 10 runtime. Registering a server that cannot start
            // would surface as a failed MCP server in every Claude session, so report instead
            // and heal on a later start once the runtime is there.
            if (!IsDotnetRuntime10Installed())
            {
                snap.Status = StatusNoRuntime;
                snap.Detail = "The .NET 10 runtime the mail server needs is not installed, so registration was skipped. Install it from " + DotnetRuntimeDownloadUrl + " and restart Outlook.";
                return snap;
            }

            if (!parsed)
            {
                // Unreachable — an unreadable file decides to do nothing above — but this is
                // the guard that matters most, so it stays: never rewrite a file we could not
                // understand, whatever route got here. That is how config gets lost.
                snap.Status = StatusParseFailed;
                snap.Detail = "Claude Code's configuration file could not be read as JSON, so it was left untouched.";
                return snap;
            }

            // The portable ${LOCALAPPDATA} spelling wherever it resolves to this very file:
            // it survives a roaming profile and a renamed user, which a resolved absolute
            // path does not. A developer build, or an install outside the default directory,
            // falls back to the resolved path so it is still registered truthfully.
            string desiredCommand = McpConfigEditor.PreferredCommand(
                serverPath, McpConfigEditor.ProcessEnvironmentLookup());

            // (7) Already correct — do nothing at all. No write, no backup churn. The
            // registered spelling has to be the one we would write; an equivalent absolute
            // path also counts, but ONLY when the portable form is not what we would write,
            // so an entry left by an earlier version is upgraded to the portable spelling
            // exactly once instead of being accepted forever.
            bool commandCorrect =
                string.Equals(snap.RegisteredCommand, desiredCommand, StringComparison.OrdinalIgnoreCase)
                || (!McpConfigEditor.ContainsEnvironmentReference(desiredCommand)
                    && IsSamePath(snap.RegisteredCommand, desiredCommand));

            bool entryIsOurs = entry == McpRegistrationDecision.EntryState.Ours;

            if (entryIsOurs && commandCorrect && IsStdioEntry(root) && !HasRemoteKeys(root))
            {
                snap.Status = StatusOk;
                snap.Detail = "Registered for all your projects, pointing at the installed mail server.";
                return snap;
            }

            string updated;
            string spliceError;
            if (!TryBuildUpdatedConfig(raw, configExists, root, desiredCommand, serverPath, entryIsOurs, out updated, out spliceError))
            {
                snap.Status = StatusError;
                snap.Detail = "Could not update the registration safely: " + spliceError + ". The file was left untouched.";
                return snap;
            }

            string writeError;
            if (!TryWriteAtomically(configPath, updated, true, out writeError))
            {
                snap.Status = StatusError;
                snap.Detail = "Could not write the registration: " + writeError;
                return snap;
            }

            snap.Status = StatusHealed;
            snap.Healed = true;
            snap.RegisteredCommand = desiredCommand;
            snap.Detail = "Registration was missing or pointing elsewhere and has been repaired.";
            return snap;
        }

        /// <summary>
        /// A question has to be asked. Hands it to the host and reports that it is outstanding;
        /// writes nothing whatsoever, here or on the way back.
        ///
        /// Some questions are not worth asking on this machine — offering to register when
        /// Claude Code is not installed, or when there is no server to register — so those are
        /// reported as the environment problem they are instead.
        /// </summary>
        private static RegistrationSnapshot AskTheUser(
            RegistrationSnapshot snap,
            McpRegistrationDecision.PromptKind prompt,
            bool configExists,
            string serverPath)
        {
            string status, detail;
            if (!CanAnswerPrompt(prompt, configExists, serverPath, out status, out detail))
            {
                snap.Status = status;
                snap.Detail = detail;
                return snap;
            }

            // Marked BEFORE the hand-off, and never unmarked: one question per Outlook
            // session, answered or not. A host that fails to show it costs this session its
            // question, which is the right way round — silence beats nagging.
            _promptHandedToHost = true;

            var host = PromptHost;
            try
            {
                // Returns immediately by contract (it posts to the UI thread); this call is
                // made under the reconcile lock and must not wait for anything.
                if (host != null)
                    host(prompt);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("MCP prompt hand-off: " + ex.Message);
            }

            snap.Status = StatusAwaitingChoice;
            snap.Detail = DescribePendingChoice(prompt);
            return snap;
        }

        /// <summary>Why the settings dialog is showing "waiting for you", in the user's terms.</summary>
        private static string DescribePendingChoice(McpRegistrationDecision.PromptKind prompt)
        {
            switch (prompt)
            {
                case McpRegistrationDecision.PromptKind.ForeignEntry:
                    return "Your Claude Code configuration has an \"outlookai\" entry pointing somewhere else. It "
                           + "was left untouched; OutlookAI will ask what to do with it.";
                case McpRegistrationDecision.PromptKind.EntryMissing:
                    return "Registration is on, but the entry is no longer in Claude Code's configuration. Nothing "
                           + "was changed; OutlookAI will ask whether to put it back.";
                case McpRegistrationDecision.PromptKind.EntryUnexpected:
                    return "Registration is off, yet an entry for OutlookAI's mail server is present. It was left "
                           + "untouched; OutlookAI will ask what to do.";
                default:
                    return "OutlookAI will ask whether to register the mail server for all your Claude Code "
                           + "projects. Nothing has been changed yet.";
            }
        }

        /// <summary>
        /// Whether a question is worth putting to the user on this machine, or whether the
        /// environment answers it first. Three of the four offer to REGISTER something, so
        /// they need a Claude Code, a server and a runtime to be answerable at all; the
        /// fourth ("this entry of ours is still here — remove it?") needs none of that.
        /// </summary>
        private static bool CanAnswerPrompt(
            McpRegistrationDecision.PromptKind prompt,
            bool configExists,
            string serverPath,
            out string status,
            out string detail)
        {
            status = "";
            detail = "";

            if (!IsClaudeCodeInstalled(configExists))
            {
                status = StatusNoClaude;
                detail = "Claude Code was not found on this machine, so there is nothing to register with.";
                return false;
            }

            if (prompt == McpRegistrationDecision.PromptKind.EntryUnexpected)
                return true;

            if (serverPath == null)
            {
                status = StatusNoServer;
                detail = "The OutlookAI mail server was not found next to the installed add-in, so there is nothing to register.";
                return false;
            }

            if (!IsDotnetRuntime10Installed())
            {
                status = StatusNoRuntime;
                detail = "The .NET 10 runtime the mail server needs is not installed, so registration was skipped. Install it from " + DotnetRuntimeDownloadUrl + " and restart Outlook.";
                return false;
            }

            return true;
        }

        /// <summary>
        /// Everything is already as it should be, or is none of our business. No write, and
        /// deliberately no removal: an <c>outlookai</c> entry that is not ours stays exactly
        /// where the user put it.
        /// </summary>
        private static RegistrationSnapshot NothingToDo(
            RegistrationSnapshot snap, McpRegistrationDecision.EntryState entry)
        {
            if (entry == McpRegistrationDecision.EntryState.Unreadable)
            {
                snap.Status = StatusParseFailed;
                snap.Detail = "Claude Code's configuration file could not be read as JSON, so it was left untouched.";
                return snap;
            }

            snap.Status = StatusDisabled;
            snap.Detail = entry == McpRegistrationDecision.EntryState.Foreign
                ? "Not registered for all your projects — that setting is off. The \"outlookai\" entry in your Claude "
                  + "Code configuration points at something else and was left untouched."
                : "Not registered for all your projects — that setting is off. Individual projects with a .mcp.json are unaffected.";

            if (entry != McpRegistrationDecision.EntryState.Foreign)
                snap.RegisteredCommand = null;

            return snap;
        }

        /// <summary>
        /// The toggle is off AND the entry is one of ours: remove it, then stop maintaining
        /// one. This is the deregistration path — there is no separate uninstall hook, and
        /// there deliberately is not one, because a user who turns this off expects the entry
        /// gone now rather than at some future uninstall.
        ///
        /// Only ever reached for <see cref="McpRegistrationDecision.EntryState.Ours"/>. That
        /// is the whole safety argument: an <c>outlookai</c> entry naming a command the user
        /// chose is never deleted by us, so answering "leave it alone" to the prompt cannot be
        /// undone by the very next Outlook start.
        /// </summary>
        private static RegistrationSnapshot ReconcileOptedOut(
            RegistrationSnapshot snap, string configPath, string raw, bool parsed)
        {
            if (!parsed)
            {
                snap.Status = StatusParseFailed;
                snap.Detail = "Claude Code's configuration file could not be read as JSON, so nothing was removed from it.";
                return snap;
            }

            string updated, changedError;
            bool changed;
            if (!McpConfigEditor.TryBuildConfigWithoutServer(raw, out updated, out changed, out changedError))
            {
                snap.Status = StatusError;
                snap.Detail = "Could not remove the registration safely: " + changedError + ". The file was left untouched.";
                return snap;
            }

            if (!changed)
            {
                snap.Status = StatusDisabled;
                snap.Detail = "Not registered for all your projects — that setting is off. Individual projects with a .mcp.json are unaffected.";
                snap.RegisteredCommand = null;
                return snap;
            }

            string writeError;
            if (!TryWriteAtomically(configPath, updated, true, out writeError))
            {
                snap.Status = StatusError;
                snap.Detail = "Could not remove the registration: " + writeError;
                return snap;
            }

            snap.Status = StatusRemoved;
            snap.Healed = true;
            snap.RegisteredCommand = null;
            snap.Detail = "Removed from Claude Code's personal configuration, because the \"all my projects\" setting is off.";
            return snap;
        }

        // ===== Paths and probes =====

        /// <summary>
        /// The user-scope configuration file. The file name comes from
        /// <see cref="AddInServerContract"/> because the MCP server reads the very same file to
        /// report whether the registration names it - see
        /// <c>HealthReporting.TryReadRegisteredCommand</c>, which resolves it the same way.
        /// </summary>
        internal static string ClaudeConfigPath()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                AddInServerContract.ClaudeConfigFileName);
        }

        /// <summary>
        /// Claude Code is present if it has written its config file, or if its CLI is where
        /// the add-in already expects it (see <see cref="ClaudeService"/>).
        /// </summary>
        private static bool IsClaudeCodeInstalled(bool configExists)
        {
            try
            {
                if (configExists)
                    return true;
                // The one spelling of that path, shared with the code that actually RUNS the
                // CLI. Two spellings meant a move would fix the executing path - which fails
                // loudly on the next click - and leave this one quietly reporting the wrong
                // status in the settings dialog forever.
                return File.Exists(ClaudeService.ClaudePath);
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// How far up from this assembly to look for {app}\McpServer\. An installed add-in
        /// lives under {app}\Application Files\OutlookAI_x_y_z_w\, which is two levels, and a
        /// developer build sits somewhere else again - so this is a search bound rather than a
        /// statement about the layout. Running out of levels returns null, which every caller
        /// already handles as "no server installed".
        /// </summary>
        private const int ServerSearchParentLevels = 4;

        /// <summary>
        /// The installed server executable, or null when there isn't one. Prefers the install
        /// directory the installer recorded; falls back to walking up from this assembly, at
        /// most <see cref="ServerSearchParentLevels"/> levels.
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
                for (int i = 0; i < ServerSearchParentLevels && !string.IsNullOrEmpty(dir); i++)
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
                // The 64-bit registry view and the 64-bit Program Files explicitly: Outlook
                // may be a 32-bit process, and under WOW64 the default view would redirect
                // into Wow6432Node and to the 32-bit Program Files - neither of which is
                // where the x64 runtime this server needs lives.
                string root = null;
                using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                using (var key = hklm.OpenSubKey(@"SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedhost"))
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
                    string programFiles = Environment.GetEnvironmentVariable("ProgramW6432");
                    if (string.IsNullOrEmpty(programFiles))
                        programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                    root = Path.Combine(programFiles, @"dotnet\shared\Microsoft.NETCore.App");
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

        /// <summary>
        /// Classifies the file text once: false when it must never be written to.
        ///
        /// A file that is on disk yet reads back with nothing in it is a FAILED READ, not a
        /// machine with no configuration — the CLI truncates before it flushes, and the shared
        /// read is what makes that window visible. Scoring it "no config yet" would have the
        /// splice write a fresh one-property document, and the atomic replace would then move
        /// the user's real configuration aside as a backup the CLI's own open handle promptly
        /// overwrites. So it is classified exactly as a file that does not parse: unreadable,
        /// and therefore untouched.
        ///
        /// A genuinely ABSENT file is readable-with-nothing-in-it: true, with a null
        /// <paramref name="root"/>. That is the one state a whole new document may be built for.
        /// </summary>
        private static bool TryParseConfig(bool configExists, string raw, out Dictionary<string, object> root)
        {
            root = null;

            if (McpConfigEditor.ExistsButReadsEmpty(configExists, raw))
                return false;
            if (raw == null || raw.Trim().Length == 0)
                return true;

            root = TryParseObject(raw);
            return root != null;
        }

        private static string ReadRegisteredCommand(Dictionary<string, object> root)
        {
            var entry = ReadServerEntry(root);
            if (entry == null)
                return null;
            object command;
            return entry.TryGetValue(AddInServerContract.CommandProperty, out command) ? command as string : null;
        }

        /// <summary>The <c>outlookai</c> member's value, whatever shape it is; false when there is none.</summary>
        private static bool TryGetServerMember(Dictionary<string, object> root, out object value)
        {
            value = null;
            if (root == null)
                return false;
            object servers;
            if (!root.TryGetValue(McpConfigEditor.ServersProperty, out servers))
                return false;
            var map = servers as Dictionary<string, object>;
            if (map == null)
                return false;
            return map.TryGetValue(ServerName, out value);
        }

        private static Dictionary<string, object> ReadServerEntry(Dictionary<string, object> root)
        {
            object value;
            if (!TryGetServerMember(root, out value))
                return null;
            return value as Dictionary<string, object>;
        }

        /// <summary>
        /// What <c>mcpServers.outlookai</c> IS — the single judgement everything else hangs off.
        ///
        /// "Ours" is deliberately narrow: a stdio entry whose <c>command</c>, once Claude Code's
        /// <c>${VAR}</c> forms are expanded, names the mail server installed on this machine.
        /// Anything else is FOREIGN, and foreign means untouchable — not adopted as a silent
        /// opt-in, not overwritten, not deleted. The case that motivates the whole rule is a
        /// user who ran <c>claude mcp add --scope user outlookai -- C:\my\wrapper.cmd</c> on
        /// purpose: treating that as our own entry would quietly destroy it, in either
        /// direction, on the next Outlook start.
        ///
        /// A remote entry (<c>url</c>) is foreign for a second reason as well: rewriting it as
        /// a stdio command while its <c>url</c> stayed behind would leave a hybrid that is
        /// neither.
        /// </summary>
        internal static McpRegistrationDecision.EntryState ClassifyEntry(
            bool parsed, Dictionary<string, object> root, string installedServerPath)
        {
            if (!parsed)
                return McpRegistrationDecision.EntryState.Unreadable;

            object value;
            if (!TryGetServerMember(root, out value))
                return McpRegistrationDecision.EntryState.Absent;

            var entry = value as Dictionary<string, object>;
            if (entry == null)
                return McpRegistrationDecision.EntryState.Foreign;

            if (entry.ContainsKey("url") || entry.ContainsKey("httpUrl"))
                return McpRegistrationDecision.EntryState.Foreign;

            object type;
            if (entry.TryGetValue("type", out type)
                && !string.Equals(type as string, "stdio", StringComparison.OrdinalIgnoreCase))
            {
                return McpRegistrationDecision.EntryState.Foreign;
            }

            object command;
            entry.TryGetValue(AddInServerContract.CommandProperty, out command);
            string text = command as string;
            if (string.IsNullOrEmpty(text))
                return McpRegistrationDecision.EntryState.Foreign;

            string expanded = McpConfigEditor.ExpandEnvironmentReferences(
                text, McpConfigEditor.ProcessEnvironmentLookup());
            if (string.IsNullOrEmpty(expanded))
                return McpRegistrationDecision.EntryState.Foreign;

            if (!string.IsNullOrEmpty(installedServerPath))
            {
                return IsSamePath(expanded, installedServerPath)
                    ? McpRegistrationDecision.EntryState.Ours
                    : McpRegistrationDecision.EntryState.Foreign;
            }

            // No installed server to compare against (an older install layout, or a developer
            // add-in build). Our executable's own file name is the only evidence left — and it
            // is enough for the one thing that still matters here: telling a stale entry of
            // OURS, which the deregistration path may remove, from a command the user chose,
            // which it may not.
            return NamesOurServerExecutable(expanded)
                ? McpRegistrationDecision.EntryState.Ours
                : McpRegistrationDecision.EntryState.Foreign;
        }

        private static bool NamesOurServerExecutable(string command)
        {
            try
            {
                return string.Equals(
                    Path.GetFileName(command),
                    Path.GetFileName(RelativeServerPath),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                // A value with invalid path characters names no executable of ours.
                System.Diagnostics.Debug.WriteLine("server name compare: " + ex.Message);
                return false;
            }
        }

        /// <summary>
        /// Whether our entry still carries keys that only belong to a REMOTE server. They are
        /// meaningless beside a <c>command</c>, and an entry holding both is exactly the hybrid
        /// an earlier version could produce by overwriting a remote entry key by key — so its
        /// presence means "not yet correct", and the rewrite below drops them.
        /// </summary>
        private static bool HasRemoteKeys(Dictionary<string, object> root)
        {
            var entry = ReadServerEntry(root);
            if (entry == null)
                return false;
            return entry.ContainsKey("url") || entry.ContainsKey("httpUrl") || entry.ContainsKey("headers");
        }

        /// <summary>
        /// Whether the registered entry is a stdio server this reconcile may keep maintaining.
        ///
        /// A hand-written entry legitimately omits <c>type</c> — stdio is what a <c>command</c>
        /// means — so requiring an explicit type would score every such entry as drift and
        /// rewrite the file on every single Outlook start. A remote entry (<c>url</c>) is the
        /// case that must still score false: taking one of those over would break it.
        /// </summary>
        private static bool IsStdioEntry(Dictionary<string, object> root)
        {
            var entry = ReadServerEntry(root);
            if (entry == null)
                return false;

            object type;
            if (entry.TryGetValue("type", out type))
                return string.Equals(type as string, "stdio", StringComparison.OrdinalIgnoreCase);

            object command;
            if (!entry.TryGetValue(AddInServerContract.CommandProperty, out command) || string.IsNullOrEmpty(command as string))
                return false;

            return !entry.ContainsKey("url") && !entry.ContainsKey("httpUrl");
        }

        internal static bool IsSamePath(string a, string b)
        {
            return McpConfigEditor.SamePath(a, b);
        }

        /// <summary>
        /// Produces the new file content by replacing ONLY the <c>mcpServers</c> value (or
        /// inserting the whole property when absent). Everything outside that span is copied
        /// through byte for byte, so no unrelated setting is reformatted. The result is
        /// re-parsed and checked before it is returned.
        ///
        /// <paramref name="command"/> is what gets written (possibly an environment-variable
        /// form); <paramref name="serverPath"/> is the real file it must resolve to, and both
        /// halves of that are verified here — an entry that reads back as something else, or
        /// that no longer names the installed server once expanded, is never written.
        ///
        /// <paramref name="configExists"/> gates the one branch that does NOT splice: a whole
        /// new document may only be produced for a file that is genuinely absent. It is belt
        /// and braces behind the classification in <see cref="ReconcileCore"/> — the branch
        /// replaces the entire file, so it must never be reachable for a file that is on disk
        /// but merely read back empty.
        ///
        /// <paramref name="keepExistingEntryKeys"/> is true ONLY when the entry being rewritten
        /// is already ours. Extra keys on our own entry — an <c>env</c> variable someone added
        /// by hand — are then carried through, which is the promise made to a user maintaining
        /// their own registration. Keys on an entry that is NOT ours are dropped, because
        /// merging into one produces exactly the nonsense this used to write: a remote server's
        /// <c>url</c> left sitting beside a <c>command</c>, an entry that is neither shape.
        /// </summary>
        internal static bool TryBuildUpdatedConfig(
            string raw,
            bool configExists,
            Dictionary<string, object> root,
            string command,
            string serverPath,
            bool keepExistingEntryKeys,
            out string updated,
            out string error)
        {
            updated = null;
            error = null;

            var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };

            // Every other server survives verbatim.
            var servers = new Dictionary<string, object>(StringComparer.Ordinal);
            object existingServers;
            if (root != null && root.TryGetValue(McpConfigEditor.ServersProperty, out existingServers))
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
            if (keepExistingEntryKeys && servers.TryGetValue(ServerName, out ours))
                entry = ours as Dictionary<string, object>;
            entry = entry == null
                ? new Dictionary<string, object>(StringComparer.Ordinal)
                : new Dictionary<string, object>(entry, StringComparer.Ordinal);

            // A stdio server never has these. Removed unconditionally so an entry that once
            // carried them cannot come back out of here as a half-remote hybrid.
            entry.Remove("url");
            entry.Remove("httpUrl");
            entry.Remove("headers");

            entry["type"] = "stdio";
            entry[AddInServerContract.CommandProperty] = command;
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
                if (McpConfigEditor.ExistsButReadsEmpty(configExists, raw))
                {
                    error = "the configuration file is on disk but read back empty";
                    return false;
                }

                updated = "{\"" + McpConfigEditor.ServersProperty + "\":" + serversJson + "}";
                expectedTopLevelKeys = 1;
            }
            else
            {
                int valueStart, valueEnd;
                if (McpConfigEditor.TryFindTopLevelValueSpan(raw, McpConfigEditor.ServersProperty, out valueStart, out valueEnd))
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
                    string separator = McpConfigEditor.HasAnyContentAfterBrace(raw, brace) ? "," : "";
                    updated = raw.Substring(0, brace + 1)
                        + "\"" + McpConfigEditor.ServersProperty + "\":" + serversJson + separator
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

            string readBack = ReadRegisteredCommand(reparsed);
            if (!string.Equals(readBack, command, StringComparison.Ordinal))
            {
                error = "the updated configuration does not name the installed server";
                updated = null;
                return false;
            }
            if (!IsSamePath(McpConfigEditor.ExpandEnvironmentReferences(readBack, McpConfigEditor.ProcessEnvironmentLookup()), serverPath))
            {
                error = "the registered command does not resolve to the installed server";
                updated = null;
                return false;
            }

            // And nobody else's server was gained or lost on the way through. A multiset
            // comparison for the same reason the project-scope builder uses one: a check the
            // wrong file can satisfy is worse than no check at all.
            var expectedServers = McpConfigEditor.ListServerNames(raw);
            if (!expectedServers.Contains(ServerName))
                expectedServers.Add(ServerName);
            if (!McpConfigEditor.SameMultiset(expectedServers, McpConfigEditor.ListServerNames(updated)))
            {
                error = "the update would have changed which MCP servers are configured";
                updated = null;
                return false;
            }

            return true;
        }

        // ===== File I/O =====

        /// <summary>Attempts, and the pause between them, for <see cref="ReadSharedSettled"/>.</summary>
        private const int EmptyReadAttempts = 3;
        private const int EmptyReadPauseMs = 50;

        /// <summary>
        /// Reads a config file, waiting out the brief window in which it reads as nothing.
        ///
        /// The CLI truncates the file and then flushes its new content; the shared read below
        /// can land in between and come back with zero bytes. Every caller already refuses to
        /// write in that case, so this only turns a refusal into a normal reconcile — and it
        /// cannot mask a genuinely empty file, because that one still reads empty after the
        /// last attempt and the refusal stands either way. Bounded at
        /// <see cref="EmptyReadAttempts"/> attempts (~100 ms worst case, and only when the
        /// file really is reading empty) because this also runs on the settings dialog's
        /// thread.
        /// </summary>
        private static string ReadSharedSettled(string path)
        {
            string text = ReadShared(path);

            for (int attempt = 1; attempt < EmptyReadAttempts && text.Trim().Length == 0; attempt++)
            {
                System.Threading.Thread.Sleep(EmptyReadPauseMs);
                try
                {
                    text = ReadShared(path);
                }
                catch (Exception ex)
                {
                    // A retry that fails tells us nothing new: keep the empty read we already
                    // have and let the caller refuse, rather than turning a settled "cannot
                    // read this" into a different error.
                    System.Diagnostics.Debug.WriteLine("config re-read: " + ex.Message);
                    break;
                }
            }

            return text;
        }

        /// <summary>
        /// Reads a config the Claude Code CLI may be holding open — hence
        /// <see cref="FileShare.ReadWrite"/>, which <see cref="File.ReadAllText(string)"/>
        /// does not grant.
        /// </summary>
        private static string ReadShared(string path)
        {
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }

        /// <summary>
        /// Writes via a sibling temp file and <see cref="File.Replace(string,string,string)"/>,
        /// which swaps the file in one operation and leaves the previous content as a backup.
        /// A half-written config is therefore not reachable even if the machine dies mid-write.
        ///
        /// <paramref name="keepBackup"/> is true for <c>~/.claude.json</c>, where the backup is
        /// the user's only way back. It is false for a project's <c>.mcp.json</c>: that file
        /// lives in source control, which is a better undo than a stray sibling file would be —
        /// so the backup is deleted once the swap has succeeded, having already done its job of
        /// covering the swap itself.
        /// </summary>
        private static bool TryWriteAtomically(string targetPath, string content, bool keepBackup, out string error)
        {
            error = null;
            string dir = Path.GetDirectoryName(targetPath);
            string name = Path.GetFileName(targetPath);
            string temp = Path.Combine(dir, name + ".outlookai-new");
            string backup = Path.Combine(dir, name + ".outlookai-backup");

            try
            {
                // No BOM: the CLI reads these files as plain UTF-8.
                File.WriteAllText(temp, content, new UTF8Encoding(false));

                if (File.Exists(targetPath))
                    File.Replace(temp, targetPath, backup, true);
                else
                    File.Move(temp, targetPath);

                if (!keepBackup)
                {
                    try { if (File.Exists(backup)) File.Delete(backup); }
                    catch (Exception ex) { System.Diagnostics.Debug.WriteLine("backup cleanup: " + ex.Message); }
                }

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
                // Value names from the shared contract: outlook_health reads these four by name
                // (Status, ResolvedServerPath, Healed, LastReconcileUtc), and Detail/Command are
                // part of the same documented key.
                key.SetValue(AddInServerContract.McpStatusValueName, snap.Status ?? "", RegistryValueKind.String);
                key.SetValue(AddInServerContract.McpDetailValueName, snap.Detail ?? "", RegistryValueKind.String);
                key.SetValue(AddInServerContract.McpCommandValueName, snap.RegisteredCommand ?? "", RegistryValueKind.String);
                key.SetValue(AddInServerContract.McpResolvedServerPathValueName, snap.ResolvedServerPath ?? "", RegistryValueKind.String);
                key.SetValue(AddInServerContract.McpHealedValueName, snap.Healed ? 1 : 0, RegistryValueKind.DWord);
                key.SetValue(AddInServerContract.McpLastReconcileUtcValueName, snap.LastReconcileUtc ?? "", RegistryValueKind.String);
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

        private static void WriteDword(string valueName, int value)
        {
            using (var key = Registry.CurrentUser.CreateSubKey(McpKeyPath))
            {
                if (key == null)
                    return;
                key.SetValue(valueName, value, RegistryValueKind.DWord);
            }
        }
    }
}
