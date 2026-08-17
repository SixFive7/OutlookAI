using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace OutlookAI.Services
{
    /// <summary>
    /// One completed edit in this compose session: what was asked for, and what came back.
    ///
    /// <see cref="Label"/> is the RENDERED instruction text as it stood when the turn ran, not an
    /// identity a label is derived from later. That is the point of it. Prompts are editable now,
    /// so re-deriving a label at send time would let one edit in settings retroactively rewrite
    /// what every past turn claims to have asked for - and the model is told this history is what
    /// actually happened.
    /// </summary>
    public class EditTurn
    {
        /// <summary>Instruction text this turn ran with, verbatim, exactly as rendered then.</summary>
        public string Label { get; set; }

        /// <summary>The selection it was applied to, or null when it applied to the whole draft.</summary>
        public string SelectedText { get; set; }

        /// <summary>What the model returned, which is what the draft was set to.</summary>
        public string Result { get; set; }
    }

    public static class ClaudeService
    {
        /// <summary>
        /// Where both machine settings live: HKCU\Software\OutlookAI, values <c>Model</c>
        /// (string) and <c>RequestTimeoutSeconds</c> (DWORD).
        ///
        /// <c>Model</c> is now also what the settings window writes, and ABSENT IS THE DEFAULT:
        /// no value means no <c>--model</c> argument, which is the documented way to let Claude
        /// Code decide (its own <c>model</c> setting, then ANTHROPIC_MODEL, then the account
        /// default). The timeout stays deliberately registry-only and undocumented in the UI -
        /// it exists so a genuinely slow link can be worked around on the spot, not as a feature.
        /// </summary>
        private const string AppKeyPath = @"Software\OutlookAI";
        private const string ModelValueName = "Model";
        private const string RequestTimeoutValueName = "RequestTimeoutSeconds";

        /// <summary>
        /// Hard cap on one CLI request. One number, and every sentence that mentions it is
        /// built from that number - the literal and the words "2 minutes" used to sit six lines
        /// apart as two spellings of the same thing.
        /// </summary>
        private const int DefaultRequestTimeoutSeconds = 120;

        /// <summary>Bounds on the override, so a typo cannot disable the cap or make it useless.</summary>
        private const int MinRequestTimeoutSeconds = 10;
        private const int MaxRequestTimeoutSeconds = 3600;

        /// <summary>How long to wait for a warm-up process to prove a prerequisite is missing by exiting.</summary>
        private const int WarmUpProbeMs = 1500;

        /// <summary>How long <c>taskkill /T /F</c> gets before we fall back to Process.Kill.</summary>
        private const int TaskKillWaitMs = 5000;

        /// <summary>
        /// How much raw CLI output an error message may quote. Bigger than
        /// <see cref="ErrorTextExcerptChars"/> because this is unparsed JSON, where the useful
        /// part is rarely at the front.
        /// </summary>
        private const int RawOutputExcerptChars = 500;

        /// <summary>How much CLI stderr, or a CLI-reported failure, an error message may quote.</summary>
        private const int ErrorTextExcerptChars = 300;

        /// <summary>
        /// How to install the CLI. Written once: two different error messages tell the user
        /// this, and a second copy of an install command is a second thing to get wrong.
        /// </summary>
        private const string InstallInstructions =
            "Install it by running this in PowerShell:\n" +
            "  irm https://claude.ai/install.ps1 | iex\n\n" +
            "Then sign in by running:\n" +
            "  claude\n\n" +
            "Then restart Outlook.";

        /// <summary>
        /// Head of the label for the three free-text buttons. Not a section: it is the one word
        /// that says which of the two shapes the request has (a bare "Draft", or a quoted
        /// instruction), and it names nothing a user would want to rewrite.
        /// </summary>
        private const string DraftLabel = "Draft";

        /// <summary>
        /// Where the Claude Code CLI is expected. Public within the assembly because
        /// <see cref="McpRegistrationService"/> asks the same question ("is Claude Code
        /// installed?") and used to spell the same path a second way - so a move would fix the
        /// executing path, which fails loudly, and leave the detecting one quietly wrong.
        /// </summary>
        internal static readonly string ClaudePath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".local", "bin", "claude.exe");

        /// <summary>
        /// Every request, minus the model. The model is APPENDED ONLY WHEN THE USER CHOSE ONE:
        /// see <see cref="BuildCliArgs"/>. There is no shipped model id anywhere in this file
        /// any more, and that is the point - a pinned id is a dated asset that outlives its
        /// model, and leaving the argument out is what makes "absent" mean "whatever Claude Code
        /// is configured to use" rather than "whatever we last shipped".
        /// </summary>
        private const string BaseCliArgs = "-p - --output-format json --max-turns 1";

        private static readonly object _argsLock = new object();
        private static string _cliArgs;

        /// <summary>
        /// The exact argument string a request is started with right now. Cached rather than
        /// rebuilt per request, and invalidated by <see cref="SetConfiguredModel"/>, so a model
        /// chosen in the settings window takes effect on the next request instead of the next
        /// Outlook session.
        ///
        /// Internal because the settings window shows what will actually be sent, and the
        /// offline harness proves it: a second place that spells these arguments out would be a
        /// second place to get them wrong.
        /// </summary>
        internal static string CliArguments
        {
            get
            {
                lock (_argsLock)
                {
                    if (_cliArgs == null)
                        _cliArgs = BuildCliArgs();
                    return _cliArgs;
                }
            }
        }

        private static string BuildCliArgs()
        {
            string model = ConfiguredModel;
            return model == null ? BaseCliArgs : BaseCliArgs + " --model \"" + model + "\"";
        }

        /// <summary>
        /// The model the user chose, or NULL for "let Claude Code decide" - which is the default
        /// and is stored as nothing at all. A stored value that could not be a model id is
        /// treated as absent rather than passed on, because it would land inside a quoted
        /// <c>--model "..."</c> argument (see <see cref="ModelCatalog.IsWellFormedModelId"/>).
        /// </summary>
        internal static string ConfiguredModel
        {
            get
            {
                string configured = ReadAppString(ModelValueName);
                return ModelCatalog.IsWellFormedModelId(configured) ? configured.Trim() : null;
            }
        }

        /// <summary>
        /// Stores the chosen model, or clears it when <paramref name="model"/> is null or blank -
        /// absent means "Claude Code decides". Returns whether the choice landed; a value that
        /// could not be a model id is refused rather than written.
        ///
        /// It also drops any pre-warmed process, and that is not optional: a warm process is
        /// already running with the arguments it was given, so re-reading the registry alone
        /// would leave the very next request answering on the model the user just changed away
        /// from - the one failure this whole setting exists to let them fix.
        /// </summary>
        internal static bool SetConfiguredModel(string model)
        {
            string trimmed = model == null ? string.Empty : model.Trim();
            bool clearing = trimmed.Length == 0;
            if (!clearing && !ModelCatalog.IsWellFormedModelId(trimmed))
                return false;

            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(AppKeyPath))
                {
                    if (key == null)
                        return false;
                    if (clearing)
                        key.DeleteValue(ModelValueName, false);
                    else
                        key.SetValue(ModelValueName, trimmed, Microsoft.Win32.RegistryValueKind.String);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Model preference: " + ex.Message);
                return false;
            }

            RefreshCliArgs();
            return true;
        }

        /// <summary>
        /// Rebuilds the argument string after the stored model changed, and - only when it really
        /// did change - throws away the pre-warmed process and starts a fresh one. Also forgets
        /// which model warning has already been shown: a different model is a different
        /// situation, and the user is owed the warning again if the new choice also fails.
        /// </summary>
        private static void RefreshCliArgs()
        {
            string rebuilt = BuildCliArgs();
            bool changed;
            lock (_argsLock)
            {
                changed = !string.Equals(_cliArgs, rebuilt, StringComparison.Ordinal);
                _cliArgs = rebuilt;
            }
            if (!changed)
                return;

            lock (_noticeLock)
            {
                _shownModelWarning = null;
                _pendingModelNotice = null;
            }
            lock (_warmLock)
            {
                KillWarmProcess();
            }
            WarmUp();
        }

        /// <summary>
        /// Seconds one request may take before the process tree is killed: the machine's
        /// override, clamped, or <see cref="DefaultRequestTimeoutSeconds"/>. The timeout and the
        /// message the user sees are both derived from this, so they cannot disagree.
        /// </summary>
        internal static int RequestTimeoutSeconds
        {
            get
            {
                int? configured = ReadAppDword(RequestTimeoutValueName);
                if (!configured.HasValue)
                    return DefaultRequestTimeoutSeconds;
                if (configured.Value < MinRequestTimeoutSeconds) return MinRequestTimeoutSeconds;
                if (configured.Value > MaxRequestTimeoutSeconds) return MaxRequestTimeoutSeconds;
                return configured.Value;
            }
        }

        /// <summary>
        /// The timeout in the words the message uses - "2 minutes", "90 seconds". Whole minutes
        /// are spelled as minutes because that is how a user reads a wait they are sitting through.
        /// </summary>
        private static string DescribeRequestTimeout(int seconds)
        {
            if (seconds >= 60 && seconds % 60 == 0)
            {
                int minutes = seconds / 60;
                return minutes == 1 ? "1 minute" : minutes + " minutes";
            }
            return seconds == 1 ? "1 second" : seconds + " seconds";
        }

        private static string ReadAppString(string valueName)
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(AppKeyPath))
                    return key?.GetValue(valueName) as string;
            }
            catch { return null; }
        }

        private static int? ReadAppDword(string valueName)
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(AppKeyPath))
                    return key?.GetValue(valueName) as int?;
            }
            catch { return null; }
        }

        // ===== The silent fallback =====
        //
        // A model the CLI cannot serve does not fail. It prints a warning on stderr, runs the
        // request on its own default instead, and exits 0 - so the user gets a perfectly good
        // reply and never learns it came from a different model than the one they picked. That
        // is the one failure this feature can produce that looks exactly like success, so the
        // warning is kept and shown rather than discarded with the rest of a successful run's
        // stderr. It is NOT an error: the request worked, and the pane says so.

        private static readonly object _noticeLock = new object();

        /// <summary>Waiting to be shown once, by whoever asks next.</summary>
        private static string _pendingModelNotice;

        /// <summary>The raw warning already surfaced this session, so it is not repeated.</summary>
        private static string _shownModelWarning;

        /// <summary>
        /// Records a model warning printed by a run that SUCCEEDED. Shown once per distinct
        /// warning per session: a bad pin produces the identical warning on every single
        /// request, and a status line that says the same thing forever is one the user stops
        /// reading. A DIFFERENT warning is a new thing to say and is shown.
        /// </summary>
        private static void NoteModelWarning(string stderr)
        {
            string warning = FindModelWarning(stderr);
            if (warning == null)
                return;
            lock (_noticeLock)
            {
                if (string.Equals(warning, _shownModelWarning, StringComparison.Ordinal))
                    return;
                _shownModelWarning = warning;
                _pendingModelNotice = DescribeModelWarning(warning);
            }
        }

        /// <summary>
        /// The pending model notice, and clears it - so the pane shows it once rather than
        /// decorating every later reply with a warning about a request that already happened.
        /// Null when there is nothing to say, which is the normal case.
        /// </summary>
        internal static string TakeModelNotice()
        {
            lock (_noticeLock)
            {
                string notice = _pendingModelNotice;
                _pendingModelNotice = null;
                return notice;
            }
        }

        /// <summary>
        /// The first stderr line that mentions the model. Matching on that word alone is
        /// deliberate: the CLI's exact wording is not documented, and a precise pattern would
        /// silently stop matching the day it is reworded - which is the same silence this exists
        /// to break. A successful run normally prints nothing at all on stderr, so a false
        /// positive costs one status line, once.
        /// </summary>
        private static string FindModelWarning(string stderr)
        {
            if (string.IsNullOrWhiteSpace(stderr))
                return null;
            foreach (string line in stderr.Split('\n'))
            {
                string trimmed = line.Trim();
                if (trimmed.Length == 0)
                    continue;
                if (trimmed.IndexOf("model", StringComparison.OrdinalIgnoreCase) < 0)
                    continue;
                return Truncate(trimmed, ErrorTextExcerptChars);
            }
            return null;
        }

        /// <summary>
        /// The warning in the user's terms. Careful about what it asserts: we know the CLI said
        /// something about the model and that the request nonetheless succeeded. We do not know
        /// which model actually answered, so it says "may have" rather than claiming a fallback
        /// we cannot see.
        /// </summary>
        private static string DescribeModelWarning(string warning)
        {
            string chosen = ConfiguredModel;
            string opening = chosen == null
                ? "Claude Code reported a problem with the model it was set to use"
                : "Claude Code reported a problem with the model you chose (" + chosen + ")";
            return opening + ", so this reply may have come from a different one. It said: " + warning;
        }

        private static readonly object _warmLock = new object();
        private static Process _warmProcess;
        private static StringBuilder _warmStdout;
        private static StringBuilder _warmStderr;
        private static volatile string _lastPrerequisiteError;
        private static int _warmingUp;            // single-flight guard for WarmUpCore (0/1)
        private static volatile bool _shutdown;   // set at Shutdown; stops a late warm-up orphaning a process

        /// <summary>
        /// Pre-warms a Claude CLI process so it's ready when the user clicks an action.
        /// Called at add-in startup and after each completed request.
        /// </summary>
        public static void WarmUp()
        {
            Task.Run(() =>
            {
                try
                {
                    WarmUpCore();
                }
                catch
                {
                    // Warm-up failure is not fatal; error will surface on first use
                }
            });
        }

        private static void WarmUpCore()
        {
            // Single-flight: only one warm-up runs at a time. Concurrent or redundant WarmUp()
            // calls (e.g. several compose windows finishing at once) skip instead of killing and
            // respawning each other's process.
            if (Interlocked.CompareExchange(ref _warmingUp, 1, 0) != 0)
                return;
            try
            {
                lock (_warmLock)
                {
                    if (_shutdown) return;
                    // A live pre-warmed process is already ready — nothing to do.
                    if (_warmProcess != null && !_warmProcess.HasExited) return;
                }

                StringBuilder stdoutBuilder, stderrBuilder;
                Process process;
                try
                {
                    process = SpawnProcess(out stdoutBuilder, out stderrBuilder);
                }
                catch (Exception ex)
                {
                    // SpawnProcess sets _lastPrerequisiteError for Win32Exception (CLI not found)
                    if (_lastPrerequisiteError == null)
                        _lastPrerequisiteError = "Failed to start Claude Code: " + ex.Message;
                    return;
                }

                // Probe for immediate exit (missing prerequisites) OUTSIDE the lock so the
                // WarmUpProbeMs wait never blocks Shutdown() (UI thread) or ExecutePrompt().
                if (process.WaitForExit(WarmUpProbeMs) && process.HasExited)
                {
                    process.WaitForExit(); // flush async output handlers
                    var stderr = stderrBuilder.ToString();
                    _lastPrerequisiteError = DiagnoseError(stderr, process.ExitCode);
                    process.Dispose();
                    return;
                }

                // NOTE: a process waiting on stdin does NOT prove prerequisites are OK — the CLI
                // validates auth only after reading the prompt. So we do not clear
                // _lastPrerequisiteError here; it is cleared only after a genuinely successful run.
                // Publish under the lock; discard if we've since shut down or another warm exists.
                lock (_warmLock)
                {
                    if (_shutdown || (_warmProcess != null && !_warmProcess.HasExited))
                    {
                        KillProcessTree(process);
                        process.Dispose();
                        return;
                    }
                    KillWarmProcess(); // clear any dead leftover
                    _warmProcess = process;
                    _warmStdout = stdoutBuilder;
                    _warmStderr = stderrBuilder;
                }
            }
            finally
            {
                Interlocked.Exchange(ref _warmingUp, 0);
            }
        }

        /// <summary>
        /// Processes an email action using iterative editing with full conversation history.
        ///
        /// <paramref name="actionLabel"/> is the fully resolved instruction, not the name of an
        /// action: for a quick button it is that button's stored prompt, and for the three
        /// free-text buttons it is <see cref="BuildDraftLabel"/> over what the user typed. The
        /// caller resolves it because the caller is also what records it in the edit history, and
        /// the two have to be the same string for the history to stay honest.
        ///
        /// A blank label is refused rather than sent. It would produce a request naming no action
        /// at all, and the model would then do something arbitrary to a real draft.
        /// </summary>
        public static async Task<string> ProcessEmailAsync(
            string actionLabel, List<EditTurn> editHistory,
            string draftText, string signatureText, string threadText,
            string selectedText = null)
        {
            var prereqError = _lastPrerequisiteError;
            if (prereqError != null)
                throw new Exception(prereqError);

            if (string.IsNullOrWhiteSpace(actionLabel))
                throw new Exception("This action has no instruction text, so there is nothing to ask for. Check the button's prompt in OutlookAI Settings.");

            var prompt = BuildIterativePrompt(actionLabel, editHistory, draftText, signatureText, threadText, selectedText);

            return await Task.Run(() => ExecutePrompt(prompt));
        }

        /// <summary>
        /// The instruction label for the three free-text buttons: a bare "Draft" when the box is
        /// empty, otherwise Draft: "what the user typed". The quotes are what keep a typed
        /// instruction visibly separate from the prompt text around it.
        ///
        /// Public because the pane needs the same string twice - once to send and once to record
        /// in the edit history - and both copies must be identical.
        /// </summary>
        public static string BuildDraftLabel(string instruction)
        {
            string typed = instruction == null ? string.Empty : instruction.Trim();
            return typed.Length == 0
                ? DraftLabel
                : DraftLabel + ": \"" + typed + "\"";
        }

        /// <summary>
        /// Picks the best installed signature for the current draft (D38 "Select the
        /// best signature" pane action). Same safety contract as the writing actions:
        /// draft, thread, recipients, and signature excerpts are fenced untrusted
        /// content; the model must answer with EXACTLY one signature name (plain text,
        /// nothing else). The caller matches the answer against the installed names.
        /// </summary>
        internal static async Task<string> SelectSignatureAsync(
            List<SignatureStore.SignatureOption> signatures,
            string draftText, string threadText, string recipientsText)
        {
            var prereqError = _lastPrerequisiteError;
            if (prereqError != null)
                throw new Exception(prereqError);

            var prompt = BuildSignatureSelectionPrompt(signatures, draftText, threadText, recipientsText);

            return await Task.Run(() => ExecutePrompt(prompt));
        }

        private static string BuildSignatureSelectionPrompt(
            List<SignatureStore.SignatureOption> signatures,
            string draftText, string threadText, string recipientsText)
        {
            var sb = new StringBuilder();

            // The editable section holds the whole instruction half INCLUDING its closing line,
            // because that is what a user editing it reads as one prompt. In the assembled prompt
            // that line belongs after the data instead: it is the output contract, and last is
            // where a one-line contract survives a long signature list and a quoted thread. So the
            // trailing copy is stripped out of the section here and the shipped line is re-emitted
            // at the end - never both, or the model is told twice in two different places.
            string instructions = WithoutTrailingClosingLine(
                PromptStore.GetSection(PromptSection.SignatureSelection));
            if (instructions.Length > 0)
            {
                sb.AppendLine(instructions);
                sb.AppendLine();
            }

            sb.AppendLine("## Available Signatures (name + excerpt)");
            sb.AppendLine();
            var listing = new StringBuilder();
            foreach (var signature in signatures)
            {
                listing.AppendLine("Name: " + signature.Name);
                if (!string.IsNullOrWhiteSpace(signature.Excerpt))
                    listing.AppendLine("Excerpt: " + signature.Excerpt);
                listing.AppendLine();
            }
            Fence(sb, listing.ToString().TrimEnd());
            sb.AppendLine();

            if (!string.IsNullOrWhiteSpace(recipientsText))
            {
                sb.AppendLine("## Recipients");
                sb.AppendLine();
                Fence(sb, recipientsText);
                sb.AppendLine();
            }

            sb.AppendLine("## Current Draft");
            sb.AppendLine();
            if (string.IsNullOrWhiteSpace(draftText))
                sb.AppendLine("(empty - nothing typed yet)");
            else
                Fence(sb, draftText);
            sb.AppendLine();

            if (!string.IsNullOrWhiteSpace(threadText))
            {
                sb.AppendLine("## Quoted Thread");
                sb.AppendLine();
                Fence(sb, threadText);
                sb.AppendLine();
            }

            sb.AppendLine(PromptDefaults.SignatureSelectionClosingLine);

            return sb.ToString();
        }

        /// <summary>
        /// One editable section, trimmed and followed by a newline; reports whether anything was
        /// written. Two rules live here rather than at each call site:
        ///
        ///  - A section the user has cleared writes nothing at all, not a blank line. An empty
        ///    override means "drop this block", and a stray blank line in the middle of the rules
        ///    is the visible residue of a block that was supposed to be gone.
        ///  - The ends are trimmed, because a multiline text box hands back the trailing newline
        ///    the user never sees, and two of those in a row would separate blocks the assembled
        ///    prompt means to keep adjacent.
        /// </summary>
        private static bool AppendSection(StringBuilder sb, PromptSection section)
        {
            string text = PromptStore.GetSection(section);
            if (string.IsNullOrWhiteSpace(text))
                return false;
            sb.AppendLine(text.Trim());
            return true;
        }

        /// <summary>
        /// The signature-selection instructions with a trailing copy of the closing line removed,
        /// so re-emitting that line after the data cannot leave a duplicate stranded above it.
        /// Case-insensitive: a user who retypes the line may well shift its capitalisation, and a
        /// contradictory-looking duplicate costs more than a strip that was not needed.
        /// </summary>
        private static string WithoutTrailingClosingLine(string text)
        {
            string trimmed = text == null ? string.Empty : text.Trim();
            string closing = PromptDefaults.SignatureSelectionClosingLine;
            if (trimmed.EndsWith(closing, StringComparison.OrdinalIgnoreCase))
                trimmed = trimmed.Substring(0, trimmed.Length - closing.Length).TrimEnd();
            return trimmed;
        }

        private static void Fence(StringBuilder sb, string content)
        {
            var fence = "---CONTENT-" + Guid.NewGuid().ToString("N").Substring(0, 12) + "---";
            sb.AppendLine(fence);
            sb.AppendLine(content);
            sb.AppendLine(fence);
        }

        private static string BuildIterativePrompt(
            string actionLabel,
            List<EditTurn> editHistory,
            string draftText, string signatureText, string threadText,
            string selectedText)
        {
            var sb = new StringBuilder();

            // The rules, from the editable sections: the preamble always, then each conditional
            // block only when the thing it talks about is actually in this prompt. Order note:
            // the hard-coded version interleaved these lines - reply, then signature, then
            // thread-is-preserved - because they were three separate AppendLine calls with two
            // separate conditions. As editable blocks the two reply lines have to travel together,
            // so the signature line now follows both of them instead of sitting between them.
            // Same rules under the same conditions; only their order on the page moved.
            bool wroteRules = AppendSection(sb, PromptSection.Preamble);
            if (!string.IsNullOrWhiteSpace(threadText))
                wroteRules |= AppendSection(sb, PromptSection.ReplyRules);
            if (!string.IsNullOrWhiteSpace(signatureText))
                wroteRules |= AppendSection(sb, PromptSection.SignatureRule);
            if (wroteRules)
                sb.AppendLine();

            // Edit history from previous turns
            if (editHistory != null && editHistory.Count > 0)
            {
                sb.AppendLine("## Edit History");
                sb.AppendLine();
                for (int i = 0; i < editHistory.Count; i++)
                {
                    var turn = editHistory[i];
                    // Verbatim, never re-derived: the label is what this turn asked for at the
                    // time, and a prompt edited since then must not rewrite the past. Trimmed only
                    // at the ends, so padding a stored prompt happens to carry cannot run the
                    // heading onto a second line.
                    string label = turn.Label == null ? string.Empty : turn.Label.Trim();
                    if (!string.IsNullOrEmpty(turn.SelectedText))
                        label += " (applied to selection)";
                    sb.AppendLine($"### Turn {i + 1} - {label}");
                    sb.AppendLine("Result:");
                    Fence(sb, turn.Result);
                    sb.AppendLine();
                }
            }

            // Current draft
            sb.AppendLine("## Current Draft");
            sb.AppendLine();
            if (string.IsNullOrWhiteSpace(draftText))
            {
                sb.AppendLine("(empty - this is a new email, compose from scratch)");
            }
            else
            {
                Fence(sb, draftText);
            }
            sb.AppendLine();

            // Signature context
            if (!string.IsNullOrWhiteSpace(signatureText))
            {
                sb.AppendLine("## Signature (for context - do NOT include in your response)");
                sb.AppendLine();
                Fence(sb, signatureText);
                sb.AppendLine();
            }

            // Thread context
            if (!string.IsNullOrWhiteSpace(threadText))
            {
                sb.AppendLine("## Quoted Thread (for context - do NOT include in your response)");
                sb.AppendLine();
                Fence(sb, threadText);
                sb.AppendLine();
            }

            // Current request
            sb.AppendLine("## Current Request");
            sb.AppendLine();
            sb.AppendLine($"Action: {actionLabel.Trim()}");
            sb.AppendLine();

            // Selection constraint
            if (!string.IsNullOrWhiteSpace(selectedText))
            {
                sb.AppendLine("The user has selected the following text in the draft. Modify ONLY that portion - keep all other text exactly as-is:");
                Fence(sb, selectedText);
                sb.AppendLine();
            }

            sb.AppendLine("Write the updated draft text only.");

            return sb.ToString();
        }

        private static string ExecutePrompt(string userMessage)
        {
            Process process = null;
            StringBuilder stdoutBuilder = null;
            StringBuilder stderrBuilder = null;

            bool usedWarm = false;
            lock (_warmLock)
            {
                if (_warmProcess != null && !_warmProcess.HasExited)
                {
                    process = _warmProcess;
                    stdoutBuilder = _warmStdout;
                    stderrBuilder = _warmStderr;
                    _warmProcess = null;
                    _warmStdout = null;
                    _warmStderr = null;
                    usedWarm = true;
                }
            }

            if (process == null)
            {
                process = SpawnProcess(out stdoutBuilder, out stderrBuilder);
            }

            try
            {
                // Write the user message to stdin as UTF-8, then close to signal EOF.
                // If a pre-warmed process died between the liveness check and now, discard it
                // and spawn a fresh one so the user gets a result instead of a raw pipe error.
                try
                {
                    using (var writer = new StreamWriter(process.StandardInput.BaseStream, new UTF8Encoding(false)))
                        writer.Write(userMessage);
                }
                catch (Exception) when (usedWarm)
                {
                    try { process.Dispose(); } catch { }
                    process = SpawnProcess(out stdoutBuilder, out stderrBuilder);
                    usedWarm = false;
                    using (var writer = new StreamWriter(process.StandardInput.BaseStream, new UTF8Encoding(false)))
                        writer.Write(userMessage);
                }

                int timeoutSeconds = RequestTimeoutSeconds;
                bool exited = process.WaitForExit(timeoutSeconds * 1000);
                if (!exited)
                {
                    KillProcessTree(process);
                    // Pre-warm the next process for the next attempt
                    WarmUp();
                    throw new Exception(
                        "Request timed out after " + DescribeRequestTimeout(timeoutSeconds) +
                        ". Please try again.");
                }

                // Flush async output buffers
                process.WaitForExit();

                if (process.ExitCode != 0)
                {
                    var stderr = stderrBuilder.ToString();
                    var error = DiagnoseError(stderr, process.ExitCode);
                    _lastPrerequisiteError = IsPrerequisiteError(stderr) ? error : null;
                    // Pre-warm for next attempt
                    WarmUp();
                    throw new Exception(error);
                }

                var stdout = stdoutBuilder.ToString().Trim();
                var result = ParseResult(stdout);

                // The run succeeded, so nothing here is an error - but a warning about the model
                // printed on the way is the one thing a successful run can be hiding. Keep it.
                NoteModelWarning(stderrBuilder.ToString());

                // A successful run is positive proof prerequisites are OK; clear any stale gate.
                _lastPrerequisiteError = null;

                // Pre-warm the next process in the background
                WarmUp();

                return result;
            }
            finally
            {
                process.Dispose();
            }
        }

        private static Process SpawnProcess(out StringBuilder stdoutBuilder, out StringBuilder stderrBuilder)
        {
            try
            {
                stdoutBuilder = new StringBuilder();
                stderrBuilder = new StringBuilder();

                var process = new Process();
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = ClaudePath,
                    Arguments = CliArguments,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8,
                };

                // Thread-safe by design: BeginOutputReadLine/BeginErrorReadLine serialize
                // their callbacks (single producer per builder), and WaitForExit() provides
                // the memory barrier before any read.
                var stdout = stdoutBuilder;
                var stderr = stderrBuilder;
                process.OutputDataReceived += (s, e) =>
                {
                    if (e.Data != null) stdout.AppendLine(e.Data);
                };
                process.ErrorDataReceived += (s, e) =>
                {
                    if (e.Data != null) stderr.AppendLine(e.Data);
                };

                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                return process;
            }
            catch (Win32Exception)
            {
                _lastPrerequisiteError =
                    "Claude Code CLI was not found at " + ClaudePath + ".\n\n" + InstallInstructions;
                throw new Exception(_lastPrerequisiteError);
            }
        }

        /// <summary>
        /// Parses the "result" field from the Claude CLI JSON output.
        /// </summary>
        private static string ParseResult(string json)
        {
            if (string.IsNullOrWhiteSpace(json))
                throw new Exception("Claude returned an empty response.");

            Dictionary<string, object> parsed;
            try
            {
                var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
                parsed = serializer.Deserialize<Dictionary<string, object>>(json);
            }
            catch (Exception ex)
            {
                throw new Exception("Could not parse Claude response: " + ex.Message + "\nRaw output:\n" + Truncate(json, RawOutputExcerptChars));
            }

            if (parsed == null)
                throw new Exception("Could not parse Claude response. Raw output:\n" + Truncate(json, RawOutputExcerptChars));

            // The CLI exits 0 but sets is_error=true for failures it handled itself
            // (e.g. subtype "error_max_turns"); without this check the error envelope's
            // text would be written into the user's email as if it were the drafted reply.
            if (parsed.TryGetValue("is_error", out var isErrorObj) && isErrorObj is bool isError && isError)
            {
                string subtype = parsed.TryGetValue("subtype", out var st) ? st as string : null;
                string detail = parsed.TryGetValue("result", out var er) ? er as string : null;
                throw new Exception(DescribeCliError(subtype, detail));
            }

            if (parsed.TryGetValue("result", out var result) && result is string resultStr)
                return resultStr;

            if (parsed.TryGetValue("text", out var text) && text is string textStr)
                return textStr;

            throw new Exception("Could not parse Claude response. Raw output:\n" + Truncate(json, RawOutputExcerptChars));
        }

        private static string DescribeCliError(string subtype, string detail)
        {
            if (subtype == "error_max_turns")
                return "Claude stopped before finishing (turn limit reached). Please simplify or rephrase the request and try again.";
            if (!string.IsNullOrWhiteSpace(detail))
                return "Claude could not complete the request: " + Truncate(detail.Trim(), ErrorTextExcerptChars);
            if (!string.IsNullOrWhiteSpace(subtype))
                return "Claude could not complete the request (" + subtype + ").";
            return "Claude could not complete the request.";
        }

        private static string DiagnoseError(string stderr, int exitCode)
        {
            if (string.IsNullOrWhiteSpace(stderr))
                return $"Claude Code exited with code {exitCode}.";

            var lower = stderr.ToLowerInvariant();

            if (lower.Contains("node") && (lower.Contains("not recognized") || lower.Contains("not found")))
                return "Node.js is required for Claude Code CLI but was not found.\n\n" +
                       "Install Node.js from https://nodejs.org\n\n" +
                       "Then restart Outlook.";

            if (lower.Contains("not recognized") || lower.Contains("not found") || lower.Contains("no such file"))
                return "Claude Code CLI is not installed or not on PATH.\n\n" + InstallInstructions;

            if (lower.Contains("unauthorized") || lower.Contains("not logged in")
                || lower.Contains("auth required") || lower.Contains("not authenticated")
                || lower.Contains("please login") || lower.Contains("authentication failed")
                || lower.Contains("api key") || lower.Contains("invalid key"))
                return "Claude Code is not authenticated.\n\n" +
                       "Run this command in a terminal:\n" +
                       "  claude\n\n" +
                       "Then sign in with your Claude subscription in the browser and restart Outlook.";

            if (lower.Contains("rate limit") || lower.Contains("too many"))
                return "Rate limit reached. Please wait a moment and try again.";

            if (lower.Contains("overloaded") || lower.Contains("capacity"))
                return "Claude is currently overloaded. Please try again in a moment.";

            return "Claude Code error: " + Truncate(stderr.Trim(), ErrorTextExcerptChars);
        }

        private static bool IsPrerequisiteError(string stderr)
        {
            if (string.IsNullOrWhiteSpace(stderr)) return false;
            var lower = stderr.ToLowerInvariant();
            return lower.Contains("not recognized") || lower.Contains("not found")
                || lower.Contains("unauthorized") || lower.Contains("not logged in")
                || lower.Contains("not authenticated") || lower.Contains("auth required")
                || lower.Contains("please login") || lower.Contains("authentication failed")
                || lower.Contains("api key") || lower.Contains("invalid key")
                || lower.Contains("no such file");
        }

        /// <summary>
        /// Kills any pre-warmed process. Called at add-in shutdown.
        /// </summary>
        public static void Shutdown()
        {
            _shutdown = true;
            lock (_warmLock)
            {
                KillWarmProcess();
            }
        }

        private static void KillProcessTree(Process proc)
        {
            try
            {
                if (proc.HasExited) return;
                int pid = proc.Id;
                try
                {
                    bool killed;
                    using (var tk = Process.Start(new ProcessStartInfo
                    {
                        FileName = "taskkill",
                        Arguments = $"/T /F /PID {pid}",
                        UseShellExecute = false,
                        CreateNoWindow = true
                    }))
                    {
                        killed = tk != null && tk.WaitForExit(TaskKillWaitMs) && tk.ExitCode == 0;
                    }
                    if (!killed && !proc.HasExited)
                        proc.Kill();
                }
                catch
                {
                    try { proc.Kill(); } catch { }
                }
            }
            catch { }
        }

        private static void KillWarmProcess()
        {
            if (_warmProcess != null)
            {
                KillProcessTree(_warmProcess);
                _warmProcess.Dispose();
                _warmProcess = null;
                _warmStdout = null;
                _warmStderr = null;
            }
        }

        private static string Truncate(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLength) return text;
            return text.Substring(0, maxLength) + "...";
        }
    }
}
