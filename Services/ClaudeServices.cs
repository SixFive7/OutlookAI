using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace OutlookAI.Services
{
    public class EditTurn
    {
        public ClaudeService.ActionType Action { get; set; }
        public string Instruction { get; set; }
        public string SelectedText { get; set; }
        public string Result { get; set; }
    }

    public static class ClaudeService
    {
        private const string Model = "claude-opus-4-6";
        private static readonly string CliArgs = "-p - --output-format json --max-turns 1 --model \"" + Model + "\"";
        private static readonly string ClaudePath = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            ".local", "bin", "claude.exe");

        private static readonly object _warmLock = new object();
        private static Process _warmProcess;
        private static volatile string _lastPrerequisiteError;

        public enum ActionType
        {
            Proofread, Revise, Draft, Shorten, Lengthen, Formal, Friendly
        }

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
            lock (_warmLock)
            {
                KillWarmProcess();

                try
                {
                    var process = SpawnProcess();

                    // Check if the process exited immediately (missing prerequisites)
                    if (process.WaitForExit(1500) && process.HasExited)
                    {
                        var stderr = process.StandardError.ReadToEnd();
                        _lastPrerequisiteError = DiagnoseError(stderr, process.ExitCode);
                        return;
                    }

                    _warmProcess = process;
                    _lastPrerequisiteError = null;
                }
                catch (Exception ex)
                {
                    // SpawnProcess sets _lastPrerequisiteError for Win32Exception (CLI not found)
                    if (_lastPrerequisiteError == null)
                        _lastPrerequisiteError = "Failed to start Claude Code: " + ex.Message;
                }
            }
        }

        /// <summary>
        /// Processes an email action using iterative editing with full conversation history.
        /// </summary>
        public static async Task<string> ProcessEmailAsync(
            ActionType action, string customPrompt,
            List<EditTurn> editHistory,
            string draftText, string signatureText, string threadText,
            string selectedText = null)
        {
            // Check for known prerequisite issues
            if (_lastPrerequisiteError != null)
            {
                // Try once more in case the user fixed it
                WarmUpCore();
                if (_lastPrerequisiteError != null)
                    throw new Exception(_lastPrerequisiteError);
            }

            var prompt = BuildIterativePrompt(action, customPrompt, editHistory, draftText, signatureText, threadText, selectedText);

            return await Task.Run(() => ExecutePrompt(prompt));
        }

        private static string BuildIterativePrompt(
            ActionType action, string customPrompt,
            List<EditTurn> editHistory,
            string draftText, string signatureText, string threadText,
            string selectedText)
        {
            var sb = new StringBuilder();

            // System instruction
            sb.AppendLine("You are an email writing assistant integrated into Microsoft Outlook. Your output is inserted directly into the user's email draft.");
            sb.AppendLine();
            sb.AppendLine("Output format:");
            sb.AppendLine("- Return only the email draft text — no commentary, no explanations, no code fences, no HTML tags.");
            sb.AppendLine("- Use blank lines between paragraphs for clean, readable structure.");
            sb.AppendLine();
            sb.AppendLine("Content:");
            sb.AppendLine("- Write in the same language as the existing draft or email thread, unless the user asks otherwise.");
            sb.AppendLine("- Match the tone and formality of the conversation unless asked to change it.");
            if (!string.IsNullOrWhiteSpace(threadText))
                sb.AppendLine("- When replying, address the content of the quoted thread.");
            if (!string.IsNullOrWhiteSpace(signatureText))
                sb.AppendLine("- The email signature is added automatically — do not include any sign-off, closing, or name at the end.");
            if (!string.IsNullOrWhiteSpace(threadText))
                sb.AppendLine("- The quoted thread is preserved automatically — do not repeat or include it.");
            sb.AppendLine();

            // Edit history from previous turns
            if (editHistory.Count > 0)
            {
                sb.AppendLine("## Edit History");
                sb.AppendLine();
                for (int i = 0; i < editHistory.Count; i++)
                {
                    var turn = editHistory[i];
                    string label = GetActionLabel(turn.Action, turn.Instruction);
                    if (!string.IsNullOrEmpty(turn.SelectedText))
                        label += " (applied to selection)";
                    sb.AppendLine($"### Turn {i + 1} — {label}");
                    sb.AppendLine("Result:");
                    sb.AppendLine("\"\"\"");
                    sb.AppendLine(turn.Result);
                    sb.AppendLine("\"\"\"");
                    sb.AppendLine();
                }
            }

            // Current draft
            sb.AppendLine("## Current Draft");
            sb.AppendLine();
            if (string.IsNullOrWhiteSpace(draftText))
            {
                sb.AppendLine("(empty — this is a new email, compose from scratch)");
            }
            else
            {
                sb.AppendLine("\"\"\"");
                sb.AppendLine(draftText);
                sb.AppendLine("\"\"\"");
            }
            sb.AppendLine();

            // Signature context
            if (!string.IsNullOrWhiteSpace(signatureText))
            {
                sb.AppendLine("## Signature (for context — do NOT include in your response)");
                sb.AppendLine();
                sb.AppendLine("\"\"\"");
                sb.AppendLine(signatureText);
                sb.AppendLine("\"\"\"");
                sb.AppendLine();
            }

            // Thread context
            if (!string.IsNullOrWhiteSpace(threadText))
            {
                sb.AppendLine("## Quoted Thread (for context — do NOT include in your response)");
                sb.AppendLine();
                sb.AppendLine("\"\"\"");
                sb.AppendLine(threadText);
                sb.AppendLine("\"\"\"");
                sb.AppendLine();
            }

            // Current request
            sb.AppendLine("## Current Request");
            sb.AppendLine();
            string currentLabel = GetActionLabel(action, customPrompt);
            sb.AppendLine($"Action: {currentLabel}");
            sb.AppendLine();

            // Selection constraint
            if (!string.IsNullOrWhiteSpace(selectedText))
            {
                sb.AppendLine("The user has selected the following text in the draft. Modify ONLY that portion — keep all other text exactly as-is:");
                sb.AppendLine("\"\"\"");
                sb.AppendLine(selectedText);
                sb.AppendLine("\"\"\"");
                sb.AppendLine();
            }

            sb.AppendLine("Write the updated draft text only.");

            return sb.ToString();
        }

        private static string GetActionLabel(ActionType action, string instruction)
        {
            switch (action)
            {
                case ActionType.Proofread:
                    return "Proofread: Fix any spelling, grammar, and punctuation errors. Keep the tone, meaning, and structure unchanged.";
                case ActionType.Revise:
                    return "Revise: Improve clarity, flow, and word choice. Preserve the original meaning and tone.";
                case ActionType.Shorten:
                    return "Shorten: Make the email more concise. Remove filler and redundancy while keeping all key points.";
                case ActionType.Lengthen:
                    return "Lengthen: Expand the email with more detail, context, or explanation. Keep the same tone and intent.";
                case ActionType.Formal:
                    return "Formal: Rewrite in a more formal, professional tone. Keep the same content and meaning.";
                case ActionType.Friendly:
                    return "Friendly: Rewrite in a warmer, more conversational tone. Keep the same content and meaning.";
                case ActionType.Draft:
                    return string.IsNullOrWhiteSpace(instruction)
                        ? "Draft"
                        : "Draft: \"" + instruction + "\"";
                default:
                    return action.ToString();
            }
        }

        private static string ExecutePrompt(string userMessage)
        {
            Process process = null;

            lock (_warmLock)
            {
                if (_warmProcess != null && !_warmProcess.HasExited)
                {
                    process = _warmProcess;
                    _warmProcess = null;
                }
            }

            if (process == null)
            {
                process = SpawnProcess();
            }

            try
            {
                var stdoutBuilder = new StringBuilder();
                var stderrBuilder = new StringBuilder();

                process.OutputDataReceived += (s, e) =>
                {
                    if (e.Data != null) stdoutBuilder.AppendLine(e.Data);
                };
                process.ErrorDataReceived += (s, e) =>
                {
                    if (e.Data != null) stderrBuilder.AppendLine(e.Data);
                };

                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                // Write the user message to stdin, then close to signal EOF
                process.StandardInput.Write(userMessage);
                process.StandardInput.Close();

                bool exited = process.WaitForExit(120_000);
                if (!exited)
                {
                    try { process.Kill(); } catch { }
                    // Pre-warm the next process for the next attempt
                    WarmUp();
                    throw new Exception("Request timed out after 2 minutes. Please try again.");
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

                // Pre-warm the next process in the background
                WarmUp();

                return result;
            }
            finally
            {
                process.Dispose();
            }
        }

        private static Process SpawnProcess()
        {
            try
            {
                var process = new Process();
                process.StartInfo = new ProcessStartInfo
                {
                    FileName = ClaudePath,
                    Arguments = CliArgs,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardInput = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8,
                };

                process.Start();
                return process;
            }
            catch (Win32Exception)
            {
                _lastPrerequisiteError =
                    "Claude Code CLI is not installed or not on PATH.\n\n" +
                    "Install it by running:\n" +
                    "  npm install -g @anthropic-ai/claude-code\n\n" +
                    "Then authenticate by running:\n" +
                    "  claude auth login";
                throw new Exception(_lastPrerequisiteError);
            }
        }

        /// <summary>
        /// Parses the "result" field from the Claude CLI JSON output.
        /// </summary>
        private static string ParseResult(string json)
        {
            var serializer = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var parsed = serializer.Deserialize<Dictionary<string, object>>(json);

            if (parsed.TryGetValue("result", out var result) && result is string resultStr)
                return resultStr;

            // Fallback: try "text" field (older CLI versions)
            if (parsed.TryGetValue("text", out var text) && text is string textStr)
                return textStr;

            throw new Exception("Could not parse Claude response. Raw output:\n" + Truncate(json, 500));
        }

        private static string DiagnoseError(string stderr, int exitCode)
        {
            if (string.IsNullOrWhiteSpace(stderr))
                return $"Claude Code exited with code {exitCode}.";

            var lower = stderr.ToLowerInvariant();

            if (lower.Contains("node") && (lower.Contains("not recognized") || lower.Contains("not found")))
                return "Node.js is required for Claude Code CLI but was not found.\n\n" +
                       "Install Node.js from https://nodejs.org";

            if (lower.Contains("not recognized") || lower.Contains("not found") || lower.Contains("no such file"))
                return "Claude Code CLI is not installed or not on PATH.\n\n" +
                       "Install it by running:\n" +
                       "  npm install -g @anthropic-ai/claude-code\n\n" +
                       "Then authenticate by running:\n" +
                       "  claude auth login";

            if (lower.Contains("auth") || lower.Contains("login") || lower.Contains("unauthorized")
                || lower.Contains("not logged in") || lower.Contains("token"))
                return "Claude Code is not authenticated.\n\n" +
                       "Run this command in a terminal:\n" +
                       "  claude auth login\n\n" +
                       "Then sign in with your Claude subscription.";

            if (lower.Contains("rate limit") || lower.Contains("too many"))
                return "Rate limit reached. Please wait a moment and try again.";

            if (lower.Contains("overloaded") || lower.Contains("capacity"))
                return "Claude is currently overloaded. Please try again in a moment.";

            return "Claude Code error: " + Truncate(stderr.Trim(), 300);
        }

        private static bool IsPrerequisiteError(string stderr)
        {
            if (string.IsNullOrWhiteSpace(stderr)) return false;
            var lower = stderr.ToLowerInvariant();
            return lower.Contains("not recognized") || lower.Contains("not found")
                || lower.Contains("auth") || lower.Contains("login")
                || lower.Contains("unauthorized") || lower.Contains("not logged in")
                || lower.Contains("no such file");
        }

        /// <summary>
        /// Kills any pre-warmed process. Called at add-in shutdown.
        /// </summary>
        public static void Shutdown()
        {
            lock (_warmLock)
            {
                KillWarmProcess();
            }
        }

        private static void KillWarmProcess()
        {
            if (_warmProcess != null)
            {
                try
                {
                    if (!_warmProcess.HasExited)
                        _warmProcess.Kill();
                }
                catch { }
                _warmProcess.Dispose();
                _warmProcess = null;
            }
        }

        private static string Truncate(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLength) return text;
            return text.Substring(0, maxLength) + "...";
        }
    }
}
