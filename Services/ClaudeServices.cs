using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAI.Services
{
    public class ClaudeService
    {
        private const string Model = "claude-opus-4-6";
        private static readonly string CliArgs = "-p - --output-format json --max-turns 1 --model \"" + Model + "\"";

        private static readonly object _warmLock = new object();
        private static Process _warmProcess;
        private static string _lastPrerequisiteError;

        public enum ActionType
        {
            Proofread, Revise, Draft, Shorten, Lengthen, Formal, Friendly, Custom
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

        public async Task<string> ProcessEmailAsync(ActionType action, string emailContent, string customPrompt = "")
        {
            // Check for known prerequisite issues
            if (_lastPrerequisiteError != null)
            {
                // Try once more in case the user fixed it
                WarmUpCore();
                if (_lastPrerequisiteError != null)
                    throw new Exception(_lastPrerequisiteError);
            }

            var userMessage = BuildUserMessage(action, emailContent, customPrompt);

            return await Task.Run(() => ExecutePrompt(userMessage));
        }

        private string ExecutePrompt(string userMessage)
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
                    FileName = "claude",
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
        private string ParseResult(string json)
        {
            // The --output-format json output has a "result" field with the text
            var resultMarker = "\"result\":\"";
            var startIndex = json.LastIndexOf(resultMarker);

            if (startIndex == -1)
            {
                // Fallback: try "text" field (older CLI versions)
                resultMarker = "\"text\":\"";
                startIndex = json.LastIndexOf(resultMarker);
            }

            if (startIndex == -1)
                throw new Exception("Could not parse Claude response. Raw output:\n" + Truncate(json, 500));

            startIndex += resultMarker.Length;
            var endIndex = startIndex;
            while (endIndex < json.Length)
            {
                endIndex = json.IndexOf("\"", endIndex);
                if (endIndex == -1) break;
                if (json[endIndex - 1] != '\\') break;
                endIndex++;
            }
            if (endIndex == -1)
                throw new Exception("Could not parse Claude response. Raw output:\n" + Truncate(json, 500));

            return UnescapeJson(json.Substring(startIndex, endIndex - startIndex));
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

        private string UnescapeJson(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text
                .Replace("\\n", "\n")
                .Replace("\\r", "\r")
                .Replace("\\t", "\t")
                .Replace("\\\"", "\"")
                .Replace("\\\\", "\\");
        }

        private static string Truncate(string text, int maxLength)
        {
            if (string.IsNullOrEmpty(text) || text.Length <= maxLength) return text;
            return text.Substring(0, maxLength) + "...";
        }

        private string GetSystemPrompt(ActionType action)
        {
            switch (action)
            {
                case ActionType.Proofread:
                    return "You are a professional editor. Review the email for grammar, spelling, punctuation, and clarity issues. Return the corrected email text only. Do not add any explanations.";
                case ActionType.Revise:
                    return "You are a professional writing assistant. Improve the email clarity, flow, and impact. Return only the revised email text without any explanations.";
                case ActionType.Draft:
                    return "You are a professional email writer. Write a clear, professional email based on the instructions. If replying to an email thread, write only your reply - do not include the previous messages. Return only the email text you are composing.";
                case ActionType.Shorten:
                    return "You are a professional editor. Condense this email to be more concise while keeping essential information. Return only the shortened email text.";
                case ActionType.Lengthen:
                    return "You are a professional writer. Expand this email with more detail while maintaining professionalism. Return only the expanded email text.";
                case ActionType.Formal:
                    return "You are a professional editor. Rewrite this email in a more formal tone suitable for business. Return only the rewritten email text.";
                case ActionType.Friendly:
                    return "You are a professional editor. Rewrite this email in a warmer, friendlier tone while remaining professional. Return only the rewritten email text.";
                case ActionType.Custom:
                    return "You are a professional email writing assistant. Follow the user's instructions exactly and apply them to the provided email content. Return only the modified email text without any explanations.";
                default:
                    return "You are a professional email writing assistant. Help the user with their email based on their instructions. Return only the result.";
            }
        }

        private string BuildUserMessage(ActionType action, string emailContent, string customPrompt)
        {
            var instructions = GetSystemPrompt(action);
            string task;

            if (action == ActionType.Draft)
            {
                if (!string.IsNullOrWhiteSpace(emailContent))
                {
                    task = "Write a reply email based on these instructions:\n\n" + customPrompt +
                           "\n\n--- Email thread for context (do NOT include this in your response, just use it for context) ---\n\n" + emailContent;
                }
                else
                {
                    task = "Write an email based on these instructions:\n\n" + customPrompt;
                }
            }
            else if (action == ActionType.Custom)
            {
                task = "Email content:\n\n" + emailContent + "\n\nInstructions: " + customPrompt;
            }
            else
            {
                task = "Email to " + action.ToString().ToLower() + ":\n\n" + emailContent;
            }

            return instructions + "\n\n" + task;
        }
    }
}
