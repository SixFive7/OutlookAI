using System.Diagnostics;
using System.Text;

namespace OutlookAI.McpServer.Tests.T2;

/// <summary>
/// A stand-in for the Claude Code CLI, put where the add-in runs it, for the live tests that drive the
/// task pane (Q168 (b), decided 2026-11-18): the pane's AI requests reach it instead of Claude, it answers
/// what the test told it to - a signature's name to a signature request, the body to any other - and it
/// keeps every prompt it was sent, for the test to read. The product has no automation entry point; this
/// is the one thing the tests put on the machine.
/// <para>
/// Only on a test machine, and only for the test that installs it: <see cref="LiveTestSettings.Load"/>
/// refuses a Production profile and a run without the per-run opt-in before anything is written, a
/// claude.exe that this class did not make refuses the install - a real CLI is never replaced - and
/// <see cref="Dispose"/> ends the stand-in's processes and removes every file it wrote. It is compiled on
/// the machine, from <see cref="Source"/>, by the .NET Framework's own C# compiler, which needs no SDK.
/// </para>
/// </summary>
internal sealed class ClaudeCliStandIn : IDisposable
{
    /// <summary>The file beside the stand-in that says this class made it.</summary>
    private const string MarkerFileName = "outlookai-claude-standin.marker";

    private const string SignatureAnswerFileName = "outlookai-claude-standin-signature.txt";
    private const string BodyAnswerFileName = "outlookai-claude-standin-body.txt";
    private const string PromptLogFileName = "outlookai-claude-standin-prompts.log";

    /// <summary>The line the stand-in writes before every prompt it logs.</summary>
    private const string PromptSeparator = "===== OUTLOOKAI CLAUDE STAND-IN PROMPT =====";

    /// <summary>
    /// The stand-in, in C# 5 for the .NET Framework's compiler: the prompt from stdin, the answer from the
    /// file the test wrote, and the CLI's JSON shape on stdout - what <c>ClaudeService.ParseResult</c> reads.
    /// </summary>
    private const string Source = """
        using System;
        using System.IO;
        using System.Text;

        internal static class Program
        {
            private static int Main()
            {
                string dir = AppDomain.CurrentDomain.BaseDirectory;
                string prompt;
                using (StreamReader reader = new StreamReader(Console.OpenStandardInput(), Encoding.UTF8))
                {
                    prompt = reader.ReadToEnd();
                }

                bool signature = prompt.IndexOf("## Available Signatures", StringComparison.Ordinal) >= 0;
                string answer = File.ReadAllText(Path.Combine(dir, signature ? "%SIGNATURE%" : "%BODY%"), Encoding.UTF8);
                File.AppendAllText(Path.Combine(dir, "%LOG%"), "%SEPARATOR%\r\n" + prompt + "\r\n", Encoding.UTF8);

                StreamWriter stdout = new StreamWriter(Console.OpenStandardOutput(), new UTF8Encoding(false));
                stdout.Write("{\"type\":\"result\",\"subtype\":\"success\",\"is_error\":false,\"result\":\"" + Escape(answer) + "\"}");
                stdout.Flush();
                return 0;
            }

            private static string Escape(string text)
            {
                StringBuilder escaped = new StringBuilder(text.Length + 16);
                foreach (char c in text)
                {
                    switch (c)
                    {
                        case '"': escaped.Append("\\\""); break;
                        case '\\': escaped.Append("\\\\"); break;
                        case '\n': escaped.Append("\\n"); break;
                        case '\r': escaped.Append("\\r"); break;
                        case '\t': escaped.Append("\\t"); break;
                        default:
                            if (c < ' ')
                            {
                                escaped.Append("\\u").Append(((int)c).ToString("x4"));
                            }
                            else
                            {
                                escaped.Append(c);
                            }

                            break;
                    }
                }

                return escaped.ToString();
            }
        }
        """;

    private readonly string _directory;
    private readonly bool _createdDirectory;

    private ClaudeCliStandIn(string directory, bool createdDirectory)
    {
        _directory = directory;
        _createdDirectory = createdDirectory;
    }

    /// <summary>Where the add-in runs the CLI from: <c>ClaudeService.ClaudePath</c>.</summary>
    public static string CliPath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".local", "bin", "claude.exe");

    /// <summary>
    /// Compiles the stand-in to <see cref="CliPath"/> and tells it its answers. Throws, having written
    /// nothing, on a machine that is not a test machine or where a real CLI is installed.
    /// </summary>
    public static ClaudeCliStandIn Install(string signatureAnswer, string bodyAnswer)
    {
        LiveTestSettings.Load();

        string directory = Path.GetDirectoryName(CliPath)!;
        if (File.Exists(CliPath) && !File.Exists(Path.Combine(directory, MarkerFileName)))
        {
            throw new InvalidOperationException(
                "A Claude Code CLI is installed at " + CliPath + ", and the stand-in never replaces one. "
                + "The pane tests run only on a test machine without Claude Code.");
        }

        bool created = !Directory.Exists(directory);
        Directory.CreateDirectory(directory);
        ClaudeCliStandIn standIn = new ClaudeCliStandIn(directory, created);
        try
        {
            File.WriteAllText(Path.Combine(directory, MarkerFileName), "made by OutlookAI's live tests; removed when they end", new UTF8Encoding(false));
            File.WriteAllText(Path.Combine(directory, SignatureAnswerFileName), signatureAnswer, new UTF8Encoding(false));
            File.WriteAllText(Path.Combine(directory, BodyAnswerFileName), bodyAnswer, new UTF8Encoding(false));
            File.Delete(Path.Combine(directory, PromptLogFileName));

            string source = Source
                .Replace("%SIGNATURE%", SignatureAnswerFileName, StringComparison.Ordinal)
                .Replace("%BODY%", BodyAnswerFileName, StringComparison.Ordinal)
                .Replace("%LOG%", PromptLogFileName, StringComparison.Ordinal)
                .Replace("%SEPARATOR%", PromptSeparator, StringComparison.Ordinal);
            string sourcePath = Path.Combine(directory, "outlookai-claude-standin.cs");
            File.WriteAllText(sourcePath, source, new UTF8Encoding(false));
            Compile(sourcePath, CliPath);
            File.Delete(sourcePath);
            return standIn;
        }
        catch
        {
            standIn.Dispose();
            throw;
        }
    }

    /// <summary>Every prompt the stand-in was sent, oldest first.</summary>
    public IReadOnlyList<string> Prompts()
    {
        string log = Path.Combine(_directory, PromptLogFileName);
        if (!File.Exists(log))
        {
            return Array.Empty<string>();
        }

        string text;
        using (FileStream stream = new FileStream(log, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
        using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
        {
            text = reader.ReadToEnd();
        }

        return text.Split(PromptSeparator, StringSplitOptions.RemoveEmptyEntries)
            .Select(p => p.Trim())
            .Where(p => p.Length > 0)
            .ToList();
    }

    /// <summary>
    /// Ends every stand-in process the add-in started - it keeps one waiting on stdin - and removes what
    /// <see cref="Install"/> wrote, the folder too when it made it.
    /// </summary>
    public void Dispose()
    {
        for (int attempt = 0; attempt < 10; attempt++)
        {
            foreach (Process process in Process.GetProcessesByName("claude"))
            {
                using (process)
                {
                    try
                    {
                        if (string.Equals(process.MainModule?.FileName, CliPath, StringComparison.OrdinalIgnoreCase))
                        {
                            process.Kill();
                            process.WaitForExit(5000);
                        }
                    }
                    catch (Exception ex) when (ex is InvalidOperationException or System.ComponentModel.Win32Exception)
                    {
                    }
                }
            }

            try
            {
                foreach (string name in new[] { "claude.exe", SignatureAnswerFileName, BodyAnswerFileName, PromptLogFileName, "outlookai-claude-standin.cs", MarkerFileName })
                {
                    File.Delete(Path.Combine(_directory, name));
                }

                if (_createdDirectory && Directory.Exists(_directory) && !Directory.EnumerateFileSystemEntries(_directory).Any())
                {
                    Directory.Delete(_directory);
                }

                return;
            }
            catch (Exception ex) when ((ex is IOException or UnauthorizedAccessException) && attempt < 9)
            {
                Thread.Sleep(500);
            }
        }
    }

    private static void Compile(string sourcePath, string exePath)
    {
        string csc = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Microsoft.NET", "Framework64", "v4.0.30319", "csc.exe");
        if (!File.Exists(csc))
        {
            throw new InvalidOperationException("The .NET Framework's C# compiler is not at " + csc + ".");
        }

        ProcessStartInfo start = new ProcessStartInfo(csc)
        {
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
        };
        start.ArgumentList.Add("/nologo");
        start.ArgumentList.Add("/target:exe");
        start.ArgumentList.Add("/out:" + exePath);
        start.ArgumentList.Add(sourcePath);
        using Process process = Process.Start(start)!;
        string output = process.StandardOutput.ReadToEnd() + process.StandardError.ReadToEnd();
        if (!process.WaitForExit(60000) || process.ExitCode != 0 || !File.Exists(exePath))
        {
            throw new InvalidOperationException("The Claude CLI stand-in did not compile: " + output);
        }
    }
}
