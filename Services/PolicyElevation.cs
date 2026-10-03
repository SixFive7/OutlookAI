using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Principal;
using Microsoft.Win32;

namespace OutlookAI.Services
{
    /// <summary>How an attempt to apply the policy values through UAC ended.</summary>
    internal enum PolicyElevationResult
    {
        /// <summary>The helper wrote every value and read it back.</summary>
        Applied,

        /// <summary>Nothing needed an administrator, so nothing was asked.</summary>
        NothingToDo,

        /// <summary>The user answered No, or closed the UAC prompt. Not an error: nothing changed.</summary>
        Cancelled,

        /// <summary>The helper ran and refused, before writing anything.</summary>
        Refused,

        /// <summary>Anything else: no helper, Windows would not start it, it failed or timed out.</summary>
        Failed,
    }

    /// <summary>The result, and one sentence for the settings dialog.</summary>
    internal sealed class PolicyElevationOutcome
    {
        internal PolicyElevationOutcome(PolicyElevationResult result, string message, int exitCode)
        {
            Result = result;
            Message = message;
            ExitCode = exitCode;
        }

        internal PolicyElevationResult Result { get; }

        internal string Message { get; }

        /// <summary>The helper's exit code when it ran, -1 when it did not.</summary>
        internal int ExitCode { get; }
    }

    /// <summary>
    /// APPLIES THE POLICY VALUES THE RECONCILE COULD NOT WRITE, THROUGH UAC - the maintainer's Q128:
    /// "If the change requires admin and the user is not admin generate a uac prompt."
    ///
    /// <para>
    /// It starts <c>OutlookAI.PolicyWriter.exe</c> with ShellExecuteEx and the "runas" verb, which
    /// is what puts the UAC prompt up: a consent prompt for an administrator whose token is
    /// filtered, a credentials prompt for a standard user. The request names the user by SID,
    /// because under an administrator's credentials the helper runs AS THAT ADMINISTRATOR and its
    /// HKCU is the administrator's hive - see <see cref="PolicyWriterCommandLine"/> and the helper's
    /// own checks. The add-in validates the request exactly as the helper will
    /// (<see cref="PolicyWriterCommandLine.Build"/>), so it never puts up a prompt for a request the
    /// helper would refuse.
    /// </para>
    ///
    /// <para>
    /// BLOCKS until the prompt is answered and the helper has exited - so the settings dialog calls
    /// it off Outlook's UI thread. Never throws. A cancelled prompt (ERROR_CANCELLED, 1223) is
    /// <see cref="PolicyElevationResult.Cancelled"/>, reported as what it is, not as a failure.
    /// </para>
    /// </summary>
    internal static class PolicyElevation
    {
        /// <summary>The helper writes a handful of values; a minute is a hang, not a slow disk.</summary>
        private const int HelperTimeoutMs = 60000;

        private const int ErrorCancelled = 1223;

        internal static PolicyElevationOutcome ApplyPending(IntPtr ownerWindow)
        {
            try
            {
                List<string> notOffered;
                List<KeyValuePair<string, int>> values = OutlookTuningService.PolicyValuesNeedingAdministrator(out notOffered);
                string skipped = notOffered.Count == 0
                    ? string.Empty
                    : " Not applied, because the desired number is not one OutlookAI offers: " + string.Join(", ", notOffered) + ".";
                if (values.Count == 0)
                {
                    return new PolicyElevationOutcome(PolicyElevationResult.NothingToDo,
                        "Nothing needs an administrator: every managed policy value is already in effect." + skipped, -1);
                }

                string helper = ResolveHelperPath();
                if (helper.Length == 0)
                {
                    return new PolicyElevationOutcome(PolicyElevationResult.Failed,
                        PolicyWriterCommandLine.ExecutableName + " was not found next to the add-in. Reinstall OutlookAI.", -1);
                }

                string sid;
                using (WindowsIdentity me = WindowsIdentity.GetCurrent())
                    sid = me.User == null ? string.Empty : me.User.Value;

                string arguments;
                try
                {
                    arguments = PolicyWriterCommandLine.Build(sid, OutlookTuningService.PolicyOfficeVersion, values);
                }
                catch (ArgumentException ex)
                {
                    return new PolicyElevationOutcome(PolicyElevationResult.Failed, "OutlookAI cannot make this request: " + ex.Message, -1);
                }

                PolicyElevationOutcome outcome = RunElevated(helper, arguments, ownerWindow);
                if (outcome.Result == PolicyElevationResult.Applied)
                {
                    // Written under a running Outlook, which reads them at its next start - and the
                    // reconcile records them as applied, which empties NeedsAdministrator.
                    OutlookTuningService.MarkRestartNeeded();
                    OutlookTuningService.ReconcileFromUi();
                }
                return skipped.Length == 0
                    ? outcome
                    : new PolicyElevationOutcome(outcome.Result, outcome.Message + skipped, outcome.ExitCode);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Policy elevation: " + ex.Message);
                return new PolicyElevationOutcome(PolicyElevationResult.Failed, "Applying as administrator failed: " + ex.Message, -1);
            }
        }

        /// <summary>
        /// The helper beside the running add-in, else beside the installed one. The add-in's own
        /// folder first, so a dev build in DevBuilds uses its own helper; the installer's InstallDir
        /// second, for a runtime that loads the assembly from a shadow copy.
        /// </summary>
        internal static string ResolveHelperPath()
        {
            var candidates = new List<string>();
            try
            {
                var assembly = typeof(PolicyElevation).Assembly;
                Uri codeBase;
                if (Uri.TryCreate(assembly.CodeBase, UriKind.Absolute, out codeBase) && codeBase.IsFile)
                    candidates.Add(Path.GetDirectoryName(codeBase.LocalPath));
                if (!string.IsNullOrEmpty(assembly.Location))
                    candidates.Add(Path.GetDirectoryName(assembly.Location));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Policy helper path (assembly): " + ex.Message);
            }

            try
            {
                using (RegistryKey app = Registry.CurrentUser.OpenSubKey(McpRegistrationService.AppKeyPath))
                {
                    string installDir = app?.GetValue(McpRegistrationService.InstallDirValueName) as string;
                    if (!string.IsNullOrEmpty(installDir))
                        candidates.Add(installDir);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Policy helper path (InstallDir): " + ex.Message);
            }

            foreach (string dir in candidates)
            {
                if (string.IsNullOrEmpty(dir))
                    continue;
                string path = Path.Combine(dir, PolicyWriterCommandLine.ExecutableName);
                if (File.Exists(path))
                    return Path.GetFullPath(path);
            }
            return string.Empty;
        }

        private static PolicyElevationOutcome RunElevated(string helper, string arguments, IntPtr ownerWindow)
        {
            var info = new SHELLEXECUTEINFO
            {
                cbSize = Marshal.SizeOf(typeof(SHELLEXECUTEINFO)),
                // A process handle to wait on and read the exit code from; no error UI of the
                // shell's own (this class reports every outcome itself); and synchronous.
                fMask = SEE_MASK_NOCLOSEPROCESS | SEE_MASK_FLAG_NO_UI | SEE_MASK_NOASYNC,
                // The settings window owns the prompt, so it is not left behind it.
                hwnd = ownerWindow,
                lpVerb = "runas",
                lpFile = helper,
                lpParameters = arguments,
                // System32, not the user-writable install folder, as the elevated process's
                // working directory.
                lpDirectory = Environment.SystemDirectory,
                nShow = SW_HIDE,
            };

            if (!ShellExecuteEx(ref info))
            {
                int error = Marshal.GetLastWin32Error();
                if (error == ErrorCancelled)
                {
                    return new PolicyElevationOutcome(PolicyElevationResult.Cancelled,
                        "Cancelled at the administrator prompt. Nothing was changed.", -1);
                }
                return new PolicyElevationOutcome(PolicyElevationResult.Failed,
                    "Windows did not start the administrator helper: " + new Win32Exception(error).Message + " (" + error + ").", -1);
            }

            if (info.hProcess == IntPtr.Zero)
            {
                return new PolicyElevationOutcome(PolicyElevationResult.Failed,
                    "Windows started the administrator helper but gave no handle to wait on, so its result is unknown.", -1);
            }

            try
            {
                uint wait = WaitForSingleObject(info.hProcess, HelperTimeoutMs);
                if (wait != WaitObject0)
                {
                    return new PolicyElevationOutcome(PolicyElevationResult.Failed,
                        "The administrator helper did not finish within " + (HelperTimeoutMs / 1000) + " seconds.", -1);
                }

                uint exitCode;
                if (!GetExitCodeProcess(info.hProcess, out exitCode))
                {
                    return new PolicyElevationOutcome(PolicyElevationResult.Failed,
                        "The administrator helper finished, but its result could not be read.", -1);
                }

                int code = unchecked((int)exitCode);
                return new PolicyElevationOutcome(Classify(code), Describe(code), code);
            }
            finally
            {
                CloseHandle(info.hProcess);
            }
        }

        private static PolicyElevationResult Classify(int exitCode)
        {
            switch (exitCode)
            {
                case PolicyWriterExit.Written:
                    return PolicyElevationResult.Applied;
                case PolicyWriterExit.RefusedArguments:
                case PolicyWriterExit.RefusedUser:
                case PolicyWriterExit.RefusedHiveNotLoaded:
                    return PolicyElevationResult.Refused;
                default:
                    return PolicyElevationResult.Failed;
            }
        }

        private static string Describe(int exitCode)
        {
            string what = PolicyWriterExit.Describe(exitCode);
            string sentence = char.ToUpperInvariant(what[0]) + what.Substring(1) + ".";
            return exitCode == PolicyWriterExit.Written
                ? sentence + " Restart Outlook for them to take effect."
                : sentence;
        }

        // ===== Win32 =====

        private const uint SEE_MASK_NOCLOSEPROCESS = 0x00000040;
        private const uint SEE_MASK_NOASYNC = 0x00000100;
        private const uint SEE_MASK_FLAG_NO_UI = 0x00000400;
        private const int SW_HIDE = 0;
        private const uint WaitObject0 = 0;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct SHELLEXECUTEINFO
        {
            public int cbSize;
            public uint fMask;
            public IntPtr hwnd;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpVerb;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpFile;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpParameters;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpDirectory;
            public int nShow;
            public IntPtr hInstApp;
            public IntPtr lpIDList;
            [MarshalAs(UnmanagedType.LPWStr)] public string lpClass;
            public IntPtr hkeyClass;
            public uint dwHotKey;
            public IntPtr hIconOrMonitor;
            public IntPtr hProcess;
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode, SetLastError = true, EntryPoint = "ShellExecuteExW")]
        private static extern bool ShellExecuteEx(ref SHELLEXECUTEINFO info);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern uint WaitForSingleObject(IntPtr handle, int milliseconds);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GetExitCodeProcess(IntPtr process, out uint exitCode);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr handle);
    }
}
