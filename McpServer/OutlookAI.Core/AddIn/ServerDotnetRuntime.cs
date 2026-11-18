using System;
using System.IO;
using Microsoft.Win32;

namespace OutlookAI.Core.AddIn
{
    /// <summary>
    /// THE MAIL SERVER'S .NET RUNTIME (Q172, decided by the maintainer 2026-11-18): which one the server
    /// runs on, where OutlookAI keeps its own copy of it, and whether this machine has a runtime the
    /// server can start on. The one copy of those answers (Q154): setup's helper
    /// (<c>OutlookAI.RuntimeSetup</c>) installs into <see cref="DefaultRoot"/> and checks its work with
    /// <see cref="IsComplete"/>, and the add-in asks <see cref="IsAvailable"/> before it registers the
    /// server with Claude Code.
    /// <para>
    /// <b>Where the server finds it.</b> The server and its COM host are framework-dependent .NET 11
    /// executables, and their apphost looks for a runtime in two places, in order
    /// (<c>AppHostDotNetSearch</c> in their project files): <c>..\..\dotnet</c> from the executable's
    /// own folder - <c>%LOCALAPPDATA%\OutlookAI\Setup\McpServer</c> once installed, so this folder - and
    /// then the machine-wide .NET, where a developer's SDK or a user's own install of .NET 11 sits.
    /// Setup puts Microsoft's own runtime package for <see cref="Version"/> in the first, per user and
    /// with no administrator rights, so an update installed silently gets it too.
    /// </para>
    /// <para>
    /// The add-in reaches Core only through this namespace (Q164 (a)): nothing here touches Outlook,
    /// COM or the index - a registry read and a few file probes.
    /// </para>
    /// </summary>
    public static class ServerDotnetRuntime
    {
        /// <summary>
        /// The Microsoft.NETCore.App version the server and its COM host ask for - their runtimeconfig
        /// names 11.0.0 alone, which any 11.x satisfies - and the version of the package setup
        /// installs. Its major is the server's target framework's: Tools/Checks/check-pinned-constants.ps1
        /// holds the two together.
        /// </summary>
        public const string Version = "11.0.0";

        /// <summary>
        /// The private copy's folder, in OutlookAI's per-user folder beside <c>Setup</c>. The server's
        /// project bakes <c>..\..\</c> and this name into its apphost (<c>AppHostRelativeDotNet</c>),
        /// and setup's uninstaller deletes it.
        /// </summary>
        public const string FolderName = "dotnet";

        /// <summary>The major version, "11", that a machine-wide runtime must have.</summary>
        public static string Major => Version.Substring(0, Version.IndexOf('.'));

        /// <summary>%LOCALAPPDATA%\OutlookAI\dotnet, where setup installs the private copy.</summary>
        public static string DefaultRoot()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookAI", FolderName);
        }

        /// <summary>
        /// Whether <paramref name="root"/> holds a runtime the apphost can start the server on: the host
        /// resolver of <see cref="Version"/>, which the apphost loads from <c>host\fxr</c>, and the shared
        /// framework of the same version, which that resolver then loads from <c>shared</c>.
        /// </summary>
        public static bool IsComplete(string? root)
        {
            if (string.IsNullOrEmpty(root))
                return false;

            try
            {
                return File.Exists(Path.Combine(root, "host", "fxr", Version, "hostfxr.dll"))
                    && File.Exists(Path.Combine(root, "shared", "Microsoft.NETCore.App", Version, "System.Private.CoreLib.dll"));
            }
            catch (ArgumentException)
            {
                return false;
            }
        }

        /// <summary>
        /// Whether the server can start on this machine: the private copy is complete, or a
        /// machine-wide Microsoft.NETCore.App of the same major is installed, which is where the
        /// apphost looks next.
        /// </summary>
        public static bool IsAvailable()
        {
            return IsComplete(DefaultRoot()) || IsMachineWideInstalled();
        }

        /// <summary>
        /// True when the machine-wide .NET holds some Microsoft.NETCore.App of <see cref="Major"/>:
        /// the server's roll-forward (Minor) takes a newer 11.x but never a 12.x or a 10.x. Probes the
        /// folders, because the <c>sharedfx</c> registry key is absent on machines that do have the
        /// runtime; the <c>sharedhost</c> key says where they are, and the 64-bit Program Files
        /// otherwise.
        /// </summary>
        public static bool IsMachineWideInstalled()
        {
            try
            {
                // The 64-bit registry view and the 64-bit Program Files explicitly: Outlook may be a
                // 32-bit process, and under WOW64 the default view would redirect into Wow6432Node and
                // to the 32-bit Program Files - neither of which is where the x64 runtime the server
                // needs lives.
                string? root = null;
                using (RegistryKey hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                using (RegistryKey? key = hklm.OpenSubKey(@"SOFTWARE\dotnet\Setup\InstalledVersions\x64\sharedhost"))
                {
                    if (key?.GetValue("Path") is string path && path.Length > 0)
                        root = Path.Combine(path, "shared", "Microsoft.NETCore.App");
                }

                if (root == null)
                {
                    string? programFiles = Environment.GetEnvironmentVariable("ProgramW6432");
                    if (string.IsNullOrEmpty(programFiles))
                        programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                    root = Path.Combine(programFiles, "dotnet", "shared", "Microsoft.NETCore.App");
                }

                if (!Directory.Exists(root))
                    return false;

                string prefix = Major + ".";
                foreach (string directory in Directory.GetDirectories(root))
                {
                    if (Path.GetFileName(directory).StartsWith(prefix, StringComparison.Ordinal))
                        return true;
                }
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException || ex is System.Security.SecurityException || ex is ArgumentException)
            {
                System.Diagnostics.Debug.WriteLine("machine-wide runtime probe: " + ex.Message);
            }

            return false;
        }
    }
}
