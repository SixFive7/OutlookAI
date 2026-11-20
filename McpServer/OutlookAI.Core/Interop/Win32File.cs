using System;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace OutlookAI.Core.Interop
{
    /// <summary>
    /// WHAT WINDOWS SAYS ABOUT A FILE OR A FOLDER, ASKED THROUGH A HANDLE: opening one for no access at all,
    /// where it really is once every link is followed, and what the file system records of it - and
    /// <c>CreateFileW</c>, which opens it. One source that Core and <c>Testbed/host/New-LiveTestSettings.ps1</c>
    /// both compile (Q171 (c), decided by the maintainer 2026-11-18): until then the attachment writer's
    /// folder check, the audit log's identity of its file and the settings script's check of where its output
    /// really goes each declared these calls themselves, and the script took the <c>\\?\</c> prefix off again
    /// in PowerShell.
    ///
    /// <para>
    /// C# 5, AND NOTHING ELSE OF CORE. The settings script compiles this file with Add-Type, in one go with
    /// a class of its own, and on Windows PowerShell 5.1 - the shell its self-test runs under on the build VM -
    /// Add-Type is the .NET Framework's C# 5 compiler: no <c>=&gt;</c> members, no <c>$""</c>, no
    /// <c>nameof</c>, no <c>?.</c>, no <c>?</c> on a reference type, no <c>out var</c>, and no type of
    /// Core's. Not for the add-in, which reaches Core only through <c>OutlookAI.Core.AddIn</c>
    /// (<c>Tools/Checks/check-pinned-constants.ps1</c>, check 15).
    /// </para>
    /// </summary>
    public static class Win32File
    {
        /// <summary>FILE_SHARE_READ, FILE_SHARE_WRITE and FILE_SHARE_DELETE: nobody else is kept out.</summary>
        public const uint ShareAll = 0x00000007;

        /// <summary>OPEN_EXISTING.</summary>
        public const uint OpenExisting = 3;

        /// <summary>FILE_FLAG_BACKUP_SEMANTICS, without which <c>CreateFileW</c> opens no folder.</summary>
        public const uint BackupSemantics = 0x02000000;

        /// <summary>BY_HANDLE_FILE_INFORMATION: what the file system records of an open file or folder.</summary>
        [StructLayout(LayoutKind.Sequential)]
        public struct ByHandleFileInformation
        {
            public uint FileAttributes;
            public uint CreationTimeLow;
            public uint CreationTimeHigh;
            public uint LastAccessTimeLow;
            public uint LastAccessTimeHigh;
            public uint LastWriteTimeLow;
            public uint LastWriteTimeHigh;
            public uint VolumeSerialNumber;
            public uint FileSizeHigh;
            public uint FileSizeLow;
            public uint NumberOfLinks;
            public uint FileIndexHigh;
            public uint FileIndexLow;
        }

        /// <summary>
        /// <paramref name="path"/>, a file or a folder, opened for no access at all: the handle can answer
        /// questions about it and do nothing else, and no byte of it is read or written. Invalid when Windows
        /// refuses; <c>Marshal.GetLastWin32Error</c>, read before any other call, then says why.
        /// </summary>
        public static SafeFileHandle OpenForQuestions(string path)
        {
            return CreateFileW(path, 0, ShareAll, IntPtr.Zero, OpenExisting, BackupSemantics, IntPtr.Zero);
        }

        /// <summary>
        /// Where what <paramref name="handle"/> is open on really is, every link and junction followed, as
        /// <see cref="WithoutWin32Prefix"/> writes it; an empty string when Windows cannot say, and
        /// <paramref name="win32Error"/> then says why.
        /// </summary>
        public static string FinalPath(SafeFileHandle handle, out int win32Error)
        {
            char[] buffer = new char[512];
            uint length = GetFinalPathNameByHandleW(handle, buffer, (uint)buffer.Length, 0);
            if (length >= buffer.Length)
            {
                buffer = new char[length + 1];
                length = GetFinalPathNameByHandleW(handle, buffer, (uint)buffer.Length, 0);
            }

            if (length == 0 || length >= buffer.Length)
            {
                win32Error = Marshal.GetLastWin32Error();
                return "";
            }

            win32Error = 0;
            return WithoutWin32Prefix(new string(buffer, 0, (int)length));
        }

        /// <summary>
        /// A path as <c>GetFinalPathNameByHandle</c> writes it, written as a person would: <c>\\?\UNC\server\share</c>
        /// as <c>\\server\share</c>, <c>\\?\C:\x</c> as <c>C:\x</c>, and no trailing separator but a drive
        /// root's (<c>C:\</c>).
        /// </summary>
        public static string WithoutWin32Prefix(string final)
        {
            string path = final;
            if (path.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase))
            {
                path = @"\\" + path.Substring(8);
            }
            else if (path.StartsWith(@"\\?\", StringComparison.Ordinal))
            {
                path = path.Substring(4);
            }

            if (path.Length > 3 && path.EndsWith(@"\", StringComparison.Ordinal))
            {
                path = path.TrimEnd('\\');
            }

            return path;
        }

        // From System32 and nowhere else, whichever assembly this is compiled into: Core says so for every
        // call it makes (CA5392), and an assembly Add-Type compiles says nothing.
        [DllImport("kernel32.dll", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode, SetLastError = true)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        public static extern SafeFileHandle CreateFileW(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetFileInformationByHandle(SafeFileHandle hFile, out ByHandleFileInformation lpFileInformation);

        [DllImport("kernel32.dll", EntryPoint = "GetFinalPathNameByHandleW", CharSet = CharSet.Unicode, SetLastError = true)]
        [DefaultDllImportSearchPaths(DllImportSearchPath.System32)]
        private static extern uint GetFinalPathNameByHandleW(
            SafeFileHandle hFile, [Out] char[] lpszFilePath, uint cchFilePath, uint dwFlags);
    }
}
