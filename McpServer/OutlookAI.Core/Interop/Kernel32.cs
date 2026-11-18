using System;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace OutlookAI.Core.Interop
{
    /// <summary>
    /// The kernel32 calls two parts of Core each declared, declared once (the code-sharing refactor's S5a,
    /// Q154): opening a file or a folder by its path, which the audit log's append handle and the attachment
    /// writer's checks of a folder and a file make, and closing a process or a thread handle, which the session
    /// registry and the start observer make. The current process's id is <c>NativeWindows</c>' in the add-in's
    /// part of Core, whose window calls need it too.
    /// </summary>
    internal static class Kernel32
    {
        [DllImport("kernel32.dll", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern SafeFileHandle CreateFileW(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool CloseHandle(IntPtr handle);
    }
}
