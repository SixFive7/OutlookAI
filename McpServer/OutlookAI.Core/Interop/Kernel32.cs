using System;
using System.Runtime.InteropServices;

namespace OutlookAI.Core.Interop
{
    /// <summary>
    /// Closing a process or a thread handle, which the session registry and the start observer do, declared
    /// once (the code-sharing refactor's S5a, Q154). Opening a file or a folder by its path, which this class
    /// declared too, is <see cref="Win32File"/>'s since Q171 (c), because <c>Testbed/host/New-LiveTestSettings.ps1</c>
    /// compiles that file as well. The current process's id is <c>NativeWindows</c>' in the add-in's part of
    /// Core, whose window calls need it too.
    /// </summary>
    internal static class Kernel32
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool CloseHandle(IntPtr handle);
    }
}
