using System;
using System.Runtime.InteropServices;
using System.Security.Principal;
using Microsoft.Win32;

namespace OutlookAI.PolicyWriter
{
    /// <summary>
    /// The Windows half of the helper's checks, and its only registry access.
    ///
    /// <para>
    /// WHO ASKED. The helper is started by UAC on behalf of the process that called ShellExecuteEx
    /// "runas" - Outlook, with the add-in in it - and the AppInfo service makes that process the
    /// helper's parent. So the requester is the parent process, and this reads three things about it
    /// without trusting anything the helper was told:
    ///  - its process id, from a Toolhelp snapshot of this process's own entry;
    ///  - that it is still the process that started this one: alive, and created no later than this
    ///    one, so a recycled process id is not mistaken for it. Checked whenever its handle can be
    ///    opened for a limited query; across users (an administrator's credentials typed for a
    ///    standard user) Windows may deny even that, and then the session and user checks below
    ///    stand alone;
    ///  - its USER and its SESSION, from the session manager (WTSEnumerateProcesses), which reads
    ///    the token on the system's side - no handle to another user's token is needed, which an
    ///    administrator's process may well be denied.
    /// The parent must run in this process's own Windows session: the session whose desktop showed
    /// the UAC prompt.
    /// </para>
    ///
    /// <para>
    /// WHERE IT WRITES. <c>Registry.Users</c> - <c>HKEY_USERS\&lt;sid&gt;</c> - and never
    /// <c>Registry.CurrentUser</c>: under an administrator's credentials HKCU is the
    /// administrator's own hive, and a write there would change the wrong person's Outlook and
    /// report success. A hive that is not loaded is refused, never loaded.
    /// </para>
    /// </summary>
    internal sealed class WindowsPolicyHost : IPolicyWriterHost
    {
        public string RequesterSid(out string why)
        {
            uint self = GetCurrentProcessId();
            uint selfSession;
            if (!ProcessIdToSessionId(self, out selfSession))
            {
                why = "this process's own session could not be read (Win32 error " + Marshal.GetLastWin32Error() + ").";
                return string.Empty;
            }

            uint parent;
            if (!TryGetParentProcessId(self, out parent, out why))
                return string.Empty;

            if (!ParentStillStartedUs(parent, out why))
                return string.Empty;

            string sid;
            uint parentSession;
            if (!TryGetProcessUser(parent, out sid, out parentSession, out why))
                return string.Empty;

            if (parentSession != selfSession)
            {
                why = "the process that started this helper (pid " + parent + ") is in session " + parentSession + ", not in this helper's session " + selfSession + ".";
                return string.Empty;
            }

            why = string.Empty;
            return sid;
        }

        public bool IsUserHiveLoaded(string sid)
        {
            using (RegistryKey hive = Registry.Users.OpenSubKey(sid, false))
            {
                return hive != null;
            }
        }

        public void WriteDword(string sid, string keyPath, string name, int value)
        {
            using (RegistryKey key = Registry.Users.CreateSubKey(sid + "\\" + keyPath))
            {
                if (key == null)
                    throw new InvalidOperationException("HKEY_USERS\\" + sid + "\\" + keyPath + " could not be opened for writing.");
                key.SetValue(name, value, RegistryValueKind.DWord);
            }
        }

        public int? ReadDword(string sid, string keyPath, string name)
        {
            using (RegistryKey key = Registry.Users.OpenSubKey(sid + "\\" + keyPath, false))
            {
                if (key == null)
                    return null;
                object value = key.GetValue(name);
                if (value == null || key.GetValueKind(name) != RegistryValueKind.DWord)
                    return null;
                return (int)value;
            }
        }

        // ===== The parent process =====

        private static bool TryGetParentProcessId(uint self, out uint parent, out string why)
        {
            parent = 0;
            IntPtr snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
            if (snapshot == InvalidHandleValue)
            {
                why = "no process snapshot (Win32 error " + Marshal.GetLastWin32Error() + ").";
                return false;
            }

            try
            {
                var entry = new PROCESSENTRY32W { dwSize = (uint)Marshal.SizeOf(typeof(PROCESSENTRY32W)) };
                if (!Process32FirstW(snapshot, ref entry))
                {
                    why = "the process snapshot is empty (Win32 error " + Marshal.GetLastWin32Error() + ").";
                    return false;
                }

                do
                {
                    if (entry.th32ProcessID == self)
                    {
                        parent = entry.th32ParentProcessID;
                        if (parent == 0)
                        {
                            why = "this process has no parent process id.";
                            return false;
                        }
                        why = string.Empty;
                        return true;
                    }
                    entry.dwSize = (uint)Marshal.SizeOf(typeof(PROCESSENTRY32W));
                }
                while (Process32NextW(snapshot, ref entry));
            }
            finally
            {
                CloseHandle(snapshot);
            }

            why = "this process is not in its own process snapshot.";
            return false;
        }

        /// <summary>
        /// False when the parent id now belongs to a process created AFTER this one - the parent
        /// exited and the id was recycled - or when it has exited. True when its handle cannot be
        /// opened at all: across users Windows may refuse even a limited query, and the session
        /// manager's answer below then stands on its own.
        /// </summary>
        private static bool ParentStillStartedUs(uint parent, out string why)
        {
            long selfCreated, ignoredExit, ignoredKernel, ignoredUser;
            if (!GetProcessTimes(GetCurrentProcess(), out selfCreated, out ignoredExit, out ignoredKernel, out ignoredUser))
            {
                why = "this process's own start time could not be read (Win32 error " + Marshal.GetLastWin32Error() + ").";
                return false;
            }

            IntPtr handle = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, false, parent);
            if (handle == IntPtr.Zero)
            {
                why = string.Empty;
                return true;
            }

            try
            {
                long parentCreated;
                if (!GetProcessTimes(handle, out parentCreated, out ignoredExit, out ignoredKernel, out ignoredUser))
                {
                    why = "the start time of the process that started this helper (pid " + parent + ") could not be read (Win32 error " + Marshal.GetLastWin32Error() + ").";
                    return false;
                }
                if (parentCreated > selfCreated)
                {
                    why = "process id " + parent + " now belongs to a process started after this helper: the process that started it has exited.";
                    return false;
                }

                uint exitCode;
                if (GetExitCodeProcess(handle, out exitCode) && exitCode != STILL_ACTIVE)
                {
                    why = "the process that started this helper (pid " + parent + ") has exited.";
                    return false;
                }

                why = string.Empty;
                return true;
            }
            finally
            {
                CloseHandle(handle);
            }
        }

        private static bool TryGetProcessUser(uint pid, out string sid, out uint session, out string why)
        {
            sid = string.Empty;
            session = 0;
            IntPtr list;
            int count;
            if (!WTSEnumerateProcessesW(IntPtr.Zero, 0, 1, out list, out count))
            {
                why = "the session manager would not list processes (Win32 error " + Marshal.GetLastWin32Error() + ").";
                return false;
            }

            try
            {
                int size = Marshal.SizeOf(typeof(WTS_PROCESS_INFOW));
                for (int i = 0; i < count; i++)
                {
                    var info = (WTS_PROCESS_INFOW)Marshal.PtrToStructure(new IntPtr(list.ToInt64() + (long)i * size), typeof(WTS_PROCESS_INFOW));
                    if (info.ProcessId != pid)
                        continue;
                    if (info.pUserSid == IntPtr.Zero)
                    {
                        why = "the session manager does not say which user process " + pid + " runs as.";
                        return false;
                    }
                    sid = new SecurityIdentifier(info.pUserSid).Value;
                    session = info.SessionId;
                    why = string.Empty;
                    return true;
                }
            }
            finally
            {
                WTSFreeMemory(list);
            }

            why = "the process that started this helper (pid " + pid + ") is no longer running.";
            return false;
        }

        // ===== Win32 =====

        private const uint TH32CS_SNAPPROCESS = 0x00000002;
        private const uint PROCESS_QUERY_LIMITED_INFORMATION = 0x1000;
        private const uint STILL_ACTIVE = 259;
        private static readonly IntPtr InvalidHandleValue = new IntPtr(-1);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct PROCESSENTRY32W
        {
            public uint dwSize;
            public uint cntUsage;
            public uint th32ProcessID;
            public IntPtr th32DefaultHeapID;
            public uint th32ModuleID;
            public uint cntThreads;
            public uint th32ParentProcessID;
            public int pcPriClassBase;
            public uint dwFlags;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szExeFile;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct WTS_PROCESS_INFOW
        {
            public uint SessionId;
            public uint ProcessId;
            public IntPtr pProcessName;
            public IntPtr pUserSid;
        }

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentProcessId();

        [DllImport("kernel32.dll")]
        private static extern IntPtr GetCurrentProcess();

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool ProcessIdToSessionId(uint processId, out uint sessionId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr CreateToolhelp32Snapshot(uint flags, uint processId);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool Process32FirstW(IntPtr snapshot, ref PROCESSENTRY32W entry);

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool Process32NextW(IntPtr snapshot, ref PROCESSENTRY32W entry);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr OpenProcess(uint access, bool inheritHandle, uint processId);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GetProcessTimes(IntPtr process, out long creation, out long exit, out long kernel, out long user);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool GetExitCodeProcess(IntPtr process, out uint exitCode);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool CloseHandle(IntPtr handle);

        [DllImport("wtsapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool WTSEnumerateProcessesW(IntPtr server, int reserved, int version, out IntPtr processInfo, out int count);

        [DllImport("wtsapi32.dll")]
        private static extern void WTSFreeMemory(IntPtr memory);
    }
}
