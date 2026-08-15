using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace OutlookAI.ComHost.Supervision
{
    /// <summary>
    /// A Windows Job Object with KILL_ON_JOB_CLOSE, used to guarantee the COM child dies
    /// with its parent.
    /// <para>
    /// This closes a leak observed in the field on 2026-08-15: 18 orphaned
    /// OutlookAI.McpServer processes were found on one machine, one of them wedged
    /// holding Outlook COM references. A killed or crashed parent cannot run cleanup
    /// code, so process-tree lifetime has to be enforced by the kernel rather than by
    /// anything we remember to call. The handle is closed when the parent exits by any
    /// means, including a hard kill, and Windows then terminates every process in the
    /// job.
    /// </para>
    /// <para>
    /// Best-effort by construction: on a system where job assignment fails (an existing
    /// job with incompatible limits, or a restricted token) the supervisor still works -
    /// it simply loses the belt-and-braces guarantee and relies on its own teardown plus
    /// the child's own orphan check.
    /// </para>
    /// </summary>
    [SupportedOSPlatform("windows")]
    internal sealed class ChildJobObject : IDisposable
    {
        private IntPtr _handle;
        private bool _disposed;

        private ChildJobObject(IntPtr handle)
        {
            _handle = handle;
        }

        /// <summary>True when a real job object backs this instance.</summary>
        internal bool IsActive => _handle != IntPtr.Zero;

        /// <summary>
        /// Creates a kill-on-close job, or an inert instance when the OS refuses. Never
        /// throws: losing the job must not prevent the server from working.
        /// </summary>
        internal static ChildJobObject CreateOrInert()
        {
            try
            {
                IntPtr handle = NativeMethods.CreateJobObject(IntPtr.Zero, null);
                if (handle == IntPtr.Zero)
                {
                    return new ChildJobObject(IntPtr.Zero);
                }

                NativeMethods.JOBOBJECT_EXTENDED_LIMIT_INFORMATION info = default;
                info.BasicLimitInformation.LimitFlags = NativeMethods.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE;

                int length = Marshal.SizeOf<NativeMethods.JOBOBJECT_EXTENDED_LIMIT_INFORMATION>();
                IntPtr buffer = Marshal.AllocHGlobal(length);
                try
                {
                    Marshal.StructureToPtr(info, buffer, fDeleteOld: false);
                    bool ok = NativeMethods.SetInformationJobObject(
                        handle,
                        NativeMethods.JobObjectExtendedLimitInformation,
                        buffer,
                        (uint)length);
                    if (!ok)
                    {
                        _ = NativeMethods.CloseHandle(handle);
                        return new ChildJobObject(IntPtr.Zero);
                    }
                }
                finally
                {
                    Marshal.FreeHGlobal(buffer);
                }

                return new ChildJobObject(handle);
            }
            catch (Exception)
            {
                // DllNotFound / EntryPointNotFound / security failures all degrade the
                // same way: no job, everything else still works.
                return new ChildJobObject(IntPtr.Zero);
            }
        }

        /// <summary>
        /// Puts <paramref name="process"/> in the job. Returns false when the process
        /// could not be assigned, which is informational only.
        /// </summary>
        internal bool TryAssign(Process process)
        {
            ArgumentNullException.ThrowIfNull(process);

            if (!IsActive)
            {
                return false;
            }

            try
            {
                return NativeMethods.AssignProcessToJobObject(_handle, process.Handle);
            }
            catch (Win32Exception)
            {
                return false;
            }
            catch (InvalidOperationException)
            {
                // The process exited between spawn and assignment.
                return false;
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            IntPtr handle = Interlocked.Exchange(ref _handle, IntPtr.Zero);
            if (handle != IntPtr.Zero)
            {
                // Closing the last handle terminates every process still in the job.
                _ = NativeMethods.CloseHandle(handle);
            }
        }

        private static class NativeMethods
        {
            internal const uint JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE = 0x00002000;
            internal const int JobObjectExtendedLimitInformation = 9;

            [StructLayout(LayoutKind.Sequential)]
            internal struct JOBOBJECT_BASIC_LIMIT_INFORMATION
            {
                public long PerProcessUserTimeLimit;
                public long PerJobUserTimeLimit;
                public uint LimitFlags;
                public nuint MinimumWorkingSetSize;
                public nuint MaximumWorkingSetSize;
                public uint ActiveProcessLimit;
                public nuint Affinity;
                public uint PriorityClass;
                public uint SchedulingClass;
            }

            [StructLayout(LayoutKind.Sequential)]
            internal struct IO_COUNTERS
            {
                public ulong ReadOperationCount;
                public ulong WriteOperationCount;
                public ulong OtherOperationCount;
                public ulong ReadTransferCount;
                public ulong WriteTransferCount;
                public ulong OtherTransferCount;
            }

            [StructLayout(LayoutKind.Sequential)]
            internal struct JOBOBJECT_EXTENDED_LIMIT_INFORMATION
            {
                public JOBOBJECT_BASIC_LIMIT_INFORMATION BasicLimitInformation;
                public IO_COUNTERS IoInfo;
                public nuint ProcessMemoryLimit;
                public nuint JobMemoryLimit;
                public nuint PeakProcessMemoryUsed;
                public nuint PeakJobMemoryUsed;
            }

            [DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
            internal static extern IntPtr CreateJobObject(IntPtr lpJobAttributes, string? lpName);

            [DllImport("kernel32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool SetInformationJobObject(
                IntPtr hJob, int jobObjectInformationClass, IntPtr lpJobObjectInformation, uint cbJobObjectInformationLength);

            [DllImport("kernel32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool AssignProcessToJobObject(IntPtr hJob, IntPtr hProcess);

            [DllImport("kernel32.dll", SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            internal static extern bool CloseHandle(IntPtr hObject);
        }
    }
}
