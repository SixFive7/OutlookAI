using System;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32.SafeHandles;

namespace OutlookAI.Core.Audit
{
    /// <summary>
    /// The audit log's file handle, opened through Windows directly (Q117) - because what makes an
    /// append safe is an access right .NET's <c>FileStream</c> never asks for.
    /// <para>
    /// <b>Why not <c>FileMode.Append</c>.</b> .NET opens an "append" stream with full write access
    /// and EMULATES append: it notes the end-of-file when the file is opened and writes there
    /// (<c>SafeFileHandle.Windows.cs</c> / <c>OSFileStreamStrategy.cs</c>; dotnet/runtime PR #55465,
    /// which would have changed it, was never merged). Two writers whose opens overlap therefore
    /// write to the same offset and one line silently overwrites the other - 5,883 of 6,000 lines
    /// survived a two-process test without a lock (2026-10-03).
    /// </para>
    /// <para>
    /// <b>What this asks for instead.</b> <c>FILE_APPEND_DATA</c> WITHOUT <c>FILE_WRITE_DATA</c>:
    /// Windows then ignores the offset of every write on the handle and puts it at the end of the
    /// file as that write runs ("If only the FILE_APPEND_DATA and SYNCHRONIZE flags are set, the caller
    /// can write only to the end of file, and any offset information about writes to the file is
    /// ignored", CreateFileW). It is the absence of <c>FILE_WRITE_DATA</c> that decides it, so the
    /// handle also carries <c>FILE_READ_DATA</c> - for the one-byte look at the tail that tells a
    /// writer a crash left a line unfinished - and <c>FILE_READ_ATTRIBUTES</c>, for the file's
    /// identity. Sharing is read and write, never delete: a rename cannot land in the middle of an
    /// append (the procedure for archiving the log rests on that - see <c>AuditLogReaderTests</c>).
    /// </para>
    /// <para>
    /// Plain <c>DllImport</c> over <c>byte[]</c> and <see cref="SafeFileHandle"/>, nothing newer:
    /// the .NET Framework 4.8 add-in is going to write this same log through this same code.
    /// </para>
    /// </summary>
    internal static class AuditLogFile
    {
        private const uint FileReadData = 0x0001;
        private const uint FileAppendData = 0x0004;
        private const uint FileReadAttributes = 0x0080;
        private const uint Synchronize = 0x00100000;
        private const uint FileShareRead = 0x00000001;
        private const uint FileShareWrite = 0x00000002;
        private const uint OpenAlways = 4;
        private const uint FileAttributeNormal = 0x00000080;
        private const uint FileFlagWriteThrough = 0x80000000;
        private const int FileIdInfoClass = 18;

        private const int ErrorFileNotFound = 2;
        private const int ErrorPathNotFound = 3;
        private const int ErrorAccessDenied = 5;
        private const int ErrorSharingViolation = 32;
        private const int ErrorLockViolation = 33;

        /// <summary>
        /// Opens (creating when absent) the log at <paramref name="path"/> for appending, the way every
        /// writer opens it. False with the Win32 error when it could not be opened.
        /// </summary>
        /// <param name="path">The log file.</param>
        /// <param name="writeThrough">
        /// <c>FILE_FLAG_WRITE_THROUGH</c>: a write returns once it is on the disk rather than in the cache.
        /// </param>
        /// <param name="handle">The open handle, or null.</param>
        /// <param name="win32Error">0, or why the open failed.</param>
        internal static bool TryOpenForAppend(string path, bool writeThrough, out SafeFileHandle? handle, out int win32Error)
        {
            SafeFileHandle opened = CreateFileW(
                Win32Path(path),
                FileAppendData | FileReadData | FileReadAttributes | Synchronize,
                FileShareRead | FileShareWrite,
                IntPtr.Zero,
                OpenAlways,
                FileAttributeNormal | (writeThrough ? FileFlagWriteThrough : 0u),
                IntPtr.Zero);
            if (opened.IsInvalid)
            {
                win32Error = Marshal.GetLastWin32Error();
                opened.Dispose();
                handle = null;
                return false;
            }

            win32Error = 0;
            handle = opened;
            return true;
        }

        /// <summary>Whether an open failed only because another handle is in the way - worth a retry.</summary>
        internal static bool IsSharingViolation(int win32Error)
        {
            return win32Error == ErrorSharingViolation || win32Error == ErrorLockViolation;
        }

        /// <summary>
        /// The exception a <c>FileStream</c> would have thrown for <paramref name="win32Error"/>, so the
        /// callers' existing handling - and the type name the health probe reports - stay what they were.
        /// </summary>
        internal static Exception ErrorFor(int win32Error, string path)
        {
            string message = new Win32Exception(win32Error).Message + " (" + path + ")";
            int hresult = unchecked((int)0x80070000 | (win32Error & 0xFFFF));
            switch (win32Error)
            {
                case ErrorAccessDenied:
                    return new UnauthorizedAccessException(message);
                case ErrorFileNotFound:
                    return new FileNotFoundException(message, path);
                case ErrorPathNotFound:
                    return new DirectoryNotFoundException(message);
                default:
                    return new IOException(message, hresult);
            }
        }

        /// <summary>
        /// True when the file is not empty and its last byte is not a line feed: a line was left
        /// unfinished - by a crash, a power cut, or a writer from before this one that wrote a long line
        /// in pieces. The next line must then start with a line feed of its own, or it would be glued
        /// to the fragment and lost with it.
        /// </summary>
        internal static bool EndsMidLine(SafeFileHandle handle)
        {
            if (!GetFileSizeEx(handle, out long size))
            {
                throw ErrorFor(Marshal.GetLastWin32Error(), "audit log");
            }

            if (size <= 0)
            {
                return false;
            }

            long at = size - 1;
            NativeOverlapped position = new NativeOverlapped
            {
                OffsetLow = unchecked((int)(at & 0xFFFFFFFFL)),
                OffsetHigh = unchecked((int)(at >> 32)),
            };
            byte[] last = new byte[1];
            if (!ReadFile(handle, last, 1, out int read, ref position))
            {
                throw ErrorFor(Marshal.GetLastWin32Error(), "audit log");
            }

            return read == 1 && last[0] != (byte)'\n';
        }

        /// <summary>
        /// Writes <paramref name="count"/> bytes in ONE <c>WriteFile</c> call - one call per line, so a
        /// line is never split across writes another writer's line could land between. Throws when the
        /// write fails or is short; it is never retried, because a retry after a partial write would
        /// write part of the line twice.
        /// </summary>
        internal static void Append(SafeFileHandle handle, byte[] bytes, int count)
        {
            if (!WriteFile(handle, bytes, count, out int written, IntPtr.Zero))
            {
                throw ErrorFor(Marshal.GetLastWin32Error(), "audit log");
            }

            if (written != count)
            {
                throw new IOException(
                    "Only " + written.ToString(CultureInfo.InvariantCulture) + " of " + count.ToString(CultureInfo.InvariantCulture)
                    + " bytes of the audit line were written.");
            }
        }

        /// <summary>
        /// The file's identity - its volume and its file id, which a rename keeps and a new file under
        /// the old name does not - or null when the file system will not say. Creation time is no
        /// substitute: NTFS "tunnelling" gives a file created under a just-renamed name the old
        /// file's creation time for fifteen seconds, which is exactly what archiving the log does.
        /// </summary>
        internal static string? TryGetIdentity(SafeFileHandle handle)
        {
            if (GetFileInformationByHandleEx(handle, FileIdInfoClass, out FileIdInfo id, Marshal.SizeOf(typeof(FileIdInfo))))
            {
                return "v" + Hex(id.VolumeSerialNumber, 16) + "-" + Hex(id.FileIdHigh, 16) + Hex(id.FileIdLow, 16);
            }

            if (GetFileInformationByHandle(handle, out ByHandleFileInformation info))
            {
                return "v" + Hex(info.VolumeSerialNumber, 8) + "-" + Hex(info.FileIndexHigh, 8) + Hex(info.FileIndexLow, 8);
            }

            return null;
        }

        /// <summary>This process's id, without <c>Environment.ProcessId</c> (which .NET Framework does not have).</summary>
        internal static int CurrentProcessId()
        {
            return unchecked((int)GetCurrentProcessId());
        }

        /// <summary>
        /// The current UTC time from <c>GetSystemTimePreciseAsFileTime</c>. <c>DateTime.UtcNow</c> is
        /// that on .NET 10 but the coarse clock on .NET Framework, which lags by up to a timer tick
        /// (15.6 ms by default) - so a Framework writer taking its turn after a .NET 10 writer could
        /// stamp an EARLIER time on the later line, and the file would stop being in time order. One
        /// clock for both runtimes keeps the lock's promise across them.
        /// </summary>
        internal static DateTime PreciseUtcNow()
        {
            GetSystemTimePreciseAsFileTime(out long fileTime);
            return DateTime.FromFileTimeUtc(fileTime);
        }

        private static string Win32Path(string path)
        {
            // CreateFileW, unlike FileStream, does not add the long-path prefix itself.
            string full = Path.GetFullPath(path);
            if (full.Length < 260 || full.StartsWith(@"\\?\", StringComparison.Ordinal))
            {
                return full;
            }

            return full.StartsWith(@"\\", StringComparison.Ordinal) ? @"\\?\UNC\" + full.Substring(2) : @"\\?\" + full;
        }

        private static string Hex(ulong value, int digits)
        {
            return value.ToString("x" + digits.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct FileIdInfo
        {
            public ulong VolumeSerialNumber;
            public ulong FileIdLow;
            public ulong FileIdHigh;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct ByHandleFileInformation
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

        [DllImport("kernel32.dll", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFileW(
            string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            IntPtr hTemplateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool WriteFile(
            SafeFileHandle hFile, byte[] lpBuffer, int nNumberOfBytesToWrite, out int lpNumberOfBytesWritten, IntPtr lpOverlapped);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ReadFile(
            SafeFileHandle hFile, byte[] lpBuffer, int nNumberOfBytesToRead, out int lpNumberOfBytesRead, ref NativeOverlapped lpOverlapped);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileSizeEx(SafeFileHandle hFile, out long lpFileSize);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandleEx(
            SafeFileHandle hFile, int fileInformationClass, out FileIdInfo lpFileInformation, int dwBufferSize);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetFileInformationByHandle(SafeFileHandle hFile, out ByHandleFileInformation lpFileInformation);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentProcessId();

        [DllImport("kernel32.dll")]
        private static extern void GetSystemTimePreciseAsFileTime(out long lpSystemTimeAsFileTime);
    }
}
