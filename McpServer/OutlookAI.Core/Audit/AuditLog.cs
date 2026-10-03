using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;

namespace OutlookAI.Core.Audit
{
    /// <summary>
    /// Structured write-op audit log (v3.MD sections 0.5/0.5.2, LIVE and load-bearing
    /// from Phase 4): every write operation of the product appends one structured line
    /// under the shared %LOCALAPPDATA%\OutlookAI state root - a gitignored,
    /// machine-local location (S6). Host-neutral: no MCP types, no console assumptions;
    /// the add-in can share the same file in v3.1.
    ///
    /// Line grammar (one line per operation, parse-friendly):
    ///   ts=2026-07-23T10:11:12.345Z op=new_draft key="value" key2="value2"
    /// Values are quoted with backslash escapes for '\', '"', CR, LF and TAB; fields
    /// with null values are omitted. <see cref="Append"/> THROWS when the line cannot
    /// be written - a write operation without its audit line must be surfaced, never
    /// silently swallowed (D4 discipline).
    ///
    /// <para>
    /// HOW THE FILE IS OPENED, because a maintainer renaming the log has to know (Q86): every
    /// append opens the file, writes one line and closes it again - no handle is ever held
    /// between two lines. The open is <c>FileMode.Append</c> with <c>FileShare.ReadWrite</c> and
    /// deliberately WITHOUT <c>FileShare.Delete</c>, so the file cannot be renamed or deleted
    /// under an append in progress; an open that meets a sharing violation is retried
    /// <see cref="WriteRetries"/> times, 15 ms apart, before the operation reports it. Appends to
    /// one log are serialized across processes by a named mutex, because two that overlap would
    /// otherwise overwrite each other (see <see cref="AppendTo"/>); the mutex holds no file handle,
    /// so it changes nothing about renaming the file.
    /// </para>
    ///
    /// <para>
    /// A PROCESS CAN BE REDIRECTED, and only a test process ever is (Q86). The non-live suite
    /// drives the real write paths through fakes, and they all end in <see cref="Append"/>; until
    /// this existed every such line landed in the maintainer's real log, which measured 64%
    /// test noise on 2026-09-28. The test assembly's module initializer now calls
    /// <see cref="RedirectThisProcess"/> once, through <c>InternalsVisibleTo</c>, before any test
    /// runs. It is not a setting: nothing a user or an installed process can reach changes it,
    /// <see cref="DefaultDirectory"/> is untouched, and a shipped process never calls it.
    /// </para>
    /// </summary>
    public static class AuditLog
    {
        /// <summary>The file name the log has in whichever directory holds it.</summary>
        public const string LogFileName = "audit.log";

        /// <summary>The <c>error</c> a probe reports for a directory a redirected process may not touch.</summary>
        internal const string RedirectRefusalError = "RedirectedForTests";

        private const int WriteRetries = 3;

        /// <summary>
        /// How long an append waits for another writer of the same log before appending anyway. An
        /// append holds the lock for one open-write-close - about a millisecond - so two seconds is
        /// only ever reached by a writer that is stuck, and a stuck writer must not stall every
        /// other session's drafts and sends behind it.
        /// </summary>
        internal const int WriterLockWaitMilliseconds = 2000;

        /// <summary>
        /// The throwaway directory this process was redirected to, or null - which it is in every
        /// shipped process. Written once, by <see cref="RedirectThisProcess"/>.
        /// </summary>
        private static string? _redirectedDirectory;

        /// <summary>Shared OutlookAI state root (v3.MD section 0.5.2).</summary>
        public static string DefaultDirectory =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookAI");

        /// <summary>Full path of the audit log file every write op appends to.</summary>
        public static string DefaultLogPath => Path.Combine(DefaultDirectory, LogFileName);

        /// <summary>
        /// The directory THIS process appends to: <see cref="DefaultDirectory"/> in every shipped
        /// process, and the throwaway directory in a test process that redirected itself.
        /// Everything that writes, probes or reads the log on this process's behalf goes through
        /// this, so the three can never disagree about which file they mean.
        /// </summary>
        internal static string EffectiveDirectory => Volatile.Read(ref _redirectedDirectory) ?? DefaultDirectory;

        /// <summary>Full path of the log THIS process appends to (see <see cref="EffectiveDirectory"/>).</summary>
        internal static string EffectiveLogPath => Path.Combine(EffectiveDirectory, LogFileName);

        /// <summary>True only in a process that redirected its audit lines - a test host, never a shipped process.</summary>
        internal static bool IsRedirected => Volatile.Read(ref _redirectedDirectory) != null;

        /// <summary>
        /// Sends every audit line this process writes from now on to <paramref name="directory"/>
        /// instead of the real log, for the rest of the process's life. Test hosts only.
        /// <para>
        /// The target has to be under the temp directory - throwaway by definition - and may not
        /// be the real audit directory or anything inside it. It is set once: a second call with
        /// the same directory is a no-op, and one naming a different directory throws, because
        /// a run whose lines are split across two directories cannot be proven isolated.
        /// </para>
        /// </summary>
        internal static void RedirectThisProcess(string directory)
        {
            string? refusal = DescribeRedirectRefusal(directory);
            if (refusal != null)
            {
                throw new ArgumentException(refusal, nameof(directory));
            }

            string target = NormalizeDirectory(directory);
            string? previous = Interlocked.CompareExchange(ref _redirectedDirectory, target, null);
            if (previous != null && !string.Equals(previous, target, StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException(
                    "This process already sends its audit lines to '" + previous + "'. A second redirect to '" + target
                    + "' would split one run's lines across two directories.");
            }
        }

        /// <summary>
        /// Why <paramref name="directory"/> cannot receive a redirected process's audit lines, or
        /// null when it can. Pure apart from resolving the two roots it compares against.
        /// </summary>
        internal static string? DescribeRedirectRefusal(string directory)
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                return "The redirect directory must not be blank.";
            }

            if (!Path.IsPathRooted(directory))
            {
                return "The redirect directory must be an absolute path: '" + directory + "' is not.";
            }

            try
            {
                if (IsSameOrUnder(directory, DefaultDirectory))
                {
                    return "'" + directory + "' is the real audit directory or inside it, which is exactly what a redirect exists to keep a test process out of.";
                }

                if (!IsSameOrUnder(directory, Path.GetTempPath()))
                {
                    return "'" + directory + "' is not under the temp directory ('" + Path.GetTempPath()
                        + "'), so it is not throwaway.";
                }
            }
            catch (Exception ex) when (ex is ArgumentException || ex is NotSupportedException || ex is PathTooLongException)
            {
                return "'" + directory + "' is not a usable path (" + ex.GetType().Name + ").";
            }

            return null;
        }

        /// <summary>
        /// Why a redirected process may not write or probe the audit log under
        /// <paramref name="directory"/>, or null when it may - and always null in a process that
        /// was never redirected. Checked BEFORE anything is created or opened.
        /// <para>
        /// An allowlist rather than a single forbidden path, on purpose: the real directory is
        /// refused in every spelling, because nothing outside the temp directory is allowed at
        /// all - a short name, a <c>\\?\</c> prefix or a path built by hand from
        /// <c>%LOCALAPPDATA%</c> lands outside it just the same.
        /// </para>
        /// </summary>
        internal static string? DescribeRefusalWhileRedirected(string directory)
        {
            string? redirectedTo = Volatile.Read(ref _redirectedDirectory);
            if (redirectedTo == null)
            {
                return null;
            }

            if (IsSameOrUnder(directory, DefaultDirectory))
            {
                return "This is a test process: its audit lines go to '" + redirectedTo + "', and the real audit log under '"
                    + DefaultDirectory + "' is refused, so no test can write or read it.";
            }

            if (!IsSameOrUnder(directory, Path.GetTempPath()))
            {
                return "This is a test process: its audit lines go to '" + redirectedTo + "', and it may write or read an audit "
                    + "log only under the temp directory - '" + directory + "' is outside it.";
            }

            return null;
        }

        /// <summary>
        /// Appends one structured line for <paramref name="operation"/> to this process's audit
        /// log (<see cref="EffectiveDirectory"/> - the default one in every shipped process).
        /// Throws <see cref="InvalidOperationException"/> when the line cannot be written
        /// (load-bearing from Phase 4 - callers surface the failure).
        /// </summary>
        public static void Append(string operation, params (string Key, string? Value)[] fields)
        {
            AppendTo(EffectiveDirectory, operation, fields);
        }

        /// <summary>
        /// Appends one structured line to <paramref name="directory"/>/audit.log
        /// (creating the directory when missing). Directory-parameterized for tests;
        /// production callers use <see cref="Append"/>.
        /// </summary>
        public static void AppendTo(string directory, string operation, IReadOnlyList<(string Key, string? Value)> fields)
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new ArgumentException("Audit directory must not be blank.", nameof(directory));
            }

            string line = FormatLine(DateTime.UtcNow, operation, fields);
            string path = Path.Combine(directory, LogFileName);
            try
            {
                // Before CreateDirectory, so a refused directory is not even created.
                string? refusal = DescribeRefusalWhileRedirected(directory);
                if (refusal != null)
                {
                    throw new InvalidOperationException("Audit line was not written to '" + path + "'. " + refusal);
                }

                Directory.CreateDirectory(directory);
                IOException? lastIo = null;

                // One writer at a time per log, across processes. FileShare.ReadWrite lets several
                // server processes (one per agent session) append to the same file, but
                // FileMode.Append only positions each handle at the end-of-file it saw when it
                // OPENED, and every write goes to that offset - so two appends whose opens overlap
                // write to the same place and one line silently overwrites the other. Measured
                // 2026-10-03 on .NET 10.0.12: two processes calling this method 3,000 times each at
                // once kept 5,883 lines of 6,000, with no exception anywhere - a silent loss in the
                // one log whose rule is that a write never goes unrecorded. See WriterLock.
                using (WriterLock.Acquire(path))
                {
                    for (int attempt = 0; attempt < WriteRetries; attempt++)
                    {
                        try
                        {
                            using (FileStream stream = new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                            using (StreamWriter writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false)))
                            {
                                writer.WriteLine(line);
                            }

                            return;
                        }
                        catch (IOException ex)
                        {
                            lastIo = ex;
                            Thread.Sleep(15);
                        }
                    }
                }

                throw new InvalidOperationException(
                    "Audit line could not be written to '" + path + "' after " +
                    WriteRetries.ToString(CultureInfo.InvariantCulture) + " attempts.", lastIo);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException
                || ex is NotSupportedException || ex is ArgumentException)
            {
                // E.g. the directory path is blocked by a same-named file. The retry
                // loop's own give-up exception above is already an
                // InvalidOperationException and passes through untouched.
                throw new InvalidOperationException("Audit line could not be written to '" + path + "'.", ex);
            }
        }

        /// <summary>
        /// Probes whether an audit line COULD be appended under
        /// <paramref name="directory"/> without writing anything: creates the directory
        /// when missing and opens (creating if absent) the log file for append with the
        /// same sharing the writers use, then closes it. Content-free error reason on
        /// failure (S4). Used by the health tool - write ops fail-closed without audit.
        /// A redirected (test) process gets <see cref="RedirectRefusalError"/> for any directory
        /// it may not touch, before anything is created or opened.
        /// </summary>
        public static bool TryProbeWritable(string directory, out string? error)
        {
            try
            {
                if (DescribeRefusalWhileRedirected(directory) != null)
                {
                    error = RedirectRefusalError;
                    return false;
                }

                Directory.CreateDirectory(directory);
                string path = Path.Combine(directory, LogFileName);
                using (new FileStream(path, FileMode.Append, FileAccess.Write, FileShare.ReadWrite))
                {
                }

                error = null;
                return true;
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException
                || ex is NotSupportedException || ex is ArgumentException)
            {
                error = ex.GetType().Name;
                return false;
            }
        }

        /// <summary>
        /// Whether <paramref name="path"/> is <paramref name="root"/> or lies under it, compared
        /// as full paths, case-insensitively, with a <c>\\?\</c> prefix and trailing separators
        /// removed. Throws what <see cref="Path.GetFullPath(string)"/> throws for an unusable path.
        /// </summary>
        internal static bool IsSameOrUnder(string path, string root)
        {
            string candidate = NormalizeDirectory(path);
            string container = NormalizeDirectory(root);
            if (string.Equals(candidate, container, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return candidate.StartsWith(container + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeDirectory(string path)
        {
            // The device prefix comes off BEFORE GetFullPath, because GetFullPath leaves a
            // \\?\ path exactly as written - '..' segments included - and the comparison
            // has to see the path the file system would.
            string spelled = path.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);
            if (spelled.StartsWith(@"\\?\UNC\", StringComparison.OrdinalIgnoreCase))
            {
                spelled = @"\\" + spelled.Substring(8);
            }
            else if (spelled.StartsWith(@"\\?\", StringComparison.Ordinal) || spelled.StartsWith(@"\\.\", StringComparison.Ordinal))
            {
                spelled = spelled.Substring(4);
            }

            return Path.GetFullPath(spelled).TrimEnd(Path.DirectorySeparatorChar);
        }

        /// <summary>
        /// The name of the mutex that serializes appends to the log at <paramref name="path"/>:
        /// in the session's own namespace (every agent session's server runs in the user's
        /// session), one per log file - a hash of its full path, case-folded - so the real log and
        /// a test run's throwaway one never wait on each other.
        /// </summary>
        internal static string WriterMutexName(string path)
        {
            string full = Path.GetFullPath(path).ToUpperInvariant();
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(full));
                return @"Local\OutlookAI.AuditLog." + BitConverter.ToString(hash).Replace("-", string.Empty);
            }
        }

        /// <summary>
        /// The one-writer-at-a-time lock around an append (see <see cref="AppendTo"/> for the loss
        /// it prevents). Best effort by design, because the APPEND is what is load-bearing, not the
        /// lock: a writer that cannot get it within <see cref="WriterLockWaitMilliseconds"/>, or
        /// cannot create it at all, appends anyway - exactly as every append did before the lock
        /// existed - rather than fail the draft or send the line records. A previous holder that
        /// died mid-append leaves the mutex abandoned, which hands ownership over normally.
        /// </summary>
        private sealed class WriterLock : IDisposable
        {
            private readonly Mutex? _mutex;
            private readonly bool _owned;

            private WriterLock(Mutex? mutex, bool owned)
            {
                _mutex = mutex;
                _owned = owned;
            }

            internal static WriterLock Acquire(string path)
            {
                Mutex? mutex = null;
                try
                {
                    mutex = new Mutex(initiallyOwned: false, WriterMutexName(path));
                    bool owned;
                    try
                    {
                        owned = mutex.WaitOne(WriterLockWaitMilliseconds);
                    }
                    catch (AbandonedMutexException)
                    {
                        owned = true;
                    }

                    return new WriterLock(mutex, owned);
                }
                catch (Exception ex) when (ex is UnauthorizedAccessException || ex is IOException
                    || ex is WaitHandleCannotBeOpenedException || ex is ArgumentException)
                {
                    // E.g. a same-named object of another type, or one created by a process this
                    // one may not open. Unserialized is how every append worked before.
                    mutex?.Dispose();
                    return new WriterLock(null, false);
                }
            }

            public void Dispose()
            {
                if (_mutex == null)
                {
                    return;
                }

                if (_owned)
                {
                    _mutex.ReleaseMutex();
                }

                _mutex.Dispose();
            }
        }

        /// <summary>
        /// Formats one audit line (pure logic, T1-tested). The operation name must be a
        /// simple token; field values are quoted and escaped, null values omitted.
        /// </summary>
        public static string FormatLine(DateTime utcTimestamp, string operation, IReadOnlyList<(string Key, string? Value)> fields)
        {
            if (string.IsNullOrWhiteSpace(operation))
            {
                throw new ArgumentException("Operation must not be blank.", nameof(operation));
            }

            foreach (char c in operation)
            {
                if (!char.IsLetterOrDigit(c) && c != '_' && c != '-')
                {
                    throw new ArgumentException("Operation must be a simple token (letters, digits, '_', '-').", nameof(operation));
                }
            }

            if (fields == null)
            {
                throw new ArgumentNullException(nameof(fields));
            }

            StringBuilder sb = new StringBuilder(128);
            sb.Append("ts=");
            sb.Append(utcTimestamp.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", CultureInfo.InvariantCulture));
            sb.Append(" op=");
            sb.Append(operation);
            for (int i = 0; i < fields.Count; i++)
            {
                (string key, string? value) = fields[i];
                if (value == null)
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(key))
                {
                    throw new ArgumentException("Field keys must not be blank.", nameof(fields));
                }

                foreach (char c in key)
                {
                    if (!char.IsLetterOrDigit(c) && c != '_' && c != '-')
                    {
                        throw new ArgumentException(
                            "Field key '" + key + "' must be a simple token (letters, digits, '_', '-').", nameof(fields));
                    }
                }

                sb.Append(' ');
                sb.Append(key);
                sb.Append("=\"");
                AppendEscaped(sb, value);
                sb.Append('"');
            }

            return sb.ToString();
        }

        private static void AppendEscaped(StringBuilder sb, string value)
        {
            foreach (char c in value)
            {
                switch (c)
                {
                    case '\\':
                        sb.Append("\\\\");
                        break;
                    case '"':
                        sb.Append("\\\"");
                        break;
                    case '\r':
                        sb.Append("\\r");
                        break;
                    case '\n':
                        sb.Append("\\n");
                        break;
                    case '\t':
                        sb.Append("\\t");
                        break;
                    default:
                        sb.Append(c);
                        break;
                }
            }
        }
    }
}
