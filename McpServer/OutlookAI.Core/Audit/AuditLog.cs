using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using Microsoft.Win32.SafeHandles;

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
    ///   ts=2026-07-23T10:11:12.345Z op=new_draft key="value" key2="value2" pid="4242" run="0a1b2c3d" seq="17" crc="89abcdef"
    /// Values are quoted with backslash escapes for '\', '"', CR, LF and TAB; fields
    /// with null values are omitted. <see cref="Append"/> THROWS when the line cannot
    /// be written - a write operation without its audit line must be surfaced, never
    /// silently swallowed (D4 discipline).
    ///
    /// <para>
    /// THE INTEGRITY TRAILER (Q117, 2026-10-03). The writer ends every line with fields no caller
    /// may use (<see cref="TrailerKeys"/>): <c>pid</c>, the writing process; <c>run</c>, eight hex
    /// digits drawn once per process, because Windows recycles process ids and one Outlook process
    /// can host the writer twice (one copy per add-in AppDomain), so a pid alone names no writer;
    /// <c>seq</c>, that writer's line number IN THIS FILE, from 1 - a new file under the same name
    /// starts it again, so a writer's lines in one file are exactly 1..n and any number missing is
    /// a line lost; <c>lock</c>, only when the writer lock was NOT held normally (see
    /// <see cref="LockTimeout"/> and its siblings); and <c>crc</c>, last, the CRC-32 of every byte of
    /// the line before it, so a line cut short, overwritten or spliced is detected rather than
    /// believed. <see cref="AuditLogReader"/> checks all of it.
    /// </para>
    ///
    /// <para>
    /// HOW THE FILE IS OPENED, because a maintainer renaming the log has to know (Q86): every
    /// append opens the file, writes one line and closes it again - no handle is ever held
    /// between two lines. Since Q117 the open is <c>CreateFileW</c> with <c>FILE_APPEND_DATA</c> and
    /// no <c>FILE_WRITE_DATA</c>, so Windows itself puts every write at the end of the file as it
    /// runs and no writer can overwrite another's line, lock or no lock (see <see cref="AuditLogFile"/>
    /// for why <c>FileMode.Append</c> could). Sharing is read and write and deliberately NOT delete,
    /// so the file cannot be renamed or deleted under an append in progress; an open that meets a
    /// sharing violation is retried <see cref="WriteRetries"/> times, 15 ms apart, before the
    /// operation reports it. Each line, terminator included, goes out in ONE write. A named mutex
    /// still serializes the appends to one log across processes - now for ORDER only: the line's
    /// timestamp is taken while it is held, so the file is in time order. The mutex holds no file
    /// handle, so it changes nothing about renaming the file.
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
        /// other session's drafts and sends behind it. Appending without the lock costs ORDER only,
        /// never a line (see <see cref="AuditLogFile"/>), and the line says so: <see cref="LockTimeout"/>.
        /// </summary>
        internal const int WriterLockWaitMilliseconds = 2000;

        /// <summary>
        /// <c>FILE_FLAG_WRITE_THROUGH</c> on the append handle: an append returns once its line is on the
        /// disk, not merely in the cache, so a power cut right after a draft or send cannot take the line
        /// recording it. Measured on the build VM before it was switched on (see the stress tests'
        /// <c>WriteThrough</c> benchmark, and CHANGELOG / the Q117 report for the numbers).
        /// </summary>
        internal const bool WriteThrough = true;

        /// <summary>
        /// The trailer fields the writer appends to every line itself, in this order (<c>lock</c> only
        /// when it applies) - reserved, so a caller passing one is refused (<see cref="FormatLine"/>).
        /// </summary>
        internal static readonly string[] TrailerKeys = { PidKey, RunKey, SeqKey, LockKey, CrcKey };

        internal const string PidKey = "pid";
        internal const string RunKey = "run";
        internal const string SeqKey = "seq";
        internal const string LockKey = "lock";
        internal const string CrcKey = "crc";

        /// <summary><c>lock="timeout"</c>: another writer held the lock longer than <see cref="WriterLockWaitMilliseconds"/>.</summary>
        internal const string LockTimeout = "timeout";

        /// <summary><c>lock="unavailable"</c>: the named mutex could not be created or opened (e.g. another account's object of that name).</summary>
        internal const string LockUnavailable = "unavailable";

        /// <summary>
        /// <c>lock="abandoned"</c>: the lock WAS held, but its previous owner died holding it - mid-append,
        /// possibly, which is when the line before this one may be a fragment.
        /// </summary>
        internal const string LockAbandoned = "abandoned";

        /// <summary><c>lock="disabled"</c>: the lock was deliberately skipped - only the multi-process stress test does that.</summary>
        internal const string LockDisabled = "disabled";

        /// <summary>The <c>lock</c> values a line may carry; anything else is not something the writer wrote.</summary>
        internal static readonly string[] LockStates = { LockTimeout, LockUnavailable, LockAbandoned, LockDisabled };

        /// <summary>This process's id, as every line it writes records it.</summary>
        internal static readonly int WriterPid = AuditLogFile.CurrentProcessId();

        /// <summary>
        /// Eight hex digits drawn once per process (per AppDomain, on .NET Framework): what tells this
        /// writer apart from an earlier process that had the same pid, or a second copy of this code in
        /// the same process.
        /// </summary>
        internal static readonly string WriterRun = NewRunId();

        private static readonly object SequenceGate = new object();

        // Per log file this process writes: the identity of the file it last appended to there, and
        // the last seq it gave a line in it. Keyed by the case-folded full path.
        private static readonly Dictionary<string, KeyValuePair<string?, long>> Sequences =
            new Dictionary<string, KeyValuePair<string?, long>>(StringComparer.Ordinal);

        private static long _lockTimeouts;
        private static long _lockUnavailable;
        private static long _lockAbandoned;

        /// <summary>Appends by this process that waited out <see cref="WriterLockWaitMilliseconds"/> and wrote without the lock (outlook_health).</summary>
        internal static long WriterLockTimeouts => Interlocked.Read(ref _lockTimeouts);

        /// <summary>Appends by this process that could not create or open the lock at all (outlook_health).</summary>
        internal static long WriterLockUnavailable => Interlocked.Read(ref _lockUnavailable);

        /// <summary>Appends by this process that inherited a lock another writer died holding (outlook_health).</summary>
        internal static long WriterLockAbandoned => Interlocked.Read(ref _lockAbandoned);

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
            AppendTo(directory, operation, fields, useWriterLock: true);
        }

        /// <summary>
        /// <see cref="AppendTo(string, string, IReadOnlyList{ValueTuple{string, string}})"/>, with the
        /// writer lock optional. Only the multi-process stress test ever passes false - to measure what
        /// the file system does with appends that nothing orders - and every line it writes that way
        /// says so (<see cref="LockDisabled"/>).
        /// </summary>
        internal static void AppendTo(
            string directory, string operation, IReadOnlyList<(string Key, string? Value)> fields, bool useWriterLock)
        {
            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new ArgumentException("Audit directory must not be blank.", nameof(directory));
            }

            // Everything a caller can get wrong is refused before the lock is taken or a seq is used.
            string body = FormatBody(operation, fields);
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

                // The lock orders appends; it no longer protects them. The append handle is what makes
                // an overwrite impossible (AuditLogFile), so a writer that cannot get the lock in time
                // still appends - and records on the line that it did so.
                using (AuditWriterLock writerLock = useWriterLock ? AuditWriterLock.Acquire(path) : AuditWriterLock.Skipped())
                {
                    WriteOneLine(path, body, writerLock.LineState);
                }
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException
                || ex is NotSupportedException || ex is ArgumentException)
            {
                // E.g. the directory path is blocked by a same-named file. The retry
                // loop's own give-up exception is already an InvalidOperationException
                // and passes through untouched.
                throw new InvalidOperationException("Audit line could not be written to '" + path + "'.", ex);
            }
        }

        /// <summary>
        /// Opens the log, takes the timestamp and the next seq, and writes the line in one call - all of
        /// it under the writer lock when the caller holds it, which is what puts the file in time order.
        /// </summary>
        private static void WriteOneLine(string path, string body, string? lockState)
        {
            SafeFileHandle? handle = null;
            int lastError = 0;
            for (int attempt = 0; attempt < WriteRetries; attempt++)
            {
                if (AuditLogFile.TryOpenForAppend(path, WriteThrough, out handle, out lastError))
                {
                    break;
                }

                if (!AuditLogFile.IsSharingViolation(lastError))
                {
                    throw AuditLogFile.ErrorFor(lastError, path);
                }

                Thread.Sleep(15);
            }

            if (handle == null)
            {
                throw new InvalidOperationException(
                    "Audit line could not be written to '" + path + "' after " +
                    WriteRetries.ToString(CultureInfo.InvariantCulture) + " attempts.", AuditLogFile.ErrorFor(lastError, path));
            }

            using (handle)
            {
                string? identity = AuditLogFile.TryGetIdentity(handle);
                bool repairTail = AuditLogFile.EndsMidLine(handle);

                // From here a failure is a line this writer meant to write and did not: the seq it
                // took stays used, so the gap it leaves is how a reader sees the loss.
                long seq = NextSequence(path, identity);
                byte[] line = EncodeSealedLine(AuditLogFile.PreciseUtcNow(), body, WriterPid, WriterRun, seq, lockState, repairTail, out int count);
                AuditLogFile.Append(handle, line, count);
            }
        }

        /// <summary>
        /// The next <c>seq</c> for <paramref name="path"/>: one more than the last, or 1 when the file
        /// there is not the one this process last appended to (it was archived and a new one begun).
        /// </summary>
        private static long NextSequence(string path, string? identity)
        {
            string key = Path.GetFullPath(path).ToUpperInvariant();
            lock (SequenceGate)
            {
                long last = 0;
                if (Sequences.TryGetValue(key, out KeyValuePair<string?, long> state)
                    && string.Equals(state.Key, identity, StringComparison.Ordinal))
                {
                    last = state.Value;
                }

                long next = last + 1;
                Sequences[key] = new KeyValuePair<string?, long>(identity, next);
                return next;
            }
        }

        /// <summary>
        /// Probes whether an audit line COULD be appended under
        /// <paramref name="directory"/> without writing anything: creates the directory
        /// when missing and opens (creating if absent) the log file exactly the way an
        /// append opens it (<see cref="AuditLogFile.TryOpenForAppend"/>), then closes it.
        /// Content-free error reason on failure (S4). Used by the health tool - write ops
        /// fail-closed without audit. A redirected (test) process gets
        /// <see cref="RedirectRefusalError"/> for any directory it may not touch, before
        /// anything is created or opened.
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
                if (!AuditLogFile.TryOpenForAppend(path, WriteThrough, out SafeFileHandle? handle, out int win32Error))
                {
                    error = AuditLogFile.ErrorFor(win32Error, path).GetType().Name;
                    return false;
                }

                handle!.Dispose();
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
        /// The name of the mutex that orders appends to the log at <paramref name="path"/>: in the
        /// GLOBAL namespace (Q117), so writers in different Windows sessions - a scheduled task, a
        /// second logon, the add-in's Outlook and an agent's server - order against each other; one
        /// per log file - a hash of its full path, case-folded - so the real log and a test run's
        /// throwaway one never wait on each other. Before Q117 it was <c>Local\</c>, per session.
        /// </summary>
        internal static string WriterMutexName(string path)
        {
            string full = Path.GetFullPath(path).ToUpperInvariant();
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(full));
                return @"Global\OutlookAI.AuditLog." + BitConverter.ToString(hash).Replace("-", string.Empty);
            }
        }

        /// <summary>
        /// The one-writer-at-a-time lock around an append, which since Q117 buys ORDER, not safety:
        /// the append-only handle already keeps every line whole and in the file however many writers
        /// race. Best effort by design: a writer that cannot get it within
        /// <see cref="WriterLockWaitMilliseconds"/>, or cannot create it at all, appends anyway rather
        /// than fail the draft or send the line records - and <see cref="LineState"/> puts that on the
        /// line, and a counter in <c>outlook_health</c>. A previous holder that died mid-append leaves
        /// the mutex abandoned, which hands ownership over normally and is recorded the same way.
        /// Internal so the stress test can hold it the way a writer that dies mid-append does.
        /// </summary>
        internal sealed class AuditWriterLock : IDisposable
        {
            private readonly Mutex? _mutex;
            private readonly bool _owned;

            private AuditWriterLock(Mutex? mutex, bool owned, string? lineState)
            {
                _mutex = mutex;
                _owned = owned;
                LineState = lineState;
            }

            /// <summary>What the line written under this lock must say about it: null when it was held normally.</summary>
            internal string? LineState { get; }

            /// <summary>Whether this writer owns the mutex (held normally, or inherited from a writer that died).</summary>
            internal bool Owned => _owned;

            internal static AuditWriterLock Acquire(string path)
            {
                Mutex? mutex = null;
                try
                {
                    mutex = new Mutex(initiallyOwned: false, WriterMutexName(path));
                    try
                    {
                        if (mutex.WaitOne(WriterLockWaitMilliseconds))
                        {
                            return new AuditWriterLock(mutex, owned: true, lineState: null);
                        }

                        Interlocked.Increment(ref _lockTimeouts);
                        return new AuditWriterLock(mutex, owned: false, lineState: LockTimeout);
                    }
                    catch (AbandonedMutexException)
                    {
                        // Ownership passed to this thread; the writer before it died holding the lock.
                        Interlocked.Increment(ref _lockAbandoned);
                        return new AuditWriterLock(mutex, owned: true, lineState: LockAbandoned);
                    }
                }
                catch (Exception ex) when (ex is UnauthorizedAccessException || ex is IOException
                    || ex is WaitHandleCannotBeOpenedException || ex is ArgumentException)
                {
                    // E.g. a same-named object of another type, or one created by an account (or an
                    // elevated process) this one may not open. The append is still safe; only its
                    // order against other writers is not guaranteed, and the line says so.
                    mutex?.Dispose();
                    Interlocked.Increment(ref _lockUnavailable);
                    return new AuditWriterLock(null, owned: false, lineState: LockUnavailable);
                }
            }

            /// <summary>No lock at all - the stress test's lock-disabled run only.</summary>
            internal static AuditWriterLock Skipped()
            {
                return new AuditWriterLock(null, owned: false, lineState: LockDisabled);
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
        /// Formats one audit line WITHOUT the integrity trailer (pure logic, T1-tested) - the line as
        /// it was written before Q117, and the part of a sealed line the trailer follows. The
        /// operation name must be a simple token; field values are quoted and escaped, null values
        /// omitted, and the trailer's own keys (<see cref="TrailerKeys"/>) are refused.
        /// </summary>
        public static string FormatLine(DateTime utcTimestamp, string operation, IReadOnlyList<(string Key, string? Value)> fields)
        {
            return FormatTimestamp(utcTimestamp) + FormatBody(operation, fields);
        }

        /// <summary>
        /// A complete line as <see cref="AppendTo(string, string, IReadOnlyList{ValueTuple{string, string}})"/>
        /// writes it, trailer and checksum included, without its terminator - for tests that need to
        /// know the exact bytes.
        /// </summary>
        internal static string FormatSealedLine(
            DateTime utcTimestamp, string operation, IReadOnlyList<(string Key, string? Value)> fields, int pid, string run, long seq, string? lockState)
        {
            byte[] bytes = EncodeSealedLine(utcTimestamp, FormatBody(operation, fields), pid, run, seq, lockState, repairTail: false, out int count);
            return Encoding.UTF8.GetString(bytes, 0, count - 1);
        }

        /// <summary>
        /// The bytes of one sealed line: an optional leading line feed (closing a fragment a crash left
        /// at the end of the file), the line, its trailer, its checksum and its terminator.
        /// </summary>
        private static byte[] EncodeSealedLine(
            DateTime utcTimestamp, string body, int pid, string run, long seq, string? lockState, bool repairTail, out int count)
        {
            StringBuilder sb = new StringBuilder(body.Length + 96);
            sb.Append(FormatTimestamp(utcTimestamp));
            sb.Append(body);
            sb.Append(' ').Append(PidKey).Append("=\"").Append(pid.ToString(CultureInfo.InvariantCulture)).Append('"');
            sb.Append(' ').Append(RunKey).Append("=\"").Append(run).Append('"');
            sb.Append(' ').Append(SeqKey).Append("=\"").Append(seq.ToString(CultureInfo.InvariantCulture)).Append('"');
            if (lockState != null)
            {
                sb.Append(' ').Append(LockKey).Append("=\"").Append(lockState).Append('"');
            }

            UTF8Encoding utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
            string covered = sb.ToString();
            int lead = repairTail ? 1 : 0;
            int coveredBytes = utf8.GetByteCount(covered);

            // ' crc="' + 8 hex digits + '"' + '\n' - every byte of it ASCII.
            byte[] bytes = new byte[lead + coveredBytes + CrcFieldBytes + 1];
            if (repairTail)
            {
                bytes[0] = (byte)'\n';
            }

            utf8.GetBytes(covered, 0, covered.Length, bytes, lead);
            string crcField = " " + CrcKey + "=\"" + AuditCrc32.ToHex(AuditCrc32.Compute(bytes, lead, coveredBytes)) + "\"";
            for (int i = 0; i < CrcFieldBytes; i++)
            {
                bytes[lead + coveredBytes + i] = (byte)crcField[i];
            }

            bytes[bytes.Length - 1] = (byte)'\n';
            count = bytes.Length;
            return bytes;
        }

        /// <summary>The length of <c> crc="xxxxxxxx"</c>, the field every sealed line ends with.</summary>
        internal const int CrcFieldBytes = 15;

        private static string FormatTimestamp(DateTime utcTimestamp)
        {
            return "ts=" + utcTimestamp.ToString("yyyy-MM-dd'T'HH:mm:ss.fff'Z'", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// " op=..." plus the caller's fields: everything of a line between its timestamp and its
        /// trailer. Validates the operation and every key.
        /// </summary>
        private static string FormatBody(string operation, IReadOnlyList<(string Key, string? Value)> fields)
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

                if (Array.IndexOf(TrailerKeys, key) >= 0)
                {
                    throw new ArgumentException(
                        "Field key '" + key + "' is reserved: the writer appends it to every line itself.", nameof(fields));
                }

                sb.Append(' ');
                sb.Append(key);
                sb.Append("=\"");
                AppendEscaped(sb, value);
                sb.Append('"');
            }

            return sb.ToString();
        }

        private static string NewRunId()
        {
            byte[] bytes = new byte[4];
            using (RandomNumberGenerator rng = RandomNumberGenerator.Create())
            {
                rng.GetBytes(bytes);
            }

            return AuditCrc32.ToHex(BitConverter.ToUInt32(bytes, 0));
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
