using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;

namespace OutlookAI.Core.Audit
{
    /// <summary>One audit line, parsed back out of the grammar <see cref="AuditLog.FormatLine"/> writes.</summary>
    public sealed class AuditLogEntry
    {
        internal AuditLogEntry(DateTime timestampUtc, string operation, IReadOnlyList<KeyValuePair<string, string>> fields)
            : this(timestampUtc, operation, fields, null, null, null, null)
        {
        }

        internal AuditLogEntry(
            DateTime timestampUtc,
            string operation,
            IReadOnlyList<KeyValuePair<string, string>> fields,
            int? pid,
            string? run,
            long? seq,
            string? writerLock)
        {
            TimestampUtc = timestampUtc;
            Operation = operation;
            Fields = fields;
            Pid = pid;
            Run = run;
            Seq = seq;
            WriterLock = writerLock;
        }

        /// <summary>The line's <c>ts</c>, UTC, to the millisecond the writer recorded.</summary>
        public DateTime TimestampUtc { get; }

        /// <summary>The line's <c>op</c> token.</summary>
        public string Operation { get; }

        /// <summary>
        /// The line's fields in line order, values unescaped - the caller's fields only: the integrity
        /// trailer (<see cref="AuditLog.TrailerKeys"/>) is split off into the properties below. A key
        /// could in principle appear twice - <see cref="AuditLog.FormatLine"/> does not forbid it,
        /// though no call site does it - so this stays a list rather than a dictionary.
        /// </summary>
        public IReadOnlyList<KeyValuePair<string, string>> Fields { get; }

        /// <summary>The writing process (<c>pid</c>), or null on a line written before lines carried a trailer.</summary>
        public int? Pid { get; }

        /// <summary>The writer's run id (<c>run</c>), or null on a line without a trailer.</summary>
        public string? Run { get; }

        /// <summary>The writer's line number in this file (<c>seq</c>, from 1), or null on a line without a trailer.</summary>
        public long? Seq { get; }

        /// <summary>The line's <c>lock</c> value - null when the writer lock was held normally, or there is no trailer.</summary>
        public string? WriterLock { get; }

        /// <summary>
        /// True when the line carried a checksum and it matched. A line whose checksum does NOT match is
        /// never returned at all; false here means a line from before checksums, which cannot be verified.
        /// </summary>
        public bool Checksummed => Run != null;
    }

    /// <summary>
    /// What to select from the log. Matching only: the caller validates the arguments, because
    /// what counts as a usable EntryID or operation name is a judgement of the tool's, not of the
    /// file format's.
    /// </summary>
    public sealed class AuditLogFilter
    {
        private readonly IReadOnlyList<string> _operationPatterns;

        /// <summary>A filter; every argument is optional and the ones given are ANDed.</summary>
        /// <param name="afterUtc">Only lines at or after this instant.</param>
        /// <param name="beforeUtc">Only lines strictly before this instant.</param>
        /// <param name="operationPatterns">
        /// Operation names, any of which may match, compared case-insensitively; one ending in
        /// <c>*</c> matches every operation starting with what precedes it. Null or empty for all.
        /// </param>
        /// <param name="entryId">
        /// Only lines naming this EntryID, case-insensitively, in a field called <c>entryId</c> or
        /// ending in <c>EntryId</c> (<c>newEntryId</c>, <c>sourceEntryId</c>).
        /// </param>
        public AuditLogFilter(DateTime? afterUtc, DateTime? beforeUtc, IReadOnlyList<string>? operationPatterns, string? entryId)
        {
            AfterUtc = afterUtc;
            BeforeUtc = beforeUtc;
            _operationPatterns = operationPatterns ?? Array.Empty<string>();
            EntryId = entryId;
        }

        /// <summary>Inclusive lower bound, or null.</summary>
        public DateTime? AfterUtc { get; }

        /// <summary>Exclusive upper bound, or null.</summary>
        public DateTime? BeforeUtc { get; }

        /// <summary>The EntryID every selected line must name, or null.</summary>
        public string? EntryId { get; }

        /// <summary>Whether <paramref name="entry"/> is selected.</summary>
        public bool Matches(AuditLogEntry entry)
        {
            if (AfterUtc.HasValue && entry.TimestampUtc < AfterUtc.Value)
            {
                return false;
            }

            if (BeforeUtc.HasValue && entry.TimestampUtc >= BeforeUtc.Value)
            {
                return false;
            }

            if (_operationPatterns.Count > 0 && !MatchesAnyOperation(entry.Operation))
            {
                return false;
            }

            return EntryId == null || NamesEntryId(entry, EntryId);
        }

        /// <summary>Whether a field key is one that holds an EntryID.</summary>
        internal static bool IsEntryIdKey(string key)
        {
            return string.Equals(key, "entryId", StringComparison.Ordinal)
                || key.EndsWith("EntryId", StringComparison.Ordinal);
        }

        private bool MatchesAnyOperation(string operation)
        {
            foreach (string pattern in _operationPatterns)
            {
                bool matched = pattern.EndsWith("*", StringComparison.Ordinal)
                    ? operation.StartsWith(pattern.Substring(0, pattern.Length - 1), StringComparison.OrdinalIgnoreCase)
                    : string.Equals(operation, pattern, StringComparison.OrdinalIgnoreCase);
                if (matched)
                {
                    return true;
                }
            }

            return false;
        }

        private static bool NamesEntryId(AuditLogEntry entry, string entryId)
        {
            foreach (KeyValuePair<string, string> field in entry.Fields)
            {
                if (IsEntryIdKey(field.Key) && string.Equals(field.Value, entryId, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }
    }

    /// <summary>
    /// Where the next older page of a newest-first read begins (Q119): every line that STARTS before
    /// <see cref="Offset"/> in the file whose identity is <see cref="FileIdentity"/>, where the line that
    /// starts at <see cref="Offset"/> is the one whose bytes have CRC-32 <see cref="Anchor"/>. A byte
    /// position, not a timestamp, so lines that share a millisecond are neither skipped nor repeated;
    /// and checked against the file twice over, so a position into a log that was archived, replaced,
    /// truncated or edited since is refused instead of read as if nothing had happened.
    /// </summary>
    public sealed class AuditLogPosition
    {
        /// <summary>A position.</summary>
        public AuditLogPosition(string? fileIdentity, long offset, uint anchor)
        {
            FileIdentity = fileIdentity;
            Offset = offset;
            Anchor = anchor;
        }

        /// <summary>The log file's volume and file id, or null when the file system would not say.</summary>
        public string? FileIdentity { get; }

        /// <summary>The byte offset of the oldest line the previous page returned.</summary>
        public long Offset { get; }

        /// <summary>The CRC-32 of that line's bytes, its terminator excluded.</summary>
        public uint Anchor { get; }
    }

    /// <summary>Whether a read that was asked to continue from a position could.</summary>
    public enum AuditLogResumeStatus
    {
        /// <summary>No position was given: the read started at the newest line.</summary>
        NotResumed,

        /// <summary>The position was found where it was left and the read continued from it.</summary>
        Resumed,

        /// <summary>There is no live log any more - it was archived and nothing has been written since.</summary>
        FileMissing,

        /// <summary>The live log is another file than the one the position points into: it was archived or replaced.</summary>
        FileReplaced,

        /// <summary>The file no longer has the position's line where it was: truncated, edited or rewritten in place.</summary>
        PositionNotFound,
    }

    /// <summary>A run of one writer's lines that are not in the file (Q117): <c>seq</c> numbers it skipped.</summary>
    public sealed class AuditLogGap
    {
        internal AuditLogGap(int pid, string run, long firstMissing, long lastMissing)
        {
            Pid = pid;
            Run = run;
            FirstMissing = firstMissing;
            LastMissing = lastMissing;
        }

        /// <summary>The writer's process id.</summary>
        public int Pid { get; }

        /// <summary>The writer's run id.</summary>
        public string Run { get; }

        /// <summary>The first missing <c>seq</c>.</summary>
        public long FirstMissing { get; }

        /// <summary>The last missing <c>seq</c>.</summary>
        public long LastMissing { get; }

        /// <summary>How many lines are missing in this run.</summary>
        public long Count => LastMissing - FirstMissing + 1;
    }

    /// <summary>What one read of the log found.</summary>
    public sealed class AuditLogScan
    {
        internal AuditLogScan()
        {
        }

        internal AuditLogScan(
            bool fileFound,
            IReadOnlyList<AuditLogEntry> entries,
            int matched,
            long linesScanned,
            long malformedLines,
            bool incompleteLastLine,
            IReadOnlyList<string> operationsSeen)
        {
            FileFound = fileFound;
            Entries = entries;
            Matched = matched;
            LinesScanned = linesScanned;
            MalformedLines = malformedLines;
            IncompleteLastLine = incompleteLastLine;
            OperationsSeen = operationsSeen;
        }

        /// <summary>False when there is no log file yet (nothing recorded since it was started or renamed).</summary>
        public bool FileFound { get; internal set; }

        /// <summary>The page: the newest matching entries before the resume position (or of the whole log), newest first.</summary>
        public IReadOnlyList<AuditLogEntry> Entries { get; internal set; } = Array.Empty<AuditLogEntry>();

        /// <summary>How many entries matched in the whole log as it was read, every page included.</summary>
        public int Matched { get; internal set; }

        /// <summary>Matching entries OLDER than the oldest one on this page - what the next pages hold.</summary>
        public long OlderMatches { get; internal set; }

        /// <summary>Where the next older page starts; null when there is none.</summary>
        public AuditLogPosition? NextPosition { get; internal set; }

        /// <summary>Whether a requested resume position could be used.</summary>
        public AuditLogResumeStatus Resume { get; internal set; }

        /// <summary>The log file's identity as read, or null when there is no file or the file system would not say.</summary>
        public string? FileIdentity { get; internal set; }

        /// <summary>Complete lines read, malformed ones included.</summary>
        public long LinesScanned { get; internal set; }

        /// <summary>
        /// Complete lines that are not lines the writer wrote - they do not parse, or their checksum does
        /// not match - and were therefore not returned. Includes <see cref="ChecksumFailures"/>.
        /// </summary>
        public long MalformedLines { get; internal set; }

        /// <summary>
        /// Lines in the writer's format whose <c>crc</c> does not match their bytes: cut short and glued
        /// to a later line, overwritten, or edited. Counted in <see cref="MalformedLines"/> too.
        /// </summary>
        public long ChecksumFailures { get; internal set; }

        /// <summary>Lines whose checksum was verified.</summary>
        public long VerifiedLines { get; internal set; }

        /// <summary>Well-formed lines from before lines carried a trailer: read, but they cannot be verified.</summary>
        public long UnverifiedLines { get; internal set; }

        /// <summary>Verified lines written without the writer lock (<c>lock</c> timeout, unavailable or disabled).</summary>
        public long LinesWithoutWriterLock { get; internal set; }

        /// <summary>Verified lines whose writer inherited the lock from a writer that died holding it.</summary>
        public long LinesAfterAbandonedLock { get; internal set; }

        /// <summary>Distinct writers (pid and run) among the verified lines.</summary>
        public int Writers { get; internal set; }

        /// <summary>
        /// Lines missing from the file by the writers' own count: every <c>seq</c> number below a writer's
        /// highest that no verified line carries. Each writer numbers its lines in a file from 1.
        /// </summary>
        public long MissingLines { get; internal set; }

        /// <summary>The missing runs, ordered by writer then seq - at most <see cref="AuditLogReader.GapsCap"/>.</summary>
        public IReadOnlyList<AuditLogGap> Gaps { get; internal set; } = Array.Empty<AuditLogGap>();

        /// <summary>Verified lines whose writer and seq another verified line already had: a line copied or replayed.</summary>
        public long DuplicateLines { get; internal set; }

        /// <summary>
        /// True when the file ended without a line terminator: a line still being written when the
        /// read began, or one a crash cut off. It is not parsed and not counted in
        /// <see cref="LinesScanned"/>.
        /// </summary>
        public bool IncompleteLastLine { get; internal set; }

        /// <summary>
        /// The distinct operation names of every parsed line, filter or not, sorted - at most
        /// <see cref="AuditLogReader.OperationsSeenCap"/>. Read off the data rather than kept as a
        /// list of what the server writes, so a reply to a filter that matched nothing can say
        /// what the log does hold without a second copy of the operation names to drift.
        /// </summary>
        public IReadOnlyList<string> OperationsSeen { get; internal set; } = Array.Empty<string>();

        internal static AuditLogScan Refused(AuditLogResumeStatus status, bool fileFound, string? fileIdentity)
        {
            return new AuditLogScan { FileFound = fileFound, Resume = status, FileIdentity = fileIdentity };
        }
    }

    /// <summary>
    /// Reads the audit log back (Q93): the parser for the line grammar <see cref="AuditLog.FormatLine"/>
    /// writes, and a scan that never gets in the way of the writers.
    /// <para>
    /// <b>Sharing.</b> The file is opened <c>FileShare.ReadWrite | FileShare.Delete</c>, read-only.
    /// The product's appends open it for writing with read and write sharing, which this reader's
    /// handle allows, and this reader asks for no access an appender's share mode excludes - so a
    /// read never blocks an append, which matters because a draft or send whose audit line cannot
    /// be written is reported as a failure. <c>FileShare.Delete</c> is there so a read in progress
    /// cannot stop the maintainer renaming the log either.
    /// </para>
    /// <para>
    /// <b>A file that grows while it is read.</b> The length is taken once, when the file is
    /// opened, and nothing past it is read: lines appended later are for the next call, and the
    /// answer is a consistent prefix of the log. A last line without its terminator - an append
    /// in flight at that instant, or one a crash cut short - is reported, not parsed.
    /// </para>
    /// <para>
    /// <b>Strict.</b> <see cref="ParseLine"/> accepts exactly what <see cref="AuditLog.FormatLine"/>
    /// and the writer's trailer can produce and nothing else: an unknown escape, a raw tab or
    /// carriage return, a missing quote, a blank line, invalid UTF-8, a trailer out of order or a
    /// checksum that does not match is malformed, counted and skipped, never guessed at. That
    /// includes the unstructured lines OutlookAI wrote before 2026-07-23.
    /// </para>
    /// <para>
    /// <b>Verified (Q117).</b> Every line written since Q117 carries <c>pid</c>, <c>run</c>,
    /// <c>seq</c> and a <c>crc</c> of its own bytes. A line whose checksum fails is counted, never
    /// returned; the writers' <c>seq</c> numbers are checked for gaps per writer, so a line that
    /// never reached the file is reported as missing rather than simply absent.
    /// </para>
    /// <para>
    /// <b>Paged by position (Q119).</b> A read returns the newest matching lines that START before a
    /// position - the end of the file, or the <see cref="AuditLogPosition"/> the previous page handed
    /// out - and the position of the next page. Byte offsets do not tie the way timestamps do.
    /// </para>
    /// </summary>
    public static class AuditLogReader
    {
        /// <summary>A line longer than this is malformed by definition; the longest real line is a few hundred bytes.</summary>
        public const int MaxLineBytes = 256 * 1024;

        /// <summary>How many distinct operation names a scan collects for <see cref="AuditLogScan.OperationsSeen"/>.</summary>
        public const int OperationsSeenCap = 40;

        /// <summary>How many missing runs a scan lists in <see cref="AuditLogScan.Gaps"/>; <see cref="AuditLogScan.MissingLines"/> counts them all.</summary>
        public const int GapsCap = 50;

        private const string TimestampFormat = "yyyy-MM-dd'T'HH:mm:ss.fff'Z'";
        private const int TimestampLength = 24;
        private const int OpenRetries = 3;
        private const int ChunkBytes = 64 * 1024;

        private static readonly UTF8Encoding StrictUtf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

        private enum LineFault
        {
            None,
            Malformed,
            ChecksumMismatch,
        }

        /// <summary>
        /// Parses one line (without its terminator), or returns null when it is not exactly a line
        /// the writer could have written - a sealed line whose checksum does not match included.
        /// </summary>
        public static AuditLogEntry? ParseLine(string line)
        {
            return Parse(line, null, 0, 0, out _);
        }

        /// <summary>
        /// Reads the log at <paramref name="path"/>: the <paramref name="top"/> newest entries
        /// <paramref name="filter"/> selects, newest first, plus what the read could not use. A
        /// missing file (or directory) is an empty log, not an error.
        /// </summary>
        /// <exception cref="InvalidOperationException">
        /// The file exists but could not be read - or this is a test process and the file is outside
        /// the temp directory, which it may not read (see <see cref="AuditLog.DescribeRefusalWhileRedirected"/>).
        /// </exception>
        public static AuditLogScan Read(string path, AuditLogFilter filter, int top)
        {
            return Read(path, filter, top, null);
        }

        /// <summary>
        /// <see cref="Read(string, AuditLogFilter, int)"/>, continuing from <paramref name="before"/>
        /// when it is given: only lines that start before it are paged, and when the position no
        /// longer fits the file the scan comes back with <see cref="AuditLogScan.Resume"/> saying why
        /// and no entries.
        /// </summary>
        public static AuditLogScan Read(string path, AuditLogFilter filter, int top, AuditLogPosition? before)
        {
            // The read half of the Q86 tripwire: a redirected process - only ever a test - may no
            // more READ the real log than write it. Checked before anything is opened, so code that
            // forgot the redirect fails here instead of scanning the maintainer's own log.
            string? refusal = AuditLog.DescribeRefusalWhileRedirected(Path.GetDirectoryName(Path.GetFullPath(path)) ?? path);
            if (refusal != null)
            {
                throw new InvalidOperationException("The audit log at '" + path + "' was not read. " + refusal);
            }

            FileStream? stream = OpenShared(path);
            if (stream == null)
            {
                return before == null
                    ? new AuditLogScan(false, Array.Empty<AuditLogEntry>(), 0, 0, 0, false, Array.Empty<string>())
                    : AuditLogScan.Refused(AuditLogResumeStatus.FileMissing, fileFound: false, fileIdentity: null);
            }

            try
            {
                using (stream)
                {
                    string? identity = AuditLogFile.TryGetIdentity(stream.SafeFileHandle);
                    if (before != null && before.FileIdentity != null && identity != null
                        && !string.Equals(before.FileIdentity, identity, StringComparison.Ordinal))
                    {
                        return AuditLogScan.Refused(AuditLogResumeStatus.FileReplaced, fileFound: true, fileIdentity: identity);
                    }

                    return Scan(stream, stream.Length, filter, top, before, identity);
                }
            }
            catch (IOException ex)
            {
                throw new InvalidOperationException("The audit log at '" + path + "' could not be read: " + ex.Message, ex);
            }
        }

        /// <summary>
        /// Opens the log for reading with the sharing every writer allows (see the class remarks),
        /// retrying a sharing violation the way the writer does; null when there is no file.
        /// </summary>
        internal static FileStream? OpenShared(string path)
        {
            IOException? last = null;
            for (int attempt = 0; attempt < OpenRetries; attempt++)
            {
                try
                {
                    return new FileStream(
                        path,
                        FileMode.Open,
                        FileAccess.Read,
                        FileShare.ReadWrite | FileShare.Delete,
                        4096,
                        FileOptions.SequentialScan);
                }
                catch (FileNotFoundException)
                {
                    return null;
                }
                catch (DirectoryNotFoundException)
                {
                    return null;
                }
                catch (IOException ex)
                {
                    last = ex;
                    Thread.Sleep(15);
                }
                catch (UnauthorizedAccessException ex)
                {
                    throw new InvalidOperationException("The audit log at '" + path + "' could not be opened for reading: " + ex.Message, ex);
                }
            }

            throw new InvalidOperationException(
                "The audit log at '" + path + "' could not be opened for reading after "
                + OpenRetries.ToString(CultureInfo.InvariantCulture) + " attempts: " + (last?.Message ?? "unknown"), last);
        }

        /// <summary>
        /// The scan itself, over the first <paramref name="length"/> bytes of
        /// <paramref name="stream"/> - the snapshot. Internal so T1 can grow a file between taking
        /// the snapshot and reading it.
        /// </summary>
        internal static AuditLogScan Scan(Stream stream, long length, AuditLogFilter filter, int top)
        {
            return Scan(stream, length, filter, top, null, null);
        }

        /// <summary>
        /// The scan, paging from <paramref name="before"/> when it is given. Every line of the snapshot
        /// is read whatever the page - the counts, the checksums and the gap check cover the whole log -
        /// but only lines that start before the position can be on the page.
        /// </summary>
        internal static AuditLogScan Scan(
            Stream stream, long length, AuditLogFilter filter, int top, AuditLogPosition? before, string? fileIdentity)
        {
            if (top < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(top), "top must be at least 1.");
            }

            long limit = length;
            bool anchorFound = before == null;
            if (before != null)
            {
                if (before.Offset < 0 || before.Offset >= length)
                {
                    // The file is shorter than the position: it was truncated or replaced since.
                    return AuditLogScan.Refused(AuditLogResumeStatus.PositionNotFound, fileFound: true, fileIdentity: fileIdentity);
                }

                limit = before.Offset;
            }

            Queue<PageCandidate> newest = new Queue<PageCandidate>(top + 1);
            SortedSet<string> operationsSeen = new SortedSet<string>(StringComparer.Ordinal);
            WriterSequences writers = new WriterSequences();
            AuditLogScan scan = new AuditLogScan { FileFound = true, FileIdentity = fileIdentity };
            int matched = 0;
            long matchedBeforeLimit = 0;
            long lines = 0;

            byte[] chunk = new byte[ChunkBytes];
            MemoryStream pending = new MemoryStream();
            bool pendingTooLong = false;
            long remaining = length;
            long chunkStart = 0;
            long lineStartInFile = 0;

            while (remaining > 0)
            {
                int read = stream.Read(chunk, 0, (int)Math.Min(chunk.Length, remaining));
                if (read <= 0)
                {
                    // Shorter than it was a moment ago - truncated under us. What was read stands.
                    break;
                }

                remaining -= read;
                int lineStart = 0;
                for (int i = 0; i < read; i++)
                {
                    if (chunk[i] != (byte)'\n')
                    {
                        continue;
                    }

                    byte[] lineBytes;
                    int lineOffset;
                    int lineCount;
                    if (pending.Length == 0 && !pendingTooLong)
                    {
                        lineBytes = chunk;
                        lineOffset = lineStart;
                        lineCount = i - lineStart;
                    }
                    else
                    {
                        AppendPending(pending, chunk, lineStart, i - lineStart, ref pendingTooLong);
                        lineBytes = pending.GetBuffer();
                        lineOffset = 0;
                        lineCount = (int)pending.Length;
                    }

                    lines++;
                    bool tooLong = pendingTooLong || lineCount > MaxLineBytes;
                    if (before != null && lineStartInFile == before.Offset && !tooLong)
                    {
                        anchorFound = AuditCrc32.Compute(lineBytes, lineOffset, lineCount) == before.Anchor;
                    }

                    LineFault fault = LineFault.Malformed;
                    AuditLogEntry? entry = tooLong ? null : DecodeAndParse(lineBytes, lineOffset, lineCount, lines == 1, out fault);
                    if (entry == null)
                    {
                        scan.MalformedLines++;
                        if (fault == LineFault.ChecksumMismatch)
                        {
                            scan.ChecksumFailures++;
                        }
                    }
                    else
                    {
                        Account(scan, writers, entry);
                        if (operationsSeen.Count < OperationsSeenCap)
                        {
                            operationsSeen.Add(entry.Operation);
                        }

                        if (filter.Matches(entry))
                        {
                            matched++;
                            if (lineStartInFile < limit)
                            {
                                matchedBeforeLimit++;
                                newest.Enqueue(new PageCandidate(
                                    entry, lineStartInFile, AuditCrc32.Compute(lineBytes, lineOffset, lineCount)));
                                if (newest.Count > top)
                                {
                                    newest.Dequeue();
                                }
                            }
                        }
                    }

                    pending.SetLength(0);
                    pendingTooLong = false;
                    lineStart = i + 1;
                    lineStartInFile = chunkStart + i + 1;
                }

                AppendPending(pending, chunk, lineStart, read - lineStart, ref pendingTooLong);
                chunkStart += read;
            }

            scan.LinesScanned = lines;
            scan.IncompleteLastLine = pending.Length > 0 || pendingTooLong;
            scan.OperationsSeen = new List<string>(operationsSeen);
            writers.Summarize(scan);
            if (!anchorFound)
            {
                // No line starts at the position any more, or a different one does: the file was
                // rewritten under the position. Nothing is paged from a position that cannot be trusted.
                AuditLogScan refused = AuditLogScan.Refused(AuditLogResumeStatus.PositionNotFound, fileFound: true, fileIdentity: fileIdentity);
                return refused;
            }

            PageCandidate[] page = newest.ToArray();
            Array.Reverse(page);
            AuditLogEntry[] entries = new AuditLogEntry[page.Length];
            for (int i = 0; i < page.Length; i++)
            {
                entries[i] = page[i].Entry;
            }

            scan.Entries = entries;
            scan.Matched = matched;
            scan.OlderMatches = matchedBeforeLimit - page.Length;
            scan.Resume = before == null ? AuditLogResumeStatus.NotResumed : AuditLogResumeStatus.Resumed;
            if (scan.OlderMatches > 0)
            {
                PageCandidate oldest = page[page.Length - 1];
                scan.NextPosition = new AuditLogPosition(fileIdentity, oldest.Offset, oldest.Anchor);
            }

            return scan;
        }

        private static void Account(AuditLogScan scan, WriterSequences writers, AuditLogEntry entry)
        {
            if (!entry.Checksummed)
            {
                scan.UnverifiedLines++;
                return;
            }

            scan.VerifiedLines++;
            if (string.Equals(entry.WriterLock, AuditLog.LockAbandoned, StringComparison.Ordinal))
            {
                scan.LinesAfterAbandonedLock++;
            }
            else if (entry.WriterLock != null)
            {
                scan.LinesWithoutWriterLock++;
            }

            writers.Add(entry.Pid!.Value, entry.Run!, entry.Seq!.Value);
        }

        private static void AppendPending(MemoryStream pending, byte[] chunk, int offset, int count, ref bool tooLong)
        {
            if (count <= 0 || tooLong)
            {
                return;
            }

            if (pending.Length + count > MaxLineBytes + 1)
            {
                // Stop buffering a line that is already too long to be real; it is counted as
                // malformed when its terminator arrives, without holding it in memory.
                tooLong = true;
                pending.SetLength(0);
                return;
            }

            pending.Write(chunk, offset, count);
        }

        private static AuditLogEntry? DecodeAndParse(byte[] bytes, int offset, int count, bool firstLine, out LineFault fault)
        {
            fault = LineFault.Malformed;
            if (count > 0 && bytes[offset + count - 1] == (byte)'\r')
            {
                count--;
            }

            // The writer never emits a byte order mark; a file saved by an editor may start with one.
            if (firstLine && count >= 3 && bytes[offset] == 0xEF && bytes[offset + 1] == 0xBB && bytes[offset + 2] == 0xBF)
            {
                offset += 3;
                count -= 3;
            }

            if (count == 0)
            {
                return null;
            }

            string line;
            try
            {
                line = StrictUtf8.GetString(bytes, offset, count);
            }
            catch (ArgumentException)
            {
                // DecoderFallbackException: not valid UTF-8, so not something the writer wrote.
                return null;
            }

            return Parse(line, bytes, offset, count, out fault);
        }

        /// <summary>
        /// The grammar, then the trailer. <paramref name="raw"/> is the line's own bytes when the caller
        /// has them, so the checksum is taken over exactly what is in the file; without them the line is
        /// re-encoded, which gives the same bytes for anything strict UTF-8 decoding accepted.
        /// </summary>
        private static AuditLogEntry? Parse(string line, byte[]? raw, int rawOffset, int rawCount, out LineFault fault)
        {
            fault = LineFault.Malformed;

            // "ts=" + 24-character timestamp + " op=" + at least one token character.
            if (line == null || line.Length < 3 + TimestampLength + 4 + 1
                || string.CompareOrdinal(line, 0, "ts=", 0, 3) != 0
                || string.CompareOrdinal(line, 3 + TimestampLength, " op=", 0, 4) != 0)
            {
                return null;
            }

            if (!DateTime.TryParseExact(
                    line.Substring(3, TimestampLength),
                    TimestampFormat,
                    CultureInfo.InvariantCulture,
                    DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                    out DateTime timestamp))
            {
                return null;
            }

            int i = 3 + TimestampLength + 4;
            int operationStart = i;
            while (i < line.Length && IsTokenChar(line[i]))
            {
                i++;
            }

            if (i == operationStart)
            {
                return null;
            }

            string operation = line.Substring(operationStart, i - operationStart);
            List<KeyValuePair<string, string>> fields = new List<KeyValuePair<string, string>>();
            StringBuilder value = new StringBuilder();
            while (i < line.Length)
            {
                // Exactly one space before every field, and none after the last.
                if (line[i] != ' ')
                {
                    return null;
                }

                i++;
                int keyStart = i;
                while (i < line.Length && IsTokenChar(line[i]))
                {
                    i++;
                }

                if (i == keyStart || i + 1 >= line.Length || line[i] != '=' || line[i + 1] != '"')
                {
                    return null;
                }

                string key = line.Substring(keyStart, i - keyStart);
                i += 2;
                value.Length = 0;
                bool closed = false;
                while (i < line.Length)
                {
                    char c = line[i];
                    if (c == '"')
                    {
                        closed = true;
                        i++;
                        break;
                    }

                    if (c == '\\')
                    {
                        if (i + 1 >= line.Length)
                        {
                            return null;
                        }

                        switch (line[i + 1])
                        {
                            case '\\':
                                value.Append('\\');
                                break;
                            case '"':
                                value.Append('"');
                                break;
                            case 'r':
                                value.Append('\r');
                                break;
                            case 'n':
                                value.Append('\n');
                                break;
                            case 't':
                                value.Append('\t');
                                break;
                            default:
                                return null;
                        }

                        i += 2;
                        continue;
                    }

                    // The writer escapes these three, so a raw one is not something it wrote.
                    if (c == '\r' || c == '\n' || c == '\t')
                    {
                        return null;
                    }

                    value.Append(c);
                    i++;
                }

                if (!closed)
                {
                    return null;
                }

                fields.Add(new KeyValuePair<string, string>(key, value.ToString()));
            }

            return ParseTrailer(line, raw, rawOffset, rawCount, timestamp, operation, fields, out fault);
        }

        /// <summary>
        /// Splits the integrity trailer off a parsed line and checks it: none of its keys at all is a
        /// line from before Q117; otherwise the caller's fields must be followed by exactly
        /// <c>pid</c>, <c>run</c>, <c>seq</c>, an optional <c>lock</c> and <c>crc</c>, each in the
        /// one form the writer produces, and the checksum must match.
        /// </summary>
        private static AuditLogEntry? ParseTrailer(
            string line,
            byte[]? raw,
            int rawOffset,
            int rawCount,
            DateTime timestamp,
            string operation,
            List<KeyValuePair<string, string>> fields,
            out LineFault fault)
        {
            fault = LineFault.Malformed;
            int n = fields.Count;
            bool sealedLine = false;
            for (int k = 0; k < n && !sealedLine; k++)
            {
                sealedLine = IsTrailerKey(fields[k].Key);
            }

            if (!sealedLine)
            {
                fault = LineFault.None;
                return new AuditLogEntry(timestamp, operation, fields);
            }

            if (!string.Equals(fields[n - 1].Key, AuditLog.CrcKey, StringComparison.Ordinal))
            {
                return null;
            }

            int at = n - 2;
            string? writerLock = null;
            if (at >= 0 && string.Equals(fields[at].Key, AuditLog.LockKey, StringComparison.Ordinal))
            {
                writerLock = fields[at].Value;
                at--;
            }

            if (at < 2
                || !string.Equals(fields[at].Key, AuditLog.SeqKey, StringComparison.Ordinal)
                || !string.Equals(fields[at - 1].Key, AuditLog.RunKey, StringComparison.Ordinal)
                || !string.Equals(fields[at - 2].Key, AuditLog.PidKey, StringComparison.Ordinal))
            {
                return null;
            }

            int callerFields = at - 2;
            for (int k = 0; k < callerFields; k++)
            {
                if (IsTrailerKey(fields[k].Key))
                {
                    return null;
                }
            }

            if (!TryParseCanonicalPositive(fields[at - 2].Value, 10, out long pid) || pid > int.MaxValue
                || !AuditCrc32.TryParseHex(fields[at - 1].Value, out _)
                || !TryParseCanonicalPositive(fields[at].Value, 19, out long seq)
                || (writerLock != null && Array.IndexOf(AuditLog.LockStates, writerLock) < 0)
                || !AuditCrc32.TryParseHex(fields[n - 1].Value, out uint stated))
            {
                return null;
            }

            // The checksum covers every byte before " crc=": a fixed fifteen ASCII bytes at the end.
            uint actual;
            if (raw != null)
            {
                actual = AuditCrc32.Compute(raw, rawOffset, rawCount - AuditLog.CrcFieldBytes);
            }
            else
            {
                byte[] covered = new UTF8Encoding(false).GetBytes(line.Substring(0, line.Length - AuditLog.CrcFieldBytes));
                actual = AuditCrc32.Compute(covered, 0, covered.Length);
            }

            if (actual != stated)
            {
                fault = LineFault.ChecksumMismatch;
                return null;
            }

            string run = fields[at - 1].Value;
            fields.RemoveRange(callerFields, n - callerFields);
            fault = LineFault.None;
            return new AuditLogEntry(timestamp, operation, fields, (int)pid, run, seq, writerLock);
        }

        private static bool IsTrailerKey(string key)
        {
            return Array.IndexOf(AuditLog.TrailerKeys, key) >= 0;
        }

        /// <summary>Digits only, no leading zero, at most <paramref name="maxDigits"/> - the one way the writer prints a positive number.</summary>
        private static bool TryParseCanonicalPositive(string text, int maxDigits, out long value)
        {
            value = 0;
            if (string.IsNullOrEmpty(text) || text.Length > maxDigits || text[0] == '0')
            {
                return false;
            }

            foreach (char c in text)
            {
                if (c < '0' || c > '9')
                {
                    return false;
                }
            }

            return long.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out value) && value > 0;
        }

        private static bool IsTokenChar(char c)
        {
            // The writer's own rule for operation names and field keys (AuditLog.FormatLine).
            return char.IsLetterOrDigit(c) || c == '_' || c == '-';
        }

        private struct PageCandidate
        {
            internal PageCandidate(AuditLogEntry entry, long offset, uint anchor)
            {
                Entry = entry;
                Offset = offset;
                Anchor = anchor;
            }

            internal AuditLogEntry Entry { get; }

            internal long Offset { get; }

            internal uint Anchor { get; }
        }

        /// <summary>Every verified line's seq, per writer, and what is missing between them.</summary>
        private sealed class WriterSequences
        {
            private readonly Dictionary<string, Track> _tracks = new Dictionary<string, Track>(StringComparer.Ordinal);

            internal void Add(int pid, string run, long seq)
            {
                string key = pid.ToString(CultureInfo.InvariantCulture) + "/" + run;
                if (!_tracks.TryGetValue(key, out Track? track))
                {
                    track = new Track(pid, run);
                    _tracks.Add(key, track);
                }

                track.Seqs.Add(seq);
            }

            internal void Summarize(AuditLogScan scan)
            {
                List<Track> tracks = new List<Track>(_tracks.Values);
                tracks.Sort((a, b) => a.Pid != b.Pid ? a.Pid.CompareTo(b.Pid) : string.CompareOrdinal(a.Run, b.Run));
                List<AuditLogGap> gaps = new List<AuditLogGap>();
                long missing = 0;
                long duplicates = 0;
                foreach (Track track in tracks)
                {
                    track.Seqs.Sort();
                    long expected = 1;
                    long previous = 0;
                    foreach (long seq in track.Seqs)
                    {
                        if (seq == previous)
                        {
                            duplicates++;
                            continue;
                        }

                        if (seq > expected)
                        {
                            missing += seq - expected;
                            if (gaps.Count < GapsCap)
                            {
                                gaps.Add(new AuditLogGap(track.Pid, track.Run, expected, seq - 1));
                            }
                        }

                        expected = seq + 1;
                        previous = seq;
                    }
                }

                scan.Writers = tracks.Count;
                scan.MissingLines = missing;
                scan.DuplicateLines = duplicates;
                scan.Gaps = gaps;
            }

            private sealed class Track
            {
                internal Track(int pid, string run)
                {
                    Pid = pid;
                    Run = run;
                }

                internal int Pid { get; }

                internal string Run { get; }

                internal List<long> Seqs { get; } = new List<long>();
            }
        }
    }
}
