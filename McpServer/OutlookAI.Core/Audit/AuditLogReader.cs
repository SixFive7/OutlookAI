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
        {
            TimestampUtc = timestampUtc;
            Operation = operation;
            Fields = fields;
        }

        /// <summary>The line's <c>ts</c>, UTC, to the millisecond the writer recorded.</summary>
        public DateTime TimestampUtc { get; }

        /// <summary>The line's <c>op</c> token.</summary>
        public string Operation { get; }

        /// <summary>
        /// The line's fields in line order, values unescaped. A key could in principle appear
        /// twice - <see cref="AuditLog.FormatLine"/> does not forbid it, though no call site does it -
        /// so this stays a list rather than a dictionary.
        /// </summary>
        public IReadOnlyList<KeyValuePair<string, string>> Fields { get; }
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

    /// <summary>What one read of the log found.</summary>
    public sealed class AuditLogScan
    {
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
        public bool FileFound { get; }

        /// <summary>The newest matching entries, newest first, at most the requested number.</summary>
        public IReadOnlyList<AuditLogEntry> Entries { get; }

        /// <summary>How many entries matched in total; more than <see cref="Entries"/> holds when the limit cut.</summary>
        public int Matched { get; }

        /// <summary>Complete lines read, malformed ones included.</summary>
        public long LinesScanned { get; }

        /// <summary>Complete lines that do not parse as an audit line, and were therefore not returned.</summary>
        public long MalformedLines { get; }

        /// <summary>
        /// True when the file ended without a line terminator: a line still being written when the
        /// read began, or one a crash cut off. It is not parsed and not counted in
        /// <see cref="LinesScanned"/>.
        /// </summary>
        public bool IncompleteLastLine { get; }

        /// <summary>
        /// The distinct operation names of every parsed line, filter or not, sorted - at most
        /// <see cref="AuditLogReader.OperationsSeenCap"/>. Read off the data rather than kept as a
        /// list of what the server writes, so a reply to a filter that matched nothing can say
        /// what the log does hold without a second copy of the operation names to drift.
        /// </summary>
        public IReadOnlyList<string> OperationsSeen { get; }
    }

    /// <summary>
    /// Reads the audit log back (Q93): the parser for the line grammar <see cref="AuditLog.FormatLine"/>
    /// writes, and a scan that never gets in the way of the writers.
    /// <para>
    /// <b>Sharing.</b> The file is opened <c>FileShare.ReadWrite | FileShare.Delete</c>, read-only.
    /// The product's appends open it for writing with <c>FileShare.ReadWrite</c>, which this
    /// reader's handle allows, and this reader asks for no access an appender's share mode
    /// excludes - so a read never blocks an append, which matters because a draft or send whose
    /// audit line cannot be written is reported as a failure. <c>FileShare.Delete</c> is there so a
    /// read in progress cannot stop the maintainer renaming the log either.
    /// </para>
    /// <para>
    /// <b>A file that grows while it is read.</b> The length is taken once, when the file is
    /// opened, and nothing past it is read: lines appended later are for the next call, and the
    /// answer is a consistent prefix of the log. A last line without its terminator - an append
    /// in flight at that instant, or one a crash cut short - is reported, not parsed.
    /// </para>
    /// <para>
    /// <b>Strict.</b> <see cref="ParseLine"/> accepts exactly what <see cref="AuditLog.FormatLine"/>
    /// can produce and nothing else: an unknown escape, a raw tab or carriage return, a missing
    /// quote, a blank line or invalid UTF-8 is malformed, counted and skipped, never guessed at.
    /// That includes the unstructured lines OutlookAI wrote before 2026-07-23.
    /// </para>
    /// </summary>
    public static class AuditLogReader
    {
        /// <summary>A line longer than this is malformed by definition; the longest real line is a few hundred bytes.</summary>
        public const int MaxLineBytes = 256 * 1024;

        /// <summary>How many distinct operation names a scan collects for <see cref="AuditLogScan.OperationsSeen"/>.</summary>
        public const int OperationsSeenCap = 40;

        private const string TimestampFormat = "yyyy-MM-dd'T'HH:mm:ss.fff'Z'";
        private const int TimestampLength = 24;
        private const int OpenRetries = 3;
        private const int ChunkBytes = 64 * 1024;

        private static readonly UTF8Encoding StrictUtf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

        /// <summary>
        /// Parses one line (without its terminator), or returns null when it is not exactly a line
        /// <see cref="AuditLog.FormatLine"/> could have written.
        /// </summary>
        public static AuditLogEntry? ParseLine(string line)
        {
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

            return new AuditLogEntry(timestamp, operation, fields);
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
                return new AuditLogScan(false, Array.Empty<AuditLogEntry>(), 0, 0, 0, false, Array.Empty<string>());
            }

            try
            {
                using (stream)
                {
                    return Scan(stream, stream.Length, filter, top);
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
            if (top < 1)
            {
                throw new ArgumentOutOfRangeException(nameof(top), "top must be at least 1.");
            }

            Queue<AuditLogEntry> newest = new Queue<AuditLogEntry>(top + 1);
            SortedSet<string> operationsSeen = new SortedSet<string>(StringComparer.Ordinal);
            int matched = 0;
            long lines = 0;
            long malformed = 0;

            byte[] chunk = new byte[ChunkBytes];
            MemoryStream pending = new MemoryStream();
            bool pendingTooLong = false;
            long remaining = length;

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
                    AuditLogEntry? entry = pendingTooLong || lineCount > MaxLineBytes
                        ? null
                        : DecodeAndParse(lineBytes, lineOffset, lineCount, lines == 1);
                    if (entry == null)
                    {
                        malformed++;
                    }
                    else
                    {
                        if (operationsSeen.Count < OperationsSeenCap)
                        {
                            operationsSeen.Add(entry.Operation);
                        }

                        if (filter.Matches(entry))
                        {
                            matched++;
                            newest.Enqueue(entry);
                            if (newest.Count > top)
                            {
                                newest.Dequeue();
                            }
                        }
                    }

                    pending.SetLength(0);
                    pendingTooLong = false;
                    lineStart = i + 1;
                }

                AppendPending(pending, chunk, lineStart, read - lineStart, ref pendingTooLong);
            }

            bool incomplete = pending.Length > 0 || pendingTooLong;

            AuditLogEntry[] entries = newest.ToArray();
            Array.Reverse(entries);
            return new AuditLogScan(true, entries, matched, lines, malformed, incomplete, new List<string>(operationsSeen));
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

        private static AuditLogEntry? DecodeAndParse(byte[] bytes, int offset, int count, bool firstLine)
        {
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

            return ParseLine(line);
        }

        private static bool IsTokenChar(char c)
        {
            // The writer's own rule for operation names and field keys (AuditLog.FormatLine).
            return char.IsLetterOrDigit(c) || c == '_' || c == '-';
        }
    }
}
