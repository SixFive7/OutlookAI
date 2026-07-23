using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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
    /// </summary>
    public static class AuditLog
    {
        private const int WriteRetries = 3;

        /// <summary>Shared OutlookAI state root (v3.MD section 0.5.2).</summary>
        public static string DefaultDirectory =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "OutlookAI");

        /// <summary>Full path of the audit log file every write op appends to.</summary>
        public static string DefaultLogPath => Path.Combine(DefaultDirectory, "audit.log");

        /// <summary>
        /// Appends one structured line for <paramref name="operation"/> to the default
        /// audit log. Throws <see cref="InvalidOperationException"/> when the line
        /// cannot be written (load-bearing from Phase 4 - callers surface the failure).
        /// </summary>
        public static void Append(string operation, params (string Key, string? Value)[] fields)
        {
            AppendTo(DefaultDirectory, operation, fields);
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
            string path = Path.Combine(directory, "audit.log");
            try
            {
                Directory.CreateDirectory(directory);
                IOException? lastIo = null;
                for (int attempt = 0; attempt < WriteRetries; attempt++)
                {
                    try
                    {
                        // FileShare.ReadWrite: multiple server processes (one per agent
                        // session) may append concurrently; FileMode.Append positions at
                        // end-of-file at open time and short lines interleave cleanly in
                        // practice.
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

                throw new InvalidOperationException(
                    "Audit line could not be written to '" + path + "' after " +
                    WriteRetries.ToString(CultureInfo.InvariantCulture) + " attempts.", lastIo);
            }
            catch (Exception ex) when (ex is UnauthorizedAccessException || ex is NotSupportedException || ex is ArgumentException)
            {
                throw new InvalidOperationException("Audit line could not be written to '" + path + "'.", ex);
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
