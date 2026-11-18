using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Microsoft.Win32.SafeHandles;
using OutlookAI.Core.Audit;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Which OutlookAI server a process works for. The server reaches Outlook through a COM host child
    /// it starts with its own process id in the environment; in-process - the live tier, the corpus
    /// tools - the server and the host are one process.
    /// </summary>
    public static class ComHostIdentity
    {
        /// <summary>
        /// The environment variable carrying the server's process id to its COM host. One definition: the
        /// host's protocol reads its own copy of the name from here.
        /// </summary>
        public const string ParentPidVariable = "OUTLOOKAI_COMHOST_PARENT_PID";

        private static readonly int Host = AuditLogFile.CurrentProcessId();

        private static readonly int Server = ReadServer();

        /// <summary>This process.</summary>
        public static int HostProcessId => Host;

        /// <summary>The OutlookAI server this process works for: its parent when it is a COM host, else itself.</summary>
        public static int ServerProcessId => Server;

        private static int ReadServer()
        {
            string? raw = Environment.GetEnvironmentVariable(ParentPidVariable);
            return int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parent) && parent > 0
                ? parent
                : Host;
        }
    }

    /// <summary>
    /// The start log (Q178 (e), decided by the maintainer 2026-11-18): one line for every step OutlookAI
    /// takes in starting or attaching to Outlook, every wait for it to settle, and every time it found
    /// Outlook not responding - under <c>%LOCALAPPDATA%\OutlookAI\outlook-start.log</c>, with no mail
    /// content.
    /// <para>
    /// <b>Why.</b> Nothing OutlookAI wrote recorded the start that hung on 2026-11-15, so the diagnosis
    /// could not name which call was in progress inside Outlook, nor what started Outlook again 67 ms
    /// after the hung one was ended. The next hang should leave a trail a user can attach to a report.
    /// </para>
    /// <para>
    /// <b>No mail content, by construction.</b> A line is a time, an event name and fields whose keys are
    /// constants and whose values are numbers, times and tokens: every string value is reduced to letters,
    /// digits and <c>. _ : / + -</c> and cut at <see cref="MaxValueLength"/>, so a sentence, a subject or a
    /// store name cannot be written even by mistake. Failures are logged by exception type and HRESULT,
    /// never by message - messages name folders and stores.
    /// </para>
    /// <para>
    /// Appended through the audit log's own append handle (<c>AuditLogFile</c>), so writers in several
    /// processes never overwrite each other, and into the audit log's directory, so a test process's
    /// lines land in its throwaway directory with its audit lines (<c>AuditLog.RedirectThisProcess</c>).
    /// Kept to about <see cref="RotateAtBytes"/>: past it the file becomes <see cref="PreviousFileName"/>,
    /// replacing the one before. Never throws, never blocks for long: a log that cannot be written must
    /// not fail or delay the start it describes.
    /// </para>
    /// </summary>
    public static class OutlookStartLog
    {
        /// <summary>The log's file name, beside the audit log.</summary>
        public const string FileName = "outlook-start.log";

        /// <summary>What the log becomes once it passes <see cref="RotateAtBytes"/>.</summary>
        public const string PreviousFileName = "outlook-start.1.log";

        /// <summary>The size past which the log starts again - about 8,000 lines.</summary>
        public const long RotateAtBytes = 1024 * 1024;

        /// <summary>The longest a field value may be.</summary>
        internal const int MaxValueLength = 120;

        /// <summary>Where a shipped OutlookAI writes the log - for health and the README.</summary>
        public static string DefaultPath => Path.Combine(AuditLog.DefaultDirectory, FileName);

        /// <summary>Where THIS process writes it: <see cref="DefaultPath"/>, or a test run's throwaway directory.</summary>
        public static string EffectivePath => Path.Combine(AuditLog.EffectiveDirectory, FileName);

        /// <summary>Appends one line. Never throws.</summary>
        /// <param name="eventName">A constant token such as <c>settle.end</c>.</param>
        /// <param name="fields">Key and value pairs; null values are left out.</param>
        public static void Write(string eventName, params KeyValuePair<string, object?>[] fields)
        {
            try
            {
                string directory = AuditLog.EffectiveDirectory;
                if (AuditLog.DescribeRefusalWhileRedirected(directory) != null)
                {
                    return;
                }

                Directory.CreateDirectory(directory);
                string path = Path.Combine(directory, FileName);
                RotateIfFull(path, Path.Combine(directory, PreviousFileName));

                string line = FormatLine(
                    AuditLogFile.PreciseUtcNow(), ComHostIdentity.ServerProcessId, ComHostIdentity.HostProcessId, eventName, fields);
                byte[] bytes = Encoding.UTF8.GetBytes(line + "\r\n");
                for (int attempt = 0; attempt < 5; attempt++)
                {
                    if (AuditLogFile.TryOpenForAppend(path, writeThrough: false, out SafeFileHandle? handle, out int error))
                    {
                        using (handle)
                        {
                            AuditLogFile.Append(handle!, bytes, bytes.Length);
                        }

                        return;
                    }

                    if (!AuditLogFile.IsSharingViolation(error))
                    {
                        return;
                    }

                    System.Threading.Thread.Sleep(10);
                }
            }
            catch (Exception ex) when (!(ex is OutOfMemoryException))
            {
                // Best effort, by design: see the class remarks.
            }
        }

        /// <summary>A field for <see cref="Write"/>.</summary>
        public static KeyValuePair<string, object?> Field(string key, object? value)
        {
            return new KeyValuePair<string, object?>(key, value);
        }

        /// <summary>
        /// One line of the log, without its line break. Pure, so T1 pins the shape and the rule that
        /// nothing but tokens gets through.
        /// </summary>
        internal static string FormatLine(
            DateTime utc, int serverProcessId, int hostProcessId, string eventName, IEnumerable<KeyValuePair<string, object?>> fields)
        {
            StringBuilder line = new StringBuilder(160);
            line.Append(DateTime.SpecifyKind(utc, DateTimeKind.Utc).ToString("o", CultureInfo.InvariantCulture));
            line.Append(' ').Append(Token(eventName));
            line.Append(" server=").Append(serverProcessId.ToString(CultureInfo.InvariantCulture));
            line.Append(" host=").Append(hostProcessId.ToString(CultureInfo.InvariantCulture));
            if (fields != null)
            {
                foreach (KeyValuePair<string, object?> field in fields)
                {
                    if (field.Value == null)
                    {
                        continue;
                    }

                    line.Append(' ').Append(Token(field.Key)).Append('=').Append(Value(field.Value));
                }
            }

            return line.ToString();
        }

        private static string Value(object value)
        {
            switch (value)
            {
                case bool flag:
                    return flag ? "true" : "false";
                case DateTime time:
                    return DateTime.SpecifyKind(time, DateTimeKind.Utc).ToString("o", CultureInfo.InvariantCulture);
                case int number:
                    return number.ToString(CultureInfo.InvariantCulture);
                case long number:
                    return number.ToString(CultureInfo.InvariantCulture);
                case double number:
                    return number.ToString("0.###", CultureInfo.InvariantCulture);
                default:
                    return Token(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
            }
        }

        /// <summary>
        /// A value reduced to a token: letters, digits and <c>. _ : / + -</c> survive, anything else becomes
        /// <c>_</c>, and it is cut at <see cref="MaxValueLength"/>.
        /// </summary>
        internal static string Token(string? text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return "-";
            }

            int length = Math.Min(text!.Length, MaxValueLength);
            StringBuilder token = new StringBuilder(length);
            for (int i = 0; i < length; i++)
            {
                char c = text[i];
                bool keep = (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= '0' && c <= '9')
                    || c == '.' || c == '_' || c == ':' || c == '/' || c == '+' || c == '-';
                token.Append(keep ? c : '_');
            }

            return token.ToString();
        }

        /// <summary>
        /// Starts the log again once it is full. Best effort: another process appending at that moment
        /// keeps the file open without delete sharing, the rename fails, and the next line tries again.
        /// </summary>
        private static void RotateIfFull(string path, string previousPath)
        {
            try
            {
                FileInfo file = new FileInfo(path);
                if (!file.Exists || file.Length < RotateAtBytes)
                {
                    return;
                }

                if (File.Exists(previousPath))
                {
                    File.Delete(previousPath);
                }

                File.Move(path, previousPath);
            }
            catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException)
            {
                // Somebody is appending: rotate on a later line.
            }
        }
    }
}
