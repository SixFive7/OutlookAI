using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using Microsoft.Win32;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// The connection that holds the start lock, or held it last: which Outlook it started or attached
    /// to, which server it works for, and which start-up step it is in, since when.
    /// </summary>
    public sealed class OutlookConnectRecord
    {
        /// <summary>Creates the record.</summary>
        public OutlookConnectRecord(
            int? outlookProcessId,
            DateTime? outlookStartUtc,
            bool startedOutlook,
            int serverProcessId,
            int hostProcessId,
            string step,
            DateTime stepSinceUtc,
            DateTime beganUtc)
        {
            OutlookProcessId = outlookProcessId;
            OutlookStartUtc = outlookStartUtc;
            StartedOutlook = startedOutlook;
            ServerProcessId = serverProcessId;
            HostProcessId = hostProcessId;
            Step = step ?? string.Empty;
            StepSinceUtc = stepSinceUtc;
            BeganUtc = beganUtc;
        }

        /// <summary>The OUTLOOK.EXE this connection started or attached to; null before it knows.</summary>
        public int? OutlookProcessId { get; }

        /// <summary>When that OUTLOOK.EXE started (UTC); null before it knows.</summary>
        public DateTime? OutlookStartUtc { get; }

        /// <summary>Whether this connection started Outlook rather than attaching to it.</summary>
        public bool StartedOutlook { get; }

        /// <summary>The OutlookAI MCP server the connection works for.</summary>
        public int ServerProcessId { get; }

        /// <summary>The process making the calls into Outlook: the server's COM host, or the server itself in-process.</summary>
        public int HostProcessId { get; }

        /// <summary>The start-up step (<see cref="OutlookStartSteps"/>).</summary>
        public string Step { get; }

        /// <summary>Since when the connection has been in <see cref="Step"/> (UTC).</summary>
        public DateTime StepSinceUtc { get; }

        /// <summary>When the connection began (UTC).</summary>
        public DateTime BeganUtc { get; }

        /// <summary>True while the connection is in a step, as opposed to finished or failed.</summary>
        public bool InProgress => !OutlookStartSteps.IsFinal(Step);

        /// <summary>The same record in a later step.</summary>
        public OutlookConnectRecord WithStep(string step, DateTime nowUtc, int? outlookProcessId, DateTime? outlookStartUtc, bool startedOutlook)
        {
            return new OutlookConnectRecord(outlookProcessId, outlookStartUtc, startedOutlook, ServerProcessId, HostProcessId, step, nowUtc, BeganUtc);
        }

        /// <summary>The record as one registry value - one value, so a reader never sees half of a write.</summary>
        public string Format()
        {
            return OutlookStartRecordFields.Format(new List<KeyValuePair<string, string?>>
            {
                Field("outlookPid", OutlookProcessId?.ToString(CultureInfo.InvariantCulture)),
                Field("outlookStartUtc", OutlookStartRecordFields.FormatTime(OutlookStartUtc)),
                Field("started", StartedOutlook ? "1" : "0"),
                Field("server", ServerProcessId.ToString(CultureInfo.InvariantCulture)),
                Field("host", HostProcessId.ToString(CultureInfo.InvariantCulture)),
                Field("step", Step),
                Field("stepSinceUtc", OutlookStartRecordFields.FormatTime(StepSinceUtc)),
                Field("beganUtc", OutlookStartRecordFields.FormatTime(BeganUtc)),
            });
        }

        /// <summary>Reads a record back; null when <paramref name="text"/> is absent or not a record.</summary>
        public static OutlookConnectRecord? Parse(string? text)
        {
            Dictionary<string, string> fields = OutlookStartRecordFields.Parse(text);
            int? server = OutlookStartRecordFields.ReadInt(fields, "server");
            int? host = OutlookStartRecordFields.ReadInt(fields, "host");
            DateTime? stepSince = OutlookStartRecordFields.ReadTime(fields, "stepSinceUtc");
            DateTime? began = OutlookStartRecordFields.ReadTime(fields, "beganUtc");
            if (server == null || host == null || stepSince == null || began == null
                || !fields.TryGetValue("step", out string? step) || step.Length == 0)
            {
                return null;
            }

            return new OutlookConnectRecord(
                OutlookStartRecordFields.ReadInt(fields, "outlookPid"),
                OutlookStartRecordFields.ReadTime(fields, "outlookStartUtc"),
                fields.TryGetValue("started", out string? started) && started == "1",
                server.Value,
                host.Value,
                step,
                stepSince.Value,
                began.Value);
        }

        private static KeyValuePair<string, string?> Field(string key, string? value)
        {
            return new KeyValuePair<string, string?>(key, value);
        }
    }

    /// <summary>
    /// One fact about one OUTLOOK.EXE, noted by whichever OutlookAI process saw it first: that OutlookAI
    /// started it, that Windows reported it not responding, or that a window of it was on screen. Keyed by
    /// the process id AND its start time, so a note about an Outlook that has gone never describes the
    /// next one, even under the same id.
    /// </summary>
    public sealed class OutlookProcessNote
    {
        /// <summary>Creates the note.</summary>
        public OutlookProcessNote(int processId, DateTime? processStartUtc, DateTime atUtc, int? serverProcessId)
        {
            ProcessId = processId;
            ProcessStartUtc = processStartUtc;
            AtUtc = atUtc;
            ServerProcessId = serverProcessId;
        }

        /// <summary>The OUTLOOK.EXE the note is about.</summary>
        public int ProcessId { get; }

        /// <summary>When that process started (UTC), when known.</summary>
        public DateTime? ProcessStartUtc { get; }

        /// <summary>When the fact was first seen (UTC).</summary>
        public DateTime AtUtc { get; }

        /// <summary>The OutlookAI server that saw it, when that matters (who started Outlook).</summary>
        public int? ServerProcessId { get; }

        /// <summary>
        /// Whether the note is about the Outlook with <paramref name="processId"/> that started at
        /// <paramref name="processStartUtc"/>. Start times are compared to the second: both come from the
        /// same system clock, one through the process's creation time and one as it was written down.
        /// </summary>
        public bool IsAbout(int? processId, DateTime? processStartUtc)
        {
            if (processId != ProcessId)
            {
                return false;
            }

            if (ProcessStartUtc is DateTime mine && processStartUtc is DateTime theirs)
            {
                return Math.Abs((mine - theirs).TotalSeconds) < 1;
            }

            return ProcessStartUtc == null && processStartUtc == null;
        }

        /// <summary>The note as one registry value.</summary>
        public string Format()
        {
            return OutlookStartRecordFields.Format(new List<KeyValuePair<string, string?>>
            {
                new KeyValuePair<string, string?>("outlookPid", ProcessId.ToString(CultureInfo.InvariantCulture)),
                new KeyValuePair<string, string?>("outlookStartUtc", OutlookStartRecordFields.FormatTime(ProcessStartUtc)),
                new KeyValuePair<string, string?>("atUtc", OutlookStartRecordFields.FormatTime(AtUtc)),
                new KeyValuePair<string, string?>("server", ServerProcessId?.ToString(CultureInfo.InvariantCulture)),
            });
        }

        /// <summary>Reads a note back; null when <paramref name="text"/> is absent or not a note.</summary>
        public static OutlookProcessNote? Parse(string? text)
        {
            Dictionary<string, string> fields = OutlookStartRecordFields.Parse(text);
            int? pid = OutlookStartRecordFields.ReadInt(fields, "outlookPid");
            DateTime? at = OutlookStartRecordFields.ReadTime(fields, "atUtc");
            if (pid == null || at == null)
            {
                return null;
            }

            return new OutlookProcessNote(
                pid.Value,
                OutlookStartRecordFields.ReadTime(fields, "outlookStartUtc"),
                at.Value,
                OutlookStartRecordFields.ReadInt(fields, "server"));
        }
    }

    /// <summary>The start-up steps a connection records, in the order it takes them.</summary>
    public static class OutlookStartSteps
    {
        /// <summary>Asking Windows for Outlook.Application - which starts OUTLOOK.EXE when none runs.</summary>
        public const string CreateApplication = "create-application";

        /// <summary>Waiting for Outlook to settle before any mailbox is opened (Q175 (a)).</summary>
        public const string Settle = "settle";

        /// <summary>Getting the MAPI namespace.</summary>
        public const string Namespace = "namespace";

        /// <summary>Logging on to the profile - the first step that can open a mailbox.</summary>
        public const string Logon = "logon";

        /// <summary>Advising the Quit event sink.</summary>
        public const string QuitSink = "quit-sink";

        /// <summary>Connected: every step done.</summary>
        public const string Connected = "connected";

        /// <summary>Failed: the connection gave up or was refused.</summary>
        public const string Failed = "failed";

        /// <summary>
        /// Whether a connection in <paramref name="step"/> had got past the settle wait - so the Outlook it was
        /// in had finished starting. "failed" does not say where it failed, so it does not count.
        /// </summary>
        public static bool IsPastSettle(string? step)
        {
            return string.Equals(step, Namespace, StringComparison.Ordinal)
                || string.Equals(step, Logon, StringComparison.Ordinal)
                || string.Equals(step, QuitSink, StringComparison.Ordinal)
                || string.Equals(step, Connected, StringComparison.Ordinal);
        }

        /// <summary>Whether <paramref name="step"/> ends a connection rather than being one of its steps.</summary>
        public static bool IsFinal(string? step)
        {
            return string.Equals(step, Connected, StringComparison.Ordinal)
                || string.Equals(step, Failed, StringComparison.Ordinal);
        }
    }

    /// <summary>
    /// The per-user record of the Outlook OutlookAI started (Q176 (c), decided by the maintainer
    /// 2026-11-18): its process and start time, which server started it, which start-up step the
    /// connection is in and since when - plus when it was first seen not responding and whether it was
    /// ever seen with a window, which is what an Outlook that hangs with no window needs reported
    /// (Q178 (e)). Shared by every OutlookAI process of the user, so the server that did not start an
    /// Outlook can still say who did, and since when it has been stuck.
    /// <para>
    /// Under <c>HKCU\Software\OutlookAI\OutlookStart</c>, one value per fact, each written whole - a reader
    /// never sees half of a write. Nothing in it is mail content, and nothing in it is personal beyond
    /// process ids and times.
    /// </para>
    /// <para>
    /// Best effort, on purpose: a record that cannot be read or written leaves the report shorter, and
    /// never fails or slows the request it describes. Every member here swallows registry failures.
    /// </para>
    /// </summary>
    public static class OutlookStartRecord
    {
        /// <summary>The record's key, under HKEY_CURRENT_USER.</summary>
        public const string KeyPath = @"Software\OutlookAI\OutlookStart";

        /// <summary>The connection that holds the start lock, or held it last (<see cref="OutlookConnectRecord"/>).</summary>
        public const string ConnectValueName = "Connect";

        /// <summary>The Outlook an OutlookAI connection started, and which server it worked for.</summary>
        public const string StartedValueName = "Started";

        /// <summary>When an Outlook was first seen not responding.</summary>
        public const string NotRespondingValueName = "NotResponding";

        /// <summary>When an Outlook was first seen with a window on screen.</summary>
        public const string WindowSeenValueName = "WindowSeen";

        /// <summary>The connection record, or null when there is none or it cannot be read.</summary>
        public static OutlookConnectRecord? ReadConnect()
        {
            return OutlookConnectRecord.Parse(ReadValue(ConnectValueName));
        }

        /// <summary>Writes the connection record. Never throws.</summary>
        public static void WriteConnect(OutlookConnectRecord record)
        {
            if (record != null)
            {
                WriteValue(ConnectValueName, record.Format());
            }
        }

        /// <summary>The note stored under <paramref name="valueName"/>, or null.</summary>
        public static OutlookProcessNote? ReadNote(string valueName)
        {
            return OutlookProcessNote.Parse(ReadValue(valueName));
        }

        /// <summary>Writes a note. Never throws.</summary>
        public static void WriteNote(string valueName, OutlookProcessNote note)
        {
            if (note != null)
            {
                WriteValue(valueName, note.Format());
            }
        }

        /// <summary>
        /// Notes what one look at Outlook saw, for the report and for health: a window on screen, and when it
        /// was first seen not responding - cleared again once it answers. Whichever OutlookAI process sees it
        /// first writes it, so the "since" is the earliest any of them saw. Reads one value per call and
        /// writes only when something changed. Never throws.
        /// </summary>
        public static void NoteWhatWasSeen(OutlookStartObservation observation, OutlookStartVerdict verdict)
        {
            if (observation == null || verdict == null || !(observation.Windows.ProcessId is int pid))
            {
                return;
            }

            DateTime? start = observation.ProcessStartUtc;
            string identity = pid.ToString(CultureInfo.InvariantCulture) + "@" + (OutlookStartRecordFields.FormatTime(start) ?? "?");
            if (observation.Windows.HasVisibleWindow
                && !string.Equals(System.Threading.Volatile.Read(ref _windowNoted), identity, StringComparison.Ordinal))
            {
                OutlookProcessNote? seen = ReadNote(WindowSeenValueName);
                if (seen == null || !seen.IsAbout(pid, start))
                {
                    WriteNote(WindowSeenValueName, new OutlookProcessNote(pid, start, observation.ObservedUtc, null));
                }

                System.Threading.Volatile.Write(ref _windowNoted, identity);
            }

            OutlookProcessNote? stuck = ReadNote(NotRespondingValueName);
            bool noted = stuck != null && stuck.IsAbout(pid, start);
            if (verdict.State == OutlookStartState.NotResponding)
            {
                if (!noted)
                {
                    WriteNote(NotRespondingValueName, new OutlookProcessNote(pid, start, observation.ObservedUtc, null));
                }
            }
            else if (noted)
            {
                ClearNote(NotRespondingValueName);
            }
        }

        private static string? _windowNoted;

        /// <summary>Removes a note - an Outlook that answers again is no longer "not responding since". Never throws.</summary>
        public static void ClearNote(string valueName)
        {
            try
            {
                using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(KeyPath, writable: true))
                {
                    key?.DeleteValue(valueName, throwOnMissingValue: false);
                }
            }
            catch (Exception ex) when (IsRegistryFailure(ex))
            {
                // Best effort: a stale note is about an Outlook that answered once, and says when.
            }
        }

        private static string? ReadValue(string valueName)
        {
            try
            {
                using (RegistryKey? key = Registry.CurrentUser.OpenSubKey(KeyPath, writable: false))
                {
                    return key?.GetValue(valueName) as string;
                }
            }
            catch (Exception ex) when (IsRegistryFailure(ex))
            {
                return null;
            }
        }

        private static void WriteValue(string valueName, string value)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.CreateSubKey(KeyPath))
                {
                    key.SetValue(valueName, value, RegistryValueKind.String);
                }
            }
            catch (Exception ex) when (IsRegistryFailure(ex))
            {
                // Best effort: the record explains a hang; it must never cause or prolong one.
            }
        }

        private static bool IsRegistryFailure(Exception ex)
        {
            return ex is System.Security.SecurityException
                || ex is UnauthorizedAccessException
                || ex is System.IO.IOException
                || ex is ObjectDisposedException
                || ex is ArgumentException;
        }
    }

    /// <summary>
    /// The <c>key=value;key=value</c> shape every value of the start record is written in. Values are
    /// numbers, round-trip UTC times and step tokens, none of which contains <c>;</c> or <c>=</c>; a value
    /// that would is refused rather than escaped, because nothing that could carry one belongs here.
    /// </summary>
    internal static class OutlookStartRecordFields
    {
        internal static string Format(IList<KeyValuePair<string, string?>> fields)
        {
            StringBuilder text = new StringBuilder();
            foreach (KeyValuePair<string, string?> field in fields)
            {
                if (field.Value == null)
                {
                    continue;
                }

                if (field.Value.IndexOf(';') >= 0 || field.Value.IndexOf('=') >= 0)
                {
                    throw new ArgumentException("A start record value may not contain ';' or '=' (" + field.Key + ").");
                }

                if (text.Length > 0)
                {
                    text.Append(';');
                }

                text.Append(field.Key).Append('=').Append(field.Value);
            }

            return text.ToString();
        }

        internal static Dictionary<string, string> Parse(string? text)
        {
            Dictionary<string, string> fields = new Dictionary<string, string>(StringComparer.Ordinal);
            if (string.IsNullOrEmpty(text))
            {
                return fields;
            }

            foreach (string part in text!.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                int equals = part.IndexOf('=');
                if (equals <= 0)
                {
                    continue;
                }

                fields[part.Substring(0, equals)] = part.Substring(equals + 1);
            }

            return fields;
        }

        /// <summary>
        /// A UTC time, round-trip. Every time in the record is UTC when it is made - <c>DateTime.UtcNow</c>,
        /// a process's creation time, a parsed record - so it is stamped UTC rather than converted, which
        /// would read the machine's zone.
        /// </summary>
        internal static string? FormatTime(DateTime? utc)
        {
            return utc == null ? null : DateTime.SpecifyKind(utc.Value, DateTimeKind.Utc).ToString("o", CultureInfo.InvariantCulture);
        }

        internal static int? ReadInt(Dictionary<string, string> fields, string key)
        {
            return fields.TryGetValue(key, out string? raw)
                && int.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out int value)
                    ? value
                    : (int?)null;
        }

        internal static DateTime? ReadTime(Dictionary<string, string> fields, string key)
        {
            return fields.TryGetValue(key, out string? raw)
                && DateTime.TryParse(raw, CultureInfo.InvariantCulture, DateTimeStyles.AdjustToUniversal | DateTimeStyles.AssumeUniversal, out DateTime value)
                    ? DateTime.SpecifyKind(value, DateTimeKind.Utc)
                    : (DateTime?)null;
        }
    }
}
