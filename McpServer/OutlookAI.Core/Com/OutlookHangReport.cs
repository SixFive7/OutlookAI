using System;
using System.Globalization;
using System.Text;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// Raised instead of a call into an Outlook that is not responding, or that has not finished starting
    /// within the limit - by the gate in front of each request, or by a connection that finds it so while
    /// it waits. Its message is the whole report: what is stuck and why, the process, how to end it and
    /// what ending it risks (<see cref="OutlookHangReport"/>).
    /// <para>
    /// Derives from <see cref="TimeoutException"/> so every layer that degrades on a deadline - the
    /// freshness sweep falls back to index-only results - degrades on this too, carrying the report as
    /// its reason. Nothing was dispatched to Outlook when it is raised.
    /// </para>
    /// </summary>
    public sealed class OutlookNotRespondingException : TimeoutException
    {
        /// <summary>
        /// What the agent is told to do with the report (Q177 (a), decided by the maintainer 2026-11-18:
        /// OutlookAI never ends Outlook; it reports clearly, and the user ends it). The maintainer: "we can
        /// then report back to the calling llm that outlook is stuck and why. Allowing for the user to
        /// report back to us or ask the llm to fix it."
        /// </summary>
        public const string AgentAdvice =
            "TELL THE USER what the message says: what is stuck and why, the process ID, how to end it and what ending it "
            + "risks. Let the user decide; end the process yourself only if the user asks you to. OutlookAI never ends "
            + "Outlook itself. search still returns indexed mail meanwhile, and outlook_health shows the details. If it "
            + "happens again, the start log named in the message is what a report to OutlookAI's developers needs.";

        /// <summary>Creates the exception.</summary>
        /// <param name="message">The report (<see cref="OutlookHangReport.Describe"/>).</param>
        /// <param name="retryAfterSeconds">When a caller should look again.</param>
        /// <param name="processId">The OUTLOOK.EXE, when known.</param>
        public OutlookNotRespondingException(string message, int retryAfterSeconds, int? processId = null)
            : base(message)
        {
            RetryAfterSeconds = retryAfterSeconds;
            ProcessId = processId;
        }

        /// <summary>When a caller should look again.</summary>
        public int RetryAfterSeconds { get; }

        /// <summary>The OUTLOOK.EXE, when known. Not carried across the COM host's pipe; the message names it.</summary>
        public int? ProcessId { get; }
    }

    /// <summary>
    /// Everything one hang report says, as plain values, so T1 pins the text without an Outlook.
    /// </summary>
    public sealed class OutlookHangFacts
    {
        /// <summary>Creates the facts.</summary>
        public OutlookHangFacts(
            OutlookStuckKind kind,
            string reason,
            int? processId,
            DateTime? processStartUtc,
            bool hasWindow,
            bool everShowedWindow,
            int? startedByServerProcessId,
            DateTime? notRespondingSinceUtc,
            OutlookConnectRecord? connect,
            bool connectInProgress,
            DateTime nowUtc,
            string logPath)
        {
            Kind = kind;
            Reason = reason ?? string.Empty;
            ProcessId = processId;
            ProcessStartUtc = processStartUtc;
            HasWindow = hasWindow;
            EverShowedWindow = everShowedWindow;
            StartedByServerProcessId = startedByServerProcessId;
            NotRespondingSinceUtc = notRespondingSinceUtc;
            Connect = connect;
            ConnectInProgress = connectInProgress;
            NowUtc = nowUtc;
            LogPath = logPath ?? string.Empty;
        }

        /// <summary>What kind of stuck.</summary>
        public OutlookStuckKind Kind { get; }

        /// <summary>The start decision's reason.</summary>
        public string Reason { get; }

        /// <summary>The OUTLOOK.EXE.</summary>
        public int? ProcessId { get; }

        /// <summary>When it started (UTC).</summary>
        public DateTime? ProcessStartUtc { get; }

        /// <summary>Whether a window of it is on screen now.</summary>
        public bool HasWindow { get; }

        /// <summary>Whether it was ever seen with a window - now, or by an OutlookAI process before.</summary>
        public bool EverShowedWindow { get; }

        /// <summary>The OutlookAI server that started it, when one did.</summary>
        public int? StartedByServerProcessId { get; }

        /// <summary>When it was first seen not responding (UTC), when it was.</summary>
        public DateTime? NotRespondingSinceUtc { get; }

        /// <summary>The OutlookAI connection that started or attached to it, when the record names this Outlook.</summary>
        public OutlookConnectRecord? Connect { get; }

        /// <summary>Whether that connection is still in its step (it holds the start lock), rather than gone.</summary>
        public bool ConnectInProgress { get; }

        /// <summary>Now (UTC).</summary>
        public DateTime NowUtc { get; }

        /// <summary>The start log a report should carry.</summary>
        public string LogPath { get; }
    }

    /// <summary>
    /// The words OutlookAI uses for an Outlook that is not responding, or that never finished starting
    /// (Q177 (a) and Q178 (e), decided by the maintainer 2026-11-18).
    /// <para>
    /// <b>Why the text depends on whether Outlook has a window.</b> The Outlook that hung on 2026-11-15 had
    /// been started by OutlookAI with no window. The user could not see it or close it, every start of
    /// Outlook was handed to it, and a deadlock does not clear. What OutlookAI said then - "restarting
    /// Outlook fixes it now", "this also clears by itself" - was true of none of that. So an Outlook with
    /// no window gets the Task Manager steps and no promise that it clears; one with a window, which the
    /// user can see and Windows offers to close, may be only busy and is told so.
    /// </para>
    /// <para>
    /// <b>Why the risks are spelled out.</b> The maintainer: "ensure to communicate back the risks of
    /// killing it when communicating back the hang." Ending Outlook is the user's call, and an agent that
    /// is asked to do it needs to say what it may cost first.
    /// </para>
    /// <para>
    /// Every time is written in UTC or as an age - never in the machine's zone, which the reader may not
    /// share.
    /// </para>
    /// </summary>
    public static class OutlookHangReport
    {
        /// <summary>What ending Outlook can cost. One sentence, the same in every variant.</summary>
        public const string EndingRisks =
            "Ending it can lose text typed in an open compose window that was not saved as a draft; mail waiting in the "
            + "Outbox is not sent until Outlook runs again; and if Outlook was writing its data file (.ost or .pst) at that "
            + "moment, the next start can take longer while it checks the file, or offer to start in safe mode.";

        /// <summary>How to end it. The process id is filled in.</summary>
        private const string TaskManagerSteps =
            "open Task Manager (Ctrl+Shift+Esc), go to Details, select OUTLOOK.EXE with PID {0}, and click End task; "
            + "then start Outlook again";

        /// <summary>The report for <paramref name="facts"/>.</summary>
        public static string Describe(OutlookHangFacts facts)
        {
            if (facts == null)
            {
                throw new ArgumentNullException(nameof(facts));
            }

            string pid = facts.ProcessId?.ToString(CultureInfo.InvariantCulture) ?? "?";
            bool windowless = !facts.HasWindow || facts.Kind == OutlookStuckKind.NoWindows;
            StringBuilder text = new StringBuilder(900);

            switch (facts.Kind)
            {
                case OutlookStuckKind.NoWindows:
                    text.Append("Outlook has not finished starting: ").Append(facts.Reason).Append('.');
                    break;
                case OutlookStuckKind.ConnectStuck:
                    text.Append("Outlook is not answering OutlookAI: ").Append(facts.Reason).Append('.');
                    break;
                default:
                    text.Append(windowless
                        ? "Outlook is running without a window and has stopped responding: "
                        : "Outlook has stopped responding: ").Append(facts.Reason).Append('.');
                    break;
            }

            text.Append(" Process: OUTLOOK.EXE, PID ").Append(pid);
            if (facts.ProcessStartUtc is DateTime started)
            {
                text.Append(", started ").Append(Ago(facts.NowUtc - started)).Append(" ago (").Append(Utc(started)).Append(')');
            }

            text.Append(facts.StartedByServerProcessId is int server
                ? ", by OutlookAI for its server process " + server.ToString(CultureInfo.InvariantCulture)
                : string.Empty);
            text.Append('.');

            if (facts.NotRespondingSinceUtc is DateTime since)
            {
                text.Append(" Windows has reported it not responding since ").Append(Utc(since))
                    .Append(", ").Append(Ago(facts.NowUtc - since)).Append(" ago.");
            }

            if (facts.Connect is OutlookConnectRecord connect && !OutlookStartSteps.IsFinal(connect.Step))
            {
                text.Append(facts.ConnectInProgress
                    ? " The OutlookAI connection of server process " + connect.ServerProcessId.ToString(CultureInfo.InvariantCulture)
                        + " has been waiting in start-up step '" + connect.Step + "' since " + Utc(connect.StepSinceUtc) + "."
                    : " The OutlookAI connection of server process " + connect.ServerProcessId.ToString(CultureInfo.InvariantCulture)
                        + " stopped in start-up step '" + connect.Step + "' at " + Utc(connect.StepSinceUtc) + " without finishing.");
            }

            if (facts.Kind == OutlookStuckKind.ConnectStuck)
            {
                text.Append(" Outlook may be waiting for an answer in a dialog - a security prompt about a program "
                    + "using it, say: if one is on screen, answering it may be enough.");
            }

            if (windowless)
            {
                text.Append(facts.EverShowedWindow
                    ? " No window of it is on screen now, so the user cannot close it,"
                    : " It has never shown a window, so the user cannot see or close it,");
                text.Append(" and starting Outlook does nothing until it is ended: Windows hands every new start to this "
                    + "process. Ending it is the only known way back. To end it, ");
            }
            else
            {
                text.Append(" It may recover if it is only busy. If it stays like this, the user can close it when Windows "
                    + "offers to, or end it: ");
            }

            text.AppendFormat(CultureInfo.InvariantCulture, TaskManagerSteps, pid).Append('.');
            text.Append(' ').Append(EndingRisks);
            text.Append(" OutlookAI never ends Outlook itself.");
            if (facts.LogPath.Length > 0)
            {
                text.Append(" The start log, with no mail content, for a report: ").Append(facts.LogPath).Append('.');
            }

            return text.ToString();
        }

        /// <summary>
        /// The report for what one look found, with the start record's notes read now - the one call the
        /// gate, a connection and health all make, so all three say the same words.
        /// </summary>
        public static string DescribeNow(OutlookStartVerdict verdict, OutlookStartObservation observation)
        {
            return Describe(FactsFor(
                verdict,
                observation,
                OutlookStartRecord.ReadNote(OutlookStartRecord.StartedValueName),
                OutlookStartRecord.ReadNote(OutlookStartRecord.NotRespondingValueName),
                OutlookStartRecord.ReadNote(OutlookStartRecord.WindowSeenValueName),
                OutlookStartLog.EffectivePath));
        }

        /// <summary>
        /// The report for what one look found, with the start record's notes. Reads nothing itself: every
        /// input is handed in, so the gate, a connection and health all say the same words.
        /// </summary>
        public static OutlookHangFacts FactsFor(
            OutlookStartVerdict verdict,
            OutlookStartObservation observation,
            OutlookProcessNote? startedNote,
            OutlookProcessNote? notRespondingNote,
            OutlookProcessNote? windowSeenNote,
            string logPath)
        {
            if (verdict == null)
            {
                throw new ArgumentNullException(nameof(verdict));
            }

            if (observation == null)
            {
                throw new ArgumentNullException(nameof(observation));
            }

            int? pid = observation.Windows.ProcessId;
            DateTime? start = observation.ProcessStartUtc;
            bool aboutThis(OutlookProcessNote? note) => note != null && note.IsAbout(pid, start);

            OutlookConnectRecord? connect = observation.Connect;
            bool connectAboutThis = connect != null
                && (connect.OutlookProcessId == null || connect.OutlookProcessId == pid);

            return new OutlookHangFacts(
                verdict.Stuck,
                verdict.Reason,
                pid,
                start,
                observation.Windows.HasVisibleWindow,
                observation.Windows.HasVisibleWindow || aboutThis(windowSeenNote),
                aboutThis(startedNote) ? startedNote!.ServerProcessId : null,
                aboutThis(notRespondingNote) ? notRespondingNote!.AtUtc : (DateTime?)null,
                connectAboutThis ? connect : null,
                observation.ConnectLockHeldElsewhere,
                observation.ObservedUtc,
                logPath);
        }

        /// <summary>A UTC time a person reads: <c>2026-11-15 15:27:39 UTC</c>.</summary>
        internal static string Utc(DateTime utc)
        {
            return DateTime.SpecifyKind(utc, DateTimeKind.Utc).ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " UTC";
        }

        /// <summary>An age a person reads: seconds under two minutes, minutes under two hours, hours after.</summary>
        internal static string Ago(TimeSpan span)
        {
            double seconds = Math.Max(0, span.TotalSeconds);
            if (seconds < 120)
            {
                return ((long)seconds).ToString(CultureInfo.InvariantCulture) + " s";
            }

            if (seconds < 7200)
            {
                return ((long)(seconds / 60)).ToString(CultureInfo.InvariantCulture) + " min";
            }

            return ((long)(seconds / 3600)).ToString(CultureInfo.InvariantCulture) + " h";
        }
    }
}
