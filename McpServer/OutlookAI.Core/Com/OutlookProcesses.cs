using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;

namespace OutlookAI.Core.Com
{
    /// <summary>
    /// The OUTLOOK.EXE processes that are this user's Outlook: the ones in the caller's own Windows session
    /// (decided 2026-11-18). Windows runs an Outlook per signed-in session, and another user signed in to the
    /// same machine - by fast user switching, or on a terminal server - has their own. Theirs cannot be
    /// attached to from here and its windows are on another desktop. Counted as this user's, it read as an
    /// Outlook still starting with no windows, then after two minutes as a stuck one - with the other user's
    /// process ID for this user to end. Every OUTLOOK.EXE lookup in the server goes through this class.
    /// </summary>
    public static class OutlookProcesses
    {
        private const string ProcessName = "OUTLOOK";

        private static readonly Lazy<int?> CallerSession = new Lazy<int?>(ReadCallerSession);

        /// <summary>This process's Windows session; null when it could not be read.</summary>
        public static int? CallerSessionId => CallerSession.Value;

        /// <summary>
        /// This session's OUTLOOK.EXE processes, in the order Windows listed them; the caller disposes them.
        /// The rule is <see cref="InSession"/>'s.
        /// </summary>
        public static Process[] InThisSession()
        {
            Process[] all = Process.GetProcessesByName(ProcessName);
            List<KeyValuePair<int, int?>> sessions = new List<KeyValuePair<int, int?>>(all.Length);
            foreach (Process process in all)
            {
                sessions.Add(new KeyValuePair<int, int?>(process.Id, SessionOf(process)));
            }

            HashSet<int> mine = new HashSet<int>(InSession(sessions, CallerSessionId));
            List<Process> kept = new List<Process>(mine.Count);
            foreach (Process process in all)
            {
                if (mine.Contains(process.Id))
                {
                    kept.Add(process);
                }
                else
                {
                    process.Dispose();
                }
            }

            return kept.ToArray();
        }

        /// <summary>Whether this session has an OUTLOOK.EXE.</summary>
        public static bool AnyInThisSession()
        {
            Process[] mine = InThisSession();
            try
            {
                return mine.Length > 0;
            }
            finally
            {
                foreach (Process process in mine)
                {
                    process.Dispose();
                }
            }
        }

        /// <summary>
        /// Of <paramref name="processes"/> - each an id and its session, null where it could not be read - the
        /// ids in <paramref name="callerSession"/>, in order. A process whose session could not be read is left
        /// out: this user's own Outlook can always be read. With the caller's own session unknown, every one
        /// is kept, as before the rule, rather than reporting no Outlook at all.
        /// </summary>
        internal static List<int> InSession(IEnumerable<KeyValuePair<int, int?>> processes, int? callerSession)
        {
            List<int> ids = new List<int>();
            foreach (KeyValuePair<int, int?> process in processes)
            {
                if (callerSession == null || process.Value == callerSession)
                {
                    ids.Add(process.Key);
                }
            }

            return ids;
        }

        private static int? SessionOf(Process process)
        {
            try
            {
                return process.SessionId;
            }
            catch (Exception ex) when (ex is InvalidOperationException || ex is Win32Exception || ex is NotSupportedException)
            {
                // Gone since the listing, or not readable from this session.
                return null;
            }
        }

        private static int? ReadCallerSession()
        {
            try
            {
                using (Process self = Process.GetCurrentProcess())
                {
                    return self.SessionId;
                }
            }
            catch (Exception ex) when (ex is InvalidOperationException || ex is Win32Exception || ex is NotSupportedException)
            {
                return null;
            }
        }
    }
}
