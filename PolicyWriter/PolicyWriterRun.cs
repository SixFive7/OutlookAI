using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security;
using OutlookAI.Services;

namespace OutlookAI.PolicyWriter
{
    /// <summary>
    /// What the helper needs from Windows, and nothing more - so the order of its checks can be
    /// tested without an administrator, a second user or a registry (the test project links this
    /// file and supplies a fake). <c>WindowsPolicyHost</c> is the real one.
    /// </summary>
    internal interface IPolicyWriterHost
    {
        /// <summary>
        /// The SID of the user whose process started this one, when that process still runs in
        /// THIS process's Windows session; empty, with <paramref name="why"/>, when that cannot be
        /// established. Never this process's own user - under an administrator's credentials that
        /// is the administrator, which is the whole problem.
        /// </summary>
        string RequesterSid(out string why);

        /// <summary>True when <c>HKEY_USERS\&lt;sid&gt;</c> is loaded - its user is logged on. The helper never loads a hive.</summary>
        bool IsUserHiveLoaded(string sid);

        /// <summary>Writes one REG_DWORD under <c>HKEY_USERS\&lt;sid&gt;\&lt;keyPath&gt;</c>, creating the key if it is missing.</summary>
        void WriteDword(string sid, string keyPath, string name, int value);

        /// <summary>Reads one value back from the same place: the number when it is a REG_DWORD, otherwise null.</summary>
        int? ReadDword(string sid, string keyPath, string name);
    }

    /// <summary>
    /// THE HELPER, START TO FINISH, in the order that makes it safe to run elevated:
    /// <list type="number">
    ///   <item>The command line must be exactly a request the add-in makes
    ///   (<see cref="PolicyWriterCommandLine.Parse"/>): the five names, their offered values, one
    ///   SID, one supported Office major. Anything else is refused WHOLE.</item>
    ///   <item>The SID must be the user who started the helper - read from the starting process,
    ///   which must be in the helper's own session - because a UAC approval with an
    ///   administrator's credentials runs the helper AS that administrator, and the request may
    ///   only ever touch the requester's own hive.</item>
    ///   <item>That user's hive must already be loaded. Nothing here loads one.</item>
    ///   <item>Only then: write each value under <c>HKEY_USERS\&lt;sid&gt;\</c>
    ///   <see cref="CachedModePolicy.KeyPath"/> - never HKCU - and read each one back.</item>
    /// </list>
    /// Every refusal happens before the first write. The exit code is the result
    /// (<see cref="PolicyWriterExit"/>), because the add-in starts the helper through
    /// ShellExecuteEx and has no pipe to read; the log line is for a person running it by hand.
    /// </summary>
    internal static class PolicyWriterRun
    {
        internal static int Run(string[] args, IPolicyWriterHost host, TextWriter log)
        {
            PolicyWriteParse parse = PolicyWriterCommandLine.Parse(args);
            if (!parse.Ok)
            {
                log.WriteLine("REFUSED (arguments), nothing written: " + parse.Refusal);
                return PolicyWriterExit.RefusedArguments;
            }

            PolicyWriteRequest request = parse.Request;

            string why;
            string requester = host.RequesterSid(out why);
            if (string.IsNullOrEmpty(requester))
            {
                log.WriteLine("REFUSED (user), nothing written: could not establish who started this helper - " + why);
                return PolicyWriterExit.RefusedUser;
            }

            if (!string.Equals(requester, request.Sid, StringComparison.Ordinal))
            {
                log.WriteLine("REFUSED (user), nothing written: the request names " + request.Sid +
                              ", but the process that started this helper runs as " + requester +
                              ". The helper only ever writes its requester's own hive.");
                return PolicyWriterExit.RefusedUser;
            }

            if (!host.IsUserHiveLoaded(request.Sid))
            {
                log.WriteLine("REFUSED (hive), nothing written: HKEY_USERS\\" + request.Sid + " is not loaded. This helper never loads a hive.");
                return PolicyWriterExit.RefusedHiveNotLoaded;
            }

            string where = "HKEY_USERS\\" + request.Sid + "\\" + request.KeyPath;
            var written = new List<string>();
            foreach (KeyValuePair<string, int> v in request.Values)
            {
                try
                {
                    host.WriteDword(request.Sid, request.KeyPath, v.Key, v.Value);
                    written.Add(v.Key);
                }
                catch (Exception ex) when (ex is UnauthorizedAccessException || ex is SecurityException)
                {
                    log.WriteLine("ACCESS DENIED writing " + v.Key + " under " + where + " - not running as an administrator? Written before it: " +
                                  Joined(written) + ". " + ex.Message);
                    return PolicyWriterExit.AccessDenied;
                }
                catch (Exception ex) when (!(ex is OutOfMemoryException))
                {
                    log.WriteLine("FAILED writing " + v.Key + " under " + where + ". Written before it: " + Joined(written) + ". " +
                                  ex.GetType().Name + ": " + ex.Message);
                    return PolicyWriterExit.Failed;
                }
            }

            foreach (KeyValuePair<string, int> v in request.Values)
            {
                int? back = host.ReadDword(request.Sid, request.KeyPath, v.Key);
                if (!back.HasValue || back.Value != v.Value)
                {
                    string shown = back.HasValue ? back.Value.ToString(CultureInfo.InvariantCulture) : "(not a REG_DWORD)";
                    log.WriteLine("READ BACK " + v.Key + " = " + shown + ", not the " + v.Value.ToString(CultureInfo.InvariantCulture) +
                                  " just written under " + where + ".");
                    return PolicyWriterExit.ReadBackMismatch;
                }
            }

            log.WriteLine("WRITTEN under " + where + ": " + Describe(request));
            return PolicyWriterExit.Written;
        }

        private static string Joined(List<string> names)
        {
            return names.Count == 0 ? "nothing" : string.Join(", ", names);
        }

        private static string Describe(PolicyWriteRequest request)
        {
            var parts = new List<string>();
            foreach (KeyValuePair<string, int> v in request.Values)
                parts.Add(v.Key + "=" + v.Value.ToString(CultureInfo.InvariantCulture));
            return string.Join(", ", parts);
        }
    }
}
