using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OutlookAI.Services
{
    /// <summary>
    /// What the add-in asks the elevated helper (<c>OutlookAI.PolicyWriter.exe</c>) to do, as a
    /// command line - and the helper's reading of it, which REFUSES everything that is not exactly
    /// such a request.
    ///
    /// <para>
    /// THE HELPER IS A PRIVILEGE BOUNDARY. A user approves its UAC prompt once, and from then on it
    /// runs with an administrator's token; whoever controls its command line controls what that token
    /// writes. So the command line is not parsed generously. It takes:
    /// <code>
    ///   --sid &lt;the original user's SID&gt;  --office &lt;Office major&gt;  Name=Value [Name=Value ...]
    /// </code>
    /// and nothing else: no path, no hive, no value type, no free text. The key is implied - always
    /// <c>HKEY_USERS\&lt;sid&gt;\</c> + <see cref="CachedModePolicy.KeyPath"/> - and every name and
    /// value comes from the closed list in <see cref="CachedModePolicy"/>. A request that deviates in
    /// any way is refused WHOLE, before anything is written.
    /// </para>
    ///
    /// <para>
    /// WHY A SID AND NOT HKCU. A standard user who approves UAC with an ADMINISTRATOR's credentials
    /// gets an elevated process running AS THAT ADMINISTRATOR, whose <c>HKEY_CURRENT_USER</c> is the
    /// administrator's own hive. Writing HKCU there would set the administrator's Outlook policy and
    /// leave the user's untouched - and report success. So the request names the user explicitly, the
    /// helper checks that it IS the user who started it (see the helper's
    /// <c>PolicyWriterRun</c>), and it writes under <c>HKEY_USERS\&lt;sid&gt;</c>. An administrator
    /// with a split token is the same user either way, so the same path serves both.
    /// </para>
    ///
    /// <para>
    /// The add-in BUILDS with <see cref="Build"/> and the helper PARSES with <see cref="Parse"/>, from
    /// this one file, so the two cannot disagree about the format; the test project links it and
    /// pins the round trip and every refusal.
    /// </para>
    ///
    /// <para>
    /// FRAMEWORK-NEUTRAL and INTERNAL, like <see cref="CachedModePolicy"/>: net48 (C# 7.3) and net10
    /// (nullable-enabled, warnings as errors), so no nullable annotations and nothing returned or
    /// assigned as null.
    /// </para>
    /// </summary>
    internal static class PolicyWriterCommandLine
    {
        /// <summary>The helper's file name, next to the add-in's own assembly in every layout.</summary>
        internal const string ExecutableName = "OutlookAI.PolicyWriter.exe";

        internal const string SidSwitch = "--sid";
        internal const string OfficeSwitch = "--office";

        /// <summary>
        /// The command line for one request, in the exact shape <see cref="Parse"/> accepts. Validates
        /// as strictly as the helper does, so the add-in never launches - and never puts a UAC prompt
        /// up for - a request the helper would refuse. Throws <see cref="ArgumentException"/> instead.
        /// </summary>
        internal static string Build(string sid, string officeVersion, IList<KeyValuePair<string, int>> values)
        {
            if (values == null || values.Count == 0)
                throw new ArgumentException("A request names at least one value.", nameof(values));

            var tokens = new List<string> { SidSwitch, sid ?? string.Empty, OfficeSwitch, officeVersion ?? string.Empty };
            foreach (KeyValuePair<string, int> v in values)
                tokens.Add(v.Key + "=" + v.Value.ToString(CultureInfo.InvariantCulture));

            PolicyWriteParse check = Parse(tokens.ToArray());
            if (!check.Ok)
                throw new ArgumentException("The helper would refuse this request: " + check.Refusal);

            // Every token is drawn from [0-9A-Za-z.=S-] by the validation above, so none needs
            // quoting - which is also why the helper can refuse any token that would.
            var sb = new StringBuilder();
            foreach (string t in tokens)
            {
                if (sb.Length > 0)
                    sb.Append(' ');
                sb.Append(t);
            }
            return sb.ToString();
        }

        /// <summary>
        /// Reads a command line, or says why it is not a request this product makes. Pure: no
        /// registry, no process, no identity - those checks are the helper's, after this.
        /// </summary>
        internal static PolicyWriteParse Parse(string[] args)
        {
            if (args == null || args.Length == 0)
                return PolicyWriteParse.Refuse("no arguments. Usage: " + SidSwitch + " <SID> " + OfficeSwitch + " <major> Name=Value [Name=Value ...]");

            string sid = string.Empty;
            string office = string.Empty;
            bool haveSid = false;
            bool haveOffice = false;
            var values = new List<KeyValuePair<string, int>>();
            var seen = new HashSet<string>(StringComparer.Ordinal);

            for (int i = 0; i < args.Length; i++)
            {
                string token = args[i] ?? string.Empty;
                if (string.Equals(token, SidSwitch, StringComparison.Ordinal) ||
                    string.Equals(token, OfficeSwitch, StringComparison.Ordinal))
                {
                    if (i + 1 >= args.Length)
                        return PolicyWriteParse.Refuse(token + " has no value after it.");
                    string value = args[++i] ?? string.Empty;
                    if (string.Equals(token, SidSwitch, StringComparison.Ordinal))
                    {
                        if (haveSid)
                            return PolicyWriteParse.Refuse(SidSwitch + " is given twice.");
                        if (!IsUserAccountSid(value))
                            return PolicyWriteParse.Refuse(SidSwitch + " '" + Shown(value) + "' is not a user account SID (S-1-5-21-... or S-1-12-1-..., in canonical form).");
                        sid = value;
                        haveSid = true;
                    }
                    else
                    {
                        if (haveOffice)
                            return PolicyWriteParse.Refuse(OfficeSwitch + " is given twice.");
                        if (!CachedModePolicy.IsSupportedOfficeVersion(value))
                            return PolicyWriteParse.Refuse(OfficeSwitch + " '" + Shown(value) + "' is not an Office major this product supports (" + string.Join(", ", OfficeVersions.Supported) + ").");
                        office = value;
                        haveOffice = true;
                    }
                    continue;
                }

                int eq = token.IndexOf('=');
                if (eq <= 0 || eq != token.LastIndexOf('='))
                    return PolicyWriteParse.Refuse("'" + Shown(token) + "' is neither " + SidSwitch + ", " + OfficeSwitch + " nor Name=Value.");
                string name = token.Substring(0, eq);
                string number = token.Substring(eq + 1);
                if (!CachedModePolicy.IsKnownName(name))
                    return PolicyWriteParse.Refuse("'" + Shown(name) + "' is not one of the five Cached Mode policy values (" + string.Join(", ", CachedModePolicy.ValueNames) + ").");
                if (!seen.Add(name))
                    return PolicyWriteParse.Refuse(name + " is given twice.");
                int parsed;
                if (!TryParseCanonicalNumber(number, out parsed))
                    return PolicyWriteParse.Refuse(name + "='" + Shown(number) + "': not a plain decimal number.");
                if (!CachedModePolicy.IsAllowed(name, parsed))
                    return PolicyWriteParse.Refuse(name + "=" + number + ": not a value OutlookAI offers for it.");
                values.Add(new KeyValuePair<string, int>(name, parsed));
            }

            if (!haveSid)
                return PolicyWriteParse.Refuse(SidSwitch + " is missing: the helper writes one named user's hive and never its own.");
            if (!haveOffice)
                return PolicyWriteParse.Refuse(OfficeSwitch + " is missing.");
            if (values.Count == 0)
                return PolicyWriteParse.Refuse("no Name=Value to write.");

            return PolicyWriteParse.Accept(new PolicyWriteRequest(sid, office, values));
        }

        /// <summary>
        /// A user account SID, as text, in the canonical form Windows itself prints: <c>S-1-5-21-</c>
        /// plus four sub-authorities (a local or domain account) or <c>S-1-12-1-</c> plus four (an
        /// Azure AD account). Every other shape is refused - well-known and service SIDs (SYSTEM is
        /// <c>S-1-5-18</c>), the <c>_Classes</c> hive's name, anything with a leading zero, a sign,
        /// whitespace or a non-ASCII digit - because it would not name a person's own hive under
        /// <c>HKEY_USERS</c>, or would name one in a spelling Windows never uses.
        /// </summary>
        internal static bool IsUserAccountSid(string text)
        {
            if (string.IsNullOrEmpty(text) || text.Length > 80)
                return false;
            string[] parts = text.Split('-');
            // S, 1, authority, then the sub-authorities.
            if (parts.Length != 8 || !string.Equals(parts[0], "S", StringComparison.Ordinal) || !string.Equals(parts[1], "1", StringComparison.Ordinal))
                return false;
            bool ntAccount = string.Equals(parts[2], "5", StringComparison.Ordinal) && string.Equals(parts[3], "21", StringComparison.Ordinal);
            bool azureAd = string.Equals(parts[2], "12", StringComparison.Ordinal) && string.Equals(parts[3], "1", StringComparison.Ordinal);
            if (!ntAccount && !azureAd)
                return false;
            for (int i = 4; i < parts.Length; i++)
            {
                uint ignored;
                if (!TryParseCanonicalUInt(parts[i], out ignored))
                    return false;
            }
            return true;
        }

        /// <summary>A non-negative decimal int in canonical form: ASCII digits, no sign, no leading zero, no whitespace.</summary>
        internal static bool TryParseCanonicalNumber(string text, out int value)
        {
            value = 0;
            uint u;
            if (!TryParseCanonicalUInt(text, out u) || u > int.MaxValue)
                return false;
            value = (int)u;
            return true;
        }

        private static bool TryParseCanonicalUInt(string text, out uint value)
        {
            value = 0;
            if (string.IsNullOrEmpty(text) || text.Length > 10)
                return false;
            foreach (char c in text)
            {
                // ASCII only: char.IsDigit would let Arabic-Indic and full-width digits through.
                if (c < '0' || c > '9')
                    return false;
            }
            if (text.Length > 1 && text[0] == '0')
                return false;
            return uint.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out value);
        }

        /// <summary>A token as it may appear in a refusal: printable ASCII, and short.</summary>
        private static string Shown(string text)
        {
            if (string.IsNullOrEmpty(text))
                return string.Empty;
            var sb = new StringBuilder();
            foreach (char c in text)
            {
                if (sb.Length >= 60)
                {
                    sb.Append("...");
                    break;
                }
                sb.Append(c >= ' ' && c <= '~' ? c : '?');
            }
            return sb.ToString();
        }
    }

    /// <summary>One validated request: whose hive, which Office, and which of the five values.</summary>
    internal sealed class PolicyWriteRequest
    {
        internal PolicyWriteRequest(string sid, string officeVersion, IList<KeyValuePair<string, int>> values)
        {
            Sid = sid;
            OfficeVersion = officeVersion;
            Values = new List<KeyValuePair<string, int>>(values);
        }

        /// <summary>The user whose hive is written - <c>HKEY_USERS\&lt;Sid&gt;</c>, never HKCU.</summary>
        internal string Sid { get; }

        internal string OfficeVersion { get; }

        /// <summary>The key under that user's hive root - always <see cref="CachedModePolicy.KeyPath"/>.</summary>
        internal string KeyPath
        {
            get { return CachedModePolicy.KeyPath(OfficeVersion); }
        }

        internal IReadOnlyList<KeyValuePair<string, int>> Values { get; }
    }

    /// <summary>What <see cref="PolicyWriterCommandLine.Parse"/> made of a command line.</summary>
    internal sealed class PolicyWriteParse
    {
        private static readonly PolicyWriteRequest NoRequest =
            new PolicyWriteRequest(string.Empty, string.Empty, new List<KeyValuePair<string, int>>());

        private PolicyWriteParse(bool ok, PolicyWriteRequest request, string refusal)
        {
            Ok = ok;
            Request = request;
            Refusal = refusal;
        }

        internal bool Ok { get; }

        /// <summary>The request when <see cref="Ok"/>; an empty one otherwise, never to be written.</summary>
        internal PolicyWriteRequest Request { get; }

        /// <summary>Why it was refused; empty when <see cref="Ok"/>.</summary>
        internal string Refusal { get; }

        internal static PolicyWriteParse Accept(PolicyWriteRequest request)
        {
            return new PolicyWriteParse(true, request, string.Empty);
        }

        internal static PolicyWriteParse Refuse(string why)
        {
            return new PolicyWriteParse(false, NoRequest, why);
        }
    }

    /// <summary>
    /// The helper's exit codes, and what the add-in tells the user about each. A closed list: the
    /// helper never exits with anything else, so an unknown code means something other than the
    /// helper ran.
    /// </summary>
    internal static class PolicyWriterExit
    {
        /// <summary>Every value written under the user's hive and read back equal.</summary>
        internal const int Written = 0;

        /// <summary>Something unexpected - not a refusal, not access denied.</summary>
        internal const int Failed = 1;

        /// <summary>The command line is not a request this product makes. Nothing written.</summary>
        internal const int RefusedArguments = 2;

        /// <summary>The SID is not the user whose process started the helper, or that could not be established. Nothing written.</summary>
        internal const int RefusedUser = 3;

        /// <summary>That user's registry hive is not loaded; the helper never loads one. Nothing written.</summary>
        internal const int RefusedHiveNotLoaded = 4;

        /// <summary>Windows refused the write - the helper was not running as an administrator.</summary>
        internal const int AccessDenied = 5;

        /// <summary>Written, but a value read back differently - something else is writing the key.</summary>
        internal const int ReadBackMismatch = 6;

        internal static string Describe(int code)
        {
            switch (code)
            {
                case Written: return "the values were written to your Outlook policy settings";
                case Failed: return "the helper failed unexpectedly";
                case RefusedArguments: return "the helper refused the request as malformed";
                case RefusedUser: return "the helper could not confirm that the request came from you, so it changed nothing";
                case RefusedHiveNotLoaded: return "your registry hive was not loaded, so the helper changed nothing";
                case AccessDenied: return "Windows refused the write - the helper did not run as an administrator";
                case ReadBackMismatch: return "the values were written but did not read back the same - something else manages them";
                default: return "the helper exited with an unknown code " + code.ToString(CultureInfo.InvariantCulture);
            }
        }
    }
}
