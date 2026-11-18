using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Microsoft.Win32;

using OfficeVersions = global::OutlookAI.Services.OfficeVersions;
using SignatureFolderLocation = global::OutlookAI.Services.SignatureFolderLocation;

namespace OutlookAI.Core.Services
{
    /// <summary>Which registry hive a signature setting lives in.</summary>
    public enum SignatureRegistryHive
    {
        /// <summary>HKEY_CURRENT_USER - every per-user Outlook and Office setting, policy included.</summary>
        CurrentUser,

        /// <summary>HKEY_LOCAL_MACHINE - only the Click-to-Run configuration (the Office build and product).</summary>
        LocalMachine,
    }

    /// <summary>
    /// Read-only access to the registry values the signature model reads - the T1 seam. The live
    /// implementation is <see cref="LiveSignatureSettingsRegistry"/>; nothing here ever writes, and
    /// no policy value is ever written by this product.
    /// </summary>
    public interface ISignatureSettingsRegistry
    {
        /// <summary>
        /// A value as stored - a string for REG_SZ (REG_EXPAND_SZ expanded), an int for REG_DWORD,
        /// a byte array for REG_BINARY - or null when the key or the value is absent or unreadable.
        /// </summary>
        object? ReadValue(SignatureRegistryHive hive, string keyPath, string valueName);
    }

    /// <summary>The live registry, read-only. Unreadable counts as absent: a key that cannot be read is not one we can honour.</summary>
    public sealed class LiveSignatureSettingsRegistry : ISignatureSettingsRegistry
    {
        /// <inheritdoc />
        public object? ReadValue(SignatureRegistryHive hive, string keyPath, string valueName)
        {
            try
            {
                RegistryKey root = hive == SignatureRegistryHive.LocalMachine ? Registry.LocalMachine : Registry.CurrentUser;
                using RegistryKey? key = root.OpenSubKey(keyPath, writable: false);
                return key?.GetValue(valueName);
            }
            catch (Exception ex) when (ex is System.Security.SecurityException or UnauthorizedAccessException or System.IO.IOException)
            {
                return null;
            }
        }
    }

    /// <summary>Where the signature files are, and why there (list_signatures' <c>folder</c>).</summary>
    public sealed class SignatureFolderInfo
    {
        /// <summary>Creates the description.</summary>
        public SignatureFolderInfo(string path, bool fromRegistry, string? recordedValue)
        {
            Path = path;
            FromRegistry = fromRegistry;
            RecordedValue = recordedValue;
        }

        /// <summary>The folder, absolute.</summary>
        public string Path { get; }

        /// <summary>True when Outlook's recorded <c>Common\General\Signatures</c> value named it; false for the default.</summary>
        public bool FromRegistry { get; }

        /// <summary>The recorded value as written, when there is one.</summary>
        public string? RecordedValue { get; }
    }

    /// <summary>
    /// The signature folder, resolved fresh on every call from the value Outlook records - the one
    /// rule the add-in and the server share (<see cref="SignatureFolderLocation"/>).
    /// </summary>
    public static class SignatureFolder
    {
        /// <summary>HKCU path of the key whose <c>Signatures</c> value names the folder, for the detected Office.</summary>
        public static string RegistryKeyPath => SignatureFolderLocation.GeneralKeyPath(OutlookProfileRegistry.OfficeVersion);

        /// <summary>The value name (<c>Signatures</c>).</summary>
        public const string RegistryValueName = SignatureFolderLocation.ValueName;

        /// <summary>The folder Outlook uses now.</summary>
        public static string CurrentPath => Current().Path;

        /// <summary>The folder Outlook uses now, with where that came from.</summary>
        public static SignatureFolderInfo Current()
        {
            return Read(new LiveSignatureSettingsRegistry(), OutlookProfileRegistry.OfficeVersion,
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData));
        }

        /// <summary>The folder a registry names, for one Office major and one %APPDATA%. Pure over the seam.</summary>
        public static SignatureFolderInfo Read(ISignatureSettingsRegistry registry, string officeVersion, string appDataRoaming)
        {
            if (registry == null)
            {
                throw new ArgumentNullException(nameof(registry));
            }

            string? recorded = registry.ReadValue(
                SignatureRegistryHive.CurrentUser, SignatureFolderLocation.GeneralKeyPath(officeVersion), SignatureFolderLocation.ValueName) as string;
            string path = SignatureFolderLocation.Resolve(appDataRoaming, recorded ?? string.Empty, out bool fromRegistry);
            return new SignatureFolderInfo(path, fromRegistry, string.IsNullOrWhiteSpace(recorded) ? null : recorded);
        }
    }

    /// <summary>Where a signature policy value came from, strongest first.</summary>
    public enum SignaturePolicySource
    {
        /// <summary>Cloud Policy (<c>HKCU\Software\Policies\Microsoft\Cloud\Office\&lt;major&gt;</c>) - overrides Group Policy.</summary>
        CloudPolicy,

        /// <summary>Group Policy (<c>HKCU\Software\Policies\Microsoft\Office\&lt;major&gt;</c>).</summary>
        GroupPolicy,

        /// <summary>
        /// A deployed preference (<c>HKCU\Software\Microsoft\Office\&lt;major&gt;\Common\MailSettings</c>) - not a
        /// policy: Outlook applies it as the default for every account, and nothing locks it.
        /// </summary>
        Preference,
    }

    /// <summary>
    /// The administrator's signature rules (Q156), as far as the registry carries them. Read-only:
    /// this product never writes a policy or a deployed preference.
    /// <list type="bullet">
    /// <item><c>DisableSignatures</c> (DWORD 1, the Group Policy "Do not allow signatures for e-mail
    /// messages") - no signature at all.</item>
    /// <item><c>NewSignature</c> / <c>ReplySignature</c> (a string, a signature name) - a default an administrator
    /// names for new mail and for replies and forwards. From a policy source it is FORCED - OutlookAI does not change
    /// an account's default under it (Q156); from the preference key it is a deployed default.</item>
    /// </list>
    /// Precedence among them: Cloud Policy over Group Policy (documented, ms-cloudpol) over the preference. But they
    /// do not decide which signature a draft gets. MEASURED 2026-11-18 on Office LTSC 2024 (Q200 (a), Q201;
    /// <c>Docs/live-tier-on-the-vm.md</c> section 4.11), with the values set before Outlook started: Outlook put an
    /// account's own default in a new mail, a reply and a forward, not theirs - as REG_SZ and as REG_EXPAND_SZ - put
    /// nothing in a reply or a forward of an account with no reply default of its own, under the policy and under the
    /// deployed setting alike, and kept its Signature button open (<see cref="EffectiveDefault"/>).
    /// </summary>
    public sealed class SignaturePolicy
    {
        /// <summary>The value forbidding signatures.</summary>
        public const string DisableSignaturesValueName = "DisableSignatures";

        /// <summary>The value naming the new-mail default.</summary>
        public const string NewSignatureValueName = "NewSignature";

        /// <summary>The value naming the reply/forward default.</summary>
        public const string ReplySignatureValueName = "ReplySignature";

        /// <summary>No policy and no preference: per-account defaults decide.</summary>
        public static readonly SignaturePolicy None = new SignaturePolicy(false, null, null, null, null, null);

        /// <summary>Creates the policy view.</summary>
        public SignaturePolicy(
            bool signaturesDisabled,
            SignaturePolicySource? disabledBy,
            string? newMessage,
            SignaturePolicySource? newMessageSource,
            string? replyForward,
            SignaturePolicySource? replyForwardSource)
        {
            SignaturesDisabled = signaturesDisabled;
            DisabledBy = disabledBy;
            NewMessage = newMessage;
            NewMessageSource = newMessageSource;
            ReplyForward = replyForward;
            ReplyForwardSource = replyForwardSource;
        }

        /// <summary>True when policy forbids signatures altogether.</summary>
        public bool SignaturesDisabled { get; }

        /// <summary>Which policy forbids them.</summary>
        public SignaturePolicySource? DisabledBy { get; }

        /// <summary>The new-mail default the policy or preference names, for every account.</summary>
        public string? NewMessage { get; }

        /// <summary>Where <see cref="NewMessage"/> came from.</summary>
        public SignaturePolicySource? NewMessageSource { get; }

        /// <summary>The reply/forward default the policy or preference names, for every account.</summary>
        public string? ReplyForward { get; }

        /// <summary>Where <see cref="ReplyForward"/> came from.</summary>
        public SignaturePolicySource? ReplyForwardSource { get; }

        /// <summary>True when the new-mail default is forced by a POLICY (not a preference).</summary>
        public bool NewMessageForced => NewMessage != null && IsPolicy(NewMessageSource);

        /// <summary>True when the reply/forward default is forced by a POLICY.</summary>
        public bool ReplyForwardForced => ReplyForward != null && IsPolicy(ReplyForwardSource);

        /// <summary>True when anything here takes defaults out of the per-account values' hands.</summary>
        public bool AnyDefaultSet => NewMessage != null || ReplyForward != null;

        /// <summary>The default named for one scope (new mail or replies), or null.</summary>
        public string? DefaultFor(bool replyOrForward) => replyOrForward ? ReplyForward : NewMessage;

        /// <summary>Where that default came from.</summary>
        public SignaturePolicySource? SourceFor(bool replyOrForward) => replyOrForward ? ReplyForwardSource : NewMessageSource;

        /// <summary>True when that default is forced by policy.</summary>
        public bool ForcedFor(bool replyOrForward) => replyOrForward ? ReplyForwardForced : NewMessageForced;

        /// <summary>
        /// The default a draft of an account gets for one scope, given the account's own (null for none, or for no
        /// account): the account's own, whatever a policy or deployed setting names - what Outlook does (measured
        /// 2026-11-18, Q200 (a), Q201; see the class). The ONE rule the draft tools (<see cref="SignatureSelector"/>),
        /// list_signatures and manage_signature share (Q154), so a later Outlook that uses these values changes one place.
        /// </summary>
        public string? EffectiveDefault(bool replyOrForward, string? accountsOwn) => accountsOwn;

        /// <summary>
        /// True when a POLICY names that scope's default, so manage_signature refuses to change the account's own (Q156:
        /// a change of default is refused under a forced default). A deployed preference does not: the account's own
        /// default ranks above it, so changing it changes what a draft gets (Q201).
        /// </summary>
        public bool TakesDefault(bool replyOrForward) => ForcedFor(replyOrForward);

        /// <summary>The source in words, for answers and refusals.</summary>
        public static string Describe(SignaturePolicySource? source)
        {
            switch (source)
            {
                case SignaturePolicySource.CloudPolicy:
                    return "Cloud Policy";
                case SignaturePolicySource.GroupPolicy:
                    return "Group Policy";
                case SignaturePolicySource.Preference:
                    return "a deployed Outlook setting (Common\\MailSettings)";
                default:
                    return "no policy";
            }
        }

        /// <summary>The short token the answers carry: cloudPolicy, groupPolicy, preference.</summary>
        public static string? Token(SignaturePolicySource? source)
        {
            switch (source)
            {
                case SignaturePolicySource.CloudPolicy:
                    return "cloudPolicy";
                case SignaturePolicySource.GroupPolicy:
                    return "groupPolicy";
                case SignaturePolicySource.Preference:
                    return "preference";
                default:
                    return null;
            }
        }

        /// <summary>HKCU path of the MailSettings key of one source.</summary>
        public static string KeyPath(SignaturePolicySource source, string officeVersion)
        {
            switch (source)
            {
                case SignaturePolicySource.CloudPolicy:
                    return @"Software\Policies\Microsoft\Cloud\Office\" + officeVersion + @"\Common\MailSettings";
                case SignaturePolicySource.GroupPolicy:
                    return @"Software\Policies\Microsoft\Office\" + officeVersion + @"\Common\MailSettings";
                default:
                    return OfficeVersions.CommonKeyPath(officeVersion) + @"\MailSettings";
            }
        }

        /// <summary>Reads the policy for one Office major. Never throws for an absent key: absent is no rule.</summary>
        public static SignaturePolicy Read(ISignatureSettingsRegistry registry, string officeVersion)
        {
            if (registry == null)
            {
                throw new ArgumentNullException(nameof(registry));
            }

            SignaturePolicySource[] order = { SignaturePolicySource.CloudPolicy, SignaturePolicySource.GroupPolicy, SignaturePolicySource.Preference };
            bool disabled = false;
            SignaturePolicySource? disabledBy = null;
            string? newMessage = null;
            SignaturePolicySource? newSource = null;
            string? reply = null;
            SignaturePolicySource? replySource = null;
            foreach (SignaturePolicySource source in order)
            {
                string key = KeyPath(source, officeVersion);

                // DisableSignatures is a POLICY: the preference location carries no such switch.
                if (!disabled && source != SignaturePolicySource.Preference
                    && IsOn(registry.ReadValue(SignatureRegistryHive.CurrentUser, key, DisableSignaturesValueName)))
                {
                    disabled = true;
                    disabledBy = source;
                }

                if (newMessage == null)
                {
                    newMessage = Name(registry.ReadValue(SignatureRegistryHive.CurrentUser, key, NewSignatureValueName));
                    newSource = newMessage != null ? source : null;
                }

                if (reply == null)
                {
                    reply = Name(registry.ReadValue(SignatureRegistryHive.CurrentUser, key, ReplySignatureValueName));
                    replySource = reply != null ? source : null;
                }
            }

            return new SignaturePolicy(disabled, disabledBy, newMessage, newSource, reply, replySource);
        }

        private static bool IsPolicy(SignaturePolicySource? source)
        {
            return source == SignaturePolicySource.CloudPolicy || source == SignaturePolicySource.GroupPolicy;
        }

        private static bool IsOn(object? value)
        {
            switch (value)
            {
                case int number:
                    return number != 0;
                case long wide:
                    return wide != 0;
                case string text:
                    return int.TryParse(text.Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed) && parsed != 0;
                case byte[] bytes:
                    return bytes.Any(b => b != 0);
                default:
                    return false;
            }
        }

        private static string? Name(object? value)
        {
            string? text = SignatureCatalog.DecodeRegistryString(value);
            return string.IsNullOrWhiteSpace(text) ? null : text!.Trim();
        }
    }

    /// <summary>The Office build and product, as Click-to-Run records them.</summary>
    public sealed class OfficeBuildInfo
    {
        /// <summary>The C2R configuration key under HKLM.</summary>
        public const string ClickToRunConfigurationKey = @"SOFTWARE\Microsoft\Office\ClickToRun\Configuration";

        /// <summary>Nothing known.</summary>
        public static readonly OfficeBuildInfo Unknown = new OfficeBuildInfo(null, null);

        /// <summary>Creates the description.</summary>
        public OfficeBuildInfo(string? version, string? productReleaseIds)
        {
            Version = string.IsNullOrWhiteSpace(version) ? null : version!.Trim();
            ProductReleaseIds = string.IsNullOrWhiteSpace(productReleaseIds) ? null : productReleaseIds!.Trim();
            Build = ParseBuild(Version);
        }

        /// <summary><c>VersionToReport</c>, e.g. 16.0.17932.20884.</summary>
        public string? Version { get; }

        /// <summary><c>ProductReleaseIds</c>, e.g. ProPlus2024Volume or O365ProPlusRetail.</summary>
        public string? ProductReleaseIds { get; }

        /// <summary>The build number - the third part of the version - or null.</summary>
        public int? Build { get; }

        /// <summary>True when a product id is a Microsoft 365 Apps subscription (O365... / Microsoft365...).</summary>
        public bool IsMicrosoft365 => ProductReleaseIds != null
            && ProductReleaseIds.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Any(id => id.Trim().StartsWith("O365", StringComparison.OrdinalIgnoreCase)
                    || id.Trim().StartsWith("Microsoft365", StringComparison.OrdinalIgnoreCase));

        /// <summary>Reads Click-to-Run's configuration. Unreadable is unknown.</summary>
        public static OfficeBuildInfo Read(ISignatureSettingsRegistry registry)
        {
            if (registry == null)
            {
                throw new ArgumentNullException(nameof(registry));
            }

            return new OfficeBuildInfo(
                registry.ReadValue(SignatureRegistryHive.LocalMachine, ClickToRunConfigurationKey, "VersionToReport") as string,
                registry.ReadValue(SignatureRegistryHive.LocalMachine, ClickToRunConfigurationKey, "ProductReleaseIds") as string);
        }

        private static int? ParseBuild(string? version)
        {
            if (version == null)
            {
                return null;
            }

            string[] parts = version.Split('.');
            return parts.Length >= 3 && int.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out int build)
                ? build
                : (int?)null;
        }
    }

    /// <summary>Whether this Outlook roams signatures through the mailbox (Q152).</summary>
    public enum SignatureRoamingState
    {
        /// <summary>
        /// It does not, as far as can be told: switched off, an Office that predates the feature, or a
        /// perpetual Office (LTSC, retail) showing no sign of it. The signature files on this machine are
        /// authoritative.
        /// </summary>
        Off,

        /// <summary>Microsoft 365 Apps on a build that roams: an Exchange Online account's signatures may live in its mailbox.</summary>
        Possible,

        /// <summary>
        /// Outlook keeps local copies of roaming signatures here - signatures named
        /// <c>Name (address)</c> after an Exchange account of this computer's profiles - whatever the
        /// Office: those accounts' signatures live in their mailboxes.
        /// </summary>
        Evident,
    }

    /// <summary>
    /// Roaming signatures, as far as this machine can say without the mailbox (Q152 (a)). Microsoft:
    /// "available for Outlook build 2303 or higher", for "mailboxes on Microsoft 365 or Outlook.com",
    /// switched off by <c>DisableRoamingSignatures</c> (DWORD 1) under <c>Outlook\Setup</c> or its policy
    /// copy; <c>DisableRoamingSignaturesTemporaryToggle</c> is the earlier name and "either key works"
    /// (slipstick). Outlook keeps a local copy of each roaming signature named
    /// <c>&lt;name&gt; (&lt;address&gt;)</c>.
    /// <para>
    /// <b>Generic on purpose (2026-11-17).</b> Office LTSC 2024 showed no sign of roaming on the
    /// Exchange test machine - Outlook's classic Signatures dialog, no account selector, no copies
    /// (stream A's R1) - but that is not proof, so a perpetual Office is OFF only until Outlook's own
    /// copies say otherwise: <see cref="WithEvidence"/> turns it EVIDENT for exactly the Exchange
    /// accounts that have them. Microsoft 365 Apps is POSSIBLE on Microsoft's word alone. Switched
    /// off, or an Office too old for the feature, stays off whatever files lie about.
    /// </para>
    /// </summary>
    public sealed class SignatureRoaming
    {
        /// <summary>The first build that roams: Version 2303 is build 16227.</summary>
        public const int FirstRoamingBuild = 16227;

        /// <summary>The switch, current name.</summary>
        public const string DisableValueName = "DisableRoamingSignatures";

        /// <summary>The switch, early-testing name - still honoured.</summary>
        public const string DisableTemporaryToggleValueName = "DisableRoamingSignaturesTemporaryToggle";

        /// <summary>Creates the view.</summary>
        public SignatureRoaming(
            SignatureRoamingState state,
            string reason,
            string? switchedOffAt,
            OfficeBuildInfo build,
            bool evidenceDecides = false,
            IReadOnlyList<string>? evidentAccounts = null)
        {
            State = state;
            Reason = reason;
            SwitchedOffAt = switchedOffAt;
            Build = build;
            EvidenceDecides = evidenceDecides;
            EvidentAccounts = evidentAccounts ?? Array.Empty<string>();
        }

        /// <summary>The verdict.</summary>
        public SignatureRoamingState State { get; }

        /// <summary>Why, in words.</summary>
        public string Reason { get; }

        /// <summary>The key and value that switched it off, when that is the reason.</summary>
        public string? SwitchedOffAt { get; }

        /// <summary>What Click-to-Run says this Office is.</summary>
        public OfficeBuildInfo Build { get; }

        /// <summary>True when Outlook's own local copies decide it - a perpetual Office, or an unreadable build.</summary>
        public bool EvidenceDecides { get; }

        /// <summary>The Exchange accounts with local copies of roaming signatures on this computer.</summary>
        public IReadOnlyList<string> EvidentAccounts { get; }

        /// <summary>True when any account's signatures may live in its mailbox.</summary>
        public bool MayRoam => State != SignatureRoamingState.Off;

        /// <summary>The state in the answers' vocabulary: off, possible, evident.</summary>
        public string Token => State == SignatureRoamingState.Off ? "off" : State == SignatureRoamingState.Possible ? "possible" : "evident";

        /// <summary>
        /// True when THIS account's signatures may live in its mailbox: every account where roaming is
        /// possible, only the accounts with copies where it is evident.
        /// </summary>
        public bool MayRoamFor(string? account)
        {
            return State == SignatureRoamingState.Possible
                || (State == SignatureRoamingState.Evident && EvidentAccounts.Any(a => ProfileAccountEntries.AddressEquals(a, account)));
        }

        /// <summary>HKCU paths of the two Setup keys that can carry the switch.</summary>
        public static IReadOnlyList<string> SwitchKeyPaths(string officeVersion)
        {
            return new[]
            {
                OfficeVersions.OutlookKeyPath(officeVersion) + @"\Setup",
                OfficeVersions.PolicyOutlookKeyPath(officeVersion) + @"\Setup",
            };
        }

        /// <summary>Reads and decides for one Office major, before any evidence.</summary>
        public static SignatureRoaming Read(ISignatureSettingsRegistry registry, string officeVersion)
        {
            if (registry == null)
            {
                throw new ArgumentNullException(nameof(registry));
            }

            string? switchedOffAt = null;
            foreach (string key in SwitchKeyPaths(officeVersion))
            {
                foreach (string value in new[] { DisableValueName, DisableTemporaryToggleValueName })
                {
                    object? raw = registry.ReadValue(SignatureRegistryHive.CurrentUser, key, value);
                    if (switchedOffAt == null && IsOne(raw))
                    {
                        switchedOffAt = "HKCU\\" + key + "\\" + value;
                    }
                }
            }

            return Decide(officeVersion, OfficeBuildInfo.Read(registry), switchedOffAt);
        }

        private static bool IsOne(object? value)
        {
            return value is int number ? number == 1
                : value is long wide ? wide == 1
                : value is string text && string.Equals(text.Trim(), "1", StringComparison.Ordinal);
        }

        /// <summary>The decision from the settings alone. Pure.</summary>
        public static SignatureRoaming Decide(string officeVersion, OfficeBuildInfo build, string? switchedOffAt)
        {
            if (build == null)
            {
                throw new ArgumentNullException(nameof(build));
            }

            if (switchedOffAt != null)
            {
                return new SignatureRoaming(SignatureRoamingState.Off,
                    "roaming signatures are switched off (" + switchedOffAt + " = 1); the signature files on this machine are authoritative",
                    switchedOffAt, build);
            }

            if (string.Equals(officeVersion, "15.0", StringComparison.Ordinal))
            {
                return new SignatureRoaming(SignatureRoamingState.Off,
                    "Outlook 2013 predates roaming signatures; the signature files on this machine are authoritative", null, build);
            }

            if (build.Build is int number && number < FirstRoamingBuild)
            {
                return new SignatureRoaming(SignatureRoamingState.Off,
                    "Office build " + build.Version + " predates Version 2303 (build " + FirstRoamingBuild.ToString(CultureInfo.InvariantCulture)
                    + "), the first that roams signatures; the signature files on this machine are authoritative", null, build);
            }

            if (build.Build != null && build.IsMicrosoft365)
            {
                return new SignatureRoaming(SignatureRoamingState.Possible,
                    "Microsoft 365 Apps " + build.Version + " roams the signatures of Exchange Online and Outlook.com accounts unless "
                    + DisableValueName + " is set", null, build);
            }

            return new SignatureRoaming(SignatureRoamingState.Off,
                (build.Build == null
                    ? "the Office build could not be read"
                    : "Office " + (build.ProductReleaseIds ?? "(product unknown)") + " " + build.Version + " is not Microsoft 365 Apps")
                + ", so the signature files on this machine are authoritative - unless Outlook keeps roaming copies here (signatures "
                + "named 'Name (address)' after an Exchange account), which would show that it roams after all",
                null, build, evidenceDecides: true);
        }

        /// <summary>
        /// This verdict with Outlook's own local copies taken into account: where the files decide, a
        /// signature named after an EXCHANGE account of this computer's profiles - <c>Name (address)</c>,
        /// the only accounts that roam - makes it EVIDENT for those accounts. A final verdict (switched
        /// off, too old, Microsoft 365) is returned unchanged. Pure.
        /// </summary>
        public SignatureRoaming WithEvidence(IEnumerable<string> signatureNames, IEnumerable<string> exchangeAccounts)
        {
            if (!EvidenceDecides)
            {
                return this;
            }

            List<string> exchange = (exchangeAccounts ?? Array.Empty<string>()).ToList();
            List<string> evident = (signatureNames ?? Array.Empty<string>())
                .Select(AccountSuffixOf)
                .Where(a => a != null && exchange.Any(e => ProfileAccountEntries.AddressEquals(e, a)))
                .Select(a => a!)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(a => a, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (evident.Count == 0)
            {
                return this;
            }

            return new SignatureRoaming(SignatureRoamingState.Evident,
                "Outlook keeps local copies of roaming signatures of " + string.Join(", ", evident) + " here (signatures named "
                + "'Name (address)'), so those accounts' signatures live in their mailboxes",
                SwitchedOffAt, Build, evidenceDecides: true, evidentAccounts: evident);
        }

        /// <summary>
        /// The account a signature name belongs to when it is shaped like Outlook's local copy of a
        /// roaming signature - <c>Name (address)</c> - or null for a plain name. Pure.
        /// </summary>
        public static string? AccountSuffixOf(string? signatureName)
        {
            if (string.IsNullOrWhiteSpace(signatureName))
            {
                return null;
            }

            string name = signatureName!.Trim();
            if (!name.EndsWith(")", StringComparison.Ordinal))
            {
                return null;
            }

            int open = name.LastIndexOf(" (", StringComparison.Ordinal);
            if (open <= 0)
            {
                return null;
            }

            string inner = name.Substring(open + 2, name.Length - open - 3).Trim();
            int at = inner.IndexOf('@');
            return at > 0 && at < inner.Length - 1 && inner.IndexOf(' ') < 0 && inner.IndexOf('@', at + 1) < 0 ? inner : null;
        }
    }

    /// <summary>
    /// Everything the signature model reads, read once and fresh: the folder, the profiles and their
    /// mail accounts, the policy, roaming, and the signatures in the folder. The read model behind
    /// list_signatures, manage_signature's refusals and every draft's signature decision. Never
    /// throws for a part that cannot be read: it comes back empty and is named in <see cref="Problems"/>.
    /// </summary>
    public sealed class SignatureEnvironment
    {
        private SignatureEnvironment(
            SignatureFolderInfo folder,
            ProfileList profiles,
            IReadOnlyList<ProfileAccountEntry> entries,
            SignaturePolicy policy,
            SignatureRoaming roaming,
            IReadOnlyList<SignatureInfo> signatures,
            IReadOnlyList<string> problems,
            string officeVersion)
        {
            Folder = folder;
            Profiles = profiles;
            Entries = entries;
            Policy = policy;
            Roaming = roaming;
            Signatures = signatures;
            Problems = problems;
            OfficeVersion = officeVersion;
        }

        /// <summary>The signature folder.</summary>
        public SignatureFolderInfo Folder { get; }

        /// <summary>Every profile, and the recorded default.</summary>
        public ProfileList Profiles { get; }

        /// <summary>Every account-manager entry of every profile, classified, each with its profile.</summary>
        public IReadOnlyList<ProfileAccountEntry> Entries { get; }

        /// <summary>The administrator's rules.</summary>
        public SignaturePolicy Policy { get; }

        /// <summary>Whether signatures can roam.</summary>
        public SignatureRoaming Roaming { get; }

        /// <summary>The signatures in <see cref="Folder"/>, name-sorted.</summary>
        public IReadOnlyList<SignatureInfo> Signatures { get; }

        /// <summary>Parts that could not be read, in words (empty when everything was).</summary>
        public IReadOnlyList<string> Problems { get; }

        /// <summary>The Office major the registry paths were built from.</summary>
        public string OfficeVersion { get; }

        /// <summary>Every profile name the model knows: the list, plus any an entry names that the list lacks.</summary>
        public IReadOnlyList<string> ProfileNames
        {
            get
            {
                List<string> names = new List<string>(Profiles.Names);
                foreach (string? name in Entries.Select(e => e.ProfileName))
                {
                    if (name != null && !names.Any(n => ProfileAccountEntries.ProfileEquals(n, name)))
                    {
                        names.Add(name);
                    }
                }

                return names.OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
            }
        }

        /// <summary>The mail accounts of one profile - never of any other.</summary>
        public IReadOnlyList<ProfileAccountEntry> MailAccountsOf(string? profile)
        {
            return Entries
                .Where(e => e.Kind == ProfileEntryKind.MailAccount && ProfileAccountEntries.ProfileEquals(e.ProfileName, profile))
                .ToList();
        }

        /// <summary>The profile's name as recorded, matched case-insensitively, or null when there is no such profile.</summary>
        public string? FindProfile(string? profile)
        {
            if (string.IsNullOrWhiteSpace(profile))
            {
                return null;
            }

            return ProfileNames.FirstOrDefault(n => ProfileAccountEntries.ProfileEquals(n, profile));
        }

        /// <summary>A signature of the folder by name, case-insensitively, or null.</summary>
        public SignatureInfo? FindSignature(string? name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            string wanted = name!.Trim();
            return Signatures.FirstOrDefault(s => string.Equals(s.Name, wanted, StringComparison.OrdinalIgnoreCase));
        }

        /// <summary>
        /// True when a signature is Outlook's local copy of a signature kept in a mailbox: its name
        /// carries an account suffix and roaming is not ruled out on this machine (Q152 (a)).
        /// </summary>
        public bool IsManagedInMailbox(SignatureInfo signature)
        {
            return signature != null && signature.ScopeAccount != null && Roaming.MayRoamFor(signature.ScopeAccount);
        }

        /// <summary>The live environment, read now.</summary>
        public static SignatureEnvironment ReadCurrent(string? directoryOverride = null)
        {
            return Read(
                new LiveSignatureSettingsRegistry(),
                new LiveProfileAccountRegistry(),
                OutlookProfileRegistry.OfficeVersion,
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                directoryOverride);
        }

        /// <summary>Reads everything through the seams. <paramref name="directoryOverride"/> replaces the resolved folder (T1).</summary>
        public static SignatureEnvironment Read(
            ISignatureSettingsRegistry settings,
            IProfileAccountRegistry accounts,
            string officeVersion,
            string appDataRoaming,
            string? directoryOverride = null)
        {
            if (settings == null)
            {
                throw new ArgumentNullException(nameof(settings));
            }

            if (accounts == null)
            {
                throw new ArgumentNullException(nameof(accounts));
            }

            List<string> problems = new List<string>();
            SignatureFolderInfo folder = directoryOverride != null
                ? new SignatureFolderInfo(directoryOverride, false, null)
                : Guard(() => SignatureFolder.Read(settings, officeVersion, appDataRoaming),
                    SignatureFolder.Read(new EmptySettings(), officeVersion, appDataRoaming), "the signature folder value", problems);
            ProfileList profiles = Guard(accounts.ReadProfiles, new ProfileList(null, Array.Empty<string>()), "the Outlook profile list", problems);
            IReadOnlyList<ProfileAccountEntry> entries = Guard(
                () => ProfileAccountEntries.ReadAll(accounts), Array.Empty<ProfileAccountEntry>(), "the profiles' account lists", problems);
            SignaturePolicy policy = Guard(() => SignaturePolicy.Read(settings, officeVersion), SignaturePolicy.None, "the signature policy", problems);
            SignatureRoaming roaming = Guard(
                () => SignatureRoaming.Read(settings, officeVersion),
                SignatureRoaming.Decide(officeVersion, OfficeBuildInfo.Unknown, null), "the roaming-signature settings", problems);
            IReadOnlyList<SignatureInfo> signatures = Guard(
                () => SignatureCatalog.ListSignatures(folder.Path), Array.Empty<SignatureInfo>(), "the signature folder", problems);
            roaming = roaming.WithEvidence(signatures.Select(s => s.Name), ExchangeAddresses(entries));
            return new SignatureEnvironment(folder, profiles, entries, policy, roaming, signatures, problems, officeVersion);
        }

        /// <summary>The addresses of every profile's Exchange accounts - the only ones whose signatures roam.</summary>
        public static IEnumerable<string> ExchangeAddresses(IEnumerable<ProfileAccountEntry> entries)
        {
            return entries
                .Where(e => e.Kind == ProfileEntryKind.MailAccount && e.AccountKind == MailAccountKind.Exchange && e.Address != null)
                .Select(e => e.Address!);
        }

        private static T Guard<T>(Func<T> read, T fallback, string what, List<string> problems)
        {
            try
            {
                return read();
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                problems.Add(what + " could not be read (" + ex.GetType().Name + ")");
                return fallback;
            }
        }

        private sealed class EmptySettings : ISignatureSettingsRegistry
        {
            public object? ReadValue(SignatureRegistryHive hive, string keyPath, string valueName) => null;
        }
    }
}
