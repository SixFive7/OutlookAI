using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Win32;

namespace OutlookAI.Core.Services
{
    /// <summary>
    /// What one entry of a profile's account-manager key is. <see cref="ProfileAccountEntries"/>
    /// says how each kind is recognised and what that rests on.
    /// </summary>
    public enum ProfileEntryKind
    {
        /// <summary>
        /// A mail account Outlook sends from - POP3, IMAP or Exchange. The ONLY kind whose
        /// <c>New Signature</c> / <c>Reply-Forward Signature</c> values are a signature default.
        /// </summary>
        MailAccount,

        /// <summary>A data file: the MAPI-service entry of a .pst (<c>MSUPST MS</c> / <c>MSPST MS</c>).</summary>
        DataFile,

        /// <summary>An address book: <c>CONTAB</c>, <c>EMABLT</c>, <c>MSPST AB</c>, or an LDAP directory.</summary>
        AddressBook,

        /// <summary>
        /// Anything else: a MAPI service this code does not know, an account class it does not
        /// know (an Exchange ActiveSync account, for one), or an entry with no readable
        /// <c>clsid</c>. Treated exactly like a data file - never written, never reported as an
        /// account - because guessing is how a default ends up where Outlook never reads it.
        /// </summary>
        Unrecognized,
    }

    /// <summary>One entry of the account-manager key, classified (data only).</summary>
    public sealed class ProfileAccountEntry
    {
        /// <summary>Creates a classified entry.</summary>
        public ProfileAccountEntry(
            string key,
            ProfileEntryKind kind,
            string? address,
            string? accountName,
            string? newSignature,
            string? replyForwardSignature)
        {
            Key = key ?? throw new ArgumentNullException(nameof(key));
            Kind = kind;
            Address = address;
            AccountName = accountName;
            NewSignature = newSignature;
            ReplyForwardSignature = replyForwardSignature;
        }

        /// <summary>
        /// Opaque handle of the entry: its HKCU-relative registry path for the live registry,
        /// any unique string for a fake.
        /// </summary>
        public string Key { get; }

        /// <summary>What the entry is.</summary>
        public ProfileEntryKind Kind { get; }

        /// <summary>
        /// The account's OWN email address - the one <c>list_accounts</c> reports and
        /// <c>set_default_for.account</c> is matched against. Null unless <see cref="Kind"/> is
        /// <see cref="ProfileEntryKind.MailAccount"/> and the value that carries the address is
        /// address-shaped.
        /// </summary>
        public string? Address { get; }

        /// <summary>
        /// The entry's <c>Account Name</c> as recorded: a DISPLAY name, which for a data file is
        /// the store's name and for a POP3/IMAP account is whatever the user called it. Carried
        /// for messages only - it is never what an account is matched on, except for Exchange,
        /// where it is the address (see <see cref="ProfileAccountEntries"/>).
        /// </summary>
        public string? AccountName { get; }

        /// <summary>Recorded new-message signature name (null = absent or blank).</summary>
        public string? NewSignature { get; }

        /// <summary>Recorded reply/forward signature name (null = absent or blank).</summary>
        public string? ReplyForwardSignature { get; }
    }

    /// <summary>One raw entry as read: its key and the values the classification needs.</summary>
    public sealed class ProfileRegistryEntry
    {
        /// <summary>Creates a raw entry.</summary>
        public ProfileRegistryEntry(string key, IReadOnlyDictionary<string, object?> values)
        {
            Key = key ?? throw new ArgumentNullException(nameof(key));
            Values = values ?? throw new ArgumentNullException(nameof(values));
        }

        /// <summary>The entry's handle (HKCU-relative registry path for the live registry).</summary>
        public string Key { get; }

        /// <summary>Value name to raw value (REG_SZ string or REG_BINARY bytes, as stored).</summary>
        public IReadOnlyDictionary<string, object?> Values { get; }
    }

    /// <summary>
    /// Raw access to the entries of the default profile's account-manager key - the T1 seam.
    /// A fake holds the value names and shapes Outlook writes, so the classification, the
    /// selection, the write guard and the read-back all run against it unchanged;
    /// <see cref="LiveProfileAccountRegistry"/> is the real registry.
    /// </summary>
    public interface IProfileAccountRegistry
    {
        /// <summary>
        /// Every entry, read NOW - never cached, so a read-back sees what the registry holds after
        /// the write rather than what the writer believed it wrote.
        /// </summary>
        IReadOnlyList<ProfileRegistryEntry> ReadEntries();

        /// <summary>Sets a REG_SZ value on an EXISTING entry; never creates an entry.</summary>
        void SetString(string key, string valueName, string value);

        /// <summary>Removes one value; a no-op when the value or the entry is absent.</summary>
        void DeleteValue(string key, string valueName);
    }

    /// <summary>
    /// The live registry: every numbered subkey of
    /// HKCU\Software\Microsoft\Office\&lt;major&gt;\Outlook\Profiles\&lt;default profile&gt;\9375CFF0413111d3B88A00104B2A6676
    /// (every profile's, when no <c>DefaultProfile</c> is recorded - as this product has always
    /// read it). The major is the detected one (<see cref="OutlookProfileRegistry.OfficeVersion"/>).
    /// Reads only <see cref="ProfileAccountEntries.ValueNamesRead"/>: an account entry also holds
    /// its sealed POP3 password, which nothing here has any business loading.
    /// </summary>
    public sealed class LiveProfileAccountRegistry : IProfileAccountRegistry
    {
        /// <inheritdoc />
        public IReadOnlyList<ProfileRegistryEntry> ReadEntries()
        {
            List<ProfileRegistryEntry> entries = new List<ProfileRegistryEntry>();
            string outlookRoot = OutlookProfileRegistry.OutlookRootKeyPath;
            using RegistryKey? outlook = Registry.CurrentUser.OpenSubKey(outlookRoot);
            if (outlook == null)
            {
                return entries;
            }

            string? defaultProfile = outlook.GetValue("DefaultProfile") as string;
            using RegistryKey? profiles = outlook.OpenSubKey(OutlookProfileRegistry.ProfilesSubKeyName);
            if (profiles == null)
            {
                return entries;
            }

            IEnumerable<string> profileNames = defaultProfile != null
                ? new[] { defaultProfile }
                : profiles.GetSubKeyNames();
            foreach (string profileName in profileNames)
            {
                string accountsPath = profileName + "\\" + OutlookProfileRegistry.AccountsSubKeyName;
                using RegistryKey? accounts = profiles.OpenSubKey(accountsPath);
                if (accounts == null)
                {
                    continue;
                }

                foreach (string subKeyName in accounts.GetSubKeyNames())
                {
                    using RegistryKey? entry = accounts.OpenSubKey(subKeyName);
                    if (entry == null)
                    {
                        continue;
                    }

                    Dictionary<string, object?> values = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
                    foreach (string valueName in ProfileAccountEntries.ValueNamesRead)
                    {
                        object? value = entry.GetValue(valueName);
                        if (value != null)
                        {
                            values[valueName] = value;
                        }
                    }

                    entries.Add(new ProfileRegistryEntry(
                        outlookRoot + "\\" + OutlookProfileRegistry.ProfilesSubKeyName + "\\" + accountsPath + "\\" + subKeyName,
                        values));
                }
            }

            return entries;
        }

        /// <inheritdoc />
        public void SetString(string key, string valueName, string value)
        {
            using RegistryKey? entry = Registry.CurrentUser.OpenSubKey(key, writable: true);
            if (entry == null)
            {
                throw new InvalidOperationException("The account's profile registry key no longer exists: " + key);
            }

            entry.SetValue(valueName, value, RegistryValueKind.String);
        }

        /// <inheritdoc />
        public void DeleteValue(string key, string valueName)
        {
            using RegistryKey? entry = Registry.CurrentUser.OpenSubKey(key, writable: true);
            entry?.DeleteValue(valueName, throwOnMissingValue: false);
        }
    }

    /// <summary>
    /// Which entries of a profile's account-manager key (<c>9375CFF0413111d3B88A00104B2A6676</c>)
    /// are MAIL ACCOUNTS, and which email address each belongs to. The one rule behind every
    /// signature-default read and write: <c>manage_signature</c>'s <c>set_default_for</c>, its
    /// delete sweep, the read-back that verifies a write, and <c>list_signatures</c>.
    /// <para>
    /// <b>Why it exists - the defect it replaces, measured 2026-09-24 on OAI-UNINDEXED.</b> The
    /// old rule was "an entry whose <c>Account Name</c> contains <c>@</c> is a mail account, and
    /// that name is its address". The key does not list only accounts: it lists every MAPI
    /// service of the profile, and a POP3 or IMAP data file is named after its address by
    /// Outlook's own default. On the tier profile, <c>set_default_for identity@vm.invalid</c>
    /// wrote <c>New Signature</c> onto entry 00000005 - the identity PST's data-file entry,
    /// <c>Account Name</c> <c>identity@vm.invalid</c> - and never saw the POP3 account 00000004
    /// at all, because its <c>Account Name</c> is the display name <c>OutlookAI identity sink</c>.
    /// <c>list_signatures</c> read back by the same rule and reported the default as set.
    /// </para>
    /// <para>
    /// <b>What tells the entries apart.</b> The entry's <c>clsid</c> - the account CLASS, one of
    /// Microsoft's Account Management API class identifiers - and, for the class that wraps a
    /// MAPI service, its <c>Service Name</c>. Never the <c>Account Name</c>:
    /// </para>
    /// <list type="bullet">
    /// <item><b>POP3 account</b> - <c>clsid</c> <c>{ED475411-B0D6-11D2-8C3B-00104B2A6676}</c>
    /// (CLSID_OlkPOP3Account). Address: <c>Email</c> (PROP_ACCT_USER_EMAIL_ADDR). MEASURED on
    /// the guests (16.0.17932): <c>Email</c>, <c>POP3 Server</c> and <c>POP3 User</c> stored as
    /// REG_SZ beside an <c>SMTP Server</c> value, and an <c>Account Name</c> (PROP_ACCT_NAME)
    /// that is the user's display name for the account - <c>OutlookAI identity sink</c> on
    /// OAI-UNINDEXED.</item>
    /// <item><b>IMAP account</b> - <c>{ED475412-B0D6-11D2-8C3B-00104B2A6676}</c>
    /// (CLSID_OlkIMAP4Account). Address: <c>Email</c>. DOCUMENTED (the class identifier is
    /// Microsoft's), NOT measured in this repository.</item>
    /// <item><b>Exchange account</b> - <c>{ED475414-B0D6-11D2-8C3B-00104B2A6676}</c>
    /// (CLSID_OlkMAPIAccount, the MAPI-service wrapper) whose <c>Service Name</c> is
    /// <c>MSEMS</c>. Address: <c>Email</c> when present, otherwise <c>Account Name</c>. MEASURED
    /// read-only on the development host's own profile, 2026-09-24 (value names and kinds only):
    /// three such entries, NONE carrying an <c>Email</c> value, each with an address-shaped
    /// <c>Account Name</c> - the value this product has always matched Exchange entries on.
    /// Whether that name always equals the <c>SmtpAddress</c> list_accounts reports (a mailbox
    /// signed into under an alias, say) is not measured; when it differs, the request is refused,
    /// never guessed. An Exchange account whose <c>Account Name</c> is not an address is
    /// unmatchable, so refused too.</item>
    /// <item><b>Data file</b> - the same wrapper with <c>MSUPST MS</c> (Unicode .pst) or
    /// <c>MSPST MS</c> (ANSI). <c>Account Name</c> is the store's display name, which CAN be an
    /// address. MEASURED on the guests: <c>identity@vm.invalid</c> on OAI-UNINDEXED (the defect's
    /// entry), and <c>tier@vm.invalid</c> on the profile the tier .prf builds (2026-09-15).</item>
    /// <item><b>Address book</b> - the wrapper with <c>CONTAB</c> (MEASURED, guest and host),
    /// <c>EMABLT</c> or <c>MSPST AB</c>; or an LDAP directory,
    /// <c>{4DB5CBF2-3B77-4852-BC8E-BB81908861F3}</c> (CLSID_OlkLDAPAccount, DOCUMENTED).</item>
    /// <item><b>Anything else</b> - <see cref="ProfileEntryKind.Unrecognized"/>. Deliberately
    /// fail-closed, the opposite of the testbed's counting rule (which counts an unknown wrapper
    /// as an account because over-counting is its safe side): for a WRITE, the safe side is not
    /// to write.</item>
    /// </list>
    /// <para>
    /// <b>Matching.</b> The whole address, trimmed, case-insensitive - never a substring, never a
    /// display name. Zero or several mail accounts with the requested address is a refusal the
    /// caller makes; nothing here picks one.
    /// </para>
    /// </summary>
    public static class ProfileAccountEntries
    {
        /// <summary>The account class (REG_SZ, a braced GUID).</summary>
        public const string ClsidValueName = "clsid";

        /// <summary>The wrapped MAPI service's name, on <see cref="MapiServiceClsid"/> entries.</summary>
        public const string ServiceNameValueName = "Service Name";

        /// <summary>PROP_ACCT_NAME: the entry's display name - an address only by coincidence or by default.</summary>
        public const string AccountNameValueName = "Account Name";

        /// <summary>PROP_ACCT_USER_EMAIL_ADDR: the account's email address (POP3/IMAP; absent on the measured Exchange entries).</summary>
        public const string EmailValueName = "Email";

        /// <summary>The new-message default signature (REG_SZ as this product writes it; REG_BINARY UTF-16LE from some writers).</summary>
        public const string NewSignatureValueName = "New Signature";

        /// <summary>The reply/forward default signature.</summary>
        public const string ReplyForwardSignatureValueName = "Reply-Forward Signature";

        /// <summary><c>Service Name</c> of an Exchange mailbox's wrapper entry.</summary>
        public const string ExchangeServiceName = "MSEMS";

        /// <summary>CLSID_OlkPOP3Account.</summary>
        public static readonly Guid Pop3AccountClsid = new Guid("ED475411-B0D6-11D2-8C3B-00104B2A6676");

        /// <summary>CLSID_OlkIMAP4Account.</summary>
        public static readonly Guid Imap4AccountClsid = new Guid("ED475412-B0D6-11D2-8C3B-00104B2A6676");

        /// <summary>CLSID_OlkMAPIAccount - the wrapper around a MAPI service: Exchange, data files, address books alike.</summary>
        public static readonly Guid MapiServiceClsid = new Guid("ED475414-B0D6-11D2-8C3B-00104B2A6676");

        /// <summary>CLSID_OlkLDAPAccount - an LDAP address book.</summary>
        public static readonly Guid LdapDirectoryClsid = new Guid("4DB5CBF2-3B77-4852-BC8E-BB81908861F3");

        /// <summary>
        /// The only values the live reader loads - the classification's inputs and the two
        /// defaults. Everything else in an entry is left unread.
        /// </summary>
        public static readonly IReadOnlyList<string> ValueNamesRead = Array.AsReadOnly(new[]
        {
            ClsidValueName,
            ServiceNameValueName,
            AccountNameValueName,
            EmailValueName,
            NewSignatureValueName,
            ReplyForwardSignatureValueName,
        });

        private static readonly string[] DataFileServiceNames = { "MSUPST MS", "MSPST MS" };

        private static readonly string[] AddressBookServiceNames = { "CONTAB", "EMABLT", "MSPST AB" };

        /// <summary>
        /// Classifies an entry by its class and wrapped service alone - see the type summary.
        /// Pure, so the whole table is assertable without a registry.
        /// </summary>
        public static ProfileEntryKind ClassifyKind(string? clsid, string? serviceName)
        {
            if (!Guid.TryParse(clsid?.Trim(), out Guid type))
            {
                return ProfileEntryKind.Unrecognized;
            }

            if (type == Pop3AccountClsid || type == Imap4AccountClsid)
            {
                return ProfileEntryKind.MailAccount;
            }

            if (type == LdapDirectoryClsid)
            {
                return ProfileEntryKind.AddressBook;
            }

            if (type != MapiServiceClsid)
            {
                return ProfileEntryKind.Unrecognized;
            }

            string service = serviceName?.Trim() ?? string.Empty;
            if (string.Equals(service, ExchangeServiceName, StringComparison.OrdinalIgnoreCase))
            {
                return ProfileEntryKind.MailAccount;
            }

            if (DataFileServiceNames.Contains(service, StringComparer.OrdinalIgnoreCase))
            {
                return ProfileEntryKind.DataFile;
            }

            return AddressBookServiceNames.Contains(service, StringComparer.OrdinalIgnoreCase)
                ? ProfileEntryKind.AddressBook
                : ProfileEntryKind.Unrecognized;
        }

        /// <summary>
        /// Classifies one raw entry and derives its address and recorded defaults. Value names
        /// are matched case-insensitively (the registry's own rule), whatever comparer the
        /// dictionary was built with. Pure.
        /// </summary>
        public static ProfileAccountEntry Classify(string key, IReadOnlyDictionary<string, object?> values)
        {
            if (values == null)
            {
                throw new ArgumentNullException(nameof(values));
            }

            string? clsid = Text(values, ClsidValueName);
            string? serviceName = Text(values, ServiceNameValueName);
            string? accountName = Blank(Text(values, AccountNameValueName));
            ProfileEntryKind kind = ClassifyKind(clsid, serviceName);

            string? address = null;
            if (kind == ProfileEntryKind.MailAccount)
            {
                string? email = Blank(Text(values, EmailValueName));
                bool exchange = Guid.TryParse(clsid?.Trim(), out Guid type) && type == MapiServiceClsid;

                // POP3 / IMAP: Email and nothing else - their Account Name is the display name,
                // and matching on it is the defect this type replaces. Exchange records no Email
                // (measured), and its Account Name is the mailbox address.
                string? candidate = IsAddressShaped(email) ? email : exchange ? accountName : null;
                address = IsAddressShaped(candidate) ? candidate : null;
            }

            return new ProfileAccountEntry(
                key ?? string.Empty,
                kind,
                address,
                accountName,
                Blank(Text(values, NewSignatureValueName)),
                Blank(Text(values, ReplyForwardSignatureValueName)));
        }

        /// <summary>Every entry of <paramref name="registry"/>, freshly read and classified.</summary>
        public static IReadOnlyList<ProfileAccountEntry> ReadAll(IProfileAccountRegistry registry)
        {
            if (registry == null)
            {
                throw new ArgumentNullException(nameof(registry));
            }

            return registry.ReadEntries().Select(e => Classify(e.Key, e.Values)).ToList();
        }

        /// <summary>The whole address, trimmed, case-insensitive. Null or blank never matches.</summary>
        public static bool AddressEquals(string? left, string? right)
        {
            if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
            {
                return false;
            }

            return string.Equals(left!.Trim(), right!.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>Entry handles compare as registry paths do: case-insensitively.</summary>
        public static bool KeyEquals(string? left, string? right)
        {
            return left != null && right != null && string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>The entry kind in words, for refusals and failure messages.</summary>
        public static string Describe(ProfileEntryKind kind)
        {
            switch (kind)
            {
                case ProfileEntryKind.MailAccount:
                    return "mail account";
                case ProfileEntryKind.DataFile:
                    return "data file";
                case ProfileEntryKind.AddressBook:
                    return "address book";
                default:
                    return "profile entry this product does not recognise as a mail account";
            }
        }

        /// <summary>The last segment of an entry handle (the numbered subkey), for short messages.</summary>
        public static string ShortKey(string key)
        {
            if (string.IsNullOrEmpty(key))
            {
                return key ?? string.Empty;
            }

            int slash = key.LastIndexOf('\\');
            return slash >= 0 && slash < key.Length - 1 ? key.Substring(slash + 1) : key;
        }

        private static bool IsAddressShaped(string? value)
        {
            if (value == null)
            {
                return false;
            }

            int at = value.IndexOf('@');
            return at > 0 && at < value.Length - 1;
        }

        private static string? Blank(string? value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value!.Trim();
        }

        private static string? Text(IReadOnlyDictionary<string, object?> values, string name)
        {
            if (values.TryGetValue(name, out object? value))
            {
                return SignatureCatalog.DecodeRegistryString(value);
            }

            foreach (KeyValuePair<string, object?> pair in values)
            {
                if (string.Equals(pair.Key, name, StringComparison.OrdinalIgnoreCase))
                {
                    return SignatureCatalog.DecodeRegistryString(pair.Value);
                }
            }

            return null;
        }
    }
}
