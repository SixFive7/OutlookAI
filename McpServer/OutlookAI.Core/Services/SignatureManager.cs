using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;

namespace OutlookAI.Core.Services
{
    /// <summary>Validated manage_signature request (soak fix D38).</summary>
    public sealed class ManageSignatureRequest
    {
        /// <summary>"create" | "update" | "delete" (case-insensitive).</summary>
        public string Action { get; set; } = string.Empty;

        /// <summary>Signature name (file base name, what Outlook's pickers show).</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>Plain-text body (create/update). Derived from <see cref="BodyHtml"/> when omitted.</summary>
        public string? BodyText { get; set; }

        /// <summary>HTML body (create/update), fragment or full document. Derived from <see cref="BodyText"/> when omitted.</summary>
        public string? BodyHtml { get; set; }

        /// <summary>SMTP address of the account to record the signature as default for (optional; create/update only).</summary>
        public string? DefaultForAccount { get; set; }

        /// <summary>"new" | "reply" | "both" - which default(s) to record for <see cref="DefaultForAccount"/>.</summary>
        public string? DefaultForScope { get; set; }
    }

    /// <summary>One profile MAIL ACCOUNT's default-signature registry state (read/write handle).</summary>
    public sealed class SignatureDefaultsRow
    {
        /// <summary>Creates a row.</summary>
        public SignatureDefaultsRow(string accountKey, string account, string? newMessage, string? replyForward)
        {
            AccountKey = accountKey;
            Account = account;
            NewMessage = newMessage;
            ReplyForward = replyForward;
        }

        /// <summary>Opaque store handle of the account (registry subkey path for the production store).</summary>
        public string AccountKey { get; }

        /// <summary>
        /// The account's email address - what <c>list_accounts</c> reports and
        /// <c>set_default_for.account</c> is matched against. <see cref="ProfileAccountEntries"/>
        /// derives it: <c>Email</c> for a POP3/IMAP account, <c>Account Name</c> for an Exchange
        /// account (which records no <c>Email</c>). It used to be <c>Account Name</c> for every
        /// entry, which is a display name - and, for a data file, the store's name.
        /// </summary>
        public string Account { get; }

        /// <summary>Currently assigned new-message signature name (null = absent).</summary>
        public string? NewMessage { get; }

        /// <summary>Currently assigned reply/forward signature name (null = absent).</summary>
        public string? ReplyForward { get; }
    }

    /// <summary>
    /// Read/write access to the per-account default-signature registry values
    /// ("New Signature" / "Reply-Forward Signature" under the profile's 9375CFF0...
    /// key - the locations D37 verified readable on this machine). The interface is
    /// the T1 seam; <see cref="ProfileSignatureDefaultsStore"/> is the live registry (or, for T1,
    /// any <see cref="IProfileAccountRegistry"/>).
    /// </summary>
    public interface ISignatureDefaultsStore
    {
        /// <summary>
        /// Enumerates the profile's MAIL ACCOUNTS - never a data file or an address book, whatever
        /// it is named (<see cref="ProfileAccountEntries"/>) - with their current assignments.
        /// </summary>
        IReadOnlyList<SignatureDefaultsRow> ReadAccounts();

        /// <summary>
        /// Re-reads ONE entry from the store, now, and classifies it from its raw values - null
        /// when it is no longer an entry of the profile's account list. The independent half of
        /// the read-back: it says what the entry a write went to actually IS, rather than what
        /// the row the write was aimed with claimed.
        /// </summary>
        ProfileAccountEntry? ReadEntry(string accountKey);

        /// <summary>Writes one default value (REG_SZ; absent value is created).</summary>
        void WriteDefault(string accountKey, string valueName, string signatureName);

        /// <summary>Removes one default value (no-op when absent).</summary>
        void ClearDefault(string accountKey, string valueName);
    }

    /// <summary>
    /// Signature defaults in the profile registry: the entries of
    /// HKCU\Software\Microsoft\Office\&lt;major&gt;\Outlook\Profiles\&lt;default profile&gt;\9375CFF0413111d3B88A00104B2A6676
    /// (<see cref="LiveProfileAccountRegistry"/>; the major is whichever Office version this
    /// machine actually has), or any <see cref="IProfileAccountRegistry"/> for T1. Which entries
    /// are accounts, and what their addresses are, is <see cref="ProfileAccountEntries"/>' rule.
    /// Writes are surgical: only the two known value names, only on an entry that - re-read at
    /// the moment of the write - classifies as a MAIL ACCOUNT, never creating an entry. The old
    /// guard was "an SMTP-shaped Account Name", which a data file named after its address
    /// passes (the 2026-09-24 defect).
    /// </summary>
    public sealed class ProfileSignatureDefaultsStore : ISignatureDefaultsStore
    {
        private readonly IProfileAccountRegistry _registry;

        /// <summary>The live registry.</summary>
        public ProfileSignatureDefaultsStore()
            : this(new LiveProfileAccountRegistry())
        {
        }

        /// <summary>Any registry - the T1 seam.</summary>
        public ProfileSignatureDefaultsStore(IProfileAccountRegistry registry)
        {
            _registry = registry ?? throw new ArgumentNullException(nameof(registry));
        }

        /// <inheritdoc />
        public IReadOnlyList<SignatureDefaultsRow> ReadAccounts()
        {
            return ProfileAccountEntries.ReadAll(_registry)
                .Where(e => e.Kind == ProfileEntryKind.MailAccount && e.Address != null)
                .Select(e => new SignatureDefaultsRow(e.Key, e.Address!, e.NewSignature, e.ReplyForwardSignature))
                .ToList();
        }

        /// <inheritdoc />
        public ProfileAccountEntry? ReadEntry(string accountKey)
        {
            return ProfileAccountEntries.ReadAll(_registry)
                .FirstOrDefault(e => ProfileAccountEntries.KeyEquals(e.Key, accountKey));
        }

        /// <inheritdoc />
        public void WriteDefault(string accountKey, string valueName, string signatureName)
        {
            RequireDefaultValueName(valueName);
            ProfileAccountEntry entry = ReadEntry(accountKey)
                ?? throw new InvalidOperationException(
                    "The account's profile registry entry is no longer there (or no longer in the default profile): " + accountKey);
            RequireMailAccount(entry, valueName);
            _registry.SetString(accountKey, valueName, signatureName);
        }

        /// <inheritdoc />
        public void ClearDefault(string accountKey, string valueName)
        {
            RequireDefaultValueName(valueName);
            ProfileAccountEntry? entry = ReadEntry(accountKey);
            if (entry == null)
            {
                // Nothing of the profile's account list there - nothing to clear.
                return;
            }

            RequireMailAccount(entry, valueName);
            _registry.DeleteValue(accountKey, valueName);
        }

        private static void RequireDefaultValueName(string valueName)
        {
            if (!string.Equals(valueName, SignatureManager.NewSignatureValueName, StringComparison.Ordinal)
                && !string.Equals(valueName, SignatureManager.ReplyForwardSignatureValueName, StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    "Only '" + SignatureManager.NewSignatureValueName + "' and '" + SignatureManager.ReplyForwardSignatureValueName
                    + "' are ever written to the profile registry, not '" + valueName + "'.",
                    nameof(valueName));
            }
        }

        private static void RequireMailAccount(ProfileAccountEntry entry, string valueName)
        {
            if (entry.Kind != ProfileEntryKind.MailAccount)
            {
                throw new InvalidOperationException(
                    "Refusing to change '" + valueName + "' on profile entry " + ProfileAccountEntries.ShortKey(entry.Key)
                    + ": it is a " + ProfileAccountEntries.Describe(entry.Kind) + ", not a mail account, and signature "
                    + "defaults are only ever written to a mail account's own entry.");
            }
        }
    }

    /// <summary>
    /// Creates, updates and deletes Outlook signatures by writing the signature file
    /// set under %APPDATA%\Microsoft\Signatures (.htm + .txt + .rtf - each mail format
    /// reads only its own rendition and silently omits the signature when it is
    /// missing, so all three are always written; the missing renditions are derived).
    /// ALWAYS-ON safety: before ANY update or delete the signature's full current file
    /// set is copied to %LOCALAPPDATA%\OutlookAI\signature-backups\&lt;utc&gt;-&lt;name&gt;\ and
    /// the backup path is returned - a failing backup ABORTS the operation. Optional
    /// default assignment writes the per-account "New Signature"/"Reply-Forward
    /// Signature" REG_SZ values (D37 locations) onto the ONE mail account whose address is
    /// the one asked for (<see cref="ProfileAccountEntries"/>' rule; none or several is
    /// refused before anything is written) and reads them back before reporting success;
    /// deleting a signature clears dangling assignments that referenced it. Pure filesystem +
    /// registry - no COM, never starts Outlook. NOTE (docs): on Microsoft 365 Apps 2303+
    /// roaming signatures can overrule local files unless DisableRoamingSignatures=1; on
    /// Office LTSC (this machine) local files are authoritative.
    /// </summary>
    public static class SignatureManager
    {
        /// <summary>Registry value name for the new-message default.</summary>
        public const string NewSignatureValueName = ProfileAccountEntries.NewSignatureValueName;

        /// <summary>Registry value name for the reply/forward default.</summary>
        public const string ReplyForwardSignatureValueName = ProfileAccountEntries.ReplyForwardSignatureValueName;

        /// <summary>Maximum signature name length (file-name discipline).</summary>
        public const int NameMaxChars = 128;

        /// <summary>Timestamp format of backup directory names (UTC, filesystem-safe).</summary>
        public const string BackupTimestampFormat = "yyyyMMdd'T'HHmmssfff'Z'";

        /// <summary>Advice returned whenever registry defaults were written or cleared.</summary>
        public const string DefaultsRestartAdvice =
            "Outlook picks up default-signature changes at its next start (or when its Signatures dialog is reopened) - "
            + "already-open compose windows keep the old default.";

        private static readonly string[] ReservedDeviceNames =
        {
            "CON", "PRN", "AUX", "NUL",
            "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
            "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
        };

        /// <summary>Default backup root: %LOCALAPPDATA%\OutlookAI\signature-backups.</summary>
        public static string DefaultBackupRoot =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookAI", "signature-backups");

        /// <summary>
        /// Executes one manage_signature operation. <paramref name="directory"/>,
        /// <paramref name="backupRoot"/>, <paramref name="defaultsStore"/> and
        /// <paramref name="utcNow"/> are T1 seams; production passes null for all.
        /// Validation failures throw <see cref="ArgumentException"/> BEFORE any file or
        /// registry work; a failing backup throws <see cref="InvalidOperationException"/>
        /// before anything is modified.
        /// </summary>
        public static ManageSignatureOutcome Manage(
            ManageSignatureRequest request,
            string? directory = null,
            string? backupRoot = null,
            ISignatureDefaultsStore? defaultsStore = null,
            Func<DateTime>? utcNow = null)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            string action = ValidateAction(request.Action);
            string name = ValidateName(request.Name);
            ValidateBodies(action, request);
            (string? defaultAccount, string? defaultScope) = ValidateDefaultRequest(action, request);

            string root = directory ?? SignatureCatalog.DefaultSignatureDirectory;
            string backups = backupRoot ?? DefaultBackupRoot;
            IReadOnlyList<string> existing = ExistingFileSet(root, name);

            if (action == "create" && existing.Count > 0)
            {
                throw new ArgumentException(
                    "Signature '" + name + "' already exists - use action 'update' to change it (an automatic backup is made).");
            }

            if (action != "create" && existing.Count == 0)
            {
                throw new ArgumentException(
                    "Signature '" + name + "' was not found. Use list_signatures for the installed signature names.");
            }

            // Resolve the defaults store lazily: only set_default_for and delete (the
            // dangling-assignment sweep) need the registry at all.
            ISignatureDefaultsStore? store = null;
            SignatureDefaultsRow? targetAccount = null;
            if (defaultAccount != null)
            {
                store = defaultsStore ?? new ProfileSignatureDefaultsStore();
                targetAccount = SelectDefaultTarget(store.ReadAccounts(), defaultAccount);
            }

            string? backupPath = null;
            if (action != "create")
            {
                backupPath = BackupFileSet(root, name, existing, backups, utcNow ?? (() => DateTime.UtcNow));
            }

            List<string> advice = new List<string>();
            ManageSignatureOutcome outcome = new ManageSignatureOutcome
            {
                Action = action,
                Name = name,
                BackupPath = backupPath,
            };

            if (action == "delete")
            {
                outcome.FilesDeleted = DeleteFileSet(root, name, existing);
                outcome.DefaultsClearedForAccounts = ClearDanglingDefaults(
                    defaultsStore ?? new ProfileSignatureDefaultsStore(), name, advice);
                if (outcome.DefaultsClearedForAccounts is { Count: > 0 })
                {
                    advice.Add(DefaultsRestartAdvice);
                }
            }
            else
            {
                if (action == "update")
                {
                    DeleteFileSet(root, name, existing);
                }

                outcome.FilesWritten = WriteFileSet(root, name, request.BodyText, request.BodyHtml);
                if (targetAccount != null && store != null)
                {
                    ApplyDefaults(store, targetAccount, defaultScope!, name);
                    VerifyDefaults(store, targetAccount, defaultAccount!, defaultScope!, name);
                    outcome.DefaultSetForAccount = targetAccount.Account;
                    outcome.DefaultSetScope = defaultScope;
                    advice.Add(DefaultsRestartAdvice);
                }
            }

            outcome.Advice = advice.Count > 0 ? string.Join(" ", advice) : null;
            return outcome;
        }

        // ------------------------------------------------------------------ validation

        private static string ValidateAction(string? action)
        {
            string normalized = (action ?? string.Empty).Trim().ToLowerInvariant();
            if (normalized != "create" && normalized != "update" && normalized != "delete")
            {
                throw new ArgumentException("action must be 'create', 'update' or 'delete'.");
            }

            return normalized;
        }

        private static string ValidateName(string? name)
        {
            string trimmed = (name ?? string.Empty).Trim();
            if (trimmed.Length == 0)
            {
                throw new ArgumentException("name is required: the signature name as shown by list_signatures.");
            }

            if (trimmed.Length > NameMaxChars)
            {
                throw new ArgumentException("name must be at most " + NameMaxChars + " characters.");
            }

            if (trimmed.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0
                || trimmed.IndexOf(Path.DirectorySeparatorChar) >= 0
                || trimmed.IndexOf(Path.AltDirectorySeparatorChar) >= 0)
            {
                throw new ArgumentException("name contains characters that are not allowed in a signature (file) name.");
            }

            if (trimmed.EndsWith(".", StringComparison.Ordinal))
            {
                throw new ArgumentException("name must not end with a dot.");
            }

            if (ReservedDeviceNames.Contains(trimmed, StringComparer.OrdinalIgnoreCase))
            {
                throw new ArgumentException("name '" + trimmed + "' is a reserved Windows device name.");
            }

            return trimmed;
        }

        private static void ValidateBodies(string action, ManageSignatureRequest request)
        {
            bool hasText = !string.IsNullOrWhiteSpace(request.BodyText);
            bool hasHtml = !string.IsNullOrWhiteSpace(request.BodyHtml);

            if (action == "delete")
            {
                if (hasText || hasHtml)
                {
                    throw new ArgumentException("delete takes no body content - body_text/body_html apply to create/update only.");
                }

                return;
            }

            if (!hasText && !hasHtml)
            {
                throw new ArgumentException(
                    "body_text and/or body_html is required for '" + action + "' - the missing rendition is derived from the given one.");
            }

            if (request.BodyText is { Length: > MailService.BodyCharsCap })
            {
                throw new ArgumentException("body_text exceeds the maximum of " + MailService.BodyCharsCap + " characters.");
            }

            if (request.BodyHtml is { Length: > MailService.BodyCharsCap })
            {
                throw new ArgumentException("body_html exceeds the maximum of " + MailService.BodyCharsCap + " characters.");
            }
        }

        private static (string? Account, string? Scope) ValidateDefaultRequest(string action, ManageSignatureRequest request)
        {
            string? account = string.IsNullOrWhiteSpace(request.DefaultForAccount) ? null : request.DefaultForAccount!.Trim();
            string? scope = string.IsNullOrWhiteSpace(request.DefaultForScope) ? null : request.DefaultForScope!.Trim().ToLowerInvariant();

            if (account == null && scope == null)
            {
                return (null, null);
            }

            if (action == "delete")
            {
                throw new ArgumentException("set_default_for cannot be combined with delete - a deleted signature cannot be a default.");
            }

            if (account == null)
            {
                throw new ArgumentException("set_default_for.account is required when a scope is given.");
            }

            if (scope == null)
            {
                throw new ArgumentException("set_default_for.scope is required: 'new', 'reply' or 'both'.");
            }

            if (scope != "new" && scope != "reply" && scope != "both")
            {
                throw new ArgumentException("set_default_for.scope must be 'new', 'reply' or 'both'.");
            }

            return (account, scope);
        }

        // ------------------------------------------------------------------ file set

        /// <summary>
        /// The signature's current on-disk entries: rendition files (.htm/.html/.rtf/.txt)
        /// plus the "&lt;name&gt;_files" resource directory - exact names only, no pattern
        /// matching (7d incident discipline).
        /// </summary>
        public static IReadOnlyList<string> ExistingFileSet(string directory, string name)
        {
            List<string> entries = new List<string>();
            if (!Directory.Exists(directory))
            {
                return entries;
            }

            foreach (string extension in new[] { ".htm", ".html", ".rtf", ".txt" })
            {
                string path = Path.Combine(directory, name + extension);
                if (File.Exists(path))
                {
                    entries.Add(path);
                }
            }

            string resourceDir = Path.Combine(directory, name + "_files");
            if (Directory.Exists(resourceDir))
            {
                entries.Add(resourceDir);
            }

            return entries;
        }

        /// <summary>
        /// Copies a signature's whole file set aside before anything touches it.
        /// <para>
        /// The claim in the failure message - the operation was aborted and nothing was
        /// modified - is TRUE of the USER'S SIGNATURES, and it is the reason this runs first:
        /// nothing below it has executed when it throws. The 2026-08-19 atomicity audit
        /// confirmed it and named the one thing it did not mention, which is that a
        /// half-written BACKUP directory can survive - <c>CreateDirectory</c> succeeds and a
        /// later <c>File.Copy</c> fails. That is not the user's data and it is not a failure,
        /// but it is a directory nobody asked for, so it is named rather than left to be
        /// found.
        /// </para>
        /// </summary>
        private static string BackupFileSet(
            string root, string name, IReadOnlyList<string> existing, string backupRoot, Func<DateTime> utcNow)
        {
            string? created = null;
            try
            {
                string stamp = utcNow().ToString(BackupTimestampFormat, CultureInfo.InvariantCulture);
                string backupPath = Path.Combine(backupRoot, stamp + "-" + name);
                for (int suffix = 2; Directory.Exists(backupPath); suffix++)
                {
                    backupPath = Path.Combine(backupRoot, stamp + "-" + name + "-" + suffix);
                }

                Directory.CreateDirectory(backupPath);

                // Recorded only once the directory really exists, so the message can tell a
                // partial backup apart from a backup root that could not be written at all.
                created = backupPath;
                foreach (string entry in existing)
                {
                    if (Directory.Exists(entry))
                    {
                        CopyDirectory(entry, Path.Combine(backupPath, Path.GetFileName(entry)));
                    }
                    else
                    {
                        File.Copy(entry, Path.Combine(backupPath, Path.GetFileName(entry)));
                    }
                }

                return backupPath;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or NotSupportedException)
            {
                throw new OperationOutcomeException(
                    Com.MutationOutcome.Unchanged,
                    "The automatic backup of signature '" + name + "' failed (" + ex.Message
                    + ") - the operation was ABORTED and your signatures were not modified."
                    + (created != null
                        ? " An INCOMPLETE BACKUP folder was left at " + created + " - delete it if you do not want it."
                        : string.Empty),
                    ex);
            }
        }

        private static void CopyDirectory(string source, string target)
        {
            Directory.CreateDirectory(target);
            foreach (string file in Directory.GetFiles(source))
            {
                File.Copy(file, Path.Combine(target, Path.GetFileName(file)));
            }

            foreach (string sub in Directory.GetDirectories(source))
            {
                CopyDirectory(sub, Path.Combine(target, Path.GetFileName(sub)));
            }
        }

        private static IReadOnlyList<string> DeleteFileSet(string root, string name, IReadOnlyList<string> existing)
        {
            List<string> deleted = new List<string>(existing.Count);
            foreach (string entry in existing)
            {
                if (Directory.Exists(entry))
                {
                    Directory.Delete(entry, recursive: true);
                }
                else if (File.Exists(entry))
                {
                    File.Delete(entry);
                }

                deleted.Add(entry);
            }

            return deleted;
        }

        private static IReadOnlyList<string> WriteFileSet(string root, string name, string? bodyText, string? bodyHtml)
        {
            Directory.CreateDirectory(root);

            string text = !string.IsNullOrWhiteSpace(bodyText)
                ? bodyText!
                : Text.HtmlToText.Convert(bodyHtml ?? string.Empty);
            string html = !string.IsNullOrWhiteSpace(bodyHtml)
                ? EnsureHtmlDocument(bodyHtml!)
                : BuildHtmlFromText(text);

            string htmPath = Path.Combine(root, name + ".htm");
            string txtPath = Path.Combine(root, name + ".txt");
            string rtfPath = Path.Combine(root, name + ".rtf");

            // .htm: UTF-8 WITHOUT BOM plus an explicit charset meta (a BOM renders as
            // mojibake in older Outlook; without the meta Outlook assumes the ANSI
            // code page). .txt: UTF-16 LE WITH BOM (what Outlook itself accepts and
            // deployment tooling writes). .rtf: pure ASCII with \uN? escapes.
            File.WriteAllText(htmPath, html, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            File.WriteAllText(txtPath, text, Encoding.Unicode);
            File.WriteAllText(rtfPath, BuildRtfFromText(text), Encoding.ASCII);

            return new[] { htmPath, txtPath, rtfPath };
        }

        /// <summary>
        /// Wraps an HTML fragment into a minimal document with a utf-8 charset meta;
        /// a full document (contains &lt;html&gt;) is written as-is.
        /// </summary>
        public static string EnsureHtmlDocument(string html)
        {
            if (html.IndexOf("<html", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return html;
            }

            return "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"></head><body>\r\n"
                + html
                + "\r\n</body></html>";
        }

        /// <summary>Derives the HTML rendition from plain text (one &lt;p&gt; per line, escaped).</summary>
        public static string BuildHtmlFromText(string text)
        {
            StringBuilder body = new StringBuilder();
            foreach (string rawLine in text.Replace("\r\n", "\n").Split('\n'))
            {
                string line = System.Net.WebUtility.HtmlEncode(rawLine.TrimEnd('\r'));
                body.Append("<p>").Append(line.Length == 0 ? "&nbsp;" : line).Append("</p>\r\n");
            }

            return "<html><head><meta http-equiv=\"Content-Type\" content=\"text/html; charset=utf-8\"></head><body>\r\n"
                + body
                + "</body></html>";
        }

        /// <summary>
        /// Derives a minimal RTF rendition from plain text (RTF mail reads ONLY the
        /// .rtf file and silently drops the signature when it is missing). ASCII-safe:
        /// specials escaped, non-ASCII as \uN?.
        /// </summary>
        public static string BuildRtfFromText(string text)
        {
            StringBuilder rtf = new StringBuilder();
            rtf.Append(@"{\rtf1\ansi\ansicpg1252\deff0{\fonttbl{\f0\fnil\fcharset0 Calibri;}}\f0\fs22 ");
            string[] lines = text.Replace("\r\n", "\n").Split('\n');
            for (int i = 0; i < lines.Length; i++)
            {
                if (i > 0)
                {
                    rtf.Append(@"\par ");
                }

                foreach (char ch in lines[i].TrimEnd('\r'))
                {
                    if (ch == '\\' || ch == '{' || ch == '}')
                    {
                        rtf.Append('\\').Append(ch);
                    }
                    else if (ch > 0x7F)
                    {
                        rtf.Append(@"\u").Append(((short)ch).ToString(CultureInfo.InvariantCulture)).Append('?');
                    }
                    else
                    {
                        rtf.Append(ch);
                    }
                }
            }

            rtf.Append('}');
            return rtf.ToString();
        }

        // ------------------------------------------------------------------ defaults

        /// <summary>
        /// The ONE entry a default is written to: the mail account whose address IS
        /// <paramref name="address"/>. The rows are mail accounts only
        /// (<see cref="ISignatureDefaultsStore.ReadAccounts"/>), so a data file or an address book
        /// never gets this far, whatever it is named. None, or more than one, is a refusal decided
        /// before any file or registry work - nothing has been touched when it is thrown. The old
        /// selection took the FIRST entry whose Account Name equalled the address, out of the
        /// entries with an '@' in that name: on the 2026-09-24 guest that was a data file, and the
        /// account itself was not among them.
        /// </summary>
        private static SignatureDefaultsRow SelectDefaultTarget(IReadOnlyList<SignatureDefaultsRow> accounts, string address)
        {
            List<SignatureDefaultsRow> matches = accounts
                .Where(r => ProfileAccountEntries.AddressEquals(r.Account, address))
                .ToList();
            if (matches.Count == 1)
            {
                return matches[0];
            }

            if (matches.Count == 0)
            {
                string known = accounts.Count > 0
                    ? " Mail accounts in the profile: "
                        + string.Join(", ", accounts.Select(r => r.Account).Distinct(StringComparer.OrdinalIgnoreCase)) + "."
                    : " The profile lists no mail account this product recognises (POP3, IMAP or Exchange).";
                throw new ArgumentException(
                    "Account '" + address + "' was not found in the Outlook profile registry - set_default_for.account must be "
                    + "one of the profile's account SMTP addresses (see list_accounts), in full." + known
                    + " Only mail-account entries count: a data file or address book named after an address is not an account. "
                    + "Nothing was written.");
            }

            throw new ArgumentException(
                matches.Count.ToString(CultureInfo.InvariantCulture) + " mail accounts in the Outlook profile registry have the "
                + "address '" + address + "' (entries "
                + string.Join(", ", matches.Select(r => ProfileAccountEntries.ShortKey(r.AccountKey)))
                + ") - refusing to guess which one to record the default on. Nothing was written. Set it in Outlook instead "
                + "(File > Options > Mail > Signatures), which lists each account separately.");
        }

        private static void ApplyDefaults(ISignatureDefaultsStore store, SignatureDefaultsRow account, string scope, string name)
        {
            try
            {
                if (scope == "new" || scope == "both")
                {
                    store.WriteDefault(account.AccountKey, NewSignatureValueName, name);
                }

                if (scope == "reply" || scope == "both")
                {
                    store.WriteDefault(account.AccountKey, ReplyForwardSignatureValueName, name);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException
                or InvalidOperationException)
            {
                // Applied, not unchanged: the signature files are already on disk, and with scope
                // 'both' the first of the two values may be too.
                throw new OperationOutcomeException(
                    Com.MutationOutcome.Applied,
                    "The signature files were written, but recording it as the default for '" + account.Account
                    + "' failed: " + ex.Message, ex);
            }
        }

        /// <summary>
        /// Reads the default back before it is reported as set - and NOT through the row the write
        /// was aimed with. That is what let the 2026-09-24 defect confirm itself: list_signatures
        /// re-read the data file by the very rule that had chosen it, and said "set". Three checks,
        /// all on a FRESH read of the store: the entry written, re-read on its own and classified
        /// from its raw values, is a MAIL ACCOUNT carrying the requested address; the selection,
        /// re-run from scratch by the same rule, still lands on that one entry and no other; and
        /// the value(s) just written read back exactly. A failure is thrown with
        /// <see cref="Com.MutationOutcome.Applied"/>: the files and a registry value were written,
        /// and the message says where.
        /// </summary>
        private static void VerifyDefaults(
            ISignatureDefaultsStore store, SignatureDefaultsRow written, string address, string scope, string name)
        {
            string? problem;
            try
            {
                problem = FindDefaultsProblem(store, written, address, scope, name);
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                problem = "the profile registry could not be read back (" + ex.GetType().Name + ": " + ex.Message + ")";
            }

            if (problem != null)
            {
                throw new OperationOutcomeException(
                    Com.MutationOutcome.Applied,
                    "The signature files were written and a default was recorded on profile entry "
                    + ProfileAccountEntries.ShortKey(written.AccountKey) + ", but reading it back did not confirm it, so it is "
                    + "NOT reported as set: " + problem + ". Check the default for '" + address + "' in Outlook "
                    + "(File > Options > Mail > Signatures).");
            }
        }

        private static string? FindDefaultsProblem(
            ISignatureDefaultsStore store, SignatureDefaultsRow written, string address, string scope, string name)
        {
            ProfileAccountEntry? entry = store.ReadEntry(written.AccountKey);
            if (entry == null)
            {
                return "that entry is no longer in the profile's account list";
            }

            if (entry.Kind != ProfileEntryKind.MailAccount)
            {
                return "that entry is a " + ProfileAccountEntries.Describe(entry.Kind) + ", not a mail account";
            }

            if (!ProfileAccountEntries.AddressEquals(entry.Address, address))
            {
                return "that entry belongs to " + Quote(entry.Address) + ", not '" + address + "'";
            }

            List<SignatureDefaultsRow> now = store.ReadAccounts()
                .Where(r => ProfileAccountEntries.AddressEquals(r.Account, address))
                .ToList();
            if (now.Count != 1 || !ProfileAccountEntries.KeyEquals(now[0].AccountKey, written.AccountKey))
            {
                return "selecting the account for '" + address + "' again finds "
                    + (now.Count == 0
                        ? "no entry"
                        : "entr" + (now.Count == 1 ? "y " : "ies ")
                            + string.Join(", ", now.Select(r => ProfileAccountEntries.ShortKey(r.AccountKey))));
            }

            if ((scope == "new" || scope == "both") && !string.Equals(entry.NewSignature, name, StringComparison.Ordinal))
            {
                return "its '" + NewSignatureValueName + "' reads " + Quote(entry.NewSignature) + ", not '" + name + "'";
            }

            if ((scope == "reply" || scope == "both") && !string.Equals(entry.ReplyForwardSignature, name, StringComparison.Ordinal))
            {
                return "its '" + ReplyForwardSignatureValueName + "' reads " + Quote(entry.ReplyForwardSignature)
                    + ", not '" + name + "'";
            }

            return null;
        }

        private static string Quote(string? value)
        {
            return value == null ? "nothing" : "'" + value + "'";
        }

        private static IReadOnlyList<string>? ClearDanglingDefaults(
            ISignatureDefaultsStore store, string name, List<string> advice)
        {
            try
            {
                List<string> cleared = new List<string>();
                foreach (SignatureDefaultsRow row in store.ReadAccounts())
                {
                    bool touched = false;
                    if (string.Equals(row.NewMessage, name, StringComparison.OrdinalIgnoreCase))
                    {
                        store.ClearDefault(row.AccountKey, NewSignatureValueName);
                        touched = true;
                    }

                    if (string.Equals(row.ReplyForward, name, StringComparison.OrdinalIgnoreCase))
                    {
                        store.ClearDefault(row.AccountKey, ReplyForwardSignatureValueName);
                        touched = true;
                    }

                    if (touched)
                    {
                        cleared.Add(row.Account);
                    }
                }

                return cleared.Count > 0 ? cleared : null;
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                // The files are gone (the intended destructive step succeeded); a
                // failing assignment sweep must not fail the whole delete - report it.
                advice.Add("Clearing per-account default assignments that referenced the deleted signature failed ("
                    + ex.GetType().Name + ") - check Outlook's Signatures dialog if a default seems stale.");
                return null;
            }
        }
    }
}
