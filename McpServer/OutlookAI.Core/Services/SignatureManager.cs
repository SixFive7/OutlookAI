using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Win32;

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

    /// <summary>One profile mail account's default-signature registry state (read/write handle).</summary>
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

        /// <summary>Account SMTP address ("Account Name" registry value).</summary>
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
    /// the T1 seam; <see cref="ProfileSignatureDefaultsStore"/> is the live registry.
    /// </summary>
    public interface ISignatureDefaultsStore
    {
        /// <summary>Enumerates the profile's mail accounts with their current assignments.</summary>
        IReadOnlyList<SignatureDefaultsRow> ReadAccounts();

        /// <summary>Writes one default value (REG_SZ; absent value is created).</summary>
        void WriteDefault(string accountKey, string valueName, string signatureName);

        /// <summary>Removes one default value (no-op when absent).</summary>
        void ClearDefault(string accountKey, string valueName);
    }

    /// <summary>
    /// Live registry implementation over
    /// HKCU\Software\Microsoft\Office\&lt;major&gt;\Outlook\Profiles\&lt;default profile&gt;\9375CFF0413111d3B88A00104B2A6676,
    /// where the major is whichever Office version this machine actually has
    /// (<see cref="OutlookProfileRegistry.OfficeVersion"/> - it used to be a hardcoded 16.0).
    /// Writes are surgical: only the two known value names, only on subkeys that carry
    /// an SMTP-shaped "Account Name", never creating subkeys.
    /// </summary>
    public sealed class ProfileSignatureDefaultsStore : ISignatureDefaultsStore
    {
        // static readonly, not const: the Office major in this path is detected at runtime now.
        private static readonly string OutlookRoot = OutlookProfileRegistry.OutlookRootKeyPath;
        private const string AccountsSubKey = OutlookProfileRegistry.AccountsSubKeyName;

        /// <inheritdoc />
        public IReadOnlyList<SignatureDefaultsRow> ReadAccounts()
        {
            List<SignatureDefaultsRow> rows = new List<SignatureDefaultsRow>();
            using RegistryKey? outlook = Registry.CurrentUser.OpenSubKey(OutlookRoot);
            if (outlook == null)
            {
                return rows;
            }

            string? defaultProfile = outlook.GetValue("DefaultProfile") as string;
            using RegistryKey? profiles = outlook.OpenSubKey("Profiles");
            if (profiles == null)
            {
                return rows;
            }

            IEnumerable<string> profileNames = defaultProfile != null
                ? new[] { defaultProfile }
                : profiles.GetSubKeyNames();
            foreach (string profileName in profileNames)
            {
                string accountsPath = profileName + "\\" + AccountsSubKey;
                using RegistryKey? accounts = profiles.OpenSubKey(accountsPath);
                if (accounts == null)
                {
                    continue;
                }

                foreach (string subKeyName in accounts.GetSubKeyNames())
                {
                    using RegistryKey? account = accounts.OpenSubKey(subKeyName);
                    if (account == null)
                    {
                        continue;
                    }

                    string? address = SignatureCatalog.DecodeRegistryString(account.GetValue("Account Name"));
                    if (address == null || address.IndexOf('@') < 0)
                    {
                        continue;
                    }

                    rows.Add(new SignatureDefaultsRow(
                        OutlookRoot + "\\Profiles\\" + accountsPath + "\\" + subKeyName,
                        address.Trim(),
                        Normalize(SignatureCatalog.DecodeRegistryString(account.GetValue(SignatureManager.NewSignatureValueName))),
                        Normalize(SignatureCatalog.DecodeRegistryString(account.GetValue(SignatureManager.ReplyForwardSignatureValueName)))));
                }
            }

            return rows;
        }

        /// <inheritdoc />
        public void WriteDefault(string accountKey, string valueName, string signatureName)
        {
            using RegistryKey? account = Registry.CurrentUser.OpenSubKey(accountKey, writable: true);
            if (account == null)
            {
                throw new InvalidOperationException("The account's profile registry key no longer exists: " + accountKey);
            }

            account.SetValue(valueName, signatureName, RegistryValueKind.String);
        }

        /// <inheritdoc />
        public void ClearDefault(string accountKey, string valueName)
        {
            using RegistryKey? account = Registry.CurrentUser.OpenSubKey(accountKey, writable: true);
            account?.DeleteValue(valueName, throwOnMissingValue: false);
        }

        private static string? Normalize(string? value)
        {
            return string.IsNullOrWhiteSpace(value) ? null : value!.Trim();
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
    /// Signature" REG_SZ values (D37 locations); deleting a signature clears dangling
    /// assignments that referenced it. Pure filesystem + registry - no COM, never
    /// starts Outlook. NOTE (docs): on Microsoft 365 Apps 2303+ roaming signatures can
    /// overrule local files unless DisableRoamingSignatures=1; on Office LTSC (this
    /// machine) local files are authoritative.
    /// </summary>
    public static class SignatureManager
    {
        /// <summary>Registry value name for the new-message default.</summary>
        public const string NewSignatureValueName = "New Signature";

        /// <summary>Registry value name for the reply/forward default.</summary>
        public const string ReplyForwardSignatureValueName = "Reply-Forward Signature";

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
                targetAccount = store.ReadAccounts()
                    .FirstOrDefault(r => string.Equals(r.Account, defaultAccount, StringComparison.OrdinalIgnoreCase));
                if (targetAccount == null)
                {
                    throw new ArgumentException(
                        "Account '" + defaultAccount + "' was not found in the Outlook profile registry - "
                        + "set_default_for.account must be one of the profile's account SMTP addresses (see list_accounts).");
                }
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

        private static string BackupFileSet(
            string root, string name, IReadOnlyList<string> existing, string backupRoot, Func<DateTime> utcNow)
        {
            try
            {
                string stamp = utcNow().ToString(BackupTimestampFormat, CultureInfo.InvariantCulture);
                string backupPath = Path.Combine(backupRoot, stamp + "-" + name);
                for (int suffix = 2; Directory.Exists(backupPath); suffix++)
                {
                    backupPath = Path.Combine(backupRoot, stamp + "-" + name + "-" + suffix);
                }

                Directory.CreateDirectory(backupPath);
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
                throw new InvalidOperationException(
                    "The automatic backup of signature '" + name + "' failed (" + ex.Message
                    + ") - the " + "operation was ABORTED and nothing was modified.", ex);
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
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException)
            {
                throw new InvalidOperationException(
                    "The signature files were written, but recording it as the default for '" + account.Account
                    + "' failed: " + ex.Message, ex);
            }
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
