using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OutlookAI.Core.Services
{
    /// <summary>One in-place text edit: every occurrence of <see cref="Find"/> becomes <see cref="Replace"/>, in every rendition that has it.</summary>
    public sealed class SignatureTextEdit
    {
        /// <summary>The visible text to find, exactly as it reads (not as the file spells it).</summary>
        public string Find { get; set; } = string.Empty;

        /// <summary>The text to put in its place (empty removes it).</summary>
        public string Replace { get; set; } = string.Empty;
    }

    /// <summary>One change to the signature's picture folder (<c>&lt;name&gt;_files</c>).</summary>
    public sealed class SignaturePictureChange
    {
        /// <summary>The file name in the picture folder, e.g. <c>image001.png</c> or <c>logo.png</c>.</summary>
        public string File { get; set; } = string.Empty;

        /// <summary>An absolute path to copy in: adds the file, or replaces it when it exists. Null with <see cref="Remove"/>.</summary>
        public string? Path { get; set; }

        /// <summary>True to delete the file from the picture folder.</summary>
        public bool Remove { get; set; }
    }

    /// <summary>Validated manage_signature request (soak fix D38, Q147 and Q159).</summary>
    public sealed class ManageSignatureRequest
    {
        /// <summary>"create" | "update" | "delete" | "set_default" | "clear_default" (case-insensitive).</summary>
        public string Action { get; set; } = string.Empty;

        /// <summary>Signature name (file base name, what Outlook's pickers show). For clear_default: the default being cleared.</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// Plain-text rendition. Create: written (derived from another when omitted). Update: the new .txt, sent
        /// with every other rendition the signature has (Q205 (d)).
        /// </summary>
        public string? BodyText { get; set; }

        /// <summary>
        /// HTML rendition, fragment or full document. Create: written. Update: the new .htm, sent with every other
        /// rendition the signature has (Q205 (d)).
        /// </summary>
        public string? BodyHtml { get; set; }

        /// <summary>
        /// RTF rendition, with <c>[[picture:N]]</c> / <c>[[picture-file:...]]</c> tokens for pictures. Update: the new
        /// .rtf, sent with every other rendition the signature has (Q205 (d)).
        /// </summary>
        public string? BodyRtf { get; set; }

        /// <summary>In-place text edits applied to every rendition (update only).</summary>
        public IReadOnlyList<SignatureTextEdit>? Replace { get; set; }

        /// <summary>Picture folder changes (create: adds only; update: add, replace, remove).</summary>
        public IReadOnlyList<SignaturePictureChange>? Pictures { get; set; }

        /// <summary>The Outlook profile whose account default is set or cleared - always explicit (Q155).</summary>
        public string? DefaultForProfile { get; set; }

        /// <summary>SMTP address of the account whose default is set or cleared.</summary>
        public string? DefaultForAccount { get; set; }

        /// <summary>"new" | "reply" | "both" - which default(s).</summary>
        public string? DefaultForScope { get; set; }
    }

    /// <summary>A manage_signature refusal made for a reason outside the request's own shape: policy, roaming.</summary>
    public sealed class SignatureRefusedException : ArgumentException
    {
        /// <summary>Creates the refusal.</summary>
        public SignatureRefusedException(string reason, string message)
            : base(message)
        {
            Reason = reason;
        }

        /// <summary>
        /// Machine-readable reason: <c>signatures_disabled_by_policy</c>, <c>default_set_by_policy</c>,
        /// <c>managed_in_mailbox</c>, <c>account_may_roam</c>, <c>signature_forced_by_policy</c>,
        /// <c>signature_of_another_account</c>.
        /// </summary>
        public string Reason { get; }
    }

    /// <summary>One MAIL ACCOUNT's default-signature registry state in one profile (read/write handle).</summary>
    public sealed class SignatureDefaultsRow
    {
        /// <summary>Creates a row.</summary>
        public SignatureDefaultsRow(
            string accountKey,
            string account,
            string? newMessage,
            string? replyForward,
            string? profile = null,
            MailAccountKind? accountKind = null,
            string? accountName = null)
        {
            AccountKey = accountKey;
            Account = account;
            NewMessage = newMessage;
            ReplyForward = replyForward;
            Profile = profile ?? ProfileAccountEntries.ProfileNameOf(accountKey);
            AccountKind = accountKind;
            AccountName = accountName;
        }

        /// <summary>Opaque store handle of the account (registry subkey path for the production store).</summary>
        public string AccountKey { get; }

        /// <summary>
        /// The account's email address - what <c>list_accounts</c> reports and
        /// <c>default_for.account</c> is matched against. <see cref="ProfileAccountEntries"/>
        /// derives it: <c>Email</c> for a POP3/IMAP account, <c>Account Name</c> for an Exchange
        /// account (which records no <c>Email</c>).
        /// </summary>
        public string Account { get; }

        /// <summary>Currently assigned new-message signature name (null = absent).</summary>
        public string? NewMessage { get; }

        /// <summary>Currently assigned reply/forward signature name (null = absent).</summary>
        public string? ReplyForward { get; }

        /// <summary>The profile this account entry belongs to.</summary>
        public string? Profile { get; }

        /// <summary>POP3, IMAP or Exchange.</summary>
        public MailAccountKind? AccountKind { get; }

        /// <summary>The entry's <c>Account Name</c> - for an Exchange account, the name Outlook shows it under.</summary>
        public string? AccountName { get; }
    }

    /// <summary>One default set or cleared, in one profile (data only).</summary>
    public sealed class SignatureDefaultChange
    {
        /// <summary>Creates the description.</summary>
        public SignatureDefaultChange(string profile, string account, string scope, string? previous, string? now)
        {
            Profile = profile;
            Account = account;
            Scope = scope;
            Previous = previous;
            Now = now;
        }

        /// <summary>The profile.</summary>
        public string Profile { get; }

        /// <summary>The account's address.</summary>
        public string Account { get; }

        /// <summary>"new" or "reply".</summary>
        public string Scope { get; }

        /// <summary>The value before the change (null = none) - what restores it.</summary>
        public string? Previous { get; }

        /// <summary>The value after the change (null = cleared).</summary>
        public string? Now { get; }
    }

    /// <summary>
    /// Read/write access to the per-account default-signature registry values
    /// ("New Signature" / "Reply-Forward Signature" under each profile's 9375CFF0...
    /// key). The interface is the T1 seam; <see cref="ProfileSignatureDefaultsStore"/> is the live
    /// registry (or, for T1, any <see cref="IProfileAccountRegistry"/>).
    /// </summary>
    public interface ISignatureDefaultsStore
    {
        /// <summary>
        /// Every profile's MAIL ACCOUNTS - never a data file or an address book, whatever it is named
        /// (<see cref="ProfileAccountEntries"/>) - with their current assignments, each row with its profile.
        /// </summary>
        IReadOnlyList<SignatureDefaultsRow> ReadAccounts();

        /// <summary>The profiles and the recorded default profile.</summary>
        ProfileList ReadProfiles();

        /// <summary>
        /// Re-reads ONE entry from the store, now, and classifies it from its raw values - null
        /// when it is no longer an entry of a profile's account list. The independent half of
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
    /// HKCU\Software\Microsoft\Office\&lt;major&gt;\Outlook\Profiles\&lt;profile&gt;\9375CFF0413111d3B88A00104B2A6676
    /// in every profile (<see cref="LiveProfileAccountRegistry"/>), or any <see cref="IProfileAccountRegistry"/>
    /// for T1. Which entries are accounts, and what their addresses are, is
    /// <see cref="ProfileAccountEntries"/>' rule. Writes are surgical: only the two known value names,
    /// only on an entry that - re-read at the moment of the write - classifies as a MAIL ACCOUNT,
    /// never creating an entry.
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

        /// <summary>The registry this store reads and writes.</summary>
        public IProfileAccountRegistry Registry => _registry;

        /// <inheritdoc />
        public IReadOnlyList<SignatureDefaultsRow> ReadAccounts()
        {
            return ProfileAccountEntries.ReadAll(_registry)
                .Where(e => e.Kind == ProfileEntryKind.MailAccount && e.Address != null)
                .Select(e => new SignatureDefaultsRow(
                    e.Key, e.Address!, e.NewSignature, e.ReplyForwardSignature, e.ProfileName, e.AccountKind, e.AccountName))
                .ToList();
        }

        /// <inheritdoc />
        public ProfileList ReadProfiles()
        {
            return _registry.ReadProfiles();
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
                    "The account's profile registry entry is no longer there: " + accountKey);
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
    /// Creates, edits and deletes Outlook signatures, and sets and clears their per-account defaults
    /// (manage_signature - soak fix D38, Q147, Q159).
    /// <para>
    /// <b>Files.</b> A signature is up to three renditions in the folder Outlook uses
    /// (<see cref="SignatureFolder"/>) - <c>.htm</c> for HTML mail, <c>.rtf</c> for Rich Text, <c>.txt</c>
    /// for plain text, each read only by its own mail format - plus a <c>&lt;name&gt;_files</c> folder
    /// holding the HTML rendition's pictures (the RTF embeds its own copies). A CREATE writes all
    /// three, deriving the ones not given, because a missing rendition silently drops the signature
    /// from that mail format. An UPDATE edits each rendition IN PLACE and touches only what was asked:
    /// a supplied rendition replaces that one file, a text edit is applied where the text appears,
    /// picture changes touch only the files named, and nothing is rebuilt from another rendition or
    /// re-encoded (Q159). It used to delete the picture folder and rebuild the .rtf from plain text on
    /// every update, which broke any signature Outlook had made with a logo.
    /// </para>
    /// <para>
    /// <b>Safety, always on.</b> Before any update or delete the signature's whole file set is copied
    /// to %LOCALAPPDATA%\OutlookAI\signature-backups\&lt;utc&gt;-&lt;name&gt;\ and the path returned - a
    /// failing backup ABORTS the operation. Every refusal is decided before anything is written.
    /// </para>
    /// <para>
    /// <b>Defaults.</b> Set and cleared on their own (set_default, clear_default) or with a create or
    /// update, always for an EXPLICIT profile and account (Q155), on the ONE mail account of that
    /// profile with that address, and read back before they are reported. Deleting a signature clears
    /// every default, in every profile, that named it.
    /// </para>
    /// <para>
    /// <b>Refused, with the reason (Q152, Q156):</b> any write while policy disables signatures; a
    /// default for a scope a policy or deployed setting fixes for every account; deleting a signature
    /// a policy forces; local writes to Outlook's copy of a signature kept in a mailbox (roaming); and
    /// a default for an Exchange account while this Outlook may roam signatures - the mailbox's
    /// signature data is never written, not even indirectly. Pure filesystem + registry - no COM,
    /// never starts Outlook.
    /// </para>
    /// </summary>
    public static class SignatureManager
    {
        /// <summary>Registry value name for the new-message default.</summary>
        public const string NewSignatureValueName = ProfileAccountEntries.NewSignatureValueName;

        /// <summary>Registry value name for the reply/forward default.</summary>
        public const string ReplyForwardSignatureValueName = ProfileAccountEntries.ReplyForwardSignatureValueName;

        /// <summary>Maximum signature name length (file-name discipline).</summary>
        public const int NameMaxChars = 128;

        /// <summary>Most text edits one call may carry.</summary>
        public const int MaxTextEdits = 50;

        /// <summary>Most picture changes one call may carry.</summary>
        public const int MaxPictureChanges = 20;

        /// <summary>Timestamp format of backup directory names (UTC, filesystem-safe).</summary>
        public const string BackupTimestampFormat = "yyyyMMdd'T'HHmmssfff'Z'";

        /// <summary>The reserved word the draft tools read as "no signature"; never a signature name.</summary>
        public const string NoSignatureKeyword = "none";

        /// <summary>Advice returned whenever registry defaults were written or cleared.</summary>
        public const string DefaultsRestartAdvice =
            "Outlook picks up default-signature changes at its next start (or when its Signatures dialog is reopened) - "
            + "already-open compose windows keep the old default. OutlookAI's own drafts use the new default at once.";

        private static readonly string[] Actions = { "create", "update", "delete", "set_default", "clear_default" };

        /// <summary>Default backup root: %LOCALAPPDATA%\OutlookAI\signature-backups.</summary>
        public static string DefaultBackupRoot =>
            Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "OutlookAI", "signature-backups");

        /// <summary>
        /// Executes one manage_signature operation. <paramref name="directory"/>,
        /// <paramref name="backupRoot"/>, <paramref name="defaultsStore"/>, <paramref name="utcNow"/> and
        /// <paramref name="settings"/> are T1 seams; production passes null for all. Every refusal is an
        /// <see cref="ArgumentException"/> thrown BEFORE any file or registry work; a failing backup is
        /// an <see cref="OperationOutcomeException"/> thrown before anything is modified.
        /// </summary>
        public static ManageSignatureOutcome Manage(
            ManageSignatureRequest request,
            string? directory = null,
            string? backupRoot = null,
            ISignatureDefaultsStore? defaultsStore = null,
            Func<DateTime>? utcNow = null,
            ISignatureSettingsRegistry? settings = null)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            string action = ValidateAction(request.Action);
            string name = ValidateName(request.Name);
            ISignatureSettingsRegistry registry = settings ?? new LiveSignatureSettingsRegistry();
            string officeVersion = OutlookProfileRegistry.OfficeVersion;
            string root = directory ?? SignatureFolder.Read(
                registry, officeVersion, Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)).Path;
            SignaturePolicy policy = SignaturePolicy.Read(registry, officeVersion);
            ISignatureDefaultsStore store = defaultsStore ?? new ProfileSignatureDefaultsStore();
            SignatureRoaming roaming = SignatureRoaming.Read(registry, officeVersion);
            if (roaming.EvidenceDecides)
            {
                // Outlook's own copies decide on a perpetual Office (Q152): the folder's names against the
                // profiles' Exchange accounts.
                roaming = roaming.WithEvidence(
                    SignatureCatalog.ListSignatures(root).Select(s => s.Name),
                    store.ReadAccounts().Where(r => r.AccountKind == MailAccountKind.Exchange).Select(r => r.Account));
            }

            // The administrator's rule first: with signatures disabled nothing below runs (Q156).
            if (policy.SignaturesDisabled)
            {
                throw new SignatureRefusedException("signatures_disabled_by_policy",
                    "Signatures are disabled on this computer by " + SignaturePolicy.Describe(policy.DisabledBy)
                    + " (DisableSignatures): Outlook inserts none and offers no way to manage them, so OutlookAI does not "
                    + "either. Nothing was written. Ask your administrator if you need a signature.");
            }

            ValidateShape(action, request);
            IReadOnlyList<string> existing = ExistingFileSet(root, name);
            string? scopeAccount = SignatureRoaming.AccountSuffixOf(name);
            bool managedInMailbox = scopeAccount != null && roaming.MayRoamFor(scopeAccount);

            switch (action)
            {
                case "create":
                    if (existing.Count > 0)
                    {
                        throw new ArgumentException(
                            "Signature '" + name + "' already exists - use action 'update' to change it (an automatic backup is made).");
                    }

                    if (managedInMailbox)
                    {
                        throw new SignatureRefusedException("managed_in_mailbox",
                            "'" + name + "' is shaped like Outlook's local copy of a roaming signature of " + scopeAccount
                            + " ('name (address)'), and " + roaming.Reason + ". Choose a name without the address in brackets. "
                            + "Nothing was written.");
                    }

                    break;
                case "update":
                case "delete":
                    if (existing.Count == 0)
                    {
                        throw new ArgumentException(
                            "Signature '" + name + "' was not found in " + root + ". Use list_signatures for the installed signature names.");
                    }

                    if (managedInMailbox)
                    {
                        throw ManagedInMailboxRefusal(name, scopeAccount!, roaming);
                    }

                    if (action == "delete")
                    {
                        RefuseDeletingAForcedSignature(name, policy);
                    }

                    break;
                case "set_default":
                    if (existing.Count == 0)
                    {
                        throw new ArgumentException(
                            "Signature '" + name + "' was not found in " + root + ", so it cannot be a default. Use list_signatures.");
                    }

                    break;
            }

            // The default this call sets or clears, decided before anything is written.
            DefaultTarget? target = null;
            if (request.DefaultForAccount != null || request.DefaultForProfile != null || request.DefaultForScope != null
                || action == "set_default" || action == "clear_default")
            {
                target = ResolveDefaultTarget(store, request, action, name, policy, roaming, scopeAccount);
            }

            ContentPlan? plan = action == "create" || action == "update"
                ? PlanContent(action, root, name, request)
                : null;

            ManageSignatureOutcome outcome = new ManageSignatureOutcome
            {
                Action = action,
                Name = name,
                Folder = root,
            };
            List<string> advice = new List<string>();

            if (action == "update" || action == "delete")
            {
                outcome.BackupPath = BackupFileSet(root, name, existing, backupRoot ?? DefaultBackupRoot, utcNow ?? (() => DateTime.UtcNow));
            }

            if (action == "delete")
            {
                outcome.FilesDeleted = DeleteFileSet(existing);
                outcome.DefaultsCleared = ClearDanglingDefaults(store, name, advice);
                if (outcome.DefaultsCleared is { Count: > 0 })
                {
                    advice.Add(DefaultsRestartAdvice);
                }
            }

            if (plan != null)
            {
                WritePlan(root, name, plan, outcome);
                advice.AddRange(plan.Warnings);
            }

            if (target != null)
            {
                outcome.DefaultsChanged = ApplyDefaults(store, target, action == "clear_default" ? null : name, plan != null);
                outcome.Profile = target.Row.Profile;
                outcome.DefaultSetForAccount = action == "clear_default" ? null : target.Row.Account;
                outcome.DefaultSetScope = action == "clear_default" ? null : target.Scope;
                advice.Add(DefaultsRestartAdvice);
            }

            outcome.Advice = advice.Count > 0 ? string.Join(" ", advice.Distinct()) : null;
            return outcome;
        }

        // ------------------------------------------------------------------ validation

        private static string ValidateAction(string? action)
        {
            string normalized = (action ?? string.Empty).Trim().ToLowerInvariant();
            if (!Actions.Contains(normalized, StringComparer.Ordinal))
            {
                throw new ArgumentException("action must be 'create', 'update', 'delete', 'set_default' or 'clear_default'.");
            }

            return normalized;
        }

        /// <summary>The checks every signature name is held to - a file name Outlook can show and Windows can store.</summary>
        public static string ValidateName(string? name)
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

            // The signature's files are <name>.htm, .rtf and .txt, and Windows reads a device name
            // before ANY extension as the device - so "CON" and "nul.backup" are refused alike.
            if (WindowsFileNames.IsReservedDeviceName(trimmed))
            {
                throw new ArgumentException("name '" + trimmed + "' is a reserved Windows device name (CON, NUL, COM1 and the like, alone or before a dot).");
            }

            if (string.Equals(trimmed, NoSignatureKeyword, StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException(
                    "name '" + trimmed + "' is reserved: the draft tools read signature=\"none\" as \"no signature\", so a signature "
                    + "of that name could never be chosen. Pick another name.");
            }

            return trimmed;
        }

        private static void ValidateShape(string action, ManageSignatureRequest request)
        {
            bool hasText = !string.IsNullOrWhiteSpace(request.BodyText);
            bool hasHtml = !string.IsNullOrWhiteSpace(request.BodyHtml);
            bool hasRtf = !string.IsNullOrWhiteSpace(request.BodyRtf);
            int edits = request.Replace?.Count ?? 0;
            int pictures = request.Pictures?.Count ?? 0;
            bool content = hasText || hasHtml || hasRtf || edits > 0 || pictures > 0;
            bool defaults = request.DefaultForAccount != null || request.DefaultForProfile != null || request.DefaultForScope != null;

            switch (action)
            {
                case "create":
                    if (!hasText && !hasHtml && !hasRtf)
                    {
                        throw new ArgumentException(
                            "body_text, body_html and/or body_rtf is required for 'create' - the renditions not given are derived from "
                            + "the ones that are.");
                    }

                    if (edits > 0)
                    {
                        throw new ArgumentException("replace edits an existing signature - use it with 'update', not 'create'.");
                    }

                    if (request.Pictures != null && request.Pictures.Any(p => p.Remove))
                    {
                        throw new ArgumentException("A new signature has no pictures to remove - 'pictures' with 'create' only adds.");
                    }

                    break;
                case "update":
                    if (!content)
                    {
                        throw new ArgumentException(
                            defaults
                                ? "'update' changes a signature's content; to change only a default use action 'set_default' or 'clear_default'."
                                : "Nothing to update: supply body_html, body_rtf and/or body_text (each replaces only its own rendition), "
                                    + "replace (in-place text edits), and/or pictures.");
                    }

                    break;
                default:
                    if (content)
                    {
                        throw new ArgumentException(
                            action + " takes no content - body_text, body_html, body_rtf, replace and pictures apply to create and update only.");
                    }

                    if (action == "delete" && defaults)
                    {
                        throw new ArgumentException("default_for cannot be combined with delete - a deleted signature cannot be a default.");
                    }

                    break;
            }

            if (defaults || action == "set_default" || action == "clear_default")
            {
                ValidateDefaultShape(request);
            }

            CheckCap(request.BodyText, "body_text");
            CheckCap(request.BodyHtml, "body_html");
            CheckCap(request.BodyRtf, "body_rtf");
            if (edits > MaxTextEdits)
            {
                throw new ArgumentException("replace may carry at most " + MaxTextEdits + " edits per call.");
            }

            if (pictures > MaxPictureChanges)
            {
                throw new ArgumentException("pictures may carry at most " + MaxPictureChanges + " changes per call.");
            }

            foreach (SignatureTextEdit edit in request.Replace ?? Array.Empty<SignatureTextEdit>())
            {
                if (edit == null || string.IsNullOrEmpty(edit.Find))
                {
                    throw new ArgumentException("every replace edit needs a non-empty 'find'.");
                }

                CheckCap(edit.Find, "replace.find");
                CheckCap(edit.Replace, "replace.replace");
            }
        }

        /// <summary>
        /// default_for's own shape - all three parts, a known scope - checked before anything is read,
        /// so a malformed request is refused the same on every machine. Whether the profile and the
        /// account exist is <see cref="ResolveDefaultTarget"/>'s question.
        /// </summary>
        private static void ValidateDefaultShape(ManageSignatureRequest request)
        {
            string? scope = string.IsNullOrWhiteSpace(request.DefaultForScope) ? null : request.DefaultForScope!.Trim().ToLowerInvariant();
            if (scope != null && scope != "new" && scope != "reply" && scope != "both")
            {
                throw new ArgumentException("default_for.scope must be 'new', 'reply' or 'both'. Nothing was written.");
            }

            if (string.IsNullOrWhiteSpace(request.DefaultForProfile))
            {
                throw new ArgumentException(
                    "default_for.profile is required: a default belongs to one Outlook profile, and OutlookAI never guesses which "
                    + "(Q155) - list_signatures lists the profiles. Nothing was written.");
            }

            if (string.IsNullOrWhiteSpace(request.DefaultForAccount))
            {
                throw new ArgumentException(
                    "default_for.account is required: the address of the account whose default changes (list_signatures lists "
                    + "each profile's accounts). Nothing was written.");
            }

            if (scope == null)
            {
                throw new ArgumentException("default_for.scope is required: 'new', 'reply' or 'both'. Nothing was written.");
            }
        }

        private static void CheckCap(string? value, string field)
        {
            if (value is { Length: > MailService.BodyCharsCap })
            {
                throw new ArgumentException(field + " exceeds the maximum of " + MailService.BodyCharsCap + " characters.");
            }
        }

        private static SignatureRefusedException ManagedInMailboxRefusal(string name, string account, SignatureRoaming roaming)
        {
            return new SignatureRefusedException("managed_in_mailbox",
                "'" + name + "' is Outlook's local copy of a signature kept in the mailbox of " + account
                + " (roaming signatures: " + roaming.Reason + "). Outlook overwrites local changes to it at its next sync, and "
                + "OutlookAI never writes a mailbox's signature data. Change it in Outlook (File > Options > Mail > Signatures) "
                + "or in Outlook on the web. Nothing was written.");
        }

        private static void RefuseDeletingAForcedSignature(string name, SignaturePolicy policy)
        {
            foreach (bool reply in new[] { false, true })
            {
                if (policy.LocksSignatureFor(reply, name))
                {
                    throw new SignatureRefusedException("signature_forced_by_policy",
                        "'" + name + "' is the signature " + SignaturePolicy.Describe(policy.SourceFor(reply)) + " names for "
                        + (reply ? "replies and forwards" : "new messages") + "; deleting it would leave the policy naming a signature "
                        + "that is not there. Nothing was written. Ask your administrator to change the policy first.");
                }
            }
        }

        // ------------------------------------------------------------------ defaults: deciding

        private sealed class DefaultTarget
        {
            public DefaultTarget(SignatureDefaultsRow row, string scope)
            {
                Row = row;
                Scope = scope;
            }

            public SignatureDefaultsRow Row { get; }

            /// <summary>new | reply | both.</summary>
            public string Scope { get; }
        }

        private static DefaultTarget ResolveDefaultTarget(
            ISignatureDefaultsStore store,
            ManageSignatureRequest request,
            string action,
            string name,
            SignaturePolicy policy,
            SignatureRoaming roaming,
            string? scopeAccount)
        {
            string? account = string.IsNullOrWhiteSpace(request.DefaultForAccount) ? null : request.DefaultForAccount!.Trim();
            string? profile = string.IsNullOrWhiteSpace(request.DefaultForProfile) ? null : request.DefaultForProfile!.Trim();
            string? scope = string.IsNullOrWhiteSpace(request.DefaultForScope) ? null : request.DefaultForScope!.Trim().ToLowerInvariant();
            IReadOnlyList<SignatureDefaultsRow> rows = store.ReadAccounts();
            IReadOnlyList<string> profiles = KnownProfiles(store, rows);

            // The shape - all three parts, a known scope - was checked before anything was read (ValidateDefaultShape).
            string? recordedProfile = profiles.FirstOrDefault(p => ProfileAccountEntries.ProfileEquals(p, profile));
            if (recordedProfile == null)
            {
                throw new ArgumentException(
                    "Outlook profile '" + profile + "' does not exist. Profiles on this computer: " + Describe(profiles) + ". Nothing was written.");
            }

            SignatureDefaultsRow row = SelectDefaultTarget(rows, recordedProfile, account!);

            foreach (bool reply in ScopesOf(scope!))
            {
                if (policy.TakesDefault(reply))
                {
                    throw new SignatureRefusedException("default_set_by_policy",
                        SignaturePolicy.Describe(policy.SourceFor(reply)) + " names the default signature for "
                        + (reply ? "replies and forwards" : "new messages") + ", '" + policy.DefaultFor(reply) + "', and OutlookAI does "
                        + "not change an account's default under a policy (Q156). Nothing was written.");
                }
            }

            if (row.AccountKind == MailAccountKind.Exchange && roaming.MayRoamFor(row.Account))
            {
                throw new SignatureRefusedException("account_may_roam",
                    row.Account + " is an Exchange account, and on this Outlook its signatures and defaults may be kept in its "
                    + "mailbox (roaming signatures: " + roaming.Reason + "). A default written here could be overruled by the "
                    + "mailbox's or copied into it, and OutlookAI never writes a mailbox's signature data. Change it in Outlook "
                    + "(File > Options > Mail > Signatures). Nothing was written.");
            }

            if (action != "clear_default" && scopeAccount != null && roaming.MayRoamFor(scopeAccount)
                && !ProfileAccountEntries.AddressEquals(scopeAccount, row.Account))
            {
                throw new SignatureRefusedException("signature_of_another_account",
                    "'" + name + "' is a roaming signature of " + scopeAccount + ", so Outlook does not offer it to " + row.Account
                    + ". Nothing was written.");
            }

            if (action == "clear_default")
            {
                foreach (bool reply in ScopesOf(scope!))
                {
                    string? current = reply ? row.ReplyForward : row.NewMessage;
                    if (!string.Equals(current, name, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new ArgumentException(
                            "The " + (reply ? "reply/forward" : "new-message") + " default of " + row.Account + " in profile '"
                            + row.Profile + "' is " + (current == null ? "not set" : "'" + current + "'") + ", not '" + name
                            + "'. clear_default clears only the default it names, so nothing is cleared by mistake. Nothing was written.");
                    }
                }
            }

            return new DefaultTarget(row, scope!);
        }

        private static IEnumerable<bool> ScopesOf(string scope)
        {
            if (scope == "new" || scope == "both")
            {
                yield return false;
            }

            if (scope == "reply" || scope == "both")
            {
                yield return true;
            }
        }

        private static IReadOnlyList<string> KnownProfiles(ISignatureDefaultsStore store, IReadOnlyList<SignatureDefaultsRow> rows)
        {
            List<string> profiles = new List<string>(store.ReadProfiles().Names);
            foreach (string? profile in rows.Select(r => r.Profile))
            {
                if (profile != null && !profiles.Any(p => ProfileAccountEntries.ProfileEquals(p, profile)))
                {
                    profiles.Add(profile);
                }
            }

            return profiles.OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static string Describe(IReadOnlyList<string> profiles)
        {
            return profiles.Count == 0 ? "none found" : string.Join(", ", profiles.Select(p => "'" + p + "'"));
        }

        /// <summary>
        /// The ONE entry a default is written to: the mail account of <paramref name="profile"/> whose
        /// address IS <paramref name="address"/>. Another profile's account is never a candidate - the
        /// profiles used to be merged when none was recorded as the default, and the first match won.
        /// None, or more than one, is a refusal decided before any file or registry work.
        /// </summary>
        private static SignatureDefaultsRow SelectDefaultTarget(IReadOnlyList<SignatureDefaultsRow> rows, string profile, string address)
        {
            List<SignatureDefaultsRow> inProfile = rows.Where(r => ProfileAccountEntries.ProfileEquals(r.Profile, profile)).ToList();
            List<SignatureDefaultsRow> matches = inProfile
                .Where(r => ProfileAccountEntries.AddressEquals(r.Account, address))
                .ToList();
            if (matches.Count == 1)
            {
                return matches[0];
            }

            if (matches.Count == 0)
            {
                string known = inProfile.Count > 0
                    ? " Mail accounts in profile '" + profile + "': "
                        + string.Join(", ", inProfile.Select(r => r.Account).Distinct(StringComparer.OrdinalIgnoreCase)) + "."
                    : " Profile '" + profile + "' lists no mail account this product recognises (POP3, IMAP or Exchange).";
                throw new ArgumentException(
                    "Account '" + address + "' is not a mail account of Outlook profile '" + profile + "' - default_for.account must "
                    + "be one of that profile's account addresses, in full." + known
                    + " Only mail-account entries count: a data file or address book named after an address is not an account. "
                    + "Nothing was written.");
            }

            throw new ArgumentException(
                matches.Count.ToString(CultureInfo.InvariantCulture) + " mail accounts in Outlook profile '" + profile + "' have the "
                + "address '" + address + "' (entries "
                + string.Join(", ", matches.Select(r => ProfileAccountEntries.ShortKey(r.AccountKey)))
                + ") - refusing to guess which one to record the default on. Nothing was written. Set it in Outlook instead "
                + "(File > Options > Mail > Signatures), which lists each account separately.");
        }

        // ------------------------------------------------------------------ defaults: writing

        private static IReadOnlyList<SignatureDefaultChange> ApplyDefaults(
            ISignatureDefaultsStore store, DefaultTarget target, string? name, bool filesWritten)
        {
            List<SignatureDefaultChange> changes = new List<SignatureDefaultChange>();
            try
            {
                foreach (bool reply in ScopesOf(target.Scope))
                {
                    string valueName = reply ? ReplyForwardSignatureValueName : NewSignatureValueName;
                    string? previous = reply ? target.Row.ReplyForward : target.Row.NewMessage;
                    if (name == null)
                    {
                        store.ClearDefault(target.Row.AccountKey, valueName);
                    }
                    else
                    {
                        store.WriteDefault(target.Row.AccountKey, valueName, name);
                    }

                    changes.Add(new SignatureDefaultChange(
                        target.Row.Profile ?? string.Empty, target.Row.Account, reply ? "reply" : "new", previous, name));
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or System.Security.SecurityException
                or InvalidOperationException)
            {
                // Applied, not unchanged: the files may already be on disk, and with scope 'both' the
                // first of the two values may be too.
                throw new OperationOutcomeException(
                    filesWritten || changes.Count > 0 ? Com.MutationOutcome.Applied : Com.MutationOutcome.Unknown,
                    (filesWritten ? "The signature files were written, but " : string.Empty)
                    + (name == null ? "clearing" : "recording") + " the default of '" + target.Row.Account + "' in profile '"
                    + target.Row.Profile + "' failed: " + ex.Message, ex);
            }

            VerifyDefaults(store, target, name, filesWritten);
            return changes;
        }

        /// <summary>
        /// Reads the default back before it is reported as set - and NOT through the row the write
        /// was aimed with. Three checks, all on a FRESH read of the store: the entry written, re-read on
        /// its own and classified from its raw values, is a MAIL ACCOUNT of the same profile carrying the
        /// requested address; the selection, re-run from scratch, still lands on that one entry; and the
        /// value(s) read back exactly (cleared ones absent). A failure is thrown with
        /// <see cref="Com.MutationOutcome.Applied"/>.
        /// </summary>
        private static void VerifyDefaults(ISignatureDefaultsStore store, DefaultTarget target, string? name, bool filesWritten)
        {
            string? problem;
            try
            {
                problem = FindDefaultsProblem(store, target, name);
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                problem = "the profile registry could not be read back (" + ex.GetType().Name + ": " + ex.Message + ")";
            }

            if (problem != null)
            {
                throw new OperationOutcomeException(
                    Com.MutationOutcome.Applied,
                    (filesWritten ? "The signature files were written and a" : "A") + " default was "
                    + (name == null ? "cleared" : "recorded") + " on profile entry " + ProfileAccountEntries.ShortKey(target.Row.AccountKey)
                    + ", but reading it back did not confirm it, so it is NOT reported as " + (name == null ? "cleared" : "set") + ": "
                    + problem + ". Check the default for '"
                    + target.Row.Account + "' in Outlook (File > Options > Mail > Signatures).");
            }
        }

        private static string? FindDefaultsProblem(ISignatureDefaultsStore store, DefaultTarget target, string? name)
        {
            ProfileAccountEntry? entry = store.ReadEntry(target.Row.AccountKey);
            if (entry == null)
            {
                return "that entry is no longer in the profile's account list";
            }

            if (entry.Kind != ProfileEntryKind.MailAccount)
            {
                return "that entry is a " + ProfileAccountEntries.Describe(entry.Kind) + ", not a mail account";
            }

            if (!ProfileAccountEntries.AddressEquals(entry.Address, target.Row.Account))
            {
                return "that entry belongs to " + Quote(entry.Address) + ", not '" + target.Row.Account + "'";
            }

            if (target.Row.Profile != null && entry.ProfileName != null
                && !ProfileAccountEntries.ProfileEquals(entry.ProfileName, target.Row.Profile))
            {
                return "that entry belongs to profile '" + entry.ProfileName + "', not '" + target.Row.Profile + "'";
            }

            List<SignatureDefaultsRow> now = store.ReadAccounts()
                .Where(r => ProfileAccountEntries.ProfileEquals(r.Profile, target.Row.Profile)
                    && ProfileAccountEntries.AddressEquals(r.Account, target.Row.Account))
                .ToList();
            if (now.Count != 1 || !ProfileAccountEntries.KeyEquals(now[0].AccountKey, target.Row.AccountKey))
            {
                return "selecting the account for '" + target.Row.Account + "' again finds "
                    + (now.Count == 0
                        ? "no entry"
                        : "entr" + (now.Count == 1 ? "y " : "ies ")
                            + string.Join(", ", now.Select(r => ProfileAccountEntries.ShortKey(r.AccountKey))));
            }

            foreach (bool reply in ScopesOf(target.Scope))
            {
                string? actual = reply ? entry.ReplyForwardSignature : entry.NewSignature;
                if (!string.Equals(actual, name, StringComparison.Ordinal))
                {
                    return "its '" + (reply ? ReplyForwardSignatureValueName : NewSignatureValueName) + "' reads " + Quote(actual)
                        + ", not " + Quote(name);
                }
            }

            return null;
        }

        private static string Quote(string? value)
        {
            return value == null ? "nothing" : "'" + value + "'";
        }

        private static IReadOnlyList<SignatureDefaultChange>? ClearDanglingDefaults(
            ISignatureDefaultsStore store, string name, List<string> advice)
        {
            try
            {
                List<SignatureDefaultChange> cleared = new List<SignatureDefaultChange>();
                foreach (SignatureDefaultsRow row in store.ReadAccounts())
                {
                    if (string.Equals(row.NewMessage, name, StringComparison.OrdinalIgnoreCase))
                    {
                        store.ClearDefault(row.AccountKey, NewSignatureValueName);
                        cleared.Add(new SignatureDefaultChange(row.Profile ?? string.Empty, row.Account, "new", row.NewMessage, null));
                    }

                    if (string.Equals(row.ReplyForward, name, StringComparison.OrdinalIgnoreCase))
                    {
                        store.ClearDefault(row.AccountKey, ReplyForwardSignatureValueName);
                        cleared.Add(new SignatureDefaultChange(row.Profile ?? string.Empty, row.Account, "reply", row.ReplyForward, null));
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

        // ------------------------------------------------------------------ content: planning

        private sealed class ContentPlan
        {
            public Dictionary<string, byte[]> Files { get; } = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

            public List<string> FormatsWritten { get; } = new List<string>();

            public List<string> FormatsDerived { get; } = new List<string>();

            public Dictionary<string, byte[]> PicturesToWrite { get; } = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);

            public List<string> PicturesAdded { get; } = new List<string>();

            public List<string> PicturesReplaced { get; } = new List<string>();

            public List<string> PicturesToRemove { get; } = new List<string>();

            public byte[]? FileList { get; set; }

            public List<SignatureReplacementView> Replacements { get; } = new List<SignatureReplacementView>();

            public List<string> Warnings { get; } = new List<string>();
        }

        private static ContentPlan PlanContent(string action, string root, string name, ManageSignatureRequest request)
        {
            ContentPlan plan = new ContentPlan();
            string resourceDirectory = Path.Combine(root, name + SignatureCatalog.ResourceDirectorySuffix);
            HashSet<string> picturesBefore = new HashSet<string>(
                Directory.Exists(resourceDirectory)
                    ? Directory.GetFiles(resourceDirectory).Select(f => Path.GetFileName(f) ?? string.Empty).Where(f => f.Length > 0)
                    : Array.Empty<string>(),
                StringComparer.OrdinalIgnoreCase);
            PlanPictures(request.Pictures, picturesBefore, plan);
            HashSet<string> picturesAfter = new HashSet<string>(picturesBefore, StringComparer.OrdinalIgnoreCase);
            picturesAfter.UnionWith(plan.PicturesToWrite.Keys);
            picturesAfter.ExceptWith(plan.PicturesToRemove);

            byte[]? PictureBytes(string reference)
            {
                string? file = SignatureHtml.ResourceFileOf(reference, name)
                    ?? (Path.IsPathRooted(reference) ? null : reference.Trim());
                if (file != null && plan.PicturesToWrite.TryGetValue(file, out byte[]? planned))
                {
                    return planned;
                }

                if (file != null && picturesAfter.Contains(file) && File.Exists(Path.Combine(resourceDirectory, file)))
                {
                    return File.ReadAllBytes(Path.Combine(resourceDirectory, file));
                }

                if (Path.IsPathRooted(reference) && File.Exists(reference))
                {
                    return SignaturePictures.ReadNewPicture(reference, Path.GetFileName(reference));
                }

                return null;
            }

            if (action == "create")
            {
                PlanCreate(root, name, request, plan, PictureBytes, picturesAfter);
            }
            else
            {
                PlanUpdate(root, name, request, plan, PictureBytes, picturesAfter);
            }

            if (plan.PicturesToWrite.Count > 0 || plan.PicturesToRemove.Count > 0)
            {
                string fileList = Path.Combine(resourceDirectory, "filelist.xml");
                if (File.Exists(fileList))
                {
                    plan.FileList = UpdateFileList(File.ReadAllBytes(fileList), plan.PicturesAdded, plan.PicturesToRemove);
                }
            }

            return plan;
        }

        private static void PlanPictures(IReadOnlyList<SignaturePictureChange>? changes, HashSet<string> existing, ContentPlan plan)
        {
            foreach (SignaturePictureChange change in changes ?? Array.Empty<SignaturePictureChange>())
            {
                string file = (change?.File ?? string.Empty).Trim();
                if (file.Length == 0 || file.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0 || file == "." || file == "..")
                {
                    throw new ArgumentException("pictures: '" + file + "' is not a plain file name (no folders) for the picture folder. Nothing was written.");
                }

                if (string.Equals(file, "filelist.xml", StringComparison.OrdinalIgnoreCase))
                {
                    throw new ArgumentException("pictures: filelist.xml is Word's index of the picture folder, kept up to date by OutlookAI. Nothing was written.");
                }

                if (plan.PicturesToWrite.ContainsKey(file) || plan.PicturesToRemove.Contains(file, StringComparer.OrdinalIgnoreCase))
                {
                    throw new ArgumentException("pictures: '" + file + "' is named twice in one call. Nothing was written.");
                }

                if (change!.Remove)
                {
                    if (change.Path != null)
                    {
                        throw new ArgumentException("pictures: '" + file + "' has both remove and a path - say one. Nothing was written.");
                    }

                    if (!existing.Contains(file))
                    {
                        throw new ArgumentException("pictures: '" + file + "' is not in the signature's picture folder, so it cannot be removed. Nothing was written.");
                    }

                    plan.PicturesToRemove.Add(file);
                    continue;
                }

                if (string.IsNullOrWhiteSpace(change.Path))
                {
                    throw new ArgumentException("pictures: '" + file + "' needs a path to copy in, or remove=true. Nothing was written.");
                }

                plan.PicturesToWrite[file] = SignaturePictures.ReadNewPicture(change.Path!.Trim(), file);
                (existing.Contains(file) ? plan.PicturesReplaced : plan.PicturesAdded).Add(file);
            }
        }

        private static void PlanCreate(
            string root, string name, ManageSignatureRequest request, ContentPlan plan, Func<string, byte[]?> pictureBytes, HashSet<string> picturesAfter)
        {
            string? html = string.IsNullOrWhiteSpace(request.BodyHtml) ? null : EnsureHtmlDocument(request.BodyHtml!);
            string? rtf = string.IsNullOrWhiteSpace(request.BodyRtf) ? null : SignatureRtf.Expand(request.BodyRtf!, string.Empty, pictureBytes);
            string? text = string.IsNullOrWhiteSpace(request.BodyText) ? null : request.BodyText;
            if (rtf != null)
            {
                RequireRtfDocument(rtf);
            }

            if (html != null)
            {
                RequireReferencesResolve(html, name, picturesAfter);
            }

            if (text == null)
            {
                text = html != null ? Text.HtmlToText.Convert(html) : SignatureRtf.ToPlainText(rtf!);
                plan.FormatsDerived.Add(SignatureFormats.Text);
            }

            if (html == null)
            {
                html = BuildHtmlFromText(text);
                plan.FormatsDerived.Add(SignatureFormats.Html);
            }

            if (rtf == null)
            {
                rtf = !string.IsNullOrWhiteSpace(request.BodyHtml) ? SignatureRtf.FromHtml(html, pictureBytes) : SignatureRtf.FromText(text);
                plan.FormatsDerived.Add(SignatureFormats.Rtf);
            }

            plan.Files[Path.Combine(root, name + ".htm")] = SignatureText.EncodeHtml(html, null);
            plan.Files[Path.Combine(root, name + ".rtf")] = SignatureText.EncodeRtf(rtf);
            plan.Files[Path.Combine(root, name + ".txt")] = SignatureText.EncodeText(text, null);
            plan.FormatsWritten.AddRange(SignatureFormats.All);
        }

        private static void PlanUpdate(
            string root, string name, ManageSignatureRequest request, ContentPlan plan, Func<string, byte[]?> pictureBytes, HashSet<string> picturesAfter)
        {
            string? htmlPath = FirstExisting(Path.Combine(root, name + ".htm"), Path.Combine(root, name + ".html"));
            string rtfPath = Path.Combine(root, name + ".rtf");
            string txtPath = Path.Combine(root, name + ".txt");
            DecodedSignatureText? htmlBefore = htmlPath != null ? SignatureText.ReadHtml(htmlPath) : null;
            string? rtfBefore = File.Exists(rtfPath) ? SignatureText.ReadRtf(rtfPath).Text : null;
            DecodedSignatureText? textBefore = File.Exists(txtPath) ? SignatureText.ReadText(txtPath) : null;

            RequireEveryRendition(name, request, htmlBefore != null, rtfBefore != null, textBefore != null);

            string? html = htmlBefore?.Text;
            string? rtf = rtfBefore;
            string? text = textBefore?.Text;
            bool htmlChanged = false;
            bool rtfChanged = false;
            bool textChanged = false;

            if (!string.IsNullOrWhiteSpace(request.BodyHtml))
            {
                html = EnsureHtmlDocument(request.BodyHtml!);
                htmlChanged = true;
            }

            if (!string.IsNullOrWhiteSpace(request.BodyRtf))
            {
                rtf = SignatureRtf.Expand(request.BodyRtf!, rtfBefore ?? string.Empty, pictureBytes);
                RequireRtfDocument(rtf);
                rtfChanged = true;
            }

            if (!string.IsNullOrWhiteSpace(request.BodyText))
            {
                text = request.BodyText;
                textChanged = true;
            }

            foreach (SignatureTextEdit edit in request.Replace ?? Array.Empty<SignatureTextEdit>())
            {
                SignatureReplacementView view = new SignatureReplacementView { Find = edit.Find };
                List<string> split = new List<string>();
                if (html != null)
                {
                    SignatureTextReplacement result = SignatureHtml.ReplaceText(html, edit.Find, edit.Replace ?? string.Empty);
                    view.Html = result.Replaced;
                    if (result.FoundButSplit)
                    {
                        split.Add(SignatureFormats.Html);
                    }

                    htmlChanged |= result.Replaced > 0;
                    html = result.Content;
                }

                if (rtf != null)
                {
                    SignatureTextReplacement result = SignatureRtf.ReplaceText(rtf, edit.Find, edit.Replace ?? string.Empty);
                    view.Rtf = result.Replaced;
                    if (result.FoundButSplit)
                    {
                        split.Add(SignatureFormats.Rtf);
                    }

                    rtfChanged |= result.Replaced > 0;
                    rtf = result.Content;
                }

                if (text != null)
                {
                    int count = CountOrdinal(text, edit.Find);
                    view.Text = count;
                    if (count > 0)
                    {
                        text = text.Replace(edit.Find, edit.Replace ?? string.Empty);
                        textChanged = true;
                    }
                }

                if (split.Count > 0)
                {
                    throw new ArgumentException(
                        "replace: '" + edit.Find + "' appears in the " + string.Join(" and ", split) + " rendition, but formatting "
                        + "splits it there (a word in bold, a spelling mark), so it cannot be replaced without moving or dropping "
                        + "that formatting. Nothing was written. Read the signature with list_signatures(name) and send that "
                        + "rendition whole as body_" + split[0] + ", edited.");
                }

                if ((view.Html ?? 0) + (view.Rtf ?? 0) + (view.Text ?? 0) == 0)
                {
                    throw new ArgumentException(
                        "replace: '" + edit.Find + "' was not found in any rendition of '" + name + "'. It is matched exactly as it "
                        + "reads, case included. Nothing was written.");
                }

                plan.Replacements.Add(view);
            }

            if (htmlChanged && html != null)
            {
                RequireReferencesResolve(html, name, picturesAfter);
            }
            else if (html != null && plan.PicturesToRemove.Count > 0)
            {
                // A picture the unchanged HTML still shows may not be removed from under it.
                List<string> stillShown = SignatureHtml.References(html, name)
                    .Where(r => r.FileName != null && plan.PicturesToRemove.Contains(r.FileName, StringComparer.OrdinalIgnoreCase))
                    .Select(r => r.FileName!)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (stillShown.Count > 0)
                {
                    throw new ArgumentException(
                        "pictures: " + string.Join(", ", stillShown) + " is still shown by the HTML rendition. Remove its reference in "
                        + "the same call (replace, or body_html) or keep the picture. Nothing was written.");
                }
            }

            if (htmlChanged && htmlPath != null)
            {
                plan.Files[htmlPath] = SignatureText.EncodeHtml(html!, htmlBefore);
                plan.FormatsWritten.Add(SignatureFormats.Html);
            }
            else if (htmlChanged)
            {
                plan.Files[Path.Combine(root, name + ".htm")] = SignatureText.EncodeHtml(html!, null);
                plan.FormatsWritten.Add(SignatureFormats.Html);
            }

            if (rtfChanged)
            {
                plan.Files[rtfPath] = SignatureText.EncodeRtf(rtf!);
                plan.FormatsWritten.Add(SignatureFormats.Rtf);
            }

            if (textChanged)
            {
                plan.Files[txtPath] = SignatureText.EncodeText(text!, textBefore);
                plan.FormatsWritten.Add(SignatureFormats.Text);
            }

            if ((plan.PicturesReplaced.Count > 0 || plan.PicturesToRemove.Count > 0) && !rtfChanged && rtfBefore != null
                && SignatureRtf.FindPictures(rtfBefore).Count > 0)
            {
                plan.Warnings.Add(
                    "The RTF rendition embeds its own copies of its pictures, so it still shows the old ones: to change them there "
                    + "too, read the signature with list_signatures(name) and send body_rtf with [[picture-file:<file>]] where the "
                    + "old [[picture:N]] stood (or without it, to remove it).");
            }
        }

        /// <summary>
        /// Q205 (d): an update that sends a rendition whole sends every rendition the signature has. Nothing
        /// is converted from one format to another (Q159), and each mail format inserts only its own, so a
        /// rendition left out would go on saying what the signature said before. Refused, with the missing
        /// ones named, before anything is written. An update that only replaces text or changes pictures
        /// sends no rendition and is not affected.
        /// </summary>
        private static void RequireEveryRendition(string name, ManageSignatureRequest request, bool hasHtml, bool hasRtf, bool hasText)
        {
            List<string> sent = new List<string>();
            if (!string.IsNullOrWhiteSpace(request.BodyHtml))
            {
                sent.Add(SignatureFormats.Html);
            }

            if (!string.IsNullOrWhiteSpace(request.BodyRtf))
            {
                sent.Add(SignatureFormats.Rtf);
            }

            if (!string.IsNullOrWhiteSpace(request.BodyText))
            {
                sent.Add(SignatureFormats.Text);
            }

            if (sent.Count == 0)
            {
                return;
            }

            List<string> has = new List<string>();
            if (hasHtml)
            {
                has.Add(SignatureFormats.Html);
            }

            if (hasRtf)
            {
                has.Add(SignatureFormats.Rtf);
            }

            if (hasText)
            {
                has.Add(SignatureFormats.Text);
            }

            List<string> missing = has.Where(f => !sent.Contains(f)).ToList();
            if (missing.Count == 0)
            {
                return;
            }

            throw new ArgumentException(
                "update: '" + name + "' has " + DescribeRenditions(has) + ", and this update sends only "
                + string.Join(" and ", sent.Select(f => "body_" + f)) + ". An update that sends a rendition whole must send every "
                + "rendition the signature has: nothing is converted from one format to another, and each mail format inserts only "
                + "its own, so a rendition left out would go on saying what it said before. Missing: "
                + string.Join(" and ", missing.Select(f => "body_" + f)) + ". Read the current ones with list_signatures(name) and "
                + "send all of them, edited - or change text in every rendition at once with replace. Nothing was written.");
        }

        /// <summary>"an HTML, an RTF and a plain-text rendition" - the renditions a signature has, in words.</summary>
        private static string DescribeRenditions(IReadOnlyList<string> formats)
        {
            List<string> words = formats
                .Select(f => f == SignatureFormats.Html ? "an HTML" : f == SignatureFormats.Rtf ? "an RTF" : "a plain-text")
                .ToList();
            string joined = words.Count <= 2
                ? string.Join(" and ", words)
                : string.Join(", ", words.Take(words.Count - 1)) + " and " + words[words.Count - 1];
            return joined + " rendition";
        }

        private static void RequireRtfDocument(string rtf)
        {
            if (!rtf.TrimStart().StartsWith(@"{\rtf", StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    "body_rtf must be a whole RTF document, starting with {\\rtf1 - as list_signatures(name) returns it. Nothing was written.");
            }
        }

        private static void RequireReferencesResolve(string html, string name, HashSet<string> picturesAfter)
        {
            List<string> missing = new List<string>();
            List<string> stray = new List<string>();
            foreach (SignatureHtmlReference reference in SignatureHtml.References(html, name))
            {
                if (reference.FileName != null)
                {
                    if (!picturesAfter.Contains(reference.FileName))
                    {
                        missing.Add(reference.FileName);
                    }
                }
                else if (SignatureHtml.IsStrayLocalReference(reference))
                {
                    stray.Add(reference.Value);
                }
            }

            if (missing.Count > 0)
            {
                throw new ArgumentException(
                    "body_html refers to " + string.Join(", ", missing.Distinct(StringComparer.OrdinalIgnoreCase)) + " in '" + name
                    + SignatureCatalog.ResourceDirectorySuffix + "', which will not exist. Add it with 'pictures' in the same call. "
                    + "Nothing was written.");
            }

            if (stray.Count > 0)
            {
                throw new ArgumentException(
                    "body_html refers to " + string.Join(", ", stray.Distinct(StringComparer.Ordinal)) + " outside the signature's own "
                    + "picture folder, which Outlook cannot follow. Add the picture with 'pictures' and refer to it as '"
                    + SignatureHtml.ReferenceFor(name, string.Empty) + "<file>'. Nothing was written.");
            }
        }

        private static int CountOrdinal(string text, string find)
        {
            int count = 0;
            for (int at = text.IndexOf(find, StringComparison.Ordinal); at >= 0; at = text.IndexOf(find, at + find.Length, StringComparison.Ordinal))
            {
                count++;
            }

            return count;
        }

        private static string? FirstExisting(params string[] paths)
        {
            return paths.FirstOrDefault(File.Exists);
        }

        /// <summary>
        /// Word's <c>filelist.xml</c> with the added pictures listed and the removed ones dropped -
        /// edited as text, never parsed as XML (no DTD processing on a file in a user folder).
        /// </summary>
        private static byte[] UpdateFileList(byte[] original, IReadOnlyList<string> added, IReadOnlyList<string> removed)
        {
            DecodedSignatureText decoded = SignatureText.DecodeHtml(original);
            string xml = decoded.Text;
            foreach (string file in removed)
            {
                xml = Regex.Replace(xml, @"[ \t]*<o:File\s+HRef=""" + Regex.Escape(file) + @"""\s*/>\r?\n?", string.Empty,
                    RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, TimeSpan.FromSeconds(2));
            }

            int anchor = xml.IndexOf("<o:File HRef=\"filelist.xml\"", StringComparison.OrdinalIgnoreCase);
            if (anchor < 0)
            {
                anchor = xml.IndexOf("</xml>", StringComparison.OrdinalIgnoreCase);
            }

            if (anchor >= 0 && added.Count > 0)
            {
                StringBuilder entries = new StringBuilder();
                foreach (string file in added)
                {
                    entries.Append(" <o:File HRef=\"").Append(file).Append("\"/>\r\n");
                }

                xml = xml.Substring(0, anchor) + entries + xml.Substring(anchor);
            }

            Encoding encoding = decoded.HasBom ? new UTF8Encoding(true) : Encoding.GetEncoding(decoded.CodePage == 20127 ? 65001 : decoded.CodePage);
            byte[] body = encoding.GetBytes(xml);
            return decoded.HasBom ? encoding.GetPreamble().Concat(body).ToArray() : body;
        }

        // ------------------------------------------------------------------ content: writing

        private static void WritePlan(string root, string name, ContentPlan plan, ManageSignatureOutcome outcome)
        {
            Directory.CreateDirectory(root);
            List<string> written = new List<string>();
            foreach (KeyValuePair<string, byte[]> file in plan.Files)
            {
                WriteFileSafely(file.Key, file.Value);
                written.Add(file.Key);
            }

            string resourceDirectory = Path.Combine(root, name + SignatureCatalog.ResourceDirectorySuffix);
            if (plan.PicturesToWrite.Count > 0)
            {
                Directory.CreateDirectory(resourceDirectory);
            }

            foreach (KeyValuePair<string, byte[]> picture in plan.PicturesToWrite)
            {
                string path = Path.Combine(resourceDirectory, picture.Key);
                WriteFileSafely(path, picture.Value);
                written.Add(path);
            }

            if (plan.FileList != null)
            {
                string path = Path.Combine(resourceDirectory, "filelist.xml");
                WriteFileSafely(path, plan.FileList);
                written.Add(path);
            }

            foreach (string picture in plan.PicturesToRemove)
            {
                string path = Path.Combine(resourceDirectory, picture);
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }

            outcome.FilesWritten = written;
            outcome.FormatsWritten = plan.FormatsWritten.Count > 0 ? plan.FormatsWritten.ToList() : null;
            outcome.FormatsDerived = plan.FormatsDerived.Count > 0 ? plan.FormatsDerived.ToList() : null;
            outcome.PicturesAdded = plan.PicturesAdded.Count > 0 ? plan.PicturesAdded.ToList() : null;
            outcome.PicturesReplaced = plan.PicturesReplaced.Count > 0 ? plan.PicturesReplaced.ToList() : null;
            outcome.PicturesRemoved = plan.PicturesToRemove.Count > 0 ? plan.PicturesToRemove.ToList() : null;
            outcome.Replacements = plan.Replacements.Count > 0 ? plan.Replacements.ToList() : null;
        }

        /// <summary>Writes beside the target and swaps it in, so a failure never leaves a half-written rendition.</summary>
        private static void WriteFileSafely(string path, byte[] content)
        {
            string temporary = path + ".outlookai-new";
            File.WriteAllBytes(temporary, content);
            try
            {
                if (File.Exists(path))
                {
                    File.Replace(temporary, path, destinationBackupFileName: null, ignoreMetadataErrors: true);
                }
                else
                {
                    File.Move(temporary, path);
                }
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException or PlatformNotSupportedException)
            {
                File.Copy(temporary, path, overwrite: true);
                File.Delete(temporary);
            }
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

            string resourceDir = Path.Combine(directory, name + SignatureCatalog.ResourceDirectorySuffix);
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

        private static IReadOnlyList<string> DeleteFileSet(IReadOnlyList<string> existing)
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

        // ------------------------------------------------------------------ derivations (create only)

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
        /// specials escaped, non-ASCII as \'xx or \uN?.
        /// </summary>
        public static string BuildRtfFromText(string text)
        {
            return SignatureRtf.FromText(text);
        }
    }
}
