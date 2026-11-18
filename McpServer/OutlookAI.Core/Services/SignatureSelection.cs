using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OutlookAI.Core.Services
{
    /// <summary>What a draft tool's <c>signature</c> argument asks for (Q157 (a)).</summary>
    public enum SignatureRequestKind
    {
        /// <summary>Left out: the sending account's effective default (update_draft: keep the draft's signature).</summary>
        AccountDefault,

        /// <summary>"none": no signature (update_draft: remove the draft's signature).</summary>
        None,

        /// <summary>A signature by name.</summary>
        Named,
    }

    /// <summary>A parsed <c>signature</c> argument.</summary>
    public sealed class SignatureRequest
    {
        private SignatureRequest(SignatureRequestKind kind, string? name)
        {
            Kind = kind;
            Name = name;
        }

        /// <summary>The account's effective default.</summary>
        public static SignatureRequest AccountDefault { get; } = new SignatureRequest(SignatureRequestKind.AccountDefault, null);

        /// <summary>No signature.</summary>
        public static SignatureRequest NoSignature { get; } = new SignatureRequest(SignatureRequestKind.None, null);

        /// <summary>What was asked for.</summary>
        public SignatureRequestKind Kind { get; }

        /// <summary>The name, for <see cref="SignatureRequestKind.Named"/>.</summary>
        public string? Name { get; }

        /// <summary>A signature by name.</summary>
        public static SignatureRequest Named(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                throw new ArgumentException("A named signature needs a name.", nameof(name));
            }

            return new SignatureRequest(SignatureRequestKind.Named, name.Trim());
        }

        /// <summary>Null or blank is the default, "none" (any case) is no signature, anything else a name. Pure.</summary>
        public static SignatureRequest Parse(string? argument)
        {
            if (string.IsNullOrWhiteSpace(argument))
            {
                return AccountDefault;
            }

            string trimmed = argument!.Trim();
            return string.Equals(trimmed, SignatureManager.NoSignatureKeyword, StringComparison.OrdinalIgnoreCase)
                ? NoSignature
                : Named(trimmed);
        }
    }

    /// <summary>The account a draft is written in, as Outlook names it.</summary>
    public sealed class SignatureSendingAccount
    {
        /// <summary>Creates the description.</summary>
        public SignatureSendingAccount(string? smtpAddress, string? displayName)
        {
            SmtpAddress = string.IsNullOrWhiteSpace(smtpAddress) ? null : smtpAddress!.Trim();
            DisplayName = string.IsNullOrWhiteSpace(displayName) ? null : displayName!.Trim();
        }

        /// <summary><c>Account.SmtpAddress</c>.</summary>
        public string? SmtpAddress { get; }

        /// <summary>
        /// <c>Account.DisplayName</c> - for an Exchange account the name its profile entry carries as
        /// <c>Account Name</c>, which is how an account signed in under an alias is still found.
        /// </summary>
        public string? DisplayName { get; }

        /// <summary>The address to show: the SMTP address, else the display name.</summary>
        public string Label => SmtpAddress ?? DisplayName ?? "(unknown account)";
    }

    /// <summary>
    /// What goes into a draft, and why (Q155 (c), Q156, Q157 (a), Q160) - the answer every draft
    /// result carries and the instruction the COM host composes by.
    /// </summary>
    public sealed class SignatureDecision
    {
        /// <summary>The draft signature came from the agent's choice.</summary>
        public const string SourceRequested = "requested";

        /// <summary>The draft signature is the sending account's own default in the profile.</summary>
        public const string SourceAccountDefault = "accountDefault";


        /// <summary>No signature, and why is in <see cref="Reason"/>.</summary>
        public const string SourceNone = "none";

        internal SignatureDecision(
            SignatureRequestKind requested,
            string? name,
            string? filePath,
            string source,
            string reason,
            string? profile,
            string? suggestion,
            string? suggestionReason,
            string? refusal,
            bool removeExisting)
        {
            Requested = requested;
            Name = name;
            FilePath = filePath;
            Source = source;
            Reason = reason;
            Profile = profile;
            Suggestion = suggestion;
            SuggestionReason = suggestionReason;
            Refusal = refusal;
            RemoveExisting = removeExisting;
        }

        /// <summary>What the agent asked for.</summary>
        public SignatureRequestKind Requested { get; }

        /// <summary>The signature to insert, or null for none.</summary>
        public string? Name { get; }

        /// <summary>Its file to insert (.htm preferred), or null.</summary>
        public string? FilePath { get; }

        /// <summary>
        /// requested | accountDefault | none. (policyDefault, a policy's or deployed default, is gone since 2026-11-18:
        /// Outlook was measured not to use one, <see cref="SignaturePolicy.EffectiveDefault"/>.)
        /// </summary>
        public string Source { get; }

        /// <summary>Why, in one sentence a person can read.</summary>
        public string Reason { get; }

        /// <summary>The Outlook profile whose defaults were consulted, or null when none could be.</summary>
        public string? Profile { get; }

        /// <summary>The sending account's own default when the draft got something else, or null.</summary>
        public string? Suggestion { get; }

        /// <summary>Why <see cref="Suggestion"/> is suggested.</summary>
        public string? SuggestionReason { get; }

        /// <summary>Set when the request is refused: the message, and nothing may be composed.</summary>
        public string? Refusal { get; }

        /// <summary>True when the draft's existing signature must go (signature="none").</summary>
        public bool RemoveExisting { get; }

        /// <summary>True when the signature is a default rather than the agent's pick (Outlook's own injection then stands).</summary>
        public bool IsDefault => Source == SourceAccountDefault;
    }

    /// <summary>
    /// The signature a draft gets - one rule for the server's draft tools now and the add-in's
    /// language pick later (Q158/Q162 (b): it will call <see cref="Candidates"/> and <see cref="Decide"/>
    /// once the add-in references Core).
    /// <list type="bullet">
    /// <item><b>Left out</b>: the sending account's EFFECTIVE default for new mail or for replies and
    /// forwards - its own in the profile Outlook is RUNNING (else the recorded default profile), whatever a
    /// policy or deployed setting names, as Outlook does (<see cref="SignaturePolicy.EffectiveDefault"/>).</item>
    /// <item><b>"none"</b>: no signature.</item>
    /// <item><b>A name</b>: that signature - only one the sending account can use: any signature of
    /// this device, and a roaming copy only for its own account.</item>
    /// </list>
    /// Policy (Q156): no signature at all while signatures are disabled. Under a forced default one message may
    /// still have another signature or none, because Outlook's own compose window allows both: its Signature button
    /// stays open (measured 2026-11-18, Q200 (a), <c>Docs/live-tier-on-the-vm.md</c> section 4.9).
    /// Delegate and shared mailboxes (Q160): no account default applies, so none goes in unless named.
    /// </summary>
    public static class SignatureSelector
    {
        /// <summary>
        /// The decision for one draft. <paramref name="sendingAccount"/> is null for mail in a store no
        /// account delivers to (a delegate or shared mailbox). <paramref name="runningProfile"/> is the
        /// profile Outlook runs on, when known. <paramref name="replyOrForward"/> is null when it is
        /// not known (update_draft): then a policy forcing either default is honoured. Pure.
        /// </summary>
        public static SignatureDecision Decide(
            SignatureEnvironment environment,
            SignatureRequest request,
            SignatureSendingAccount? sendingAccount,
            bool? replyOrForward,
            string? runningProfile)
        {
            if (environment == null)
            {
                throw new ArgumentNullException(nameof(environment));
            }

            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            SignaturePolicy policy = environment.Policy;
            string? profile = ProfileFor(environment, runningProfile, out string profileNote);

            if (policy.SignaturesDisabled)
            {
                string disabled = "signatures are disabled on this computer by " + SignaturePolicy.Describe(policy.DisabledBy) + " (DisableSignatures)";
                return request.Kind == SignatureRequestKind.Named
                    ? Refuse(request, "signature '" + request.Name + "' was not added: " + disabled + ", so a draft carries none. "
                        + "Leave signature out, or pass \"none\".", profile)
                    : None(request, request.Kind == SignatureRequestKind.None ? "signature=\"none\" was asked for" : disabled, profile, null, null);
            }

            ProfileAccountEntry? entry = null;
            string? entryNote = null;
            if (sendingAccount != null && profile != null)
            {
                entry = FindAccountEntry(environment, profile, sendingAccount, out entryNote);
            }

            Effective effective = EffectiveDefault(environment, policy, entry, sendingAccount, replyOrForward, profile, profileNote, entryNote);

            switch (request.Kind)
            {
                case SignatureRequestKind.Named:
                {
                    SignatureInfo? signature = environment.FindSignature(request.Name);
                    if (signature?.PreferredFilePath == null)
                    {
                        return Refuse(request, "signature '" + request.Name + "' was not found" + UsableList(environment, sendingAccount), profile);
                    }

                    string? otherAccount = OwnerIfNotUsable(environment, signature, sendingAccount);
                    if (otherAccount != null)
                    {
                        return Refuse(request, "signature '" + signature.Name + "' is a roaming signature of " + otherAccount
                            + ", and Outlook offers it only to that account" + UsableList(environment, sendingAccount), profile);
                    }

                    (string? suggestion, string? why) = Suggest(effective, signature.Name);
                    return new SignatureDecision(
                        request.Kind, signature.Name, signature.PreferredFilePath, SignatureDecision.SourceRequested,
                        "signature '" + signature.Name + "' was asked for", profile, suggestion, why, null, false);
                }

                case SignatureRequestKind.None:
                {
                    (string? suggestion, string? why) = Suggest(effective, null);
                    return None(request, "signature=\"none\" was asked for", profile, suggestion, why);
                }

                default:
                {
                    if (effective.Name == null)
                    {
                        // Q155 (c): suggest what matches the sending account - its own roaming signature, when it has one.
                        SignatureInfo? own = sendingAccount?.SmtpAddress == null || !environment.Roaming.MayRoamFor(sendingAccount.SmtpAddress)
                            ? null
                            : environment.Signatures.FirstOrDefault(
                                s => ProfileAccountEntries.AddressEquals(s.ScopeAccount, sendingAccount.SmtpAddress));
                        return None(request, effective.Reason, profile, own?.Name,
                            own == null ? null : "'" + own.Name + "' is a roaming signature of " + sendingAccount!.Label);
                    }

                    SignatureInfo? signature = environment.FindSignature(effective.Name);
                    if (signature?.PreferredFilePath == null)
                    {
                        return None(request, effective.Reason + ", but no signature of that name is in " + environment.Folder.Path
                            + ", so none was added", profile, null, null);
                    }

                    return new SignatureDecision(
                        request.Kind, signature.Name, signature.PreferredFilePath, effective.Source, effective.Reason, profile, null, null,
                        null, false);
                }
            }
        }

        /// <summary>
        /// Checks a request BEFORE any COM work, as far as it can be checked without Outlook: that a
        /// named signature exists, that policy allows it, and - when the sending account is already
        /// known (new_draft) - that the account can use it. Throws the refusal as an
        /// <see cref="ArgumentException"/>; returns the signature to hand the COM host.
        /// </summary>
        public static SignatureInfo? Validate(
            SignatureEnvironment environment, SignatureRequest request, string? knownSendingAccount, bool? replyOrForward)
        {
            if (request.Kind == SignatureRequestKind.AccountDefault)
            {
                return null;
            }

            SignatureSendingAccount? account = knownSendingAccount == null ? null : new SignatureSendingAccount(knownSendingAccount, null);
            SignatureDecision decision = Decide(environment, request, account, replyOrForward, null);
            if (decision.Refusal != null)
            {
                throw new SignatureRefusedException("signature_not_usable", decision.Refusal);
            }

            return request.Kind == SignatureRequestKind.Named ? environment.FindSignature(request.Name) : null;
        }

        /// <summary>
        /// The signatures an account can use, its effective defaults, and the policy limits - what the
        /// add-in's language pick will choose from (Q158/Q162 (b)), and what a draft result suggests
        /// from. Pure.
        /// </summary>
        public static SignatureCandidates Candidates(
            SignatureEnvironment environment, SignatureSendingAccount? sendingAccount, string? runningProfile)
        {
            if (environment == null)
            {
                throw new ArgumentNullException(nameof(environment));
            }

            string? profile = ProfileFor(environment, runningProfile, out string profileNote);
            ProfileAccountEntry? entry = sendingAccount != null && profile != null
                ? FindAccountEntry(environment, profile, sendingAccount, out _)
                : null;
            Effective newMessage = EffectiveDefault(environment, environment.Policy, entry, sendingAccount, false, profile, profileNote, null);
            Effective reply = EffectiveDefault(environment, environment.Policy, entry, sendingAccount, true, profile, profileNote, null);
            List<SignatureInfo> usable = environment.Policy.SignaturesDisabled
                ? new List<SignatureInfo>()
                : environment.Signatures.Where(s => OwnerIfNotUsable(environment, s, sendingAccount) == null).ToList();
            return new SignatureCandidates(
                profile,
                usable,
                newMessage.Name,
                reply.Name,
                environment.Policy.SignaturesDisabled,
                environment.Policy.NewMessageForced || environment.Policy.ReplyForwardForced);
        }

        // ------------------------------------------------------------------ the parts

        private readonly struct Effective
        {
            public Effective(string? name, string source, string reason)
            {
                Name = name;
                Source = source;
                Reason = reason;
            }

            public string? Name { get; }

            public string Source { get; }

            public string Reason { get; }
        }

        /// <summary>
        /// The profile whose defaults apply: the running one, else the recorded default, else the
        /// only one. Null - with the reason - when none of those exists.
        /// </summary>
        private static string? ProfileFor(SignatureEnvironment environment, string? runningProfile, out string note)
        {
            string? running = environment.FindProfile(runningProfile) ?? (string.IsNullOrWhiteSpace(runningProfile) ? null : runningProfile!.Trim());
            if (running != null)
            {
                note = "the profile Outlook is running, '" + running + "'";
                return running;
            }

            string? recorded = environment.FindProfile(environment.Profiles.DefaultProfile);
            if (recorded != null)
            {
                note = "the default profile '" + recorded + "' (Outlook's running profile was not known)";
                return recorded;
            }

            IReadOnlyList<string> names = environment.ProfileNames;
            if (names.Count == 1)
            {
                note = "the only profile, '" + names[0] + "'";
                return names[0];
            }

            note = names.Count == 0
                ? "no Outlook profile was found"
                : "Outlook's running profile was not known and none of the " + names.Count + " profiles is recorded as the default";
            return null;
        }

        /// <summary>
        /// The sending account's entry in one profile: by its address first, then - for an Exchange
        /// account whose profile entry is named after another address of the mailbox (an alias) - by
        /// the name Outlook shows it under. Null, with the reason, when neither finds exactly one.
        /// </summary>
        private static ProfileAccountEntry? FindAccountEntry(
            SignatureEnvironment environment, string profile, SignatureSendingAccount account, out string? note)
        {
            IReadOnlyList<ProfileAccountEntry> accounts = environment.MailAccountsOf(profile);
            List<ProfileAccountEntry> byAddress = accounts.Where(e => ProfileAccountEntries.AddressEquals(e.Address, account.SmtpAddress)).ToList();
            if (byAddress.Count == 1)
            {
                note = null;
                return byAddress[0];
            }

            List<ProfileAccountEntry> byName = byAddress.Count == 0 && account.DisplayName != null
                ? accounts.Where(e => e.AccountKind == MailAccountKind.Exchange
                    && string.Equals(e.AccountName, account.DisplayName, StringComparison.OrdinalIgnoreCase)).ToList()
                : new List<ProfileAccountEntry>();
            if (byName.Count == 1)
            {
                note = null;
                return byName[0];
            }

            int candidates = byAddress.Count > 0 ? byAddress.Count : byName.Count;
            note = candidates > 1
                ? candidates + " accounts of profile '" + profile + "' match " + account.Label + ", so OutlookAI cannot tell whose default applies"
                : "no account of profile '" + profile + "' matches " + account.Label + " (by address, or for Exchange by the name Outlook "
                    + "shows it under)";
            return null;
        }

        private static Effective EffectiveDefault(
            SignatureEnvironment environment,
            SignaturePolicy policy,
            ProfileAccountEntry? entry,
            SignatureSendingAccount? account,
            bool? replyOrForward,
            string? profile,
            string profileNote,
            string? entryNote)
        {
            bool reply = replyOrForward ?? false;
            string scope = reply ? "replies and forwards" : "new messages";
            string? own = account != null && profile != null && entry != null
                ? (reply ? entry.ReplyForwardSignature : entry.NewSignature)
                : null;
            string? name = policy.EffectiveDefault(reply, own);
            if (account == null)
            {
                return new Effective(null, SignatureDecision.SourceNone,
                    "the mail is in a delegate or shared mailbox that no account of the profile delivers to, so no account's default "
                    + "applies and OutlookAI adds none for now (Q160) - name one with signature to add it");
            }

            if (profile == null)
            {
                return new Effective(null, SignatureDecision.SourceNone, profileNote + ", so no account default could be read");
            }

            if (entry == null)
            {
                return new Effective(null, SignatureDecision.SourceNone, (entryNote ?? "the account was not found in profile '" + profile + "'")
                    + ", so no default was added");
            }

            if (name == null)
            {
                return new Effective(null, SignatureDecision.SourceNone,
                    account.Label + " has no default signature for " + scope + " in " + profileNote);
            }

            return new Effective(name, SignatureDecision.SourceAccountDefault,
                "'" + name + "' is the default of " + account.Label + " for " + scope + " in " + profileNote);
        }

        /// <summary>The account a roaming copy belongs to when it is not the sending account's, else null.</summary>
        private static string? OwnerIfNotUsable(SignatureEnvironment environment, SignatureInfo signature, SignatureSendingAccount? account)
        {
            string? owner = signature.ScopeAccount;
            if (owner == null || !environment.Roaming.MayRoamFor(owner))
            {
                return null;
            }

            return account?.SmtpAddress != null && ProfileAccountEntries.AddressEquals(owner, account.SmtpAddress) ? null : owner;
        }


        private static (string? Suggestion, string? Reason) Suggest(Effective effective, string? used)
        {
            if (effective.Name == null || string.Equals(effective.Name, used, StringComparison.OrdinalIgnoreCase))
            {
                return (null, null);
            }

            return (effective.Name, effective.Reason);
        }

        private static string UsableList(SignatureEnvironment environment, SignatureSendingAccount? account)
        {
            List<string> usable = environment.Signatures
                .Where(s => OwnerIfNotUsable(environment, s, account) == null)
                .Select(s => s.Name)
                .ToList();
            return usable.Count > 0
                ? ". Signatures " + (account?.SmtpAddress != null ? account.SmtpAddress + " can use" : "available") + ": "
                    + string.Join(", ", usable) + " - or \"none\". See list_signatures."
                : " - no signatures are installed in " + environment.Folder.Path + " (see list_signatures).";
        }

        private static SignatureDecision Refuse(SignatureRequest request, string message, string? profile)
        {
            return new SignatureDecision(request.Kind, null, null, SignatureDecision.SourceNone, message, profile, null, null, message, false);
        }

        private static SignatureDecision None(SignatureRequest request, string reason, string? profile, string? suggestion, string? suggestionReason)
        {
            return new SignatureDecision(
                request.Kind, null, null, SignatureDecision.SourceNone, reason, profile, suggestion, suggestionReason, null,
                request.Kind == SignatureRequestKind.None);
        }
    }

    /// <summary>
    /// What one account can choose from (data only) - the input the add-in's language pick will
    /// use (Q158/Q162 (b)): the signatures it may use, with their excerpts for the language, its
    /// effective defaults, and whether policy leaves any choice at all.
    /// </summary>
    public sealed class SignatureCandidates
    {
        /// <summary>Creates the description.</summary>
        public SignatureCandidates(
            string? profile,
            IReadOnlyList<SignatureInfo> usable,
            string? newMessageDefault,
            string? replyForwardDefault,
            bool signaturesDisabled,
            bool defaultsForced)
        {
            Profile = profile;
            Usable = usable;
            NewMessageDefault = newMessageDefault;
            ReplyForwardDefault = replyForwardDefault;
            SignaturesDisabled = signaturesDisabled;
            DefaultsForced = defaultsForced;
        }

        /// <summary>The profile the defaults come from.</summary>
        public string? Profile { get; }

        /// <summary>The signatures the account may use (empty when policy disables signatures).</summary>
        public IReadOnlyList<SignatureInfo> Usable { get; }

        /// <summary>The effective new-message default.</summary>
        public string? NewMessageDefault { get; }

        /// <summary>The effective reply/forward default.</summary>
        public string? ReplyForwardDefault { get; }

        /// <summary>True when policy disables signatures: nothing may be inserted.</summary>
        public bool SignaturesDisabled { get; }

        /// <summary>
        /// True when a policy names a default, so OutlookAI does not change the account's own (Q156). One message may
        /// still have another signature, as Outlook's own compose window allows (measured 2026-11-18, Q200 (a)).
        /// </summary>
        public bool DefaultsForced { get; }
    }
}
